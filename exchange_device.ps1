param(
    [string]$Users,
    [string]$DeviceIds
)

$logDir    = "C:\Logs"
$log       = Join-Path $logDir "awx-clean-mobiledevices.log"
$reportCsv = Join-Path $logDir ("mobiledevice-cleanup-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

function Log($msg) {
    Add-Content $log "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
}

Log "==================================================================="
Log "SCRIPT VERSION: 2.1-diagnostic"
Log "Raw -Users param    : [$Users]"
Log "Raw -DeviceIds param: [$DeviceIds]"

# --- Load Exchange snap-in in THIS session ---
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        Log "Snap-in loaded successfully."
    } catch {
        Log "FATAL: Could not load Exchange snap-in: $($_.Exception.Message)"
        Write-Output "===RESULT_JSON_START==="
        @([PSCustomObject]@{ User='N/A'; DeviceId='N/A'; OverallResult="FATAL: snap-in load failed - $($_.Exception.Message)" }) | ConvertTo-Json -Depth 4
        Write-Output "===RESULT_JSON_END==="
        exit 1
    }
}

if (-not (Get-Command Get-MobileDevice -ErrorAction SilentlyContinue)) {
    Log "FATAL: Get-MobileDevice not available after snap-in load."
    Write-Output "===RESULT_JSON_START==="
    @([PSCustomObject]@{ User='N/A'; DeviceId='N/A'; OverallResult='FATAL: Get-MobileDevice not available' }) | ConvertTo-Json -Depth 4
    Write-Output "===RESULT_JSON_END==="
    exit 1
}

# --- Clean input ---
$UserList     = $Users     -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
$InputDevices = $DeviceIds -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

Log "START PROCESS"
Log "Users to process (count=$($UserList.Count)): $($UserList -join ', ')"
Log "Device filter (count=$($InputDevices.Count)): $(if ($InputDevices.Count -gt 0) { $InputDevices -join ', ' } else { 'NONE (all devices)' })"
Log "Block policy: always block DeviceId + clear any pre-existing allow entry (no reliance on org Quarantine default alone)"

if ($UserList.Count -eq 0) {
    Log "FATAL: No users parsed from -Users input. Check the value being passed from Ansible."
}

if ($InputDevices.Count -eq 0) {
    Log "FATAL: No DeviceIds provided. Refusing to run - this script only acts on DeviceIds explicitly supplied via the survey, never on 'all devices for a user'."
    Write-Output "===RESULT_JSON_START==="
    @([PSCustomObject]@{ User = ($UserList -join ','); DeviceId = 'N/A'; OverallResult = 'FATAL: no DeviceIds supplied - refusing to process all devices' }) | ConvertTo-Json -Depth 4
    Write-Output "===RESULT_JSON_END==="
    exit 1
}

function Test-IsDesktopClient {
    # Get-MobileDevice also returns Mac Mail.app and Windows Mail app partnerships,
    # since both connect over ActiveSync (EAS) just like a phone. This function flags
    # those so the script never wipes/removes/blocks a desktop mail session.
    #
    # IMPORTANT: DeviceType "Outlook" alone is ambiguous - the modern "Outlook for iOS
    # and Android" mobile app also reports DeviceType = "Outlook" (with DeviceModel =
    # "Outlook for iOS and Android"). So we check DeviceModel/DeviceOS for explicit
    # mobile-OS indicators FIRST and treat those as an override that always wins.
    param($DeviceType, $DeviceOS, $DeviceUserAgent, $DeviceModel)

    $mobileIndicatorPatterns = @(
        'iOS', 'iPhone', 'iPad', 'Android'
    )
    foreach ($p in $mobileIndicatorPatterns) {
        if ($DeviceModel -and $DeviceModel -match $p)      { return $false }
        if ($DeviceOS -and $DeviceOS -match $p)            { return $false }
        if ($DeviceUserAgent -and $DeviceUserAgent -match $p) { return $false }
    }

    $desktopTypePatterns = @(
        '^WindowsMail$', '^MacOutlook$', '^OutlookMac$'
    )
    $desktopOsPatterns = @(
        'Mac OS X', 'macOS', 'Windows NT', 'Windows 10\.0', 'Windows 11'
    )

    foreach ($p in $desktopTypePatterns) {
        if ($DeviceType -and $DeviceType -match $p) { return $true }
    }
    foreach ($p in $desktopOsPatterns) {
        if ($DeviceOS -and $DeviceOS -match $p)          { return $true }
        if ($DeviceUserAgent -and $DeviceUserAgent -match $p) { return $true }
        if ($DeviceModel -and $DeviceModel -match $p)    { return $true }
    }
    return $false
}

$Results = [System.Collections.Generic.List[object]]::new()

function New-Result {
    param($User, $DeviceId, $DeviceType, $DeviceModel, $LastSync)
    [PSCustomObject]@{
        Timestamp     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        User          = $User
        DeviceId      = $DeviceId
        DeviceType    = $DeviceType
        DeviceModel   = $DeviceModel
        LastSync      = $LastSync
        WipeStatus    = 'NotAttempted'
        WipeDetail    = ''
        RemoveStatus  = 'NotAttempted'
        RemoveDetail  = ''
        BlockStatus   = 'Skipped'
        BlockDetail   = ''
        OverallResult = 'Pending'
    }
}

foreach ($User in $UserList) {
    Log "----------------------------------------"
    Log "USER: [$User]"

    # --- Diagnostic: confirm the mailbox actually resolves ---
    try {
        $mbx = Get-Mailbox -Identity $User -ErrorAction Stop
        Log "Mailbox resolved: $($mbx.DisplayName) <$($mbx.PrimarySmtpAddress)>"
    } catch {
        Log "ERROR: Get-Mailbox could not resolve identity '$User' -> $($_.Exception.Message)"
        $r = New-Result -User $User -DeviceId 'N/A' -DeviceType 'N/A' -DeviceModel 'N/A' -LastSync 'N/A'
        $r.OverallResult = "FAILED: mailbox identity did not resolve - $($_.Exception.Message)"
        $Results.Add($r)
        continue
    }

    try {
        $Devices = Get-MobileDevice -Mailbox $User -ErrorAction Stop
        Log "Get-MobileDevice returned $($Devices.Count) device(s) for $User"
    } catch {
        Log "ERROR: Get-MobileDevice failed for $User -> $($_.Exception.Message)"
        $r = New-Result -User $User -DeviceId 'N/A' -DeviceType 'N/A' -DeviceModel 'N/A' -LastSync 'N/A'
        $r.OverallResult = "FAILED: Get-MobileDevice error - $($_.Exception.Message)"
        $Results.Add($r)
        continue
    }

    if (-not $Devices -or $Devices.Count -eq 0) {
        Log "No mobile devices found for $User"
        $r = New-Result -User $User -DeviceId 'N/A' -DeviceType 'N/A' -DeviceModel 'N/A' -LastSync 'N/A'
        $r.OverallResult = 'No devices found'
        $Results.Add($r)
        continue
    }

    foreach ($Device in $Devices) {
        $id = $Device.DeviceId
        Log "Found device: Id=$id Type=$($Device.DeviceType) OS=$($Device.DeviceOS) Model=$($Device.DeviceModel)"

        # --- Guard: never process Mac/PC desktop mail clients, only real mobile devices ---
        $isDesktop = Test-IsDesktopClient -DeviceType $Device.DeviceType -DeviceOS $Device.DeviceOS `
                                           -DeviceUserAgent $Device.DeviceUserAgent -DeviceModel $Device.DeviceModel
        Log "Classification for $id : DeviceType=$($Device.DeviceType) DeviceModel=$($Device.DeviceModel) DeviceOS=$($Device.DeviceOS) -> IsDesktop=$isDesktop"
        if ($isDesktop) {
            Log "SKIP (desktop/Mac/PC client, NOT a mobile device - will not wipe/remove/block): $id Type=$($Device.DeviceType) OS=$($Device.DeviceOS)"
            $r = New-Result -User $User -DeviceId $id -DeviceType $Device.DeviceType -DeviceModel $Device.DeviceModel -LastSync 'N/A'
            $r.OverallResult = 'Skipped - desktop/Mac/PC client, not a mobile device'
            $Results.Add($r)
            continue
        }

        # --- Guard: only act on DeviceIds explicitly supplied via the survey.
        # Every other device on this mailbox is skipped, no exceptions. ---
        $normId     = ($id -replace '-', '').ToUpper()
        $normFilter = $InputDevices | ForEach-Object { ($_ -replace '-', '').ToUpper() }

        if ($normFilter -notcontains $normId) {
            Log "SKIP (DeviceId not in survey-supplied list): $id -- left untouched"
            $r = New-Result -User $User -DeviceId $id -DeviceType $Device.DeviceType -DeviceModel $Device.DeviceModel -LastSync 'N/A'
            $r.OverallResult = 'Skipped - DeviceId not requested in survey'
            $Results.Add($r)
            continue
        }
        Log "MATCH: $id is in the survey-supplied DeviceIds list, proceeding"

        # IMPORTANT: DeviceId (bare string) is NOT accepted by Clear-MobileDevice /
        # Remove-MobileDevice -Identity on most on-prem Exchange builds. Those cmdlets
        # need the full device Identity path (Device.Identity), not Device.DeviceId.
        # DeviceId is still used above for filtering since that's the human-friendly
        # value operators supply, but the action cmdlets must use $Device.Identity.
        $identity = $Device.Identity
        if ([string]::IsNullOrWhiteSpace($identity)) {
            Log "WARN: Device.Identity was empty for $id, falling back to DeviceId (may fail)"
            $identity = $id
        }
        Log "Resolved Identity for action cmdlets: $identity"

        $stats = $null
        try {
            $stats = Get-MobileDeviceStatistics -Identity $identity -ErrorAction Stop
        } catch {
            Log "WARN: could not read stats for $id -> $($_.Exception.Message)"
        }

        $r = New-Result -User $User -DeviceId $id `
             -DeviceType  $(if ($stats) { $stats.DeviceType } else { $Device.DeviceType }) `
             -DeviceModel $(if ($stats) { $stats.DeviceModel } else { $Device.DeviceModel }) `
             -LastSync    $(if ($stats) { $stats.LastSuccessSync } else { 'Unknown' })

        # 1. Account-only wipe is NOT supported by on-prem Exchange's Clear-MobileDevice
        #    (-AccountOnly is an Exchange Online-only parameter). On-prem only offers a
        #    FULL device wipe via Clear-MobileDevice, which was declined as too destructive
        #    relative to the goal (remove only the account). Wipe step is intentionally
        #    skipped; Remove-MobileDevice + explicit block below is what actually protects
        #    the mailbox going forward.
        $r.WipeStatus = 'Skipped'
        $r.WipeDetail = 'On-prem Exchange does not support account-only wipe (Online-only parameter); full wipe declined as too destructive'
        Log "WIPE SKIPPED for $id (on-prem limitation, full-wipe not authorized)"

        # 2. Remove partnership -> kills the active EAS session
        try {
            Remove-MobileDevice -Identity $identity -Confirm:$false -ErrorAction Stop
            $r.RemoveStatus = 'Success'
            $r.RemoveDetail = 'Partnership removed, session terminated'
            Log "REMOVED: $id"
        } catch {
            $r.RemoveStatus = 'Failed'
            $r.RemoveDetail = $_.Exception.Message
            Log "ERROR removing $id -> $($_.Exception.Message)"
        }

        # 3. Force explicit block for this DeviceId. We no longer rely solely on the
        #    org-wide Quarantine default: a per-mailbox ActiveSyncAllowedDeviceIDs entry,
        #    or an org-wide Device Access Rule allowing this DeviceType (e.g. "Outlook
        #    for iOS and Android" is commonly pre-allowed), can silently bypass Quarantine
        #    and let the same DeviceId reconnect with zero approval. So: always clear any
        #    existing allow-entry, then always block, regardless of $PermanentlyBlock.
        try {
            $cas = Get-CASMailbox -Identity $User -ErrorAction Stop
            if ($cas.ActiveSyncAllowedDeviceIDs -contains $id) {
                Set-CASMailbox -Identity $User -ActiveSyncAllowedDeviceIDs @{remove=$id} -ErrorAction Stop
                Log "Removed pre-existing ALLOW entry for $id on $User (was bypassing Quarantine)"
            }
        } catch {
            Log "WARN: could not check/clear ActiveSyncAllowedDeviceIDs for $User -> $($_.Exception.Message)"
        }

        try {
            Set-CASMailbox -Identity $User -ActiveSyncBlockedDeviceIDs @{add=$id} -ErrorAction Stop
            $r.BlockStatus = 'Success'
            $r.BlockDetail = 'DeviceID blocked - requires manual admin re-approval to reconnect'
            Log "BLOCKED: $id"
        } catch {
            $r.BlockStatus = 'Failed'
            $r.BlockDetail = $_.Exception.Message
            Log "ERROR blocking $id on $User -> $($_.Exception.Message)"
        }

        if ($r.RemoveStatus -eq 'Success' -and $r.BlockStatus -eq 'Success') {
            $r.OverallResult = 'SUCCESS'
        } elseif ($r.RemoveStatus -eq 'Success') {
            $r.OverallResult = 'PARTIAL (session killed, block failed - device could reconnect unblocked)'
        } else {
            $r.OverallResult = 'FAILED'
        }

        $Results.Add($r)
    }
}

Log "END PROCESS"
$desktopSkipped  = ($Results | Where-Object { $_.OverallResult -match 'desktop/Mac/PC' }).Count
$notRequested    = ($Results | Where-Object { $_.OverallResult -match 'not requested in survey' }).Count
$actuallyProcessed = ($Results | Where-Object { $_.OverallResult -in @('SUCCESS','FAILED') -or $_.OverallResult -match 'PARTIAL' }).Count
Log "Summary: $actuallyProcessed device(s) actually processed, $desktopSkipped desktop/Mac/PC device(s) skipped, $notRequested device(s) skipped (not requested in survey)"

$Results | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
Log "Report written: $reportCsv"

Write-Output "===RESULT_JSON_START==="
$Results | ConvertTo-Json -Depth 4
Write-Output "===RESULT_JSON_END==="

if ($Results | Where-Object { $_.OverallResult -like 'FAILED*' }) {
    exit 1
} else {
    exit 0
}