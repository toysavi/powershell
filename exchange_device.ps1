param(
    [string]$Users,
    [string]$DeviceIds
)

$log = "C:\Logs\awx-clean-mobiledevices.log"
function Log($msg) {
    Add-Content $log "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
}

# --- Load Exchange snap-in in THIS session (this was missing) ---
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        Log "Snap-in loaded successfully."
    } catch {
        Log "FATAL: Could not load Exchange snap-in: $($_.Exception.Message)"
        exit 1
    }
}

if (-not (Get-Command Get-MobileDevice -ErrorAction SilentlyContinue)) {
    Log "FATAL: Get-MobileDevice not available after snap-in load."
    exit 1
}

# --- Clean split, trim, drop blanks ---
$UserList     = $Users     -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
$InputDevices = $DeviceIds -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

Log "START PROCESS"
Log "Users to process: $($UserList -join ', ')"
Log "Device filter: $(if ($InputDevices.Count -gt 0) { $InputDevices -join ', ' } else { 'NONE (all devices)' })"

foreach ($User in $UserList) {
    Log "USER: $User"

    try {
        $Devices = Get-MobileDevice -Mailbox $User -ErrorAction Stop
    } catch {
        Log "ERROR: Get-MobileDevice failed for $User -> $($_.Exception.Message)"
        continue
    }

    if (-not $Devices -or $Devices.Count -eq 0) {
        Log "No mobile devices found for $User"
        continue
    }

    foreach ($Device in $Devices) {
        $id = $Device.DeviceId

        if ($InputDevices.Count -gt 0) {
            # Normalize both sides (strip dashes, uppercase) before comparing
            $normId     = ($id -replace '-', '').ToUpper()
            $normFilter = $InputDevices | ForEach-Object { ($_ -replace '-', '').ToUpper() }

            if ($normFilter -notcontains $normId) {
                Log "SKIP (not in filter): $id"
                continue
            }
        }

        try {
            Clear-MobileDevice -Identity $id -AccountOnly -ErrorAction Stop
            Log "WIPED: $id"
        } catch {
            Log "ERROR wiping $id -> $($_.Exception.Message)"
            continue
        }

        try {
            Remove-MobileDevice -Identity $id -Confirm:$false -ErrorAction Stop
            Log "REMOVED: $id"
        } catch {
            Log "ERROR removing $id -> $($_.Exception.Message)"
        }

        try {
            Set-CASMailbox -Identity $User -ActiveSyncBlockedDeviceIDs @{add=$id} -ErrorAction Stop
            Log "BLOCKED: $id"
        } catch {
            Log "ERROR blocking $id on $User -> $($_.Exception.Message)"
        }
    }
}
Log "END PROCESS"