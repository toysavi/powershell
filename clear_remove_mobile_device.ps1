param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

# =========================
# Configuration
# =========================

$ExchangeServer = "EXCH01.yourdomain.local"

$LogFolder = "C:\Logs"
$LogFile = Join-Path $LogFolder "Exchange_MobileDevice_Cleanup.log"

# =========================
# Logging
# =========================

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log {
    param([string]$Message)

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$Time | $Message"
}

# =========================
# Connect to Exchange
# =========================

try {

    Write-Log "Connecting to Exchange server $ExchangeServer..."

    $Session = New-PSSession `
        -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "http://$ExchangeServer/PowerShell/" `
        -Authentication Kerberos `
        -ErrorAction Stop

    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

    Get-Command Get-MobileDevice -ErrorAction Stop | Out-Null

    Write-Log "Connected successfully."

}
catch {

    Write-Log "ERROR: Failed to connect to Exchange."
    Write-Log $_.Exception.Message
    exit 1

}

# =========================
# Validate CSV
# =========================

if (!(Test-Path $CsvPath)) {

    Write-Log "ERROR: CSV not found: $CsvPath"

    if ($Session) {
        Remove-PSSession $Session
    }

    exit 1
}

$Rows = Import-Csv $CsvPath

if (!$Rows) {

    Write-Log "ERROR: CSV is empty."

    if ($Session) {
        Remove-PSSession $Session
    }

    exit 1
}

Write-Log "CSV Loaded. Total records: $($Rows.Count)"

# =========================
# Process Devices
# =========================

foreach ($Row in $Rows) {

    $UserSAM = $Row.sAMAccountName
    $DeviceId = $Row.DeviceId
    $Reason = $Row.Reason

    Write-Log "------------------------------------------------------"
    Write-Log "Processing User=$UserSAM DeviceId=$DeviceId"

    try {

        $Devices = Get-MobileDevice -Mailbox $UserSAM -ErrorAction Stop

        $Device = $Devices | Where-Object {
            $_.DeviceId -eq $DeviceId
        }

        if (!$Device) {

            Write-Log "WARNING: Device not found."

            continue
        }

        $Identity = $Device.Identity

        $Stats = Get-MobileDeviceStatistics -Identity $Identity -ErrorAction Stop

        Write-Log "Device Model      : $($Stats.DeviceModel)"
        Write-Log "Device OS         : $($Stats.DeviceOS)"
        Write-Log "Device Type       : $($Stats.DeviceType)"
        Write-Log "Last Sync         : $($Stats.LastSuccessSync)"
        Write-Log "Reason            : $Reason"

        # ---------------------------------
        # Account-only wipe
        # ---------------------------------

        Clear-MobileDevice `
            -Identity $Identity `
            -AccountOnly `
            -Confirm:$false `
            -ErrorAction Stop

        Write-Log "Account-only wipe command submitted."

        Start-Sleep -Seconds 3

        try {

            $Check = Get-MobileDeviceStatistics -Identity $Identity

            Write-Log "Wipe Status       : $($Check.DeviceWipeStatus)"

        }
        catch {

            Write-Log "Unable to read wipe status."

        }

        # ---------------------------------
        # Remove partnership
        # ---------------------------------

        Remove-MobileDevice `
            -Identity $Identity `
            -Confirm:$false `
            -ErrorAction Stop

        Write-Log "SUCCESS: Device removed."

        # ---------------------------------
        # Flush ActiveSync token
        # ---------------------------------

        try {

            Set-CASMailbox `
                -Identity $UserSAM `
                -ActiveSyncEnabled $false `
                -ErrorAction Stop

            Start-Sleep -Seconds 2

            Set-CASMailbox `
                -Identity $UserSAM `
                -ActiveSyncEnabled $true `
                -ErrorAction Stop

            Write-Log "ActiveSync cache refreshed."

        }
        catch {

            Write-Log "WARNING: Unable to refresh ActiveSync."

        }

    }
    catch {

        Write-Log "ERROR processing $UserSAM"

        Write-Log $_.Exception.Message

    }

}

# =========================
# Cleanup
# =========================

if ($Session) {

    Remove-PSSession $Session

    Write-Log "Exchange session closed."

}

Write-Log "========== Script Completed =========="