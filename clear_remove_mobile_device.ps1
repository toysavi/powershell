param(
    [Parameter(Mandatory = $true)]
    [string]$UserSAM,

    [Parameter(Mandatory = $true)]
    [string]$DeviceId
)

#====================================================
# CONFIGURATION
#====================================================

$ExchangeServer = "EXCH01.yourdomain.local"

$LogFolder = "C:\Logs"
$LogFile = Join-Path $LogFolder "Exchange_MobileDevice_Cleanup.log"

#====================================================
# CREATE LOG FOLDER
#====================================================

if (!(Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

function Write-Log {

    param([string]$Message)

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    Add-Content -Path $LogFile -Value "$Time | $Message"

}

Write-Log "====================================================="
Write-Log "Starting Mobile Device Cleanup"
Write-Log "User      : $UserSAM"
Write-Log "Device ID : $DeviceId"

#====================================================
# CONNECT TO EXCHANGE
#====================================================

try {

    Write-Log "Connecting to Exchange..."

    $Session = New-PSSession `
        -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "http://$ExchangeServer/PowerShell/" `
        -Authentication Kerberos `
        -ErrorAction Stop

    Import-PSSession $Session `
        -DisableNameChecking `
        -AllowClobber | Out-Null

    Get-Command Get-MobileDevice -ErrorAction Stop | Out-Null

    Write-Log "Connected to Exchange."

}
catch {

    Write-Log "ERROR: Unable to connect to Exchange."

    Write-Log $_.Exception.Message

    exit 1

}

#====================================================
# PROCESS DEVICE
#====================================================

try {

    Write-Log "Searching for mobile device..."

    $Device = Get-MobileDevice -Mailbox $UserSAM -ErrorAction Stop |
        Where-Object { $_.DeviceId -eq $DeviceId }

    if (!$Device) {

        Write-Log "WARNING: Device not found."

        return

    }

    $Identity = $Device.Identity

    Write-Log "Device found."

    #==========================================
    # DEVICE INFORMATION
    #==========================================

    try {

        $Stats = Get-MobileDeviceStatistics `
            -Identity $Identity `
            -ErrorAction Stop

        Write-Log "Model            : $($Stats.DeviceModel)"
        Write-Log "Device OS        : $($Stats.DeviceOS)"
        Write-Log "Device Type      : $($Stats.DeviceType)"
        Write-Log "User Agent       : $($Stats.DeviceUserAgent)"
        Write-Log "First Sync       : $($Stats.FirstSyncTime)"
        Write-Log "Last Sync        : $($Stats.LastSuccessSync)"
        Write-Log "Wipe Status      : $($Stats.DeviceWipeStatus)"

    }
    catch {

        Write-Log "Unable to retrieve device statistics."

    }

    #==========================================
    # ACCOUNT ONLY WIPE
    #==========================================

    Write-Log "Sending Account-Only Remote Wipe..."

    Clear-MobileDevice `
        -Identity $Identity `
        -AccountOnly `
        -Confirm:$false `
        -ErrorAction Stop

    Write-Log "Account-only wipe command sent."

    Start-Sleep -Seconds 5

    #==========================================
    # VERIFY WIPE STATUS
    #==========================================

    try {

        $Check = Get-MobileDeviceStatistics -Identity $Identity

        Write-Log "Current Wipe Status : $($Check.DeviceWipeStatus)"

    }
    catch {

        Write-Log "Unable to verify wipe status."

    }

    #==========================================
    # REMOVE DEVICE PARTNERSHIP
    #==========================================

    Write-Log "Removing mobile device partnership..."

    Remove-MobileDevice `
        -Identity $Identity `
        -Confirm:$false `
        -ErrorAction Stop

    Write-Log "Device partnership removed."

    #==========================================
    # REFRESH ACTIVESYNC
    #==========================================

    try {

        Write-Log "Refreshing ActiveSync..."

        Set-CASMailbox `
            -Identity $UserSAM `
            -ActiveSyncEnabled $false `
            -ErrorAction Stop

        Start-Sleep -Seconds 2

        Set-CASMailbox `
            -Identity $UserSAM `
            -ActiveSyncEnabled $true `
            -ErrorAction Stop

        Write-Log "ActiveSync refreshed."

    }
    catch {

        Write-Log "WARNING: Unable to refresh ActiveSync."

        Write-Log $_.Exception.Message

    }

    Write-Log "SUCCESS: Device cleanup completed."

}
catch {

    Write-Log "ERROR during device cleanup."

    Write-Log $_.Exception.Message

}
finally {

    if ($Session) {

        Remove-PSSession $Session

        Write-Log "Exchange PowerShell session closed."

    }

    Write-Log "Completed."
    Write-Log "====================================================="

}