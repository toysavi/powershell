param(
    [string]$Users,
    [string]$DeviceIds
)

$log = "C:\Logs\awx-clean-mobiledevices.log"

function Log($msg){
    Add-Content $log "$(Get-Date) | $msg"
}

$UserList = $Users -split "`n"
$InputDevices = $DeviceIds -split "`n"

Log "START PROCESS"

foreach ($User in $UserList) {

    if ([string]::IsNullOrWhiteSpace($User)) { continue }

    Log "USER: $User"

    # Get all devices for user
    $Devices = Get-MobileDevice -Mailbox $User

    foreach ($Device in $Devices) {

        $id = $Device.DeviceId

        # If device filter provided → restrict
        if ($InputDevices.Count -gt 0 -and $InputDevices[0] -ne "") {
            if ($InputDevices -notcontains $id) {
                continue
            }
        }

        # =========================
        # ALWAYS EXECUTE (NO MODE)
        # =========================

        # 1. Account-only wipe
        Clear-MobileDevice -Identity $id -AccountOnly
        Log "WIPED: $id"

        # 2. Remove device
        Remove-MobileDevice -Identity $id -Confirm:$false
        Log "REMOVED: $id"

        # 3. Block device
        Set-CASMailbox -Identity $User `
            -ActiveSyncBlockedDeviceIDs @{add=$id}

        Log "BLOCKED: $id"
    }
}

Log "END PROCESS"