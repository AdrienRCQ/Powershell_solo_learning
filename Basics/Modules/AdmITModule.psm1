function CheckService {
    param(
    [string]$ServiceName
)
Get-Service -Name $ServiceName
if ((Get-Service -Name $ServiceName).Status -ne 'Running') {
    Start-Service -Name $ServiceName
    Write-Host "Service $ServiceName strated !"
}
}


function Get-DiskAlerte {
    $disks = Get-Volume | Where-Object { $_.DriveType -eq "Fixed" }
    foreach ($disk in $disks) {
        $freeSpacePercent = [math]::Round(($disk.SizeRemaining / $disk.Size) * 100, 2)
        if ($freeSpacePercent -lt 10) {
            Write-Log -message "ALERTE : Espace disque faible sur $($disk.DriveLetter) ($freeSpacePercent%)"
        }
    }

}
