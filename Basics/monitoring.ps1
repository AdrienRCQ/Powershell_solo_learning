function Get-DiskAlerte {
    $disks = Get-Volume | Where-Object { $_.DriveType -eq "Fixed" }
    foreach ($disk in $disks) {
        $freeSpacePercent = [math]::Round(($disk.SizeRemaining / $disk.Size) * 100, 2)
        if ($freeSpacePercent -lt 10) {
            Write-Log -message "ALERTE : Espace disque faible sur $($disk.DriveLetter) ($freeSpacePercent%)"
        }
    }

}




