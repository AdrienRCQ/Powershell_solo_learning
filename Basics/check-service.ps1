param(
    [string]$ServiceName
)
Get-Service -Name $ServiceName
if ((Get-Service -Name $ServiceName).Status -ne 'Running') {
    Start-Service -Name $ServiceName
    Write-Host "Service $ServiceName strated !"
}