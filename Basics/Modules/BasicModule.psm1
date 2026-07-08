function RequestName{
    $Name = Read-Host "What is your name?"
    Write-Output "Your name is : ${Name}"

}

function RequestMdp{
    $MDP = Read-Host "Enter your password" -AsSecureString
    $MyClearPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($MDP))
    Write-Output "Your password is: ${MyClearPass}"
}

function TestInternet{
    Test-Connection -ComputerName 8.8.8.8 -ErrorAction SilentlyContinue
    if ($? -eq $true) {
        Write-Output "Internet: OK"
    }
    else {
        Write-Output "Internet: Error"
    }
}

function File{
    $search= Read-Host 'De quel document voulez vous voir le contenu?'
    $searchpath = Join-Path $PSScriptRoot $search
    get-content $searchpath
}