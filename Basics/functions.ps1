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
    get-content C:\Users\adrie\Desktop\$search
}

function Menu{
    Write-Output "Bienvenu dans le menu du script de fonctions powershell :"
    $Choice = Read-Host 'Alors, on fait quoi ? '
    switch ($Choice) {
        1 {File}
        2 {TestInternet}
        3 {RequestMdp}
        4 {RequestName}
        Default {Menu}
    }
}

# Chargez le script dans la session actuelle
. .\fonctionsps.ps1

# Appelez la fonction avec les paramètres appropriés
Menu