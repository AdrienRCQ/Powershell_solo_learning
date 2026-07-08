Import-Module BasicModule

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

Menu