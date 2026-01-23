$banner = @"
 ________  ________          _________  ________  ________  ___       ________      
|\   __  \|\   ___ \        |\___   ___|\   __  \|\   __  \|\  \     |\   ____\     
\ \  \|\  \ \  \_|\ \       \|___ \  \_\ \  \|\  \ \  \|\  \ \  \    \ \  \___|_    
 \ \   __  \ \  \ \\ \           \ \  \ \ \  \\\  \ \  \\\  \ \  \    \ \_____  \   
  \ \  \ \  \ \  \_\\ \           \ \  \ \ \  \\\  \ \  \\\  \ \  \____\|____|\  \  
   \ \__\ \__\ \_______\           \ \__\ \ \_______\ \_______\ \_______\____\_\  \ 
    \|__|\|__|\|_______|            \|__|  \|_______|\|_______|\|_______|\_________\
                                                                        \|_________|
                                                                                    
                                                                                                                                                                                                                                                                                                       
"@

Write-Host $banner

function Menu{
    Write-Output "Bienvenu dans le menu d'AD TOOLS :"
    Write-Output "
    Génération de l'arbo `n
    Créer un utilisateur `n
    Importer des utilisateurs via CSV `n
    Tester Internet `n
    "
    $Choice = Read-Host 'Alors, on fait quoi ? '

    switch ($Choice) {
        1 {CreateADArbo}
        2 {Write-Output "Pas encore disponible"}
        3 {Write-Output "Pas encore disponible"}
        4 {TestInternet}
        Default {Menu}
    }
}

# Chargez le script dans la session actuelle
. .\script1.ps1
. .\arboAD.ps1

# Appelez la fonction avec les paramètres appropriés
Menu