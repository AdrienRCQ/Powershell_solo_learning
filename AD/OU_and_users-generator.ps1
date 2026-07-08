Import-Module ADModule

Import-Module -Name  PSWriteWord -Force

$fqdn = Get-ADDomain
$fulldomain = $fqdn.DNSRoot
$domain = $fulldomain.Split(".")
$Dom = $domain[0] 
$Dom1 = $fulldomain.Split(".")[0]
$EXT = $domain[1]
$EXT1 = $(Get-ADDomain).DNSRoot.Split(".")[1]

write-output $fqdn
write-output $fulldomain 
write-output $domain 
write-output $Dom  
write-output $Dom1 
write-output $EXT  
write-output $EXT1 

$OUpremierNiveau = 'Site Lyon'
$OUsecondNiveaux = @('Informaticiens', 'Clients Entreprises', 'Services', 'Groupes')
$Services = @('Direction', 'Comptabilite', 'Commerciale')
$Informatiques = @('Techniciens informatiques', 'Administrateurs informatiques')
$Groupes = @('Groupes globaux', 'Groupes domaine locaux')


function CreateADArbo {

    New-ADOrganizationalUnit -Name $OUpremierNiveau -ProtectedFromAccidentalDeletion $false

    foreach ($OU in $OUsecondNiveaux) {


        New-ADOrganizationalUnit -Name $OU -Path "Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -ProtectedFromAccidentalDeletion $false

        switch ($OU) {
            'Informaticiens' {
            
                foreach ($item in $Informatiques) {

                    New-ADOrganizationalUnit -Name $item -Path "OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -ProtectedFromAccidentalDeletion $false
                
                    switch ($item) {
                        'Techniciens informatiques' { $utilisateurs = New-RandomUser -Amount 15 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results }
                        'Administrateurs informatiques' { $utilisateurs = New-RandomUser -Amount 16 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results } 
                        Default {}
                    }
                    
                    foreach ($user in $utilisateurs) {

                        $userPassword = New-Password
                    
                        $newUserProperties = @{
                            Name              = "$($user.name.first) $($user.name.last)"
                            City              = "Poitiers" # a modifier au besoin
                            GivenName         = $user.name.first
                            Surname           = $user.name.last
                            Path              = "OU=$Item,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT"  # a modifier au besoin
                            title             = "Employees"  # a modifier au besoin
                            department        = "$item"
                            OfficePhone       = $user.phone
                            MobilePhone       = $user.cell
                            Company           = "$Dom"
                            EmailAddress      = "$($user.name.first).$($user.name.last)@$($fulldomain)"
                            AccountPassword   = (ConvertTo-SecureString $userPassword -AsPlainText -Force)
                            SamAccountName    = $($user.name.first).Substring(0, 1) + $($user.name.last)
                            UserPrincipalName = "$(($user.name.first).Substring(0,1)+$($user.name.last))@$($fulldomain)"
                            Enabled           = $true
                        } 
                        # Pour detecter tout problème on utilse un Try Catch pour capturer les erreurs
                        Try { 
                            New-ADUser @newUserProperties     
                        } 
                        catch {}

                            
                        if (!(Test-Path -Path "C:\DocumentsUser\Employes\$OU\$Item")) {
                            New-Item -Path "C:\DocumentsUser\Employes\$OU\$Item" -ItemType Directory | Out-Null
                        }
                        else {
                            #"The directory exist" 
                        }

                        #Emplacement du Template 
                        $FilePathTemplate = "C:\Template\template.docx"

                        $WordDocument = Get-WordDocument -FilePath $FilePathTemplate

                        $FilePathInvoice = "C:\DocumentsUser\Employes\$Ou\$Item\$($user.name.last) $($user.name.first).docx"
                        Add-WordText -WordDocument $WordDocument -Text 'Creation de Compte' -FontSize 15 -HeadingType Heading1 -FontFamily 'Arial' -Italic $true | Out-Null


                        Add-WordText -WordDocument $WordDocument -Text 'Voici les informations qui vous permettrons de vous connecter au Domaine Active Directory', " $fulldomain" `
                            -FontSize 12, 13 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -SpacingBefore 15 `
                            -Supress $True

                        Add-WordText -WordDocument $WordDocument -Text 'Login : ', "$(($user.name.first).Substring(0,1)+$($user.name.last))" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -Supress $True

                        Add-WordText -WordDocument $WordDocument -Text 'Mot de passe : ', "$userPassword" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -Supress $True
                        Add-WordText -WordDocument $WordDocument -Text 'Adresse de messagerie : ', "$($user.name.first).$($user.name.last)@$($fulldomain)" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -SpacingAfter 15 `
                            -Supress $True

                        Add-WordText -WordDocument $WordDocument -Text "Le Service Informatique." `
                            -FontSize 12 `
                            -Supress $True
                    
                        Add-WordProtection -WordDocument $WordDocument -EditRestrictions readOnly -Password 'P@$$Sio123*'

                        Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice -Supress $true -Language 'fr-FR'

                    
                    }
        
                }
                

            }
            'Services' {
                
                foreach ($Serv in $Services) {

                    New-ADOrganizationalUnit -Name $Serv -Path "OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -ProtectedFromAccidentalDeletion $false
                
                    switch ($Serv) {
                        'Direction' { $utilisateurs = New-RandomUser -Amount 15 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results }
                        'Comptabilite' { $utilisateurs = New-RandomUser -Amount 10 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results }
                        'Commerciale' { $utilisateurs = New-RandomUser -Amount 20 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results }
                        Default {}
                    }

                    foreach ($user in $utilisateurs) {

                        $userPassword = New-Password
                    
                        $newUserProperties = @{
                            Name              = "$($user.name.first) $($user.name.last)"
                            City              = "Poitiers" # a modifier au besoin
                            GivenName         = $user.name.first
                            Surname           = $user.name.last
                            Path              = "OU=$Serv,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT"  # a modifier au besoin
                            title             = "Employees"  # a modifier au besoin
                            department        = "$Serv"
                            OfficePhone       = $user.phone
                            MobilePhone       = $user.cell
                            Company           = "$Dom"
                            EmailAddress      = "$($user.name.first).$($user.name.last)@$($fulldomain)"
                            AccountPassword   = (ConvertTo-SecureString $userPassword -AsPlainText -Force)
                            SamAccountName    = $($user.name.first).Substring(0, 1) + $($user.name.last)
                            UserPrincipalName = "$(($user.name.first).Substring(0,1)+$($user.name.last))@$($fulldomain)"
                            Enabled           = $true
                        } 
                        # Pour detecter tout problème on utilse un Try Catch pour capturer les erreurs
                        Try { 
                            New-ADUser @newUserProperties     
                        } 
                        catch {}
                        if (!(Test-Path -Path "C:\DocumentsUser\Employes\$OU\$Serv")) {
                            New-Item -Path "C:\DocumentsUser\Employes\$OU\$Serv" -ItemType Directory | Out-Null
                        }
                        else {
                            #"The directory exist" 
                        }
        
                        #Emplacement du Template 
                        $FilePathTemplate = "C:\Template\template.docx"
        
                        $WordDocument = Get-WordDocument -FilePath $FilePathTemplate
        
                        $FilePathInvoice = "C:\DocumentsUser\Employes\$OU\$Serv\$($user.name.last) $($user.name.first).docx"
                        Add-WordText -WordDocument $WordDocument -Text 'Creation de Compte' -FontSize 15 -HeadingType Heading1 -FontFamily 'Arial' -Italic $true  | Out-Null
        
        
                        Add-WordText -WordDocument $WordDocument -Text 'Voici les informations qui vous permettrons de vous connecter au domaine Active Directory', " $fulldomain" `
                            -FontSize 12, 13 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -SpacingBefore 15 `
                            -Supress $True
        
                        Add-WordText -WordDocument $WordDocument -Text 'Login : ', "$(($user.name.first).Substring(0,1)+$($user.name.last))" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -Supress $True
        
                        Add-WordText -WordDocument $WordDocument -Text 'Mot de passe : ', "$userPassword" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -Supress $True
                        Add-WordText -WordDocument $WordDocument -Text 'Adresse de messagerie : ', "$($user.name.first).$($user.name.last)@$($fulldomain)" `
                            -FontSize 12, 10 `
                            -Color Black, Blue `
                            -Bold $false, $true `
                            -SpacingAfter 15 `
                            -Supress $True
        
                        Add-WordText -WordDocument $WordDocument -Text "Le Service Informatique." `
                            -FontSize 12 `
                            -Supress $True
                        
                        Add-WordProtection -WordDocument $WordDocument -EditRestrictions readOnly -Password 'P@$$Sio123*'
        
                        Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice -Supress $true -Language 'fr-FR'
        
                    
                    }


                }

            }
            'Groupes' {

                foreach ($Grp in $Groupes) {

                    New-ADOrganizationalUnit -Name $Grp -Path "OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -ProtectedFromAccidentalDeletion $false

                    switch ($Grp) {
                        'Groupes globaux' { 
                            
                            foreach ($item in $Services) {

                                $item = $item.replace(" ", "_")

                                New-ADGroup -Name "Gg_$item" -DisplayName "G_$item" -GroupScope Global -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Global $item"
                            }
                            foreach ($item in $Informatiques) {

                                $item = $item.replace(" ", "_")

                                New-ADGroup -Name "Gg_$item" -DisplayName "G_$item" -GroupScope Global -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Global $item"
                            }

                        }
                        'Groupes domaine locaux' { 

                            foreach ($item in $Services) {
                                
                                $item = $item.replace(" ", "_")

                                New-ADGroup -Name "DL_$item`_L" -DisplayName "DL_$item`_L" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item Lecture"
                                New-ADGroup -Name "DL_$item`_LM" -DisplayName "DL_$item`_LM" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item Lecture et modification"
                                New-ADGroup -Name "DL_$item`_CT" -DisplayName "DL_$item`_CT" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item controle totale"

                            }
                            foreach ($item in $Informatiques) {
                                
                                $item = $item.replace(" ", "_")

                                New-ADGroup -Name "DL_$item`_L" -DisplayName "DL_$item`_L" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item Lecture"
                                New-ADGroup -Name "DL_$item`_LM" -DisplayName "DL_$item`_LM" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item Lecture et modification"
                                New-ADGroup -Name "DL_$item`_CT" -DisplayName "DL_$item`_CT" -GroupScope DomainLocal -GroupCategory Security -Path "OU=$Grp,OU=$OU,Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT" -Description "Groupe Domaine local $item controle totale"

                            }

                        }
                        Default {}
                    }
                }

            }
            Default {}
        }
        
    }

    $Users = Get-ADUser -Filter 'department -like "*"' -Properties * -SearchBase "Ou=$OUpremierNiveau,dc=$Dom,dc=$EXT"

    foreach ($item in $Users) {
        
        switch ($item.department) {
            'Techniciens informatiques' { $item | Add-ADPrincipalGroupMembership  -MemberOf "Gg_Techniciens_informatiques" }
            'Administrateurs informatiques' { $item | Add-ADPrincipalGroupMembership  -MemberOf "Gg_Administrateurs_informatiques" }
            'Direction' { $item | Add-ADPrincipalGroupMembership  -MemberOf "Gg_Direction" }
            'Comptabilite' { $item | Add-ADPrincipalGroupMembership  -MemberOf "Gg_Comptabilite" }
            'Commerciale' { $item | Add-ADPrincipalGroupMembership  -MemberOf "Gg_Commerciale" }
            Default {}
        }

    }

    $chemin = "Commun"
    $Drive = "C:\"
    $cheminF = $drive + $chemin 
    New-Item -path "$cheminf" -ItemType directory | Out-Null

    foreach ($item in $Services) {
        Set-Location $cheminF | Out-Null
        New-Item -Name $item.Replace(" ", "_") -ItemType directory | Out-Null
        Set-Location $cheminF | Out-Null

        #Ajout de la securite NTFS
        $item2 = $item.Replace(" ", "_")
        $chemin = "$cheminF" + "\$item2"
        Set-NTFSInheritance -Path $item2.Replace(" ", "_") -AccessInheritanceEnabled $false -AuditInheritanceEnabled $true 
        $groupeL = "DL_" + $item.Replace(" ", "_") + "_L" 
        $groupeLM = "DL_" + $item.Replace(" ", "_") + "_LM" 
        $groupeCT = "DL_" + $item.Replace(" ", "_") + "_CT" 
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeL" -AccessRights ReadAndExecute -AccessType Allow 
            
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeLM" -AccessRights Modify -AccessType Allow 
                
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeCT" -AccessRights FullControl -AccessType Allow 
                
        Add-NTFSAccess -Path $chemin -Account "$Dom\Administrateur" -AccessRights FullControl -AccessType Allow
        
    }

    foreach ($item in $Informatiques) {
        Set-Location $cheminF | Out-Null
        New-Item -Name $item.Replace(" ", "_") -ItemType directory | Out-Null
        Set-Location $cheminF | Out-Null
    
        #Ajout de la securite NTFS
        $item2 = $item.Replace(" ", "_")
        $chemin = "$cheminF" + "\$item2"
        Set-NTFSInheritance -Path $item2.Replace(" ", "_") -AccessInheritanceEnabled $false -AuditInheritanceEnabled $true 
        $groupeL = "DL_" + $item.Replace(" ", "_") + "_L" 
        $groupeLM = "DL_" + $item.Replace(" ", "_") + "_LM" 
        $groupeCT = "DL_" + $item.Replace(" ", "_") + "_CT" 
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeL" -AccessRights ReadAndExecute -AccessType Allow 
            
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeLM" -AccessRights Modify -AccessType Allow 
                
        Add-NTFSAccess -Path $chemin -Account "$Dom\$groupeCT" -AccessRights FullControl -AccessType Allow 
                
        Add-NTFSAccess -Path $chemin -Account "$Dom\Administrateur" -AccessRights FullControl -AccessType Allow
    }

    #Il nous manque des Groupes de domaine locaux pour le premier dossier
    foreach ($Y in $chemin) {
                                
        $name = "Access_" + $Y
        New-ADGroup -Name "DL_$name`_L" -DisplayName "DL_$name`_L" -GroupScope DomainLocal -GroupCategory Security -Path "OU=Groupes domaine locaux,OU=Groupes,OU=Site Lyon,DC=Tierslieux86,DC=fr" -Description "Groupe Domaine local $item Lecture"
        New-ADGroup -Name "DL_$name`_LM" -DisplayName "DL_$name`_LM" -GroupScope DomainLocal -GroupCategory Security -Path "OU=Groupes domaine locaux,OU=Groupes,OU=Site Lyon,DC=Tierslieux86,DC=fr" -Description "Groupe Domaine local $item Lecture et modification"
        New-ADGroup -Name "DL_$name`_CT" -DisplayName "DL_$name`_CT" -GroupScope DomainLocal -GroupCategory Security -Path "OU=Groupes domaine locaux,OU=Groupes,OU=Site Lyon,DC=Tierslieux86,DC=fr" -Description "Groupe Domaine local $item controle totale"
        
    }
        
    New-SmbShare -Name $chemin -Path $cheminf -ChangeAccess "AUTORITE NT\Utilisateurs authentifies" -Description "Accès $chemin " | Out-Null
    # Ajout de la Securite NTFS pour le dossier Commun
    Add-NTFSAccess -Path $cheminF -Account "$Dom\Administrateur" -AccessRights FullControl -AccessType Allow
    Add-NTFSAccess -Path $cheminF -Account "BUILTIN\Administrateurs" -AccessRights FullControl -AccessType Allow -AppliesTo ThisFolderSubfoldersAndFiles
    Add-NTFSAccess -Path $cheminf -Account "$Dom\DL_Access_Commun_CT" -AccessRights FullControl -AccessType Allow -AppliesTo ThisFolderSubfoldersAndFiles
    Add-NTFSAccess -Path $cheminf -Account "$Dom\DL_Access_Commun_LM" -AccessRights Modify -AccessType Allow -AppliesTo ThisFolderSubfoldersAndFiles
    Add-NTFSAccess -Path $cheminf -Account "$Dom\DL_Access_Commun_L" -AccessRights ReadAndExecute -AccessType Allow -AppliesTo ThisFolderSubfoldersAndFiles

    # Important il faut supprimer l'heritage.
    Set-NTFSInheritance -Path $cheminF -AccessInheritanceEnabled $false -AuditInheritanceEnabled $true 
    

    #Ajout des deux groupes Globaux 
    New-ADGroup -Name "Gg_Employes" -DisplayName "Gg_Employes" -GroupScope Global -GroupCategory Security -Path "OU=Groupes globaux,OU=Groupes,OU=Site Lyon,DC=Tierslieux86,DC=fr" -Description "Groupe Domaine local $item controle totale"
    New-ADGroup -Name "Gg_Employes_informatiques" -DisplayName "Gg_Employes_informatiques" -GroupScope Global -GroupCategory Security -Path "OU=Groupes globaux,OU=Groupes,OU=Site Lyon,DC=Tierslieux86,DC=fr" -Description "Groupe Domaine local $item controle totale"

    #Recherche des Groupes Globaux et groupes de domaine Locaux
    $groupeGL = Get-ADGroup -filter "*" -SearchBase "OU=Groupes globaux,OU=Groupes,OU=Site Lyon,DC=$Dom,DC=$Ext"
    $groupeDL = Get-ADGroup -filter "*" -SearchBase "OU=Groupes domaine locaux,OU=Groupes,OU=Site Lyon,DC=$Dom,DC=$Ext"


    foreach ($gl in $groupeGL) {
        
        switch ($gl) { 
            { (($gl -match "GG_Direction")) } { Add-ADGroupMember -Identity ( Get-ADGroup "Gg_Employes") -Members $gl } 
            { (($gl -match "GG_Commerciale")) } { Add-ADGroupMember -Identity ( Get-ADGroup "Gg_Employes") -Members $gl }
            { (($gl -match "GG_Comptabilite")) } { Add-ADGroupMember -Identity ( Get-ADGroup "Gg_Employes") -Members $gl }

            { (($gl -match "Info") -and ($gl -notmatch "Gg_Employes_informatiques")) } { Add-ADGroupMember -Identity ( Get-ADGroup "Gg_Employes_informatiques") -Members $gl } 
            
            default {}
        }
    }


    foreach ($gl in $groupeGL) {
        foreach ($dl in $groupeDL) {
            switch ($gl) { 
                { (($gl -match "Direc") -And ($dl -match "Direction_LM")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "Direc") -And ($dl -match "Commerciale_LM")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "Direc") -And ($dl -match "Comptabilite_LM")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "Compta") -And ($dl -match "Comptabilite_LM")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "Commer") -And ($dl -match "Commerciale_LM")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "Commer") -And ($dl -match "Comptabilite_L")) } { Add-ADGroupMember -Identity $dl -Members $gl } 
                { (($gl -match "info") -And ($dl -match "CT")) } { Add-ADGroupMember -Identity  $dl -Members $gl }
            
                default {}
            }
        }
    }
}