Import-Module ADModule

$fqdn = Get-ADDomain
$fulldomain = $fqdn.DNSRoot
$domain = $fulldomain.Split(".")
$Dom = $domain[0]
$EXT = $domain[1]

New-ADOrganizationalUnit -Name "New User Demo" -ProtectedFromAccidentalDeletion $false

$utilisateurs = New-RandomUser -Amount 10 -Nationality FR -IncludeFields name, dob, phone, cell -ExcludeFields picture | Select-Object -ExpandProperty results

foreach ($user in $utilisateurs) {

    $userPassword = New-Password

    $newUserProperties = @{
        Name              = "$($user.name.first) $($user.name.last)"
        City              = "Lyon" # a modifier au besoin
        GivenName         = $user.name.first
        Surname           = $user.name.last
        Path              = "OU=New User Demo,dc=$Dom,dc=$EXT"  # a modifier au besoin
        title             = "Employees"  # a modifier au besoin
        department        = "Test"
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

}

