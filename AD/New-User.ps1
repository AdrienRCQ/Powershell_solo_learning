function New-Password
{

$Alphabets = 'a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z'
$numbers = 0..9
$specialCharacters = '~,!,@,#,$,%,^,&,*,(,),>,<,?,\,/,_,-,=,+'
$array = @()
$array += $Alphabets.Split(',') | Get-Random -Count 4
$array[0] = $array[0].ToUpper()
$array[-1] = $array[-1].ToUpper()
$array += $numbers | Get-Random -Count 3
$array += $specialCharacters.Split(',') | Get-Random -Count 3
($array | Get-Random -Count $array.Count) -join ""
}
function New-RandomUser {
<#
.SYNOPSIS
Generate random user data from Https://randomuser.me/.
.DESCRIPTION
This function uses the free API for generating random user data from https://randomuser.me/
.EXAMPLE
Get-RandomUser 10
.EXAMPLE
Get-RandomUser -Amount 25 -Nationality us,gb 
.LINK
https://randomuser.me/
#>
[CmdletBinding()]
param (
[Parameter(Position = 0)]
[ValidateRange(1,500)]
[int] $Amount,

[Parameter()]
[ValidateSet('Male','Female')]
[string] $Gender,

# Supported nationalities: AU, BR, CA, CH, DE, DK, ES, FI, FR, GB, IE, IR, NL, NZ, TR, US
[Parameter()]
[string[]] $Nationality,


[Parameter()]
[ValidateSet('json','csv','xml')]
[string] $Format = 'json',

# Fields to include in the results.
# Supported values: gender, name, location, email, login, registered, dob, phone, cell, id, picture, nat
[Parameter()]
[string[]] $IncludeFields,

# Fields to exclude from the the results.
# Supported values: gender, name, location, email, login, registered, dob, phone, cell, id, picture, nat
[Parameter()]
[string[]] $ExcludeFields
)

$rootUrl = "http://api.randomuser.me/?format=$($Format)"

if ($Amount) {
$rootUrl += "&results=$($Amount)"
}

if ($Gender) {
$rootUrl += "&gender=$($Gender)"
}


if ($Nationality) {
$rootUrl += "&nat=$($Nationality -join ',')"
}

if ($IncludeFields) {
$rootUrl += "&inc=$($IncludeFields -join ',')"
}

if ($ExcludeFields) {
$rootUrl += "&exc=$($ExcludeFields -join ',')"
}

Invoke-RestMethod -Uri $rootUrl
}

$fqdn = Get-ADDomain
$fulldomain = $fqdn.DNSRoot
$domain = $fulldomain.Split(".")
$Dom = $domain[0]
$EXT  = $domain[1]

New-ADOrganizationalUnit -Name "New User Demo" -ProtectedFromAccidentalDeletion $false

$utilisateurs = New-RandomUser -Amount 10 -Nationality FR -IncludeFields name,dob,phone,cell -ExcludeFields picture | Select-Object -ExpandProperty results

foreach ($user in $utilisateurs) {

    $userPassword = New-Password

            $newUserProperties = @{
            Name = "$($user.name.first) $($user.name.last)"
            City = "Lyon" # a modifier au besoin
            GivenName = $user.name.first
            Surname = $user.name.last
            Path = "OU=New User Demo,dc=$Dom,dc=$EXT"  # a modifier au besoin
            title = "Employees"  # a modifier au besoin
            department="Test"
            OfficePhone = $user.phone
            MobilePhone = $user.cell
            Company="$Dom"
            EmailAddress="$($user.name.first).$($user.name.last)@$($fulldomain)"
            AccountPassword = (ConvertTo-SecureString $userPassword -AsPlainText -Force)
            SamAccountName = $($user.name.first).Substring(0,1)+$($user.name.last)
            UserPrincipalName = "$(($user.name.first).Substring(0,1)+$($user.name.last))@$($fulldomain)"
            Enabled = $true
        } 
        # Pour detecter tout problème on utilse un Try Catch pour capturer les erreurs
        Try
        { New-ADUser @newUserProperties     
        } 
        catch{}

}

