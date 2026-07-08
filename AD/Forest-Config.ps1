function ForestConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DomainNameDNS,

        [Parameter(Mandatory)]
        [string]$DomainNameNetbios,

        [Parameter(Mandatory)]
        [SecureString]$SafeModeAdministratorPassword
    )

    try {
        # Vérification du rôle AD DS
        $adDsFeature = Get-WindowsFeature -Name AD-Domain-Services

        if (-not $adDsFeature.Installed) {
            throw "Le rôle 'Active Directory Domain Services' n'est pas installé."
        }

        Import-Module ADDSDeployment -ErrorAction Stop

        $forestParams = @{
            DomainName                    = $DomainNameDNS
            DomainNetbiosName             = $DomainNameNetbios
            SafeModeAdministratorPassword = $SafeModeAdministratorPassword
            InstallDns                    = $true
            CreateDnsDelegation           = $false
            DatabasePath                  = 'C:\Windows\NTDS'
            LogPath                       = 'C:\Windows\NTDS'
            SysvolPath                    = 'C:\Windows\SYSVOL'
            Force                         = $true
            NoRebootOnCompletion          = $false
            ErrorAction                   = 'Stop'
        }

        Install-ADDSForest @forestParams
    }
    catch {
        Write-Error "Échec de la création de la forêt Active Directory : $_"
        throw
    }
}

# Use example :
# $DSRMPassword = ConvertTo-SecureString `
#     "P@ssw0rd123!" `
#     -AsPlainText `
#     -Force

# ForestConfig `
#     -DomainNameDNS "contoso.local" `
#     -DomainNameNetbios "CONTOSO" `
#     -SafeModeAdministratorPassword $DSRMPassword