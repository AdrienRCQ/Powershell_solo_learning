$DomainNameDNS = "RCQ-HomeLab.local"
$DomainNameNetbios = "RCQ-HOMELAB"

$ForestConfiguration = @{
    '-DatabasePath' = 'C:\Windows\NTDS';
    '-DomainMode' = 'Default';
    '-DomainName' = $DomainNameDNS;
    '-DomainNetbiosName' = $DomainNameNetbios;
    '-ForestMode' = 'Default';
    '-InstallDns' = $true;
    '-LogPath' = 'C:\Windows\NTDS';
    '-NoRebootOnCompletion' = $false;
    '-SysvolPath' = 'C:\Windows\SYSVOL';
    '-Force' = $true;
    '-CreateDnsDelegation' = $false
}

Import-Module ADDSDeployment
Install-ADDSForest @ForestConfiguration