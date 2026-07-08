$filter = @{
    LogName   = "Security"
    Id        = 4625
    StartTime = (Get-Date).AddHours(-24)
}

$Events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop

$Results = foreach ($Event in $Events) {

    $Xml = [xml]$Event.ToXml()
    $Data = @{}
    foreach($Item in $Xml.Event.EventData.Data){
        $Data[$Item.Name] = $Item.'#text'
    }
    
    [PSCustomObject]@{
        DateHeure      = $Event.TimeCreated
        Utilisateur    = $Data.TargetUserName
        Domaine        = $Data.TargetDomainName
        AdresseIP      = $Data.IpAddress
        PosteSource    = $Data.WorkstationName
        Processus      = $Data.ProcessName
        LogonType      = $Data.LogonType
        Statut         = $Data.Status
        SousStatut     = $Data.SubStatus
    }

}

$Results |
    Sort-Object DateHeure -Descending |
    Format-Table -AutoSize

$CsvPath = Join-Path $PSScriptRoot "Evenements_4625.csv"
$Results |
    Export-Csv $CsvPath -NoTypeInformation -Encoding UTF8