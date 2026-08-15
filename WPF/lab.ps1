Add-Type -AssemblyName 'presentationframework'
[xml]$MonXAML = get-content -path "lab.xaml"

$reader = New-Object System.Xml.XmlNodeReader $MonXAML
$window = [Windows.Markup.XamlReader]::Load($reader)

$computerName = $window.FindName("ComputerName")
$button = $window.FindName("TestButton")
$result = $window.FindName("Result")

$button.Add_Click({
    $machine = $computerName.Text

    if (Test-Connection -ComputerName $machine -Count 1 -Quiet) {
        $result.Text = "$machine est accessible."
    }
    else {
        $result.Text = "$machine n'est pas accessible"
    }
})

$window.ShowDialog()