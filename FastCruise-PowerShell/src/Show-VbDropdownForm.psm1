function Show-VbDropdownForm
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,
        [Parameter(Position = 1)]
        [string]$InputFile = "C:\Users\erika\OneDrive\Documents\GitHub\ITPS.OMCS.FastCruise\Configfiles\computerlocation.json",  # Path to JSON or text file
        [Parameter(Position = 2)]
        [string]$TitleBar = 'Fast Cruise'
    )

    Add-Type -AssemblyName Microsoft.VisualBasic

    # Read options from file
    if ($InputFile -like '*.json') {
        $options = Get-Content -Raw -Path $InputFile | ConvertFrom-Json
        if ($options -is [System.Collections.IDictionary]) {
            $options = $options.PSObject.Properties.Name
        }
        elseif ($options -is [System.Collections.IEnumerable]) {
            $options = @($options)
        }
    } else {
        $options = Get-Content -Path $InputFile
    }

    # Create a Windows Form with a ComboBox
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $TitleBar
    $form.Size = New-Object System.Drawing.Size(400,150)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,20)
    $form.Controls.Add($label)

    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(10,50)
    $comboBox.Size = New-Object System.Drawing.Size(360,20)
    $comboBox.DropDownStyle = 'DropDownList'
    $comboBox.Items.AddRange($options)
    $form.Controls.Add($comboBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(300,80)
    $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK })
    $form.Controls.Add($okButton)

    $form.AcceptButton = $okButton

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $comboBox.SelectedItem
    } else {
        return $null
    }
}

Export-ModuleMember -Function Show-VbDropdownForm