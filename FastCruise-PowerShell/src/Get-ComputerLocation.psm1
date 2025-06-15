function Get-ComputerLocation {
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [String]$jsonFilePath = "C:\Users\erika\OneDrive\Documents\GitHub\ITPS.OMCS.FastCruise\Configfiles\computerlocation.json"
    )

    function Show-LocalDropdown {
        param(
            [string]$Message,
            [array]$Options,
            [string]$TitleBar = 'Fast Cruise'
        )
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
        $comboBox.Items.AddRange($Options)
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

    if (Test-Path -Path $jsonFilePath -ErrorAction SilentlyContinue) {
        Write-Verbose -Message 'Using JSON File'
        $location = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

        $deptNames = $location.Department.PSObject.Properties.Name
        $LclDept = Show-LocalDropdown -Message 'Select Department:' -Options $deptNames -TitleBar 'Department'

        $buildNames = $location.Department.$LclDept.Building.PSObject.Properties.Name
        $LclBuild = Show-LocalDropdown -Message 'Select Building:' -Options $buildNames -TitleBar 'Building'

        $roomList = $location.Department.$LclDept.Building.$LclBuild.Room
        $LclRm = Show-LocalDropdown -Message 'Select Room:' -Options $roomList -TitleBar 'Room'

        $Desk = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q')
        $LclDesk = Show-LocalDropdown -Message 'Select Desk:' -Options $Desk -TitleBar 'Desk'
    }
    else {
        Write-Error "Unable to find or use JSON File: $jsonFilePath"
        # [TAGGED: ManualInputFallback]
        return
    }

    [PSCustomObject]@{
        Department = $LclDept
        Building   = $LclBuild
        Room       = $LclRm
        Desk       = $LclDesk
    }
}

Export-ModuleMember -Function Get-ComputerLocation