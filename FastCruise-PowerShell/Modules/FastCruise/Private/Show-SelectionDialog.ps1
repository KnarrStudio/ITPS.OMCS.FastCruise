function Show-SelectionDialog
{
    <#
        .SYNOPSIS
        Prompts the user to pick one value from a list.

        .DESCRIPTION
        Displays a Windows Forms dialog with a label, a drop-down list of
        choices, and OK/Cancel buttons. This single function replaces every
        Out-GridView picker and the old Show-VbDropdownForm prototype so all
        FastCruise selection prompts (department, building, room, desk,
        application-test result, exit-strategy menu) share one look and one
        implementation.

        .PARAMETER Message
        Prompt text shown above the drop-down.

        .PARAMETER Options
        The list of choices to offer.

        .PARAMETER Title
        Dialog window title. Defaults to 'Fast Cruise'.

        .EXAMPLE
        Show-SelectionDialog -Message 'Select Department:' -Options @('Sales', 'Shipping')

        .OUTPUTS
        System.String. The selected option, or an empty string if the dialog
        was cancelled or no option was chosen.
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Message,
        [Parameter(Mandatory, Position = 1)]
        [Object[]]$Options,
        [Parameter(Position = 2)]
        [String]$Title = 'Fast Cruise'
    )
    Write-FCLog -Message ('Prompting for selection: {0}' -f $Message) -Level Debug

    if ($Options.Count -eq 0)
    {
        Write-FCLog -Message 'Show-SelectionDialog called with no options to choose from.' -Level Warning
        return ''
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $Form = [System.Windows.Forms.Form]::new()
    $Form.Text = $Title
    $Form.Size = [System.Drawing.Size]::new(420, 150)
    $Form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'FixedDialog'
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    $Form.TopMost = $true

    $Label = [System.Windows.Forms.Label]::new()
    $Label.Text = $Message
    $Label.AutoSize = $true
    $Label.MaximumSize = [System.Drawing.Size]::new(380, 0)
    $Label.Location = [System.Drawing.Point]::new(10, 15)
    $Form.Controls.Add($Label)

    $ComboBox = [System.Windows.Forms.ComboBox]::new()
    $ComboBox.Location = [System.Drawing.Point]::new(10, 55)
    $ComboBox.Size = [System.Drawing.Size]::new(380, 20)
    $ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $null = $ComboBox.Items.AddRange($Options)
    $ComboBox.SelectedIndex = 0
    $Form.Controls.Add($ComboBox)

    $OkButton = [System.Windows.Forms.Button]::new()
    $OkButton.Text = 'OK'
    $OkButton.Location = [System.Drawing.Point]::new(230, 85)
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($OkButton)
    $Form.AcceptButton = $OkButton

    $CancelButton = [System.Windows.Forms.Button]::new()
    $CancelButton.Text = 'Cancel'
    $CancelButton.Location = [System.Drawing.Point]::new(315, 85)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.Controls.Add($CancelButton)
    $Form.CancelButton = $CancelButton

    $Result = $Form.ShowDialog()
    $Form.Dispose()

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK -and $null -ne $ComboBox.SelectedItem)
    {
        return $ComboBox.SelectedItem.ToString()
    }
    return ''
}
