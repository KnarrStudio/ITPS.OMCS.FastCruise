function Show-InputDialog
{
    <#
        .SYNOPSIS
        Prompts the user for a single line of free-text input.

        .DESCRIPTION
        Displays a small Windows Forms dialog with a label, a text box, and
        OK/Cancel buttons. Replaces the old Microsoft.VisualBasic InputBox so
        that every FastCruise prompt uses the same dialog technology on both
        Windows PowerShell 5.1 and PowerShell 7+.

        .PARAMETER Message
        Prompt text shown above the input box.

        .PARAMETER Title
        Dialog window title. Defaults to 'Fast Cruise'.

        .PARAMETER DefaultValue
        Text pre-filled in the input box.

        .EXAMPLE
        Show-InputDialog -Message 'Nearest Phone Number:' -DefaultValue '757-555-1234'

        .OUTPUTS
        System.String. The text entered, or an empty string if the dialog was
        cancelled.
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Message,
        [Parameter(Position = 1)]
        [String]$Title = 'Fast Cruise',
        [Parameter(Position = 2)]
        [String]$DefaultValue = ''
    )
    Write-FCLog -Message ('Prompting for input: {0}' -f $Message) -Level Debug

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

    $TextBox = [System.Windows.Forms.TextBox]::new()
    $TextBox.Location = [System.Drawing.Point]::new(10, 55)
    $TextBox.Size = [System.Drawing.Size]::new(380, 20)
    $TextBox.Text = $DefaultValue
    $Form.Controls.Add($TextBox)

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

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        return $TextBox.Text
    }
    return ''
}
