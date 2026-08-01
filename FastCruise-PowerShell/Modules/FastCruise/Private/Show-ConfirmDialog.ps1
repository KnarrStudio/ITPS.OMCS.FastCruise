function Show-ConfirmDialog
{
    <#
        .SYNOPSIS
        Asks the user a Yes/No question.

        .DESCRIPTION
        Displays a Windows Forms message box with Yes/No buttons. Replaces the
        old Microsoft.VisualBasic MsgBox YesNo call so every FastCruise prompt
        uses the same dialog technology on both Windows PowerShell 5.1 and
        PowerShell 7+.

        .PARAMETER Message
        Question text shown in the dialog.

        .PARAMETER Title
        Dialog window title. Defaults to 'Fast Cruise'.

        .EXAMPLE
        Show-ConfirmDialog -Message 'Is the location above correct?'

        .OUTPUTS
        System.Boolean. $true for Yes, $false for No.
    #>
    [CmdletBinding()]
    [OutputType([Boolean])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Message,
        [Parameter(Position = 1)]
        [String]$Title = 'Fast Cruise'
    )
    Write-FCLog -Message ('Prompting for confirmation: {0}' -f $Message) -Level Debug

    Add-Type -AssemblyName System.Windows.Forms

    $Result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
        [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
    )
    return ($Result -eq [System.Windows.Forms.DialogResult]::Yes)
}
