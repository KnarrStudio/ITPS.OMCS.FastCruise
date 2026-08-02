function Start-ApplicationTest
{
    <#
        .SYNOPSIS
        Launches a program with a test file to confirm it opens correctly.

        .DESCRIPTION
        Starts TestProgram with TestFile as its argument, optionally waits for
        the process to close, then asks the user whether the test succeeded.

        .PARAMETER WaitTest
        If set, waits for ProcessName to exit before asking for a result.

        .PARAMETER TestFile
        Path of the file to open (e.g. a sample PDF or PPTX).

        .PARAMETER TestProgram
        Path to the program executable to launch.

        .PARAMETER ProcessName
        Process name to wait on when -WaitTest is used, and to label the
        result prompt with.

        .EXAMPLE
        Start-ApplicationTest -WaitTest -TestFile 'C:\Test.pdf' -TestProgram 'C:\Program Files\Adobe\Acrobat.exe' -ProcessName 'Acrobat'

        .OUTPUTS
        System.String. 'Good' or 'Failed'.
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Position = 0)]
        [Switch]$WaitTest,
        [Parameter(Mandatory, Position = 1)]
        [String]$TestFile,
        [Parameter(Mandatory, Position = 2)]
        [String]$TestProgram,
        [Parameter(Mandatory, Position = 3)]
        [String]$ProcessName
    )
    try
    {
        Write-FCLog -Message ('Attempting to open {0} with {1}' -f $TestFile, $TestProgram) -Level Verbose
        Start-Process -FilePath $TestProgram -ArgumentList $TestFile
    }
    catch
    {
        Write-FCLog -Message $_.Exception.Message -Level Error
        return 'Failed'
    }

    if ($WaitTest)
    {
        Write-FCLog -Message ('Waiting for {0} to close before continuing.' -f $ProcessName) -Level Info
        Wait-Process -Name $ProcessName -ErrorAction SilentlyContinue
    }

    return Show-SelectionDialog -Message ('Did {0} open and display correctly?' -f $ProcessName) -Options @('Good', 'Failed') -Title $ProcessName
}
