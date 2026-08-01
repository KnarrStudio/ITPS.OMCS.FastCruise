#Requires -Version 5.1
Set-StrictMode -Version Latest

# Module-scoped state. Set by Initialize-FCLog, read by Write-FCLog / Get-FCLogPath.
$Script:FCLogFilePath   = $null
$Script:FCMinimumLevel  = 'Info'
$Script:FCLevelWeight   = [Ordered]@{
    Debug   = 0
    Verbose = 1
    Info    = 2
    Warning = 3
    Error   = 4
}
$Script:FCLevelColor    = @{
    Debug   = 'Gray'
    Verbose = 'Cyan'
    Info    = 'White'
    Warning = 'Yellow'
    Error   = 'Red'
}

function Set-FCLogFileContent
{
    <#
        .SYNOPSIS
        Appends a line of text to a file using UTF-8 without a byte-order-mark.

        .DESCRIPTION
        Windows PowerShell 5.1 and PowerShell 7+ disagree on what "-Encoding UTF8"
        means (5.1 always writes a BOM, 7+ does not by default). Writing straight
        to the .NET file APIs sidesteps that difference so every log file this
        module produces looks identical regardless of which PowerShell wrote it.
        This is a private, module-internal helper and is not exported.

        .PARAMETER Path
        Full path of the file to append to. The file is created if it does not
        already exist.

        .PARAMETER Line
        The line of text to append. A newline is added automatically.

        .EXAMPLE
        Set-FCLogFileContent -Path 'C:\Logs\fastcruise.log' -Line 'hello'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Path,
        [Parameter(Mandatory, Position = 1)]
        [String]$Line
    )
    $Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::AppendAllText($Path, $Line + [Environment]::NewLine, $Utf8NoBom)
}

function Initialize-FCLog
{
    <#
        .SYNOPSIS
        Configures where and how Write-FCLog writes for the rest of the session.

        .DESCRIPTION
        Creates the log directory and the current log file if they do not exist,
        then remembers the resulting path and the minimum severity to record.
        Call this once near the start of a script; every later call to
        Write-FCLog in the same session uses this configuration. The log file
        name is rolled monthly (FastCruise_yyyy-MMMM.log) to match the reporting
        convention already used for the FastCruise CSV reports, so log retention
        and report retention line up.

        .PARAMETER LogDirectory
        Directory the log file should live in. Created automatically if missing.

        .PARAMETER LogName
        Base name of the log file, without extension or date. Defaults to
        'FastCruise'.

        .PARAMETER MinimumLevel
        The lowest severity that should actually be written. One of Debug,
        Verbose, Info, Warning, Error. Defaults to 'Info'.

        .EXAMPLE
        Initialize-FCLog -LogDirectory 'S:\FastCruise\Logs' -MinimumLevel Verbose
        Starts logging Verbose and above to S:\FastCruise\Logs.

        .OUTPUTS
        System.String. The full path of the active log file.
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$LogDirectory,
        [Parameter(Position = 1)]
        [String]$LogName = 'FastCruise',
        [Parameter(Position = 2)]
        [ValidateSet('Debug', 'Verbose', 'Info', 'Warning', 'Error')]
        [String]$MinimumLevel = 'Info'
    )
    if (-not (Test-Path -Path $LogDirectory))
    {
        $null = New-Item -Path $LogDirectory -ItemType Directory -Force
    }
    $YearMonth = Get-Date -Format 'yyyy-MMMM'
    $FileName  = '{0}_{1}.log' -f $LogName, $YearMonth
    $FullPath  = Join-Path -Path $LogDirectory -ChildPath $FileName
    if (-not (Test-Path -Path $FullPath))
    {
        $null = New-Item -Path $FullPath -ItemType File -Force
    }
    $Script:FCLogFilePath  = $FullPath
    $Script:FCMinimumLevel = $MinimumLevel
    Write-FCLog -Message ('Logging initialized. Minimum level: {0}' -f $MinimumLevel) -Level Info
    return $FullPath
}

function Get-FCLogPath
{
    <#
        .SYNOPSIS
        Returns the path of the log file currently in use.

        .DESCRIPTION
        Returns whatever Initialize-FCLog most recently configured, or $null if
        Initialize-FCLog has not been called yet in this session.

        .EXAMPLE
        Get-FCLogPath

        .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param()
    return $Script:FCLogFilePath
}

function Write-FCLog
{
    <#
        .SYNOPSIS
        Writes one log entry to the console and, once initialized, to the log file.

        .DESCRIPTION
        This is the single logging entry point every FastCruise script and
        module should call instead of Write-Verbose, Write-Warning, or
        Write-Host directly. It writes a timestamped, leveled line to the
        console (color-coded by severity) and, if Initialize-FCLog has been
        called, appends the same line to the active log file. Messages below
        the configured minimum level are written to the file only when the
        common -Verbose or -Debug switches are used; console visibility
        otherwise follows the configured MinimumLevel.

        .PARAMETER Message
        The text to log.

        .PARAMETER Level
        Severity of the message: Debug, Verbose, Info, Warning, or Error.
        Defaults to Info.

        .PARAMETER Source
        Name of the function or script raising the message. Defaults to the
        immediate caller's command name so callers do not need to supply it.

        .EXAMPLE
        Write-FCLog -Message 'Starting Fast Cruise' -Level Info

        .EXAMPLE
        Write-FCLog -Message 'Network path not available' -Level Warning -Source 'Start-FastCruise'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Message,
        [Parameter(Position = 1)]
        [ValidateSet('Debug', 'Verbose', 'Info', 'Warning', 'Error')]
        [String]$Level = 'Info',
        [Parameter(Position = 2)]
        [String]$Source = (Get-PSCallStack)[1].Command
    )
    if ($Script:FCLevelWeight[$Level] -lt $Script:FCLevelWeight[$Script:FCMinimumLevel])
    {
        return
    }
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line = '[{0}] [{1,-7}] [{2}] {3}' -f $Timestamp, $Level, $Source, $Message
    Write-Host -Object $Line -ForegroundColor $Script:FCLevelColor[$Level]
    if ($Script:FCLogFilePath)
    {
        Set-FCLogFileContent -Path $Script:FCLogFilePath -Line $Line
    }
}

Export-ModuleMember -Function 'Initialize-FCLog', 'Write-FCLog', 'Get-FCLogPath'
