function Set-FCContent
{
    <#
        .SYNOPSIS
        Writes text to a file using a single, uniform encoding.

        .DESCRIPTION
        Windows PowerShell 5.1 and PowerShell 7+ pick different default
        encodings for Out-File/Set-Content (5.1 writes a UTF-8 byte-order-mark,
        7+ generally does not). That made the JSON and CSV files this project
        produces inconsistent with each other depending on which PowerShell
        wrote them. Set-FCContent and Add-FCContent always write UTF-8 without
        a byte-order-mark, on both editions, so every file FastCruise produces
        looks the same regardless of which PowerShell version ran it.

        .PARAMETER Path
        Full path of the file to write.

        .PARAMETER Content
        Text content to write. Replaces the file if it already exists.

        .EXAMPLE
        Set-FCContent -Path 'S:\ComputerLocation.json' -Content $JsonText
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Path,
        [Parameter(Mandatory, Position = 1)]
        [AllowEmptyString()]
        [String]$Content
    )
    $Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function Add-FCContent
{
    <#
        .SYNOPSIS
        Appends a line of text to a file using a single, uniform encoding.

        .DESCRIPTION
        Companion to Set-FCContent. Appends UTF-8 (no byte-order-mark) text to
        a file, creating the file first if it does not exist yet, so append
        behavior is identical on Windows PowerShell 5.1 and PowerShell 7+.

        .PARAMETER Path
        Full path of the file to append to.

        .PARAMETER Line
        The line of text to append. A newline is added automatically.

        .EXAMPLE
        Add-FCContent -Path $ReportPath -Line '- Lights out in room 102'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String]$Path,
        [Parameter(Mandatory, Position = 1)]
        [AllowEmptyString()]
        [String]$Line
    )
    if (-not (Test-Path -Path $Path))
    {
        $null = New-Item -Path $Path -ItemType File -Force
    }
    $Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::AppendAllText($Path, $Line + [Environment]::NewLine, $Utf8NoBom)
}
