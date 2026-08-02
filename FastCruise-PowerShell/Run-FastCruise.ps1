<#
    .SYNOPSIS
    Entry point for the ITPS.OMCS.FastCruise operational check.

    .DESCRIPTION
    Imports the FCLogging and FastCruise modules, runs an initial Fast
    Cruise check, and then (unless -NoMenu is used) presents a simple
    selection menu so the person running the check can repeat the check,
    record a facility issue, or restart the computer without leaving
    PowerShell.

    .PARAMETER ConfigPath
    Path to FastCruise.config.psd1. Defaults to Config\FastCruise.config.psd1
    next to this script.

    .PARAMETER ManualInput
    Passed through to Start-FastCruise: skip automatic detection and prompt
    for a computer name instead.

    .PARAMETER NoMenu
    Runs a single Fast Cruise check and exits instead of showing the
    follow-up menu.

    .EXAMPLE
    .\Run-FastCruise.ps1

    .EXAMPLE
    .\Run-FastCruise.ps1 -ManualInput -NoMenu
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [String]$ConfigPath,
    [Parameter(Position = 1)]
    [Switch]$ManualInput,
    [Parameter(Position = 2)]
    [Switch]$NoMenu
)

$ProjectRoot = $PSScriptRoot
Import-Module -Name (Join-Path -Path $ProjectRoot -ChildPath 'Modules\FCLogging\FCLogging.psd1') -Force
Import-Module -Name (Join-Path -Path $ProjectRoot -ChildPath 'Modules\FastCruise\FastCruise.psd1') -Force

if (-not $ConfigPath)
{
    $ConfigPath = Join-Path -Path $ProjectRoot -ChildPath 'Config\FastCruise.config.psd1'
}

$FastCruiseParams = @{
    ConfigPath  = $ConfigPath
    ManualInput = $ManualInput
}
Start-FastCruise @FastCruiseParams

if (-not $NoMenu)
{
    Show-FastCruiseMenu -ConfigPath $ConfigPath
}
