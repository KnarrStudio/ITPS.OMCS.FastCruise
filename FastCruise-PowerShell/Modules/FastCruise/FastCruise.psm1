#Requires -Version 5.1
<#
    FastCruise.psm1 - root module file.

    Responsible only for wiring the project together:
      1. Locate the project root (so default config/log paths can be
         resolved without hardcoding an absolute path anywhere).
      2. Make sure the FCLogging module is loaded, since every function in
         this module calls Write-FCLog.
      3. Dot-source every Private helper and every Public function into the
         module's scope, then export only the Public function names.

    All actual behavior lives in Public\*.ps1 and Private\*.ps1 - each of
    those files is self-contained and documented with its own
    comment-based help.
#>
Set-StrictMode -Version Latest

$Script:ModuleRoot       = $PSScriptRoot
$Script:ProjectRoot      = Split-Path -Path (Split-Path -Path $Script:ModuleRoot -Parent) -Parent
$Script:DefaultConfigPath = Join-Path -Path $Script:ProjectRoot -ChildPath 'Config\FastCruise.config.psd1'

if (-not (Get-Module -Name FCLogging))
{
    $LoggingModulePath = Join-Path -Path (Split-Path -Path $Script:ModuleRoot -Parent) -ChildPath 'FCLogging\FCLogging.psd1'
    Import-Module -Name $LoggingModulePath -Force -ErrorAction Stop
}

$PrivateFunctions = Get-ChildItem -Path (Join-Path -Path $Script:ModuleRoot -ChildPath 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue
$PublicFunctions  = Get-ChildItem -Path (Join-Path -Path $Script:ModuleRoot -ChildPath 'Public')  -Filter '*.ps1' -ErrorAction SilentlyContinue

foreach ($Function in @($PrivateFunctions) + @($PublicFunctions))
{
    try
    {
        . $Function.FullName
    }
    catch
    {
        Write-Error -Message ('Failed to load {0}: {1}' -f $Function.FullName, $_.Exception.Message)
    }
}

Export-ModuleMember -Function $PublicFunctions.BaseName
