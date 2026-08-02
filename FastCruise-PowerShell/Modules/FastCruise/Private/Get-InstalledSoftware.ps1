function Get-InstalledSoftware
{
    <#
        .SYNOPSIS
        Returns installed-software entries from the local uninstall registry key.

        .DESCRIPTION
        Reads HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* and
        returns the install date, version, and display name for each entry.
        Optionally filters to names matching one or more search terms and can
        return just the display name or just the version instead of full
        objects.

        .PARAMETER SoftwareName
        One or more substrings to match against installed-software display
        names. If omitted, every installed program is returned.

        .PARAMETER SelectParameter
        When set to 'DisplayName' or 'DisplayVersion', returns only that
        property's value(s) instead of the full object.

        .EXAMPLE
        Get-InstalledSoftware -SoftwareName 'Mozilla Firefox' -SelectParameter DisplayVersion
        Returns just the installed Firefox version string.

        .EXAMPLE
        Get-InstalledSoftware
        Returns every installed program on the local machine.

        .OUTPUTS
        System.String or PSCustomObject, depending on -SelectParameter.
    #>
    [CmdletBinding(SupportsPaging)]
    param(
        [Parameter(ValueFromPipeline, Position = 0)]
        [AllowNull()]
        [String[]]$SoftwareName,
        [Parameter(Position = 1)]
        [ValidateSet('DisplayName', 'DisplayVersion')]
        [AllowNull()]
        [String]$SelectParameter
    )
    Begin
    {
        Write-FCLog -Message 'Reading installed software from the registry.' -Level Debug
        $SoftwareOutput    = @()
        $UninstallKeyPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
        $InstalledSoftware = Get-ItemProperty -Path $UninstallKeyPath
    }
    Process
    {
        try
        {
            if ($null -eq $SoftwareName)
            {
                $SoftwareOutput = $InstalledSoftware |
                    Select-Object -Property InstallDate, DisplayVersion, DisplayName
            }
            else
            {
                foreach ($Item in $SoftwareName)
                {
                    $SoftwareOutput += $InstalledSoftware |
                        Where-Object -Property DisplayName -Match -Value $Item |
                        Select-Object -Property InstallDate, DisplayVersion, DisplayName
                }
            }
        }
        catch
        {
            Write-FCLog -Message $_.Exception.Message -Level Error
        }
    }
    End
    {
        switch ($SelectParameter)
        {
            'DisplayName'    { $SoftwareOutput.DisplayName }
            'DisplayVersion' { $SoftwareOutput.DisplayVersion }
            default          { $SoftwareOutput }
        }
    }
}
