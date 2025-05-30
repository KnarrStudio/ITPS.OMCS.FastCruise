function Get-InstalledSoftware    
{
    [cmdletbinding(SupportsPaging)]
    Param(
        [Parameter(ValueFromPipeline, Position = 0)]
        [AllowNull()]
        [String[]]$SoftwareName,
        [ValidateSet('DisplayName','DisplayVersion')] 
        [AllowNull()]
        [String]$SelectParameter
    )
    Begin { 
        $SoftwareOutput = @()
        $InstalledSoftware = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*
    }
    Process {
        Try 
        {
            if($SoftwareName -eq $null) 
            {
                $SoftwareOutput = $InstalledSoftware |
                Select-Object -Property Installdate, DisplayVersion, DisplayName 
            }
            Else 
            {
                foreach($Item in $SoftwareName)
                {
                    $SoftwareOutput += $InstalledSoftware |
                    Where-Object -Property DisplayName -Match -Value $Item |
                    Select-Object -Property Installdate, DisplayVersion, DisplayName 
                }
            }
        }
        Catch 
        {
            $ErrorMessage  = $_.exception.message
        }
    }
    End{ 
        Switch ($SelectParameter){
            'DisplayName' 
            {
                $SoftwareOutput.displayname
            }
            'DisplayVersion' 
            {
                $SoftwareOutput.DisplayVersion
            }
            default 
            {
                $SoftwareOutput
            }
        }
    }
}

Export-ModuleMember -Function Get-InstalledSoftware  