function Get-LastComputerStatus    
{
    param
    (
        [Parameter(Mandatory, Position = 0)]
        [String]$LastCruiseStatus
    )
    Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
    Write-Verbose -Message 'Importing the Fast Cruise Report'
    $CompImport = Import-Csv -Path $LastCruiseStatus
    Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"
    try
    {
        $LatestStatus = $CompImport |
        Where-Object -FilterScript {
            $PSItem.ComputerName -eq $env:COMPUTERNAME
        } |
        Select-Object -Last 1 
        if($LatestStatus -eq $null)
        {
            Write-Output -InputObject 'Unable to find an existing record for this system.'
            $Script:Ans = 'NoHistory'
        }
    }
    Catch
    {
        $ErrorMessage  = $_.exception.message
        Write-Verbose -Message ('Error Message: {0}' -f $ErrorMessage)
    }
    Return $LatestStatus
}

Export-ModuleMember -Function Get-LastComputerStatus    