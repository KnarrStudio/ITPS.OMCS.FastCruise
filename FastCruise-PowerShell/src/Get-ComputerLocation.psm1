function Get-ComputerLocation     
{
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowNull()]
        [String]$jsonFilePath
    )

    if(Test-Path -Path $jsonFilePath -ErrorAction SilentlyContinue)
    {
        Write-Verbose -Message 'Using JSON File'
        $location = Convert-JSONToHash -root $(Get-Content -Path $jsonFilePath -ErrorAction SilentlyContinue | ConvertFrom-Json)
        [string]$Script:LclDept = $location.Department.keys | Out-GridView -Title 'Department' -OutputMode Single
        [string]$Script:LclBuild = $location.Department[$LclDept].Building.Keys | Out-GridView -Title 'Building' -OutputMode Single
        [string]$Script:LclRm = $location.Department[$LclDept].Building[$LclBuild].Room | Out-GridView -Title 'Room' -OutputMode Single
        [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
    }
    else
    {
        Write-Verbose -Message 'Unable to find or use JSON File'
        [string]$Script:LclDept = Show-VbForm -InputBox -Message 'Department: Produce, Bakery, Dairy' -TitleBar 'Department' -DefaultValue 'Other'
        [string]$Script:LclBuild = Show-VbForm -InputBox -Message 'Building: Office-4, Bay-34' -TitleBar 'Building' -DefaultValue 'Office'
        [string]$Script:LclRm = Show-VbForm -InputBox -Message 'Room Number:' -TitleBar 'Room' -DefaultValue 1
        [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
    }
}

Export-ModuleMember -Function Get-ComputerLocation  