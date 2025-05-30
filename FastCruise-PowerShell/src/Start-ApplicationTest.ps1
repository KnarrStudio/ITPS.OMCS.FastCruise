function Start-ApplicationTest    
{
    param
    (
        [Parameter(Mandatory, Position = 0)]
        [Switch]$WaitTest,
        [Parameter(Mandatory, Position = 1)]
        [string]$TestFile,
        [Parameter(Mandatory, Position = 2)]
        [string]$TestProgram,
        [Parameter(Mandatory, Position = 3)]
        [string]$ProcessName
    )
    
    $DescriptionLists = [Ordered]@{
        FunctionResult = 'Good', 'Failed'
    }
    
    try
    {
        Start-Process -FilePath $TestProgram -ArgumentList $TestFile
        if($WaitTest)
        {
            Wait-Process -Name $ProcessName
        }
        return $DescriptionLists.FunctionResult[0]
    }
    catch
    {
        return $DescriptionLists.FunctionResult[1]
    }
}