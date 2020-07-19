function Export-ComputerDescription
{
  <#
      .SYNOPSIS
      Creates the Description for AD

      .DESCRIPTION
      Creates the Description from input CSV file then outputs another csv that can be used to update the computer descriptions in Active Directory or where ever

      .PARAMETER InputReportFile
      Is a CSV file that has the following columns:
      ComputerName, Department, Building, Room, Desk

      .PARAMETER OutputListFile
      Is a CSV file that has the following columns:
      ComputerName, Description

      .EXAMPLE
      Export-ComputerDescriptionInputReportFile Value -OutputListFile Value
      Describe what this call does

      .NOTES
      This was built to be used after the Fast Cruise script to capture all of the locations for the systems

      .INPUTS
      .CSV

      .OUTPUTS
      .CSV
  #>


  param
  (
    [Parameter(Mandatory, Position = 0)]
    [ValidateScript({
          If($_ -match '.csv')
          {
            $true
          }
          Else
          {
            Throw 'Input file needs to be CSV'
          }
    })][String]$InputReportFile,
    [Parameter(Mandatory, Position = 1)]
    [String]$OutputListFile 
  )
  $FullDescriptionList = @()
  
  function Get-LastFour 
  {
    param(
      [Parameter(Mandatory,Position = 0)]
      [String]$MacAddress
    )
    $MacInfo = (($MacAddress.Split('-',5))[4]).replace('-',':')
    $MacInfo
  }

  $FastCruiseData = (Import-Csv -Path $InputReportFile) | Sort-Object -Property Department, Building 
  
  $FastCruiseData |
  ForEach-Object -Process {
    $MacFour = Get-LastFour -MacAddress $_.MacAddress
    $ADDescription = New-Object -TypeName System.Object
    $ADDescription | Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $_.ComputerName
    $ADDescription | Add-Member -MemberType NoteProperty -Name 'ComputerDescription' -Value ('OSD-OMC-{0}-{1}-{2}{3} [{4}]' -f $_.Department, $_.Building, $_.Room, $_.Desk, $MacFour) 

    $FullDescriptionList += $ADDescription
  }
  
  #Finally, use Export-Csv to export the data to a csv file
  $FullDescriptionList | Export-Csv -NoTypeInformation -Path $OutputListFile
  #Return $OutputListFile
}

Export-ComputerDescription -InputReportFile 'C:\temp\Reports\FastCruise_2020-June.csv' -OutputListFile $env:HOMEDRIVE\temp\Reports\computerDescriptions.csv


