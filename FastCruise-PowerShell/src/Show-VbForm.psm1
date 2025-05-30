function Show-VbForm    
{
  [cmdletbinding(DefaultParameterSetName = 'Message')]
  param(
    [Parameter(Position = 0,ParameterSetName = 'Message')]
    [Switch]$YesNoBox,
    [Parameter(Position = 0,ParameterSetName = 'Input')]
    [Switch]$InputBox,
    [Parameter(Mandatory,Position = 1)]
    [string]$Message,
    [Parameter(Position = 2)]
    [string]$TitleBar = 'Fast Cruise',
    [Parameter(Position = 3,ParameterSetName = 'Input')]
    [string]$DefaultValue
  )
  Add-Type -AssemblyName Microsoft.VisualBasic
  switch($PSBoundParameters.Keys){
    'InputBox'
    {
      $Response = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $TitleBar, $DefaultValue)
    }
    'YesNoBox'
    {
      $Response = [Microsoft.VisualBasic.Interaction]::MsgBox($Message, 'YesNo, SystemModal, MsgBoxSetForeground', $TitleBar)
    }
  }
  $Response
}

Export-ModuleMember -Function Show-VbForm