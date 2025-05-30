function Get-MacAddress {
     param(
         [Parameter(Position = 0)]
         [Switch]$LastFour
     )

     $MacAddress = (Get-NetAdapter -Physical | Where-Object status -EQ 'Up').MacAddress

     if ($MacAddress) {
         if ($LastFour) {
             return ($MacAddress -replace '^(.+-)(..)-(..)$', '$2:$3').tolower()
         }
         return $MacAddress.tolower()
     } else {
         Write-Error "No active network adapters found."
     }
 }

Export-ModuleMember -Function Get-MacAddress
