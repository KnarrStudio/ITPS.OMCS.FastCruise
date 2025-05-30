function Get-MacAddress {
    param(
        [Parameter(Position = 0)]
        [Switch]$LastFour
    )
    $MacAddress = (Get-NetAdapter -Physical | Where-Object -Property status -EQ -Value 'Up').macaddress
    if($LastFour) {
        $MacInfo = (($MacAddress.Split('-',5))[4]).replace('-',':')
    } else {
        $MacInfo = $MacAddress
    }
    $MacInfo
}