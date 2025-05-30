function Convert-JSONToHash {
    param(
        [AllowNull()]
        [Object]$root
    )
    $hash = @{}
    $keys = $root |
    Get-Member -MemberType NoteProperty |
    Select-Object -ExpandProperty Name
    $keys | ForEach-Object -Process {
        $key = $_
        $obj = $root.$($_)
        if($obj -match '@{') {
            $nesthash = Convert-JSONToHash -root $obj
            $hash.add($key, $nesthash)
        } else {
            $hash.add($key, $obj)
        }
    }
    return $hash
}