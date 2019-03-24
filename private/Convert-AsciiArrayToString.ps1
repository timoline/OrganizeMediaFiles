function Convert-AsciiArrayToString
{
    [CmdletBinding()] 
    param ( 
        #CharArray 
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $CharArray 
    )

    $ReturnVal = ""
    foreach ($Char in $CharArray) 
    {
        $ReturnVal += [char]$Char
    }
    return $ReturnVal
}