function New-FilePath
{
    [CmdletBinding()] 
    param ( 
        # Path
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        [string]
        $Path,  
        # Date
        [Parameter(Position = 1, Mandatory = $true)]
        $Date,  
        # Affix
        [Parameter(Position = 2)]
        $Affix,
        # Extension
        [Parameter(Position = 3, Mandatory = $true)]
        $Extension
    )

    if ($Affix)
    {
        return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $Affix, $Extension.ToLower())
    }
    else
    {
        return [String]::Format("{0}\{1}{2}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $Extension.ToLower())
    }
}