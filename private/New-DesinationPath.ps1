function New-DesinationPath
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
        # FolderName
        [Parameter(Position = 2, Mandatory = $true)]
        $FolderName 
    )

    if ($FolderName) 
    {
        return [String]::Format("{0}\{1}\{2}", $Path, $Date.Year, $FolderName)
    }
    else 
    {
        return [String]::Format("{0}\{1}\{2}{3}", $Path, $Date.Year, $Date.Year, $Date.ToString("MM"))
    }
}
