function Get-CreatedDateFromFileInfo
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $File 
    )

    return $File.CreationTime
}