function Get-CreationDate
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $File  
    )

    switch ($File.Extension) 
    { 
        ".jpg" { $CreationDate = Get-DateTakenFromExifData($File) } 
        ".3gp" { $CreationDate = Get-CreatedDateFromFilename($File) }
        ".mov" { $CreationDate = Get-CreatedDateFromFileInfo($File) }
        default { $CreationDate = Get-MediaCreatedDate($File) }
    }
    return $CreationDate
}