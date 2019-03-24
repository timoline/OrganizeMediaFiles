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
        ".jpg" { $CreationDate = Get-DateTakenFromExifData -File $File } 
        ".3gp" { $CreationDate = Get-CreatedDateFromFilename -File $File }
        ".mov" { $CreationDate = Get-CreatedDateFromFileInfo -File $File }
        default { $CreationDate = Get-MediaCreatedDate -File $File }
    }
    return $CreationDate
}