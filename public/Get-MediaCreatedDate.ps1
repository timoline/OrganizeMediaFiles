function Get-MediaCreatedDate
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $File  
    )
    $Shell = New-Object -ComObject Shell.Application
    $Folder = $Shell.Namespace($File.DirectoryName)
    $CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), 4).Replace([char]8206, ' ').Replace([char]8207, ' ')

    if ($null -ne ($CreatedDate -as [DateTime])) 
    {
        return [DateTime]::Parse($CreatedDate)
    } 
    else 
    {
        return $null
    }
}