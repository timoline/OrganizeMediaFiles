function Get-AllSourceFiles
{
    [CmdletBinding()] 
    param ( 
        # SourceRootPath
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        [string] 
        $Path,  
        # FileTypesToOrganize
        [parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $FileExtensions
    )
    Write-Verbose "Getting Source files..."
    return @(Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize)
}