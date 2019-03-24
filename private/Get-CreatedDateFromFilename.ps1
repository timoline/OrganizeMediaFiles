function Get-CreatedDateFromFilename
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $File 
    )

    $Filename = $File.Name.Substring(0, 11).Replace("_", " ") + $File.Name.Substring(11, 8).Replace("-", ":")
    Write-Host $Filename
    Write-Host ($Filename -as [DateTime])
    if ( $null -ne ($Filename -as [DateTime])) 
    {
        return [DateTime]::ParseExact($Filename, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture) 
    } 
    else 
    {
        return $null
    }
}