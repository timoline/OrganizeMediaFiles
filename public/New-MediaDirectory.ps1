function New-MediaDirectory
{
    [CmdletBinding()] 
    param ( 
        # Path
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        [string] 
        $Path
    )

    if (!(Test-Path $Path)) 
    {
        New-Item $Path -Type Directory | out-null
    }
}