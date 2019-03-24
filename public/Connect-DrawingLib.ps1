function Connect-DrawingLib
{
    [reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") | out-null
    Write-Verbose "Connecting to drawing library..."
}

