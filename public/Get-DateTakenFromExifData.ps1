function Get-DateTakenFromExifData
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $File 
    )

    $FileDetail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $File.Fullname 
    $DateTimePropertyItem = $FileDetail.GetPropertyItem(36867)
    $FileDetail.Dispose()

    $Year = Convert-AsciiArrayToString $DateTimePropertyItem.value[0..3]
    $Month = Convert-AsciiArrayToString $DateTimePropertyItem.value[5..6]
    $Day = Convert-AsciiArrayToString $DateTimePropertyItem.value[8..9]
    $Hour = Convert-AsciiArrayToString $DateTimePropertyItem.value[11..12]
    $Minute = Convert-AsciiArrayToString $DateTimePropertyItem.value[14..15]
    $Second = Convert-AsciiArrayToString $DateTimePropertyItem.value[17..18]
	
    $DateString = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
	
    if ($null -ne ($DateString -as [DateTime])) 
    {
        return [DateTime]::Parse($DateString)
    } 
    else 
    {
        return $null
    }
}