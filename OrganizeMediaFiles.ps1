# ==============================================================================================
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. 
#
# JPG files contain EXIF data which has a DateTaken value. 
# Other media files have a MediaCreated date.  
#
# It will also append a sequenced number to the end of the file name if the name already exists to avoid name collisions. 
# The script will look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under DestinationRootPath, e.g. :
#
# "SourceRootPath\\IMG_2011-02-09_21-41-47_680.jpg"
# Will be changed to:
# "DestinationRootPath\2011\201102\20110209_214147.jpg"
# Or if already exists:
# "DestinationRootPath\2011\201102\20110209_214147_1.jpg"
#
#
# The code for extracting the EXIF DateTaken is based on a script by Kim Oppalfens:
# #http://blogcastrepository.com/blogs/kim_oppalfenss_systems_management_ideas/archive/2007/12/02/organize-your-digital-photos-into-folders-usi#ng-powershell-and-exif-data.aspx
# ============================================================================================== 

[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") | out-null

$global:ConfirmAll = $false

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

function Convert-AsciiArrayToString
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
        $CharArray 
    )

    $ReturnVal = ""
    foreach ($Char in $CharArray) 
    {
        $ReturnVal += [char]$Char
    }
    return $ReturnVal
}

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

function New-FilePath
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
        # Affix
        [Parameter(Position = 2)]
        $Affix,
        # Extension
        [Parameter(Position = 3, Mandatory = $true)]
        $Extension
    )

    if ($Affix)
    {
        return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $Affix, $Extension.ToLower())
    }
    else
    {
        return [String]::Format("{0}\{1}{2}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $Extension.ToLower())
    }
}

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

function Confirm-ContinueProcessing
{
    if ($global:ConfirmAll -eq $false) 
    {
        $Response = Read-Host "Continue? (Y/N/A)"
        if ($Response.Substring(0, 1).ToUpper() -eq "A") 
        {
            $global:ConfirmAll = $true
        }
        if ($Response.Substring(0, 1).ToUpper() -eq "N") 
        { 
            break 
        }
    }
}

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


function Add-OrganizeMedia
{
    [CmdletBinding()] 
    param ( 
        # SourceRootPath
        [parameter(Position = 0, ValueFromPipelineByPropertyName = $true )]
        [string]
        $SourceRootPath = "\\diskstation\photo\upload\Marcel",
        # DestinationRootPath
        [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true )]
        $DestinationRootPath = "\\diskstation\photo",
        # FileTypesToOrganize
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true )]
        [ValidateSet("*.jpg", "*.avi", "*.mp4", "*.3gp", "*.mov")]
        $FileTypesToOrganize = @("*.jpg"),        
        # FolderName
        [Parameter(Position = 3)]
        $FolderName = ""
    )

    Write-Verbose "Begin Organizing media..."
    
    $Files = Get-AllSourceFiles -Path $SourceRootPath -FileExtensions $FileTypesToOrganize
    foreach ($File in $Files) 
    {
        $CreationDate = Get-CreationDate -File $File
        if ($null -ne ($CreationDate -as [DateTime])) 
        {
            $DestinationPath = New-DesinationPath -Path $DestinationRootPath -Date $CreationDate -FolderName $FolderName
            New-MediaDirectory -Path $DestinationPath
            $NewFilePath = New-FilePath -Path $DestinationPath -Date $CreationDate -Extension $File.Extension
           
            $i = 1
            while (Test-Path $NewFilePath)
            {
                $NewFilePath = New-FilePath -Path $DestinationPath -Date $CreationDate -Affix $i -Extension $File.Extension  
                $i++                     
            }

            Move-Item -Path $File.FullName -Destination $NewFilePath -Verbose

        } 
        else 
        {
            Write-Warning "Unable to determine creation date of file: $File" 
            Confirm-ContinueProcessing
        }
    } 

    Write-Verbose "Finished Organizing media..."
}

# ============================================================================================== 
# Main
# ============================================================================================== 
# $Source = "P:\Upload\henk"
# $Source = "P:\Upload\lisa"
# 
# $Dest = "P:\"
# $Folder = ""

# Add-OrganizeMedia -SourceRootPath $Source -DestinationRootPath $Dest -FolderName $Folder -verbose