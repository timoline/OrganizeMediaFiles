# ==============================================================================================
# 
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. It will also append
# a random number to the end of the file name just to avoid name collisions. The script will
# look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under
# DestinationRootPath, e.g. :
# DestinationRootPath\2011\02_February\2011-02-09_21-41-47_680.jpg
# I changed it to: DestinationRootPath\2011\201102\20110209_214147_680.jpg
#
# JPG files contain EXIF data which has a DateTaken value. Other media files have a MediaCreated
# date. 
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $File  
    )
    #Write-Verbose "Getting Media CreatedDate..."
	$Shell = New-Object -ComObject Shell.Application
	$Folder = $Shell.Namespace($File.DirectoryName)
	$CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), 191).Replace([char]8206, ' ').Replace([char]8207, ' ')

	if (($CreatedDate -as [DateTime]) -ne $null) 
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $File 
    )
    #Write-Verbose "Getting CreatedDate From Filename..."

	$Filename = $File.Name.Substring(0, 11).Replace("_", " ") + $File.Name.Substring(11, 8).Replace("-", ":")
	Write-Host $Filename
	Write-Host ($Filename -as [DateTime])
	if (($Filename -as [DateTime]) -ne $null) 
	{
		return [DateTime]::ParseExact($Filename,"yyyy-MM-dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) 
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $File 
    )

    #Write-Verbose "Getting CreatedDate From FileInfo..."
	return $File.CreationTime
}

function Convert-AsciiArrayToString
{
    [CmdletBinding()] 
    param ( 
        # File
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $File 
    )
    #Write-Verbose "Getting DateTaken From ExifData..."

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
	
	if (($DateString -as [DateTime]) -ne $null) 
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $File  
    )

    #Write-Verbose "Checking extension..."
	switch ($File.Extension) 
    { 
        ".jpg" { $CreationDate = Get-DateTakenFromExifData($File) } 
		".3gp" { $CreationDate =  Get-CreatedDateFromFilename($File) }
		".mov" { $CreationDate =  Get-CreatedDateFromFileInfo($File) }
        default { $CreationDate = Get-MediaCreatedDate($File) }
    }
	return $CreationDate
}

function Build-DesinationPath
{
    [CmdletBinding()] 
    param ( 
        # Path
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        [string] 
        $Path,  
        # Date
        [Parameter(Position=1,Mandatory=$true)]
        $Date,  
        # FolderName
        [Parameter(Position=2,Mandatory=$true)]
        $FolderName 
    )

	#return [String]::Format("{0}\{1}\{2}_{3}", $Path, $Date.Year, $Date.ToString("MM"), $Date.ToString("MMMM"))
	if ($FolderName) 
	{
        #Write-Verbose "Creating new foldername: $FolderName"
		return [String]::Format("{0}\{1}\{2}", $Path, $Date.Year, $FolderName)
	}
	else 
	{
        #Write-Verbose "Creating new foldername..."
		return [String]::Format("{0}\{1}\{2}{3}", $Path, $Date.Year, $Date.Year, $Date.ToString("MM"))
	}
}

function Build-NewFilePath
{
    [CmdletBinding()] 
    param ( 
        # Path
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        [string]
        $Path,  
        # Date
        [Parameter(Position=1,Mandatory=$true)]
        $Date,  
        # Affix
        [Parameter(Position=2)]
        $Affix,
        # Extension
        [Parameter(Position=3,Mandatory=$true)]
        $Extension
    )
    #Write-Verbose "Creating new filename..."
    #$RandomGenerator = New-Object System.Random
	#return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $RandomGenerator.Next(100, 1000).ToString(), $Extension)
    if ($Affix)
    {
        return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyyMMdd_HHmmss"), $Affix, $Extension)
    }
    else
    {
        return [String]::Format("{0}\{1}{2}", $Path, $Date.ToString("yyyyMMdd_HHmmss"),$Extension)
    }
}

function Create-Directory
{
    [CmdletBinding()] 
    param ( 
        # Path
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        [string] 
        $Path
    )

	if (!(Test-Path $Path)) 
	{
		New-Item $Path -Type Directory | out-null
        #Write-Verbose "Folder created: $Path"
	}
}

function Confirm-ContinueProcessing
{
	if ($global:ConfirmAll -eq $false) 
	{
		$Response = Read-Host "Continue? (Y/N/A)"
		if ($Response.Substring(0,1).ToUpper() -eq "A") 
		{
			$global:ConfirmAll = $true
		}
		if ($Response.Substring(0,1).ToUpper() -eq "N") 
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
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        [string] 
        $Path,  
        # FileTypesToOrganize
        [parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName=$true )]
        $FileExtensions
    )
    Write-Verbose "Getting Source files..."
	return @(Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize)
}


function OrganizeMedia
{
    [CmdletBinding()] 
    param ( 
        # SourceRootPath
        [parameter(Position=0,ValueFromPipelineByPropertyName=$true )]
        [string]
        $SourceRootPath = "\\diskstation\photo\upload\Marcel",
        # DestinationRootPath
        [Parameter(Position=1,ValueFromPipelineByPropertyName=$true )]
        $DestinationRootPath = "\\diskstation\photo",
        # FileTypesToOrganize
        [Parameter(Position=2,ValueFromPipelineByPropertyName=$true )]
        [ValidateSet("*.jpg","*.avi","*.mp4", "*.3gp", "*.mov")]
        $FileTypesToOrganize = @("*.jpg","*.avi","*.mp4", "*.3gp", "*.mov"),        
        # FolderName
        [Parameter(Position=3)]
        $FolderName = ""
    )

    Write-Verbose "Begin Organizing media..."
    
    $Files = Get-AllSourceFiles -Path $SourceRootPath -FileExtensions $FileTypesToOrganize
    foreach ($File in $Files) 
    {
	    $CreationDate = Get-CreationDate -File $File
	    if (($CreationDate -as [DateTime]) -ne $null) 
	    {
		    $DestinationPath = Build-DesinationPath -Path $DestinationRootPath -Date $CreationDate -FolderName $FolderName
		    Create-Directory -Path $DestinationPath
		    $NewFilePath = Build-NewFilePath -Path $DestinationPath -Date $CreationDate -Extension $File.Extension
           
            $i=1
            while (Test-Path $NewFilePath)
            {
                $NewFilePath = Build-NewFilePath -Path $DestinationPath -Date $CreationDate -Affix $i -Extension $File.Extension  
                $i++                     
            }

            Move-Item -Path $File.FullName -Destination $NewFilePath -Verbose
<# 		
		    if (-not(Test-Path $NewFilePath)) 
		    {
			    Move-Item -Path $File.FullName -Destination $NewFilePath -Verbose
                #Write-Verbose "$File -> $NewFilePath"
		    } 
		    else 
		    {
			    Write-Warning "Unable to rename file. File already exists."
			    Confirm-ContinueProcessing
		    }
#>
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

OrganizeMedia -SourceRootPath "\\diskstation\photo\upload\Marcel" -DestinationRootPath "\\diskstation\photo" -verbose