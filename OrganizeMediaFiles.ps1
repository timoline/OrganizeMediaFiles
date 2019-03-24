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



$global:ConfirmAll = $false


# ============================================================================================== 
# Main
# ============================================================================================== 
# $Source = "P:\Upload\henk"
# $Source = "P:\Upload\lisa"
# 
# $Dest = "P:\"
# $Folder = ""

# Add-OrganizeMedia -SourceRootPath $Source -DestinationRootPath $Dest -FolderName $Folder -verbose