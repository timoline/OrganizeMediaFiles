function Add-OrganizeMedia
{
    [CmdletBinding()] 
    param ( 
        # SourceRootPath
        [parameter(Position = 0, ValueFromPipelineByPropertyName = $true )]
        [string]
        $SourceRootPath = "\\diskstation\photo\upload",
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

    Connect-DrawingLib

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