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