param (
    [string]$Path
)

$ErrorActionPreference = "Stop"
$host.UI.RawUI.WindowTitle = "SPO Documents Report 📄"
Write-Host $Path -ForegroundColor Green
Set-location $Path
Get-ChildItem -Recurse | Unblock-File
. .\Misc\PS-Forms.ps1
Import-Module (Get-ChildItem -Recurse -Filter "*.psd1").FullName
Clear-Host


Write-host      
Write-host  "      __    ___  ___      ___                                      _       " -ForegroundColor Yellow
Write-host  "     / _\  / _ \/___\    /   \___   ___ _   _ _ __ ___   ___ _ __ | |_ ___ " -ForegroundColor Yellow
Write-host  "     \ \  / /_)//  //   / /\ / _ \ / __| | | | '_ `` _ \ / _ \ '_ \| __/ __|" -ForegroundColor Yellow
Write-host  "     _\ \/ ___/ \_//   / /_// (_) | (__| |_| | | | | | |  __/ | | | |_\__ \" -ForegroundColor Yellow
Write-host  "     \__/\/   \___/   /___,' \___/ \___|\__,_|_| |_| |_|\___|_| |_|\__|___/" -ForegroundColor Yellow

Write-host  "                         __                       _                       " -ForegroundColor Cyan   
Write-host  "                        /__\ ___ _ __   ___  _ __| |_                     " -ForegroundColor Cyan   
Write-host  "                       / \/// _ \ '_ \ / _ \| '__| __|                   " -ForegroundColor Cyan    
Write-host  "                      / _  \  __/ |_) | (_) | |  | |_                   " -ForegroundColor Cyan     
Write-host  "                      \/ \_/\___| .__/ \___/|_|   \__|                  " -ForegroundColor Cyan     
Write-host  "                                |_|                                       " -ForegroundColor Cyan                                        
Write-host                                                                       
Write-host "--------------------------------------------------------------------------------------"
Write-host                                                                       
# Created by Denis Molodtsov, 2022
                                            
Write-host "                Please, select one site at a time." -ForegroundColor Yellow
Write-host "                You will have to run this app one time per site." -ForegroundColor Yellow
Write-host   
Write-host "--------------------------------------------------------------------------------------"

$inputs = @{
    Site_URL = "https://contoso.sharepoint.com/sites/site/"
}

$inputs = Get-FormItemProperties -item $inputs -dialogTitle "Fill these required fields"

Connect-PnPOnline -UseWebLogin -Url $inputs.Site_URL -WarningAction Ignore

try {
    # This does not work for some reason. bug in the PnP Library?
    # $WebsCollection = Get-PnPSubWeb -Recurse 
    
    $web = Get-PnPWeb
    $webTitle = $web.Title
    $lists = Get-PnPList -Includes Title, BaseTemplate, Hidden
    $libraries = $lists | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }


    $documents = @()
    foreach ($library in $libraries) {
        $library = $libraries[0]
        $documents = @()
        
        $files = Get-PnPListItem -List $library -PageSize 200  -Fields Title, Id, FileRef, FileLeafRef, File_x0020_Size, LinkFilenameNoMenu, GUID

        foreach ($file in $files) {

            $Document = [PSCustomObject]@{
                FilePath           = $file["FileRef"]
                FileName           = $file["FileLeafRef"]
                SizeMegabytes      = ($file["File_x0020_Size"] / 1MB).ToString("#.##")
                Title              = $file["Title"]
                Extension          = [System.IO.Path]::GetExtension($file["FileLeafRef"])
                LibraryName        = $library.Title
            }

            if ($file["File_x0020_Size"]) {
                $documents += $Document
            }
        }
    }

    $documents | Export-Csv -Path "$webTitle.csv" -NoTypeInformation
   
    
}
catch {
    Write-Host
    Write-Host $_.Exception -ForegroundColor Red

    Write-Host Information on the failed email:
    Write-Host $row -ForegroundColor Yellow


    Write-Host Press any key...
    [System.Console]::ReadKey()
    Write-Host Press any key...
    [System.Console]::ReadKey()
    return
}


Write-host $reportwas successfully generated -ForegroundColor Green
Write-host You may close this window
Write-Host Press any key...
[System.Console]::ReadKey()