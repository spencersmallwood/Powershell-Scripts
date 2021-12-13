<#
.Synopsis
   Scans the entire SharePoint tennat for files older than a certain date and then sends them to an Azure Blob
.DESCRIPTION
   
.INPUTS
    $tenantSiteURL: This is the top level tennant site url
    $libraryExclusions: This is to exclude any libraries we do not want to pull files from
    $startDate: The date you would like to use as a filter for the files that are moving, e.g. 4/29/2020 for all files modified less than a year ago
    $localPath: Files are downloaded to the local drive and then pushed to the Onedrive of choice. This is because onedrive uses a different URL even though it is in the same tennant
    $username: User name under which the script will run
    $pwd: password of the user listed above
.NOTES
--This uses the classic PnP SharPoint Powershell module, not the newer cross platform version which did not work
--Install SharePoint Online PnP PowerShell:
----Run $PSVersionTable.PSVersion
----Ensure you are above v3 and then run the command below:
------Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/pnp/PnP-PowerShell/master/Samples/Modules.Install/Install-SharePointPnPPowerShell.ps1')
----Reference:
------https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
--Install Azure Modules
----Install-Module Az -AllowClobber -Scope CurrentUser
----Reference:
------https://thedatacrew.com/articles/upload-multiple-files-to-azure-storage-with-powershell/
#>

#This function sets the connection for sharepoint online. The connection is set globaly in the script in memory which is when there is a function
#and you see the switching when you go from SP online to onedrive and back
function Set-Connection{
    Param($url)

    #----------------Remove this section-----------------
    $username = 'xx@xx.onmicrosoft.com'
    $pwd = 'xx' #App Password
    #----------------------------------------------------

	[SecureString]$securePass = ConvertTo-SecureString $pwd -AsPlainText -Force
	[System.Management.Automation.PSCredential]$PSCredentials = New-Object System.Management.Automation.PSCredential($username, $securePass)  
    #Connects to SharePoint online and stores the connection in memory
    Connect-PnPOnline -Url $url -Credentials $PSCredentials
}

#clear the screen
Clear-Host

##Set the varaibles
#allows us to cycle through the site collections in the tennant
$tenantSiteURL = "siteURL"
#these are the libraries we do not want the script to remove the old files. Add to this comma delimited list any libraries that you want to exclude
$libraryExclusions = @('Form Templates','Style Library','Organization Logos','Site Assets','Site Pages')
#set the start date for the file move. Any file modified less than the date specified will be moved
$startDate = Get-Date "10/01/2020"
#a local path is needed when moving to one drive. This is because the copy or move functions require a relative path and one drive accounts
#have their own URL even though they are on the same tennant
$localPath = "C:\CodeFiles\temp"
#URL to the onedrive folder when you would like the files to arrive
#$destinationUrl = "https://xx-my.sharepoint.com/personal/xx_xx_onmicrosoft_com"
#Azure storage varables
$storageAccountName = "xx"
$storageAccountKey = "xx"
$containerName = "xx"

#Check if the storage module is imported, if it is not, import it
$module = Get-module -name "Az.Storage" -Verbose

If (($module.Name) -match ("Az.Storage")) {
    Write-Host "Module Az.Storage is loaded"
}
else {
    Write-Host "Loading module Az.Storage...."
	Import-Module Az.Storage
}

#set the Azure blob storage context
$ctx = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

#set the connection to the tennant
Set-Connection -url $tenantSiteURL
#get all of the site collections in the tennant
$sites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites

foreach($site in $sites)
{
    write-host "Site collection: $($site.Url)" -BackgroundColor Cyan -ForegroundColor Black

    #Set the commection to the site collection
    $siteUrl = $site.Url
    Set-Connection -url $siteUrl

    #get the root web of the site collection
    $siteWeb = Get-PnPWeb
    #get the lists at the root web of the site collection
    $sitelists = Get-PnPList -Web $siteWeb

    #For each list in the root web of the site collection
    foreach($siteList in $sitelists)
    {
        #get only the Document Libraries, ones that are not hidden (Which are usually system libraries), and exclude ones from our exclusion list
        if($siteList.BaseType -eq "DocumentLibrary" -and $siteList.Hidden -eq $false -and $libraryExclusions.Contains($siteList.Title) -eq $false)
        {
            #if the library has more than 5k items the Get Items will fail, thus only get 5k items
            $query = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>"
            $siteItems = Get-PnPListItem -List $siteList.Title -Web $siteWeb

            #each item from the results of the query above
            foreach($siteItem in $siteItems)
            {
                #get the modified date
                $itemModified = $siteItem.FieldValues["Modified"]

                $siteItem.FieldValues["Modified"]

                #filter the items out whose modified date are less that the variable start date
                if($itemModified -lt $startDate)
                {
                    #File Ref is the relative file path
                    $sourceUrl = $siteItem.FieldValues["FileRef"]
                    #File Leaf Ref is the file name
                    $fileName = $siteItem.FieldValues["FileLeafRef"]

                    #Downloads the file to the local drive, spfile is used here because it causes an error without it
                    if($item.FieldValues["FSObjType"] -eq 0) #make sure it is a file object type
                        {
                            try{
                                $spfile = Get-PnPFile -Url $sourceUrl -Path $localPath -AsFile -Force
                            }
                            catch{
                                $_
                            }
                        }
                    
                    #remove the downloaded item from the original list. It goes into the recycle bin
                    #********************#
					#Remove-PnPListItem -List $siteList -Identity $siteItem -Force -Recycle
					#********************#
                }
            }
        }
    }

    #get all the sub web sites in the site collection
    $webs = Get-PnPSubWebs

    #For each web in the site collection
    foreach($w in $webs)
    {
        write-host "Processing Web: $($w.url)" -BackgroundColor Green -ForegroundColor Black

        #For each list in the web site of the site collection 
        $lists = Get-PnPList -Web $w
        foreach($list in $lists)
        {
            #get only the Document Libraries, ones that are not hidden (Which are usually system libraries), and exclude ones from our exclusion list
            if($list.BaseType -eq "DocumentLibrary" -and $list.Hidden -eq $false -and $libraryExclusions.Contains($list.Title) -eq $false)
            {
                #if the library has more than 5k items the Get Items will fail, thus only get 5k items
                $query = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>"
                $items = Get-PnPListItem -List $list.Title -Web $w -Query $query
                
                #each item from the results of the query above
                foreach($item in $items)
                {
                    #get the modified date
                    $itemModified = $item.FieldValues["Modified"]

                    $itemModified = $item.FieldValues["Modified"]

                    #filter the items out whose modified date are less that the variable start date
                    if($itemModified -lt $startDate)
                    {
                        #File Ref is the relative file path
                        $sourceUrl = $item.FieldValues["FileRef"]
                        #File Leaf Ref is the file name
                        $fileName = $item.FieldValues["FileLeafRef"]
                        #Downloads the file to the local drive, spfile is used here because it causes an error without it
                        if($item.FieldValues["FSObjType"] -eq 0) #make sure it is a file object type
                        {
                            try{
                                $spfile = Get-PnPFile -Url $sourceUrl -Path $localPath -AsFile -Force
                            }
                            catch{
                                $_
                            }
                        }
                        #remove the downloaded item from the original list. It goes into the recycle bin
						#********************#
						#Remove-PnPListItem -List $siteList -Identity $siteItem -Force -Recycle
						#********************#
                    }
                }
            }
        }
    }

}

#now go through all of the files and send them to the blob
$files = Get-ChildItem $localPath

foreach ($f in $files){
    write-host "Sending file to the blob: $($f.FullName)"
    Set-AzStorageBlobContent -File $($f.FullName) -Container $containerName -Context $ctx -Force
    #remove local file
    Remove-Item -Path $($f.FullName)
}


#All done!
Write-Host "Completed running script" -BackgroundColor Green -ForegroundColor Black
