#$credentials = Get-Credential
function Set-Connection{
    Connect-PnPOnline -Url SITEURL –UseWebLogin -WarningAction Ignore
    
}

#clear the screen
Clear-Host

##Set the varaibles
$siteUrl = "SITEURL"
#these are the libraries we do not want the script to remove the old files. Add to this comma delimited list any libraries that you want to exclude
$libraryExclusions = @('Form Templates','Style Library','Organization Logos','Site Assets','Site Pages')
#set the start date for the file move. Any file modified less than the date specified will be moved
$startDate = Get-Date "07/10/2021"
$endDate = Get-Date "07/19/2021"

$localPathStatic = "C:\Files"


    write-host "Site collection: $siteUrl" -BackgroundColor Cyan -ForegroundColor Black

    #Set the connection to the site collection

    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    
    #get the root web of the site collection
    $siteWeb = Get-PnPWeb  

    #get the lists at the root web of the site collection
    $sitelists = Get-PnPList #-Web $siteWeb
     write-host "Site list gotten" -BackgroundColor Cyan -ForegroundColor Black
    #For each list in the root web of the site collection
    foreach($siteList in $sitelists)
    {
    write-host "In the foreach siteList......................" -BackgroundColor Cyan -ForegroundColor Black
        #get only the Document Libraries, ones that are not hidden (Which are usually system libraries), and exclude ones from our exclusion list
        
        if($siteList.BaseType -eq "DocumentLibrary" -and $siteList.Hidden -eq $false -and $libraryExclusions.Contains($siteList.Title) -eq $false)
        {
        write-host "*************In the IF Library " -BackgroundColor Cyan -ForegroundColor Black
            #if the library has more than 5k items the Get Items will fail, thus only get 5k items
            $query = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>"
            $siteItems = Get-PnPListItem -List $siteList.Title -Web $siteWeb



            #each item from the results of the query above
            foreach($siteItem in $siteItems)
            {
                write-host "....................................In the foreach ITEM " -BackgroundColor Cyan -ForegroundColor Black
                #get the modified date only without time
                $itemModified = $siteItem.FieldValues["Modified"].Date
                 write-host "Modified Date --------------------> $itemModified" -BackgroundColor Green -ForegroundColor Black
                #filter the items out whose modified date are less that the variable start date
                if(($itemModified -ge $startDate) -and ($itemModified -le $endDate))
                {
                    #File Ref is the relative file path
                    $sourceUrl = $siteItem.FieldValues["FileRef"]

                    #File Leaf Ref is the file name
                    $fileName = $siteItem.FieldValues["FileLeafRef"]

                    $fileType = $siteItem.FieldValues["Folder"].Text

                    write-host "//// ----- ///// item name  $fileName " -BackgroundColor Green -ForegroundColor Black
                    
                    #delete the name of the file to get only folders name 
                    $FilePathLocal = $sourceUrl.Replace($fileName,'')

                    $ParentFolderPath = $localPathStatic 

                    foreach($FolderName in $FilePathLocal.split("/")) 
                    {
                        $CreateFolderPath = $ParentFolderPath + "\$FolderName"
                        $ParentFolderPath = $CreateFolderPath

                        New-Item -ItemType Directory -Force -Path $CreateFolderPath
                        #write-host "==============================================LOCAAAAL Folder $CreateFolderPath " -BackgroundColor White -ForegroundColor Black

                    }
                    
                   # write-host "the siteItem.FileSystemObjectType is :::  $($siteItem.Folder) " -BackgroundColor White -ForegroundColor Black

                    

                    $CreateFolderPath = $CreateFolderPath.Replace("\\",'\')

                    Try 
                    {
                        $spfile = Get-PnPFile -Url $sourceUrl -Path $CreateFolderPath -FileName $fileName -AsFile -Force -ErrorAction Stop
                        Remove-PnPListItem -List $siteList -Identity $siteItem -Force -Recycle
                    }
                    Catch 
                    {
                        write-host "IT IS A FOLDER " -BackgroundColor White -ForegroundColor Black
                    }

					}
                }
            }
        }

