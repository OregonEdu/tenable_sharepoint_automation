## Process Tenable WAS Scan Results to SharePoint Automatically - v4.1
## Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov> - IT Operations & Support Network Team Lead - Oregon Department of Education
#
#  Requirements: An Entra application with SharePoint delegate API permissions
#
#  This script will automatically process a Tenable WAS scan results JSON file to a SharePoint site. It automated SharePoint List creation as well uploading files to the Document library. It is extremely modular and has functions
#  for virtually every aspect of SharePoint automation and thus could be dissected for any type of SharePoint automation tasks.

# Add .NET Forms
Add-Type -AssemblyName System.Windows.Forms

# Create our class for the imported Tenable Report JSON
class ODETenableItem 
{
    [string]$Title
    [datetime]$ScanDate
    [string]$PluginID
    [string]$CVE
    [double]$CVSSv3
    [string]$Risk
    [string]$URI    
    [string]$Synopsis
    [string]$Information
    [string]$Solution
    [string]$Remediated
    [string]$DateRemediated
    [string]$Tracking       
}

# Create our class for the Tenable WAS Summary List Template
class ODETenableSummaryItem
{
    [string]$Title
    [datetime]$LastScanDate
    [string]$Status
    [int]$CriticalCount
    [int]$HighCount
    [int]$MediumCount
    [string]$ScanNotes
    [string]$ScanHistory
}

# Ordered hashtable so that it actually creates them in the Vulnerability Tracker List in this order as opposed to being random -- used a hashtable since originally I was going to create the columns in the script. The number
# values define the type of column they are (https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee540543(v=office.15)). I scrapped this idea though and used Content Types which offered more options (colored labels).
$odeTenableSPOList = [ordered]@{
    "Scan_x0020_Date" = 4
    "Plugin_x0020_ID" = 1
    CVE = 4
    CVSSv3 = 9
    Risk = 2
    URI = 3
    Synopsis = 3
    Information = 3
    Solution = 3
    Remediated = 2
    "Date_x0020_Remediated" = 4
    Tracking = 3
}

# Scan History List columns
$odeTenableSummaryList = [ordered]@{
    "Last Scan Date" = 4
    Status = 2
    "%23 of CRITICAL" = 1
    "%23 of HIGH" = 1
    "%23 of MEDIUM" = 1
    "Scan Notes" = 3    
}

# Yes, 4 versions.
$scriptVersion = "v4.1"
# Base tenant SharePoint URL
$tenantDomain = "https://odemail.sharepoint.com"
# Name of the SharePoint List housing the schedule of the Tenable WAS schedule -- %20 is a space
$spoScanScheduleList = "Scan%20Schedule"
# Name of the SharePoint List housing the summary of the Tenable WAS scans
$spoScanSummaryList = "Scan%20Summary"
# Name of the SharePoint Document Library housing the Tenable Report PDF's
$spoReportDocumentLibary = "Scan%20Reports"
# SharePoint List Data Type for Scan Schedule List needed to MERGE (update) the hyperlink columns -- _x0020_ is a space in SharePoint List nomenclature
$spoListDataType = "Scan_x0020_Schedule"
# SharePoint List Data Type for Scan Summary List needed to MERGE (update) the hyperlink column
$spoSummaryListDataType = "Scan_x0020_Summary"
# Name of the SharePoint site collection
$odeTenableWasSpoSite = "ODETenableWebApplicationScanning"
# Base URL of the Reports Document Library to be used as a clickable hyperlink in the Scan Schedule List
$reportFolderBaseUrl = "$($tenantDomain)/sites/$($odeTenableWasSpoSite)/$($spoReportDocumentLibary)/Forms/AllItems.aspx?id=%2Fsites%2F$($odeTenableWasSpoSite)%2F$($spoReportDocumentLibary)%2F"
# Base REST API URL for Lists
$tenableBaseList = "$($tenantDomain)/sites/$($odeTenableWasSpoSite)/_api/lists/" 
# Base REST API URL for web operations
$tenableBaseUrl = "$($tenantDomain)/sites/$($odeTenableWasSpoSite)/_api/web/"
# Content Type ID of the ODE Tenable WAS List Template
$odeTenableWasContentType = "0x01003583A55101FF6447A0019F999BDDB1B7"
# Content Type ID of the ODE Tenable WAS Summary List Template
$odeTenableWasSummaryContentType = "0x0100C7B3916115D23E4587BB46A268B5EC0B"
# My kick-ass Amiga font ascii for the splash screen
$odeLogo = ".\ODE_WAS.png"
# Client ID of odeSPO Azure application
$clientId = "ae024a63-55de-4205-a6a2-e5fc1394f389"
# ODE's tenant ID in Azure
$tenantId = "b4f51418-b269-49a2-935a-fa54bf584fc8"
# Resource for token authentication
$resource = "https://graph.microsoft.com/"

# Function to display said kick-ass ODE Amiga font ascii
function odeSplash
{    
    $odeAmiga = (Get-Item -Path $odeLogo)
    $img = [System.Drawing.Image]::FromFile($odeAmiga)

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $odeAmigaForm = New-Object Windows.Forms.Form
    $odeAmigaForm.Text = "ODE Tenable.io WAS Report Processor"
    $odeAmigaForm.Width = $img.Size.Width
    $odeAmigaForm.Height = $img.Size.Height
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width = $img.Size.Width
    $pictureBox.Height = $img.Size.Height
    $pictureBox.Image = $img
    $odeAmigaForm.Controls.Add($pictureBox)
    $odeAmigaForm.Add_Shown({$odeAmigaForm.Activate()})
    $odeAmigaForm.ShowDialog()
}

# Provide a GUI interface for opening the Tenable WAS Report JSON export
function openTenableReport($initialDirectory)
{
    $openFile = New-Object System.Windows.Forms.OpenFileDialog
    $openFile.Filter = "JSON (*.json)|*.json"
    $openFile.ShowDialog() | Out-Null
    
    return $openFile.FileName
}

# Provide a GUI interface for opening and uploading the Tenable WAS Report PDF export
function uploadTenableReport($initialDirectory, $reportName)
{
    $uploadFile = New-Object System.Windows.Forms.OpenFileDialog
    $uploadFile.Filter = "PDF (*.pdf)|*.pdf"
    $uploadFile.ShowDialog() | Out-Null

    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }
    
    try
    {
        $fileName = Split-Path -Path $uploadFile.FileName -Leaf
        $reportFolder = $tenableBaseUrl + "GetFolderByServerRelativeUrl" + "('" + $spoReportDocumentLibary + "/" + $reportName + "')/files/add(overwrite=true,url='$($fileName)')"
        $null = Invoke-RestMethod -Method POST -Uri $reportFolder -InFile $uploadFile.FileName -Headers $headers
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not upload PDF report to $($reportName) library!"
    }
}

# MFA to authenticate to the Azure application
function getAzureDeviceCode
{
    $deviceCodeRequestParams = @{
        Method = "POST"
        Uri = "https://login.microsoftonline.com/$tenantId/oauth2/devicecode"
        Body = @{
            client_id = $clientId
            resource = $resource
        }
    }

    $deviceCodeRequest = Invoke-RestMethod @deviceCodeRequestParams
    Write-Host -ForegroundColor Cyan $deviceCodeRequest.Message

    return $deviceCodeRequest
}

# Once we have MFA'ed we can request a token with delegate permissions (meaning the app runs as YOU) to the Azure application -- the Azure application requires both Graph and SharePoint permissions
function getAzureToken($azureDeviceCode)
{
    $tokenRequestParams = @{
        Method = "POST"
        Uri = "https://login.microsoftonline.com/$tenantId/oauth2/token"
        Body = @{
            grant_type = "urn:ietf:params:oauth:grant-type:device_code"
            code = $azureDeviceCode.device_code
            client_id = $clientId
        }
    }

    $tokenRequest = Invoke-RestMethod @tokenRequestParams

    return $tokenRequest
}

# This was the hardest part to figure out -- the Graph token does not play well with the token needed for the SharePoint REST API so we use the Graph token as the refresh token request to generate one that SharePoint likes and will use
function getSPOToken($azureToken)
{
    $spoTokenRequestParams = @{
        Method = "POST"
        Uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        Body = @{        
            refresh_token = $azureToken.refresh_token
            grant_type = "refresh_token"
            client_id = $clientId
            scope = "https://odemail.sharepoint.com/.default"
        }
    }

    $spoTokenRequest = Invoke-RestMethod @spoTokenRequestParams

    return $spoTokenRequest
}

# This function searches the Scan Schedule SharePoint List for the List item that matches the title of the scan in the JSON export -- it has logic to detect if there's multiple entries and builds a hashtable menu to allow selection
function getScheduledScan($scanName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }
    
    try
    {
        $getScheduledScans = Invoke-RestMethod -Method GET -Uri ($tenableBaseList + "GetByTitle('$spoScanScheduleList')/items") -Headers $headers -ContentType "application/json;odata=verbose"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not retrieve $($spoScanScheduleList) List!"
    }

    $getSpecificScan = $getScheduledScans | Where-Object { $_.content.properties.Title -eq $scanName }

    If ($getSpecificScan.Count -gt 1)
    {
        Write-Host -ForegroundColor Cyan "Multiple scans with the name $($scanName) found! Which one are we working with?"
        $i = 1
        $makeMenuHash = @{}
        
        ForEach ($scan in $getSpecificScan)
        {
            $addMe = "$($scan.content.properties.Title) - $($scan.content.properties.field_1)"            
            $makeMenuHash.Add($i, $addMe)
            $i++
        }
        
        $padLength = $makeMenuHash.Count.ToString().Length
        For ($i = 1; $i -le $makeMenuHash.Count; $i++)
        {               
            Write-Host -ForegroundColor DarkYellow "$($i.ToString().PadLeft($padLength))) $($makeMenuHash[$i])"
        }

        $getScanArrayIndex = $(Read-Host "Select number of scan")

        return $getSpecificScan[[int]$getScanArrayIndex - 1]
    }
    Else
    {
        return $getSpecificScan
    }
}

# Function to check if the scan already exists in Scan Summary List -- if it does we will delete it after copying it to the Scan History for that application to make room for the new scan results 
function getScanSummary($scanName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }
    
    try
    {
        $getScheduledScans = Invoke-RestMethod -Method GET -Uri ($tenableBaseList + "GetByTitle('$spoScanSummaryList')/items") -Headers $headers -ContentType "application/json;odata=verbose"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not retrieve $($spoScanScheduleList) list!"
    }

    $getSpecificScan = $getScheduledScans | Where-Object { $_.content.properties.Title -eq $scanName }

    If ($getSpecificScan)
    {
        Write-Host -ForegroundColor DarkCyan "Found $($scanName) in Scan Summary List! Will replace entry with new results..."
        return $getSpecificScan
    }
    Else
    {
        Write-Host -ForegroundColor DarkCyan "Could not find $($scanName) in Scan Summary List! Will make a new entry for it..." 
        return 0
    }
}

# Create the Vulnerability Tracker List if it does not exist and it shouldn't since we delete it in another function first
function checkCreateTenableSPOList($spoListName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }

    $body = @{
        AllowContentTypes = $true
        BaseTemplate = 100
        ContentTypesEnabled = $true        
        Title = $spoListName
    }

    $checkSpoList = ($tenableBaseList + "GetByTitle('$spoListName')")

    try 
    {         
        $null = Invoke-RestMethod -Method GET -Uri $checkSpoList -Headers $headers -ContentType "application/json"
        Write-Host -ForegroundColor Cyan "$($spoListName) List exists!"
        return 0
    }
    catch
    {
        Write-Host -ForegroundColor Red "$($spoListName) List does NOT exist! Create it?"
        $yesNo = $null
        while ($yesNo -ne 'n' -and $yesNo -ne 'y')
        {
            $yesNo = Read-Host -Prompt "[Y/N]"
            switch ($yesNo)
            {
                "y"
                {       
                    Write-Host -ForegroundColor Cyan "Creating List for $($spoListName)..."

                    try
                    {
                        $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseUrl + "lists") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)
                    }
                    catch
                    {
                        Write-Host -ForegroundColor Red "Could not create SharePoint List: $($spoListName)!"
                    }                    
                }
                "n"
                {
                    Write-Host -ForegroundColor DarkMagenta "Exiting..."
                    exit
                }                                       
            }
        }
    
        return 1
    }    
}

# PITA. SharePoint truncates/formats List URL's with no documentation on why/how. This was a very tricky way to generate a working clickable hyperlink for the Vulnerability Tracker List in the Scan Schedule that doesn't care if there's
# special characters in the List name or the length of the list name.
function getTenableSPOListURL($spoListName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"
        Accept = "application/json;odata=verbose"            
    }

    Write-Host -ForegroundColor Cyan "Generating URL for $($spoListName)..."    

    try
    {
        $getRequest = Invoke-WebRequest -Method GET -Uri ($tenableBaseList + "GetByTitle('$spoListName')?`$select=RootFolder/ServerRelativeUrl&`$expand=RootFolder") -Headers $headers -ContentType "application/json;odata=verbose"
        $getServerRelativeUrl = ($getRequest.Content | ConvertFrom-Json).d.RootFolder.ServerRelativeUrl
        $getUrl = $tenantDomain + $getServerRelativeUrl

        return $getUrl
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not generate URL for $($spoListName)!"
        return 0
    }    
}

# Add the Tenable WAS List Template Content Type to the newly created Vulnerability Tracker List we've made
function addContentType($spoListName, $listContentType)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"                    
    }

    If ($listContentType -eq 0)
    {
        $body = @{
            contentTypeId = $odeTenableWasContentType
        }
    }
    ElseIf ($listContentType -eq 1)
    {
        $body = @{
            contentTypeId = $odeTenableWasSummaryContentType
        }
    }            

    Write-Host -ForegroundColor Cyan "Adding Content Type for $($spoListName)..."

    try
    {
       $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$spoListName')/ContentTypes/AddAvailableContentType") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not set Content Type!"
    }    
}

# Remove the default "Item" Content Type from the newly created Vulnerability Tracker List
function removeContentType($spoListName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }    
    
    Write-Host -ForegroundColor Cyan "Removing Item Content Type from $($spoListName)..."

    try
    {
        $getItemId = (Invoke-RestMethod -Method GET -Uri ($tenableBaseList + "GetByTitle('$spoListName')/contenttypes?`$filter=Name eq 'Item'") -Headers $headers -ContentType "application/json").content.properties.Id.StringValue
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not retrieve ID of Item Content Type!"                
    }
    
    try
    {
        $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$spoListName')/contenttypes('$getItemId')/deleteObject()") -Headers $headers -ContentType "application/json"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not delete Item Content Type!"
    }    
}  

# This is the function we don't use to create the Vulnerability Tracker List fields -- I left it in for reference though in case we ever have a need for it elsewhere
function createTenableSPOListFields($spoListName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }

    ForEach ($tenableSPOListField in $odeTenableSPOList.GetEnumerator())
    {
        $body = @{
            Title = $tenableSPOListField.Name
            FieldTypeKind = $tenableSPOListField.Value
        }
        
        try
        {
            $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$spoListName')/fields") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)            
        }
        catch
        {
            Write-Host -ForegroundColor Red "Could not create SharePoint List field!"
        }
    }
}

# This function updated the hyperlink for either the "Scan Reports" column or the "Vulnerability Tracker" column depending on what's passed to it in $scanItemType
function updateScheduledScanItem($scheduledScan, $scanItemType)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"
        "If-Match" = "*"
    }

    $getTitle = $scheduledScan.content.properties.Title
        
    If ($scanItemType -eq 0)
    {
        Write-Host -ForegroundColor Cyan "Adding link to $($getTitle) Tenable WAS Report..."
        $getForm = $reportFolderBaseUrl + [uri]::EscapeDataString($getTitle)
        $body = "{ '__metadata': { 'type': 'SP.Data.$($spoListDataType)ListItem' }, 'ScanReports': { '__metadata': { 'type': 'SP.FieldUrlValue' }, 'Url': '$getForm' } }"        
    }
    ElseIf ($scanItemType -eq 1)
    {
        Write-Host -ForegroundColor Cyan "Adding link to $($getTitle) Tenable WAS Vulnerabilities Tracker..."
        $getUrl = getTenableSPOListURL $getTitle        
        $body = "{ '__metadata': { 'type': 'SP.Data.$($spoListDataType)ListItem' }, 'VulnerabilityTracker': { '__metadata': { 'type': 'SP.FieldUrlValue' }, 'Url': '$getUrl'} }"
    }
    Else
    {
        Write-Host -ForegroundColor Red "Not a known type!"        
    }

    try
    {
        $itemId = ($scheduledScan.content.properties.id | Select-Object -First 1)."#text"
        $null = Invoke-RestMethod -Method MERGE -Uri ($tenableBaseList + "GetByTitle('$spoScanScheduleList')/items(" + $itemId + ")") -Headers $headers -ContentType "application/json;odata=verbose" -Body $body
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not update Scheduled Scan Item!"
    }    
}

# Function to copy the current Scan Summary item for the application to the Scan History List for said application before deleting it to make room for the new results
function copyScanSummaryItem($summaryScan)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"
        "If-Match" = "*"
    }
    
    $getTitle = $summaryScan.content.properties.Title + " Scan History"

    $body = @{
        Title = $getTitle                
        "Last_x0020_Scan_x0020_Date" = $summaryScan.content.properties.Last_x0020_Scan_x0020_Date."#text"
        Status = $summaryScan.content.properties.Status
        "OData__x0023__x0020_of_x0020_CRITICAL" = $summaryScan.content.properties.OData__x0023__x0020_of_x0020_CRITICAL."#text"
        "OData__x0023__x0020_of_x0020_HIGH" = $summaryScan.content.properties.OData__x0023__x0020_of_x0020_HIGH."#text"
        "OData__x0023__x0020_of_x0020_MEDIUM" = $summaryScan.content.properties.OData__x0023__x0020_of_x0020_MEDIUM."#text"
        "Scan_x0020_Notes" = $summaryScan.content.properties.Scan_x0020_Notes."#text"
    }
        
    $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$getTitle')/items") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)
}

# Function to check if a Vulnerability Tracker List for an application exists and if so delete it to make room for the scan results   
function deleteVulnerabilityTrackerList($spoListName)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"
        "If-Match" = "*"
    }
    
    Write-Host -ForegroundColor Cyan "Checking if Vulnerability Tracker List exists for $($spoListName)..."

    try 
    {         
        $null = Invoke-RestMethod -Method GET -Uri ($tenableBaseUrl + "lists/GetByTitle('$spoListName')") -Headers $headers -ContentType "application/json;odata=verbose" 
        Write-Host -ForegroundColor Magenta "$($spoListName) Vulnerability Tracker List exists!"
    }
    catch
    {
        Write-Host -ForegroundColor Magenta "$($spoListName) has no Vulnerability Tracker List! Proceeding..."
        return
    }
               
    Write-Host -ForegroundColor Red "Deleting $($getTitle) Vulnerability Tracker List! Are you sure?"
    $yesNo = $null
    while ($yesNo -ne 'n' -and $yesNo -ne 'y')
    {
        $yesNo = Read-Host -Prompt "[Y/N]"
        switch ($yesNo)
        {
            "y"
             {       
                try
                {
                    $null = Invoke-RestMethod -Method DELETE -Uri ($tenableBaseList + "GetByTitle('$spoListName')") -Headers $headers -ContentType "application/json;odata=verbose"
                }
                catch
                {
                    Write-Host -ForegroundColor Red "Could not delete $($spoListName) Vulnerability Tracker List!"
                }
            }
            "n"
            {
                Write-Host -ForegroundColor Red "Nothing more to do!"
                exit
            }
        }
    }
}

# Function to delete the Scan Summary item once we have copied it so we can post the results from the latest scan
function deleteScanSummaryItem($summaryScan)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"
        "If-Match" = "*"
    }

    $getTitle = $summaryScan.content.properties.Title
    
    Write-Host -ForegroundColor Red "Deleting $($getTitle) Scan Summary item! Are you sure?"
    $yesNo = $null
    while ($yesNo -ne 'n' -and $yesNo -ne 'y')
    {
        $yesNo = Read-Host -Prompt "[Y/N]"
        switch ($yesNo)
        {
            "y"
             {       
                try
                {
                    $itemId = ($summaryScan.content.properties.id | Select-Object -First 1)."#text"
                    $null = Invoke-RestMethod -Method DELETE -Uri ($tenableBaseList + "GetByTitle('$spoScanSummaryList')/items(" + $itemId + ")") -Headers $headers -ContentType "application/json;odata=verbose"
                }
                catch
                {
                    Write-Host -ForegroundColor Red "Could not delete Scan Summary item!"
                }
            }
            "n"
            {
                Write-Host -ForegroundColor DarkRed "Aborting!"
            }
        }
    }    
}

# Function to unhide columns from the View of the Vulnerability Tracker List
function revealTenableSPOListFields($spoListName, $listContentType)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }

    Write-Host -ForegroundColor Cyan "Making Tenable WAS columns viewable in List..."

    If ($listContentType -eq 0)
    {
        $getListHashTable = $odeTenableSPOList
    }
    ElseIf ($listContentType -eq 1)
    {        
        $getListHashTable = $odeTenableSummaryList
    }
    Else
    {
        Write-Host -ForegroundColor Red "Did not recognize Content Type!"
    }

    ForEach ($tenableSPOListField in $getListHashTable.GetEnumerator())
    {
        try
        {
            $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$spoListName')/views/GetByTitle('All%20Items')/viewfields/addViewField('$($tenableSPOListField.Name)')") -Headers $headers -ContentType "application/json"
        }
        catch
        {
            Write-Host -ForegroundColor Red "Could not unhide SharePoint List field!"
        }        
    }    
}

# Function to check if the application Reports folder exists in the Document Library and create it if it does not
function checkCreateReportFolder($scheduledScan, $scanName)
{    
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"            
    }
    
    If ($scheduledScan)
    {
        Write-Host -ForegroundColor Magenta "Scheduled Scan found! Proceeding as normal..."
        $reportName = $scheduledScan.content.properties.Title
    }
    Else
    {
        Write-Host -ForegroundColor Magenta "No Scheduled Scan found! Treating as a 1-off..."
        $reportName = $scanName
    }
    
    $checkReportFolder = $tenableBaseUrl + "GetFolderByServerRelativeUrl" + "('" + $spoReportDocumentLibary + "/" + $reportName + "')"

    Write-Host -ForegroundColor Cyan "Checking if reports folder exists for $($reportName)..."

    try 
    {         
        $null = Invoke-RestMethod -Method GET -Uri $checkReportFolder -Headers $headers -ContentType "application/json;odata=verbose" 
        Write-Host -ForegroundColor Cyan "$($reportName) Reports folder exists!"
    }
    catch
    {
        Write-Host -ForegroundColor Red "$($reportName) Reports folder does NOT exist! Create it?"
        $yesNo = $null
        while ($yesNo -ne 'n' -and $yesNo -ne 'y')
        {
            $yesNo = Read-Host -Prompt "[Y/N]"
            switch ($yesNo)
            {
                "y"
                {
                    $createReportFolder = $tenableBaseUrl + "folders/add" + "('" + $spoReportDocumentLibary + "/" + $reportName + "')"
                    try
                    {
                        $null = Invoke-RestMethod -Method POST -Uri $createReportFolder -Headers $headers -ContentType "application/json;odata=verbose"
                        Write-Host -ForegroundColor Magenta "Created $($reportName) folder!"
                        If ($scheduledScan)
                        {
                            Write-Host -ForegroundColor Cyan "Update Scan Schedule with Report folder?"
                            $yesNo = $null
                            while ($yesNo -ne 'n' -and $yesNo -ne 'y')
                            {
                                $yesNo = Read-Host -Prompt "[Y/N]"
                                switch ($yesNo)
                                {
                                    "y"
                                    {                                    
                                        updateScheduledScanItem $scheduledScan 0                                    
                                    }
                                    "n"
                                    {
                                        Write-Host -ForegroundColor Red "Not updating Scan Schedule with Reports link!"
                                    }
                                }
                            }
                        }
                    }                                                        
                    catch
                    {
                        Write-Host -ForegroundColor Red "Unable to create $($createReportFolder)!"
                    }
                }
                "n"
                {
                    Write-Host -ForegroundColor DarkMagenta "Exiting..."
                    exit
                }
            }
        }
    }
}

# Function to post the Vulnerability Tracker List items
function addTenableItem($newTenableItem)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"        
    }

    $body = @{ 
        Title = $newTenableItem.Title       
        "Scan_x0020_Date" = $newTenableItem.ScanDate
        "Plugin_x0020_ID" = $newTenableItem.PluginID
        CVE = $newTenableItem.CVE        
        CVSSv3 = $newTenableItem.CVSSv3
        Risk = $newTenableItem.Risk
        URI = $newTenableItem.URI
        Synopsis = $newTenableItem.Synopsis
        Information = $newTenableItem.Information
        Solution = $newTenableItem.Solution  
        Remediated = $newTenableItem.Remediated
        Tracking = $newTenableItem.Tracking            
    }

    $getList = $newTenableItem.Title           
    $null = Invoke-RestMethod  -Method POST -Uri ($tenableBaseList + "GetByTitle('$getList')/items") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)
}

# Function to populate the values from the Tenable WAS Report JSON that will go into the Vulnerability Tracker List
function addVulnerabilities($importTenable)
{
    $getTitle = $importTenable.config.name
    $getDate = Get-Date -Date $importTenable.scan.started_at -Format "o"
    $getTextInfo = (Get-Culture).TextInfo

    ForEach ($tenableFinding in $importTenable.findings)
    {  
        Write-Host -ForegroundColor Cyan "$($tenableFinding.risk_factor) vulnerability detected! Processing..."
    
        $newTenableItem = [ODETenableItem]::new()
        $newTenableItem.Title = [string]$getTitle
        $newTenableItem.ScanDate = [datetime]$getDate
        $newTenableItem.PluginID = [string]$tenableFinding.plugin_id
        $newTenableItem.CVE = [string]$tenableFinding.cves     
        $newTenableItem.CVSSv3 = [double]$tenableFinding.cvssv3
        $newTenableItem.Risk = [string]$getTextInfo.ToTitleCase($tenableFinding.risk_factor)
        $newTenableItem.URI = [string]$tenableFinding.uri
        $newTenableItem.Synopsis = [string]$tenableFinding.synopsis
        $newTenableItem.Information = [string]$tenableFinding.description
        $newTenableItem.Solution = [string]$tenableFinding.solution    
    
        If ($tenableFinding.risk_factor -eq "info")
        {        
            $newTenableItem.Remediated = "N/A"
        }
        Else
        {
            $newTenableItem.Remediated = $null
        }
       
        addTenableItem $newTenableItem
    }
}

# Function to populate the values to the Scan Summary List from our function that calculates information from Tenable WAS Report JSON
function addTenableSummaryItem($newTenableSummaryItem)
{
    $headers = @{
        Authorization = "Bearer $($spoToken.access_token)"        
    }

    $body = @{ 
        Title = $newTenableSummaryItem.Title       
        "Last_x0020_Scan_x0020_Date" = $newTenableSummaryItem.LastScanDate
        Status = $newTenableSummaryItem.Status
        "OData__x0023__x0020_of_x0020_CRITICAL" = $newTenableSummaryItem.CriticalCount
        "OData__x0023__x0020_of_x0020_HIGH" = $newTenableSummaryItem.HighCount
        "OData__x0023__x0020_of_x0020_MEDIUM" = $newTenableSummaryItem.MediumCount
        "Scan_x0020_History" = @{
            "Url" = $newTenableSummaryItem.ScanHistory
        }        
    }
        
    $null = Invoke-RestMethod -Method POST -Uri ($tenableBaseList + "GetByTitle('$spoScanSummaryList')/items") -Headers $headers -ContentType "application/json" -Body ($body | ConvertTo-Json)
}

# Function to calculate values for the Tenable Scan Summary List
function addScanSummary($importTenable)
{    
    $getTitle = $importTenable.config.name
    $getDate = Get-Date -Date $importTenable.scan.started_at -Format "o"    
    $criticalCount = 0
    $highCount = 0
    $mediumCount = 0

    ForEach ($tenableFinding in $importTenable.findings)
    {          
        switch ($tenableFinding.risk_factor)
        {
            "critical"
            {
                $criticalCount++
            }
            "high"
            {
                $highCount++
            }
            "medium"
            {
                $mediumCount++
            }
        }
    }
        
    $newTenableSummaryItem = [ODETenableSummaryItem]::new()
    $newTenableSummaryItem.Title = [string]$getTitle
    $newTenableSummaryItem.LastScanDate = [datetime]$getDate

    If (($criticalCount -eq 0) -and ($highCount -eq 0))
    {
        $newTenableSummaryItem.Status = "Passed"
    }
    Else
    {
        $newTenableSummaryItem.Status = "In Review"
    }

    $newTenableSummaryItem.CriticalCount = $criticalCount
    $newTenableSummaryItem.HighCount = $highCount
    $newTenableSummaryItem.MediumCount = $mediumCount
    $newTenableSummaryItem.ScanHistory = getTenableSPOListURL ($getTitle + " Scan History")
           
    addTenableSummaryItem $newTenableSummaryItem
}

odeSplash

$tenableReport = openTenableReport $PWD.Path

try 
{
    $importTenable = (Get-Content -Path $tenableReport) | ConvertFrom-Json
}
catch 
{
    Write-Host -ForegroundColor Red "Error importing JSON! Exiting..."
    exit
}

Write-Host -ForegroundColor Magenta "ODE Tenable WAS Automation Script $($scriptVersion) - Benjamin Barshaw " -NoNewline
Write-Host -ForegroundColor DarkGray "<" -NoNewline
Write-Host -ForegroundColor Cyan "benjamin.barshaw@ode.oregon.gov" -NoNewline
Write-Host -ForegroundColor DarkGray ">"

Write-Host -ForegroundColor Cyan "Is this a scheduled scan or a 1-off scan?"

$makeMenuHash = @{}
$makeMenuHash.Add(1, "Scheduled Scan")
$makeMenuHash.Add(2, "1-off Scan")                
$padLength = $makeMenuHash.Count.ToString().Length
For ($i = 1; $i -le $makeMenuHash.Count; $i++)
{               
    Write-Host -ForegroundColor DarkYellow "$($i.ToString().PadLeft($padLength))) $($makeMenuHash[$i])"
}

$getScanType = [int]$(Read-Host "Select number of scan")
$getScanType--

$azureDeviceCode = getAzureDeviceCode

Write-Host -ForegroundColor DarkYellow "Press enter key only when you have authenticated"
$null = Read-Host "Waiting..."

$azureToken = getAzureToken $azureDeviceCode

$spoToken = getSPOToken $azureToken

$getTitle = $importTenable.config.name 

If ($getScanSummary = getScanSummary $getTitle)
{
    copyScanSummaryItem $getScanSummary
    deleteScanSummaryItem $getScanSummary
}

# Originally the menu would not add a Scan History if 1-off was selected -- I've removed this check as we did want to view this information even for 1-offs

#If ($getType -eq 0) 
#{
    $addScanHistory = $getTitle + " Scan History"

    If (checkCreateTenableSPOList $addScanHistory)
    {
        addContentType $addScanHistory 1

        revealTenableSPOListFields $addScanHistory 1

        removeContentType $addScanHistory
    }
#}

addScanSummary $importTenable

$getScheduledScan = getScheduledScan $getTitle

checkCreateReportFolder $getScheduledScan $getTitle

uploadTenableReport $PWD.Path $getTitle

deleteVulnerabilityTrackerList $getTitle

If (checkCreateTenableSPOList $getTitle)
{
    If ($getscanType -eq 0)
    {
        updateScheduledScanItem $getScheduledScan 1
    }
    Else
    {
        Write-Host -ForegroundColor Magenta "1-off scan detected! Skipping updating Scheduled Scan List..."
    }

    addContentType $getTitle 0

    revealTenableSPOListFields $getTitle 0

    removeContentType $getTitle

    addVulnerabilities $importTenable
}