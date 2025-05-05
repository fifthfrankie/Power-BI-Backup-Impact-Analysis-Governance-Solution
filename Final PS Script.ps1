# Define the base folder and script paths at the beginning of the script

$baseFolderPath = "C:\Power BI Backups"
$Script1Path = "$baseFolderPath\Config\Report Detail Extract Script.csx"
$Script2Path = "$baseFolderPath\Config\Model Detail Extract Script.csx"
$Script3Path = "$baseFolderPath\Config\Measure Dependency Extract Script.csx"
$TabularEditor2Path = "$baseFolderPath\Config\TabularEditor\TabularEditor.exe"

# Enter Workspace ID between quotation marks if you only want script to run in 1 or 2 workspaces. Leave BOTH empty if you want to loop through all.

$SpecificWorkspaceID1 = ""  # Replace with your actual workspace ID or leave empty and the script will loop through every workspace

$SpecificWorkspaceID2 = ""  # Replace with your actual workspace ID or leave empty and the script will loop through every workspace


$ErrorActionPreference = "SilentlyContinue"

# Check for PBI Tools folder and extract if necessary
$PBIToolsFolderPath = Join-Path -Path $baseFolderPath\Config -ChildPath "PBI Tools"
$PBIToolsZipPattern = "pbi-tools*.zip"

if (-not (Test-Path -Path $PBIToolsFolderPath)) {
    # Check if there is any zip file matching the pattern
    $zipFile = Get-ChildItem -Path $baseFolderPath\Config -Filter $PBIToolsZipPattern | Select-Object -First 1
    if ($zipFile) {
        # Create the PBI Tools folder
        New-Item -Path $PBIToolsFolderPath -ItemType Directory -Force
        # Extract the zip file contents into the PBI Tools folder
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $PBIToolsFolderPath)
        Write-Output "Extracted $($zipFile.Name) to $PBIToolsFolderPath"
    }
}

# Define the folder path and zip pattern for Tabular Editor
$TabularEditorFolderPath = Join-Path -Path $baseFolderPath\Config -ChildPath "TabularEditor"
$TabularEditorZipPattern = "TabularEditor*.zip"

if (-not (Test-Path -Path $TabularEditorFolderPath)) {
    # Check if there is any zip file matching the pattern
    $zipFile = Get-ChildItem -Path $baseFolderPath\Config -Filter $TabularEditorZipPattern | Select-Object -First 1
    if ($zipFile) {
        # Create the Tabular Editor folder
        New-Item -Path $TabularEditorFolderPath -ItemType Directory -Force
        # Extract the zip file contents into the Tabular Editor folder
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $TabularEditorFolderPath)
        Write-Output "Extracted $($zipFile.Name) to $TabularEditorFolderPath"
    } else {
        Write-Output "No zip file matching the pattern $TabularEditorZipPattern found in $baseFolderPath\Config"
    }
} else {
    Write-Output "Tabular Editor folder already exists at $TabularEditorFolderPath"
}


# Add the PBI Tools folder to the PATH environment variable only if it's not already there
$existingPath = [System.Environment]::GetEnvironmentVariable("PATH", [System.EnvironmentVariableTarget]::User)
if ($existingPath -notlike "*$PBIToolsFolderPath*") {
    $env:PATH += ";$PBIToolsFolderPath"
    [System.Environment]::SetEnvironmentVariable("PATH", $env:PATH, [System.EnvironmentVariableTarget]::User)
    Write-Output "Added $PBIToolsFolderPath to the PATH environment variable."
} else {
    Write-Output "$PBIToolsFolderPath is already in the PATH environment variable."
}

# Temporarily set execution policy to Bypass for this session
if ((Get-ExecutionPolicy) -ne 'Bypass') {
    Write-Host "Temporarily setting Execution Policy to Bypass for this session..."
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}

# Ensure required modules are installed, and imports them. If import fails, error and exit early
$requiredModules = @( 'ImportExcel', 'MicrosoftPowerBIMgmt' )
foreach ($module in $requiredModules) {
    if( -not (Import-Module $module -PassThru -EA ignore) ) {
       Install-Module -Name $module -Scope CurrentUser -Force
    }

    Import-Module $Module -ErrorAction 'stop' # In the rare case Install-Module fails, you probably want a terminating error
}


# Connect to the Power BI Service
function Connect-PowerBI {
    Connect-PowerBIServiceAccount
    $global:accessTokenObject = Get-PowerBIAccessToken
    $global:accessToken = $accessTokenObject.Authorization -replace 'Bearer ', ''
    # Write the access token to a temporary file
    Set-Content -Path $env:TEMP\PowerBI_AccessToken.txt -Value $global:accessToken
}

# Track script start time
$scriptStartTime = Get-Date
Connect-PowerBI

# Function to refresh the token in a background job
function Start-TokenRefreshJob {
    $jobScript = {
        function Connect-PowerBI {
            Connect-PowerBIServiceAccount
            $global:accessTokenObject = Get-PowerBIAccessToken
            $global:accessToken = $accessTokenObject.Authorization -replace 'Bearer ', ''
            # Write the access token to a temporary file
            Set-Content -Path $env:TEMP\PowerBI_AccessToken.txt -Value $global:accessToken
        }
        while ($true) {
            Start-Sleep -Seconds 3300  # Sleep for 55 minutes
            Connect-PowerBI
        }
    }
    Start-Job -ScriptBlock $jobScript -Name "TokenRefreshJob"
}

# Start the background job to refresh the token
Start-TokenRefreshJob

# Function to get the current access token
function Get-CurrentAccessToken {
    $global:accessToken = Get-Content -Path $env:TEMP\PowerBI_AccessToken.txt
    return $global:accessToken
}

# Create a variable date
$date = (Get-Date -UFormat "%Y-%m-%d")








#### Start of Power BI Environment Detail Extract ####








# Define the Information Extract Excel file path
$excelFile = "$baseFolderPath\Power BI Environment Detail.xlsx"

# Function to rename properties in objects and handle duplicates
function Rename-Properties {
    param ($object, $renameMap)
    $newObject = New-Object PSObject
    foreach ($originalName in $renameMap.Keys) {
        $newPropertyName = $renameMap[$originalName]
        $propertyValue = if ($object.PSObject.Properties[$originalName]) { $object.$originalName } else { $null }
        if ($newObject.PSObject.Properties[$newPropertyName]) { $newPropertyName += "_duplicate" }
        $newObject | Add-Member -MemberType NoteProperty -Name $newPropertyName -Value $propertyValue
    }
    foreach ($property in $object.PSObject.Properties) {
        if (-not $renameMap.ContainsKey($property.Name)) {
            $newObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }
    return $newObject
}

# Define renaming maps for each type of object
$workspaceRenameMap = @{
    "id" = "WorkspaceId";
    "name" = "WorkspaceName";
    "isReadOnly" = "WorkspaceIsReadOnly";
    "isOnDedicatedCapacity" = "WorkspaceIsOnDedicatedCapacity";
    "capacityId" = "WorkspaceCapacityId";
    "defaultDatasetStorageFormat" = "WorkspaceDefaultDatasetStorageFormat";
    "type" = "WorkspaceType"
}

$datasetRenameMap = @{
    "id" = "DatasetId";
    "name" = "DatasetName";
    "description" = "DatasetDescription";
    "webUrl" = "DatasetWebUrl";
    "addRowsAPIEnabled" = "DatasetAddRowsAPIEnabled";
    "configuredBy" = "DatasetConfiguredBy";
    "isRefreshable" = "DatasetIsRefreshable";
    "isEffectiveIdentityRequired" = "DatasetIsEffectiveIdentityRequired";
    "isEffectiveIdentityRolesRequired" = "DatasetIsEffectiveIdentityRolesRequired";
    "isOnPremGatewayRequired" = "DatasetIsOnPremGatewayRequired";
    "targetStorageMode" = "DatasetTargetStorageMode";
    "queryScaleOutSettings" = "DatasetQueryScaleOutSettings";
    "createdDate" = "DatasetCreatedDate"
}

$datasetDatasourceRenameMap = @{
    "datasourceType" = "DatasetDatasourceType";
    "datasourceId" = "DatasetDatasourceId";
    "gatewayId" = "DatasetDatasourceGatewayId";
    "connectionDetails" = "DatasetDatasourceConnectionDetails"
}

$dataflowDatasourceRenameMap = @{
    "datasourceType" = "DataflowDatasourceType";
    "datasourceId" = "DataflowDatasourceId";
    "gatewayId" = "DataflowDatasourceGatewayId";
    "connectionDetails" = "DataflowDatasourceConnectionDetails"
}

$datasetRefreshRenameMap = @{
    "requestId" = "DatasetRefreshRequestId";
    "id" = "DatasetRefreshId";
    "startTime" = "DatasetRefreshStartTime";
    "endTime" = "DatasetRefreshEndTime";
    "status" = "DatasetRefreshStatus";
    "refreshType" = "DatasetRefreshType"
}

$dataflowRefreshRenameMap = @{
    "requestId" = "DataflowRefreshRequestId";
    "id" = "DataflowRefreshId";
    "startTime" = "DataflowRefreshStartTime";
    "endTime" = "DataflowRefreshEndTime";
    "status" = "DataflowRefreshStatus" ;
    "refreshType" = "DataflowRefreshType" ;
    "errorInfo" = "DataflowErrorInfo"
}

$dataflowRenameMap = @{
    "configuredBy"      = "DataflowConfiguredBy";
    "description"       = "DataflowDescription";
    "modelUrl"         = "DataflowJsonURL";
    "modifiedBy"       = "DataflowModifiedBy";
    "modifiedDateTime" = "DataflowModifiedDateTime";
    "name"             = "DataflowName";
    "objectId"         = "DataflowId";
    "generation" = "DataflowGeneration"
}

$dataflowLineageRenameMap = @{
    "datasetObjectId"   = "DatasetId";
    "dataflowObjectId"  = "DataflowId";
    "workspaceObjectId" = "WorkspaceId"
}

$reportRenameMap = @{
    "id" = "ReportId";
    "name" = "ReportName";
    "description" = "ReportDescription";
    "webUrl" = "ReportWebUrl";
    "embedUrl" = "ReportEmbedUrl";
    "isFromPbix" = "ReportIsFromPbix";
    "isOwnedByMe" = "ReportIsOwnedByMe";
    "datasetId" = "DatasetId";
    "datasetWorkspaceId" = "DatasetWorkspaceId";
    "reportType" = "ReportType"
}

$pageRenameMap = @{
    "name" = "PageName";
    "displayName" = "PageDisplayName";
    "order" = "PageOrder"
}

$appRenameMap = @{
    "id" = "AppId";
    "name" = "AppName";
    "lastUpdate" = "AppLastUpdate";
    "description" = "AppDescription";
    "publishedBy" = "AppPublishedBy";
    "workspaceId" = "AppWorkspaceId";
    "users" = "AppUsers"
}

$appReportRenameMap = @{
    "id" = "AppReportId";
    "reportType" = "AppReportType";
    "name" = "ReportName";
    "webUrl" = "AppReportWebUrl";
    "embedUrl" = "AppReportEmbedUrl";
    "isOwnedByMe" = "AppReportIsOwnedByMe";
    "datasetId" = "AppReportDatasetId";
    "originalReportObjectId" = "ReportId";
    "users" = "AppUsers";
    "subscriptions" = "AppReportSubscriptions";
    "sections" = "AppReportSections"
}

# Fetch and filter workspaces
$workspacesUrl = "https://api.powerbi.com/v1.0/myorg/groups"
$workspacesResponse = Invoke-PowerBIRestMethod -Method GET -Url $workspacesUrl | ConvertFrom-Json
$workspacesInfo = @()

foreach ($workspace in $workspacesResponse.value) {
    # Check if we should use specific workspace IDs for filtering
# If only SpecificWorkspaceID1 is provided, filter on that alone
if ($SpecificWorkspaceID1 -and -not $SpecificWorkspaceID2 -and $workspace.id -ne $SpecificWorkspaceID1) {
    continue
}
# If only SpecificWorkspaceID2 is provided, filter on that alone
elseif ($SpecificWorkspaceID2 -and -not $SpecificWorkspaceID1 -and $workspace.id -ne $SpecificWorkspaceID2) {
    continue
}
# If both workspace IDs are provided, filter based on either ID
elseif ($SpecificWorkspaceID1 -and $SpecificWorkspaceID2 -and 
        ($workspace.id -ne $SpecificWorkspaceID1 -and $workspace.id -ne $SpecificWorkspaceID2)) {
    continue
}
    # Add the workspace to workspacesInfo if it passes the checks
    $workspacesInfo += Rename-Properties -object $workspace -renameMap $workspaceRenameMap
}

# Initialize collections for all necessary information
$datasetsInfo = @()
$datasetSourcesInfo = @()
$dataflowsInfo = @()
$dataflowLineage = @()
$dataflowSourcesInfo = @()
$reportsInfo = @()
$reportPagesInfo = @()
$appsInfo = @()
$reportsInAppInfo = @()
$datasetNameLookup = @{}
$dataflowNameLookup = @{}
$datasetRefreshHistory = @()
$dataflowRefreshHistory = @()

# Loop through filtered workspaces
foreach ($workspace in $workspacesInfo) {
    # Fetch datasets
    $datasetsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/datasets"
    $datasets = Invoke-PowerBIRestMethod -Method GET -Url $datasetsUrl | ConvertFrom-Json

    foreach ($dataset in $datasets.value) {
        $renamedDataset = Rename-Properties -object $dataset -renameMap $datasetRenameMap
        $renamedDataset | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
        $renamedDataset | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
        
    # Store the DatasetId and DatasetName in the lookup table
        $datasetNameLookup[$dataset.id] = $dataset.name
        $datasetsInfo += $renamedDataset

        # Fetch dataset sources
        $datasourcesUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/datasets/$($dataset.id)/datasources"
        $datasources = Invoke-PowerBIRestMethod -Method GET -Url $datasourcesUrl | ConvertFrom-Json

        foreach ($datasource in $datasources.value) {
            $renamedDatasource = Rename-Properties -object $datasource -renameMap $datasetDatasourceRenameMap
            $renamedDatasource | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
            $renamedDatasource | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
            $renamedDatasource | Add-Member -NotePropertyName "DatasetId" -NotePropertyValue $dataset.id -Force
            $renamedDatasource | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $dataset.name -Force
            if ($datasource.connectionDetails) {
                $renamedDatasource.DatasetDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
            }
            $datasetSourcesInfo += $renamedDatasource
        }
    }

    # Fetch reports
    $reportsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/reports"
    $reports = Invoke-PowerBIRestMethod -Method GET -Url $reportsUrl | ConvertFrom-Json

	# Create a hash set to store Report IDs
	$reportIds = @{}

    foreach ($report in $reports.value) {
        $renamedReport = Rename-Properties -object $report -renameMap $reportRenameMap
        $renamedReport | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
        $renamedReport | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force


        # Retrieve and add the correct DatasetName from the lookup table if DatasetId exists
        $datasetId = $report.datasetId
        if ($datasetId -and $datasetNameLookup.ContainsKey($datasetId)) {
            $renamedReport | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $datasetNameLookup[$datasetId] -Force
        } else {
            $renamedReport | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue "Unknown Dataset" -Force
        }

        $reportsInfo += $renamedReport

        # Fetch report pages
        $pagesUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/reports/$($report.id)/pages"
        $pages = Invoke-PowerBIRestMethod -Method GET -Url $pagesUrl | ConvertFrom-Json
        foreach ($page in $pages.value) {
            $renamedPage = Rename-Properties -object $page -renameMap $pageRenameMap
            $renamedPage | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
            $renamedPage | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
            $renamedPage | Add-Member -NotePropertyName "ReportId" -NotePropertyValue $report.id -Force
            $renamedPage | Add-Member -NotePropertyName "ReportName" -NotePropertyValue $report.name -Force
            $reportPagesInfo += $renamedPage

            # Store the report ID in the hash set
            $reportIds[$report.id] = $true
        }
    }
}

# Fetch Apps and App Reports that are in filtered workspaces
$appsUrl = "https://api.powerbi.com/v1.0/myorg/apps"
$apps = Invoke-PowerBIRestMethod -Method GET -Url $appsUrl | ConvertFrom-Json

# Create a hash set to store App Report IDs
$appReportIds = @{}
$originalReportObjectIds = @{}

foreach ($app in $apps.value) {
    if ($workspacesInfo.WorkspaceId -contains $app.workspaceId) {
        $renamedApp = Rename-Properties -object $app -renameMap $appRenameMap
        $appsInfo += $renamedApp

        # Fetch reports within each app
        $appReportsUrl = "https://api.powerbi.com/v1.0/myorg/apps/$($app.id)/reports"
        $appReports = Invoke-PowerBIRestMethod -Method GET -Url $appReportsUrl | ConvertFrom-Json

        foreach ($report in $appReports.value) {
            $renamedAppReport = Rename-Properties -object $report -renameMap $appReportRenameMap
            $renamedAppReport | Add-Member -NotePropertyName "AppId" -NotePropertyValue $app.id -Force
            $renamedAppReport | Add-Member -NotePropertyName "AppName" -NotePropertyValue $app.name -Force
            $reportsInAppInfo += $renamedAppReport

            # Store the app report ID in the hash set
            $appReportIds[$report.id] = $true
            $originalReportObjectIds[$report.originalReportObjectId] = $true
        }
    }
}


# Fetch Refresh History for Datasets
foreach ($workspace in $workspacesInfo) {
    foreach ($dataset in $datasetsInfo | Where-Object { $_.WorkspaceId -eq $workspace.WorkspaceId }) {
        $refreshHistoryUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/datasets/$($dataset.DatasetId)/refreshes"
        $refreshHistoryResponse = Invoke-PowerBIRestMethod -Method GET -Url $refreshHistoryUrl | ConvertFrom-Json

        foreach ($refresh in $refreshHistoryResponse.value) {
            $renamedRefreshRecord = Rename-Properties -object $refresh -renameMap $datasetRefreshRenameMap
            $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
            $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
            $renamedRefreshRecord | Add-Member -NotePropertyName "DatasetId" -NotePropertyValue $dataset.DatasetId -Force
            $renamedRefreshRecord | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $dataset.DatasetName -Force

            $datasetRefreshHistory += $renamedRefreshRecord
        }
    }
}  


# Fetch Dataflows for Workspaces
foreach ($workspace in $workspacesInfo) {
    $dataflowsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/dataflows"
    $dataflowsResponse = Invoke-PowerBIRestMethod -Method GET -Url $dataflowsUrl | ConvertFrom-Json

    # Ensure response is not null before looping
    if ($dataflowsResponse.value) {
        foreach ($dataflow in $dataflowsResponse.value) {
            $renamedDataflow = Rename-Properties -object $dataflow -renameMap $dataflowRenameMap
            $renamedDataflow | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
            $renamedDataflow | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force

            # Store DataflowId and DataflowName in a lookup table
            if ($dataflow.objectId) {  
                $dataflowNameLookup[$dataflow.objectId] = $dataflow.name  
            }

            $dataflowsInfo += $renamedDataflow

            # Fetch Dataflow Datasources
            $dataflowSourcesUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/dataflows/$($dataflow.objectId)/datasources"
            $dataflowSourcesResponse = Invoke-PowerBIRestMethod -Method GET -Url $dataflowSourcesUrl | ConvertFrom-Json

            # Ensure response is not null before looping
            if ($dataflowSourcesResponse.value) {
                foreach ($datasource in $dataflowSourcesResponse.value) {
                    $renamedDataflowDatasource = Rename-Properties -object $datasource -renameMap $dataflowDatasourceRenameMap

                    # Ensure required fields are included
                    $renamedDataflowDatasource | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
                    $renamedDataflowDatasource | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
                    $renamedDataflowDatasource | Add-Member -NotePropertyName "DataflowId" -NotePropertyValue $dataflow.objectId -Force

            		if ($datasource.connectionDetails) {
               		 $renamedDataflowDatasource.DataflowDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
            }
                    

                    if ($dataflowNameLookup.ContainsKey($dataflow.objectId)) {
                        $renamedDataflowDatasource | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue $dataflowNameLookup[$dataflow.objectId] -Force
                    } else {
                        $renamedDataflowDatasource | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue "Unknown Dataflow" -Force
                    }

                    # Store in collection
                    $dataflowSourcesInfo += $renamedDataflowDatasource
                }
            }
        }
    }
} 

# Fetch Dataflow Lineage (Upstream Dataflows)
foreach ($workspace in $workspacesInfo) {
        $dataflowLineageUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/datasets/upstreamDataflows"
        $dataflowLineageResponse = Invoke-PowerBIRestMethod -Method GET -Url $dataflowLineageUrl | ConvertFrom-Json

        # Ensure response is not null before looping
        if ($dataflowLineageResponse.value) {
            foreach ($dataflow in $dataflowLineageResponse.value) {
                $renamedDataflow = Rename-Properties -object $dataflow -renameMap $dataflowLineageRenameMap
                $renamedDataflow | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
                $renamedDataflow | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force


                $dataflowId = $dataflow.dataflowObjectId  
                if ($dataflowId -and $dataflowNameLookup.ContainsKey($dataflowId)) {
                    $renamedDataflow | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue $dataflowNameLookup[$dataflowId] -Force
                } else {
                    $renamedDataflow | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue "Unknown Dataflow" -Force
                }

            $datasetId = $dataflow.datasetObjectId  
            if ($datasetId -and $datasetNameLookup.ContainsKey($datasetId)) {
                $renamedDataflow | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $datasetNameLookup[$datasetId] -Force
            } else {
                $renamedDataflow | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue "Unknown Dataset" -Force
            }

                $dataflowLineage += $renamedDataflow
            }
        }
    }


# Fetch Dataflow Refresh History
foreach ($workspace in $workspacesInfo) {
    foreach ($dataflow in $dataflowsInfo | Where-Object { $_.WorkspaceId -eq $workspace.WorkspaceId }) {
        $refreshHistoryUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.WorkspaceId)/dataflows/$($dataflow.DataflowId)/transactions"
        $refreshHistoryResponse = Invoke-PowerBIRestMethod -Method GET -Url $refreshHistoryUrl | ConvertFrom-Json

        # Ensure response is not null before looping
        if ($refreshHistoryResponse.value) {
            foreach ($refresh in $refreshHistoryResponse.value) {
                $renamedRefreshRecord = Rename-Properties -object $refresh -renameMap $dataflowRefreshRenameMap
                $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspace.WorkspaceId -Force
                $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspace.WorkspaceName -Force
                $renamedRefreshRecord | Add-Member -NotePropertyName "DataflowId" -NotePropertyValue $dataflow.DataflowId -Force
                

                if ($dataflowNameLookup.ContainsKey($dataflow.DataflowId)) {
                    $renamedRefreshRecord | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue $dataflowNameLookup[$dataflow.DataflowId] -Force
                } else {
                    $renamedRefreshRecord | Add-Member -NotePropertyName "DataflowName" -NotePropertyValue "Unknown Dataflow" -Force
                }

                # Store in collection
                $dataflowRefreshHistory += $renamedRefreshRecord
            }
        }
    }
}









#### Start of 'My Workspace' detail extract ####

# Check if either variable is filled out, if so, skip this section
if (-not $SpecificWorkspaceID1 -and -not $SpecificWorkspaceID2) {

# Define "My Workspace" constants
$myWorkspaceId = "My Workspace"
$myWorkspaceName = "My Workspace"

# Manually add "My Workspace" breakdown to workspacesInfo
$myWorkspaceDetails = [PSCustomObject]@{
    WorkspaceId                 = $myWorkspaceId
    WorkspaceName               = $myWorkspaceName
    WorkspaceType               = "Workspace"
    WorkspaceIsReadOnly         = $false
    WorkspaceIsOnDedicatedCapacity = $false
}
$workspacesInfo += $myWorkspaceDetails

# Fetch datasets from "My Workspace"
$myWorkspaceDatasetsUrl = "https://api.powerbi.com/v1.0/myorg/datasets"
$myWorkspaceDatasets = Invoke-PowerBIRestMethod -Method GET -Url $myWorkspaceDatasetsUrl | ConvertFrom-Json

foreach ($dataset in $myWorkspaceDatasets.value) {
    $renamedDataset = Rename-Properties -object $dataset -renameMap $datasetRenameMap
    $renamedDataset | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $myWorkspaceId -Force
    $renamedDataset | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $myWorkspaceName -Force

    # Store the DatasetId and DatasetName in the lookup table
    $datasetNameLookup[$dataset.id] = $dataset.name
    $datasetsInfo += $renamedDataset

    # Fetch dataset sources
    $datasourcesUrl = "https://api.powerbi.com/v1.0/myorg/datasets/$($dataset.id)/datasources"
    $datasources = Invoke-PowerBIRestMethod -Method GET -Url $datasourcesUrl | ConvertFrom-Json

    foreach ($datasource in $datasources.value) {
        $renamedDatasource = Rename-Properties -object $datasource -renameMap $datasetDatasourceRenameMap
        $renamedDatasource | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $myWorkspaceId -Force
        $renamedDatasource | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $myWorkspaceName -Force
        $renamedDatasource | Add-Member -NotePropertyName "DatasetId" -NotePropertyValue $dataset.id -Force
        $renamedDatasource | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $dataset.name -Force
        if ($datasource.connectionDetails) {
            $renamedDatasource.DatasetDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
        }
        $datasetSourcesInfo += $renamedDatasource
    }
}

# Fetch reports from "My Workspace"
$myWorkspaceReportsUrl = "https://api.powerbi.com/v1.0/myorg/reports"
$myWorkspaceReports = Invoke-PowerBIRestMethod -Method GET -Url $myWorkspaceReportsUrl | ConvertFrom-Json

# Flag to track if any shared report exists
$sharedReportExists = $false

foreach ($report in $myWorkspaceReports.value) {
    # Skip reports that exist in either the Report list or the App Report list
    if ($appReportIds.ContainsKey($report.id) -or $reportIds.ContainsKey($report.id) -or $originalReportObjectIds.ContainsKey($report.id)) {
        continue
    }

    # Check if the report is owned by me
    if ($report.isOwnedByMe -eq $false) {
        $workspaceIdValue = "Shared Reports (No Workspace Access)"
        $workspaceNameValue = "Shared Reports (No Workspace Access)"
        $sharedReportExists = $true  # Set flag if a shared report is found
    } else {
        $workspaceIdValue = $myWorkspaceId
        $workspaceNameValue = $myWorkspaceName
    }

    $renamedReport = Rename-Properties -object $report -renameMap $reportRenameMap
    $renamedReport | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspaceIdValue -Force
    $renamedReport | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspaceNameValue -Force

    # Retrieve and add the correct DatasetName from the lookup table if DatasetId exists
    $datasetId = $report.datasetId
    if ($datasetId -and $datasetNameLookup.ContainsKey($datasetId)) {
        $renamedReport | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $datasetNameLookup[$datasetId] -Force
    } else {
        $renamedReport | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue "Unknown Dataset" -Force
    }

    $reportsInfo += $renamedReport

    # Fetch report pages
    $pagesUrl = "https://api.powerbi.com/v1.0/myorg/reports/$($report.id)/pages"
    $pages = Invoke-PowerBIRestMethod -Method GET -Url $pagesUrl | ConvertFrom-Json
    foreach ($page in $pages.value) {
        $renamedPage = Rename-Properties -object $page -renameMap $pageRenameMap
        $renamedPage | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $workspaceIdValue -Force
        $renamedPage | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $workspaceNameValue -Force
        $renamedPage | Add-Member -NotePropertyName "ReportId" -NotePropertyValue $report.id -Force
        $renamedPage | Add-Member -NotePropertyName "ReportName" -NotePropertyValue $report.name -Force
        $reportPagesInfo += $renamedPage
    }
}

# After processing all reports, add a single row to workspacesInfo if at least one shared report exists

if ($sharedReportExists) {
    # Define "Shared Reports (My Workspace)" constants
    $sharedWorkspaceId = "Shared Reports (No Workspace Access)"
    $sharedWorkspaceName = "Shared Reports (No Workspace Access)"

    # Add one entry to workspacesInfo for all shared reports
    $sharedWorkspaceDetails = [PSCustomObject]@{
        WorkspaceId                 = $sharedWorkspaceId
        WorkspaceName               = $sharedWorkspaceName
        WorkspaceType               = "Workspace"
        WorkspaceIsReadOnly         = $false
        WorkspaceIsOnDedicatedCapacity = $false
    }
    $workspacesInfo += $sharedWorkspaceDetails
}



# Add refresh history for each dataset in "My Workspace"
foreach ($dataset in $datasetsInfo | Where-Object { $_.WorkspaceId -eq $myWorkspaceId }) {
    $refreshHistoryUrl = "https://api.powerbi.com/v1.0/myorg/datasets/$($dataset.DatasetId)/refreshes"
    $refreshHistoryResponse = Invoke-PowerBIRestMethod -Method GET -Url $refreshHistoryUrl | ConvertFrom-Json

    foreach ($refresh in $refreshHistoryResponse.value) {
        # Rename properties based on the map and dynamically include dataset context
        $renamedRefreshRecord = Rename-Properties -object $refresh -renameMap $datasetRefreshRenameMap
        $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceId" -NotePropertyValue $myWorkspaceId -Force
        $renamedRefreshRecord | Add-Member -NotePropertyName "WorkspaceName" -NotePropertyValue $myWorkspaceName -Force
        $renamedRefreshRecord | Add-Member -NotePropertyName "DatasetId" -NotePropertyValue $dataset.DatasetId -Force
        $renamedRefreshRecord | Add-Member -NotePropertyName "DatasetName" -NotePropertyValue $dataset.DatasetName -Force

        $datasetRefreshHistory += $renamedRefreshRecord
	    }
	}
    } else {
	Write-Host "Skipping 'My Workspace' processing because Specific Workspace ID is provided."
}


# Export all collections to Excel
if (Test-Path $excelFile) {
    Remove-Item $excelFile -Force
}

$workspacesInfo | Export-Excel -Path $excelFile -WorksheetName "Workspaces" -AutoSize
$datasetsInfo | Export-Excel -Path $excelFile -WorksheetName "Datasets" -AutoSize -Append
$datasetSourcesInfo | Export-Excel -Path $excelFile -WorksheetName "DatasetSourcesInfo" -AutoSize -Append
$datasetRefreshHistory | Export-Excel -Path $excelFile -WorksheetName "DatasetRefreshHistory" -AutoSize -Append
$dataflowsInfo | Export-Excel -Path $excelFile -WorksheetName "Dataflows" -AutoSize -Append
$dataflowLineage | Export-Excel -Path $excelFile -WorksheetName "DataflowLineage" -AutoSize -Append
$dataflowSourcesInfo | Export-Excel -Path $excelFile -WorksheetName "DataflowSourcesInfo" -AutoSize -Append
$dataflowRefreshHistory | Export-Excel -Path $excelFile -WorksheetName "DataflowRefreshHistory" -AutoSize -Append
$reportsInfo | Export-Excel -Path $excelFile -WorksheetName "Reports" -AutoSize -Append
$reportPagesInfo | Export-Excel -Path $excelFile -WorksheetName "ReportPages" -AutoSize -Append
$appsInfo | Export-Excel -Path $excelFile -WorksheetName "Apps" -AutoSize -Append
$reportsInAppInfo | Export-Excel -Path $excelFile -WorksheetName "AppReports" -AutoSize -Append

Write-Host "Export completed. Data is saved to $excelFile"
 








#### Start of Model Backup ####









# Loop through datasetsInfo collection to perform model export
foreach ($dataset in $datasetsInfo) {
    # Get workspace details
    $workspace = $workspacesInfo | Where-Object { $_.WorkspaceId -eq $dataset.WorkspaceId }

    # Check if the workspace is Premium or Fabric capacity before proceeding
    if ($workspace.WorkspaceIsOnDedicatedCapacity -eq $true) {

        $workspaceName = $dataset.WorkspaceName -replace '\[', '%5B' -replace '\]', '%5D' -replace ' ', '%20'
        $datasetId = $dataset.DatasetId
        $datasetName = $dataset.DatasetName

        # Clean up workspace name
        $cleanDatasetWorkspaceName = $dataset.WorkspaceName -replace '\[', '(' -replace '\]', ')'
        $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
        $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName.TrimStart()

        # Clean up dataset name
        $cleanDatasetName = $datasetName -replace '\[', '(' -replace '\]', ')'
        $cleanDatasetName = $cleanDatasetName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
        $cleanDatasetName = $cleanDatasetName.TrimStart()

        # Construct the folder path and create it if it doesn't exist
        $modelBackupsPath = "$baseFolderPath\Model Backups"
        if (-not (Test-Path $modelBackupsPath)) {
            New-Item -ItemType Directory -Force -Path $modelBackupsPath
        }

        # Construct the date model backup folder path and create it if it doesn't exist
        $folderPath = "$modelBackupsPath\$date"
        if (-not (Test-Path $folderPath)) {
            New-Item -ItemType Directory -Force -Path $folderPath
        }

        # Define the new model database name
        $newModelDatabaseName = "$cleanDatasetWorkspaceName ~ $cleanDatasetName"

        # Create the C# script to rename the Model.Database.Name
        $csharpScript = @"
Model.Database.Name = `"$newModelDatabaseName`";
"@

        # Save the C# script to a temporary file
        $tempScriptPath = [System.IO.Path]::GetTempFileName()
        $tempScriptPath = [System.IO.Path]::ChangeExtension($tempScriptPath, ".cs")
        Set-Content -Path $tempScriptPath -Value $csharpScript

        # Construct the argument list for the model export with renaming
        $modelExportArgs = "`"Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/$workspaceName;Password=$(Get-CurrentAccessToken)`" $datasetId -S `"$tempScriptPath`" -B `"$folderPath\$cleanDatasetWorkspaceName ~ $cleanDatasetName.bim`""
        
        # Start the Tabular Editor process for model export and renaming
        Start-Process -FilePath "$TabularEditor2Path" -Wait -NoNewWindow -PassThru -ArgumentList $modelExportArgs

        # Clean up the temporary script file
        Remove-Item -Path $tempScriptPath
    }
}








#### Start of Report Backup ####








# Define the report backups path
$reportBackupsPath = Join-Path -Path $baseFolderPath -ChildPath "Report Backups"

# Check if the base folder exists, if not create it
if (-not (Test-Path -Path $baseFolderPath)) {
    New-Item -Path $baseFolderPath -ItemType Directory -Force
}

# Check if the "Report Backups" folder exists, if not create it
if (-not (Test-Path -Path $reportBackupsPath)) {
    New-Item -Path $reportBackupsPath -ItemType Directory -Force
}

# Create a new sub folder for the date
$newDateFolder = Join-Path -Path $reportBackupsPath -ChildPath $date
if (-not (Test-Path -Path $newDateFolder)) {
    New-Item -Path $newDateFolder -ItemType Directory -Force
}

# Define the temporary extraction folder
$tempExtractFolder = "$baseFolderPath\Config\Temp"

# Check if the temp extract folder exists, if not create it
if (-not (Test-Path -Path $tempExtractFolder)) {
    New-Item -Path $tempExtractFolder -ItemType Directory -Force
}

foreach ($workspace in $workspacesInfo) {
    $workspaceName = $workspace.WorkspaceName
    $workspaceId = $workspace.WorkspaceId

    # Clean up workspace name
    $cleanWorkspaceName = $workspaceName -replace '\[', '(' -replace '\]', ')'
    $cleanWorkspaceName = $cleanWorkspaceName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $cleanWorkspaceName = $cleanWorkspaceName.TrimStart()

    # Fetch reports from the existing list, NOT from Get-PowerBIReport
    $reports = $reportsInfo | Where-Object { $_.WorkspaceId -eq $workspaceId }

    # Export each report in the workspace
    foreach ($report in $reports) {

    $reportName = $report.ReportName
    $reportId = $report.ReportId

        # Clean up report name
        $cleanReportName = $reportName -replace '\[', '(' -replace '\]', ')'
        $cleanReportName = $cleanReportName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
        $cleanReportName = $cleanReportName.TrimStart()

        # Determine the file extension based on the report type
        $fileExtension = if ($report.WebUrl -like "*/rdlreports/*") { "rdl" } else { "pbix" }
        $filename = "$cleanWorkspaceName ~ $cleanReportName.$fileExtension"
        $filepath = Join-Path -Path $newDateFolder -ChildPath $filename
        $extractFolder = Join-Path -Path $tempExtractFolder -ChildPath "$cleanWorkspaceName ~ $cleanReportName"

        # Check if the file exists and remove it if it does
        if (Test-Path $filepath) {
            Remove-Item $filepath -Force
        }

        Write-Output "Exporting $cleanWorkspaceName ~ $cleanReportName"
        Export-PowerBIReport -Id $reportId -OutFile $filepath

        # Only process model extraction if the Workspace is not Premium or Fabric Capacity (i.e. only Pro Workspaces)

        if ($workspace.WorkspaceIsOnDedicatedCapacity -eq $false) {
            # Extract the pbix file
            Write-Output "Extracting $cleanWorkspaceName ~ $cleanReportName"
            pbi-tools extract $filepath -extractFolder $extractFolder -modelSerialization Raw

            # Define a different drive letter to use (e.g., Q:)
            $driveLetter = "Q:"
            $quotedExtractFolder = "`"$extractFolder`""

            # Ensure the drive letter is not already mapped
            if (Test-Path $driveLetter) {
                subst $driveLetter /D
            }

            # Log the command for debugging
            $substCommand = "subst $driveLetter $quotedExtractFolder"
            Write-Output "Running command: $substCommand"

            # Execute the subst command directly
            cmd.exe /c $substCommand

            # Confirm mapping success
            if (-not (Test-Path $driveLetter)) {
                Write-Error "Failed to map $extractFolder to $driveLetter"
                continue
            }

            # Generate the .bim file using the mapped drive
            Write-Output "Generating .bim for $cleanWorkspaceName ~ $cleanReportName using $driveLetter"
            pbi-tools generate-bim $driveLetter -transforms RemovePBIDataSourceVersion

            # Rename and move the generated .bim files
            $bimFilesGenerated = Get-ChildItem -Path $extractFolder -Filter *.bim
            foreach ($bimFile in $bimFilesGenerated) {
                $newBimName = "$cleanWorkspaceName ~ $cleanReportName.bim"
                $newBimPath = Join-Path -Path $tempExtractFolder -ChildPath $newBimName

                # Rename and move the .bim file
                Move-Item -Path $bimFile.FullName -Destination $newBimPath -Force
                Write-Output "Moved .bim file to $newBimPath"
            }

            # Cleanup: Unmap the drive
            subst $driveLetter /D
        } else {
        }
    }

    Write-Output "$cleanWorkspaceName Reports exported and processed"
}

# Move .bim files to the report backups path and clean up temp folder
$bimFiles = Get-ChildItem -Path $tempExtractFolder -Recurse -Filter *.bim
foreach ($bimFile in $bimFiles) {
    $destinationPath = Join-Path -Path $newDateFolder -ChildPath $bimFile.Name
    Move-Item -Path $bimFile.FullName -Destination $folderPath -Force
}

$tempDestinationRemovalPath = "$baseFolderPath\Config"

# Remove the temporary extraction folder
Remove-Item -Path $tempDestinationRemovalPath\Temp -Recurse -Force

# Remove the temporary localhost folder
Remove-Item -Path $tempDestinationRemovalPath\localhost -Recurse -Force

# Remove the temporary localhost folder
Remove-Item -Path $baseFolderPath\localhost -Recurse -Force








#### Start of Tabular Editor Report Detail (Visual Object Layer) Extract ####








# Start Tabular Editor process
$p = Start-Process -filePath "$TabularEditor2Path" -Wait -NoNewWindow -PassThru `
       -ArgumentList "`"$baseFolderPath\Config\Blank Model.bim`" -S `"$Script1Path`" -B `"localhost\tabular`" `"MyModel`""

# Define the output Excel file path in the parent folder
$outputExcelFile = Join-Path -Path $baseFolderPath -ChildPath "Report Detail.xlsx"

# Check if the Excel file already exists and delete it if it does
if (Test-Path -Path $outputExcelFile) {
    Remove-Item $outputExcelFile -Force
    $excelExists = $false
} else {
    $excelExists = $false
}

# Get the latest folder by date
$latestDateFolder = Get-ChildItem -Path $reportBackupsPath -Directory |
    Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |  # Filter for folders with a date pattern
    Sort-Object { [datetime]::Parse($_.Name) } -Descending | # Sort by date, descending
    Select-Object -First 1

if ($latestDateFolder) {
    $newDateFolder = $latestDateFolder.FullName
    Write-Output "Using the latest dated folder: $newDateFolder"
} else {
    Write-Error "No valid dated folders found in 'Report Backups'."
    exit
}

# Function to clean worksheet and property names for Excel
function Clean-ExcelName {
    param (
        [string]$name
    )
    # Remove invalid characters and trim
    $cleanName = $name -replace '[^A-Za-z0-9_]', '_' # Only allow letters, numbers, underscore
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        $cleanName = 'Sheet1'
    }
    return $cleanName
}

# When exporting TXT files to Excel, Clean worksheet names
foreach ($txtFile in (Get-ChildItem -Path $newDateFolder -Filter *.txt)) {
    # Get the base name of the file (without extension) for the worksheet name
    $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($txtFile.FullName)
    $worksheetName = Clean-ExcelName $worksheetName
    
    # Import the TXT file, assuming it's tab-delimited
    $txtData = Get-Content -Path $txtFile.FullName -Encoding UTF8 | ConvertFrom-Csv -Delimiter "`t"

    if ($excelExists) {
        # Append data to the existing Excel file
        $txtData | Export-Excel -Path $outputExcelFile -WorksheetName $worksheetName -AutoNameRange -Append
    } else {
        # Create a new Excel file with the data
        $txtData | Export-Excel -Path $outputExcelFile -WorksheetName $worksheetName -AutoNameRange
        $excelExists = $true
    }
}


Write-Output "TXT files appended to $outputExcelFile"


#### Cleanup Visual Object folders Remaining ####


# Check if the any remaining VOL folders exists
if (Test-Path -Path $newDateFolder) {
    # Get all subfolders
    $subfolders = Get-ChildItem -Path $newDateFolder -Directory

    foreach ($subfolder in $subfolders) {
        Remove-Item -Path $subfolder.FullName -Recurse -Force
    }

    Write-Output "Additional Subfolders in $newDateFolder deleted"
} else {
    Write-Output "Additional Subfolders in $newDateFolder do not exist"
}








#### Start of Model Detail Script Run with Script2Path ####








$latestModelDateFolder = Get-ChildItem -Path $modelBackupsPath -Directory |
    Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |  # Filter for folders with a date pattern
    Sort-Object { [datetime]::Parse($_.Name) } -Descending | # Sort by date, descending
    Select-Object -First 1

if ($latestModelDateFolder) {
    # Override $folderPath with the latest-dated folder
    $folderPath = $latestModelDateFolder.FullName
    Write-Output "Using the latest dated folder: $folderPath"
} else {
    # Throw an error if no valid dated folders are found
    Write-Error "No valid dated folders found in 'Model Backups'."
    exit
}

# Check if there are any .bim files in the $folderPath
$folderPathBimFiles = Get-ChildItem -Path $folderPath -Filter *.bim
if ($folderPathBimFiles.Count -eq 0) {
    $sourceFolderPath = $newDateFolder
} else {
    $sourceFolderPath = $folderPath
}


# Loop through datasetsInfo collection to run the first model script
foreach ($dataset in $datasetsInfo) {
    $workspaceName = $dataset.WorkspaceName
    $datasetName = $dataset.DatasetName

    # Clean up workspace name
    $cleanDatasetWorkspaceName = $workspaceName -replace '\[', '(' -replace '\]', ')'
    $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName.TrimStart()

    # Clean up dataset name
    $cleanDatasetName = $datasetName -replace '\[', '(' -replace '\]', ')'
    $cleanDatasetName = $cleanDatasetName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $cleanDatasetName = $cleanDatasetName.TrimStart()

    # Construct the argument list for the first model script run
    $modelScriptArgs = "`"$sourceFolderPath\$cleanDatasetWorkspaceName ~ $cleanDatasetName.bim`" -S `"$Script2Path`" -B `"localhost\tabular`" `"MyModel`""
    
    # Start the Tabular Editor process to run the first model script
    Start-Process -FilePath "$TabularEditor2Path" -Wait -NoNewWindow -PassThru -ArgumentList $modelScriptArgs
}

Write-Host "First model export and script run completed."

#### Start of PowerShell Combining to Semantic Models Worksheet ####

# Define the output Excel file path in the base folder
$outputExcelFile = Join-Path -Path $baseFolderPath -ChildPath "Model Detail.xlsx"

# Check if the Excel file already exists and delete it if it does
if (Test-Path -Path $outputExcelFile) {
    Remove-Item $outputExcelFile -Force
}

# Initialize an empty array to store all CSV data for Semantic Models
$semanticModelsCsvData = @()

foreach ($csvFile in (Get-ChildItem -Path $folderPath -Filter *.csv)) {
    # Exclude files that end with "_MD.csv" for Semantic Models
    if ($csvFile.Name -notlike "*_MD.csv") {
        # Import the CSV file
        $csvData = Import-Csv -Path $csvFile.FullName

        # Append the data to the array
        $semanticModelsCsvData += $csvData
    }
}

# Create a new Excel file with the data in the "Semantic Models" worksheet
$semanticModelsCsvData | Export-Excel -Path $outputExcelFile -WorksheetName "Semantic Models" -AutoNameRange

Write-Output "CSV files combined into the 'Semantic Models' worksheet in $outputExcelFile"








#### Start of Model Detail Script Run with Script3Path ####








# Loop through datasetsInfo collection to run the second model script
foreach ($dataset in $datasetsInfo) {
    $workspaceName = $dataset.WorkspaceName
    $datasetName = $dataset.DatasetName

    # Clean up workspace name
    $cleanDatasetWorkspaceName = $workspaceName -replace '\[', '(' -replace '\]', ')'
    $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $cleanDatasetWorkspaceName = $cleanDatasetWorkspaceName.TrimStart()

    # Clean up dataset name
    $cleanDatasetName = $datasetName -replace '\[', '(' -replace '\]', ')'
    $cleanDatasetName = $cleanDatasetName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $cleanDatasetName = $cleanDatasetName.TrimStart()

    # Construct the argument list for the second model script run
    $modelScriptArgs = "`"$sourceFolderPath\$cleanDatasetWorkspaceName ~ $cleanDatasetName.bim`" -S `"$Script3Path`" -B `"localhost\tabular`" `"MyModel`""
    
    # Start the Tabular Editor process to run the second model script
    Start-Process -FilePath "$TabularEditor2Path" -Wait -NoNewWindow -PassThru -ArgumentList $modelScriptArgs
}

Write-Host "Second model export and script run completed."

#### Start of PowerShell Combining to Measure Dependencies Worksheet ####

# Initialize an empty array to store all CSV data for Measure Dependencies
$measureDependenciesCsvData = @()

foreach ($csvFile in (Get-ChildItem -Path $folderPath -Filter *.csv)) {
    # Include only files that end with "_MD.csv" for Measure Dependencies
    if ($csvFile.Name -like "*_MD.csv") {
        # Import the CSV file
        $csvData = Import-Csv -Path $csvFile.FullName

        # Append the data to the array
        $measureDependenciesCsvData += $csvData
    }
}

# Append data to the existing Excel file in the "Measure Dependencies" worksheet
$measureDependenciesCsvData | Export-Excel -Path $outputExcelFile -WorksheetName "Measure Dependencies" -AutoNameRange -Append

Write-Output "CSV files combined into the 'Measure Dependencies' worksheet in $outputExcelFile"








#### Start of Power BI Dataflow Backup and Detail Extract ####








# Define the dataflow backups path
$dataflowBackupsPath = Join-Path -Path $baseFolderPath -ChildPath "Dataflow Backups"

# Check if the "Dataflow Backups" folder exists, if not create it
if (-not (Test-Path -Path $dataflowBackupsPath)) {
    New-Item -Path $dataflowBackupsPath -ItemType Directory
}

# Create a variable for end of week (Friday) date
$date = (Get-Date -UFormat "%Y-%m-%d")

# Create a new folder for the backups
$dataflow_new_date_folder = Join-Path -Path $dataflowBackupsPath -ChildPath $date
New-Item -Path $dataflow_new_date_folder -ItemType Directory -Force

# Set the base output file path
$baseOutputFilePath = $dataflow_new_date_folder

# Get the latest folder by date
$latestDataflowDateFolder = Get-ChildItem -Path $dataflowBackupsPath -Directory |
    Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |  # Filter for folders with a date pattern
    Sort-Object { [datetime]::Parse($_.Name) } -Descending | # Sort by date, descending
    Select-Object -First 1

if ($latestDataflowDateFolder) {
    # Override $baseOutputFilePath with the latest-dated folder
    $folderPath = $latestDataflowDateFolder.FullName
    Write-Output "Using the latest dated folder: $baseOutputFilePath"
} else {
    # Throw an error if no valid dated folders are found
    Write-Error "No valid dated folders found in 'Dataflow Backups'."
    exit
}

# Set the combined Excel output path
$combinedExcelOutputPath = Join-Path -Path $dataflow_new_date_folder -ChildPath "Dataflow Detail.xlsx"

# Define the headers
$headers = @("Dataflow ID", "Dataflow Name", "Query Name", "Query", "Report Date", "Workspace Name - Dataflow Name")

# Initialize a combined DataTable with the specified headers
$combinedDataTable = New-Object System.Data.DataTable
foreach ($header in $headers) {
    $combinedDataTable.Columns.Add($header, [System.String])
}

# Get the current date
$currentDate = [datetime]::Parse($latestDataflowDateFolder.Name)

# Function to check if a position is within curly braces
function IsInsideCurlyBraces {
    param (
        [string]$text,
        [int]$position
    )
    $openBraces = 0
    for ($i = 0; $i -lt $position; $i++) {
        if ($text[$i] -eq '{') { $openBraces++ }
        elseif ($text[$i] -eq '}') { $openBraces-- }
    }
    return $openBraces -gt 0
}

# Loop through all workspaces to fetch dataflows


foreach ($workspace in $workspacesInfo) {
    $workspaceName = $workspace.WorkspaceName
    $workspaceId = $workspace.WorkspaceId

    # Set the Power BI REST API URL for the dataflow details
    $dataflowDetailsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/dataflows"

    # Get the list of dataflows in the workspace
    $dataflowsResponse = Invoke-PowerBIRestMethod -Url $dataflowDetailsUrl -Method Get

    # Parse the JSON response
    $dataflows = $dataflowsResponse | ConvertFrom-Json

    # Check if the response is valid and contains dataflows
    if ($dataflows -and $dataflows.value) {
        Write-Host "Dataflows found in workspace '$workspaceName': $($dataflows.value.Count)"
        
        # Iterate through the dataflows
        foreach ($dataflow in $dataflows.value) {
            $dataflowId = $dataflow.objectId
            $dataflowName = $dataflow.name

            # Clean up workspace name
            $cleanWorkspaceName = $workspaceName -replace '\[', '(' -replace '\]', ')'
            $cleanWorkspaceName = $cleanWorkspaceName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
            $cleanWorkspaceName = $cleanWorkspaceName.TrimStart()

            # Clean up dataflow name
            $cleanDataFlowName = $dataflowName -replace '\[', '(' -replace '\]', ')'
            $cleanDataFlowName = $cleanDataFlowName -replace "[^a-zA-Z0-9\(\)&,.-]", " "
            $cleanDataFlowName = $cleanDataFlowName.TrimStart()
            
            # Define output file path specific to the dataflow
            $dataflowOutputFilePath = Join-Path -Path $baseOutputFilePath -ChildPath "$cleanWorkspaceName ~ $cleanDataFlowName.txt"
            
            # Set the Power BI REST API URL for the specific dataflow
            $apiUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/dataflows/$dataflowId"
            
            # Get the dataflow
            $response = Invoke-PowerBIRestMethod -Url $apiUrl -Method Get
            
            # Convert the response to JSON string
            $jsonString = $response | ConvertTo-Json
            
            # Write the JSON string to a text file
            $jsonString | Out-File -FilePath $dataflowOutputFilePath -Encoding UTF8
            
            # Extract the data from the JSON response without writing intermediate files
            $startMarker = '"document\":'
            $endMarker = '\\r\\n\"'
            $startIndex = $jsonString.IndexOf($startMarker) + $startMarker.Length
            $endIndex = $jsonString.IndexOf($endMarker, $startIndex)
	    $fallbackEndMarker = '"connectionOverrides\"'

		# Extract start and end indices
		$startIndex = $jsonString.IndexOf($startMarker) + $startMarker.Length
		$endIndex = $jsonString.IndexOf($endMarker, $startIndex)

		# Check if $endMarker is valid; if not, look for fallback marker
		if ($endIndex -lt 0) {
		    $endIndex = $jsonString.IndexOf($fallbackEndMarker, $startIndex)
		}
            
            # Check if the startIndex and endIndex are valid
            if ($startIndex -ge 0 -and $endIndex -ge 0 -and $endIndex -gt $startIndex) {
                $documentContent = $jsonString.Substring($startIndex, $endIndex - $startIndex)
            } else {
                Write-Host "Invalid Start/End mark for dataflow '$dataflowName' in workspace '$workspaceName'."
                continue
            }

            # Format the extracted content
            $formattedText = $documentContent -replace '\\r\\n', "`n" `
                                                    -replace '\\\"', '"' `
                                                    -replace '\\\\', '\' `
                                                    -replace '(?<=\w)(=|then|else)(?=\w)', ' $1 ' `
                                                    -replace '(?<=then)\s+', ' ' `
                                                    -replace '\s+(?=else)', "`n    "

            # Additional formatting
            $insideQuotes = $false
            $formattedTextStep7 = ""
            for ($i = 0; $i -lt $formattedText.Length; $i++) {
                $char = $formattedText[$i]
                if ($char -eq '"') {
                    $insideQuotes = -not $insideQuotes
                }
                if ($char -eq ',' -and -not $insideQuotes -and -not (IsInsideCurlyBraces -text $formattedText -position $i)) {
                    $formattedTextStep7 += "$char`n    "
                } else {
                    $formattedTextStep7 += $char
                }
            }

            $formattedTextStep8 = $formattedTextStep7 -replace 'nshared', 'QueryStartandEndMarker' `
                                                      -replace '\\r', ' ' `
                                                      -replace '\\n', ' ' `
                                                      -replace '\\', '' `
                                                      -replace '\r\n', "`n" `
                                                      -replace '(?<!["\w])r(\s)(?![\w"])', '$1' `
                                                      -replace '\;r', ''

            $formattedTextStep9 = $formattedTextStep8 -replace '(?<=let|in|each)\s+', "`n    " `
                                                      -replace '(?<=,\s*)#', "`n    #"

            $formattedTextStep9 = $formattedTextStep9 -replace '(?<!\n)QueryStartandEndMarker\s', "`nQueryStartandEndMarker`n`n"

            $formattedTextStep9 = $formattedTextStep9 -replace '(^|\s)let', "`nlet"

            $formattedTextStep9 = $formattedTextStep9 -replace '\)  in', ")`n    in"

            # Read the content directly from the formatted text
            $fileContent = $formattedTextStep9

            # Initialize variables
            $inQuery = $false
            $queryName = ""
            $query = ""
            $data = @()

            # Split the content into lines
            $lines = $fileContent -split "`n"

            # Iterate over the lines
            foreach ($line in $lines) {
                if ($line -match 'QueryStartandEndMarker\s*') {
                    # If we find a new query, save the previous one
                    if ($queryName -ne "") {
                        $data += [PSCustomObject]@{
                            "Dataflow ID" = $dataflowId
                            "Dataflow Name" = $dataflowName
                            "Query Name" = $queryName
                            "Query" = $query.Trim()
                            "Report Date" = $currentDate
                            "Workspace Name - Dataflow Name" = "$cleanWorkspaceName ~ $cleanDataFlowName"  # Add data to new column
                        }
                    }
                    # Reset variables for new query
                    $inQuery = $true
                    $queryName = ""
                    $query = ""
                } elseif ($inQuery -and $line.Trim() -ne "") {
                    # Set the query name as the first non-empty line after QueryStartandEndMarker
                    if ($queryName -eq "") {
                        $queryName = $line.Trim()
                    } else {
                        # Append the line to the query
                        $query += "$line`n"
                    }
                } elseif ($queryName -ne "") {
                    # Append the line to the query
                    $query += "$line`n"
                }
            }

            # Add the last query
            if ($queryName -ne "") {
                $data += [PSCustomObject]@{
                    "Dataflow ID" = $dataflowId
                    "Dataflow Name" = $dataflowName
                    "Query Name" = $queryName
                    "Query" = $query.Trim()
                    "Report Date" = $currentDate
                    "Workspace Name - Dataflow Name" = "$cleanWorkspaceName ~ $cleanDataFlowName"  # Add data to new column
                }
            }

            # Fill the combined DataTable with data
            foreach ($item in $data) {
                $row = $combinedDataTable.NewRow()
                $row["Dataflow ID"] = $item."Dataflow ID"
                $row["Dataflow Name"] = $item."Dataflow Name"
                $row["Query Name"] = $item."Query Name"
                $row["Query"] = $item.Query
                $row["Report Date"] = $item."Report Date"
                $row["Workspace Name - Dataflow Name"] = $item."Workspace Name - Dataflow Name"  # Add data to new column
                $combinedDataTable.Rows.Add($row)
            }
        }
    } else {
        Write-Host "No dataflows found in workspace '$workspaceName'."
    }
}

# Check if the combined DataTable has any rows, if not add a dummy row with headers only
if ($combinedDataTable.Rows.Count -eq 0) {
    $row = $combinedDataTable.NewRow()
    foreach ($header in $headers) {
        $row[$header] = ""
    }
    $combinedDataTable.Rows.Add($row)
}



# Export the combined DataTable to an Excel file
$combinedDataTable | Export-Excel -Path $combinedExcelOutputPath -AutoSize
Write-Output "Data exported to $combinedExcelOutputPath"

# Combine files if both are there
$fileName = "Dataflow Detail.xlsx"
$sourceFilePath = Join-Path -Path $dataflow_new_date_folder -ChildPath $fileName
$destinationFilePath = Join-Path -Path $baseFolderPath -ChildPath $fileName

# Check if the source file exists
if (-not (Test-Path -Path $sourceFilePath)) {
    Write-Error "Source file not found: $sourceFilePath"
    exit
}

# Remove the destination file if it already exists
if (Test-Path -Path $destinationFilePath) {
    Remove-Item -Path $destinationFilePath -Force
}

# Copy the source file to the destination
Copy-Item -Path $sourceFilePath -Destination $destinationFilePath

# Load the source and destination Excel files
$sourceData = Import-Excel -Path $sourceFilePath
$destinationData = Import-Excel -Path $destinationFilePath

# Combine the data
$combinedData = $destinationData + $sourceData

# Export the combined data to the destination file
$combinedData | Export-Excel -Path $destinationFilePath -WorksheetName "Sheet1"

# Stop the background job after script completion
Stop-Job -Name "TokenRefreshJob"
Remove-Job -Name "TokenRefreshJob"

Write-Output "Excel files processed and combined successfully."

