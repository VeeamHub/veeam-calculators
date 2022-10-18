<# 
    .SYNOPSIS
        Script uses the Microsoft Graph SDK to gather relevant data and outputs a json file that can be imported into the Veeam M365 calculator.
    .DESCRIPTION
        Script gathers connects to Microsoft Graph via the PowerShell SDK and downloads relevant reports to csv format.
        The scope of the report is 30 days for all Graph Reports by default but can be modified using the -Days parameter 
        It then processes these files into json that can be uploaded via the "import" function on the Veeam Calculator.
        Requires the PowerShell Graph SDK
        Uses "user.read.all", "reports.read.all" permissions
    .PARAMETER Days
        Changes the quantity of days the analysis is based on.
        Choices are 7, 30, 90, and 180.
    .PARAMETER Local
        To be used if the csv files have been downloaded manually and placed in the same folder as the script.
    .EXAMPLE
        ./GetVB365Data.ps1
    .EXAMPLE
        ./GetVB365Data.ps1 -Days 90 
    .LINK
        https://calculator.veeam.com/vbo/manual
    .LINK
        https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0
    .OUTPUTS
        vb365-environment-info.json
    .Notes 
    Version:        1.0
    Author:         Ed Howard (edward.x.howard@veeam.com)
    Creation Date:  06.06.2022
    Purpose/Change: 06.06.2022 - 1.0 - Initial script development
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $False)]
    [string]$Days = 30,
    [Parameter(Mandatory = $False)]
    [bool]$Local = $False
)

$options = @(7, 30, 90, 180)


if ($Days -notin $options) {
    Write-Host "Please choose either 7, 30, 90, or 180"
    return
}

if ($Local -eq $False) {

    # Add check for modules loaded
    if (!(Get-Module -Name Microsoft.Graph*)) { # if no modules found with this name.
        Write-Host "Microsoft.Graph module does not appear to be loaded."
        Write-Host "Refer to README @ https://github.com/VeeamHub/veeam-calculators/tree/master/Get-VB365Data for guidance on installing/loading these modules."
        Exit
    }

    try {
        Connect-MgGraph -Scopes "user.read.all", "reports.read.all"

        # GetOffice365ActiveUserDetail
        Get-MgReportOffice365ActiveUserDetail -Period D$Days -OutFile active_user_detail.csv

        # GetOffice365ActiveUserCounts
        Get-MgReportOffice365ActiveUserCount -Period D$Days -OutFile active_user_counts.csv

        # GetOffice365GroupsActivityGroupCounts
        Get-MgReportOffice365GroupActivityCount -Period D$Days -OutFile group_activity_counts.csv

        # GetMailboxUsageDetail
        Get-MgReportMailboxUsageDetail -Period D$Days -OutFile mailbox_usage_detail.csv 
        
        # GetMailboxUsageStorage
        Get-MgReportMailboxUsageStorage -Period D$Days -OutFile mailbox_usage_storage.csv

        # GetOneDriveUsageStorage
        Get-MgReportOneDriveUsageStorage -Period D$Days -OutFile onedrive_usage_storage.csv

        # GetSharePointSiteUsageStorage
        Get-MgReportSharePointSiteUsageStorage -Period D$Days -OutFile sharepoint_site_storage.csv

        # GetSharePointSiteUseageSiteCounts
        Get-MgReportSharePointSiteUsageSiteCount -Period D$Days -OutFile sharepoint_site_counts.csv

        # GetSharePointSitesDetail
        Get-MgReportSharePointSiteUsageDetail -Period D$Days -OutFile sharepoint_sites_detail.csv
    } catch {
        Write-Host "There was an issue executing Microsoft.Graph cmdlets."
        Write-Host "Check to ensure the modules are loaded, you are authenticated, and you have the required Graph permissions."
        Write-Host "Refer to README @ https://github.com/VeeamHub/veeam-calculators/tree/master/Get-VB365Data for details."
        Exit
    }
}
else {
    Write-Host "Running in Local mode"
}

function Get-Change {
    param (
        [decimal]$Newest,
        [decimal]$Oldest,
        [int]$Days
    )
    [float]$diff = $newest - $oldest
    [float]$change = ((($diff / $oldest) * 100) / $Days) * 7
    if ($change -lt 0.01) {
        $change = 0.01
    }
    $change
}

function Test-Size {
    param (
        [float]$Cap  
    )

    if ($Cap -le 0.01) {
        return 0.01
    }
    else {
        return $Cap
    }
}

class RetentionSize {
    [int]$value
    [string]$unit

    RetentionSize(
        [int]$v,
        [string]$u
    ) {
        $this.value = $v
        $this.unit = $u
    }
}

function Get-Value {
    param(
        [array]$data,
        [bool]$old
    )

    if ($old) {
        [decimal]$oldestStorage = ($data | Select-Object -Last 1).'Storage Used (Byte)'
        return $oldestStorage
    }
    else {
        [decimal]$newestStorage = ($data | Select-Object -First 1).'Storage Used (Byte)'
        return $newestStorage
    }
}

class TableItem {
    [string]$name
    [float]$totalSize
    [float]$weeklyChangeRate
    [int]$quantity
    [RetentionSize]$retentionSize

    TableItem(
        [string]$n,
        [float]$t,
        [float]$w,
        [int]$q,
        [RetentionSize]$r
    ) {
        $this.name = $n
        $this.totalSize = $t
        $this.weeklyChangeRate = $w
        $this.quantity = $q
        $this.retentionSize = $r
    }
}

class TableState {
    [TableItem]$primaryMailbox
    [TableItem]$archiveMailbox
    [TableItem]$sharepoint
    [TableItem]$onedrive

    TableState(
        [TableItem]$pm,
        [TableItem]$am,
        [TableItem]$sp,
        [TableItem]$od
    ) {
        $this.primaryMailbox = $pm
        $this.archiveMailbox = $am
        $this.sharepoint = $sp
        $this.onedrive = $od
    }
}

class CostState {
    [string]$regionName
    [string]$performanceTier
    [string]$capacityTier
    [string]$replication
    [int]$backupServerConfig
    [int]$proxyServerConfig
    [string]$offer

    CostState(
        [string]$rg,
        [string]$pt,
        [string]$ct,
        [string]$re,
        [int]$bs,
        [int]$ps,
        [string]$of
    ) {
        $this.regionName = $rg
        $this.performanceTier = $pt
        $this.capacityTier = $ct
        $this.replication = $re
        $this.backupServerConfig = $bs
        $this.proxyServerConfig = $ps
        $this.offer = $of
    }
}

class SaveData {
    [TableState]$tableState
    [int]$teamsState
    [bool]$recommended
    [CostState]$costingState

    SaveData(
        [TableState]$ts,
        [int]$tes,
        [bool]$re,
        [CostState]$cs
    ) {
        $this.tableState = $ts
        $this.teamsState = $tes
        $this.recommended = $re
        $this.costingState = $cs
    }
}


# VB365 active user counts

Try {
    $userCounts = Import-Csv -Path ./active_user_counts.csv 
    $userDetail = Import-Csv -Path ./active_user_detail.csv
    $sharePointSites = Import-Csv -Path ./sharepoint_site_counts.csv
} 
Catch {
    Write-Host "One or more user counts files CSV files are missing."
    Write-Host "active_user_counts, active_user_detail and sharepoint_site_counts files required"
    Exit
}

# removes user information from the active_user_detail.csv file
$userDetail | Select-Object 'Has Exchange License', 'Has OneDrive License', 'Has SharePoint License', 'Has Teams License' | Export-Csv -Path .\active_user_detail.csv -NoTypeInformation

$exchange = @()
$sharepoint = @()
$onedrive = @()
$teams = @()

foreach ($item in $userCounts) {
    $exchange += [int]$item.Exchange
    $sharepoint += [int]$item.SharePoint
    $onedrive += [int]$item.OneDrive
    $teams += [int]$item.Teams
}

# licensed users
$exchangeAll = ($userDetail | Where-Object { $_.'Has Exchange License' -eq 'True' } ).count
$onedriveAll = ($userDetail | Where-Object { $_.'Has OneDrive License' -eq 'True' } ).count
$teamsAll = ($userDetail | Where-Object { $_.'Has Teams License' -eq 'True' } ).count

$sharepointSites = ($sharePointSites | Where-Object { $_.'Site Type' -like 'All' } | Select-Object -First 1).Total 

try {
    # get capacity info for each
    $mailboxStorage = Import-Csv -Path ./mailbox_usage_storage.csv
    $sharepointStorage = Import-Csv -Path ./sharepoint_site_storage.csv
    $onedriveStorage = Import-Csv -Path ./onedrive_usage_storage.csv
}
catch {
    Write-Host "One or more storage CSV files are missing"
    Write-Host "mailbox_usage_storage, sharepoint_site_storage and onedrive_usage_storage files required"
    Exit
}


# Filter out any zero values, removes possible error if report length does not meet the period
$mailboxStorage = $mailboxStorage | Where-Object { $_.'Storage Used (Byte)' -gt 0 }
$sharepointStorage = $sharepointStorage | Where-Object { $_.'Storage Used (Byte)' -gt 0 }
$onedriveStorage = $onedriveStorage | Where-Object { $_.'Storage Used (Byte)' -gt 0 }

# exchange
[decimal]$oldestMailbox = Get-Value -data $mailboxStorage -old $true
[decimal]$newestMailbox = Get-Value -data $mailboxStorage -old $false
$mailBoxChange = Get-Change -Oldest $oldestMailbox -Newest $newestMailbox -Days $Days
$newestMailboxTb = $newestMailbox / [Math]::Pow(1024, 4)
$newestMailboxTb = Test-Size $newestMailboxTb

#sharepoint
[decimal]$oldestSharepoint = Get-Value -data $sharepointStorage -old $true
[decimal]$newestSharepoint = Get-Value -data $sharepointStorage -old $false
$sharepointChange = Get-Change -Oldest $oldestSharepoint -Newest $newestSharepoint -Days $Days
$newestSharepointTb = $newestSharepoint / [Math]::Pow(1024, 4)
$newestSharepointTb = Test-Size $newestSharepointTb

#onedrive
[decimal]$oldestOnedrive = Get-Value -data $onedriveStorage -old $true
[decimal]$newestOnedrive = Get-Value -data $onedriveStorage -old $false
$onedriveChange = Get-Change -Oldest $oldestOnedrive -Newest $newestOnedrive -Days $Days
$newestOnedriveTb = $newestOnedrive / [Math]::Pow(1024, 4)
$newestOnedriveTb = Test-Size $newestOnedriveTb

# create the retention size object
$retentionSize = [RetentionSize]::new(
    5, 
    "Years"
)

# create the main object
$primaryMailbox = [TableItem]::new(
    "primaryMailbox",
    [math]::Round($newestMailboxTb, 2),
    [math]::Round($mailBoxChange, 2),
    $exchangeAll,
    $retentionSize
)

# Archive Mailbox 
$archiveMailbox = [TableItem]::new(
    "archiveMailbox",
    0, 
    0,
    0,
    $retentionSize
)

# sharepoint
$sharepoint = [TableItem]::new(
    "sharepoint",
    [math]::Round($newestSharepointTb, 2),
    [math]::Round($sharepointChange, 2),
    $sharepointSites,
    $retentionSize
)

$onedrive = [TableItem]::new(
    "onedrive",
    [math]::Round($newestOnedriveTb, 2),
    [math]::Round($onedriveChange, 2),
    $onedriveAll,
    $retentionSize
)

$tableState = [TableState]::new(
    $primaryMailbox,
    $archiveMailbox,
    $sharepoint,
    $onedrive
)

# legacy
$costState = [CostState]::new(
    "US Central",
    "Tiered Block Blob",
    "Hot",
    "LRS Data Stored",
    0,
    0,
    "0003P"
)

$saveData = [SaveData]::new(
    $tableState,
    $teamsAll,
    $true,
    $costState
)

$saveData | ConvertTo-Json -Depth 9 | Out-File -FilePath vb365-environment-info.json

Write-Host "Complete, please upload to vb365-environment-info.json file to the Veeam VB365 Calculator"