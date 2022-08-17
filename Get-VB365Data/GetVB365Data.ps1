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
    [Parameter(Mandatory=$False)]
    [string]$Days=30
)

$options = @(7,30,90,180)

if ($Days -notin $options) {
    Write-Host "Please choose either 7, 30, 90, or 180"
    return
}

Connect-MgGraph -Scopes "user.read.all", "reports.read.all"

# GetOffice365ActiveUserDetail
Invoke-MgGraphRequest -Uri  "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D$Days')" -OutputFilePath active_user_detail.csv

# GetOffice365ActiveUserCounts
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserCounts(period='D$Days')" -OutputFilePath active_user_counts.csv

# GetOffice365GroupsActivityGroupCounts
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityGroupCounts(period='D$Days')" -OutputFilePath group_activity_counts.csv

# GetMailboxUsageDetail
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D$Days')" -OutputFilePath mailbox_usage_detail.csv

# GetMailboxUsageStorage
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageStorage(period='D$Days')" -OutputFilePath mailbox_useage_storage.csv

# GetOneDriveUsageStorage
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageStorage(period='D$Days')" -OutputFilePath onedrive_usage_storage.csv

# GetSharePointSiteUsageStorage
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageStorage(period='D$Days')" -OutputFilePath sharepoint_site_storage.csv

# GetSharePointSiteUseageSiteCountds
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageSiteCounts(period='D$Days')" -OutputFilePath sharepoint_site_counts.csv

# GetSharePointSitesDetail
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D$Days')" -OutputFilePath sharepoint_sites_detail.csv

function Get-Change {
    param (
        [decimal]$Newest,
        [decimal]$Oldest,
        [int]$Days
    )
    [float]$diff = $newest - $oldest
    [float]$change = ((($diff / $oldest) * 100) / $Days) * 7
    if($change -lt 0.01) {
        $change = 0.01
    }
    $change
}

function Test-Size {
    param (
        [int]$Cap  
    )

    if($Cap -le 0) {
        return 1
    } else {
        return $Cap
    }
}

class RetentionSize {
    [int]$value
    [string]$unit

    RetentionSize(
        [int]$v,
        [string]$u
    ){
        $this.value = $v
        $this.unit = $u
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
    ){
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
    ){
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

$userCounts = Import-Csv -Path ./active_user_counts.csv 
$userDetail = Import-Csv -Path ./active_user_detail.csv
$sharePointSites = Import-Csv -Path ./sharepoint_site_counts.csv

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
$exchangeAll = ($userDetail | Where-Object {$_.'Has Exchange License' -eq 'True'} ).count
$onedriveAll = ($userDetail | Where-Object {$_.'Has OneDrive License' -eq 'True'} ).count
$teamsAll = ($userDetail | Where-Object {$_.'Has Teams License' -eq 'True'} ).count

$sharepointSites = ($sharePointSites | Where-Object {$_.'Site Type' -like 'ALL'} | ForEach-Object {[int]$_.Total }  | Measure-Object -Maximum).Maximum

# get capacity info for each
$mailboxStorage = Import-Csv -Path ./mailbox_useage_storage.csv
$sharepointStorage = Import-Csv -Path ./sharepoint_site_storage.csv
$onedriveStorage = Import-Csv -Path ./onedrive_usage_storage.csv

# Filter out any zero values, removes possible error if report length does not meet the period
$mailboxStorage = $mailboxStorage | Where-Object {$_.'Storage Used (Byte)' -gt 0}
$sharepointStorage = $sharepointStorage | Where-Object {$_.'Storage Used (Byte)' -gt 0}
$onedriveStorage = $onedriveStorage | Where-Object {$_.'Storage Used (Byte)' -gt 0}

# exchange
[decimal]$oldestMailbox = $mailboxStorage | Select-Object -Last 1 | ForEach-Object {$_.'Storage Used (Byte)'}
[decimal]$newestMailbox = $mailboxStorage | Select-Object -First 1 | ForEach-Object {$_.'Storage Used (Byte)'}
$mailBoxChange = Get-Change -Oldest $oldestMailbox -Newest $newestMailbox -Days $Days
$newestMailboxTb = $newestMailbox / [Math]::Pow(1024, 4)
$newestMailboxTb = Test-Size $newestMailboxTb

#sharepoint
[decimal]$oldestSharepoint = $sharepointStorage | Select-Object -Last 1 | ForEach-Object {$_.'Storage Used (Byte)'}
[decimal]$newestSharepoint = $sharepointStorage | Select-Object -First 1 | ForEach-Object {$_.'Storage Used (Byte)'}
$sharepointChange = Get-Change -Oldest $oldestSharepoint -Newest $newestSharepoint -Days $Days
$newestSharepointTb = $oldestSharepoint / [Math]::Pow(1024, 4)
$newestSharepointTb = Test-Size $newestSharepointTb

#onedrive
[decimal]$oldestOnedrive = $onedriveStorage | Select-Object -Last 1 | ForEach-Object {$_.'Storage Used (Byte)'}
[decimal]$newestOnedrive = $onedriveStorage | Select-Object -First 1 | ForEach-Object {$_.'Storage Used (Byte)'}
$onedriveChange = Get-Change -Oldest $oldestOnedrive -Newest $newestOnedrive -Days $Days
$newestOnedriveTb = $newestOnedrive / [Math]::Pow(1024, 4)
$newestOnedriveTb = Test-Size $newestOnedriveTb

# create the retention size object
$retentionSize = New-Object -TypeName RetentionSize -ArgumentList 5,"Years"

# create the main object
$primaryMailbox = [TableItem]::new(
    "primaryMailbox",
    [math]::Round($newestMailboxTb,2),
    [math]::Round($mailBoxChange,2),
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
    [math]::Round($newestOnedriveTb,2),
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