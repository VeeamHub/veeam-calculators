<# 
    .SYNOPSIS
        Script counts the number of archive mailboxes and their size. Values can be used in the Veeam Backup for M365 calculator.
    .DESCRIPTION
        The script uses the ExchangeOnlineManagement PowerShell module to collect the necessary information on archive mailboxes from Exchange Online.            
        The script will ask for a user principal to connect to Exchange Online, authentication will be handled for MFA and non-MFA accounts via OAuth.
    .EXAMPLE
        ./getArchiveMailboxStats.ps1    
    .LINK
        https://calculator.veeam.com/vbo/manual
    .LINK
        https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
    .LINK
        https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.5    
    .LICENSE
        MIT
    .NOTES 
    Version:        1.1
    Author:         Ed Howard (edward.x.howard@veeam.com), Stefan Zimmermann (stefan.zimmermann@veeam.com)
    Creation Date:  08.06.2022
    Purpose/Change: 06.06.2022 - 1.0 - Initial script development
                    08.06.2022 - 1.1 - Output and script enhancements
#>

Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5
Import-Module ExchangeOnlineManagement

$userPrincipal = Read-Host "Enter User Principal Name to connect to Exchange Online"

Connect-ExchangeOnline -UserPrincipalName $userPrincipal -ShowBanner:$false

Write-Host "Gathering Stats, Please Wait.." 
$mailboxes = Get-EXOMailbox -Archive -resultsize unlimited | Select-Object name,@{n="size";e={(Get-EXOMailboxStatistics -archive $_.identity).TotalItemSize.value.toBytes()}}

$sizes = $mailboxes | Measure-Object -Property size -Sum -Maximum -Minimum -Average

Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Total number of archive mailboxes: $($sizes.Count)"
Write-Output "Total size of archive mailboxes (rounded to full GB): $([Math]::Ceiling(($sizes.Sum / 1GB)))"