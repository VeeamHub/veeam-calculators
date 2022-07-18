# Gather VB365 calculator data - beta

***NOTE: The reason project is in beta is because further testing is needed on the results of the json file. It is recommended to double check the results manually using the downloaded CSV files.
If you find any discrepancies, please raise an issue.***

This repository provides information on how to gather VB365 data for the use with the Veeam Backup for Microsoft 365 calculator. The script is provided under MIT, please review license before proceeding.

https://calculator.veeam.com/vb365

The bulk of the information can be gathered using the Graph API with the exception of Exchange Archive Mailboxes.

Exchange Archive Mailboxes uses the Exchange Online Management PowerShell module. 

The GetVB365Data.ps1 script outputs a json file that can be imported via the "Import Settings" button.

If you are not comfortable running the script, please use it as a guide to running the commands manually.

## Installation of PowerShell Graph API SDK

You will need to install the Graph Module prior to running the script.

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

    Install-Module Microsoft.Graph -Scope CurrentUser

Microsoft Documentation

https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0 

Installation of PowerShell SDK

https://docs.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0

Overview of the PowerShell SDK

https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0

## Installation of the Exchange Online Management Module

You will need to install the Exchange Online Management Module prior to running the script.

    Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5

Microsoft Documentation

https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps 


## Graph API Permissions

Required Graph permissions

    "user.read.all", "reports.read.all"

You will need to used Admin credentials when granting access to PowerShell for the above scopes. See this blog for further details.

https://practical365.com/connect-microsoft-graph-powershell-sdk/ 

The script uses the following command to Log into Graph:

    Connect-MgGraph -Scopes "user.read.all", "reports.read.all"

https://docs.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0#sign-in


## What is gathered from the Graph API

The following information is gathered from the Graph API.

| Item | Link |
| ---- | ---- |
| Mailbox Usage Detail | https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.reports/get-mgreportmailboxusagedetail?view=graph-powershell-1.0 |
| Active User Detail | https://docs.microsoft.com/en-us/graph/api/reportroot-getoffice365activeuserdetail?view=graph-rest-1.0 | 
| Active User Counts | https://docs.microsoft.com/en-us/graph/api/reportroot-getoffice365activeusercounts?view=graph-rest-1.0 |
| Mailbox Usage Storage | https://docs.microsoft.com/en-us/graph/api/reportroot-getmailboxusagestorage?view=graph-rest-1.0  |
| OneDrive Usage Storage | https://docs.microsoft.com/en-us/graph/api/reportroot-getonedriveusagestorage?view=graph-rest-1.0 | 
| SharePoint Site Usage Storage | https://docs.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagestorage?view=graph-rest-1.0 |
| SharePoint Site Detail | https://docs.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail?view=graph-rest-1.0 |


### User Data 

active_user_detail.csv is downloaded with usernames and email addresses; however, that information is **removed** as part of the script.

## How to use the Graph Script

Run the script the following command in PowerShell:

    ./GetM365Data.ps1

The script defaults to analyzing the last 30 days. To modify this you can modify this with the -Days flag with either 7, 30, 90, and 180 days. 

    ./GetM365Data.ps1 -Days 90

You will be prompted to log into your MS account which will require Admin privileges. Doing this will create a new PowerShell related Enterprise Application in your Azure AD. 

It will call the API and download a series of csv files, these are then imported back into the script and are used export the json file in the correct format.

Exchange Archive will be set to zeros, you will need to run the second script for this information and manually input the results into the calculator after importing the json file into the calculator.

### Sign out of Graph

To sign out of Graph you can manually enter the following after you have finished with the script.

    Disconnect-MgGraph

### Removing PowerShell SDK Access

Please check your Azure AD account and delete the PowerShell Enterprise Application named "Microsoft Graph PowerShell" if you do not wish to retain access.     

## How to use Exchange Online Management Script

Similar to the above the above, run the script:

    ./GetArchiveMailboxStates.ps1

Log in when prompted, once complete it will display the results that can be inputted manually into the calculator. 

## How the calculation works

Change-Rate
- The first and last capacity values are taken from the storage reports
- The difference in values is calculated
- The difference is divided by the most recent storage value
- The value is converted to a percentage, divided by the report's scoped days, the multiplied by 7

Capacity 
- The capacity figures used for the inputs are based on the last reported value for each application.

Sharepoint Sites
- Highest value from the Sharepoint Site counts is used from the period.

User Counts
- Values are derived from the active_user_detail.csv file
- All users which are marked as having an Exchange, OneDrive and Teams license are counted

General Notes
- If the capacity of any of the applications is less than 1TB, it will be rounded to 1TB.
- If the weekly change-rate is less that 0.01%, it will be rounded to 0.01%.

## Other Useful Data

Also, the sharepoint_sites_detail.csv is not used in the calculation but provides useful information to help planning SharePoint backups.

## Note on high change-rate

The change-rate calculation doesn't account for any unrepresentative periods of high growth e.g. migration. It is recommended to check the storage results using a graph in cases of high change-rate.

## Project Notes
Author: Ed X Howard (Veeam), Stefan Zimmermann (Veeam)

## ✍ Contributions

We welcome contributions from the community! We encourage you to create [issues](https://github.com/VeeamHub/veeam-calculators/issues/new/choose) for Bugs & Feature Requests and submit Pull Requests. For more detailed information, refer to our [Contributing Guide](CONTRIBUTING.md).

## 🤝🏾 License

* [MIT License](LICENSE)

## 🤔 Questions

If you have any questions or something is unclear, please don't hesitate to [create an issue](https://github.com/VeeamHub/veeam-calculators/issues/new/choose) and let us know!
