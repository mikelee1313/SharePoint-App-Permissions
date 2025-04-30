# SharePoint-App-Permissions
Scripts to Add and check Site Level permissions to an Entra App in SharePoint. 

This script is very useful when adding  permissions to SharePoint Sites when the App is configured with Sites.Selected.

**Grant-AppSitePermissions.ps1**
This script uses Microsoft Graph PowerShell module. 
Connects to Microsoft Graph API and grants specified permissions (read or write) to an Azure AD application for a particular SharePoint site. 
It retrieves the SharePoint site ID based on the provided URL, checks existing permissions, grants the requested permissions, and verifies the assignment.

All you need is to specify these paramaters:

Example:

$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"  # The Azure AD tenant ID

$Appid = 'b8c630cd-a668-4e6a-8574-1f3cbdb43c89'     # The Azure AD application (client) ID

$AppDisplayName = "Sites.Selected - App"             # The display name of the application

$siteUrl = "https://contoso.sharepoint.com/sites/it"  # The SharePoint site URL

$approle = "write"  # Permission level: "write" grants read/write access, "read" grants read-only access


**Test-SiteAccess.ps1**
This script authenticates to Microsoft Graph API using client credentials flow and tests both read and write access to a specified SharePoint site collection. 
It performs read tests by retrieving site lists and document library items, and write tests by uploading and optionally deleting a temporary test file in the default document library.
