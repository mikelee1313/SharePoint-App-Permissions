# SharePoint-App-Permissions
Scripts to Add and check Site Level permissions to an Entra App in SharePoint

**Grant-AppSitePermissions.ps1**
This script authenticates to Microsoft Graph API using client credentials flow and tests both read and write access to a specified SharePoint site collection based on app credentials.
It performs read tests by retrieving site lists and document library items, and write tests by uploading and optionally deleting a temporary test file in the default document library.

**Test-SiteAccess.ps1**




