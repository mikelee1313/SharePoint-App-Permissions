# SharePoint-App-Permissions

Scripts to Add and check Site Level permissions from an Entra App in SharePoint using Graph PowerShell.

These scripts are very useful when adding permissions to SharePoint Sites when the App is configured with Sites.Selected.

Example:

![image](https://github.com/user-attachments/assets/eff92e25-8bf8-4098-8b3e-0e5eb7a29668)


![image](https://github.com/user-attachments/assets/a9587c22-50ed-40e6-bd76-8c273747725a)



**Grant-AppSitePermissions.ps1**

This script uses Microsoft Graph PowerShell module. 
Connects to Microsoft Graph API and grants specified permissions (read or write) to an Azure AD application for a particular SharePoint site. 
It retrieves the SharePoint site ID based on the provided URL, checks existing permissions, grants the requested permissions, and verifies the assignment.

**Requirements:**

- Requires PowerShell 7.X
- Tested with PNP 2.12.0  
- Requires Microsoft Graph PowerShell module.
- Requires appropriate administrative permissions to grant application permissions to SharePoint sites.
- Ensure the application ID and tenant ID are correct and that the application is registered in Azure AD.
  



Specify these parameters:

Example:

$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"  # The Azure AD tenant ID

$Appid = 'b8c630cd-a668-4e6a-8574-1f3cbdb43c89'     # The Azure AD application (client) ID

$AppDisplayName = "Sites.Selected - App"             # The display name of the application

$siteUrl = "https://contoso.sharepoint.com/sites/it"  # The SharePoint site URL

$approle = "write"  # Permission level: "write" grants read/write access, "read" grants read-only access

Example Output:

![image](https://github.com/user-attachments/assets/16f0d01f-23c4-4762-9789-764fdd6663b5)




**Test-SiteAccess.ps1**

This script authenticates to Microsoft Graph API using client credentials flow and tests both read and write access to a specified SharePoint site collection. 
It performs read tests by retrieving site lists and document library items, and write tests by uploading and optionally deleting a temporary test file in the default document library.

You will need your app information to create your access token to perform the test.

Example:

$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"     # Your Azure AD tenant ID

$ClientId = 'b8c630cd-a668-4e6a-8574-1f3cbdb43c89'      # App registration client ID

$clientSecret = '' # App client secret

$siteUrl = "https://m365cpi13246019.sharepoint.com/sites/it" # Target SharePoint site URL

Then the the script will use "Connect-MgGraph -AccessToken $secureToken" to gain access to the site collection.

Example commands:

- Test-SiteAccess

  Performs both read and write access tests on the specified SharePoint site.

- Test-SiteAccess -TestType "Read"

  Performs only read access tests on the specified SharePoint site.

- Test-SiteAccess -TestType "Write"

  Performs only write access tests on the specified SharePoint site.


Here is an example of the output:

![image](https://github.com/user-attachments/assets/55513015-77a6-491f-a14f-c64bf18c371a)


