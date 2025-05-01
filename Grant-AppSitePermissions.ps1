<#
.SYNOPSIS
Grants an Azure AD application specific permissions (read or write) to a SharePoint site using Microsoft Graph API.

.DESCRIPTION
This script connects to Microsoft Graph API and grants specified permissions (read or write) to an Azure AD application for a particular SharePoint site. 
It retrieves the SharePoint site ID based on the provided URL, checks existing permissions, grants the requested permissions, and verifies the assignment.

.PARAMETER tenantId
The Azure AD tenant ID where the SharePoint site and application reside.

.PARAMETER Appid
The Application (client) ID of the Azure AD application to which permissions will be granted.

.PARAMETER AppDisplayName
The display name of the Azure AD application.

.PARAMETER siteUrl
The full URL of the SharePoint site to which permissions will be granted.

.PARAMETER approle
The permission level to grant to the application. Accepted values are "read" or "write".

.EXAMPLE
.\Grant-AppSitePermissions.ps1

This example runs the script with predefined parameters to grant the specified application write permissions to the SharePoint site.

.NOTES
Authors: Mike Lee
 Date: 5/1/2025

- Requires Microsoft Graph PowerShell module.
- Requires appropriate administrative permissions to grant application permissions to SharePoint sites.
- Ensure the application ID and tenant ID are correct and that the application is registered in Azure AD.

Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

#> 
# Script to grant an application access to a SharePoint site with write permissions using Microsoft Graph API

# These values should be replaced with your actual tenant, application, and site information
$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"  # The Azure AD tenant ID
$Appid = 'b8c630cd-a668-4e6a-8574-1f3cbdb43c89'     # The Azure AD application (client) ID
$AppDisplayName = "Sites.Selected - App"             # The display name of the application
$siteUrl = "https://contoso.sharepoint.com/sites/Contoso"  # The SharePoint site URL
$approle = "write"  # Permission level: "write" grants read/write access, "read" grants read-only access

# Connect to Microsoft Graph API with the necessary permissions
# Directory.ReadWrite.All - For accessing directory objects
# AppRoleAssignment.ReadWrite.All - For managing application permissions
# Sites.FullControl.All - For managing SharePoint site permissions
Connect-MgGraph -TenantId $tenantId -Scopes "Directory.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Sites.FullControl.All" -NoWelcome

# Retrieve the SharePoint site ID based on the URL
# This section uses the hostname and site path approach to get the site ID
try {
    $siteId = @()  # Initialize the site ID variable to store the retrieved site ID
    # Remove the protocol (http:// or https://) from the URL
    $siteRelativeUrl = $siteUrl -replace "^https?://", ""
    
    # Extract the hostname (e.g., contoso.sharepoint.com)
    $hostname = $siteRelativeUrl.Split('/')[0]
    
    # Extract the site path (e.g., /sites/teamsite)
    $path = "/" + ($siteRelativeUrl.Split('/', 2)[1])
    
    Write-Host "Attempting to get site using hostname: $hostname and path: $path" -ForegroundColor Green
    Write-Host ""

    # Call the Microsoft Graph API to get site information
    # Format: {hostname}:{path} - This is the Graph API format for identifying a SharePoint site
    $siteInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$path"
    
    # Extract the site ID from the response
    $siteId = $siteInfo.id  # The site ID is returned in the response
    
    # Display the retrieved site ID for confirmation
    Write-Host "Successfully retrieved site ID $siteId for URL: $siteUrl" -ForegroundColor Magenta

}
catch {
    # Error handling if the site cannot be found
    Write-Host "Error retrieving site information: $_"
}

# Function to check and display current permissions for a SharePoint site
function Get-SharePointSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteId  # The SharePoint site ID to check permissions for
    )
    
    try {
        # Get all permissions for the specified site using the Graph API
        $permissions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions"
        Write-Host "Current Site Permissions:" -ForegroundColor Yellow
        
        # Iterate through each permission entry
        foreach ($perm in $permissions.value) {
            # Combine all roles into a comma-separated string
            $perm = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions/$($perm.id)"
            $roles = $perm.roles -join ', '
            
            # Check permissions granted to specific identities (users, groups, applications)
            if ($perm.grantedToIdentities) {
                foreach ($identity in $perm.grantedToIdentities) {
                    # Display application-specific permissions
                    if ($identity.application) {
                        Write-Host "  - Application: $($identity.application.displayName) (ID: $($identity.application.id)) - Roles: $roles"
                    }
                }
            }
        }
    }
    catch {
        # Error handling for permission retrieval issues
        Write-Host "Error retrieving site permissions: $_"
    }
}

# Display existing permissions before making changes
if ($siteId) {
    Write-Host "`n--- Current permissions before granting access ---" -ForegroundColor cyan
    Get-SharePointSitePermissions -SiteId $siteId
    try {
        # Create the permission request body for the Graph API call
        # This defines what permissions to grant and to which application
        $grantAccessBody = @{
            roles               = @($approle) # The roles to grant (read or write)
            grantedToIdentities = @(
                @{
                    application = @{
                        id          = $Appid           # Application ID to grant permissions to
                        displayName = $AppDisplayName  # Display name of the application (for reference)
                    }
                }
            )
        } | ConvertTo-Json -Depth 10  # Convert to JSON with sufficient depth for nested objects

        # Attempt to grant the specified access to the site

        # Call the Graph API to create a new permission entry
        $addperms = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -Body $grantAccessBody -ContentType "application/json"
  
        # Verify that the application has been granted the requested permissions
        Write-Host "`n--- Verifying application permissions ---"
        if ($addperms) {
            try {
                # Get updated permissions
                $permissions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
        
                # Find the specific permission entry for our application
                # This filters the permissions to find the one matching our application ID
                $appPermission = $permissions.value | Where-Object { 
                    $_.grantedToIdentities -and 
                    ($_.grantedToIdentities | ForEach-Object { $_.application.id -eq $Appid }) -contains $true 
                }
        
                # Check if the application permission was found and display the results
                if ($appPermission) {
                    $roles = $appPermission.roles -join ', '
                    Write-Host "Verification successful: Application '$($Appid)' has been granted ($approle) permissions to" $siteUrl -ForegroundColor Green
                }
                else {
                    # If the permission wasn't found, it might indicate a problem with the grant
                    Write-Host "Verification failed: Could not find permissions for application ID: $Appid" -ForegroundColor Red
                }
            }
            catch {
                # Handle errors during verification
                Write-Host "Error verifying application permissions: $_"
            }
        }
        # Final status message
        Write-Host "`n--- Script completed!! ---"
        Write-Host "Please consider running the (Test-SiteAccess.ps1) from the same GitHub repository further test access." -ForegroundColor Green
    }
    catch {
        # Handle errors during the permission granting process
        Write-Host "Error granting site access: $_"
    }
}
# Clean up by disconnecting from the Microsoft Graph API
# This is important to release any authentication tokens and resources
Disconnect-MgGraph | Out-Null
