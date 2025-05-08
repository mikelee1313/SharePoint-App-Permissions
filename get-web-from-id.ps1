<#
.SYNOPSIS
    Retrieves the SharePoint site collection URL from a given site/web ID.

.DESCRIPTION
    This script connects to SharePoint Online Admin center using app-based authentication
    and queries the Microsoft Graph API to find a SharePoint site by its web ID.
    It then extracts and returns the site collection URL.

.PARAMETER appID
    The Azure AD application ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER tenant
    The tenant ID for the Microsoft 365 tenant.

.PARAMETER AdminUrl
    The SharePoint Admin center URL.

.PARAMETER WebId
    The GUID of the SharePoint web/site to locate.

.EXAMPLE
    .\get-web-from-id.ps1

.NOTES
    Requirements:
    - PnP PowerShell module must be installed
    - A registered Azure AD app with proper SharePoint and Graph API permissions
    - A certificate associated with the Azure AD app

#>

# Define parameters for the script
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"
$AdminUrl = "https://m365cpi13246019-admin.sharepoint.com"
$WebId = '0f4ba734-330f-47b7-af7c-c20a5ce46e25'

try {
    # Connect to SharePoint Online using certificate-based authentication
    Write-Host "Connecting to SharePoint Online Admin center..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $AdminUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
    Write-Host "Looking up web with ID: $WebId" -ForegroundColor Yellow

    # Get a Graph API access token from the existing PnP connection
    # This reuses the authentication we already established
    $graphAccessToken = Get-PnPGraphAccessToken
    
    # Prepare HTTP headers for Graph API calls
    # The token authorizes our request and we specify JSON content type
    $headers = @{
        "Authorization" = "Bearer $graphAccessToken"
        "Content-Type"  = "application/json"
    }
    
    # Define direct Graph API endpoint for site lookup using WebId
    # This URL can directly address a site if we know its exact identifier
    $graphDirectApiUrl = "https://graph.microsoft.com/v1.0/sites/$WebId"
    
    Write-Host "Querying Microsoft Graph API..." -ForegroundColor Yellow
    
    # Define search endpoint to find site by WebId
    # Search is more flexible and can find sites even with partial information
    $graphApiUrl = "https://graph.microsoft.com/v1.0/sites?search=$WebId"
        
    try {
        # Execute the Graph API request to search for sites matching our WebId
        $searchResponse = Invoke-RestMethod -Uri $graphApiUrl -Headers $headers -Method Get
        
        if ($searchResponse.value -and $searchResponse.value.Count -gt 0) {
            # Take the first result as our target site
            $site = $searchResponse.value[0]
            
            Write-Host "Site found through search!" -ForegroundColor Green
            Write-Host "Web URL: $($site.webUrl)" -ForegroundColor Green
            
            # Pattern matching to extract the site collection URL from the web URL
            # This handles both regular sites and team sites formats
            if ($site.webUrl -match "https://[^/]+(/sites/[^/]+|/teams/[^/]+)?") {
                # Extract the base match which could be just the tenant URL
                $siteCollectionUrl = $Matches[0]
                
                # Look for the standard /sites/ pattern for site collections
                if ($site.webUrl -match "(https://[^/]+/sites/[^/]+)") {
                    $siteCollectionUrl = $Matches[1]
                }
                # Look for the teams site pattern
                elseif ($site.webUrl -match "(https://[^/]+/teams/[^/]+)") {
                    $siteCollectionUrl = $Matches[1]
                }
                # If neither pattern matched, we'll use the base URL as determined earlier
            }
            else {
                $siteCollectionUrl = $site.webUrl
            }
            
            Write-Host "Site Collection URL: $siteCollectionUrl" -ForegroundColor Green
            return $siteCollectionUrl
        }
        else {
            Write-Error "No site found with ID: $WebId"
        }
    }
    catch {
        Write-Error "Error searching for site: $_"
    }
}
catch {
    Write-Error "Error retrieving site information: $_"
}
finally {

    Disconnect-PnPOnline
}
