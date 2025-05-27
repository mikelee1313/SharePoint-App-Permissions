<#
.SYNOPSIS
    Tests read and write access to a SharePoint site collection using Microsoft Graph API.

.DESCRIPTION
This script authenticates to Microsoft Graph API using client credentials flow and tests both read and write access to a specified SharePoint site collection. 
It performs read tests by retrieving site lists and document library items, and write tests by uploading and optionally deleting a temporary test file in the default document library.

.PARAMETER TestType
    Specifies the type of access test to perform. Valid values are:
    - Read: Performs only read access tests.
    - Write: Performs only write access tests.
    - Both (default): Performs both read and write access tests.

.EXAMPLE
    Test-SiteAccess
    Performs both read and write access tests on the specified SharePoint site.

.EXAMPLE
    Test-SiteAccess -TestType "Read"
    Performs only read access tests on the specified SharePoint site.

.EXAMPLE
    Test-SiteAccess -TestType "Write"
    Performs only write access tests on the specified SharePoint site.

.NOTES

Authors: Mike Lee
 Date: 4/30/2025

- Requires Microsoft Graph PowerShell SDK (Connect-MgGraph)
- Appropriate permissions configured in Azure AD.
- Ensure the provided client credentials have sufficient permissions to perform the intended operations.
- Requires PowerShell 7.X


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

# Configuration parameters for connecting to SharePoint through Microsoft Graph
$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"     # Your Azure AD tenant ID
$ClientId = 'b8c630cd-a668-4e6a-8574-1f3cbdb43c89'      # App registration client ID
$clientSecret = '' # App client secret
$siteUrl = "https://contoso.sharepoint.com/sites/it" # Target SharePoint site URL

# Parse the site URL into components needed for Graph API calls
$siteRelativeUrl = $siteUrl -replace "^https?://", ""   # Remove protocol (http:// or https://)
$hostname = $siteRelativeUrl.Split('/')[0]              # Extract the hostname portion
$path = "/" + ($siteRelativeUrl.Split('/', 2)[1])       # Extract the path portion

# Authentication parameters for OAuth token acquisition
$scopes = "https://graph.microsoft.com/.default"        # Default Graph API scope
$loginURL = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" # OAuth endpoint
$body = @{grant_type = "client_credentials"; client_id = $ClientId; client_secret = $ClientSecret; scope = $scopes }
    
# Acquire OAuth token using client credentials flow
$Token = Invoke-RestMethod -Method Post -Uri $loginURL -Body $body
$headerParams = @{'Authorization' = "$($Token.token_type) $($Token.access_token)" } # Create authorization header

# Connect to Microsoft Graph with the acquired token
$secureToken = ConvertTo-SecureString $Token.access_token -AsPlainText -Force
Connect-MgGraph -AccessToken $secureToken -NoWelcome # Connect without welcome message

# Initial site information retrieval to validate connectivity and get the site ID
Write-Host "Attempting to get site using hostname: $hostname and path: $path"
$siteInfo = @()
$siteId = @()

try {
    # Retrieve site information using Microsoft Graph API
    $siteInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$path" -Headers $headerParams

    $siteId = $siteInfo.id
    Write-Host "Site ID: $siteId"
}
catch {
    Write-Host "Error retrieving site information: $_"
}

function Test-SiteAccess {
    <#
    .SYNOPSIS
        Main function to test access permissions to a SharePoint site.
    .DESCRIPTION
        Tests read and/or write access to the specified SharePoint site based on the TestType parameter.
        Performs detailed checks and provides comprehensive output on the results.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet("Read", "Write", "Both")]
        [string]$TestType = "Both" # Default to testing both read and write access
    )

    # Results tracking object to store test outcomes
    $results = @{
        ReadSuccess  = $null  # Will be set to $true or $false based on read test results
        WriteSuccess = $null  # Will be set to $true or $false based on write test results
    }
    
    # Display test information
    Write-Host "Testing access to site: $siteUrl" -ForegroundColor Cyan
    Write-Host "Test type: $TestType" -ForegroundColor Cyan
    
    #region READ ACCESS TESTING
    # Perform Read Test if TestType is Read or Both
    if ($TestType -eq "Read" -or $TestType -eq "Both") {
        Write-Host "`n=== READ ACCESS TEST ===" -ForegroundColor Cyan
        
        try {
            # Test 1: Get site lists - this tests basic read permission to the site
            Write-Host "Test 1: Retrieving site lists..." -NoNewline
            $listsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists"
            $lists = Invoke-MgGraphRequest -Method GET -Uri $listsUri -Headers $headerParams
            
            # Evaluate results of the lists retrieval
            if ($lists -and $lists.value) {
                Write-Host "SUCCESS" -ForegroundColor Green
                Write-Host "Found $($lists.value.Count) lists in the site."
                
                # Display sample lists to provide more context in the output
                if ($lists.value.Count -gt 0) {
                    Write-Host "Sample lists:" -ForegroundColor Cyan
                    $lists.value | Select-Object -First 3 | ForEach-Object {
                        Write-Host "  - $($_.displayName) (ID: $($_.id))"
                    }
                }
            }
            else {
                Write-Host "FAILED" -ForegroundColor Red
                Write-Host "Could not retrieve lists, but no error was thrown."
                $results.ReadSuccess = $false
            }
            
            # Test 2: Get document library items - tests more specific read permissions
            Write-Host "Test 2: Retrieving document library items..." -NoNewline
            try {
                # First get the drives (document libraries) in the site
                $driveInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Headers $headerParams
                # Find the default document library (usually named "Documents" or "Shared Documents")
                $defaultDrive = $driveInfo.value | Where-Object { $_.name -eq 'Documents' -or $_.name -eq 'Shared Documents' } | Select-Object -First 1
                
                if ($defaultDrive) {
                    # If we found a default document library, get its contents
                    $driveId = $defaultDrive.id
                    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root/children"
                    $items = Invoke-MgGraphRequest -Method GET -Uri $itemsUri -Headers $headerParams
                    
                    # Successfully retrieved document library items
                    Write-Host "SUCCESS" -ForegroundColor Green
                    Write-Host "Found $($items.value.Count) items in the default document library."
                    $results.ReadSuccess = $true
                }
                else {
                    # Document libraries exist but couldn't find the default one
                    Write-Host "PARTIAL" -ForegroundColor Yellow
                    Write-Host "Could retrieve drives but default document library not found."
                    $results.ReadSuccess = $false
                }
            }
            catch {
                # Error accessing document library items
                Write-Host "FAILED" -ForegroundColor Red
                Write-Host "Could not retrieve document library items: $_"
                $results.ReadSuccess = $false
            }
            
            # Final read access result summary
            if ($results.ReadSuccess) {
                Write-Host "Read access to the site collection is working properly." -ForegroundColor Green
            }
            else {
                Write-Host "Read access test failed." -ForegroundColor Red
            }
        }
        catch {
            # Catch-all for any unexpected errors during read testing
            Write-Host "ERROR testing read access: $_" -ForegroundColor Red
            $results.ReadSuccess = $false
        }
    }
    #endregion
    
    #region WRITE ACCESS TESTING
    # Perform Write Test if TestType is Write or Both
    if ($TestType -eq "Write" -or $TestType -eq "Both") {
        Write-Host "`n=== WRITE ACCESS TEST ===" -ForegroundColor Cyan
        
        try {
            # Create a temporary test file with timestamp to avoid name conflicts
            $tempFileName = "WriteAccessTest_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            # Generate file content that includes the creation timestamp
            $tempFileContent = [System.Text.Encoding]::UTF8.GetBytes("This is a test file created at $(Get-Date) to verify write access.")
            
            # Get the default document library for upload testing
            Write-Host "Retrieving default document library..." -NoNewline
            $driveInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Headers $headerParams
            $defaultDrive = $driveInfo.value | Where-Object { $_.name -eq 'Documents' -or $_.name -eq 'Shared Documents' } | Select-Object -First 1
            
            if (-not $defaultDrive) {
                # Can't proceed with write test if document library isn't found
                Write-Host "FAILED" -ForegroundColor Red
                Write-Host "Default document library not found"
                $results.WriteSuccess = $false
            }
            else {
                Write-Host "SUCCESS" -ForegroundColor Green
                
                $driveId = $defaultDrive.id
                
                # Upload the test file to verify write permissions
                Write-Host "Attempting to upload test file ($tempFileName)..." -NoNewline
                $uploadUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root:/$tempFileName`:/content"
                
                try {
                    # Use PUT request to upload file content directly
                    $response = Invoke-MgGraphRequest -Method PUT -Uri $uploadUrl -Headers $headerParams -Body $tempFileContent -ContentType "text/plain"
                    
                    # Upload succeeded - write access confirmed
                    Write-Host "SUCCESS" -ForegroundColor Green
                    Write-Host "Test file uploaded successfully: $($response.webUrl)"
                    $results.WriteSuccess = $true
                    
                    # Clean up by deleting the test file (good practice)
                    Write-Host "Cleaning up test file..." -NoNewline
                    $deleteUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$($response.id)"
                    try {
                        Invoke-MgGraphRequest -Method DELETE -Uri $deleteUrl -Headers $headerParams
                        Write-Host "SUCCESS" -ForegroundColor Green
                    }
                    catch {
                        # Cleanup failed, but the write test was still successful
                        Write-Host "FAILED (Cleanup failed but write test succeeded)" -ForegroundColor Yellow
                    }
                }
                catch {
                    # File upload failed - write access denied
                    Write-Host "FAILED" -ForegroundColor Red
                    Write-Host "Could not upload test file: $_"
                    $results.WriteSuccess = $false
                }
            }
            
            # Final write access result summary
            if ($results.WriteSuccess) {
                Write-Host "Write access to the site collection is working properly." -ForegroundColor Green
            }
            else {
                Write-Host "Write access test failed." -ForegroundColor Red
            }
        }
        catch {
            # Catch-all for any unexpected errors during write testing
            Write-Host "ERROR testing write access: $_" -ForegroundColor Red
            $results.WriteSuccess = $false
        }
    }
    #endregion
    
    #region TEST SUMMARY
    # Display final summary of all completed tests
    Write-Host "`n=== TEST SUMMARY ===" -ForegroundColor Cyan
    if ($TestType -eq "Read" -or $TestType -eq "Both") {
        # Use ternary operator to display SUCCESS or FAILED based on test results
        Write-Host "Read Access: $($results.ReadSuccess -eq $true ? "SUCCESS" : "FAILED")" -ForegroundColor ($results.ReadSuccess -eq $true ? "Green" : "Red")
    }
    if ($TestType -eq "Write" -or $TestType -eq "Both") {
        Write-Host "Write Access: $($results.WriteSuccess -eq $true ? "SUCCESS" : "FAILED")" -ForegroundColor ($results.WriteSuccess -eq $true ? "Green" : "Red")
    }
    #endregion
    
    # Return results object for potential use by calling code
    return $results
}

# Example usage of the Test-SiteAccess function
# Run default test (both read and write)
Test-SiteAccess

# Additional usage examples (commented out)
# Test only write access
#Test-SiteAccess -TestType "Write"

# Test only read access
#Test-SiteAccess -TestType "Read"
