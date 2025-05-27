<#
.SYNOPSIS
Tests SharePoint site access permissions using the Microsoft Graph API with certificate-based authentication.

.DESCRIPTION
This script verifies if an Azure AD application can access a specified SharePoint site using certificate-based authentication.
It performs read and/or write tests to validate proper permissions have been configured for the application.

The script performs these main functions:
1. Authenticates to Microsoft Graph API using a certificate
2. Retrieves SharePoint site information
3. Tests read access by retrieving site lists and document library content
4. Tests write access by creating and deleting a temporary file
5. Provides a detailed summary of test results

.PARAMETER tenantId
The Azure AD/Microsoft Entra ID tenant ID.

.PARAMETER ClientId
The application (client) ID of the Azure AD app registration.

.PARAMETER siteUrl
The complete URL of the SharePoint site to test.

.PARAMETER pfxPath
The file path to the PFX certificate file used for authentication.

.PARAMETER pfxPassword
The password for the PFX certificate file.

.PARAMETER testType
The type of tests to perform. Valid values: "Read", "Write", "Both". Default is "Both".

.EXAMPLE
This runs the script with default parameters, testing both read and write access to the specified SharePoint site.
.\Test-Graph-SiteAccess-with-CertAuth.ps1

.NOTES
File Name      : Test-Graph-SiteAccess-with-CertAuth.ps1
Author         : Mike Lee
Date:          : 5/25/25
 Prerequisite   : PowerShell 5.1 or later
                     Valid Azure AD app registration with certificate authentication
                     PFX certificate file

Author: Microsoft SharePoint Customer Engineering Team
#>

#######################################################
# CONFIGURATION SETTINGS - MODIFY THESE VALUES TO MATCH YOUR ENVIRONMENT
#######################################################

# Azure AD / Microsoft Entra ID configuration
$tenantId = "85612ccb-4c28-4a34-88df-a538cc139a51"  # Your Azure AD/Microsoft Entra ID tenant ID
$ClientId = '5baa1427-1e90-4501-831d-a8e67465f0d9'  # Your app registration client ID (Application ID)

# SharePoint site configuration
$siteUrl = "https://m365x61250205.sharepoint.com/sites/commsite1"  # The SharePoint site URL to test

# Certificate configuration
$pfxPath = "C:\Users\michlee\OneDrive - Microsoft\SfMC_Docs\EEEU\08391064707223D84E33F271936DE80E92ED4F9C.pfx"  # Path to your PFX certificate file
$pfxPassword = "pass"  # Password for the PFX file

# Optional configuration
$testType = "Both"  # Valid values: "Read", "Write", "Both"

#######################################################
# END OF CONFIGURATION SETTINGS
#######################################################

# Parse the site URL into components needed for Graph API calls
$siteRelativeUrl = $siteUrl -replace "^https?://", ""   # Remove protocol (http:// or https://)
$hostname = $siteRelativeUrl.Split('/')[0]              # Extract the hostname portion
$path = "/" + ($siteRelativeUrl.Split('/', 2)[1])       # Extract the path portion

# Load the certificate first
try {
    # Create a secure string for the password
    $securePassword = ConvertTo-SecureString -String $pfxPassword -AsPlainText -Force
    # Load the certificate with the secure password
    $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, $securePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
    Write-Host "Certificate loaded successfully: $($certificate.Subject)" -ForegroundColor Green
}
catch {
    Write-Host "Error loading certificate: $_" -ForegroundColor Red
    exit
}

# Authentication using certificate
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Create a JWT header
$jwtHeader = @{
    alg = "RS256"
    typ = "JWT"
    x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
} | ConvertTo-Json -Compress

# Convert to Base64
$jwtHeaderBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)).
Replace('+', '-').
Replace('/', '_').
TrimEnd('=')

# Create JWT payload with claims
$now = [int](Get-Date -UFormat %s)
$exp = $now + 3600  # Token valid for 1 hour

$jwtPayload = @{
    aud = $tokenEndpoint
    exp = $exp
    iss = $clientId
    jti = [guid]::NewGuid()
    nbf = $now
    sub = $clientId
} | ConvertTo-Json -Compress

# Convert to Base64
$jwtPayloadBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)).
Replace('+', '-').
Replace('/', '_').
TrimEnd('=')

# Create the signature
$toSign = $jwtHeaderBase64 + "." + $jwtPayloadBase64
$rsa = $certificate.PrivateKey
$signature = [Convert]::ToBase64String($rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($toSign), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)) -replace '\+', '-' -replace '/', '_' -replace '='

# Create the complete JWT token
$jwt = $jwtHeaderBase64 + "." + $jwtPayloadBase64 + "." + $signature

# Get OAuth token
$body = @{
    client_id             = $clientId
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    client_assertion      = $jwt
    scope                 = "https://graph.microsoft.com/.default"
    grant_type            = "client_credentials"
}

# Acquire OAuth token using certificate authentication
try {
    $Token = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body
    $headerParams = @{'Authorization' = "$($Token.token_type) $($Token.access_token)" }
    
    # Connect to Microsoft Graph with the acquired token
    $secureToken = ConvertTo-SecureString $Token.access_token -AsPlainText -Force
    Connect-MgGraph -AccessToken $secureToken -NoWelcome # Connect without welcome message
    
    Write-Host "Successfully authenticated using certificate" -ForegroundColor Green
}
catch {
    Write-Host "Error authenticating with certificate: $_" -ForegroundColor Red
    exit
}

# Initial site information retrieval to validate connectivity and get the site ID
Write-Host "Getting site information for: $hostname$path"
try {
    # Retrieve site information using Microsoft Graph API
    $siteInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$path" -Headers $headerParams

    $siteId = $siteInfo.id
    Write-Host "Site ID: $siteId" -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving site information: $_" -ForegroundColor Red
    exit
}

function Test-SiteAccess {
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
            $lists = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists" -Headers $headerParams
            
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
                    $items = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root/children" -Headers $headerParams
                    
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
                try {
                    # Use PUT request to upload file content directly
                    $response = Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root:/$tempFileName`:/content" -Headers $headerParams -Body $tempFileContent -ContentType "text/plain"
                    
                    # Upload succeeded - write access confirmed
                    Write-Host "SUCCESS" -ForegroundColor Green
                    Write-Host "Test file uploaded successfully: $($response.webUrl)"
                    $results.WriteSuccess = $true
                    
                    # Clean up by deleting the test file (good practice)
                    Write-Host "Cleaning up test file..." -NoNewline
                    try {
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$($response.id)" -Headers $headerParams
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
