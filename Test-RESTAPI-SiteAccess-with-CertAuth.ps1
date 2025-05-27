<#
.SYNOPSIS
    Generates a JWT client assertion token for Azure AD authentication and performs SharePoint REST API operations.

.DESCRIPTION
    This script demonstrates how to authenticate with Azure AD using certificate-based authentication and perform
    operations against the SharePoint REST API. It creates a client assertion JWT token, acquires access tokens
    for both Microsoft Graph and SharePoint, and then performs read and write operations against a SharePoint site.

.PARAMETER None
    This script does not accept parameters from the command line. Configure the script variables in the USER
    CONFIGURATION section before running.

.NOTES
    File Name      : Test-RESTAPI-SiteAccess-with-CertAuth.ps1
    Author         : Mike Lee
    Date:          : 5/25/25
    Prerequisite   : PowerShell 5.1 or later
                     Valid Azure AD app registration with certificate authentication
                     PFX certificate file
.EXAMPLE
    PS> .\Test-RESTAPI-SiteAccess-with-CertAuth.ps1

.COMPONENT
    - Requires an Azure AD application registration with appropriate permissions:
    - Microsoft Graph: Application permissions
    - SharePoint: Application permissions (Read and Write)

.INPUTS
    None. Configuration is done via variables in the script.

.OUTPUTS
    Console output showing the results of SharePoint REST API operations.

.LINK
    https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/using-csom-for-dotnet-standard

.NOTES
    Configuration Variables:
    - $tenantId: Azure AD tenant ID
    - $clientId: Azure AD application (client) ID
    - $pfxPath: Path to the PFX certificate file
    - $pfxPassword: Password for the PFX file
    - $siteUrl: URL of the SharePoint site
#>
# This script generates a client assertion JWT for Azure AD using a certificate and requests an access token from the v2.0 endpoint.

#########################################################################
# USER CONFIGURATION - MODIFY THESE SETTINGS BEFORE RUNNING THE SCRIPT
#########################################################################

# Azure AD / App Registration settings
$tenantId = "85612ccb-4c28-4a34-88df-a538cc139a51"     # Your Azure AD tenant ID
$clientId = '5baa1427-1e90-4501-831d-a8e67465f0d9'     # App registration client ID

# Certificate settings
$pfxPath = "C:\temp\08391064707223D84E33F271936DE80E92ED4F9C.pfx"  # Path to your PFX file
$pfxPassword = ""  # Password for the PFX file, leave empty if no password

# SharePoint settings
$siteUrl = "https://m365x61250205.sharepoint.com/sites/commsite1"  # URL of the SharePoint site

#########################################################################
# END OF USER CONFIGURATION
#########################################################################

# Extract hostname from site URL (don't modify)
$hostname = ($siteUrl -replace "^https?://", "").Split('/')[0]

# Load certificate from PFX file
if ([string]::IsNullOrEmpty($pfxPassword)) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath)
}
else {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, $pfxPassword)
}

# Create JWT header and payload
$now = [System.DateTime]::UtcNow
$exp = $now.AddMinutes(10)
$jwtHeader = @{
    alg = "RS256"
    typ = "JWT"
    x5t = [System.Convert]::ToBase64String($cert.GetCertHash())
}
$jwtPayload = @{
    aud = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"  # Updated audience for v2.0 endpoint
    iss = $clientId
    sub = $clientId
    jti = [guid]::NewGuid().ToString()
    nbf = [int][double]::Parse((Get-Date $now -UFormat %s))
    exp = [int][double]::Parse((Get-Date $exp -UFormat %s))
}

# Convert to JSON and base64url encode
$headerJson = ($jwtHeader | ConvertTo-Json -Compress)
$payloadJson = ($jwtPayload | ConvertTo-Json -Compress)
$headerEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
$payloadEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
$unsignedJwt = "$headerEncoded.$payloadEncoded"

# Sign the JWT using a more compatible approach
$sha256 = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
$signatureFormatter = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($cert.PrivateKey)
$signatureFormatter.SetHashAlgorithm("SHA256")
$bytes = [System.Text.Encoding]::UTF8.GetBytes($unsignedJwt)
$hash = $sha256.ComputeHash($bytes)
$signature = $signatureFormatter.CreateSignature($hash)
$signatureEncoded = [Convert]::ToBase64String($signature).TrimEnd('=').Replace('+', '-').Replace('/', '_')
$clientAssertion = "$unsignedJwt.$signatureEncoded"

# Prepare token request body
$body = @{
    grant_type            = "client_credentials"
    client_id             = $clientId
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    client_assertion      = $clientAssertion
    scope                 = "https://graph.microsoft.com/.default"  # Standard scope for Microsoft Graph
}

# Request token from modern v2.0 endpoint
$response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body -ContentType "application/x-www-form-urlencoded"
$graphToken = $response.access_token

# We can use the same client assertion as for Graph
# Get SharePoint token using v2.0 endpoint - just change the scope/audience
$spBody = @{
    grant_type            = "client_credentials"
    client_id             = $clientId  # No tenant suffix needed for v2.0 endpoint
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    client_assertion      = $clientAssertion  # Use the same assertion as for Graph
    scope                 = "https://$hostname/.default"  # Use SharePoint domain scope
}

# Use the v2.0 endpoint - same as Graph
$spResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $spBody -ContentType "application/x-www-form-urlencoded"
$spToken = $spResponse.access_token

# Create headers for SharePoint REST API calls
$spHeaders = @{
    'Authorization' = "Bearer $spToken"
    'Accept'        = 'application/json;odata=verbose'
}

# Function to perform SharePoint REST API read operations
function Test-SharePointRestRead {
    Write-Host "`n=== SharePoint REST API Read Operations ===" -ForegroundColor Cyan
    
    # Test 1: Get site information
    Write-Host "Test 1: Getting site information..." -ForegroundColor Yellow
    try {
        $siteInfoUri = "$siteUrl/_api/web?`$select=Title,Id"
        $siteInfo = Invoke-RestMethod -Uri $siteInfoUri -Headers $spHeaders -Method Get
        
        Write-Host "SUCCESS: Retrieved site information" -ForegroundColor Green
        Write-Host "  Site Title: $($siteInfo.d.Title)" -ForegroundColor Green
        Write-Host "  Site ID: $($siteInfo.d.Id)" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to get site information: $_" -ForegroundColor Red
    }
    
    # Test 2: Get lists in the site
    Write-Host "`nTest 2: Getting site lists..." -ForegroundColor Yellow
    try {
        $listsUri = "$siteUrl/_api/web/lists?`$select=Title,Id,ItemCount&`$filter=(Title eq 'Documents' or Title eq 'Shared Documents')&`$top=5"
        $lists = Invoke-RestMethod -Uri $listsUri -Headers $spHeaders -Method Get
        
        Write-Host "SUCCESS: Retrieved $($lists.d.results.Count) lists" -ForegroundColor Green
        
        # Display some list details
        foreach ($list in $lists.d.results | Select-Object -First 3) {
            Write-Host "  - List: $($list.Title) (Items: $($list.ItemCount))" -ForegroundColor Green
        }
        
        # Return the lists for use by the write operation
        return $lists.d.results
    }
    catch {
        Write-Host "ERROR: Failed to get lists: $_" -ForegroundColor Red
        return $null
    }
}

# Function to perform SharePoint REST API write operations
function Test-SharePointRestWrite {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Lists
    )
    
    Write-Host "`n=== SharePoint REST API Write Operations ===" -ForegroundColor Cyan
    
    # Find the Documents library
    $docsLibrary = $Lists | Where-Object { $_.Title -eq "Documents" -or $_.Title -eq "Shared Documents" } | Select-Object -First 1
    
    if (-not $docsLibrary) {
        Write-Host "No Documents library found for testing write operations" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Using library '$($docsLibrary.Title)' for write operations" -ForegroundColor Yellow
    
    # Get request digest (required for write operations)
    try {
        $digestUri = "$siteUrl/_api/contextinfo"
        $digestHeaders = $spHeaders.Clone()
        $digestResult = Invoke-RestMethod -Uri $digestUri -Headers $digestHeaders -Method Post
        $requestDigest = $digestResult.d.GetContextWebInformation.FormDigestValue
        
        Write-Host "Successfully acquired request digest" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to get request digest: $_" -ForegroundColor Red
        return
    }
    
    # Create headers for write operations
    $writeHeaders = $spHeaders.Clone()
    $writeHeaders["X-RequestDigest"] = $requestDigest
    $writeHeaders["Content-Type"] = "application/json;odata=verbose"
    
    # Generate a unique filename
    $fileName = "TestFile_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $fileContent = "This is a test file created by PowerShell on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    
    # Create a temporary file
    $tempFilePath = [System.IO.Path]::GetTempFileName()
    Set-Content -Path $tempFilePath -Value $fileContent -Encoding UTF8
    
    # Upload file to document library
    Write-Host "`nTest 1: Uploading test file..." -ForegroundColor Yellow
    
    try {
        # Get the server relative URL of the document library
        $libraryUrl = "$siteUrl/_api/web/lists(guid'$($docsLibrary.Id)')/RootFolder"
        $libraryInfo = Invoke-RestMethod -Uri $libraryUrl -Headers $spHeaders -Method Get
        $serverRelativeUrl = $libraryInfo.d.ServerRelativeUrl
        
        Write-Host "  Document library path: $serverRelativeUrl" -ForegroundColor Gray
        
        $uploadUri = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$serverRelativeUrl')/Files/add(url='$fileName',overwrite=true)"
        
        # For file upload, we need different headers
        $uploadHeaders = $spHeaders.Clone()
        $uploadHeaders["X-RequestDigest"] = $requestDigest
        
        # Read file content as bytes
        $fileBytes = [System.IO.File]::ReadAllBytes($tempFilePath)
        
        # Upload the file
        $uploadResponse = Invoke-RestMethod -Uri $uploadUri -Headers $uploadHeaders -Method Post -Body $fileBytes
        
        Write-Host "SUCCESS: Uploaded file '$fileName'" -ForegroundColor Green
        
        # Get the server relative URL of the uploaded file
        $fileServerRelativeUrl = $uploadResponse.d.ServerRelativeUrl
        
        # Delete the file
        Write-Host "`nTest 2: Deleting test file..." -ForegroundColor Yellow
        
        $deleteHeaders = $writeHeaders.Clone()
        $deleteHeaders["IF-MATCH"] = "*"
        $deleteHeaders["X-HTTP-Method"] = "DELETE"
        
        $deleteFileUri = "$siteUrl/_api/web/GetFileByServerRelativeUrl('$fileServerRelativeUrl')"
        
        Invoke-RestMethod -Uri $deleteFileUri -Headers $deleteHeaders -Method Post
        Write-Host "SUCCESS: Deleted file '$fileName'" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: SharePoint file operation failed: $_" -ForegroundColor Red
    }
    finally {
        # Clean up temporary file
        if (Test-Path $tempFilePath) {
            Remove-Item -Path $tempFilePath -Force
        }
    }
}

# Execute the tests
Write-Host "`n==== SharePoint REST API Tests ====" -ForegroundColor Cyan
Write-Host "Site URL: $siteUrl" -ForegroundColor Cyan

# Perform read tests and capture the lists for write tests
$siteLists = Test-SharePointRestRead

# Perform write tests if we have lists
if ($siteLists) {
    Test-SharePointRestWrite -Lists $siteLists
}
