# Kasenga Kapansa - July 2021
# This script connects to Azure AD to obtain a token and uses it to call MsGraph using the client credentials grant flow
# It also writes the results to a json file.
# Goal: to obtain a list of all objectIds in a tenant and write them to a json file. This file should then be modified to include only those users who need to have their sign-in sessions and refresh tokens revoked.

# To run:
# 1. Create an Azure AD app registration and assign the Microsoft Graph Application permission: Directory.Read.All or Directory.ReadWrite.All and grant admin consent
# 2. Generate and take note of the client secret
# 3. Populate the values below: client id (app id), tenant id, client secret
# 4. Create a folder for your json file.


# This function obtains an access token using the client credentials flow
# Populate the tenantId, clientId and clientSecret
Function Get-Token
{
    $tenantId = 'tenantID'
    $clientId = 'clientID'
    $clientSecret = 'clientSecret'

    $Uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token";

    $Body = @{
            grant_type = 'client_credentials'
            client_id = $clientId
            client_secret = $clientSecret
            scope = 'https://graph.microsoft.com/.default'
        }
    $AuthResult = Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body
    
    $accessToken = $AuthResult.access_token
    return $accessToken;
}

# Build out the request header, attaching the bearer token.
$global:headerParams = @{}
Function Set-AuthHeader 
{
    $global:headerParams = @{
        "Authorization" = "Bearer $accessToken"
        }
}

# Request an access token
$accessToken = Get-Token

# Show the token to make sure we have one
Write-Host "Access Token: " $accessToken


$queryUrl = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,displayName"

$pages = 1
$theFile = "c:\temp\users.json"

do
{
    Set-AuthHeader
    Write-Host "Getting Page " $pages
    $result = Invoke-RestMethod -Method Get -Uri $queryUrl -Headers $global:headerParams
    $content = $result.value
    # Write out content to a json file
    $content | ConvertTo-Json -depth 100 | Out-File $theFile -Append
    # Get the next page
    $queryUrl = $result."@odata.nextLink"
    $pages++
 } 
until (!($queryUrl))

Write-Host "The file been saved here: " $theFile
Write-Host "End"