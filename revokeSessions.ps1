# Kasenga Kapansa - July 2021
# This script connects to Azure AD to obtain a token and uses it to call MsGraph using the client credentials grant flow
# It also reads object ids from a json file
#Goal: to revoke all sign-ins and refresh tokens listed in a json file

# To run:
# 1. Create an Azure AD app registration and assign the Microsoft Graph Application permission: Directory.Read.All or Directory.ReadWrite.All and grant admin consent
# 2. Generate and take note of the client secret
# 3. Populate the values below: client id (app id), tenant id, client secret
# 4. Create a json file with a headings of "id", "userPrincipalName" and include the list of object ids you intend to revoke sign-ins for.


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

# Build out the request header, attaching the access token and setting the content type ...
$global:headerParams = @{}
Function Set-AuthHeader {
$global:headerParams = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type" = "application/json"
    }
}

#Request an access token
$accessToken = Get-Token

# Show the access token - just to make sure that we have one
Write-Host "Access Token: " $accessToken

# The calls that we will make ...
$revokeSignInSessionsQuery = "https://graph.microsoft.com/v1.0/{id}/revokeSignInSessions"

Set-AuthHeader

# Read the list of objectIds from the csv file. You may update the location
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertfrom-json?view=powershell-7.1
$users = Get-Content "c:\temp\users.json" | ConvertFrom-Json
 
# Loop through the file and make two graph calls for each object
foreach($user in $users)
{
    try 
    {
        $Id = $user.id
        $revokeSignInSessionsQuery = "https://graph.microsoft.com/v1.0/users/$Id/revokeSignInSessions"          # This is the new call.
        
        # Show the revoke sign-ins call ...
        Write-Host "Post $revokeSignInSessionsQuery"
        Invoke-RestMethod -Method Post -Uri $revokeSignInSessionsQuery -Headers $global:headerParams
    }
    catch 
    {
        # Check the error code, looking for a 429 ...
        $errorCode = $_.Exception.Response.StatusCode.value__
        Write-Host "Caught an exception: $errorCode"
        
        # https://stackoverflow.com/questions/29613572/error-handling-for-invoke-restmethod-powershell
        # Error handling ...
        switch ($errorCode) 
        {
            429 
            {
                # Error handling for 429, set initial delay to 10 seconds
                $delay = 10
                # Show assumed delay
                Write-Host ${"$errorCode: Throttled: $delay"}

                [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0])
                # Show delay sent from MsGraph
                Write-Host "429: Throttled: $delay"

                #Wait, then try again
                Start-Sleep $delay
                
                #Re-try the calls here ...
                Write-Host "Re-trying Post $revokeSignInSessionsQuery"
                Invoke-RestMethod -Method Post -Uri $revokeSignInSessionsQuery -Headers $global:headerParams

                # Show the invalidate refresh token call ... [This seems to be for legacy]
                Write-Host "Re-trying Post $invalidateAllRefreshTokens"
                Invoke-RestMethod -Method Post -Uri $invalidateAllRefreshTokens -Headers $global:headerParams
                break
            }

            400 
            {
                # Error handling for 400. For the sign-in endpoint, this could also mean that the internally gnerated skiptoken, used to go to the next page of results, is not yet valid
                #The skip token is different from the access token used to make the Graph call.
                Write-Host "400 invalid clause or invalid skip token"
                break
            }

            404 
            {
                # Error handling for 404
                Write-Host "404: Not found"
            }
            401 
            {
                # Error handling for 401 - the token has expired, renew it here ...
                Write-Host "401: Unauthorized"
                $accessToken = Get-Token
                Set-AuthHeader
                break
            }
            
            default 
            {
                # Handle misc errors ...
            }
        }
    }
    
}

Write-Host "end"