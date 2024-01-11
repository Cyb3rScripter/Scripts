# First create an app registration with Mail.Send permissions. Authentication using certificate
# Acquire an access token to interact with the app
# PREREQUISITE - Certificate is required from the user's personal store
# Run script (Create App Registration Tickets.ps1) to create cert in currentuser cert store
# update cert thumbprint in ClientCertificate below - current cert expires 01/01/2023
# Requires CSV file with corresponding headers
Import-Module -Name MSAL.PS -Force
$appRegistration = @{
    TenantId          = "57ed4d2a-bc9c-4fee-87a8-536e35d07ddc"
    ClientId          = "c292639b-b8e1-4233-804c-7de3a1e6d3bd"
    ClientCertificate = Get-Item "Cert:\CurrentUser\My\475da948e4ba44d9b5bc31ab4b8006113fd5f538"
}

$msalToken = Get-msaltoken @appRegistration -ForceRefresh
$vulnerabilities = Import-Csv -Path "C:\Vulnerabilities.csv"

foreach ($vuln in $vulnerabilities) {

    # request body which contains our message
    $requestBody = @{
        "message"         = [PSCustomObject]@{
            "subject"      = "Vulnerability Remediation - $($vuln.Vulnerability)"
            "body"         = [PSCustomObject]@{
                "contentType" = "Text"
                "content"     = "This is a request to remediate vulnerabilities in the organization environment. Below is a table of details regarding the vulnerability.

            Route To: $($vuln.RouteTo)
            Priority: $($vuln.Priority)
            Action Required: $($vuln.Action)
            Technique / Remediation Steps: $($vuln.Technique)
            Asset: $($Vuln.Device)
            User: $($vuln.User)
            Notes: $($vuln.Notes)
            
            "
            }
            "toRecipients" = @(
                [PSCustomObject]@{
                    "emailAddress" = [PSCustomObject]@{
                        "address" = "support@yourorganization.org"
                    }
                }
            )
        }
        "saveToSentItems" = "true"
    }

    # make the graph request
    $request = @{
        "Headers"     = @{Authorization = $msalToken.CreateAuthorizationHeader() }
        "Method"      = "Post"
        "Uri"         = "https://graph.microsoft.com/v1.0/users/username@organization.com/sendMail"
        "Body"        = $requestBody | ConvertTo-Json -Depth 5
        "ContentType" = "application/json"
    }

    Invoke-RestMethod @request
}