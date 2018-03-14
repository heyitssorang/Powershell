# Must include this block or REST API will not authenticate properly
#------------------------------------------------------------------------------------
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
#-------------------------------------------------------------------------------------



function CallJIRAJQL
{
    param( 
        $query
    )
    $objTable = @()

    $server = "url"
    $Method = "GET"
    $userPassword = ConvertTo-SecureString -String "pass" -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential (“user”, $userPassword)

    $objTable = @()
    $counter = 0
    $maxCounter = 1

    While ($counter -lt $maxCounter)
    {
        $StartIndex = $counter
        $MaxResults = 100
        $escapedQuery = [System.Web.HttpUtility]::UrlPathEncode($Query)
        $Uri = "$($server)/rest/api/latest/search?jql=$escapedQuery&validateQuery=true&expand=transitions&startAt=$StartIndex&maxResults=$MaxResults"

        $headers = @{
        'Content-Type' = 'application/json; charset=utf-8'
        }

        if ($Credential)
            {
                Write-Host "[JIRA REST API] Using HTTP Basic authentication with provided credentials for $($Credential.UserName)"
                [String] $Username = $Credential.UserName
                $token = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("${Username}:$($Credential.GetNetworkCredential().Password)"))
                $headers.Add('Authorization', "Basic $token")
                Write-Host "Using HTTP Basic authentication with username $($Credential.UserName)"
            } else {
                Write-Host "[JIRA REST API] Credentials were not provided. Please Provide Credentials"
            }

        $iwrSplat = @{
                Uri             = $Uri
                Headers         = $headers
                Method          = $Method
                UseBasicParsing = $true
        }

        try
            {
                Write-Host "[Invoke-JiraMethod] Invoking JIRA method $Method to URI $URI"
                $webResponse = Invoke-WebRequest @iwrSplat
            } catch {
                # Invoke-WebRequest is hard-coded to throw an exception if the Web request returns a 4xx or 5xx error.
                # This is the best workaround I can find to retrieve the actual results of the request.
                $webResponse = $_.Exception.Response
            }
    
            # Convert the JSON to Powershell Obj
        $content = ConvertFrom-Json -InputObject $webResponse.content
        $maxCounter = $content.total

        # Get Timesheet
        foreach ($i in $content.issues)
            {
                $obj = New-Object -TypeName PSObject -Property @{
                    IssueKey = $i.key
                    AssigneeEmail = $i.fields.assignee.EmailAddress
                    AssigneeName = $i.fields.assignee.DisplayName
                    Status = $i.fields.status.name
                    Component = $i.fields.components.name
                    ReporterEmail = $i.fields.assignee.EmailAddress
                    ReporterName = $i.fields.assignee.DisplayName
                    IssueType = $i.fields.issuetype.name
                    CreatedDate = (Get-Date ($i.fields.created))
                    CreatedMonth = (Get-Date ($i.fields.created)).Month
                    CreatedDay = (Get-Date ($i.fields.created)).Day
                    CreatedYear = (Get-Date ($i.fields.created)).Year
                    UpdatedDate = (Get-Date ($i.fields.updated))
                    UpdatedMonth = (Get-Date ($i.fields.updated)).Month
                    UpdatedDay = (Get-Date ($i.fields.updated)).Day
                    UpdatedYear = (Get-Date ($i.fields.updated)).Year
                    Summary = $i.fields.summary
                    }
                $obj.PSObject.TypeNames.Insert(0, 'Report.JIRATicketInfo')
                $objTable += $obj
            }

        $counter += $MaxResults
    }

    

    return $objTable
}
