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


# Test Uri
#$tempoUri = 'https://jiraurl/rest/tempo-timesheets/3/worklogs?dateFrom=2017-11-26&dateTo=2017-12-10&teamId=4'

$server = "https://touton.atlassian.net"
$Method = "GET"
#$Credential = Get-Credential
$userPassword = ConvertTo-SecureString -String "pass" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential (“user”, $userPassword)

# Get Current Date of the Week
$s = (Get-Date -hour 0 -minute 0 -second 0).AddDays(-7)
$sd = $s.AddDays(-($s).DayOfWeek.value__)
$ed = $s.AddDays(7-($s.AddSeconds(86399)).DayOfWeek.value__)   
$dateFrom = Get-Date $sd -Format yyyy-M-dd
$dateTo = Get-Date $ed -Format yyyy-M-dd

# REST API formatted
$Uri = "$($server)/rest/tempo-timesheets/3/worklogs?dateFrom=$($dateFrom)&dateTo=$($dateTo)&teamId=4"

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
        #Uri             = $TempoUri
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
$content = ConvertFrom-Json -InputObject $webResponse.Content

# Get Timesheet
$objTable = @()
foreach ($i in $content)
    {
        $obj = New-Object -TypeName PSObject -Property @{
            MemberName = $i.author.DisplayName
            MemberEmail = $i.author.name
            TimeSpent = ($i.timeSpentSeconds / 60)
            WorkStarted = Get-Date ($i.dateStarted)
            WorkComment = $i.comment
            IssueKey = $i.issue.key
            IssueSummary = $i.issue.summary
            DateMonth = (Get-Date ($i.dateStarted)).Month
            DateDay = (Get-Date ($i.dateStarted)).Day
            DateYear = (Get-Date ($i.dateStarted)).Year
            TimeHour = ((Get-Date).Date.AddHours(($i.timeSpentSeconds / 60) / 60)).Hour
            TimeMinute = ((Get-Date).Date.AddHours(($i.timeSpentSeconds / 60) / 60)).Minute
            }
        $obj.PSObject.TypeNames.Insert(0, 'Report.JIRATimeSheet')
        $objTable += $obj
    }

Write-Output $objTable | ft
