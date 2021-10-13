$culture = Get-Culture

$2008R2 = "https://support.microsoft.com/en-us/feed/rss/f825ca23-c7d1-aab8-4513-64980e1c3007"
$2012 = "https://support.microsoft.com/en-us/feed/rss/0cfbf2af-24ea-3e18-17e6-02df7331b571"
$2012R2 = "https://support.microsoft.com/en-us/feed/rss/3ec8448d-ebc8-8fc0-e0b7-9e8ef6c79918"
$2016 = "https://support.microsoft.com/en-us/feed/rss/c3a1be8a-50db-47b7-d5eb-259debc3abcc"
$2019 = "https://support.microsoft.com/en-us/feed/rss/eb958e25-cff9-2d06-53ca-f656481bb31f"


$rssUrl = $2019

$feedData = Invoke-WebRequest -Uri $rssUrl -UseBasicParsing -ContentType "application/xml"
$feedData.Content | Out-File -FilePath "$PSScriptRoot\rss.xml"
$feedXMLString = Get-Content -Path "$PSScriptRoot\rss.xml"
   
$feedXML = $feedXMLString[1..$feedXMLString.Length]
$formattedXML = [xml]$feedXML
$feed = $formattedXML.rss.channel

$OSName = $feed.title.Split("-")[1].trim(" ")
$patchList = New-Object -TypeName "System.Collections.ArrayList"

ForEach ($msg in $feed.Item){
    If ($msg.title -like "*Monthly Rollup*" -or $msg.title -like "*OS Build*") {
        $obj = [PSCustomObject]@{
            'Link'          = $msg.link
            'Description'   = $msg.title
            'LastUpdated'   = [datetime]$msg.pubDate
            'OS'            = $OSName
        }
        $patchList.Add($obj)
    }
}

$latestMonthPatchData = Invoke-WebRequest -Uri $patchList[0].Link
$patchFileCSVUrl = $latestMonthPatchData.Links.href -like "*.csv"
Invoke-WebRequest -Uri $patchFileCSVUrl[0] -OutFile "$PSScriptRoot\patchFile.csv"
$patchIssuesHTMLTable = ($latestMonthPatchData.AllElements | Where-Object {$_.tagname -eq 'td'})
$patchSymptomsTable = New-Object -TypeName "System.Collections.ArrayList"
$patchWorkAroundTable = New-Object -TypeName "System.Collections.ArrayList"

for ($i=2; $i -lt $patchIssuesHTMLTable.count; $i++){
    if ($patchIssuesHTMLTable[$i].innerText -notlike "*Release Channel*") {
        if ($i % 2 -eq 0 ) {
            $obj = [PSCustomObject]@{
                'Symptom' = $patchIssuesHTMLTable[$i].innerText
            }
            $patchSymptomsTable.Add($obj)
        }
        else {
            $obj = [PSCustomObject]@{
                'Workaround' = $patchIssuesHTMLTable[$i].innerText
            }
            $patchWorkAroundTable.Add($obj)
        }
    }
    else { break }
}

$patchFileImported = Import-Csv -Path "$PSScriptRoot\patchFile.csv" -Delimiter ","
$patchFileImported = Get-Content -Path "C:\Users\MK10453\OneDrive - Point72 Asset Management, L.P\Documents\vsworkspace\powershell\project\PatchingReport\patchFile.csv"
$patchFileData = $patchFileImported | Select-Object -Skip 2 | ConvertFrom-Csv -Delimiter "," -Header "FileName", "FileVersion", "Date", "Time", "FileSize"
$patchFileTable = $patchFileData | Where-Object -FilterScript {$_.FileVersion -ne "" -and $_.FileVersion -ne "Not versioned" -and $_.Date.ToDateTime($culture) -gt ((Get-Date).AddMonths(-2))}

$UBRTable = New-Object -TypeName "System.Collections.ArrayList"

$patchFileTable | ForEach-Object {
    $obj = [PSCustomObject]@{
        'FileName'  = $_.FileName
        'FileVer'   = $_.FileVersion
        'Date'      = $_.Date.ToDateTime($culture)
        'BuildNum'  = $_.FileVersion.Split(".")[-2].Trim(" ")
        'UBR'       = $_.FileVersion.Split(".")[-1].Trim(" ")
    }
    $UBRTable.Add($obj)
}