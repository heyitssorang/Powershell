$2008R2 = "https://support.microsoft.com/en-us/feed/rss/f825ca23-c7d1-aab8-4513-64980e1c3007"
$2012 = "https://support.microsoft.com/en-us/feed/rss/0cfbf2af-24ea-3e18-17e6-02df7331b571"
$2012R2 = "https://support.microsoft.com/en-us/feed/rss/3ec8448d-ebc8-8fc0-e0b7-9e8ef6c79918"
$2016 = "https://support.microsoft.com/en-us/feed/rss/c3a1be8a-50db-47b7-d5eb-259debc3abcc"
$2019 = "https://support.microsoft.com/en-us/feed/rss/eb958e25-cff9-2d06-53ca-f656481bb31f"

function Parse-MSFTRSSXml {
    param (
        $Uri
    )

    $feedData = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ContentType "application/xml"
    $feedData.Content | Out-File -FilePath "$PSScriptRoot\rss.xml"
    $feedXMLString = Get-Content -Path "$PSScriptRoot\rss.xml"

    $feedXML = $feedXMLString[1..$feedXMLString.Length]
    $formattedXML = [xml]$feedXML
    $feed = $formattedXML.rss.channel

    $OSName = $feed.title.Split("-")[1].trim(" ")
    $patchList = New-Object -TypeName "System.Collections.ArrayList"

    ForEach ($msg in $feed.Item){
        If ($msg.title -like "*Monthly Rollup*" -or $msg.title -like "*OS Build*" -and $msg.title -notmatch "Windows Server base*|Preview$|Expired$") {
            $obj = [PSCustomObject]@{
                'Link'          = $msg.link
                'Description'   = $msg.title
                'LastUpdated'   = [datetime]$msg.pubDate
                'OS'            = $OSName
            }
            $patchList += $obj
        }
    }
    return $patchList
}

#($msg.title -like "*Monthly Rollup*" -or $msg.title -like "*OS Build*" -and $msg.title -notlike "Windows Server base*")
#$tempDate = [datetime]::parseexact($msg.title.Split('-KB')[0].Trim('—'), 'MMMM d, yyyy', $null)

$MRList_2019 = Parse-MSFTRSSXml -Uri $2008R2