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
        If ($msg.title -like "*Monthly Rollup*" -or $msg.title -like "*OS Build*" -and $msg.title -notlike "Windows Server base*") {
            $tempDate = [datetime]::parseexact($msg.title.Split('-KB')[0].Trim('—'), 'MMMM d, yyyy', $null)
            $obj = [PSCustomObject]@{
                'Link'          = $msg.link
                'Description'   = $msg.title
                'LastUpdated'   = [datetime]$msg.pubDate
                'PatchDate'     = $tempDate
                'OS'            = $OSName
            }
            $patchList += $obj
        }
    }
    $patchList = $patchList | Sort-Object -Property PatchDate
    return $patchList
}

$MRList_2019 = Parse-MSFTRSSXml -Uri $2019