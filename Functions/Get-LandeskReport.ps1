# Declaring File/Folder Locations
$ActiveFile = "C:\temp\Active.csv"		
$FailedFile = "C:\temp\Failed.csv"
$PendingFile = "C:\temp\Pending.csv"
$SuccessfulFile = "C:\temp\Successful.csv"
$FailedCodeTableFile = "C:\temp\FailedCodeTable.csv"
		
# Getting Fridays Job Run Date		
$a = Get-Date		
while ($a.DayOfWeek -ne "Friday") {$a = $a.AddDays(-1)}
$PRDate = $a.ToString("dd/MM/yyyy")
		
# Importing Required Data
$Active = Import-Csv -Path $ActiveFile
$Pending = Import-Csv -Path $PendingFile
$Failed = Import-Csv -Path $FailedFile
$Success = Import-Csv -Path $SuccessfulFile
$FailedCodeTable = Import-Csv -Path $FailedCodeTableFile

# Calculating Counts For Output		
$OutstandingCount = $Active.Name.Count + $Failed.Name.Count + $Pending.Name.Count
$TotalCount = $OutstandingCount + $Success.Name.Count		
		
#Create Failed Result Details Table	
$FailedTable = New-Object 'System.Collections.Generic.List[System.Object]'
ForEach ($Fail in $Failed) {
    $FailedMsg = $FailedCodeTable.Where({$PSItem.Code -eq $Fail.'Return code'})
    $obj = [PSCustomObject]@{
        'Server Name'	= $Fail.Name
        'Code Msg'	    = $FailedMsg.Msg
        'Detailed Msg'	= $FailedMsg.Detail
      }
    $FailedTable.Add($obj)
}

$title = "<h1>Landesk Report</h1>"

$text1 =
"
<h3>
Job Run Date           - $($PRDate)<br>
Completed Successfully - $($Success.Name.Count)<br>
Outstanding            - $OutstandingCount<br>
Total Servers          - $TotalCount<br>
</h3>
"

$FailedTable = $FailedTable | ConvertTo-Html -Fragment -PreContent "<h2>Failed Results</h2>"

#CSS codes
$header = @"
<style>
    h1 {
        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;
    }

    h2 {
        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;
    }

   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}

    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

    #CreationDate {
        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }
</style>
"@

$Report = ConvertTo-HTML -Body "$title $text1 $FailedTable" -Head $header -Title "Landesk Report" -PostContent "<p id='CreationDate'>Creation Date: $(Get-Date)</p>"
$Report | Out-File .\Report.html