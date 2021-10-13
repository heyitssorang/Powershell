$Set1 = '1','2','4'

$Set2 = '1','2','3','4'

Compare-Object $set1 $set2 | Where-Object -FilterScript {$_.SideIndicator -like "=>"}
