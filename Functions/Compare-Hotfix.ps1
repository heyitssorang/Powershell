function Compare-Hotfix {
    param (
        $ServerName = 'localhost',
        $HotfixList
    )

    $Installed = Get-HotFix -Id $HotfixList -ComputerName $ServerName
    $Missing = Compare-Object -ReferenceObject $HotfixList -DifferenceObject $Installed.HotfixId -PassThru

    $HotfixResult = New-Object 'System.Collections.Generic.List[System.Object]'
    If ($Installed.count -ne $null) {
        ForEach ($Hotfix in $Installed) {
            $obj = [PSCustomObject]@{
                'Server Name'	= $Hotfix.Source
                'HotfixID'	    = $Hotfix.Msg
                'Installed On'	= $Hotfix.InstalledOn
                'State'         = "Installed"
            }
            $HotfixResult.Add($obj)
        }
    }
    Else {
        $obj = [PSCustomObject]@{
            'Server Name'	= $Installed.Source
            'HotfixID'	    = $Installed.Msg
            'Installed On'	= $Installed.InstalledOn
            'State'         = "Installed"
        }
    }

}