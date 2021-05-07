function Get-DuplicateFiles {
    param (
        $SourchPath,
        $DestinationPath
    )
    
    #Initialize Stopwatch
    $Stopwatch = [system.diagnostics.stopwatch]::StartNew()

    #Initialize Window Forms and VisualBasic Assembly for Popups
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    #---

    $sourceFileList = Get-ChildItem -Path $SourchPath -File -Recurse | Get-FileHash -Algorithm SHA1
    $destFileList = Get-ChildItem -Path $DestinationPath -File -Recurse | Get-FileHash -Algorithm SHA1

    $duplicates = Compare-Object -ReferenceObject $destFileList -DifferenceObject $sourceFileList -Property Hash -IncludeEqual -ExcludeDifferent -PassThru | Select-Object -Property Hash, Path

    #---

    #Stop Stopwatch
    $Stopwatch.Stop()
    $Message = "Script has finished.`nRun Time: $($Stopwatch.Elapsed.TotalSeconds) Seconds"
    [System.Windows.Forms.MessageBox]::Show("$Message" , "Script Finished" , "Ok")

    return $duplicates
}

Get-DuplicateFiles -SourchPath "C:\temp\Test1" -DestinationPath "C:\temp\Test2"