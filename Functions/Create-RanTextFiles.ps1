function Create-RanTextFiles {
    param (
        [int]$Count,
        [string]$Path
    )

    $ranN = (0..$Count)

    ForEach ($n in $ranN) {
        Write-Host $n
        $fileName = "$($Path)\file$($n).txt"
        New-Item -Path $fileName
        Set-Content -Path $fileName -Value "$fileName"
    }
}

Create-RanTextFiles -Count 1000 -Path "C:\temp\Test1"
