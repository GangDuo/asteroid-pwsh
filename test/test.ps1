. ..\Split-By-Line.ps1

$dest = '.\dest'
New-Item $dest -ItemType Directory -ErrorAction SilentlyContinue

Split-By-Line -ChunkSize 3 -Destination $dest -File .\source.csv -IsSkipLine { param($line)
    $xs = $line.split(",")
    if($xs[1] -eq '') {
        return $true
    }
    return $false
}

#Get-Help Split-By-Line