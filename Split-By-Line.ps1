<#
.SYNOPSIS
    テキストファイルを行単位で分割する。
.DESCRIPTION
    分割後のファイルに含めない行を指定できる。
#>
function Split-By-Line {
    param(
        [PSDefaultValue(Help = '100')]
        [UInt64]$ChunkSize = 100,

        [PSDefaultValue(Help = 'splitFile_')]
        [string]$Prefix = 'splitFile_',

        [PSDefaultValue(Help = '.csv')]
        [string]$Extension = '.csv',        

        [Parameter(Mandatory)]
        [string]$Destination,

        [Parameter(Mandatory)]
        [string]$File,

        [System.Management.Automation.ScriptBlock]$IsSkipLine = {$false}
    )

    begin {
        # ファイルから先頭行を読み取る
        $header = Get-Content -Path $File | Select-Object -First 1

        # ファイルを読み込み、先頭行を除外して配列に格納
        $content = Get-Content -Path $File | Select-Object -Skip 1

        $index = 0
        $buf = New-Object System.Collections.ArrayList
        $cursor = -1
        $currentCursor = 0

        $chunks = New-Object System.Collections.ArrayList

        # n行ずつ分割して新しいファイルに保存
        foreach ($line in $content) {
            if(Invoke-Command -ScriptBlock $IsSkipLine -ArgumentList $line) {
                continue
            }

            $currentCursor = [math]::Ceiling(($index + 1) / $ChunkSize)
            if($currentCursor -ne $cursor) {
                if($cursor -ne -1) {
                    # 上限に達したので分割レコードをキャッシュに保存
                    $chunks.Add($buf) > $null
                }

                $buf = New-Object System.Collections.ArrayList
                $buf.Add($header) > $null
            }
            $buf.Add($line) > $null

            $index++
            $cursor = $currentCursor
        }
        $chunks.Add($buf) > $null

        $i = 0
        foreach ($chunk in $chunks) {
            Add-Content (Join-Path $Destination ($Prefix + (++$i) + ${Extension})) $chunk
        }
    }

    process {}

    end {
        $chunks.Count
    }
}
