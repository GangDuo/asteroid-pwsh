# asteroid-pwsh
PowerShellの便利モジュール群

## Set-ClipboardWithImage

画像をクリップボードに設定します。

### 使用例

```powershell
$PSScriptRoot = ''
Import-Module "$($PSScriptRoot)\Clipboard"

Set-ClipboardWithImage -Path [filepath]
```

# Split-By-Line

## 例

### 例 1: テキストファイルを999行単位で複数ファイルへ分割する

LineNumbers.csv は先頭行にヘッダありとします。
LineNumbers.csv を読み込み1ファイル1000行（ヘッダを含む）の複数ファイルへ分割し、dest フォルダへ保存します。

```powershell
Split-By-Line -ChunkSize 999 -Destination .\dest -File .\LineNumbers.csv
```

### 例 2: テキストファイル2列目の値が空欄のレコードを除外して、100行単位で複数ファイルへ分割する

IsSkipLine パラメーターは引数にテキストファイル行を受け取ります。
返り値が $true なら分割後のファイルにその行は含まれません。
$false なら分割後のファイルにその行を含みます。

```powershell
Split-By-Line -Destination .\dest -File .\LineNumbers.csv -IsSkipLine { param($line)
    $clms = $line.split(",")
    if($clms[1] -eq '') {
        return $true
    }
    return $false
}
```
