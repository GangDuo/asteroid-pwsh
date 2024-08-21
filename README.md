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
