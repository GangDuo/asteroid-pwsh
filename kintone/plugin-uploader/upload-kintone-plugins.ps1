# https://cybozu.dev/ja/kintone/sdk/development-environment/plugin-uploader/

@{
BASE_URL=$env:KINTONE_BASE_URL
USERNAME=$env:KINTONE_USERNAME
PASSWORD=$env:KINTONE_PASSWORD
PROXY=$env:HTTP_PROXY
PLUGINS_HOME=$env:PLUGINS_HOME
}

Get-ChildItem -Path "${env:PLUGINS_HOME}\*.zip" -Recurse -File | `
  %{ kintone-plugin-uploader $_.FullName; Start-Sleep -Seconds 3 }