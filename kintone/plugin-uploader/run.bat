@echo off

rem end without '\'.
rem --------------------------------
set PLUGINS_HOME=<required>

rem https://github.com/kintone/js-sdk/blob/main/packages/plugin-uploader/README.md
rem --------------------------------
set SUBDOMAIN=<required>
set KINTONE_USERNAME=<required>
set KINTONE_PASSWORD=<required>

rem Comments Off if enabled.
rem --------------------------------
rem set KINTONE_BASIC_AUTH_USERNAME=
rem set KINTONE_BASIC_AUTH_PASSWORD=

rem HTTPS_PROXY or HTTP_PROXY
rem --------------------------------
rem set HTTPS_PROXY=
rem set HTTP_PROXY=

set KINTONE_BASE_URL=https://%SUBDOMAIN%.cybozu.com

pushd %~dp0
rem Powershell -ExecutionPolicy Bypass -File .\upload-kintone-plugins.ps1

pause