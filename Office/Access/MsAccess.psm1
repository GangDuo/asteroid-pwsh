using namespace  System.Data
using namespace  System.Data.OleDb

function New-MsAccessConnection {
    [OutputType([System.Data.OleDb.OleDbConnection])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }#,
#            ErrorMessage = ": {0} はAccessファイルではありません。"
        )]
        [string]
        $acPath
    )
    # 接続文字列作成
    $builder = [OleDbConnectionStringBuilder]::new()

    # インストールされている最新のMicrosoft.ACE.OLEDBプロバイダ名を取得
    $builder["Provider"] = ([System.Data.OleDb.OleDbEnumerator]::new().GetElements().SOURCES_NAME -match "^Microsoft.ACE.OLEDB." )[-1]
    # 対象Accessデータベース
    $builder["Data Source"] = $acPath
    # システムデータベース(指定することでMSysRelationshipsなどにクエリ可能になる)
    $builder["Jet OLEDB:System database"] = "$env:APPDATA\Microsoft\Access\System.mdw"

    return [OleDbConnection]::new($builder.ConnectionString)
}

<#
.SYNOPSIS
実行するpowershell.exeは、インストールされているOfficeと同じビット数でなければならない。
◎PowerShell実行ファイルの場所
▼32 ビット
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ep bypass

▼64bit
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

.EXAMPLE
$dataSet = Get-MsAccessERInfo ".\sample.accdb"
$dataSet.Tables["Tables"] | ForEach-Object { $_ | FL }

#>
function Get-MsAccessERInfo {
    [OutputType([System.Data.DataSet])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }
        )]
        [string]
        $acPath
    )
    try {
        # セットアップ
        [OleDbConnection]$connection = New-MsAccessConnection $acPath
        $connection.Open()
        [OleDbTransaction]$Transaction = $connection.BeginTransaction()

        $adapter = [OleDbDataAdapter]::new(
            [OleDbCommand]::new("Select * From MSysRelationships", $connection, $Transaction)
        )

        # テーブルデータ取得
        $dataSet = [DataSet]::new()
        $adapter.Fill($dataSet, "MSysRelationships") > $null
        $dataSet.Tables.Add($connection.GetSchema("Indexes"))  > $null
        $dataSet.Tables.Add($connection.GetSchema("Columns"))  > $null
        $dataSet.Tables.Add($connection.GetSchema("Tables"))  > $null

        $Transaction.Commit()
    } catch {
        if ($Transaction) {
            $Transaction.Rollback()
        }
        $PSCmdlet.ThrowTerminatingError($_)
    } finally {
        if ($Transaction) {
            $Transaction.Dispose()
        }

        if ($connection) {
            $connection.Close()
            $connection.Dispose()
        }  
    }

    # 返り値はDataset
    return [DataSet]$dataSet
}

<#
.SYNOPSIS

.EXAMPLE
using namespace  System.Data.OleDb

Use-OleDbConnection -ScriptBlock {param([OleDbConnection]$connection, $Options)
    Write-Host $Options.prop
    $cmd = [OleDbCommand]::new("DELETE FROM table_name", $connection)
    $cmd.ExecuteNonQuery()
} -File .\sample.accdb -Options @{ prop = 1}

#>
function Use-OleDbConnection {
    param($ScriptBlock, $File, $Options)

    try {
        [OleDbConnection]$connection = New-MsAccessConnection $File
        $connection.Open()

        Invoke-Command $ScriptBlock -ArgumentList $connection, $Options
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    } finally {
        if ($connection) {
            $connection.Close()
            $connection.Dispose()
        }  
    }
}