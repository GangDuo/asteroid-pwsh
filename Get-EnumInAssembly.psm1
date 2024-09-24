<#
.Description
アセンブリ内のEnumを取得する。

.OUTPUTS
System.Collections.Generic.Dictionary[string, enum]

.EXAMPLE
$enumDic = Get-EnumInAssembly
$enumDic.msoControlOLEUsageNeither
#>
function Get-EnumInAssembly {
    # Enum値を格納するDictionaryを生成
    $enumDic = [System.Collections.Generic.Dictionary[string, enum]]::new()

    # パス指定でアセンブリを読み込む
    # PassThruオプションで返り値が読み込んだアセンブリ内の型になる
    Add-Type -Path @(
        "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
        "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
        "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Access.Dao\*\Microsoft.Office.Interop.Access.Dao.dll"
        "C:\Windows\assembly\GAC\ADODB\*\ADODB.dll"
        "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Access\*\Microsoft.Office.Interop.Access.dll"
        # IsEnumでフィルタリングしてからEnum値を取得
    ) -PassThru | Where-Object IsEnum | ForEach-Object GetEnumValues | ForEach-Object {
        # 値をセット(重複値は上書き)
        $enumDic[$_] = $_
    }
    return $enumDic
}