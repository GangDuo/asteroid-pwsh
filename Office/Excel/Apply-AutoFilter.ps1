<#
.SYNOPSIS

.EXAMPLE
Apply-AutoFilter -Path ".\sample.xlsx" -RangeAsText '$A$1:$C$20' -Field 2 -Criteria2 @(1, "2/6/2024", 1, "3/7/2024")

#>
function Apply-AutoFilter {
    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$RangeAsText,

        [Parameter(Mandatory)]
        [uint32]$Field,

        [Parameter(Mandatory)]
        $Criteria2
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Open($Path)

    $workSheet = $workbook.ActiveSheet

    $range = $workSheet.Range($RangeAsText)

    # https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.range.autofilter?view=excel-pia
    $range.AutoFilter.GetType().FullName
    $range.AutoFilter.Invoke(@($Field, $null, 7, $Criteria2, $true))

    $workbook.Save()
    $workbook.Close($false)
    $excel.Quit()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
