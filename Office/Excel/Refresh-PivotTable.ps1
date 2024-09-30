<#
.SYNOPSIS

.EXAMPLE
Refresh-PivotTable -Path .\sample.xlsx

#>
function Refresh-PivotTable {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Open($Path)

    $workSheet = $workbook.ActiveSheet

    for($i = 1; $i -le $workSheet.PivotTables.Count; $i++) {
	    $pivot = $workSheet.PivotTables.Invoke(@($i))
	    $pivot.RefreshTable.Invoke() | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pivot)
	    $pivot = $null
	}

    $workbook.Save()
    $workbook.Close($false)
    $excel.Quit()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}