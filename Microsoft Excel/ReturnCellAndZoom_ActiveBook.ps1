$excel = New-Object -ComObject Excel.Application
$workbook = $excel.ActiveWorkbook
Write-Output $workbook.Name()
foreach ($worksheet in $workbook.worksheets) {
    $worksheet.Activate()
    $worksheet.Range("A1").Activate()
    $excel.activeWindow.zoom = 75
    
}
$workbook.worksheets(1).Activate()