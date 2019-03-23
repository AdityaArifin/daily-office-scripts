Set excel = GetObject(,"Excel.Application")
Set workbook = excel.activeWorkbook    
for each worksheet in workbook.worksheets
    worksheet.Activate
    worksheet.Range("A1").Activate
    excel.activeWindow.zoom = 75
next
workbook.worksheets(1).Activate