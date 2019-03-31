Set excel = GetObject(,"Excel.Application")
Set workbook = excel.activeWorkbook
firstVisibleIndex = -1
on error resume next    
for each worksheet in workbook.worksheets:do
    if (worksheet.visible = 0) then exit do
    worksheet.Activate
    worksheet.Range("A1").Activate
    excel.activeWindow.zoom = 75
    if( firstVisibleIndex = -1 ) then firstVisibleIndex = worksheet.Index
loop while false: next
workbook.worksheets(firstVisibleIndex).Activate