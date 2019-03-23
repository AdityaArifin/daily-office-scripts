Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = objFSO.GetParentFolderName(Wscript.ScriptFullName)

for each file in objFSO.GetFolder(strFolder).Files
    processFile(file)
next

Sub processFile(ByRef file)
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    fileExtension = objFSO.GetExtensionName(file)
    if(fileExtension <> "xls" and fileExtension <> "xlsx" ) then 
        exit sub
    end if
    
    Set excel = CreateObject("Excel.Application")
    Set workbook = excel.workbooks.open(file)
    if (workbook.ReadOnly = True) then
        msgbox "Workbook " & file & "is read only/open"
        err.raise 1
    end if
    
    for each worksheet in workbook.worksheets
        worksheet.Activate
        worksheet.Range("A1").Activate
        excel.activeWindow.zoom = 75
    next
    workbook.worksheets(1).Activate
    workbook.save
    workbook.close
End Sub