Attribute VB_Name = "Exports_handling"
Function last()
'last = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(Rows.Count, "I").End(xlUp).Row
last = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).Row
End Function
Function first()
first = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Row
End Function
Sub HandleExport(ByVal i As Integer, sheetname As String, exportname As String)
Dim lastrow, firstrow As Long
Dim myPath, FolderPath, FilePath, FileName As String
Dim oFSO As Object
Dim oFolder, oFile As Object
Dim wb_exp As Object
Dim wb As Object
Dim xlApp As Excel.Application
Dim xlInstances As Excel.Application

Application.DisplayAlerts = False
Application.CutCopyMode = True

'Get Folder Dynamically
FolderPath = GetLocalPath(ThisWorkbook.path)

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(FolderPath & "\SAP_Exports")

FileName = exportname & i & ".xlsx"


On Error Resume Next
Set xlInstances = GetObject(, "Excel.Application")

For Each xlApp In xlInstances
    For Each wb In xlApp.Workbooks
        If wb.Name = FileName Then
            Debug.Print FileName
            Set wb_exp = Workbooks(FileName)
            
            wb_exp.Activate
            wb_exp.Sheets(1).Activate
            'Remove Yellow rows
            ActiveSheet.Range("A1:S1").AutoFilter Field:=1, Criteria1:=RGB(255, _
                255, 0), Operator:=xlFilterCellColor
                
            lastrow = last
            firstrow = first
            
            ActiveSheet.Range("A" & firstrow & ":A" & lastrow).EntireRow.Delete
            ActiveSheet.ShowAllData
            'Remove First row
            ActiveSheet.Rows(1).EntireRow.Delete
            'Copy whats left and paste
            ActiveSheet.Range("A1").CurrentRegion.Select
            Selection.Copy
            ThisWorkbook.Sheets(sheetname).Activate
            lastrow = last
            ThisWorkbook.Sheets(sheetname).Range("A" & lastrow + 1).PasteSpecial xlPasteAll
            wb_exp.Close SaveChanges:=False
            'Move File after Processing
            oFSO.MoveFile FolderPath & "\SAP_Exports\" & FileName, FolderPath & "\Processed\" & FileName
        End If
    Next wb
Next xlApp

Application.CutCopyMode = False

End Sub
