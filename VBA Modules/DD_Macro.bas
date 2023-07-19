Attribute VB_Name = "Module3"
Sub New_macro()

Dim activerange As Range
Dim name, erp_name, goldID, location As String
Dim colName, valCell, startdate, comment As String
Dim firstrowoutput, lastrow, iteration, firstrow, lastrowmain, nextrowoutput As Long
Dim i As Integer
Dim Col As New Collection
Dim nameArray(), filenames() As Variant
Dim FolderPath, FilePath, myPath, LastCol As String
Dim oFSO, oFolder, oFile As Object

FolderPath = GetLocalPath(ActiveWorkbook.path)

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(FolderPath & "\Input")

For Each oFile In oFolder.Files

    Workbooks.Open oFile

Dim main_ws As Worksheet: Set main_ws = Workbooks(oFSO.GetFileName(oFile)).Worksheets("1.survey")
Dim output_ws As Worksheet: Set output_ws = ThisWorkbook.Worksheets("Output")



main_ws.Activate

On Error Resume Next
'Pick Up Name

LastCol = Cells(2, ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column).Address(False, False)


iteration = Application.CountIfs(Range("A2:" & LastCol), "*Name*", Range("A2:" & LastCol), "<>*Entity*")


lastrowmain = main_ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim totalhourscol As String

totalhourscol = "F"

colName = GetColumnName(4, 0, "Comment")
startdate = GetColumnName(4, 0, "Start Date")
comment = GetColumnName(4, 1, "Comment")

Dim diff As Integer
diff = Columns(GetColumnName(4, 0, "Comment")).Column - Columns("E").Column - 1  ' find length of table to copy

'' Following code used to figure out if the number of iterations should be used in case no name was written
x = iteration
For i = 1 To iteration
If i > 1 Then
    totalhourscol = ColumnNumberToLetter(Columns(totalhourscol).Column + diff)
    If IsEmpty(Range(GetColumnNamebyrange(Range(totalhourscol & "2:" & ColumnNumberToLetter(Columns(GetColumnName(4, 0, "Comment")).Column + (diff * i)) & "2"), 0, "Name:") & "2")) = "" _
    And WorksheetFunction.CountA(totalhourscol & "5:" & totalhourscol & lastrowmain) = 0 Then
        x = x - 1
    End If
End If

Next i


iteration = x
'------------------------------------------------------------------------------------------------

totalhourscol = "F"
For i = 1 To iteration
    
With main_ws.Range(firstcol & "2:" & colName & "2")
    
    Set activerange = main_ws.Range(totalhourscol & "2:" & colName & "2")
    name = FindAstring(activerange, "Name:", 1)
    erp_name = FindAstring(activerange, "ERP:", 1)
    goldID = FindAstring(activerange, "GoldID", 1)
    location = FindAstring(activerange, "Location", 1)
    
    
    


    firstrow = main_ws.UsedRange.Offset(4, 0).SpecialCells(xlCellTypeVisible).Row
    firstrowoutput = output_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
    lastrowmain = main_ws.Cells(Rows.Count, 1).End(xlUp).Row
    main_ws.Range("A" & firstrow & ":E" & lastrowmain).SpecialCells(xlCellTypeVisible).Copy Destination:=output_ws.Range("A" & firstrowoutput)
    main_ws.Range(totalhourscol & firstrow & ":" & totalhourscol & lastrowmain).SpecialCells(xlCellTypeVisible).Copy Destination:=output_ws.Range("F" & firstrowoutput)
    lastrow = output_ws.Cells(Rows.Count, 1).End(xlUp).Row
    output_ws.Range("V" & firstrowoutput & ":V" & lastrow) = name
    output_ws.Range("Y" & firstrowoutput & ":Y" & lastrow) = erp_name
    output_ws.Range("Z" & firstrowoutput & ":Z" & lastrow) = goldID
    output_ws.Range("AA" & firstrowoutput & ":AA" & lastrow) = location
    main_ws.Range(startdate & firstrow & ":" & comment & lastrowmain).SpecialCells(xlCellTypeVisible).Copy Destination:=output_ws.Range("J" & firstrowoutput)
    nextrowoutput = output_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

    

    If FindAstring(output_ws.Range("A" & lastrow), "Total", 0) <> "" Then
    output_ws.Rows(lastrow).EntireRow.Delete
    End If

End With
        
        If iteration > 1 Then
        
            totalhourscol = ColumnNumberToLetter(Columns(comment).Column + diff * (i))
            
            colName = GetColumnName(4, 1 + diff, "Comment")
            
        
            startdate = ColumnNumberToLetter(Range(GetColumnName(4, 0, "Start Date") & 4).Column + diff)
            comment = ColumnNumberToLetter(Range(GetColumnName(4, 0, "Comment") & 4).Column + diff)
        End If
    
Next i

    
    output_ws.Range("AB" & firstrowoutput & ":AB" & output_ws.Cells(Rows.Count, 1).End(xlUp).Row) = oFSO.GetFileName(oFile)
    Workbooks(oFSO.GetFileName(oFile)).Close SaveChanges:=False
    oFSO.MoveFile FolderPath & "\Input\" & oFSO.GetFileName(oFile), FolderPath & "\Processed\" & oFSO.GetFileName(oFile)
Next oFile
    output_ws.Activate
    output_ws.Range("G2:G" & output_ws.Cells(Rows.Count, 1).End(xlUp).Row).Formula = "=F2/480"
    output_ws.Range("G2:G" & output_ws.Cells(Rows.Count, 1).End(xlUp).Row).Style = "Comma"
    output_ws.Range("G2:G" & output_ws.Cells(Rows.Count, 1).End(xlUp).Row).NumberFormat = "_(* #,##0.000_);_(* (#,##0.000);_(* ""-""??_);_(@_)"
    Debug.Print "Finished."
End Sub




Sub create_pivot()

Sheets("FTE Breakdown").Activate

Set pt = Sheets("Output").PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=Sheets("Output").Range("A1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("FTE Breakdown").Range("A1"), _
TableName:="PivotTable1")


With Sheets("FTE Breakdown").PivotTables("PivotTable1").PivotFields("Team")
    .Orientation = xlRowField
    .position = 1
End With
Sheets("FTE Breakdown").PivotTables("PivotTable1").AddDataField Sheets("FTE Breakdown").PivotTables( _
        "PivotTable1").PivotFields("FTE"), "Sum of FTE", xlSum


Set pt = Sheets("Output").PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=Sheets("Output").Range("F1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("FTE Breakdown").Range("F1"), _
TableName:="PivotTable2")
With Sheets("FTE Breakdown").PivotTables("PivotTable2").PivotFields("Team")
        .Orientation = xlRowField
        .position = 1
End With
Sheets("FTE Breakdown").PivotTables("PivotTable2").AddDataField Sheets("FTE Breakdown").PivotTables( _
        "PivotTable2").PivotFields("FTE"), "Sum of FTE", xlSum
With Sheets("FTE Breakdown").PivotTables("PivotTable2").PivotFields("Activities/Recon")
        .Orientation = xlColumnField
        .position = 1
End With


Range("B3:B5,G3:I5").NumberFormat = "0.00"

Application.CutCopyMode = False
End Sub

