Attribute VB_Name = "Module1"
Sub pignone()

Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws_gl As Worksheet: Set ws_gl = wb.Sheets("GL report")
Dim ws_raw As Worksheet: Set ws_raw = wb.Sheets("RAW")
Dim ws_ich As Worksheet: Set ws_ich = wb.Sheets("ICH")
Dim LastRow, LastCol, LR_Reversal As Long
Dim exists As Boolean
Dim Found As Range
Dim colName, f, conc1Col, conc2Col As String
Dim pt, PvtTbl As PivotTable
Dim PvtItm      As PivotItem

Application.ScreenUpdating = False
Application.EnableEvents = False


MsgBox "Beginning to Execute..."
ws_gl.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Create Column Target Account and format

Range("V:V").EntireColumn.Insert
Range("V1").Select
Selection.Value = "Target Account"
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 255
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Vlookup and fill column

Range("V2").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(MID(RC[-1],8,9),'Account list'!C[-21],1,0)"
Range("V2").Select
Selection.AutoFill Destination:=Range("V2:V" & LastRow)

'Filter
ws_gl.Range("$A$1:$AI$" & LastRow).AutoFilter Field:=22, Criteria1:="<>#N/A"

'Creating Helper Sheet

Worksheets.Add.Name = "Helper"

Timeout (1)

ws_gl.Range("O1:O" & LastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("Helper").Range("A1:A" & LastRow).PasteSpecial

ws_gl.Range("V1:V" & LastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("Helper").Range("B1:B" & LastRow).PasteSpecial Paste:=xlPasteValues

Sheets("Helper").Activate

Sheets("Helper").Range("B1").Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 255
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

ws_gl.Activate
ws_gl.ShowAllData

Timeout (1)

'Create Column Target invoice and format

Range("W:W").EntireColumn.Insert
Range("W1").Select
Selection.Value = "Target invoice"
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 255
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Vlookup and fill column

Range("W2").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],'Helper'!C[-22],1,0)"
Range("W2").Select
Selection.AutoFill Destination:=Range("W2:W" & LastRow)



'Creating PivotTable Reversal


Worksheets.Add.Name = "Reversal"

Timeout (1)
Set pt = Sheets("GL report").PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=Sheets("GL report").Range("A1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("Reversal").Range("A1"), _
TableName:="PivotTable1")
Sheets("Reversal").Activate
With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ERP Invoice ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Currency")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Flexfield")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Entered DR"), "Sum of Entered DR", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Entered CR"), "Sum of Entered CR", xlSum
    With ActiveSheet.PivotTables("PivotTable1").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Status").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Period Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Source Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Category Name").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Batch Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Batch Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Journal Number").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Journal Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Journal Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Ledger ID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Ledger Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Org Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Type").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ERP Invoice ID").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Number").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Date").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Currency"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Functional Currency"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Class"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Flexfield"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Account").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Entered DR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Entered CR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Total Entered Amount"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounted DR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounted CR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Total Accounted Amount"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Currency Conversion Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Currency Conversion Type"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Currency Conversion Rate"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Line Item Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("BL Journal ID").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Je_header_ID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice")
        .PivotItems("#N/A").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice"). _
        EnableMultiplePageItems = True



'Amount Consistency Column

Sheets("Reversal").Range("F5").Value = "Amount consistency"

Sheets("Reversal").Range("F:F").Columns.AutoFit
LR_Reversal = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Reversal").Range("F5").Select
ActiveCell.Formula = "=SUMIF(A:A,A5,D:D)=SUMIF(A:A,A5,E:E)"
Selection.AutoFill Destination:=Range("F5:F" & LR_Reversal - 1)

'Entry Cr DR columns
Sheets("Reversal").Range("G4").Value = "Entry dr"

Sheets("Reversal").Range("G:G").Columns.AutoFit
LR_Reversal = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Reversal").Range("G5").Select
ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,"""",RC[-2])"
Selection.AutoFill Destination:=Range("G5:G" & LR_Reversal - 1)

Sheets("Reversal").Range("H4").Value = "Entry Cr"

Sheets("Reversal").Range("H:H").Columns.AutoFit
LR_Reversal = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Reversal").Range("H5").Select
ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,"""",RC[-4])"
Selection.AutoFill Destination:=Range("H5:H" & LR_Reversal - 1)

'Account Breakdown
Sheets("Reversal").Range("I4").Value = "Account breakdown"

Sheets("Reversal").Range("I5").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-6],6)"
Selection.AutoFill Destination:=Range("I5:I" & LR_Reversal - 1)

Sheets("Reversal").Range("J5").Select
ActiveCell.Formula = "=MID($C5,8,9)"
Selection.AutoFill Destination:=Range("J5:J" & LR_Reversal - 1)

Sheets("Reversal").Range("K5").Select
ActiveCell.Formula = "=MID($C5,28,1)"
Selection.AutoFill Destination:=Range("K5:K" & LR_Reversal - 1)

Sheets("Reversal").Range("L5").Select
ActiveCell.Formula = "=MID($C5,30,1)"
Selection.AutoFill Destination:=Range("L5:L" & LR_Reversal - 1)

Sheets("Reversal").Range("M5").Select
ActiveCell.Formula = "=MID($C5,22,5)"
Selection.AutoFill Destination:=Range("M5:M" & LR_Reversal - 1)

Sheets("Reversal").Range("N5").Select
ActiveCell.Formula = "=MID($C5,28,1)"
Selection.AutoFill Destination:=Range("N5:N" & LR_Reversal - 1)

Sheets("Reversal").Range("O5").Select
ActiveCell.Formula = "=MID($C5,30,1)"
Selection.AutoFill Destination:=Range("O5:O" & LR_Reversal - 1)

Sheets("Reversal").Range("P5").Select
ActiveCell.Formula = "=MID($C5,32,1)"
Selection.AutoFill Destination:=Range("P5:P" & LR_Reversal - 1)


Range(Range("D5:E5"), Range("D5:E5").End(xlDown)).Select
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "0.000"
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "#,##0.00"

Range(Range("G5:H5"), Range("G5:H5").End(xlDown)).Select
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "0.000"
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "#,##0.00"

'RAW Sheet

ws_raw.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Find Column Name that includes ERP Invoice ID as header to avoid different table layout errors

Set Found = Rows(1).Find(what:="ERP INVOICE ID (Journal Lines)", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column + 1)
conc1Col = ColumnNumberToLetter(Found.Column)
'Create Column Target Invoice

Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Target Invoice"
'Vlookup and fill column

Range(colName & "2").Select
ActiveCell.Formula = "=VLOOKUP(" & conc1Col & "2,'Helper'!A:B,1,0)"
Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

'ICH SHEET


ws_ich.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Find Column Name that includes ERP_INVOICE_ID as header to avoid different table layout errors

Set Found = Rows(1).Find(what:="ERP_INVOICE_ID", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column + 1)
conc1Col = ColumnNumberToLetter(Found.Column)

'Create Column Target Invoice

Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Target Invoice"
'Vlookup and fill column

Range(colName & "2").Select
ActiveCell.Formula = "=VLOOKUP(" & conc1Col & "2,'Helper'!A:B,1,0)"
Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

'Concatenate


Set Found = Rows(1).Find(what:="Seller_UEI", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column)
Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Conc"

Set Found = Rows(1).Find(what:="ERP_INVOICE_ID", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)
Set Found = Rows(1).Find(what:="Invoice_Line_Item_Number", LookIn:=xlValues, lookat:=xlWhole)
conc2Col = ColumnNumberToLetter(Found.Column)
Set Found = Rows(1).Find(what:="Seller_UEI", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column - 1)

Range(colName & "2").Select

Dim str1, str2 As String
str1 = conc1Col & "2"
str2 = conc2Col & "2"

ActiveCell.Formula = "=" & str1 & "&" & "-" & str2

Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

'Move Concatenate

Set Found = Rows(1).Find(what:="Conc", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column)

Set Found = Rows(1).Find(what:="Line_Item_Amount", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)

Columns(colName & ":" & colName).Cut
Columns(conc1Col & ":" & conc1Col).Insert

'RAW Sheet Creating Recalculatd Debit and Recaclculated Debit



'Concatenate

ws_raw.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Found = Rows(1).Find(what:="ERP INVOICE ID (Journal Lines)", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column + 1)

Set Found = Rows(1).Find(what:="invoice line item number", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)

Set Found = Rows(1).Find(what:="ERP INVOICE ID (Journal Lines)", LookIn:=xlValues, lookat:=xlWhole)
conc2Col = ColumnNumberToLetter(Found.Column)

Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Concatenate"

Timeout (1)


Range(colName & "2").Formula = "=" & conc2Col & "2" & "& ""-"" &" & conc1Col & "2"
Range(colName & "2").AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

Timeout (1)

'Line amount
Dim conCol As String
Set Found = Rows(1).Find(what:="Concatenate", LookIn:=xlValues, lookat:=xlWhole)
conCol = ColumnNumberToLetter(Found.Column)

Set Found = Rows(1).Find(what:="Concatenate", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column + 1)
Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Line Amount"

ws_ich.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Found = Rows(1).Find(what:="Conc", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)


Set Found = Rows(1).Find(what:="Line_Item_Amount", LookIn:=xlValues, lookat:=xlWhole)
conc2Col = ColumnNumberToLetter(Found.Column)

ws_ich.Range(conc1Col & "1:" & conc2Col & LastRow).SpecialCells(xlCellTypeVisible).Copy

Timeout (1)

'creation  of transitory sheet to be deleted later
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Transit" Then
        exists = True
    End If
Next i

If Not exists Then
    Worksheets.Add.Name = "Transit"
End If

Sheets("Transit").Activate

Sheets("Transit").Range("A1:B" & LastRow).PasteSpecial Paste:=xlPasteValues

ws_raw.Activate
Set Found = Rows(1).Find(what:="Concatenate", LookIn:=xlValues, lookat:=xlWhole)
conCol = ColumnNumberToLetter(Found.Column)

Range(colName & "2").Select


ActiveCell.Formula = "=VLOOKUP(" & conCol & "2,Transit!A:B,2,0)"

Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

Timeout (1)
'Currency Column
Set Found = Rows(1).Find(what:="Line Amount", LookIn:=xlValues, lookat:=xlWhole)
conCol = ColumnNumberToLetter(Found.Column + 1)
Range(conCol & ":" & conCol).EntireColumn.Insert
Range(conCol & "1").Select
Selection.Value = "Currency"

ws_ich.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Found = Rows(1).Find(what:="Conc", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)


Set Found = Rows(1).Find(what:="Currency", LookIn:=xlValues, lookat:=xlWhole)
conc2Col = ColumnNumberToLetter(Found.Column)

ws_ich.Range(conc1Col & "1:" & conc1Col & LastRow).SpecialCells(xlCellTypeVisible).Copy

'creation  of transitory sheet to be deleted later
For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Transit2" Then
        exists = True
    End If
Next i

If Not exists Then
    Worksheets.Add.Name = "Transit2"
End If

Sheets("Transit2").Activate

Sheets("Transit2").Range("A1:A" & LastRow).PasteSpecial Paste:=xlPasteValues


ws_ich.Activate
ws_ich.Range(conc2Col & "1:" & conc2Col & LastRow).SpecialCells(xlCellTypeVisible).Copy

Sheets("Transit2").Activate
Sheets("Transit2").Range("B1:B" & LastRow).PasteSpecial Paste:=xlPasteValues



ws_raw.Activate
Timeout (1)

Set Found = Rows(1).Find(what:="Currency", LookIn:=xlValues, lookat:=xlWhole)
conCol = ColumnNumberToLetter(Found.Column)

Range(conCol & "2").Select
Set Found = Rows(1).Find(what:="Concatenate", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column)

ActiveCell.Formula = "=VLOOKUP(" & colName & "2,Transit2!A:B,2,0)"

Range(conCol & "2").Select
Selection.AutoFill Destination:=Range(conCol & "2:" & conCol & LastRow)

ws_raw.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Found = Rows(1).Find(what:="Line Amount", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)

Timeout (1)

'Recalculated Debit
Set Found = Rows(1).Find(what:="Line Item Desc (Journal Lines)", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column)


Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Recalculated debit"
Set Found = Rows(1).Find(what:="Line Amount", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)
Range(colName & "2").Select
ActiveCell.Formula = "=IF(" & conc1Col & "2>0, H2,I2)"
Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

'Recalculated Credit
Set Found = Rows(1).Find(what:="Line Item Desc (Journal Lines)", LookIn:=xlValues, lookat:=xlWhole)
colName = ColumnNumberToLetter(Found.Column)

Range(colName & ":" & colName).EntireColumn.Insert
Range(colName & "1").Select
Selection.Value = "Recalculated credit"
Set Found = Rows(1).Find(what:="Line Amount", LookIn:=xlValues, lookat:=xlWhole)
conc1Col = ColumnNumberToLetter(Found.Column)
Range(colName & "2").Select
ActiveCell.Formula = "=IF(" & conc1Col & "2>0, I2,H2)"
Range(colName & "2").Select
Selection.AutoFill Destination:=Range(colName & "2:" & colName & LastRow)

Sheets("Transit").Visible = False
Sheets("Transit2").Visible = False




Worksheets.Add.Name = "Reclass"



Timeout (1)
Set pt = Sheets("RAW").PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=Sheets("RAW").Range("A1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("Reclass").Range("A1"), _
TableName:="PivotTable1")

Sheets("Reclass").Activate
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Invoice")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "ERP INVOICE ID (Journal Lines)")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Currency")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Account Flexfield (Journal Lines)")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Recalculated debit"), "Sum of Recalculated debit", _
        xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Recalculated credit"), "Sum of Recalculated credit" _
        , xlSum
    With ActiveSheet.PivotTables("PivotTable1").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Invoice"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Invoice")
        .PivotItems("#N/A").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Invoice"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Period End Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Company* (Journal Header)") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Currency (FC)* (Journal Header)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Destination ERP (Journal Lines)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Company Name (Display) (Journal Header)").Subtotals = Array(False, False, False, _
        False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Org ID (Journal Lines)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "BL Journal ID (Display) (Journal Header)").Subtotals = Array(False, False, False _
        , False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Debit (Journal Lines)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Credit (Journal Lines)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Recalculated debit"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Recalculated credit"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Line Item Desc (Journal Lines)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Match Txn ID (Display) (Journal Lines)").Subtotals = Array(False, False, False, _
        False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "ImportDate (Display) (Journal Lines)").Subtotals = Array(False, False, False, _
        False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Account Flexfield (Journal Lines)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("invoice line item number"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Invoice Date* (Journal Lines)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Invoice Number").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "ERP INVOICE ID (Journal Lines)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Concatenate").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Line Amount").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Currency").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target Invoice").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder1 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder2 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder3 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder4 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder5 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Placeholder6 (Journal Lines)").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Process Status").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Error Message").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels




'Amount Consistency Column
Sheets("Reclass").Activate
Sheets("Reclass").Range("F4").Value = "Amount consistency"

Sheets("Reclass").Range("F:F").Columns.AutoFit
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Reclass").Range("F5").Select
ActiveCell.Formula = "=SUMIF(A:A,A4,D:D)=SUMIF(A:A,A4,E:E)"
Selection.AutoFill Destination:=Range("F5:F" & LastRow - 1)


'Account Breakdown
Sheets("Reclass").Range("G4").Value = "Account breakdown"

Sheets("Reclass").Range("G5").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-4],6)"
Selection.AutoFill Destination:=Range("G5:G" & LastRow - 1)

Sheets("Reclass").Range("H5").Select
ActiveCell.Formula = "=MID($C5,8,9)"
Selection.AutoFill Destination:=Range("H5:H" & LastRow - 1)

Sheets("Reclass").Range("I5").Select
ActiveCell.Formula = "=MID($C5,28,1)"
Selection.AutoFill Destination:=Range("I5:I" & LastRow - 1)

Sheets("Reclass").Range("J5").Select
ActiveCell.Formula = "=MID($C5,30,1)"
Selection.AutoFill Destination:=Range("J5:J" & LastRow - 1)

Sheets("Reclass").Range("K5").Select
ActiveCell.Formula = "=MID($C5,22,5)"
Selection.AutoFill Destination:=Range("K5:K" & LastRow - 1)

Sheets("Reclass").Range("L5").Select
ActiveCell.Formula = "=MID($C5,28,1)"
Selection.AutoFill Destination:=Range("L5:L" & LastRow - 1)

Sheets("Reclass").Range("M5").Select
ActiveCell.Formula = "=MID($C5,30,1)"
Selection.AutoFill Destination:=Range("M5:M" & LastRow - 1)

Sheets("Reclass").Range("N5").Select
ActiveCell.Formula = "=MID($C5,32,1)"
Selection.AutoFill Destination:=Range("N5:N" & LastRow - 1)

Range(Range("D5:E5"), Range("D5:E5").End(xlDown)).Select
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "0.000"
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "#,##0.00"
    

'Creating Receivable Consistency

Worksheets.Add.Name = "Receivable Consistency"


Timeout (1)

Set pt = ws_gl.PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=ws_gl.Range("A1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("Receivable Consistency").Range("A1"), _
TableName:="PivotTable1")
Worksheets("Receivable Consistency").Activate
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Flexfield")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Entered DR"), "Sum of Entered DR", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Entered CR"), "Sum of Entered CR", xlSum
    With ActiveSheet.PivotTables("PivotTable1").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice")
        .PivotItems("#N/A").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Target invoice"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Flexfield"). _
        CurrentPage = "(All)"

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Accounting Flexfield"). _
        EnableMultiplePageItems = True




f = "030031100"

Set PvtTbl = Sheets("Receivable Consistency").PivotTables("PivotTable1")

With PvtTbl.PivotFields("Accounting Flexfield")
    .ClearAllFilters

    For Each PvtItm In .PivotItems
        If Not PvtItm.Name Like "*" & f & "*" Then
            PvtItm.Visible = False

        End If
    Next PvtItm
End With

Set pt = ws_raw.PivotTableWizard( _
SourceType:=xlDatabase, _
SourceData:=ws_raw.Range("A1").CurrentRegion, _
TableDestination:=ThisWorkbook.Sheets("Receivable Consistency").Range("H1"), _
TableName:="PivotTable2")
Worksheets("Receivable Consistency").Activate
Range("H1").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Target Invoice")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Account Flexfield (Journal Lines)")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Recalculated debit"), "Sum of Recalculated debit", _
        xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Recalculated credit"), "Sum of Recalculated credit" _
        , xlSum
    With ActiveSheet.PivotTables("PivotTable2").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Target Invoice"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Target Invoice")
        .PivotItems("#N/A").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Target Invoice"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Account Flexfield (Journal Lines)").CurrentPage = "(All)"

    With ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Account Flexfield (Journal Lines)")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Account Flexfield (Journal Lines)").EnableMultiplePageItems = True


'Pivot Filtration Criteria
f = "030031100"

Set PvtTbl = Sheets("Receivable Consistency").PivotTables("PivotTable2")

With PvtTbl.PivotFields("Account Flexfield (Journal Lines)")
    .ClearAllFilters

    For Each PvtItm In .PivotItems
        If Not PvtItm.Name Like "*" & f & "*" Then
            PvtItm.Visible = False

        End If
    Next PvtItm
End With


Range("D6").FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("K6").FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("B6:D6").Style = "Comma"
Range("I6:K6").Style = "Comma"
Range("L5").FormulaR1C1 = "Consistency"
Range("L6").FormulaR1C1 = "=RC[-8]=RC[-1]"
Range("N5").FormulaR1C1 = "Delta"
Range("N6").FormulaR1C1 = "=RC[-10]-RC[-3]"
Range("D6").NumberFormat = "#,##0.00_);(#,##0.00)"
Range("K6").NumberFormat = "#,##0.00_);(#,##0.00)"


   
   
   
Dim ws_output As Worksheet: Set ws_output = ThisWorkbook.Sheets("Output")

LastRow = Sheets("Reclass").Cells(Rows.Count, 1).End(xlUp).Row - 1



Sheets("Reclass").Range("G5:N" & LastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("Output").Range("F3:M" & LastRow + 2).PasteSpecial Paste:=xlPasteValues


Sheets("Output").Range("C3:C" & LastRow - 2).Value = "Reclass"
Sheets("Output").Range("D3:D" & LastRow - 2).Value = Sheets("Reclass").Range("B5:B" & LastRow).Value


Sheets("Output").Range("N3:O" & LastRow - 2).Value = Sheets("Reclass").Range("D5:E" & LastRow).Value


LR_Reversal = Sheets("Output").Cells(Rows.Count, 3).End(xlUp).Row + 1

LastRow = Sheets("Reversal").Cells(Rows.Count, 1).End(xlUp).Row - 1



Sheets("Reversal").Range("I5:P" & LastRow).SpecialCells(xlCellTypeVisible).Copy
Sheets("Output").Range("F" & LR_Reversal & ":M" & LastRow + 2 + LR_Reversal).PasteSpecial Paste:=xlPasteValues


Sheets("Output").Range("C" & LR_Reversal & ":C" & LastRow + LR_Reversal - 2).Value = "Reversal"
Sheets("Output").Range("D" & LR_Reversal & ":D" & LastRow - 2 + LR_Reversal).Value = Sheets("Reversal").Range("B5:B" & LastRow).Value


Sheets("Output").Range("N" & LR_Reversal & ":O" & LastRow - 2 + LR_Reversal).Value = Sheets("Reversal").Range("G5:H" & LastRow).Value

Sheets("Output").Activate

LastRow = Sheets("Output").Cells(Rows.Count, 3).End(xlUp).Row
Range("$A$1:$Y$" & LastRow).AutoFilter Field:=14, Criteria1:=Array("<>(Blanks)", "<>#N/A")

Range("N3:O" & LastRow).Select
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "0.000"
Selection.NumberFormat = "0.00"
Selection.NumberFormat = "#,##0.00"



Application.CutCopyMode = False


MsgBox "Macro Finished Running."
End Sub



