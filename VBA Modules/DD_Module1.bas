Attribute VB_Name = "Module2"
Sub consolidate()
Dim Myfile, Mypath, MAIN As String
Dim sor, oszlop, SORMAIN As Variant
Dim a As Integer

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationAutomatic


Mypath = ActiveWorkbook.Path
'Mypath = Mypath & "\hatobbvan"

MAIN = ActiveWorkbook.Name

Myfile = Dir(Mypath & "\")


Do While Myfile <> ""

If Myfile <> MAIN Then 'I change it
Workbooks.Open Mypath & "\" & Myfile
Sheets("1.survey").Select
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0

Do While Range("g2") <> ""
SSO = Range("g2").Value
Name = Range("i2").Value
ERP = Range("k2").Value
LE = Range("o3").Value

Range("a4").Select
sor = Range(Selection, Selection.End(xlDown)).Count + 3
oszlop = Range(Selection, Selection.End(xlToRight)).Count

Rows("4:4").Select
Selection.AutoFilter
ActiveSheet.Range("a4:q" & sor).AutoFilter Field:=6, Criteria1:="<>"

If ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1) <> "" Then ' if all the wos are empty the macro just jump through the codes below and continueing from the do while:)

Range(Cells(5, 1), Cells(sor, oszlop)).SpecialCells(xlCellTypeVisible).Copy

Workbooks(MAIN).Activate

If Range("a5") <> "" Then
Range("a" & sor2 + 1).PasteSpecial xlPasteValuesAndNumberFormats
Range("a4").Select

sor3 = Range(Selection, Selection.End(xlDown)).Count + 3
Range(Cells(sor2 + 1, "r"), Cells(sor3, "r")) = SSO
Range(Cells(sor2 + 1, "s"), Cells(sor3, "s")) = Name
Range(Cells(sor2 + 1, "t"), Cells(sor3, "t")) = ERP
Range(Cells(sor2 + 1, "u"), Cells(sor3, "u")) = LE
sor2 = sor3
Else:

Range("a5").PasteSpecial xlPasteValuesAndNumberFormats
Range("a4").Select
sor2 = Range(Selection, Selection.End(xlDown)).Count + 3
Range(Cells(5, "r"), Cells(sor2, "r")) = SSO
Range(Cells(5, "s"), Cells(sor2, "s")) = Name
Range(Cells(5, "t"), Cells(sor2, "t")) = ERP
Range(Cells(5, "u"), Cells(sor2, "u")) = LE
Range("a4").Select
sor2 = Range(Selection, Selection.End(xlDown)).Count + 3
End If

Workbooks(Myfile).Activate
Columns("F:Q").Delete xlToLeft
Else: End If
Loop
Workbooks(Myfile).Close False
Else: End If
Myfile = Dir

Loop

Workbooks(MAIN).Activate
a = 4
Do While Cells(a, 1) <> ""
If Cells(a, 1) = "Total FTE" Then
    Rows(a).Delete Shift:=xlUp
    'Row counter should not be incremented if row was just deleted
Else
    'Increment a for next row only if row not deleted
    a = a + 1
End If

Loop

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
End Sub

