Attribute VB_Name = "Module1"
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.ThisWorkbook.Saved = True
Dim answer As Integer
Dim cellcheck, columncheck, thissheet As String
Dim lastrow As Long
Dim CellsArray() As Variant
Dim count As Integer

CellsArray = Array("G2", "K2", "M2")
lastrow = Sheets("1.survey").Cells(Rows.count, "A").End(xlUp).Row - 1

cellcheck = ""
columncheck = ""
count = 0

sso = ThisWorkbook.Worksheets("1.survey").Range("G2")

For j = 1 To ThisWorkbook.Sheets.count


    If InStr(1, ThisWorkbook.Sheets(j).Name, sso) <> 0 Or ThisWorkbook.Sheets(j).Name = "1.survey" Then
        thissheet = ThisWorkbook.Sheets(j).Name

        For i = 0 To 2
        If Sheets(thissheet).Range(CellsArray(i)) = "" Then
        cellcheck = "Cell " & CellsArray(i) & " in Sheet """ & thissheet & """ is empty." & vbNewLine & cellcheck
        End If
        Next i

        columncheck = columncheck & vbNewLine & "In Sheet """ & thissheet & """ these columns are empty:"
        If WorksheetFunction.CountA(Worksheets(thissheet).Range("F5:F" & lastrow)) = 0 Then
             columncheck = columncheck & vbNewLine & "Column F, ""Total Hrs per quarter"""
        Else
        count = Inc(0)
        End If

        If WorksheetFunction.CountA(Worksheets(thissheet).Range("L5:L" & lastrow)) = 0 Then
             columncheck = columncheck & vbNewLine & "Column L, ""Company Code"""
        Else
        count = Inc(1)
        End If
        If WorksheetFunction.CountA(Worksheets(thissheet).Range("T5:T" & lastrow)) = 0 Then
            columncheck = columncheck & vbNewLine & "Column T, ""Activities/Recons?"""
        Else
        count = Inc(2)
        End If
        If WorksheetFunction.CountA(Worksheets(thissheet).Range("U5:U" & lastrow)) = 0 Then
             columncheck = columncheck & vbNewLine & "Column U, ""Functional Team"""
        Else
        count = Inc(3)
        End If
        If WorksheetFunction.CountA(Worksheets(thissheet).Range("V5:V" & lastrow)) = 0 Then
             columncheck = columncheck & vbNewLine & "Column V,""Functional Team Lead"""
        Else
        count = Inc(4)
        End If

        If count = 5 Then
        columncheck = ""
        End If
        count = 0
    End If

Next j

If IsEmpty(Sheets("1.survey").Range("G2")) Then
    answer = MsgBox("Click Yes to Continue Editing, No to Close the file without saving changes.", vbQuestion + vbYesNo + vbDefaultButton2, "Check Before Saving")
    If answer = vbYes Then
      Cancel = True
      MsgBox "Cell G2 ""4-3-1/SSO"" is needed"
    Else
        ThisWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
ElseIf Not IsEmpty(Sheets("1.survey").Range("G2")) And (Not IsEmpty(cellcheck) Or Not IsEmpty(columncheck)) Then
    answer = MsgBox("Click Yes to Continue Editing, Click No to Close the file without saving changes.", vbQuestion + vbYesNo + vbDefaultButton2, "Check Before Saving")
    If answer = vbYes Then
        Cancel = True
        SaveAsUI = False
        If Not cellcheck = "" Then
            MsgBox cellcheck
        End If
        If Not columncheck = "" Then
            MsgBox columncheck
        End If
    Else
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
    End If
End If


End Sub



