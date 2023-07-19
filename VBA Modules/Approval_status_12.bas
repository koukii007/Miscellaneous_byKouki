Attribute VB_Name = "Approval_status_12"
Public filename, sheet As String
Dim FolderPath, FilePath, colName, s As String
Dim oFSO, oFolder, oFile As Object
Dim i As Integer
Dim lastrow, firstrow As Long
Dim item As Variant
Function first()
first = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Row
End Function
Function last()
'last = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(Rows.Count, "I").End(xlUp).Row
last = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
End Function
Sub Approval_Status_check()
Dim main_ws As Worksheet: Set main_ws = ThisWorkbook.Sheets("AR_Invoice_Export")

'filter on non N/A

colName = GetColumnName(1, 0, "Non PO Check")
main_ws.Range(colName & ":" & colName).AutoFilter Field:=GetColumnindex(1, "Non PO Check"), Criteria1:="<>#N/A"


''Create Approval Status Column
colName = GetColumnName(1, 1, "Non PO Check")
main_ws.Columns(colName).Insert
main_ws.Range(colName & "1") = "Approval Status"


firstrow = first
lastrow = last
nonPOCol = GetColumnName(1, 0, "NONPO_ICH_WF_NUMBER (AR_INVOICES)")
nonPOfile = GetFileName("Non PO WF")
main_ws.Range(colName & firstrow & ":" & colName & lastrow).SpecialCells(xlCellTypeVisible) = "=XLOOKUP(" & nonPOCol & firstrow & ",'" & nonPOfile & "'!$D:$D,'" & nonPOfile & "'!$E:$E)"

'NON PO WF is not approved
colName = GetColumnName(1, 0, "Approval Status")
main_ws.Range(colName & ":" & colName).AutoFilter Field:=GetColumnindex(1, "Approval Status"), Criteria1:=Array("Rejected", "Prepared", _
                                    "Not Prepared"), _
                    Operator:=xlFilterValues


If ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2) <> "" Then
    firstrow = first
    lastrow = last

    If WorksheetFunction.CountA(main_ws.Range("B2:B" & lastrow).SpecialCells(xlCellTypeVisible)) > 1 Then

       main_ws.Range("A" & firstrow & ":A" & lastrow).SpecialCells(xlCellTypeVisible) = "Non PO WF is not approved"

    ElseIf WorksheetFunction.CountA(main_ws.Range("B2:B" & lastrow).SpecialCells(xlCellTypeVisible)) = 1 Then

       main_ws.Range("A" & firstrow & ":A" & lastrow) = "Non PO WF is not approved"

    End If

End If



'Approved/ Reviewed Approval Status
colName = GetColumnName(1, 0, "Approval Status")
main_ws.Range(colName & ":" & colName).AutoFilter Field:=GetColumnindex(1, "Approval Status"), Criteria1:=Array("Approved", "Reviewed"), _
    Operator:=xlFilterValues

'Add 4 Checks
'1-Non PO initiator
colName = GetColumnName(1, 1, "Approval Status")
main_ws.Columns(colName).Insert
main_ws.Range(colName & "1") = "Non PO Initiator"
firstrow = first
lastrow = last
nonPOCol = GetColumnName(1, 0, "NONPO_ICH_WF_NUMBER (AR_INVOICES)")
main_ws.Range(colName & firstrow & ":" & colName & lastrow).SpecialCells(xlCellTypeVisible) = "=LEFT(VLOOKUP(TEXT(" & nonPOCol & firstrow & ",""@""),'" & nonPOfile & "'!$D:$F,3,0),10)"


''2- Non PO Recipient
colName = GetColumnName(1, 1, "Non PO Initiator")
main_ws.Columns(colName).Insert
main_ws.Range(colName & "1") = "Non PO Recipient"
firstrow = first
lastrow = last
nonPOCol = GetColumnName(1, 0, "NONPO_ICH_WF_NUMBER (AR_INVOICES)")
main_ws.Range(colName & firstrow & ":" & colName & lastrow).SpecialCells(xlCellTypeVisible) = "=LEFT(VLOOKUP(TEXT(" & nonPOCol & firstrow & ",""@""),'" & nonPOfile & "'!$D:$G,4,0),10)"

'3- Initiator Check
colName = GetColumnName(1, 1, "Non PO Initiator")
main_ws.Columns(colName).Insert
main_ws.Range(colName & "1") = "Initiator Check"
firstrow = first
lastrow = last
initiatorCol = GetColumnName(1, 0, "Non PO Initiator")
sellerUEICol = GetColumnName(1, 0, "SELLER_UEI (AR_INVOICES)")
main_ws.Range(colName & firstrow & ":" & colName & lastrow).SpecialCells(xlCellTypeVisible) = "=TEXT(" & initiatorCol & firstrow & ",""@"")=TEXT(" & sellerUEICol & firstrow & ",""@"")"


'4- Recipient Check
colName = GetColumnName(1, 1, "Non PO Recipient")
main_ws.Columns(colName).Insert
main_ws.Range(colName & "1") = "Recipient Check"
firstrow = first
lastrow = last
recipientCol = GetColumnName(1, 0, "Non PO Recipient")
buyerUEICol = GetColumnName(1, 0, "BUYER_UEI (AR_INVOICES)")
main_ws.Range(colName & firstrow & ":" & colName & lastrow).SpecialCells(xlCellTypeVisible) = "=TEXT(" & recipientCol & firstrow & ",""@"")=TEXT(" & buyerUEICol & firstrow & ",""@"")"

Debug.Print "Approval status , recipient initatior..."
End Sub
