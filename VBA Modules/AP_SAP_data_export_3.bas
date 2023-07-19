Attribute VB_Name = "AP_SAP_data_export_3"
Function first()
first = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Row
End Function
Function last()
'last = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(Rows.Count, "I").End(xlUp).Row
last = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).Row
End Function
Sub Sap_data_export_AP()
Dim x, j As Integer
Dim lastrow, firstrow As Long
Dim i, limit As Integer
Dim currentpath As String
Dim dur As String

Application.DisplayAlerts = False


Dim main_ws As Worksheet: Set main_ws = ThisWorkbook.Sheets("Proposed_IC_Settlements")

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection


session.findById("wnd[0]").resizeWorkingPane 128, 39, False

main_ws.Activate
lastrow = last
firstrow = first
x = 0


If lastrow < 40000 Then
j = 10
ElseIf 40000 <= lastrow < 60000 Then
j = 12
ElseIf 60000 <= lastrow < 100000 Then
j = 20
End If

OpenStatusBar
currtime = Time()
firstrow = first

For i = j To 1 Step -1:

    dur = Format(Now() - currtime, "hh:mm:ss")
    DoEvents
    Call RunStatusBar(x, j, dur)
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfbl1n"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "204040644"
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").collapseNode "          1"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").collapseNode "         28"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         60"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         54"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         60"
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    Application.CutCopyMode = True
    main_ws.Activate
    lastrow = last
    
    limit = Int(lastrow - ((lastrow / j) * i) + (lastrow / j))
    
    
    'Get earliest and latest year to set period
    main_ws.Range("N" & firstrow & ":N" & limit).NumberFormat = "m/d/yyyy"
    main_ws.Range("N" & firstrow & ":N" & limit).NumberFormat = "yyyy"
    
    xmin = Application.WorksheetFunction.Min(main_ws.Range("N" & firstrow & ":N" & limit))
    xmax = Application.WorksheetFunction.Max(main_ws.Range("N" & firstrow & ":N" & limit))
    xmax = Format(xmax, "yyyy")
    xmin = Format(xmin, "yyyy")
    main_ws.Range("N" & firstrow & ":N" & limit).NumberFormat = "m/d/yyyy h:mm:ss"
    
    main_ws.Range("I" & firstrow & ":I" & limit).Copy
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press 'Reference
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press 'Upload from clipboard
    session.findById("wnd[1]/tbar[0]/btn[8]").press ' Copy
    session.findById("wnd[0]/usr/radX_CLSEL").Select
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-LOW").Text = "0101" & xmin
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-LOW").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").Text = "3106" & xmax ' change it later
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    'Layout
    session.findById("wnd[0]/tbar[1]/btn[32]").press
    session.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(1).Selected = True
    session.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(2).Selected = True
    session.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST/txtGT_FIELD_LIST-SELTEXT[0,2]").SetFocus
    session.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST/txtGT_FIELD_LIST-SELTEXT[0,2]").caretPosition = 0
    session.findById("wnd[1]/usr/btnAPP_WL_SING").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    'Export
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = dynamicPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "AP_Export" & i & ".xlsx"
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    Debug.Print firstrow, limit
    firstrow = limit + 1
    
    Application.CutCopyMode = False
    Timeout (15)
    x = x + 1
    
    dur = Format(Now() - currtime, "hh:mm:ss")
    DoEvents
    Call RunStatusBar(x, j, dur)

    Call HandleExport(i, "AP_SAP vendor line", "AP_Export")
    Debug.Print

Next i

Unload StatusBar
Application.DisplayAlerts = True


End Sub

