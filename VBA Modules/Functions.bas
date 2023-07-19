Attribute VB_Name = "Functions"
Function GetWorksheetName(ByVal workbookName As String)
        Dim sheetName As String
        For i = 1 To Workbooks(workbookName).Sheets.Count
            If InStr(1, Workbooks(workbookName).Sheets(i).Name, "PO_Repository") <> 0 Then
                sheetName = Workbooks(workbookName).Sheets(i).Name
                Exit For
            End If
        Next i
        GetWorksheetName = sheetName
End Function
Function GetFileName(ByVal lookfor As String)
FolderPath = GetLocalPath(ActiveWorkbook.path)

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(FolderPath & "\")

For Each oFile In oFolder.Files
    If InStr(1, oFile, lookfor) <> 0 Then
        thisfilename = oFSO.GetFileName(oFile)
    End If
Next oFile
GetFileName = thisfilename
End Function


Function GetColumnName(ByVal rownb, position As Integer, lookfor As String)  ' rownb is for the row number where you're looking for the match

Set Found = Rows(rownb).Find(what:="*" & lookfor & "*", LookIn:=xlValues, lookat:=xlWhole)

If Not Found Is Nothing Then
GetColumnName = Replace(Replace(Cells(1, Found.Column + position).Address, "1", ""), "$", "") ' position 0 if you need the actuall column, 1 if you need the Next etc
Else
GetColumnName = ""
End If

End Function

Function GetColumnindex(ByVal rownb As Integer, lookfor As String)  ' rownb is for the row number where you're looking for the match

Set Found = Rows(rownb).Find(what:="*" & lookfor & "*", LookIn:=xlValues, lookat:=xlWhole)

If Not Found Is Nothing Then
GetColumnindex = Found.Column
Else
GetColumnindex = 0
End If

End Function


Function ColumnstoDelete() As Variant


Dim ColumnNames(44) As Variant

ColumnNames(0) = "ERP_Batch_ID (AR_Invoices)"
ColumnNames(1) = "Seller_VAT_NumberTax_ID (AR_Invoices)"
ColumnNames(2) = "Buyer_VAT_NumberTAX_ID (AR_Invoices)"
ColumnNames(3) = "Buyer_SSO (AR_Invoices)"
ColumnNames(4) = "Ship_To_Address (AR_Invoices)"
ColumnNames(5) = "Ship_To_City (AR_Invoices)"
ColumnNames(6) = "Ship_To_State (AR_Invoices)"
ColumnNames(7) = "Ship_To_Zip (AR_Invoices)"
ColumnNames(8) = "Ship_From_Address (AR_Invoices)"
ColumnNames(9) = "Ship_From_City (AR_Invoices)"
ColumnNames(10) = "Ship_From_State (AR_Invoices)"
ColumnNames(11) = "Ship_From_Zip (AR_Invoices)"
ColumnNames(12) = "Date_Shipped (AR_Invoices)"
ColumnNames(13) = "Shipping_Costs (AR_Invoices)"
ColumnNames(14) = "Bill_of_Lading (AR_Invoices)"
ColumnNames(15) = "Shipped_Via (AR_Invoices)"
ColumnNames(16) = "Transportation_Terms (AR_Invoices)"
ColumnNames(17) = "Invoice_Discount_Percent (AR_Invoices)"
ColumnNames(18) = "Invoice_Discount_Amount (AR_Invoices)"
ColumnNames(19) = "Book_Currency_Code (AR_Invoices)"
ColumnNames(20) = "Book_Amount (AR_Invoices)"
ColumnNames(21) = "Supplier_Product_Code (AR_Invoices)"
ColumnNames(22) = "Cur_of_Record_Unit_Price (AR_Invoices)"
ColumnNames(23) = "PO_Number_Multiple (AR_Invoices)"
ColumnNames(24) = "Unit_of_Measure (AR_Invoices)"
ColumnNames(25) = "Extended_Amount (AR_Invoices)"
ColumnNames(26) = "Price_Per (AR_Invoices)"
ColumnNames(27) = "UNSPSC_Code (AR_Invoices)"
ColumnNames(28) = "Discount_Percent (AR_Invoices)"
ColumnNames(29) = "Discount_Amount (AR_Invoices)"
ColumnNames(30) = "Special_Charge_Code (AR_Invoices)"
ColumnNames(31) = "Special_Charge_Text (AR_Invoices)"
ColumnNames(32) = "Special_Charge_Amount (AR_Invoices)"
ColumnNames(33) = "ADN (AR_Invoices)"
ColumnNames(34) = "Org_ID (AR_Invoices)"
ColumnNames(35) = "Fiscal_Code (AR_Invoices)"
ColumnNames(36) = "UniqueKey (AR_Invoices)"
ColumnNames(37) = "Entity_Relationship (AR_Invoices)"
ColumnNames(38) = "Entity_Relationship_ID (AR_Invoices)"
ColumnNames(39) = "Receiving_Status (AR_Invoices)"
ColumnNames(40) = "BLDisputeID (AR_Invoices)"
ColumnNames(41) = "SourceFileName (AR_Invoices)"
ColumnNames(42) = "ImportDate (AR_Invoices)"
ColumnNames(43) = "JobProfileID (AR_Invoices)"
ColumnNames(44) = "IsUpdatedRow (AR_Invoices)"

ColumnstoDelete = ColumnNames

End Function
Public Function GetLocalPath(ByVal path As String, _
                    Optional ByVal rebuildCache As Boolean = False, _
                    Optional ByVal returnAll As Boolean = False, _
                    Optional ByVal preferredMountPointOwner As String = "") _
                             As String
    #If Mac Then
        Const vbErrPermissionDenied As Long = 70
        Const vbErrInvalidFormatInResourceFile As Long = 325
        Const ps As String = "/"
    #Else
        Const ps As String = "\"
    #End If
    Const vbErrFileNotFound As Long = 53
    Static locToWebColl As Collection, lastTimeNotFound As Collection
    Static lastCacheUpdate As Date
    Dim resColl As Object, webRoot As String, locRoot As String
    Dim vItem As Variant, s As String, keyExists As Boolean
    Dim pmpo As String: pmpo = LCase(preferredMountPointOwner)

    If Not locToWebColl Is Nothing And Not rebuildCache Then
        Set resColl = New Collection: GetLocalPath = ""
        For Each vItem In locToWebColl
            locRoot = vItem(0): webRoot = vItem(1)
            If InStr(1, path, webRoot, vbTextCompare) = 1 Then _
                resColl.Add Key:=vItem(2), _
                   item:=Replace(Replace(path, webRoot, locRoot, , 1), "/", ps)
        Next vItem
        If resColl.Count > 0 Then
            If returnAll Then
                For Each vItem In resColl: s = s & "//" & vItem: Next vItem
                GetLocalPath = Mid(s, 3): Exit Function
            End If
            On Error Resume Next: GetLocalPath = resColl(pmpo): On Error GoTo 0
            If GetLocalPath <> "" Then Exit Function
            GetLocalPath = resColl(1): Exit Function
        End If
        If Not lastTimeNotFound Is Nothing Then
            On Error Resume Next: lastTimeNotFound path
            keyExists = (Err.Number = 0): On Error GoTo 0
            If keyExists Then
                If DateAdd("s", 1, lastTimeNotFound(path)) > Now() Then _
                    GetLocalPath = path: Exit Function
            End If
        End If
        GetLocalPath = path
    End If

    Dim cid As String, fileNum As Long, line As Variant, parts() As String
    Dim tag As String, mainMount As String, relPath As String, email As String
    Dim b() As Byte, n As Long, i As Long, size As Long, libNr As String
    Dim parentID As String, folderID As String, folderName As String
    Dim folderIdPattern As String, filename As String, folderType As String
    Dim siteID As String, libID As String, webID As String, lnkID As String
    Dim odFolders As Object, cliPolColl As Object, libNrToWebColl As Object
    Dim sig1 As String: sig1 = StrConv(Chr$(&H2), vbFromUnicode)
    Dim sig2 As String: sig2 = ChrW$(&H1) & String(3, vbNullChar)
    Dim vbNullByte As String: vbNullByte = MidB$(vbNullChar, 1, 1)
    #If Mac Then
        Dim utf16() As Byte, utf32() As Byte, j As Long, k As Long, m As Long
        Dim charCode As Long, lowSurrogate As Long, highSurrogate As Long
        ReDim b(0 To 3): b(0) = &HAB&: b(1) = &HAB&: b(2) = &HAB&: b(3) = &HAB&
        Dim sig3 As String: sig3 = b: sig3 = vbNullChar & vbNullChar & sig3
    #Else
        ReDim b(0 To 1): b(0) = &HAB&: b(1) = &HAB&
        Dim sig3 As String: sig3 = b: sig3 = vbNullChar & sig3
    #End If

    Dim settPath As String, wDir As String, clpPath As String
    #If Mac Then
        s = Environ("HOME")
        settPath = Left(s, InStrRev(s, "/Library/Containers")) & _
                   "Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
                   "Library/Application Support/OneDrive/settings/"
        clpPath = s & "/Library/Application Support/Microsoft/Office/CLP/"
    #Else
        settPath = Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
        clpPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
    #End If

    #If Mac Then
        Dim possibleDirs(0 To 11) As String: possibleDirs(0) = settPath
        For i = 1 To 9: possibleDirs(i) = settPath & "Business" & i & ps: Next i
       possibleDirs(10) = settPath & "Personal" & ps: possibleDirs(11) = clpPath
        If Not GrantAccessToMultipleFiles(possibleDirs) Then _
            Err.Raise vbErrPermissionDenied
    #End If

    Dim oneDriveSettDirs As Collection: Set oneDriveSettDirs = New Collection
    Dim dirName As Variant: dirName = Dir(settPath, vbDirectory)
    Do Until dirName = ""
        If dirName = "Personal" Or dirName Like "Business#" Then _
            oneDriveSettDirs.Add dirName
        dirName = Dir(, vbDirectory)
    Loop

    #If Mac Then
        s = ""
        For Each dirName In oneDriveSettDirs
            wDir = settPath & dirName & ps
            cid = IIf(dirName = "Personal", "????????????????", _
                      "????????-????-????-????-????????????")
           If dirName = "Personal" Then s = s & "//" & wDir & "GroupFolders.ini"
            s = s & "//" & wDir & "global.ini"
            filename = Dir(wDir, vbNormal)
            Do Until filename = ""
                If filename Like cid & ".ini" Or _
                   filename Like cid & ".dat" Or _
                   filename Like "ClientPolicy*.ini" Then _
                    s = s & "//" & wDir & filename
                filename = Dir
            Loop
        Next dirName
        If Not GrantAccessToMultipleFiles(Split(Mid(s, 3), "//")) Then _
            Err.Raise vbErrPermissionDenied
    #End If

    If Not locToWebColl Is Nothing And Not rebuildCache Then
        s = ""
        For Each dirName In oneDriveSettDirs
            wDir = settPath & dirName & ps
            cid = IIf(dirName = "Personal", "????????????????", _
                      "????????-????-????-????-????????????")
            If Dir(wDir & "global.ini") <> "" Then _
                s = s & "//" & wDir & "global.ini"
            filename = Dir(wDir, vbNormal)
            Do Until filename = ""
                If filename Like cid & ".ini" Then _
                    s = s & "//" & wDir & filename
                filename = Dir
            Loop
        Next dirName
        For Each vItem In Split(Mid(s, 3), "//")
            If FileDateTime(vItem) > lastCacheUpdate Then _
                rebuildCache = True: Exit For
        Next vItem
        If Not rebuildCache Then
            If lastTimeNotFound Is Nothing Then _
                Set lastTimeNotFound = New Collection
            On Error Resume Next: lastTimeNotFound.Remove path: On Error GoTo 0
            lastTimeNotFound.Add item:=Now(), Key:=path
            Exit Function
        End If
    End If
    
    lastCacheUpdate = Now()
    Set lastTimeNotFound = Nothing

    Set locToWebColl = New Collection
    For Each dirName In oneDriveSettDirs
        wDir = settPath & dirName & ps
        If Dir(wDir & "global.ini", vbNormal) = "" Then GoTo NextFolder
        fileNum = FreeFile()
        Open wDir & "global.ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then
            b = StrConv(b, vbUnicode)
        #End If
        For Each line In Split(b, vbNewLine)
            If line Like "cid = *" Then cid = Mid(line, 7): Exit For
        Next line

        If cid = "" Then GoTo NextFolder
        If (Dir(wDir & cid & ".ini") = "" Or _
            Dir(wDir & cid & ".dat") = "") Then GoTo NextFolder
        If dirName Like "Business#" Then
            folderIdPattern = Replace(Space(32), " ", "[a-f0-9]")
        ElseIf dirName = "Personal" Then
            folderIdPattern = Replace(Space(16), " ", "[A-F0-9]") & "!###*"
        End If

        filename = Dir(clpPath, vbNormal)
        Do Until filename = ""
            If InStr(1, filename, cid) And cid <> "" Then _
                email = LCase(Left(filename, InStr(filename, cid) - 2)): Exit Do
            filename = Dir
        Loop

        Set cliPolColl = New Collection
        filename = Dir(wDir, vbNormal)
        Do Until filename = ""
            If filename Like "ClientPolicy*.ini" Then
                fileNum = FreeFile()
                Open wDir & filename For Binary Access Read As #fileNum
                    ReDim b(0 To LOF(fileNum)): Get fileNum, , b
                Close #fileNum: fileNum = 0
                #If Mac Then
                    b = StrConv(b, vbUnicode)
                #End If
                cliPolColl.Add Key:=filename, item:=New Collection
                For Each line In Split(b, vbNewLine)
                    If InStr(1, line, " = ", vbBinaryCompare) Then
                        tag = Left(line, InStr(line, " = ") - 1)
                        s = Mid(line, InStr(line, " = ") + 3)
                        Select Case tag
                        Case "DavUrlNamespace"
                            cliPolColl(filename).Add Key:=tag, item:=s
                        Case "SiteID", "IrmLibraryId", "WebID"
                            s = Replace(LCase(s), "-", "")
                            If Len(s) > 3 Then s = Mid(s, 2, Len(s) - 2)
                            cliPolColl(filename).Add Key:=tag, item:=s
                        End Select
                    End If
                Next line
            End If
            filename = Dir
        Loop

        fileNum = FreeFile
        Open wDir & cid & ".dat" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b: s = b: size = LenB(s)
        Close #fileNum: fileNum = 0
        Set odFolders = New Collection
        For Each vItem In Array(16, 8)
            i = InStrB(vItem, s, sig2)
            Do While i > vItem And i < size - 168
                If MidB$(s, i - vItem, 1) = sig1 Then
                    i = i + 8: n = InStrB(i, s, vbNullByte) - i
                    If n < 0 Then n = 0
                    If n > 39 Then n = 39
                    folderID = StrConv(MidB$(s, i, n), vbUnicode)
                    i = i + 39: n = InStrB(i, s, vbNullByte) - i
                    If n < 0 Then n = 0
                    If n > 39 Then n = 39
                    parentID = StrConv(MidB$(s, i, n), vbUnicode)
                    i = i + 121: n = -Int(-(InStrB(i, s, sig3) - i) / 2) * 2
                    If n < 0 Then n = 0
                    #If Mac Then
                        utf32 = MidB$(s, i, n)
                        ReDim utf16(LBound(utf32) To UBound(utf32))
                        j = LBound(utf32): k = LBound(utf32)
                        Do While j < UBound(utf32)
                            If utf32(j + 2) = 0 And utf32(j + 3) = 0 Then
                                utf16(k) = utf32(j): utf16(k + 1) = utf32(j + 1)
                                k = k + 2
                            Else
                                If utf32(j + 3) <> 0 Then Err.Raise _
                                    vbErrInvalidFormatInResourceFile
                                charCode = utf32(j + 2) * &H10000 + _
                                           utf32(j + 1) * &H100& + utf32(j)
                                m = charCode - &H10000
                                highSurrogate = &HD800& + (m \ &H400&)
                                lowSurrogate = &HDC00& + (m And &H3FF)
                                utf16(k) = CByte(highSurrogate And &HFF&)
                                utf16(k + 1) = CByte(highSurrogate \ &H100&)
                                utf16(k + 2) = CByte(lowSurrogate And &HFF&)
                                utf16(k + 3) = CByte(lowSurrogate \ &H100&)
                                k = k + 4
                            End If
                            j = j + 4
                        Loop
                        ReDim Preserve utf16(LBound(utf16) To k - 1)
                        folderName = utf16
                    #Else
                        folderName = MidB$(s, i, n)
                    #End If
                    If folderID Like folderIdPattern Then
                        odFolders.Add VBA.Array(parentID, folderName), folderID
                    End If
                End If
                i = InStrB(i + 1, s, sig2)
            Loop
            If odFolders.Count > 0 Then Exit For
        Next vItem

        fileNum = FreeFile()
        Open wDir & cid & ".ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then
            b = StrConv(b, vbUnicode)
        #End If
        Select Case True
        Case dirName Like "Business#"
            mainMount = "": Set libNrToWebColl = New Collection
            For Each line In Split(b, vbNewLine)
                webRoot = "": locRoot = ""
                Select Case Left$(line, InStr(line, " = ") - 1)
                Case "libraryScope"
                    parts = Split(line, """"): locRoot = parts(9)
                    If locRoot = "" Then libNr = Split(line, " ")(2)
                    folderType = parts(3): parts = Split(parts(8), " ")
                    siteID = parts(1): webID = parts(2): libID = parts(3)
                    If mainMount = "" And folderType = "ODB" Then
                        mainMount = locRoot: filename = "ClientPolicy.ini"
                    Else: filename = "ClientPolicy_" & libID & siteID & ".ini"
                    End If
                    On Error Resume Next
                    webRoot = cliPolColl(filename)("DavUrlNamespace")
                    On Error GoTo 0
                    If webRoot = "" Then
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID And vItem("WebID") = _
                            webID And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace"): Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = "" Then Err.Raise vbErrFileNotFound
                    If locRoot = "" Then
                        libNrToWebColl.Add VBA.Array(libNr, webRoot), libNr
                    Else: locToWebColl.Add VBA.Array(locRoot, webRoot, email), _
                                           locRoot
                    End If
                Case "libraryFolder"
                    locRoot = Split(line, """")(1): libNr = Split(line, " ")(3)
                    For Each vItem In libNrToWebColl
                        If vItem(0) = libNr Then
                            s = "": parentID = Left(Split(line, " ")(4), 32)
                            Do
                                On Error Resume Next: odFolders parentID
                                keyExists = (Err.Number = 0): On Error GoTo 0
                                If Not keyExists Then Exit Do
                                s = odFolders(parentID)(1) & "/" & s
                                parentID = odFolders(parentID)(0)
                            Loop
                            webRoot = vItem(1) & s: Exit For
                        End If
                    Next vItem
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email), locRoot
                Case "AddedScope"
                    parts = Split(line, """")
                    relPath = parts(5): If relPath = " " Then relPath = ""
                    parts = Split(parts(4), " "): siteID = parts(1)
                    webID = parts(2): libID = parts(3): lnkID = parts(4)
                    filename = "ClientPolicy_" & libID & siteID & lnkID & ".ini"
                    On Error Resume Next
                    webRoot = cliPolColl(filename)("DavUrlNamespace") & relPath
                    On Error GoTo 0
                    If webRoot = "" Then
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID And vItem("WebID") = _
                            webID And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace") & relPath
                                Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = "" Then Err.Raise vbErrFileNotFound
                    s = "": parentID = Left(Split(line, " ")(3), 32)
                    Do
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do
                        s = odFolders(parentID)(1) & ps & s
                        parentID = odFolders(parentID)(0)
                    Loop
                    locRoot = mainMount & ps & s
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email), locRoot
                Case Else
                    Exit For
                End Select
            Next line
        Case dirName = "Personal"
            For Each line In Split(b, vbNewLine)
                If line Like "library = *" Then _
                    locRoot = Split(line, """")(3): Exit For
            Next line
            On Error Resume Next
            webRoot = cliPolColl("ClientPolicy.ini")("DavUrlNamespace")
            On Error GoTo 0
            If locRoot = "" Or webRoot = "" Or cid = "" Then GoTo NextFolder
            locToWebColl.Add VBA.Array(locRoot, webRoot & "/" & cid, email), _
                             locRoot
            If Dir(wDir & "GroupFolders.ini") = "" Then GoTo NextFolder
            cid = "": fileNum = FreeFile()
            Open wDir & "GroupFolders.ini" For Binary Access Read As #fileNum
                ReDim b(0 To LOF(fileNum)): Get fileNum, , b
            Close #fileNum: fileNum = 0
            #If Mac Then
                b = StrConv(b, vbUnicode)
            #End If
            For Each line In Split(b, vbNewLine)
                If InStr(line, "BaseUri = ") And cid = "" Then
                    cid = LCase(Mid(line, InStrRev(line, "/") + 1, 16))
                    folderID = Left(line, InStr(line, "_") - 1)
                ElseIf cid <> "" Then
                    locToWebColl.Add VBA.Array(locRoot & ps & odFolders( _
                                     folderID)(1), webRoot & "/" & cid & "/" & _
                                     Mid(line, Len(folderID) + 9), email), _
                                     locRoot & ps & odFolders(folderID)(1)
                    cid = "": folderID = ""
                End If
            Next line
        End Select
NextFolder:
        cid = "": s = "": email = "": Set odFolders = Nothing
    Next dirName

    Dim tmpColl As Collection: Set tmpColl = New Collection
    For Each vItem In locToWebColl
        locRoot = vItem(0): webRoot = vItem(1): email = vItem(2)
       If Right(webRoot, 1) = "/" Then webRoot = Left(webRoot, Len(webRoot) - 1)
        If Right(locRoot, 1) = ps Then locRoot = Left(locRoot, Len(locRoot) - 1)
        tmpColl.Add VBA.Array(locRoot, webRoot, email), locRoot
    Next vItem
    Set locToWebColl = tmpColl

    GetLocalPath = GetLocalPath(path, False, returnAll, pmpo): Exit Function
End Function



