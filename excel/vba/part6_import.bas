
'==============================================================================
' Import vendor request files
'==============================================================================
Private Function FieldKeys() As Variant
    FieldKeys = Array("id", "vendorName", "projectName", "category", "itemName", "spec", "quantity", _
        "requestedDate", "requestedTime", "location", "contactName", "contactPhone", "contactEmail", "note")
End Function

Private Function NormHeader(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, " ", "")
    s = Replace(s, "*", "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr$(160), "")
    NormHeader = s
End Function

Private Function HeaderMatches(ByVal h As String, ByVal fieldKey As String) As Boolean
    Dim names() As String, i As Long
    names = Split(T("hdr." & fieldKey), "|")
    For i = 0 To UBound(names)
        If NormHeader(names(i)) = h Then HeaderMatches = True: Exit Function
    Next i
End Function

Public Sub ImportStart()
    Dim files As Variant, fv As Variant, ws As Worksheet
    EnsureEvents
    files = Application.GetOpenFilename(T("filter.excel"), , T("imp.pick"), , True)
    If Not IsArray(files) Then Exit Sub
    Set ws = SheetOf("import")
    ImportClearRows ws
    Application.ScreenUpdating = False
    For Each fv In files
        ReadVendorFile ws, CStr(fv)
    Next fv
    ImportRecheck
    Application.ScreenUpdating = True
    ws.Activate
    ws.Cells(I_FIRST, 1).Select
End Sub

Public Sub ImportClear()
    EnsureEvents
    If Msg(T("msg.imp.clear"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    ImportClearRows SheetOf("import")
End Sub

Private Sub ImportClearRows(ByVal ws As Worksheet)
    Dim last As Long
    last = LastRow(ws, I_FILE)
    If LastRow(ws, I_KIND) > last Then last = LastRow(ws, I_KIND)
    If last >= I_FIRST Then
        ws.Range(ws.Cells(I_FIRST, 1), ws.Cells(last, I_LAST)).ClearContents
        ws.Range(ws.Cells(I_FIRST, 1), ws.Cells(last, I_LAST)).Interior.ColorIndex = xlNone
    End If
End Sub

Private Function ImportLast(ByVal ws As Worksheet) As Long
    Dim a As Long, b As Long
    a = LastRow(ws, I_FILE): b = LastRow(ws, I_KIND)
    If b > a Then a = b
    If a < I_FIRST Then a = I_FIRST - 1
    ImportLast = a
End Function

Private Sub ReadVendorFile(ByVal ws As Worksheet, ByVal path As String)
    Dim wb As Workbook, sh As Worksheet, hr As Long, best As Long, bestHits As Long, bestSheet As Worksheet
    Dim keys As Variant, map() As Long, bestMap() As Long, r As Long, c As Long, k As Long, h As String, hits As Long
    Dim lastR As Long, lastC As Long, out As Long, v As Variant, anyVal As Boolean, fname As String
    Dim miss As String, req As Variant, item As String, spec As String

    fname = Mid$(path, InStrRev(path, "\") + 1)
    keys = FieldKeys
    ReDim map(0 To UBound(keys)): ReDim bestMap(0 To UBound(keys))
    Application.EnableEvents = False
    On Error GoTo openFail
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error GoTo readFail

    For Each sh In wb.Worksheets
        If sh.Name <> "_meta" Then
            For hr = 1 To 10
                For k = 0 To UBound(keys): map(k) = 0: Next k
                hits = 0
                For c = 1 To 60
                    h = NormHeader(Nz(sh.Cells(hr, c).Value))
                    If h <> "" Then
                        For k = 0 To UBound(keys)
                            If map(k) = 0 Then
                                If HeaderMatches(h, CStr(keys(k))) Then map(k) = c: hits = hits + 1: Exit For
                            End If
                        Next k
                    End If
                Next c
                If hits > bestHits Then
                    bestHits = hits: best = hr: Set bestSheet = sh
                    For k = 0 To UBound(keys): bestMap(k) = map(k): Next k
                End If
            Next hr
        End If
    Next sh

    out = ImportLast(ws) + 1
    If bestHits < 4 Then
        ws.Cells(out, I_KIND).Value = T("imp.kind.error")
        ws.Cells(out, I_MSG).Value = T("imp.noHeader")
        ws.Cells(out, I_FILE).Value = fname
        ws.Cells(out, I_PATH).Value = path
        ws.Range(ws.Cells(out, 1), ws.Cells(out, I_LAST)).Interior.Color = CLR_ERR
        GoTo closeIt
    End If

    req = Array("vendorName", "category", "itemName", "quantity", "requestedDate", "contactName", "contactEmail")
    For k = 0 To UBound(keys)
        If bestMap(k) = 0 And InList(CStr(keys(k)), Join(req, ",")) Then miss = miss & IIf(miss = "", "", ", ") & Split(T("hdr." & keys(k)), "|")(0)
    Next k
    If miss <> "" Then
        ws.Cells(out, I_KIND).Value = T("imp.kind.error")
        ws.Cells(out, I_MSG).Value = F("imp.missingCols", miss)
        ws.Cells(out, I_FILE).Value = fname
        ws.Cells(out, I_PATH).Value = path
        ws.Range(ws.Cells(out, 1), ws.Cells(out, I_LAST)).Interior.Color = CLR_ERR
        GoTo closeIt
    End If

    lastR = bestSheet.UsedRange.Row + bestSheet.UsedRange.Rows.Count - 1
    For r = best + 1 To lastR
        anyVal = False
        For k = 0 To UBound(keys)
            If bestMap(k) > 0 Then
                If Nz(bestSheet.Cells(r, bestMap(k)).Value) <> "" Then anyVal = True: Exit For
            End If
        Next k
        If anyVal Then
            If bestMap(0) > 0 Then
                If Nz(bestSheet.Cells(r, bestMap(0)).Value) = T("imp.example") Then anyVal = False
            End If
        End If
        If anyVal Then
            ws.Cells(out, I_FILE).Value = fname
            ws.Cells(out, I_ROW).Value = r
            ws.Cells(out, I_PATH).Value = path
            ws.Cells(out, I_ID).Value = CellOf(bestSheet, r, bestMap(0))
            ws.Cells(out, I_VENDOR).Value = CellOf(bestSheet, r, bestMap(1))
            ws.Cells(out, I_PROJECT).Value = CellOf(bestSheet, r, bestMap(2))
            ws.Cells(out, I_CATEGORY).Value = CellOf(bestSheet, r, bestMap(3))
            item = Nz(CellOf(bestSheet, r, bestMap(4))): spec = Nz(CellOf(bestSheet, r, bestMap(5)))
            If spec <> "" And InStr(1, item, spec, vbTextCompare) = 0 Then
                If item = "" Then item = spec Else item = item & " (" & spec & ")"
            End If
            ws.Cells(out, I_ITEM).Value = item
            ws.Cells(out, I_QTY).Value = CellOf(bestSheet, r, bestMap(6))
            ws.Cells(out, I_DATE).Value = CellOf(bestSheet, r, bestMap(7))
            ws.Cells(out, I_TIME).Value = CellOf(bestSheet, r, bestMap(8))
            ws.Cells(out, I_LOC).Value = CellOf(bestSheet, r, bestMap(9))
            ws.Cells(out, I_CONTACT).Value = CellOf(bestSheet, r, bestMap(10))
            ws.Cells(out, I_PHONE).Value = "'" & Nz(CellOf(bestSheet, r, bestMap(11)))
            ws.Cells(out, I_EMAIL).Value = CellOf(bestSheet, r, bestMap(12))
            ws.Cells(out, I_NOTE).Value = CellOf(bestSheet, r, bestMap(13))
            out = out + 1
        End If
    Next r
closeIt:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Exit Sub
readFail:
    out = ImportLast(ws) + 1
    ws.Cells(out, I_KIND).Value = T("imp.kind.error")
    ws.Cells(out, I_MSG).Value = F("imp.openFail", Err.Description)
    ws.Cells(out, I_FILE).Value = fname
    ws.Cells(out, I_PATH).Value = path
    Resume closeIt
openFail:
    Application.EnableEvents = True
    out = ImportLast(ws) + 1
    ws.Cells(out, I_KIND).Value = T("imp.kind.error")
    ws.Cells(out, I_MSG).Value = F("imp.openFail", Err.Description)
    ws.Cells(out, I_FILE).Value = fname
    ws.Cells(out, I_PATH).Value = path
    ws.Range(ws.Cells(out, 1), ws.Cells(out, I_LAST)).Interior.Color = CLR_ERR
End Sub

Private Function CellOf(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim v As Variant
    If c <= 0 Then CellOf = Empty: Exit Function
    v = sh.Cells(r, c).Value
    If IsError(v) Then CellOf = Empty Else CellOf = v
End Function

' Validate every preview row; marks cells and sets kind/message. Returns error count.
Public Function ImportRecheck() As Long
    Dim ws As Worksheet, last As Long, r As Long, errs As String, warns As String, info As String, ok As Boolean
    Dim x As Variant, s As String, lr As Long, L As Worksheet, items As Variant, idSeen As New Collection
    Dim nErr As Long, kind As String, i As Long, n As Long, isNew As Boolean
    Dim sDate() As Long, sTime() As String, sLoc() As String, sInsp() As String, sVendor() As String
    Dim sQty() As Double, sId() As String, msgs As Variant, bse As Long, lv As Variant, lastL As Long, st As String

    EnsureEvents
    Set ws = SheetOf("import"): Set L = ListWs
    last = ImportLast(ws)
    If last < I_FIRST Then Exit Function
    On Error Resume Next
    items = ThisWorkbook.Names("set_items").RefersToRange.Value
    On Error GoTo 0

    For r = I_FIRST To last
        If Nz(ws.Cells(r, I_VENDOR).Value) = "" And Nz(ws.Cells(r, I_ITEM).Value) = "" And Nz(ws.Cells(r, I_ROW).Value) = "" Then GoTo nextRow
        errs = "": warns = "": info = ""
        ws.Range(ws.Cells(r, I_ID), ws.Cells(r, I_NOTE)).Interior.ColorIndex = xlNone
        ' required
        CheckReq ws, r, I_VENDOR, "vendorName", errs
        CheckReq ws, r, I_CATEGORY, "category", errs
        CheckReq ws, r, I_ITEM, "itemName", errs
        CheckReq ws, r, I_QTY, "quantity", errs
        CheckReq ws, r, I_DATE, "requestedDate", errs
        CheckReq ws, r, I_CONTACT, "contactName", errs
        CheckReq ws, r, I_EMAIL, "contactEmail", errs
        ' formats
        x = ParseQtyV(ws.Cells(r, I_QTY).Value, ok)
        If ok Then
            If Not IsEmpty(x) Then ws.Cells(r, I_QTY).Value = x
        Else
            AddIssue ws, r, I_QTY, errs, T("imp.e.qty"), CLR_ERR
        End If
        x = ParseDateV(ws.Cells(r, I_DATE).Value, ok)
        If ok Then
            If Not IsEmpty(x) Then
                ws.Cells(r, I_DATE).Value = x
                If CLng(x) < CLng(Date) Then AddIssue ws, r, I_DATE, warns, F("imp.w.past", YMD(Date)), CLR_WARN
            End If
        Else
            AddIssue ws, r, I_DATE, errs, T("imp.e.date"), CLR_ERR
        End If
        s = ParseTimeV(ws.Cells(r, I_TIME).Value, ok)
        If ok Then
            ws.Cells(r, I_TIME).Value = s
        Else
            AddIssue ws, r, I_TIME, errs, T("imp.e.time"), CLR_ERR
        End If
        If Nz(ws.Cells(r, I_EMAIL).Value) <> "" And Not IsEmailOk(Nz(ws.Cells(r, I_EMAIL).Value)) Then AddIssue ws, r, I_EMAIL, errs, T("imp.e.email"), CLR_ERR
        If Nz(ws.Cells(r, I_CATEGORY).Value) <> "" And IsArray(items) Then
            ok = False
            For i = 1 To UBound(items, 1)
                If Nz(items(i, 1)) = Nz(ws.Cells(r, I_CATEGORY).Value) Then ok = True
            Next i
            If Not ok Then AddIssue ws, r, I_CATEGORY, warns, T("imp.w.category"), CLR_WARN
        End If
        ' id matching
        s = UCase$(Nz(ws.Cells(r, I_ID).Value))
        isNew = True
        If s <> "" Then
            ws.Cells(r, I_ID).Value = s
            lr = FindRowById(s)
            If lr = 0 Then
                AddIssue ws, r, I_ID, errs, T("imp.e.noId"), CLR_ERR
            Else
                isNew = False
                info = DiffText(ws, r, L, lr)
                If info = "" Then info = T("imp.noChange")
            End If
            If CHas(idSeen, s) Then
                AddIssue ws, r, I_ID, warns, T("imp.w.dupId"), CLR_WARN
            Else
                idSeen.Add r, s
            End If
        End If
        If errs <> "" Then
            kind = T("imp.kind.error"): nErr = nErr + 1
        ElseIf isNew Then
            kind = T("imp.kind.new")
        Else
            kind = T("imp.kind.update")
        End If
        ws.Cells(r, I_KIND).Value = kind
        ws.Cells(r, I_MSG).Value = Trim$(errs & IIf(warns <> "", " " & warns, "") & IIf(info <> "", " " & info, ""))
        If Nz(ws.Cells(r, I_INCL).Value) = "" Then ws.Cells(r, I_INCL).Value = IIf(errs = "", T("yes"), T("no"))
        If errs <> "" Then
            ws.Cells(r, I_KIND).Interior.Color = CLR_ERR
        ElseIf warns <> "" Then
            ws.Cells(r, I_KIND).Interior.Color = CLR_WARN
        Else
            ws.Cells(r, I_KIND).Interior.ColorIndex = xlNone
        End If
nextRow:
    Next r

    ' conflicts with the list + other preview rows
    lastL = ListLast
    bse = 0
    If lastL >= FIRST Then bse = lastL - FIRST + 1
    n = bse + (last - I_FIRST + 1)
    ReDim sDate(1 To n): ReDim sTime(1 To n): ReDim sLoc(1 To n): ReDim sInsp(1 To n)
    ReDim sVendor(1 To n): ReDim sQty(1 To n): ReDim sId(1 To n)
    If bse > 0 Then
        lv = L.Range(L.Cells(FIRST, 1), L.Cells(lastL, C_LAST)).Value
        For i = 1 To bse
            sId(i) = Nz(lv(i, C_ID))
            st = StatusKey(Nz(lv(i, C_STATUS)))
            If st <> "canceled" And sId(i) <> "" And IsRealDate(lv(i, C_SDATE)) Then sDate(i) = CLng(DateOf(lv(i, C_SDATE)))
            sTime(i) = Nz(lv(i, C_STIME)): sLoc(i) = Nz(lv(i, C_SLOC)): sInsp(i) = Nz(lv(i, C_INSP))
            sVendor(i) = Nz(lv(i, C_VENDOR))
            If IsNumeric(lv(i, C_QTY)) And Not IsEmpty(lv(i, C_QTY)) Then sQty(i) = CDbl(lv(i, C_QTY))
        Next i
    End If
    For r = I_FIRST To last
        i = bse + r - I_FIRST + 1
        sId(i) = Nz(ws.Cells(r, I_ID).Value)
        If sId(i) = "" Then sId(i) = F("imp.rowLabel", r)
        If ws.Cells(r, I_KIND).Value <> T("imp.kind.error") And IsRealDate(ws.Cells(r, I_DATE).Value) Then
            sDate(i) = CLng(DateOf(ws.Cells(r, I_DATE).Value))
            ' an update replaces the existing row's request: drop the old slot
            If Nz(ws.Cells(r, I_ID).Value) <> "" Then
                lr = FindRowById(Nz(ws.Cells(r, I_ID).Value))
                If lr > 0 Then
                    If StatusKey(Nz(L.Cells(lr, C_STATUS).Value)) = "registered" Or StatusKey(Nz(L.Cells(lr, C_STATUS).Value)) = "nego" Then
                        sDate(lr - FIRST + 1) = 0
                    Else
                        sDate(i) = 0
                    End If
                End If
            End If
        End If
        sTime(i) = Nz(ws.Cells(r, I_TIME).Value): sLoc(i) = Nz(ws.Cells(r, I_LOC).Value)
        sVendor(i) = Nz(ws.Cells(r, I_VENDOR).Value)
        If IsNumeric(ws.Cells(r, I_QTY).Value) And Not IsEmpty(ws.Cells(r, I_QTY).Value) Then sQty(i) = CDbl(ws.Cells(r, I_QTY).Value)
    Next r
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)
    For r = I_FIRST To last
        i = bse + r - I_FIRST + 1
        If msgs(i) <> "" Then
            ws.Cells(r, I_MSG).Value = Trim$(Nz(ws.Cells(r, I_MSG).Value) & " " & F("imp.w.conflict", msgs(i)))
            ws.Cells(r, I_DATE).Interior.Color = CLR_CONFLICT
            If ws.Cells(r, I_KIND).Interior.ColorIndex = xlNone Then ws.Cells(r, I_KIND).Interior.Color = CLR_WARN
        End If
    Next r
    ImportRecheck = nErr
End Function

Private Sub CheckReq(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal key As String, ByRef errs As String)
    If Nz(ws.Cells(r, col).Value) = "" Then AddIssue ws, r, col, errs, F("imp.e.req", Split(T("hdr." & key), "|")(0)), CLR_ERR
End Sub

Private Sub AddIssue(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByRef acc As String, ByVal text As String, ByVal clr As Long)
    acc = acc & IIf(acc = "", "", " ") & "[" & text & "]"
    ws.Cells(r, col).Interior.Color = clr
End Sub

Private Function DiffText(ByVal ws As Worksheet, ByVal r As Long, ByVal L As Worksheet, ByVal lr As Long) As String
    Dim pc As Variant, lc As Variant, i As Long, a As String, b As String, s As String
    pc = Array(I_VENDOR, I_PROJECT, I_CATEGORY, I_ITEM, I_QTY, I_DATE, I_TIME, I_LOC, I_CONTACT, I_PHONE, I_EMAIL, I_NOTE)
    lc = Array(C_VENDOR, C_PROJECT, C_CATEGORY, C_ITEM, C_QTY, C_REQDATE, C_REQTIME, C_REQLOC, C_CONTACT, C_PHONE, C_EMAIL, C_NOTE)
    For i = 0 To UBound(pc)
        a = ValText(L.Cells(lr, lc(i)).Value): b = ValText(ws.Cells(r, pc(i)).Value)
        If b <> "" And a <> b Then s = s & IIf(s = "", "", ", ") & ColLabel(CLng(lc(i))) & " " & IIf(a = "", "-", a) & "->" & b
    Next i
    If s <> "" Then DiffText = F("imp.diff", s)
End Function

Public Sub ImportCommit()
    Dim ws As Worksheet, L As Worksheet, last As Long, r As Long, nErr As Long, incErr As Long, inc As Long
    Dim lr As Long, id As String, isNew As Boolean, src As String, newIds As New Collection, nUpd As Long, nSkip As Long
    Dim files As New Collection, fv As Variant, pc As Variant, lc As Variant, i As Long, a As String, b As Variant
    Dim st As String, sk As Long, ok As Boolean, x As Variant, copied As String, failCopy As String
    EnsureEvents
    Set ws = SheetOf("import"): Set L = ListWs
    last = ImportLast(ws)
    If last < I_FIRST Then Msg T("msg.imp.empty"), vbExclamation: Exit Sub
    If ThisWorkbook.ReadOnly Then Msg T("msg.readonly.edit"), vbExclamation: Exit Sub
    nErr = ImportRecheck
    For r = I_FIRST To last
        If Nz(ws.Cells(r, I_INCL).Value) = T("yes") And Nz(ws.Cells(r, I_ROW).Value) <> "" Then
            inc = inc + 1
            If Nz(ws.Cells(r, I_KIND).Value) = T("imp.kind.error") Then incErr = incErr + 1
        End If
    Next r
    If inc = 0 Then Msg T("msg.imp.nothing"), vbExclamation: Exit Sub
    If incErr > 0 Then
        If Msg(F("msg.imp.withErrors", incErr), vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Sub
    ElseIf Msg(F("msg.imp.confirm", inc), vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If

    pc = Array(I_VENDOR, I_PROJECT, I_CATEGORY, I_ITEM, I_QTY, I_DATE, I_TIME, I_LOC, I_CONTACT, I_PHONE, I_EMAIL, I_NOTE)
    lc = Array(C_VENDOR, C_PROJECT, C_CATEGORY, C_ITEM, C_QTY, C_REQDATE, C_REQTIME, C_REQLOC, C_CONTACT, C_PHONE, C_EMAIL, C_NOTE)
    Application.ScreenUpdating = False
    gBusy = True
    Application.EnableEvents = False
    For r = I_FIRST To last
        If Nz(ws.Cells(r, I_INCL).Value) <> T("yes") Or Nz(ws.Cells(r, I_ROW).Value) = "" Then
            If Nz(ws.Cells(r, I_ROW).Value) <> "" Then nSkip = nSkip + 1
            GoTo nextRow
        End If
        ' unparsable values are saved blank
        x = ParseQtyV(ws.Cells(r, I_QTY).Value, ok)
        If Not ok Then ws.Cells(r, I_QTY).Value = Empty
        x = ParseDateV(ws.Cells(r, I_DATE).Value, ok)
        If Not ok Then ws.Cells(r, I_DATE).Value = Empty
        x = ParseTimeV(ws.Cells(r, I_TIME).Value, ok)
        If Not ok Then ws.Cells(r, I_TIME).Value = Empty
        src = F("hist.import", Nz(ws.Cells(r, I_FILE).Value), Nz(ws.Cells(r, I_ROW).Value))
        id = Nz(ws.Cells(r, I_ID).Value)
        lr = 0
        If id <> "" Then lr = FindRowById(id)
        If lr > 0 Then
            st = StatusKey(Nz(L.Cells(lr, C_STATUS).Value))
            For i = 0 To UBound(pc)
                b = ws.Cells(r, pc(i)).Value
                If Nz(b) <> "" And ValText(L.Cells(lr, lc(i)).Value) <> ValText(b) Then
                    AddHist id, ColLabel(CLng(lc(i))), L.Cells(lr, lc(i)).Value, b, src
                    L.Cells(lr, lc(i)).Value = b
                    If st = "registered" Or st = "nego" Then
                        sk = 0
                        If lc(i) = C_REQDATE Then sk = C_SDATE
                        If lc(i) = C_REQTIME Then sk = C_STIME
                        If lc(i) = C_REQLOC Then sk = C_SLOC
                        If sk > 0 Then L.Cells(lr, sk).Value = b
                    End If
                End If
            Next i
            L.Cells(lr, C_UPDATED).Value = Format$(Now, "yyyy-mm-dd hh:nn") & " " & CurrentUser
            nUpd = nUpd + 1
        Else
            id = NextId
            lr = ListLast + 1
            L.Cells(lr, C_ID).Value = id
            L.Cells(lr, C_STATUS).Value = T("st.registered")
            For i = 0 To UBound(pc)
                L.Cells(lr, lc(i)).Value = ws.Cells(r, pc(i)).Value
            Next i
            L.Cells(lr, C_SDATE).Value = L.Cells(lr, C_REQDATE).Value
            L.Cells(lr, C_STIME).Value = L.Cells(lr, C_REQTIME).Value
            L.Cells(lr, C_SLOC).Value = L.Cells(lr, C_REQLOC).Value
            L.Cells(lr, C_CREATED_AT).Value = Now
            L.Cells(lr, C_CREATED_BY).Value = CurrentUser
            L.Cells(lr, C_STATUS_AT).Value = Now
            L.Cells(lr, C_SRCFILE).Value = Nz(ws.Cells(r, I_FILE).Value)
            L.Cells(lr, C_SRCROW).Value = ws.Cells(r, I_ROW).Value
            SaveOriginal id, ws, r
            AddHist id, T("f.created"), "", id, src
            newIds.Add id
        End If
        UpsertVendor ws, r
        If Not CHas(files, Nz(ws.Cells(r, I_PATH).Value)) Then files.Add Nz(ws.Cells(r, I_PATH).Value), Nz(ws.Cells(r, I_PATH).Value)
nextRow:
    Next r
    Application.EnableEvents = True
    gBusy = False

    ' keep a copy of each original file
    For Each fv In files
        If CopyOriginal(CStr(fv)) Then copied = copied & "1" Else failCopy = failCopy & vbCrLf & "- " & Mid$(fv, InStrRev(fv, "\") + 1)
    Next fv
    ImportClearRows ws
    RecalcAll
    RefreshHome
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly And ThisWorkbook.Path <> "" Then ThisWorkbook.Save
    On Error GoTo 0
    Msg F("msg.imp.done", newIds.Count, nUpd, nSkip, JoinCol(newIds, ", ")) & IIf(failCopy <> "", vbCrLf & vbCrLf & T("msg.imp.copyFail") & failCopy, "")
    If newIds.Count > 0 Then
        If Msg(T("msg.imp.mailAsk"), vbYesNo + vbQuestion) = vbYes Then
            MailDraftForIds newIds
        Else
            ListWs.Activate
        End If
    Else
        ListWs.Activate
    End If
End Sub

Private Sub SaveOriginal(ByVal id As String, ByVal ws As Worksheet, ByVal r As Long)
    Dim o As Worksheet, orow As Long
    Set o = SheetOf("orig")
    orow = LastRow(o, 1) + 1
    o.Cells(orow, 1).Value = id
    o.Cells(orow, 2).Value = Now
    o.Cells(orow, 3).Value = ws.Cells(r, I_VENDOR).Value
    o.Cells(orow, 4).Value = ws.Cells(r, I_PROJECT).Value
    o.Cells(orow, 5).Value = ws.Cells(r, I_CATEGORY).Value
    o.Cells(orow, 6).Value = ws.Cells(r, I_ITEM).Value
    o.Cells(orow, 7).Value = ws.Cells(r, I_QTY).Value
    o.Cells(orow, 8).Value = ws.Cells(r, I_DATE).Value
    o.Cells(orow, 9).Value = ws.Cells(r, I_TIME).Value
    o.Cells(orow, 10).Value = ws.Cells(r, I_LOC).Value
    o.Cells(orow, 11).Value = ws.Cells(r, I_CONTACT).Value
    o.Cells(orow, 12).Value = ws.Cells(r, I_PHONE).Value
    o.Cells(orow, 13).Value = ws.Cells(r, I_EMAIL).Value
    o.Cells(orow, 14).Value = ws.Cells(r, I_NOTE).Value
    o.Cells(orow, 15).Value = ws.Cells(r, I_FILE).Value & " " & ws.Cells(r, I_ROW).Value
End Sub

Private Sub UpsertVendor(ByVal ws As Worksheet, ByVal r As Long)
    Dim v As Worksheet, i As Long, nm As String, row As Long
    nm = Nz(ws.Cells(r, I_VENDOR).Value)
    If nm = "" Then Exit Sub
    Set v = SheetOf("vendor")
    For i = 2 To LastRow(v, 1)
        If Nz(v.Cells(i, 1).Value) = nm Then row = i: Exit For
    Next i
    If row = 0 Then row = LastRow(v, 1) + 1: v.Cells(row, 1).Value = nm
    If Nz(ws.Cells(r, I_CONTACT).Value) <> "" Then v.Cells(row, 2).Value = ws.Cells(r, I_CONTACT).Value
    If Nz(ws.Cells(r, I_PHONE).Value) <> "" Then v.Cells(row, 3).Value = "'" & Nz(ws.Cells(r, I_PHONE).Value)
    If Nz(ws.Cells(r, I_EMAIL).Value) <> "" Then v.Cells(row, 4).Value = ws.Cells(r, I_EMAIL).Value
    v.Cells(row, 6).Value = Date
End Sub

Private Function CopyOriginal(ByVal path As String) As Boolean
    Dim dirp As String, nm As String, dest As String
    If BaseDir = "" Or path = "" Then Exit Function
    dirp = BaseDir & "\" & T("dir.orig") & "\" & Format$(Date, "yyyy-mm")
    If Not EnsureDir(dirp) Then Exit Function
    nm = Mid$(path, InStrRev(path, "\") + 1)
    dest = dirp & "\" & Format$(Now, "yyyymmdd_hhnnss") & "_" & nm
    On Error GoTo fail
    FileCopy path, dest
    CopyOriginal = True
    Exit Function
fail:
    CopyOriginal = False
End Function
