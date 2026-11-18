
'==============================================================================
' Startup / events
'==============================================================================
Public Sub Auto_Open()
    On Error Resume Next
    EnsureEvents
    ReloadTexts
    EnsureSetup False
    If ThisWorkbook.ReadOnly Then Msg T("msg.readonly"), vbExclamation
    RefreshAll
    SheetOf("home").Activate
    StartupPopup
End Sub

Public Sub EnsureEvents()
    If gEvents Is Nothing Then
        Set gEvents = New clsAppEvents
        Set gEvents.App = Application
        On Error Resume Next
        ApplyProtection
        On Error GoTo 0
    End If
End Sub

Public Sub OnSheetChange(ByVal Sh As Object, ByVal Target As Range)
    If gBusy Then Exit Sub
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    On Error GoTo done
    If Sh.Name = T("sh.list") Then
        ListChanged Target
    ElseIf Sh.Name = T("sh.cal") Then
        If Not Intersect(Target, Sh.Range("C2,F2")) Is Nothing Then DrawCalendar
    ElseIf Sh.Name = T("sh.detail") Then
        If Not Intersect(Target, Sh.Range("C2")) Is Nothing Then ShowDetail Nz(Sh.Range("C2").Value), False
    End If
done:
    If Err.Number <> 0 Then Msg F("msg.error", Err.Description), vbExclamation
    gBusy = False
    Application.EnableEvents = True
End Sub

Public Sub OnSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If gBusy Then Exit Sub
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    If Sh.Name <> T("sh.list") Then Exit Sub
    CaptureOld Target
End Sub

Private Sub CaptureOld(ByVal Target As Range)
    Dim a As Range
    On Error GoTo clear
    Set a = Target.Areas(1)
    If a.Cells.CountLarge > 3000 Then GoTo clear
    gOldAddr = a.Address
    gOldVals = a.Value
    Exit Sub
clear:
    gOldAddr = ""
    gOldVals = Empty
End Sub

Private Function OldValueOf(ByVal c As Range) As Variant
    Dim o As Range
    If gOldAddr = "" Then Exit Function
    Set o = c.Worksheet.Range(gOldAddr)
    If Intersect(o, c) Is Nothing Then Exit Function
    If o.Cells.CountLarge = 1 Then
        OldValueOf = gOldVals
    Else
        OldValueOf = gOldVals(c.Row - o.Row + 1, c.Column - o.Column + 1)
    End If
End Function

Public Sub OnDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim s As String, d As Variant
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    On Error GoTo done
    If Sh.Name = T("sh.list") Then
        If Target.Column = C_ID And Target.Row >= FIRST Then
            s = Nz(Target.Value)
            If s <> "" Then Cancel = True: ShowDetail s, True
        End If
    ElseIf Sh.Name = T("sh.home") Then
        s = Nz(Target.Cells(1, 1).Value)
        If FindRowById(s) > 0 Then Cancel = True: ShowDetail s, True
    ElseIf Sh.Name = T("sh.cal") Then
        d = CalendarDateAt(Target)
        If Not IsEmpty(d) Then Cancel = True: FilterListByDate CDate(d)
    End If
done:
End Sub

Public Sub OnAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If gSavingCopy Then Exit Sub
    If Not Success Then Exit Sub
    If Not (Wb Is ThisWorkbook) Then Exit Sub
    MakeBackup
End Sub

'==============================================================================
' List edits
'==============================================================================
Private Sub ListChanged(ByVal Target As Range)
    Dim ws As Worksheet, rng As Range, c As Range
    Set ws = ListWs
    Set rng = Intersect(Target, ws.Range(ws.Cells(FIRST, 1), ws.Cells(MAXROW, C_LAST)))
    If rng Is Nothing Then Exit Sub
    If ThisWorkbook.ReadOnly And Not gRoWarned Then
        gRoWarned = True
        Msg T("msg.readonly.edit"), vbExclamation
    End If
    gBusy = True
    Application.EnableEvents = False
    If rng.Cells.CountLarge > 3000 Then
        Msg T("msg.bigpaste"), vbExclamation
    Else
        For Each c In rng.Cells
            HandleCell ws, c.Row, c.Column, OldValueOf(c), c.Value
        Next c
    End If
    Application.EnableEvents = True
    gBusy = False
    RecalcAll
    CaptureOld Selection
End Sub

Private Function IsAutoCol(ByVal col As Long) As Boolean
    IsAutoCol = (col = C_ID Or col = C_TODO Or col = C_CONFLICT Or col = C_NOTICED Or col >= C_CONF_AT)
End Function

Private Sub HandleCell(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal oldV As Variant, ByVal newV As Variant)
    Dim id As String, ok As Boolean, reason As String, x As Variant, s As String
    If ValText(oldV) = ValText(newV) Then Exit Sub
    If IsAutoCol(col) Then
        ws.Cells(r, col).Value = oldV
        Exit Sub
    End If
    id = Nz(ws.Cells(r, C_ID).Value)
    If id = "" Then
        If Nz(newV) = "" Then Exit Sub
        id = RegisterManualRow(ws, r)
    End If

    Select Case col
        Case C_STATUS
            ok = HandleStatus(ws, r, id, Nz(oldV), Nz(newV))
            If Not ok Then ws.Cells(r, col).Value = oldV
            GoTo stamp
        Case C_RESULT
            ok = HandleResult(ws, r, id, Nz(oldV), Nz(newV))
            If Not ok Then ws.Cells(r, col).Value = oldV
            GoTo stamp
        Case C_SDATE, C_STIME, C_SLOC, C_INSP
            ok = HandleSchedule(ws, r, id, col, oldV, newV)
            If Not ok Then ws.Cells(r, col).Value = oldV
            GoTo stamp
        Case C_REQTIME
            s = ParseTimeV(newV, ok)
            If Not ok Then
                Msg T("msg.time.invalid"), vbExclamation
                ws.Cells(r, col).Value = oldV
                Exit Sub
            End If
            ws.Cells(r, col).Value = s
            newV = s
        Case C_REQDATE, C_RDATE
            x = ParseDateV(newV, ok)
            If Not ok Then
                Msg T("msg.date.invalid"), vbExclamation
                ws.Cells(r, col).Value = oldV
                Exit Sub
            End If
            ws.Cells(r, col).Value = x
            newV = x
        Case C_QTY
            x = ParseQtyV(newV, ok)
            If Not ok Then
                Msg T("msg.qty.invalid"), vbExclamation
                ws.Cells(r, col).Value = oldV
                Exit Sub
            End If
            ws.Cells(r, col).Value = x
            newV = x
        Case C_EMAIL
            If Nz(newV) <> "" And Not IsEmailOk(Nz(newV)) Then Msg T("msg.email.invalid"), vbExclamation
    End Select
    AddHist id, ColLabel(col), oldV, newV, ""
    ok = True
stamp:
    If ok Then ws.Cells(r, C_UPDATED).Value = Format$(Now, "yyyy-mm-dd hh:nn") & " " & CurrentUser
End Sub

Private Function RegisterManualRow(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim id As String
    id = NextId
    ws.Cells(r, C_ID).Value = id
    If Nz(ws.Cells(r, C_STATUS).Value) = "" Then ws.Cells(r, C_STATUS).Value = T("st.registered")
    ws.Cells(r, C_CREATED_AT).Value = Now
    ws.Cells(r, C_CREATED_BY).Value = CurrentUser
    ws.Cells(r, C_STATUS_AT).Value = Now
    ws.Cells(r, C_SRCFILE).Value = T("w.manual")
    AddHist id, T("f.created"), "", id, T("w.manual")
    RegisterManualRow = id
End Function

Private Function HandleStatus(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByVal oldS As String, ByVal newS As String) As Boolean
    Dim fromK As String, toK As String, reason As String, res As String
    fromK = StatusKey(oldS): toK = StatusKey(newS)
    If toK = "" Then
        Msg T("msg.status.invalid"), vbExclamation
        Exit Function
    End If
    If fromK = toK Then HandleStatus = True: Exit Function
    If Not InList(toK, AllowedNext(fromK)) Then
        Msg F("msg.status.notallowed", IIf(oldS = "", "-", oldS), newS), vbExclamation
        Exit Function
    End If
    Select Case toK
        Case "confirmed"
            If Not ConfirmChecks(ws, r, id, reason) Then Exit Function
            ws.Cells(r, C_CONF_AT).Value = Now
            ws.Cells(r, C_CONF_BY).Value = CurrentUser
        Case "canceled"
            reason = AskReason(F("ask.cancel", id))
            If reason = "" Then Exit Function
        Case "changed"
            reason = AskReason(F("ask.change", id))
            If reason = "" Then Exit Function
        Case "done"
            If Nz(ws.Cells(r, C_RESULT).Value) = "" Then
                res = Trim$(InputBox(F("ask.result", id), T("app.title"), "1"))
                If res = "1" Then
                    res = T("res.pass")
                ElseIf res = "2" Then
                    res = T("res.fail")
                Else
                    Exit Function
                End If
                If Not ApplyResult(ws, r, id, "", res) Then Exit Function
                ws.Cells(r, C_RESULT).Value = res
            End If
    End Select
    ws.Cells(r, C_STATUS).Value = newS
    ws.Cells(r, C_STATUS_AT).Value = Now
    AddHist id, T("f.status"), oldS, newS, reason
    HandleStatus = True
End Function

' Required fields + conflict check before confirming. reason gets the force reason.
Private Function ConfirmChecks(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByRef reason As String) As Boolean
    Dim miss As String, tv As String, ok As Boolean, c As String
    If Not IsRealDate(ws.Cells(r, C_SDATE).Value) Then miss = miss & vbLf & "- " & ColLabel(C_SDATE)
    tv = ParseTimeV(ws.Cells(r, C_STIME).Value, ok)
    If Not ok Or tv = "" Then
        miss = miss & vbLf & "- " & ColLabel(C_STIME)
    Else
        ws.Cells(r, C_STIME).Value = tv
    End If
    If Nz(ws.Cells(r, C_SLOC).Value) = "" Then miss = miss & vbLf & "- " & ColLabel(C_SLOC)
    If Nz(ws.Cells(r, C_INSP).Value) = "" Then
        miss = miss & vbLf & "- " & ColLabel(C_INSP)
    ElseIf Not ActiveInspector(Nz(ws.Cells(r, C_INSP).Value)) Then
        miss = miss & vbLf & "- " & F("msg.insp.inactive", Nz(ws.Cells(r, C_INSP).Value))
    End If
    If miss <> "" Then
        Msg F("msg.confirm.missing", id) & miss, vbExclamation
        Exit Function
    End If
    c = ConflictsForRow(r)
    If c <> "" Then
        If Not SetYes("set_allowForce") Then
            Msg F("msg.conflict.block", id, Replace(c, " / ", vbLf & "- ")), vbExclamation
            Exit Function
        End If
        If Msg(F("msg.conflict.ask", id, Replace(c, " / ", vbLf & "- ")), vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Function
        reason = AskReason(T("ask.force"))
        If reason = "" Then Exit Function
        AddHist id, T("f.force"), "", c, reason
    End If
    ConfirmChecks = True
End Function

' Conflicts the given list row would have with its current values
Public Function ConflictsForRow(ByVal row As Long) As String
    Dim ws As Worksheet, last As Long, n As Long, v As Variant, i As Long
    Dim sDate() As Long, sTime() As String, sLoc() As String, sInsp() As String, sVendor() As String
    Dim sQty() As Double, sId() As String, msgs As Variant, st As String, ok As Boolean
    Set ws = ListWs
    last = ListLast
    n = last - FIRST + 1
    If n < 1 Then Exit Function
    v = ws.Range(ws.Cells(FIRST, 1), ws.Cells(last, C_LAST)).Value
    ReDim sDate(1 To n): ReDim sTime(1 To n): ReDim sLoc(1 To n): ReDim sInsp(1 To n)
    ReDim sVendor(1 To n): ReDim sQty(1 To n): ReDim sId(1 To n)
    For i = 1 To n
        sId(i) = Nz(v(i, C_ID))
        st = StatusKey(Nz(v(i, C_STATUS)))
        If (st <> "canceled" Or FIRST + i - 1 = row) And sId(i) <> "" And IsRealDate(v(i, C_SDATE)) Then sDate(i) = CLng(DateOf(v(i, C_SDATE)))
        sTime(i) = ParseTimeV(v(i, C_STIME), ok): sLoc(i) = Nz(v(i, C_SLOC)): sInsp(i) = Nz(v(i, C_INSP))
        sVendor(i) = Nz(v(i, C_VENDOR))
        If IsNumeric(v(i, C_QTY)) And Not IsEmpty(v(i, C_QTY)) Then sQty(i) = CDbl(v(i, C_QTY))
    Next i
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)
    ConflictsForRow = msgs(row - FIRST + 1)
End Function

Private Function HandleSchedule(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByVal col As Long, _
        ByVal oldV As Variant, ByVal newV As Variant) As Boolean
    Dim st As String, reason As String, ok As Boolean, x As Variant
    st = StatusKey(Nz(ws.Cells(r, C_STATUS).Value))
    If st = "done" Or st = "canceled" Then
        Msg F("msg.schedule.locked", id, Nz(ws.Cells(r, C_STATUS).Value)), vbExclamation
        Exit Function
    End If
    If col = C_STIME Then
        x = ParseTimeV(newV, ok)
        If Not ok Then Msg T("msg.time.invalid"), vbExclamation: Exit Function
        ws.Cells(r, col).Value = x
        newV = x
    ElseIf col = C_SDATE Then
        x = ParseDateV(newV, ok)
        If Not ok Then Msg T("msg.date.invalid"), vbExclamation: Exit Function
        ws.Cells(r, col).Value = x
        newV = x
    ElseIf col = C_INSP Then
        If Nz(newV) <> "" And Not ActiveInspector(Nz(newV)) Then Msg F("msg.insp.inactive", Nz(newV)), vbExclamation
    End If
    If ValText(oldV) = ValText(newV) Then HandleSchedule = True: Exit Function
    If st = "confirmed" Then
        reason = AskReason(F("ask.change", id))
        If reason = "" Then Exit Function
        ws.Cells(r, C_STATUS).Value = T("st.changed")
        ws.Cells(r, C_STATUS_AT).Value = Now
        AddHist id, T("f.status"), T("st.confirmed"), T("st.changed"), reason
    End If
    AddHist id, ColLabel(col), oldV, newV, reason
    HandleSchedule = True
End Function

Private Function HandleResult(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByVal oldR As String, ByVal newR As String) As Boolean
    Dim st As String
    st = StatusKey(Nz(ws.Cells(r, C_STATUS).Value))
    If newR = "" Then
        If st = "done" Then
            If Msg(F("msg.result.clear", id), vbYesNo + vbQuestion) <> vbYes Then Exit Function
            ws.Cells(r, C_STATUS).Value = T("st.confirmed")
            ws.Cells(r, C_STATUS_AT).Value = Now
            AddHist id, T("f.status"), T("st.done"), T("st.confirmed"), T("w.resultCleared")
        End If
        AddHist id, ColLabel(C_RESULT), oldR, "", ""
        HandleResult = True
        Exit Function
    End If
    If newR <> T("res.pass") And newR <> T("res.fail") Then
        Msg T("msg.result.invalid"), vbExclamation
        Exit Function
    End If
    If st <> "confirmed" And st <> "changed" And st <> "done" Then
        Msg F("msg.result.notconfirmed", id), vbExclamation
        Exit Function
    End If
    If Not ApplyResult(ws, r, id, oldR, newR) Then Exit Function
    If st <> "done" Then
        AddHist id, T("f.status"), Nz(ws.Cells(r, C_STATUS).Value), T("st.done"), ""
        ws.Cells(r, C_STATUS).Value = T("st.done")
        ws.Cells(r, C_STATUS_AT).Value = Now
    End If
    HandleResult = True
End Function

Private Function ApplyResult(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByVal oldR As String, ByVal newR As String) As Boolean
    Dim reason As String
    If newR = T("res.fail") And Nz(ws.Cells(r, C_FAIL).Value) = "" Then
        reason = AskReason(F("ask.fail", id))
        If reason = "" Then Exit Function
        ws.Cells(r, C_FAIL).Value = reason
    End If
    If Not IsRealDate(ws.Cells(r, C_RDATE).Value) Then
        If IsRealDate(ws.Cells(r, C_SDATE).Value) Then
            ws.Cells(r, C_RDATE).Value = DateOf(ws.Cells(r, C_SDATE).Value)
        Else
            ws.Cells(r, C_RDATE).Value = Date
        End If
    End If
    AddHist id, ColLabel(C_RESULT), oldR, newR, Nz(ws.Cells(r, C_FAIL).Value)
    ApplyResult = True
End Function
