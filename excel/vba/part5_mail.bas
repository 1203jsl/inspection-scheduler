
'==============================================================================
' Mail drafts (no sending; drafts are copied into web mail)
'==============================================================================
Private Function TplText(ByVal key As String, ByVal part As Long) As String
    Dim ws As Worksheet, i As Long
    Set ws = SheetOf("tpl")
    For i = 2 To LastRow(ws, 1)
        If Nz(ws.Cells(i, 1).Value) = key Then TplText = Replace(CStr(ws.Cells(i, 2 + part).Value), vbCr, ""): Exit Function
    Next i
End Function

Private Function TplLabel(ByVal key As String) As String
    Dim ws As Worksheet, i As Long
    Set ws = SheetOf("tpl")
    For i = 2 To LastRow(ws, 1)
        If Nz(ws.Cells(i, 1).Value) = key Then TplLabel = Nz(ws.Cells(i, 2).Value): Exit Function
    Next i
    TplLabel = key
End Function

Private Function Fill(ByVal s As String, ByVal names As Collection, ByVal vals As Collection) As String
    Dim i As Long
    For i = 1 To names.Count
        s = Replace(s, "{" & names(i) & "}", vals(i))
    Next i
    Fill = s
End Function

Private Sub PutVar(ByVal names As Collection, ByVal vals As Collection, ByVal phKey As String, ByVal value As String)
    names.Add T("ph." & phKey)
    If value = "" Then value = "-"
    vals.Add value
End Sub

Private Function LastReason(ByVal id As String, ByVal toStatus As String) As String
    Dim ws As Worksheet, i As Long
    Set ws = SheetOf("hist")
    For i = LastRow(ws, 1) To 2 Step -1
        If UCase$(Nz(ws.Cells(i, 3).Value)) = UCase$(id) And Nz(ws.Cells(i, 4).Value) = T("f.status") And Nz(ws.Cells(i, 6).Value) = toStatus Then
            LastReason = Nz(ws.Cells(i, 7).Value)
            Exit Function
        End If
    Next i
End Function

Private Function ItemLine(ByVal L As Worksheet, ByVal r As Long, ByVal kind As String) As String
    Dim s As String, sep As String
    sep = T("sep")
    s = "- " & Nz(L.Cells(r, C_ID).Value)
    If Nz(L.Cells(r, C_PROJECT).Value) <> "" Then s = s & sep & T("w.project") & ": " & Nz(L.Cells(r, C_PROJECT).Value)
    s = s & sep & Nz(L.Cells(r, C_ITEM).Value) & sep & T("w.qty") & " " & Nz(L.Cells(r, C_QTY).Value)
    If kind = "confirmed" Or kind = "changed" Then
        s = s & vbLf & "  " & T("w.when") & ": " & DateKo(L.Cells(r, C_SDATE).Value) & " " & Nz(L.Cells(r, C_STIME).Value) & _
            sep & T("w.place") & ": " & Nz(L.Cells(r, C_SLOC).Value)
        If Nz(L.Cells(r, C_INSP).Value) <> "" Then s = s & sep & T("w.insp") & ": " & Nz(L.Cells(r, C_INSP).Value)
    ElseIf kind = "registered" Then
        s = s & vbLf & "  " & T("w.wish") & ": " & DateKo(L.Cells(r, C_REQDATE).Value) & " " & Nz(L.Cells(r, C_REQTIME).Value)
    End If
    ItemLine = s
End Function

Private Function KindForRow(ByVal L As Worksheet, ByVal r As Long) As String
    Select Case StatusKey(Nz(L.Cells(r, C_STATUS).Value))
        Case "registered", "nego", "": KindForRow = "registered"
        Case "confirmed": KindForRow = "confirmed"
        Case "changed": KindForRow = "changed"
        Case "canceled": KindForRow = "canceled"
        Case "done"
            If Nz(L.Cells(r, C_RESULT).Value) = T("res.fail") Then KindForRow = "fail" Else KindForRow = "pass"
    End Select
End Function

' Selected list rows (visible only) -> mail drafts
Public Sub MailForSelection()
    Dim ids As New Collection, c As Range, L As Worksheet, s As String, area As Range
    EnsureEvents
    Set L = ListWs
    If ActiveSheet.Name = T("sh.list") Then
        For Each area In Selection.Areas
            For Each c In Intersect(area.EntireRow, L.Columns(C_ID)).Cells
                If c.Row >= FIRST And Not c.EntireRow.Hidden Then
                    s = Nz(c.Value)
                    If s <> "" And Not CHas(ids, s) Then ids.Add s, s
                End If
                If c.Row > ListLast Then Exit For
            Next c
        Next area
    Else
        s = ActiveId
        If s <> "" Then ids.Add s, s
    End If
    If ids.Count = 0 Then Msg T("msg.pick.row"), vbExclamation: Exit Sub
    MailDraftForIds ids
End Sub

Public Sub MailDraftForIds(ByVal ids As Collection)
    Dim L As Worksheet, groups As New Collection, gkeys As New Collection, id As Variant, r As Long
    Dim kind As String, gk As String, grp As Collection, g As Variant, n As Long
    Set L = ListWs
    For Each id In ids
        r = FindRowById(CStr(id))
        If r > 0 Then
            kind = KindForRow(L, r)
            If kind = "registered" Or kind = "confirmed" Then
                gk = kind & "|" & Nz(L.Cells(r, C_VENDOR).Value) & "|" & LCase$(Nz(L.Cells(r, C_EMAIL).Value))
            Else
                gk = kind & "|" & CStr(id)
            End If
            If Not CHas(groups, gk) Then
                Set grp = New Collection
                groups.Add grp, gk
                gkeys.Add gk
            End If
            groups(gk).Add r
        End If
    Next id
    MailClearDrafts False
    For Each g In gkeys
        Set grp = groups(CStr(g))
        kind = Split(CStr(g), "|")(0)
        WriteVendorDraft L, grp, kind
        n = n + 1
    Next g
    SheetOf("mail").Activate
    SheetOf("mail").Cells(M_FIRST, M_BODY).Select
    Msg F("msg.mail.drafted", n)
End Sub

Private Sub WriteVendorDraft(ByVal L As Worksheet, ByVal rows As Collection, ByVal kind As String)
    Dim names As New Collection, vals As New Collection, r0 As Long, r As Variant, list As String
    Dim idsTxt As String, tos As New Collection, e As String, toTxt As String, res As String
    r0 = rows(1)
    For Each r In rows
        If list <> "" Then list = list & vbLf
        list = list & ItemLine(L, CLng(r), kind)
        If idsTxt <> "" Then idsTxt = idsTxt & ", "
        idsTxt = idsTxt & Nz(L.Cells(r, C_ID).Value)
        e = Nz(L.Cells(r, C_EMAIL).Value)
        If e <> "" And Not CHas(tos, LCase$(e)) Then tos.Add e, LCase$(e)
    Next r
    toTxt = JoinCol(tos, "; ")
    res = Nz(L.Cells(r0, C_RESULT).Value)
    PutVar names, vals, "vendor", Nz(L.Cells(r0, C_VENDOR).Value)
    PutVar names, vals, "project", Nz(L.Cells(r0, C_PROJECT).Value)
    PutVar names, vals, "id", idsTxt
    PutVar names, vals, "item", Nz(L.Cells(r0, C_ITEM).Value)
    PutVar names, vals, "qty", Nz(L.Cells(r0, C_QTY).Value)
    PutVar names, vals, "category", Nz(L.Cells(r0, C_CATEGORY).Value)
    PutVar names, vals, "wishDate", DateKo(L.Cells(r0, C_REQDATE).Value)
    PutVar names, vals, "wishTime", Nz(L.Cells(r0, C_REQTIME).Value)
    If IsRealDate(L.Cells(r0, C_RDATE).Value) And (kind = "pass" Or kind = "fail") Then
        PutVar names, vals, "date", DateKo(L.Cells(r0, C_RDATE).Value)
    Else
        PutVar names, vals, "date", DateKo(L.Cells(r0, C_SDATE).Value)
    End If
    PutVar names, vals, "time", Nz(L.Cells(r0, C_STIME).Value)
    PutVar names, vals, "place", Nz(L.Cells(r0, C_SLOC).Value)
    PutVar names, vals, "insp", Nz(L.Cells(r0, C_INSP).Value)
    PutVar names, vals, "contact", Nz(L.Cells(r0, C_CONTACT).Value)
    PutVar names, vals, "list", list
    PutVar names, vals, "count", CStr(rows.Count)
    PutVar names, vals, "confirmedOn", DateKo(Date)
    PutVar names, vals, "result", res
    PutVar names, vals, "failReason", Nz(L.Cells(r0, C_FAIL).Value)
    PutVar names, vals, "changeReason", LastReason(Nz(L.Cells(r0, C_ID).Value), T("st.changed"))
    PutVar names, vals, "cancelReason", LastReason(Nz(L.Cells(r0, C_ID).Value), T("st.canceled"))
    PutVar names, vals, "sender", Nz(SetV("set_sender"))
    AddDraft TplLabel(kind), toTxt, Nz(SetV("set_cc")), Fill(TplText(kind, 1), names, vals), Fill(TplText(kind, 2), names, vals), idsTxt, kind
End Sub

Private Sub AddDraft(ByVal label As String, ByVal toTxt As String, ByVal cc As String, ByVal subj As String, _
        ByVal body As String, ByVal ids As String, ByVal key As String)
    Dim ws As Worksheet, r As Long
    Set ws = SheetOf("mail")
    r = LastRow(ws, M_SUBJ) + 1
    If r < M_FIRST Then r = M_FIRST
    ws.Cells(r, M_NO).Value = r - M_FIRST + 1
    ws.Cells(r, M_KIND).Value = label
    ws.Cells(r, M_TO).Value = toTxt
    ws.Cells(r, M_CC).Value = cc
    ws.Cells(r, M_SUBJ).Value = subj
    ws.Cells(r, M_BODY).Value = body
    ws.Cells(r, M_IDS).Value = ids
    ws.Cells(r, M_SENT).Value = ""
    ws.Cells(r, M_KEY).Value = key
    ws.Rows(r).RowHeight = 180
End Sub

Public Sub MailClearDrafts(Optional ByVal ask As Boolean = True)
    Dim ws As Worksheet, last As Long
    EnsureEvents
    Set ws = SheetOf("mail")
    last = LastRow(ws, M_SUBJ)
    If last < M_FIRST Then Exit Sub
    If ask Then
        If Msg(T("msg.mail.clear"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    ws.Range(ws.Cells(M_FIRST, 1), ws.Cells(last, M_KEY)).ClearContents
    ws.Range(ws.Cells(M_FIRST, 1), ws.Cells(last, M_KEY)).Rows.RowHeight = ws.StandardHeight
End Sub

Private Function DraftRow() As Long
    If ActiveSheet.Name <> T("sh.mail") Then Msg T("msg.mail.pick"), vbExclamation: Exit Function
    If ActiveCell.Row < M_FIRST Or Nz(ActiveSheet.Cells(ActiveCell.Row, M_SUBJ).Value) = "" Then
        Msg T("msg.mail.pick"), vbExclamation
        Exit Function
    End If
    DraftRow = ActiveCell.Row
End Function

Public Sub MailCopyBody(): MailCopy M_BODY: End Sub
Public Sub MailCopySubject(): MailCopy M_SUBJ: End Sub
Public Sub MailCopyTo(): MailCopy M_TO: End Sub

Private Sub MailCopy(ByVal col As Long)
    Dim r As Long, s As String
    EnsureEvents
    r = DraftRow
    If r = 0 Then Exit Sub
    s = CStr(ActiveSheet.Cells(r, col).Value)
    If col = M_BODY Then s = Replace(s, vbLf, vbCrLf)
    If ToClipboard(s) Then
        Application.StatusBar = F("msg.copied", Nz(ActiveSheet.Cells(M_HDR, col).Value))
        Msg F("msg.copied", Nz(ActiveSheet.Cells(M_HDR, col).Value))
    Else
        ActiveSheet.Cells(r, col).Select
        Msg T("msg.copy.manual"), vbExclamation
    End If
End Sub

Public Function ToClipboard(ByVal s As String) As Boolean
    If ClipHtml(s) Then ToClipboard = True: Exit Function
    ToClipboard = ClipForms(s)
End Function

Private Function ClipHtml(ByVal s As String) As Boolean
    Dim o As Object
    On Error GoTo fail
    Set o = CreateObject("htmlfile")
    o.parentWindow.clipboardData.SetData "text", s
    ClipHtml = True
fail:
End Function

Private Function ClipForms(ByVal s As String) As Boolean
    Dim o As Object
    On Error GoTo fail
    Set o = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    o.SetText s
    o.PutInClipboard
    ClipForms = True
fail:
End Function

Public Sub OpenWebmail()
    Dim url As String
    EnsureEvents
    url = Nz(SetV("set_webmail"))
    If url = "" Then Msg T("msg.webmail.none"), vbExclamation: Exit Sub
    On Error GoTo fail
    ThisWorkbook.FollowHyperlink url
    Exit Sub
fail:
    Msg F("msg.open.fail", url), vbExclamation
End Sub

' Record selected drafts as sent
Public Sub MailMarkSent()
    Dim ws As Worksheet, lg As Worksheet, c As Range, r As Long, done As New Collection, n As Long
    Dim key As String, ids() As String, i As Long, L As Worksheet, lr As Long
    EnsureEvents
    If ActiveSheet.Name <> T("sh.mail") Then Msg T("msg.mail.pick"), vbExclamation: Exit Sub
    Set ws = ActiveSheet: Set lg = SheetOf("maillog"): Set L = ListWs
    For Each c In Selection.Cells
        r = c.Row
        If r >= M_FIRST And Nz(ws.Cells(r, M_SUBJ).Value) <> "" And Not CHas(done, CStr(r)) Then
            done.Add r, CStr(r)
            If Nz(ws.Cells(r, M_SENT).Value) <> "" Then
                If Msg(F("msg.mail.already", ws.Cells(r, M_NO).Value), vbYesNo + vbQuestion) <> vbYes Then GoTo nxt
            End If
            key = Nz(ws.Cells(r, M_KEY).Value)
            lr = LastRow(lg, 1) + 1
            lg.Cells(lr, 1).Value = Now
            lg.Cells(lr, 2).Value = key
            lg.Cells(lr, 3).Value = Nz(ws.Cells(r, M_KIND).Value)
            lg.Cells(lr, 4).Value = Nz(ws.Cells(r, M_IDS).Value)
            lg.Cells(lr, 5).Value = Nz(ws.Cells(r, M_TO).Value)
            lg.Cells(lr, 6).Value = Nz(ws.Cells(r, M_SUBJ).Value)
            lg.Cells(lr, 7).Value = CurrentUser
            ws.Cells(r, M_SENT).Value = F("mail.sentAt", Format$(Now, "yyyy-mm-dd hh:nn"))
            ids = Split(Nz(ws.Cells(r, M_IDS).Value), ",")
            For i = 0 To UBound(ids)
                If key = "pass" Or key = "fail" Then
                    lr = FindRowById(Trim$(ids(i)))
                    If lr > 0 Then
                        gBusy = True
                        L.Cells(lr, C_NOTICED).Value = Date
                        gBusy = False
                    End If
                End If
                If key <> "reminder" And key <> "resultReminder" Then AddHist Trim$(ids(i)), T("f.mail"), "", Nz(ws.Cells(r, M_KIND).Value), ""
            Next i
            n = n + 1
        End If
nxt:
    Next c
    If n = 0 Then Msg T("msg.mail.pick"), vbExclamation: Exit Sub
    RecalcAll
    Msg F("msg.mail.marked", n)
End Sub

'==============================================================================
' Inspector reminders / result notice reminders
'==============================================================================
Private Function InspectorRows() As Variant
    Dim ws As Worksheet
    Set ws = SheetOf("insp")
    InspectorRows = ws.Range("A2:E" & (LastRow(ws, 1) + 1)).Value
End Function

Public Sub DraftInspectorReminders()
    Dim ins As Variant, L As Worksheet, v As Variant, last As Long, i As Long, k As Long, days As Long
    Dim today As Long, d As Long, st As String, list As String, cnt As Long, ids As String, n As Long
    Dim names As Collection, vals As Collection, nm As String, lines As String
    EnsureEvents
    Set L = ListWs
    days = SetN("set_remindDays"): If days < 0 Then days = 0
    today = CLng(Date)
    last = ListLast
    If last < FIRST Then Msg T("msg.none"): Exit Sub
    v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value
    ins = InspectorRows
    MailClearDrafts False
    For k = 1 To UBound(ins, 1)
        nm = Nz(ins(k, 1))
        If nm <> "" And Nz(ins(k, 5)) <> T("no") And Nz(ins(k, 3)) <> "" Then
            lines = "": cnt = 0: ids = ""
            For i = 1 To UBound(v, 1)
                st = StatusKey(Nz(v(i, C_STATUS)))
                If Nz(v(i, C_INSP)) = nm And (st = "confirmed" Or st = "changed") And IsRealDate(v(i, C_SDATE)) Then
                    d = CLng(DateOf(v(i, C_SDATE)))
                    If d >= today And d <= today + days Then
                        lines = InsertSorted(lines, "- " & MDKo(CDate(d)) & " " & Nz(v(i, C_STIME)) & T("sep") & Nz(v(i, C_VENDOR)) & _
                            IIf(Nz(v(i, C_PROJECT)) <> "", T("sep") & Nz(v(i, C_PROJECT)), "") & T("sep") & Nz(v(i, C_ITEM)) & T("sep") & _
                            Nz(v(i, C_SLOC)) & " (" & Nz(v(i, C_ID)) & ")" & IIf(st = "changed", " [" & T("st.changed") & "]", ""))
                        cnt = cnt + 1
                        ids = ids & IIf(ids = "", "", ", ") & Nz(v(i, C_ID))
                    End If
                End If
            Next i
            If cnt > 0 Then
                Set names = New Collection: Set vals = New Collection
                PutVar names, vals, "insp", nm
                PutVar names, vals, "count", CStr(cnt)
                PutVar names, vals, "list", lines
                PutVar names, vals, "period", MDKo(CDate(today)) & " ~ " & MDKo(CDate(today + days))
                PutVar names, vals, "sender", Nz(SetV("set_sender"))
                AddDraft TplLabel("reminder"), Nz(ins(k, 3)), "", Fill(TplText("reminder", 1), names, vals), Fill(TplText("reminder", 2), names, vals), ids, "reminder"
                n = n + 1
            End If
        End If
    Next k
    If n = 0 Then Msg T("msg.reminder.none"): Exit Sub
    SheetOf("mail").Activate
    Msg F("msg.mail.drafted", n)
End Sub

Public Sub DraftResultReminders()
    Dim ins As Variant, L As Worksheet, v As Variant, last As Long, i As Long, k As Long
    Dim today As Long, st As String, cnt As Long, ids As String, n As Long, rd As Long, dl As Long
    Dim names As Collection, vals As Collection, nm As String, lines As String, what As String, bse As Variant
    EnsureEvents
    Set L = ListWs
    RecalcAll
    today = CLng(Date)
    rd = SetN("set_resultDays"): If rd <= 0 Then rd = 7
    last = ListLast
    If last < FIRST Then Msg T("msg.none"): Exit Sub
    v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value
    ins = InspectorRows
    MailClearDrafts False
    For k = 1 To UBound(ins, 1)
        nm = Nz(ins(k, 1))
        If nm <> "" And Nz(ins(k, 3)) <> "" Then
            lines = "": cnt = 0: ids = ""
            For i = 1 To UBound(v, 1)
                st = StatusKey(Nz(v(i, C_STATUS)))
                what = ""
                If Nz(v(i, C_INSP)) = nm And IsRealDate(v(i, C_SDATE)) Then
                    If (st = "confirmed" Or st = "changed") And CLng(DateOf(v(i, C_SDATE))) < today And Nz(v(i, C_RESULT)) = "" Then
                        what = T("w.noResult"): bse = v(i, C_SDATE)
                    ElseIf st = "done" And Nz(v(i, C_NOTICED)) = "" Then
                        what = T("w.noNotice")
                        If IsRealDate(v(i, C_RDATE)) Then bse = v(i, C_RDATE) Else bse = v(i, C_SDATE)
                    End If
                End If
                If what <> "" Then
                    dl = CLng(DateOf(bse)) + rd
                    lines = lines & IIf(lines = "", "", vbLf) & "- " & Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM)) & " (" & Nz(v(i, C_ID)) & ")" & _
                        T("sep") & T("w.inspDate") & " " & MDKo(bse) & T("sep") & LeftLabel(dl - today) & T("sep") & what
                    cnt = cnt + 1
                    ids = ids & IIf(ids = "", "", ", ") & Nz(v(i, C_ID))
                End If
            Next i
            If cnt > 0 Then
                Set names = New Collection: Set vals = New Collection
                PutVar names, vals, "insp", nm
                PutVar names, vals, "count", CStr(cnt)
                PutVar names, vals, "list", lines
                PutVar names, vals, "deadline", CStr(rd)
                PutVar names, vals, "sender", Nz(SetV("set_sender"))
                AddDraft TplLabel("resultReminder"), Nz(ins(k, 3)), "", Fill(TplText("resultReminder", 1), names, vals), _
                    Fill(TplText("resultReminder", 2), names, vals), ids, "resultReminder"
                n = n + 1
            End If
        End If
    Next k
    If n = 0 Then Msg T("msg.result.none"): Exit Sub
    SheetOf("mail").Activate
    Msg F("msg.mail.drafted", n)
End Sub

' Popup at startup for the inspector matching this PC user
Public Sub StartupPopup()
    Dim ins As Variant, k As Long, nm As String, acct As String, me_ As String, L As Worksheet, v As Variant
    Dim i As Long, last As Long, st As String, d As Long, today As Long, days As Long, up As String, rs As String
    Dim rd As Long, nUp As Long, nRs As Long, dl As Long
    If Not SetYes("set_popup") Then Exit Sub
    acct = LCase$(Environ$("USERNAME")): me_ = CurrentUser
    ins = InspectorRows
    For k = 1 To UBound(ins, 1)
        If Nz(ins(k, 1)) <> "" And (LCase$(Nz(ins(k, 4))) = acct Or Nz(ins(k, 1)) = me_) Then nm = Nz(ins(k, 1)): Exit For
    Next k
    If nm = "" Then Exit Sub
    Set L = ListWs
    last = ListLast
    If last < FIRST Then Exit Sub
    v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value
    today = CLng(Date)
    days = SetN("set_remindDays"): If days < 0 Then days = 0
    rd = SetN("set_resultDays"): If rd <= 0 Then rd = 7
    For i = 1 To UBound(v, 1)
        st = StatusKey(Nz(v(i, C_STATUS)))
        If Nz(v(i, C_INSP)) = nm And IsRealDate(v(i, C_SDATE)) Then
            d = CLng(DateOf(v(i, C_SDATE)))
            If (st = "confirmed" Or st = "changed") And d >= today And d <= today + days Then
                up = up & vbCrLf & "- " & MDKo(CDate(d)) & " " & Nz(v(i, C_STIME)) & T("sep") & Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM)) & T("sep") & Nz(v(i, C_SLOC))
                nUp = nUp + 1
            End If
            If (st = "confirmed" Or st = "changed") And d < today And Nz(v(i, C_RESULT)) = "" Then
                rs = rs & vbCrLf & "- " & Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM)) & T("sep") & LeftLabel(d + rd - today) & T("sep") & T("w.noResult")
                nRs = nRs + 1
            ElseIf st = "done" And Nz(v(i, C_NOTICED)) = "" Then
                rs = rs & vbCrLf & "- " & Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM)) & T("sep") & LeftLabel(d + rd - today) & T("sep") & T("w.noNotice")
                nRs = nRs + 1
            End If
        End If
    Next i
    If nUp + nRs = 0 Then Exit Sub
    Msg F("popup.head", nm) & IIf(nUp > 0, vbCrLf & vbCrLf & F("popup.up", nUp) & up, "") & _
        IIf(nRs > 0, vbCrLf & vbCrLf & F("popup.result", nRs, rd) & rs, ""), vbInformation
End Sub
