
'==============================================================================
' Buttons: navigation and refresh
'==============================================================================
Public Sub RefreshAll()
    EnsureEvents
    Application.ScreenUpdating = False
    RecalcAll
    RefreshHome
    DrawCalendar
    Application.ScreenUpdating = True
End Sub

Public Sub GoHome(): EnsureEvents: RefreshHome: SheetOf("home").Activate: End Sub
Public Sub GoList(): EnsureEvents: ListWs.Activate: End Sub
Public Sub GoCalendar(): EnsureEvents: DrawCalendar: SheetOf("cal").Activate: End Sub
Public Sub GoMail(): EnsureEvents: SheetOf("mail").Activate: End Sub
Public Sub GoImport(): EnsureEvents: SheetOf("import").Activate: End Sub

Public Function ActiveId() As String
    Dim sh As Object, s As String
    Set sh = ActiveSheet
    If sh.Name = T("sh.list") Then
        If ActiveCell.Row >= FIRST Then ActiveId = Nz(sh.Cells(ActiveCell.Row, C_ID).Value)
    ElseIf sh.Name = T("sh.detail") Then
        ActiveId = Nz(sh.Range("C2").Value)
    Else
        s = Nz(ActiveCell.Value)
        If FindRowById(s) > 0 Then ActiveId = s
    End If
End Function

Private Function NeedId() As String
    NeedId = ActiveId
    If NeedId = "" Then Msg T("msg.pick.row"), vbExclamation
End Function

Public Sub ClearFilters()
    Dim ws As Worksheet
    EnsureEvents
    Set ws = ListWs
    Unprot ws
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
    ProtectSheet ws
End Sub

Public Sub FilterListByDate(ByVal d As Date)
    Dim ws As Worksheet
    Set ws = ListWs
    Unprot ws
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(HDR, 1), ws.Cells(MAXROW, C_LAST)).AutoFilter
    ws.Range(ws.Cells(HDR, 1), ws.Cells(MAXROW, C_LAST)).AutoFilter Field:=C_SDATE, _
        Criteria1:=">=" & CLng(d), Operator:=xlAnd, Criteria2:="<" & CLng(d + 1)
    On Error GoTo 0
    ProtectSheet ws
    ws.Activate
    Application.StatusBar = F("msg.filtered.date", DateKo(d))
End Sub

Public Sub ShowDetailForActive()
    Dim id As String
    EnsureEvents
    id = NeedId
    If id <> "" Then ShowDetail id, True
End Sub

'==============================================================================
' Home (dashboard)
'==============================================================================
Public Sub RefreshHome()
    Dim ws As Worksheet, L As Worksheet, last As Long, v As Variant, i As Long, n As Long
    Dim today As Long, wk0 As Long, wk1 As Long, m0 As Long, m1 As Long, d As Long, st As String
    Dim nToday As Long, nTodayDone As Long, nWeek As Long, nTodo As Long
    Dim cnt(1 To 6) As Long, keys As Variant, k As Long
    Dim todoRows As Long, upRows As Long, cfDates As New Collection, cfMsg As New Collection
    Dim upArr() As Variant, upN As Long, j As Long, tmp As Variant

    Set ws = SheetOf("home"): Set L = ListWs
    today = CLng(Date)
    wk0 = today - Weekday(Date, vbMonday) + 1: wk1 = wk0 + 6
    m0 = CLng(DateSerial(Year(Date), Month(Date), 1)): m1 = CLng(DateSerial(Year(Date), Month(Date) + 1, 0))
    keys = Array("registered", "nego", "confirmed", "changed", "done", "canceled")

    gBusy = True
    ws.Range("A12:E41,G12:K41,A45:E60").ClearContents
    ws.Range("A12:E41,G12:K41,A45:E60").Interior.ColorIndex = xlNone

    last = ListLast
    If last >= FIRST Then
        v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value
        n = UBound(v, 1)
        ReDim upArr(1 To n, 1 To 6)
        todoRows = 12
        For i = 1 To n
            If Nz(v(i, C_ID)) <> "" Then
                st = StatusKey(Nz(v(i, C_STATUS)))
                d = 0
                If IsRealDate(v(i, C_SDATE)) Then d = CLng(DateOf(v(i, C_SDATE)))
                If st <> "canceled" Then
                    If d = today Then
                        nToday = nToday + 1
                        If st = "done" Then nTodayDone = nTodayDone + 1
                    End If
                    If d >= wk0 And d <= wk1 Then nWeek = nWeek + 1
                End If
                If d >= m0 And d <= m1 Then
                    For k = 0 To 5
                        If keys(k) = st Then cnt(k + 1) = cnt(k + 1) + 1
                    Next k
                End If
                If Nz(v(i, C_TODO)) <> "" Then
                    nTodo = nTodo + 1
                    If todoRows <= 41 Then
                        ws.Cells(todoRows, 1).Value = Nz(v(i, C_ID))
                        ws.Cells(todoRows, 2).Value = Nz(v(i, C_VENDOR))
                        ws.Cells(todoRows, 3).Value = Nz(v(i, C_ITEM))
                        ws.Cells(todoRows, 4).Value = Nz(v(i, C_STATUS))
                        ws.Cells(todoRows, 5).Value = Nz(v(i, C_TODO))
                        If InStr(Nz(v(i, C_TODO)), T("left.overWord")) > 0 Then ws.Range(ws.Cells(todoRows, 1), ws.Cells(todoRows, 5)).Interior.Color = CLR_ERR
                        todoRows = todoRows + 1
                    End If
                End If
                If (st = "confirmed" Or st = "changed") And d >= today And d <= today + 7 Then
                    upN = upN + 1
                    upArr(upN, 1) = d
                    upArr(upN, 2) = Nz(v(i, C_STIME))
                    upArr(upN, 3) = Nz(v(i, C_ID))
                    upArr(upN, 4) = Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM))
                    upArr(upN, 5) = Nz(v(i, C_SLOC)) & IIf(Nz(v(i, C_INSP)) <> "", T("sep") & Nz(v(i, C_INSP)), "")
                    upArr(upN, 6) = st
                End If
                If Nz(v(i, C_CONFLICT)) <> "" And d >= today Then
                    If Not CHas(cfDates, CStr(d)) Then
                        cfDates.Add d, CStr(d)
                        cfMsg.Add Nz(v(i, C_CONFLICT)), CStr(d)
                    End If
                End If
            End If
        Next i
    End If

    ' sort upcoming by date/time
    For i = 1 To upN - 1
        For j = i + 1 To upN
            If upArr(j, 1) < upArr(i, 1) Or (upArr(j, 1) = upArr(i, 1) And upArr(j, 2) < upArr(i, 2)) Then
                For k = 1 To 6
                    tmp = upArr(i, k): upArr(i, k) = upArr(j, k): upArr(j, k) = tmp
                Next k
            End If
        Next j
    Next i
    upRows = 12
    For i = 1 To upN
        If upRows > 41 Then Exit For
        d = upArr(i, 1)
        If d = today Then ws.Cells(upRows, 7).Value = T("w.dday") Else ws.Cells(upRows, 7).Value = "D-" & (d - today)
        ws.Cells(upRows, 8).Value = MDKo(CDate(d)) & " " & upArr(i, 2)
        ws.Cells(upRows, 9).Value = upArr(i, 3)
        ws.Cells(upRows, 10).Value = upArr(i, 4)
        ws.Cells(upRows, 11).Value = upArr(i, 5) & IIf(upArr(i, 6) = "changed", " [" & T("st.changed") & "]", "")
        If d = today Then ws.Range(ws.Cells(upRows, 7), ws.Cells(upRows, 11)).Interior.Color = CLR_TODAY
        upRows = upRows + 1
    Next i

    ' conflicts, sorted by date
    Dim cfArr() As Long, cn As Long
    cn = cfDates.Count
    If cn > 0 Then
        ReDim cfArr(1 To cn)
        For i = 1 To cn: cfArr(i) = cfDates(i): Next i
        For i = 1 To cn - 1
            For j = i + 1 To cn
                If cfArr(j) < cfArr(i) Then k = cfArr(i): cfArr(i) = cfArr(j): cfArr(j) = k
            Next j
        Next i
        For i = 1 To cn
            If 44 + i > 60 Then Exit For
            ws.Cells(44 + i, 1).Value = DateKo(CDate(cfArr(i)))
            ws.Cells(44 + i, 2).Value = cfMsg(CStr(cfArr(i)))
        Next i
    End If

    ws.Range("A6").Value = F("w.count", nToday)
    ws.Range("A7").Value = F("home.todaySub", nTodayDone, nToday - nTodayDone)
    ws.Range("D6").Value = F("w.count", nWeek)
    ws.Range("D7").Value = MDKo(CDate(wk0)) & " ~ " & MDKo(CDate(wk1))
    ws.Range("G6").Value = F("w.days", cn)
    If cn > 0 Then ws.Range("G7").Value = F("home.cfFirst", MDKo(CDate(cfArr(1)))) Else ws.Range("G7").Value = T("home.cfNone")
    ws.Range("J6").Value = F("w.count", nTodo)
    ws.Range("J7").Value = IIf(nTodo > 0, T("home.todoSee"), T("home.todoNone"))
    ws.Range("A9").Value = F("home.month", Format$(Date, "yyyy-mm"), cnt(1), cnt(2), cnt(3), cnt(4), cnt(5), cnt(6))
    ws.Range("A2").Value = F("home.updated", Format$(Now, "yyyy-mm-dd hh:nn"), CurrentUser) & _
        IIf(ThisWorkbook.ReadOnly, "   " & T("home.readonly"), "")
    gBusy = False
End Sub

'==============================================================================
' Detail
'==============================================================================
Public Sub ShowDetail(ByVal id As String, ByVal doActivate As Boolean)
    Dim ws As Worksheet, L As Worksheet, r As Long, o As Worksheet, orow As Long, i As Long, rr As Long
    Dim labels As Variant, cur As Variant, org As Variant, cols As Variant, ocols As Variant
    Dim v As Variant, sec As Long

    Set ws = SheetOf("detail"): Set L = ListWs
    gBusy = True
    Application.EnableEvents = False
    ws.Range("A4:H400").ClearContents
    ws.Range("A4:H400").Interior.ColorIndex = xlNone
    ws.Range("A4:H400").Font.Bold = False
    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0
    ws.Range("C2").Value = id
    r = FindRowById(id)
    If r = 0 Then
        ws.Range("B4").Value = F("detail.notfound", id)
        GoTo fin
    End If
    ws.Range("E2").Value = Nz(L.Cells(r, C_STATUS).Value) & IIf(Nz(L.Cells(r, C_RESULT).Value) <> "", T("sep") & Nz(L.Cells(r, C_RESULT).Value), "")
    ws.Range("B4").Value = Nz(L.Cells(r, C_VENDOR).Value) & T("sep") & Nz(L.Cells(r, C_ITEM).Value)
    ws.Range("B4").Font.Bold = True

    ' original request snapshot
    Set o = SheetOf("orig")
    orow = 0
    For i = 2 To LastRow(o, 1)
        If UCase$(Nz(o.Cells(i, 1).Value)) = UCase$(id) Then orow = i
    Next i

    ' label, list column, original column (0 = none)
    cols = Array(C_VENDOR, C_PROJECT, C_CATEGORY, C_ITEM, C_QTY, C_SDATE, C_STIME, C_SLOC, C_INSP, _
                 C_CONTACT, C_PHONE, C_EMAIL, C_NOTE, C_RESULT, C_RDATE, C_FAIL, C_RMEMO, C_NOTICED, C_TODO, C_CONFLICT)
    ocols = Array(3, 4, 5, 6, 7, 8, 9, 10, 0, 11, 12, 13, 14, 0, 0, 0, 0, 0, 0, 0)
    ws.Range("B6").Value = T("detail.colItem"): ws.Range("C6").Value = T("detail.colOrig"): ws.Range("D6").Value = T("detail.colCur")
    ws.Range("B6:D6").Font.Bold = True
    ws.Range("B6:D6").Interior.Color = CLR_GRAY
    rr = 7
    For i = 0 To UBound(cols)
        ws.Cells(rr, 2).Value = ColLabel(cols(i))
        If ocols(i) > 0 And orow > 0 Then ws.Cells(rr, 3).Value = "'" & ValText(o.Cells(orow, ocols(i)).Value) Else ws.Cells(rr, 3).Value = "-"
        ws.Cells(rr, 4).Value = "'" & ValText(L.Cells(r, cols(i)).Value)
        If ocols(i) > 0 And orow > 0 Then
            If ValText(o.Cells(orow, ocols(i)).Value) <> ValText(L.Cells(r, cols(i)).Value) Then ws.Cells(rr, 4).Interior.Color = CLR_WARN
        End If
        rr = rr + 1
    Next i
    If orow = 0 Then ws.Cells(rr, 2).Value = T("detail.noOrig"): rr = rr + 1
    rr = rr + 1

    rr = DetailSection(ws, rr, T("detail.nego"), SheetOf("nego"), id, 2, Array(1, 4, 3), Array(T("w.when2"), T("w.writer"), T("w.content")))
    rr = DetailSection(ws, rr, T("detail.att"), SheetOf("att"), id, 2, Array(1, 3, 4, 6), Array(T("w.when2"), T("w.kind"), T("w.file"), T("w.writer")))
    rr = DetailSection(ws, rr, T("detail.mail"), SheetOf("maillog"), id, 4, Array(1, 3, 6, 5, 7), Array(T("w.when2"), T("w.kind"), T("w.subject"), T("w.to"), T("w.writer")))
    rr = DetailSection(ws, rr, T("detail.hist"), SheetOf("hist"), id, 3, Array(1, 2, 4, 5, 6, 7), Array(T("w.when2"), T("w.writer"), T("w.field"), T("w.before"), T("w.after"), T("w.reason")))
fin:
    Application.EnableEvents = True
    gBusy = False
    If doActivate Then ws.Activate: ws.Range("C2").Select
End Sub

' Writes rows of src whose idCol contains id (newest first). Returns next free row.
Private Function DetailSection(ByVal ws As Worksheet, ByVal rr As Long, ByVal title As String, ByVal src As Worksheet, _
        ByVal id As String, ByVal idCol As Long, ByVal cols As Variant, ByVal heads As Variant) As Long
    Dim last As Long, i As Long, k As Long, cnt As Long, cell As Range, hit As Boolean, s As String
    ws.Cells(rr, 2).Value = title
    ws.Cells(rr, 2).Font.Bold = True
    rr = rr + 1
    For k = 0 To UBound(heads)
        ws.Cells(rr, 2 + k).Value = heads(k)
    Next k
    ws.Range(ws.Cells(rr, 2), ws.Cells(rr, 2 + UBound(heads))).Interior.Color = CLR_GRAY
    rr = rr + 1
    last = LastRow(src, 1)
    For i = last To 2 Step -1
        s = UCase$(Nz(src.Cells(i, idCol).Value))
        hit = (s = UCase$(id)) Or (InStr(1, "," & Replace(s, " ", "") & ",", "," & UCase$(id) & ",") > 0)
        If hit Then
            For k = 0 To UBound(cols)
                Set cell = ws.Cells(rr, 2 + k)
                cell.Value = "'" & ValText(src.Cells(i, cols(k)).Value)
                If src.Cells(i, cols(k)).Hyperlinks.Count > 0 Then
                    ws.Hyperlinks.Add Anchor:=cell, Address:=src.Cells(i, cols(k)).Hyperlinks(1).Address, TextToDisplay:=Nz(src.Cells(i, cols(k)).Value)
                End If
            Next k
            rr = rr + 1
            cnt = cnt + 1
            If cnt >= 40 Then Exit For
        End If
    Next i
    If cnt = 0 Then ws.Cells(rr, 2).Value = T("w.none"): rr = rr + 1
    DetailSection = rr + 1
End Function

'==============================================================================
' Negotiation notes and attachments
'==============================================================================
Public Sub AddNegotiation()
    Dim id As String, s As String, ws As Worksheet, r As Long, L As Worksheet, lr As Long
    EnsureEvents
    id = NeedId
    If id = "" Then Exit Sub
    s = Trim$(InputBox(F("ask.nego", id), T("app.title")))
    If s = "" Then Exit Sub
    Set ws = SheetOf("nego")
    r = LastRow(ws, 1) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = id
    ws.Cells(r, 3).Value = s
    ws.Cells(r, 4).Value = CurrentUser
    Set L = ListWs
    lr = FindRowById(id)
    If lr > 0 Then
        If StatusKey(Nz(L.Cells(lr, C_STATUS).Value)) = "registered" Then
            gBusy = True
            L.Cells(lr, C_STATUS).Value = T("st.nego")
            L.Cells(lr, C_STATUS_AT).Value = Now
            gBusy = False
            AddHist id, T("f.status"), T("st.registered"), T("st.nego"), T("w.negoStart")
        End If
    End If
    RecalcAll
    If ActiveSheet.Name = T("sh.detail") Then ShowDetail id, False
End Sub

Public Sub AddAttachment()
    Dim id As String, files As Variant, kind As String, dirp As String, fv As Variant, nm As String
    Dim dest As String, ws As Worksheet, r As Long, n As Long, k As Long, bse As String, ext As String
    EnsureEvents
    id = NeedId
    If id = "" Then Exit Sub
    If BaseDir = "" Then Msg T("msg.save.first"), vbExclamation: Exit Sub
    files = Application.GetOpenFilename(T("filter.all"), , F("att.pick", id), , True)
    If Not IsArray(files) Then Exit Sub
    kind = Trim$(InputBox(T("ask.attkind"), T("app.title"), "1"))
    Select Case kind
        Case "1": kind = T("att.photo")
        Case "2": kind = T("att.report")
        Case "3": kind = T("att.etc")
        Case "": Exit Sub
    End Select
    dirp = BaseDir & "\" & T("dir.att") & "\" & SafeName(id)
    If Not EnsureDir(dirp) Then Msg F("msg.dir.fail", dirp), vbCritical: Exit Sub
    Set ws = SheetOf("att")
    For Each fv In files
        nm = Mid$(fv, InStrRev(fv, "\") + 1)
        dest = dirp & "\" & nm
        k = 1
        Do While Len(Dir(dest)) > 0
            k = k + 1
            If InStrRev(nm, ".") > 0 Then
                bse = Left$(nm, InStrRev(nm, ".") - 1): ext = Mid$(nm, InStrRev(nm, "."))
            Else
                bse = nm: ext = ""
            End If
            dest = dirp & "\" & bse & "_" & k & ext
        Loop
        On Error GoTo copyFail
        FileCopy fv, dest
        On Error GoTo 0
        r = LastRow(ws, 1) + 1
        ws.Cells(r, 1).Value = Now
        ws.Cells(r, 2).Value = id
        ws.Cells(r, 3).Value = kind
        ws.Cells(r, 4).Value = Mid$(dest, InStrRev(dest, "\") + 1)
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 4), Address:=T("dir.att") & "\" & SafeName(id) & "\" & Mid$(dest, InStrRev(dest, "\") + 1), _
            TextToDisplay:=Mid$(dest, InStrRev(dest, "\") + 1)
        ws.Cells(r, 5).Value = dest
        ws.Cells(r, 6).Value = CurrentUser
        AddHist id, T("f.att"), "", kind & " " & Mid$(dest, InStrRev(dest, "\") + 1), ""
        n = n + 1
nextFile:
    Next fv
    Msg F("msg.att.done", id, n)
    If ActiveSheet.Name = T("sh.detail") Then ShowDetail id, False
    Exit Sub
copyFail:
    Msg F("msg.copy.fail", fv, Err.Description), vbExclamation
    Resume nextFile
End Sub

'==============================================================================
' Calendar
'==============================================================================
Private Function CalStart(ByVal ws As Worksheet) As Long
    Dim m0 As Date, v As Variant
    v = ws.Range("C2").Value
    If IsRealDate(v) Then m0 = DateSerial(Year(DateOf(v)), Month(DateOf(v)), 1) Else m0 = DateSerial(Year(Date), Month(Date), 1)
    CalStart = CLng(m0) - (Weekday(m0) - 1)
End Function

Public Function CalendarDateAt(ByVal Target As Range) As Variant
    Dim ws As Worksheet, w As Long, d As Long
    Set ws = Target.Worksheet
    If Target.Row < 5 Or Target.Row > 10 Or Target.Column < 2 Or Target.Column > 8 Then Exit Function
    w = Target.Row - 5: d = Target.Column - 2
    CalendarDateAt = CDate(CalStart(ws) + w * 7 + d)
End Function

Public Sub DrawCalendar()
    Dim ws As Worksheet, L As Worksheet, m0 As Date, start As Long, last As Long, v As Variant
    Dim i As Long, w As Long, d As Long, dd As Long, cell As Range, txt As String, filt As String
    Dim lines As Collection, cnt As Long, conflictDay As Boolean, hol As Variant, hn As String, st As String
    Dim dayItems() As String, dayCf() As Boolean, nDays As Long, idx As Long

    Set ws = SheetOf("cal"): Set L = ListWs
    gBusy = True
    Application.EnableEvents = False
    If Not IsRealDate(ws.Range("C2").Value) Then ws.Range("C2").Value = DateSerial(Year(Date), Month(Date), 1)
    m0 = DateSerial(Year(DateOf(ws.Range("C2").Value)), Month(DateOf(ws.Range("C2").Value)), 1)
    ws.Range("C2").Value = m0
    start = CalStart(ws)
    filt = Nz(ws.Range("F2").Value)
    If filt = T("w.all") Then filt = ""
    nDays = 42
    ReDim dayItems(0 To nDays - 1): ReDim dayCf(0 To nDays - 1)

    last = ListLast
    If last >= FIRST Then
        v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value
        ' collect items per dd, sorted by time via simple insertion
        For i = 1 To UBound(v, 1)
            st = StatusKey(Nz(v(i, C_STATUS)))
            If Nz(v(i, C_ID)) <> "" And st <> "canceled" And IsRealDate(v(i, C_SDATE)) Then
                idx = CLng(DateOf(v(i, C_SDATE))) - start
                If idx >= 0 And idx < nDays And (filt = "" Or Nz(v(i, C_INSP)) = filt) Then
                    txt = Nz(v(i, C_STIME)) & " " & Nz(v(i, C_VENDOR)) & T("sep") & Nz(v(i, C_ITEM))
                    If Nz(v(i, C_INSP)) <> "" Then txt = txt & " (" & Nz(v(i, C_INSP)) & ")"
                    If st = "changed" Then txt = txt & " [" & T("st.changed") & "]"
                    If st = "registered" Or st = "nego" Then txt = txt & " [" & T("w.unconfirmed") & "]"
                    dayItems(idx) = InsertSorted(dayItems(idx), txt)
                    If Nz(v(i, C_CONFLICT)) <> "" Then dayCf(idx) = True
                End If
            End If
        Next i
    End If

    On Error Resume Next
    hol = ThisWorkbook.Names("set_holidays").RefersToRange.Value
    On Error GoTo 0
    For w = 0 To 5
        For d = 0 To 6
            idx = w * 7 + d
            dd = start + idx
            Set cell = ws.Cells(5 + w, 2 + d)
            hn = HolidayNameOf(dd, hol)
            txt = Format$(CDate(dd), "m/d") & IIf(hn <> "", " " & hn, "")
            If dayItems(idx) <> "" Then
                cnt = UBound(Split(dayItems(idx), vbLf)) + 1
                txt = txt & vbLf & dayItems(idx)
            End If
            cell.Value = txt
            cell.Font.Color = IIf(Month(CDate(dd)) = Month(m0), RGB(31, 41, 55), RGB(160, 160, 160))
            If dayCf(idx) Then
                cell.Interior.Color = CLR_CONFLICT
            ElseIf hn <> "" Or Weekday(CDate(dd)) = vbSunday Or Weekday(CDate(dd)) = vbSaturday Then
                cell.Interior.Color = CLR_GRAY
            ElseIf dd = CLng(Date) Then
                cell.Interior.Color = CLR_TODAY
            Else
                cell.Interior.ColorIndex = xlNone
            End If
            cell.Characters(1, Len(Format$(CDate(dd), "m/d"))).Font.Bold = True
        Next d
    Next w
    ws.Range("H2").Value = Format$(m0, "yyyy") & T("w.year") & " " & Format$(m0, "m") & T("w.month")
    Application.EnableEvents = True
    gBusy = False
End Sub

Private Function InsertSorted(ByVal existing As String, ByVal ln As String) As String
    Dim parts() As String, i As Long, out As String, placed As Boolean
    If existing = "" Then InsertSorted = ln: Exit Function
    parts = Split(existing, vbLf)
    For i = 0 To UBound(parts)
        If Not placed And ln < parts(i) Then
            out = out & IIf(out = "", "", vbLf) & ln
            placed = True
        End If
        out = out & IIf(out = "", "", vbLf) & parts(i)
    Next i
    If Not placed Then out = out & vbLf & ln
    InsertSorted = out
End Function

Private Function HolidayNameOf(ByVal d As Long, ByVal hol As Variant) As String
    Dim i As Long
    If Not IsArray(hol) Then Exit Function
    For i = 1 To UBound(hol, 1)
        If IsRealDate(hol(i, 1)) Then
            If CLng(DateOf(hol(i, 1))) = d Then
                HolidayNameOf = Nz(hol(i, 2))
                If HolidayNameOf = "" Then HolidayNameOf = T("w.holiday")
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub CalPrev()
    Call CalMove(-1)
End Sub
Public Sub CalNext()
    Call CalMove(1)
End Sub
Public Sub CalToday()
    Call CalMove(0)
End Sub

Private Sub CalMove(ByVal delta As Long)
    Dim ws As Worksheet, m0 As Date
    EnsureEvents
    Set ws = SheetOf("cal")
    If delta = 0 Or Not IsRealDate(ws.Range("C2").Value) Then
        m0 = DateSerial(Year(Date), Month(Date), 1)
    Else
        m0 = DateSerial(Year(DateOf(ws.Range("C2").Value)), Month(DateOf(ws.Range("C2").Value)) + delta, 1)
    End If
    gBusy = True
    ws.Range("C2").Value = m0
    gBusy = False
    DrawCalendar
End Sub
