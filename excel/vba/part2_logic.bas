
'==============================================================================
' Parsing: dates, times, quantities, e-mail
'==============================================================================
Private Function DigitGroups(ByVal s As String) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            cur = cur & ch
        ElseIf cur <> "" Then
            c.Add cur
            cur = ""
        End If
    Next i
    If cur <> "" Then c.Add cur
    Set DigitGroups = c
End Function

Private Function MakeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long, ByRef ok As Boolean) As Variant
    Dim x As Date
    ok = False
    If y < 2000 Or y > 2100 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    x = DateSerial(y, m, d)
    If Month(x) <> m Or Day(x) <> d Then Exit Function
    ok = True
    MakeDate = x
End Function

' Returns a Date, or Empty for a blank cell. ok=False when the text is not a date.
Public Function ParseDateV(ByVal v As Variant, ByRef ok As Boolean) As Variant
    Dim s As String, g As Collection
    ok = True
    ParseDateV = Empty
    If IsError(v) Then ok = False: Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then ParseDateV = DateOf(v): Exit Function
    If VarType(v) <> vbString Then
        If IsNumeric(v) Then
            If v > 36526 And v < 73051 Then ParseDateV = DateOf(v): Exit Function
        End If
        ok = False: Exit Function
    End If
    s = Trim$(v)
    If s = "" Then Exit Function
    Set g = DigitGroups(s)
    If g.Count = 1 Then
        If Len(g(1)) = 8 Then
            ParseDateV = MakeDate(CLng(Left$(g(1), 4)), CLng(Mid$(g(1), 5, 2)), CLng(Right$(g(1), 2)), ok)
            Exit Function
        End If
    ElseIf g.Count >= 3 Then
        If Len(g(1)) = 4 Then
            ParseDateV = MakeDate(CLng(g(1)), CLng(g(2)), CLng(g(3)), ok)
            Exit Function
        End If
    End If
    ok = False
End Function

' Returns "HH:MM", or "" for a blank cell. ok=False when not a valid time.
Public Function ParseTimeV(ByVal v As Variant, ByRef ok As Boolean) As String
    Dim s As String, g As Collection, h As Long, m As Long, total As Long, n As Double
    Dim pm As Boolean, am As Boolean
    ok = True
    If IsError(v) Then ok = False: Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then
        If Not IsNumeric(v) Then ok = False: Exit Function
        n = CDbl(v)
        If n >= 1 And n = Int(n) Then
            s = CStr(CLng(n))
        Else
            total = CLng((n - Int(n)) * 1440)
            If total >= 1440 Then total = 0
            ParseTimeV = Format$(total \ 60, "00") & ":" & Format$(total Mod 60, "00")
            Exit Function
        End If
    Else
        s = Trim$(v)
    End If
    If s = "" Then Exit Function
    pm = (InStr(1, s, T("time.pm"), vbTextCompare) > 0) Or (InStr(1, UCase$(s), "PM") > 0)
    am = (InStr(1, s, T("time.am"), vbTextCompare) > 0) Or (InStr(1, UCase$(s), "AM") > 0)
    Set g = DigitGroups(s)
    If g.Count = 0 Then ok = False: Exit Function
    If g.Count = 1 Then
        If Len(g(1)) >= 3 And Len(g(1)) <= 4 Then
            h = CLng(Left$(g(1), Len(g(1)) - 2)): m = CLng(Right$(g(1), 2))
        ElseIf Len(g(1)) <= 2 Then
            h = CLng(g(1)): m = 0
            If InStr(1, s, T("time.half"), vbTextCompare) > 0 Then m = 30
        Else
            ok = False: Exit Function
        End If
    Else
        If Len(g(1)) > 2 Or Len(g(2)) > 2 Then ok = False: Exit Function
        h = CLng(g(1)): m = CLng(g(2))
    End If
    If pm And h < 12 Then h = h + 12
    If am And h = 12 Then h = 0
    If h < 0 Or h > 23 Or m < 0 Or m > 59 Then ok = False: Exit Function
    ParseTimeV = Format$(h, "00") & ":" & Format$(m, "00")
End Function

Public Function TimeMinutes(ByVal s As String) As Long
    If Len(s) = 5 And Mid$(s, 3, 1) = ":" Then
        TimeMinutes = CLng(Left$(s, 2)) * 60 + CLng(Right$(s, 2))
    Else
        TimeMinutes = -1
    End If
End Function

' Returns a Long, or Empty for blank. ok=False when not a whole number >= 1.
Public Function ParseQtyV(ByVal v As Variant, ByRef ok As Boolean) As Variant
    Dim s As String, i As Long, digits As String, rest As String, ch As String, n As Double
    ok = True
    ParseQtyV = Empty
    If IsError(v) Then ok = False: Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then
        If Not IsNumeric(v) Then ok = False: Exit Function
        n = CDbl(v)
        If n >= 1 And n = Int(n) And n < 100000000# Then ParseQtyV = CLng(n) Else ok = False
        Exit Function
    End If
    s = Replace(Trim$(v), ",", "")
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            rest = Mid$(s, i)
            Exit For
        End If
    Next i
    If digits = "" Or Len(digits) > 8 Then ok = False: Exit Function
    For i = 1 To Len(rest)
        ch = Mid$(rest, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then ok = False: Exit Function
    Next i
    n = CDbl(digits)
    If n < 1 Then ok = False: Exit Function
    ParseQtyV = CLng(n)
End Function

Public Function IsEmailOk(ByVal s As String) As Boolean
    Dim at As Long, dot As Long
    s = Trim$(s)
    If s = "" Or InStr(s, " ") > 0 Then Exit Function
    at = InStr(s, "@")
    If at < 2 Then Exit Function
    If InStr(at + 1, s, "@") > 0 Then Exit Function
    dot = InStrRev(s, ".")
    If dot < at + 2 Or dot > Len(s) - 2 Then Exit Function
    IsEmailOk = True
End Function

'==============================================================================
' Conflict engine (works on plain arrays so the list and the import preview
' can share it)
'==============================================================================
' Slot arrays are 1-based with n entries:
'   sDate (Long serial, 0 = none), sTime ("HH:MM"), sLoc, sInsp, sVendor, sQty, sId
' Result: msgs(i) = conflict text for slot i ("" when none)
Public Function ComputeConflicts(ByVal n As Long, ByVal sDate As Variant, ByVal sTime As Variant, ByVal sLoc As Variant, _
        ByVal sInsp As Variant, ByVal sVendor As Variant, ByVal sQty As Variant, ByVal sId As Variant) As Variant
    Dim maxDay As Long, maxInsp As Long, slotMin As Long, maxVendor As Long, maxQty As Double
    Dim weekendWarn As Boolean, hol As Variant, i As Long, j As Long, k As Long, d As Long
    Dim days As New Collection, dayKey As Variant, members As Collection
    Dim idx() As Long, cnt As Long, label As String, visits As Long, s As String, msgs() As String

    ReDim msgs(1 To IIf(n < 1, 1, n))
    ComputeConflicts = msgs
    If n < 1 Then Exit Function
    maxDay = SetN("set_maxPerDay"): maxInsp = SetN("set_maxPerInsp"): slotMin = SetN("set_slotMin")
    maxVendor = SetN("set_maxPerVendor"): maxQty = SetN("set_maxQty")
    weekendWarn = SetYes("set_weekend")
    On Error Resume Next
    hol = ThisWorkbook.Names("set_holidays").RefersToRange.Value
    On Error GoTo 0

    ' group slots by date
    For i = 1 To n
        If sDate(i) > 0 Then
            If Not CHas(days, CStr(sDate(i))) Then
                Set members = New Collection
                days.Add members, CStr(sDate(i))
            End If
            Set members = days(CStr(sDate(i)))
            members.Add i
        End If
    Next i

    For Each members In days
        cnt = members.Count
        ReDim idx(1 To cnt)
        For k = 1 To cnt: idx(k) = members(k): Next k
        d = sDate(idx(1))

        ' holiday / weekend
        s = HolidayName(d, hol)
        If s <> "" Then
            For k = 1 To cnt: AddMsg msgs, idx(k), F("cf.holiday", MDKo(CDate(d)), s): Next k
        ElseIf weekendWarn And (Weekday(CDate(d)) = vbSaturday Or Weekday(CDate(d)) = vbSunday) Then
            For k = 1 To cnt: AddMsg msgs, idx(k), F("cf.weekend", MDKo(CDate(d))): Next k
        End If

        ' max per day (same vendor + same place = 1 visit)
        If maxDay > 0 Then
            visits = VisitCount(idx, cnt, sVendor, sLoc, "", sInsp)
            If visits > maxDay Then
                label = CountLabel(visits, cnt)
                For k = 1 To cnt: AddMsg msgs, idx(k), F("cf.day", maxDay, label): Next k
            End If
        End If

        ' max per inspector / vendor
        If maxInsp > 0 Then GroupRule idx, cnt, sInsp, sVendor, sLoc, maxInsp, "cf.insp", msgs
        If maxVendor > 0 Then GroupRule idx, cnt, sVendor, sVendor, sLoc, maxVendor, "cf.vendor", msgs

        ' quantity per day
        If maxQty > 0 Then
            Dim sumQ As Double
            sumQ = 0
            For k = 1 To cnt: sumQ = sumQ + sQty(idx(k)): Next k
            If sumQ > maxQty Then
                For k = 1 To cnt: AddMsg msgs, idx(k), F("cf.qty", maxQty, sumQ): Next k
            End If
        End If

        ' same time slot at the same place (different vendors only)
        If slotMin > 0 Then
            Dim a As Long, b As Long, ta As Long, tb As Long
            For a = 1 To cnt
                For b = a + 1 To cnt
                    i = idx(a): j = idx(b)
                    If sLoc(i) <> "" And sLoc(i) = sLoc(j) And sVendor(i) <> sVendor(j) Then
                        ta = TimeMinutes(sTime(i)): tb = TimeMinutes(sTime(j))
                        If ta >= 0 And tb >= 0 Then
                            If Abs(ta - tb) < slotMin Then
                                AddMsg msgs, i, F("cf.slot", slotMin, sTime(j), sId(j))
                                AddMsg msgs, j, F("cf.slot", slotMin, sTime(i), sId(i))
                            End If
                        End If
                    End If
                Next b
            Next a
        End If
    Next members
    ComputeConflicts = msgs
End Function

Private Sub AddMsg(msgs() As String, ByVal i As Long, ByVal s As String)
    If InStr(1, msgs(i), s, vbBinaryCompare) > 0 Then Exit Sub
    If msgs(i) <> "" Then msgs(i) = msgs(i) & " / "
    msgs(i) = msgs(i) & s
End Sub

Private Function HolidayName(ByVal d As Long, ByVal hol As Variant) As String
    Dim i As Long
    If Not IsArray(hol) Then Exit Function
    For i = 1 To UBound(hol, 1)
        If IsRealDate(hol(i, 1)) Then
            If CLng(DateOf(hol(i, 1))) = d Then
                HolidayName = Nz(hol(i, 2))
                If HolidayName = "" Then HolidayName = T("w.holiday")
                Exit Function
            End If
        End If
    Next i
End Function

Private Function VisitCount(idx As Variant, ByVal cnt As Long, sVendor As Variant, sLoc As Variant, _
        ByVal onlyKey As String, sKey As Variant) As Long
    Dim seen As New Collection, k As Long, key As String
    For k = 1 To cnt
        If onlyKey = "" Or sKey(idx(k)) = onlyKey Then
            If sVendor(idx(k)) = "" Then key = "#" & idx(k) Else key = sVendor(idx(k)) & "|" & sLoc(idx(k))
            If Not CHas(seen, key) Then seen.Add key, key
        End If
    Next k
    VisitCount = seen.Count
End Function

Private Function CountLabel(ByVal visits As Long, ByVal items As Long) As String
    If visits = items Then CountLabel = F("w.count", visits) Else CountLabel = F("w.countItems", visits, items)
End Function

Private Sub GroupRule(idx As Variant, ByVal cnt As Long, sKey As Variant, sVendor As Variant, sLoc As Variant, _
        ByVal limit As Long, ByVal msgKey As String, msgs() As String)
    Dim keys As New Collection, k As Long, kk As Variant, v As Long, items As Long, j As Long
    For k = 1 To cnt
        If sKey(idx(k)) <> "" Then
            If Not CHas(keys, sKey(idx(k))) Then keys.Add sKey(idx(k)), sKey(idx(k))
        End If
    Next k
    For Each kk In keys
        v = VisitCount(idx, cnt, sVendor, sLoc, CStr(kk), sKey)
        If v > limit Then
            items = 0
            For j = 1 To cnt
                If sKey(idx(j)) = CStr(kk) Then items = items + 1
            Next j
            For j = 1 To cnt
                If sKey(idx(j)) = CStr(kk) Then AddMsg msgs, idx(j), F(msgKey, limit, CStr(kk), CountLabel(v, items))
            Next j
        End If
    Next kk
End Sub

'==============================================================================
' Recalculate auto columns (conflicts, to-do) for the whole list
'==============================================================================
Public Sub RecalcAll()
    Dim ws As Worksheet, last As Long, n As Long, v As Variant, i As Long, r As Long
    Dim sDate() As Long, sTime() As String, sLoc() As String, sInsp() As String, sVendor() As String
    Dim sQty() As Double, sId() As String, msgs As Variant, outC() As Variant, outT() As Variant
    Dim st As String, todo As String, today As Long, resultDays As Long, negoLate As Long, dl As Long
    Dim mails As Collection, k As String, changedAt As Double

    Set ws = ListWs
    last = ListLast
    If last < FIRST Then Exit Sub
    n = last - FIRST + 1
    v = ws.Range(ws.Cells(FIRST, 1), ws.Cells(last, C_LAST)).Value
    ReDim sDate(1 To n): ReDim sTime(1 To n): ReDim sLoc(1 To n): ReDim sInsp(1 To n)
    ReDim sVendor(1 To n): ReDim sQty(1 To n): ReDim sId(1 To n)
    ReDim outC(1 To n, 1 To 1): ReDim outT(1 To n, 1 To 1)

    For i = 1 To n
        sId(i) = Nz(v(i, C_ID))
        st = StatusKey(Nz(v(i, C_STATUS)))
        If st <> "canceled" And sId(i) <> "" And IsRealDate(v(i, C_SDATE)) Then sDate(i) = CLng(DateOf(v(i, C_SDATE)))
        sTime(i) = Nz(v(i, C_STIME)): sLoc(i) = Nz(v(i, C_SLOC)): sInsp(i) = Nz(v(i, C_INSP))
        sVendor(i) = Nz(v(i, C_VENDOR))
        If IsNumeric(v(i, C_QTY)) And Not IsEmpty(v(i, C_QTY)) Then sQty(i) = CDbl(v(i, C_QTY))
    Next i
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)

    Set mails = MailLogIndex()
    today = CLng(Date)
    resultDays = SetN("set_resultDays"): If resultDays <= 0 Then resultDays = 7
    negoLate = SetN("set_negoLate"): If negoLate <= 0 Then negoLate = 3

    For i = 1 To n
        outC(i, 1) = msgs(i)
        todo = ""
        st = StatusKey(Nz(v(i, C_STATUS)))
        changedAt = 0
        If IsRealDate(v(i, C_STATUS_AT)) Or VarType(v(i, C_STATUS_AT)) = vbDate Then changedAt = CDbl(v(i, C_STATUS_AT))
        If sId(i) <> "" Then
            Select Case st
                Case "registered", ""
                    todo = T("todo.nego")
                Case "nego"
                    If changedAt > 0 Then
                        dl = today - Int(changedAt)
                        If dl >= negoLate Then todo = F("todo.negoLate", dl)
                    End If
                Case "confirmed", "changed"
                    If st = "changed" Then todo = T("todo.reconfirm")
                    If Not MailDone(mails, sId(i), st, changedAt) Then todo = JoinTodo(todo, T("todo.mail." & st))
                    If IsRealDate(v(i, C_SDATE)) Then
                        If CLng(DateOf(v(i, C_SDATE))) < today And Nz(v(i, C_RESULT)) = "" Then
                            dl = CLng(DateOf(v(i, C_SDATE))) + resultDays
                            todo = JoinTodo(todo, F("todo.noResult", MDKo(CDate(dl)), LeftLabel(dl - today)))
                        End If
                    End If
                Case "done"
                    If Nz(v(i, C_NOTICED)) = "" Then
                        If IsRealDate(v(i, C_RDATE)) Then
                            dl = CLng(DateOf(v(i, C_RDATE))) + resultDays
                        ElseIf IsRealDate(v(i, C_SDATE)) Then
                            dl = CLng(DateOf(v(i, C_SDATE))) + resultDays
                        Else
                            dl = today
                        End If
                        todo = F("todo.notice", MDKo(CDate(dl)), LeftLabel(dl - today))
                    End If
                Case "canceled"
                    If Not MailDone(mails, sId(i), st, changedAt) Then todo = T("todo.mail.canceled")
            End Select
        End If
        outT(i, 1) = todo
    Next i

    gBusy = True
    Application.EnableEvents = False
    ws.Range(ws.Cells(FIRST, C_CONFLICT), ws.Cells(last, C_CONFLICT)).Value = outC
    ws.Range(ws.Cells(FIRST, C_TODO), ws.Cells(last, C_TODO)).Value = outT
    Application.EnableEvents = True
    gBusy = False
End Sub

Private Function JoinTodo(ByVal a As String, ByVal b As String) As String
    If a = "" Then JoinTodo = b Else JoinTodo = a & " / " & b
End Function

Public Function LeftLabel(ByVal daysLeft As Long) As String
    If daysLeft > 0 Then
        LeftLabel = F("left.days", daysLeft)
    ElseIf daysLeft = 0 Then
        LeftLabel = T("left.today")
    Else
        LeftLabel = F("left.over", -daysLeft)
    End If
End Function

' id|kind -> latest logged time (Double)
Private Function MailLogIndex() As Collection
    Dim ws As Worksheet, v As Variant, i As Long, ids() As String, j As Long, key As String, c As New Collection
    Set ws = SheetOf("maillog")
    Set MailLogIndex = c
    If LastRow(ws, 1) < 2 Then Exit Function
    v = ws.Range("A2:D" & LastRow(ws, 1)).Value
    For i = 1 To UBound(v, 1)
        ids = Split(Nz(v(i, 4)), ",")
        For j = 0 To UBound(ids)
            key = UCase$(Trim$(ids(j))) & "|" & Nz(v(i, 2))
            If CHas(c, key) Then c.Remove key
            If IsRealDate(v(i, 1)) Or VarType(v(i, 1)) = vbDate Then
                c.Add CDbl(v(i, 1)), key
            Else
                c.Add CDbl(0), key
            End If
        Next j
    Next i
End Function

Private Function MailDone(ByVal mails As Collection, ByVal id As String, ByVal kind As String, ByVal since As Double) As Boolean
    Dim key As String
    key = UCase$(id) & "|" & kind
    If Not CHas(mails, key) Then Exit Function
    MailDone = (mails(key) >= since - 0.0001)
End Function
