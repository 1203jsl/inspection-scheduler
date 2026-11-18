Attribute VB_Name = "modInspection"
'==============================================================================
' Inspection schedule manager (Excel macro edition)
' All user-facing text lives in the hidden sheet "_text" (key -> text) so that
' this module stays ASCII-only and imports cleanly on any Windows code page.
'==============================================================================
Option Explicit

' ---- list sheet layout ----
Public Const HDR As Long = 3
Public Const FIRST As Long = 4
Public Const MAXROW As Long = 5000

Public Const C_ID As Long = 1
Public Const C_STATUS As Long = 2
Public Const C_VENDOR As Long = 3
Public Const C_PROJECT As Long = 4
Public Const C_CATEGORY As Long = 5
Public Const C_ITEM As Long = 6
Public Const C_QTY As Long = 7
Public Const C_SDATE As Long = 8
Public Const C_STIME As Long = 9
Public Const C_SLOC As Long = 10
Public Const C_INSP As Long = 11
Public Const C_RESULT As Long = 12
Public Const C_RDATE As Long = 13
Public Const C_FAIL As Long = 14
Public Const C_TODO As Long = 15
Public Const C_CONFLICT As Long = 16
Public Const C_NOTICED As Long = 17
Public Const C_REQDATE As Long = 18
Public Const C_REQTIME As Long = 19
Public Const C_REQLOC As Long = 20
Public Const C_CONTACT As Long = 21
Public Const C_PHONE As Long = 22
Public Const C_EMAIL As Long = 23
Public Const C_NOTE As Long = 24
Public Const C_RMEMO As Long = 25
Public Const C_CONF_AT As Long = 26
Public Const C_CONF_BY As Long = 27
Public Const C_STATUS_AT As Long = 28
Public Const C_CREATED_AT As Long = 29
Public Const C_CREATED_BY As Long = 30
Public Const C_SRCFILE As Long = 31
Public Const C_SRCROW As Long = 32
Public Const C_UPDATED As Long = 33
Public Const C_LAST As Long = 33

' ---- import preview sheet layout ----
Public Const I_HDR As Long = 4
Public Const I_FIRST As Long = 5
Public Const I_INCL As Long = 1
Public Const I_KIND As Long = 2
Public Const I_MSG As Long = 3
Public Const I_FILE As Long = 4
Public Const I_ROW As Long = 5
Public Const I_ID As Long = 6
Public Const I_VENDOR As Long = 7
Public Const I_PROJECT As Long = 8
Public Const I_CATEGORY As Long = 9
Public Const I_ITEM As Long = 10
Public Const I_QTY As Long = 11
Public Const I_DATE As Long = 12
Public Const I_TIME As Long = 13
Public Const I_LOC As Long = 14
Public Const I_CONTACT As Long = 15
Public Const I_PHONE As Long = 16
Public Const I_EMAIL As Long = 17
Public Const I_NOTE As Long = 18
Public Const I_PATH As Long = 19
Public Const I_LAST As Long = 19

' ---- mail sheet layout ----
Public Const M_HDR As Long = 4
Public Const M_FIRST As Long = 5
Public Const M_NO As Long = 1
Public Const M_KIND As Long = 2
Public Const M_TO As Long = 3
Public Const M_CC As Long = 4
Public Const M_SUBJ As Long = 5
Public Const M_BODY As Long = 6
Public Const M_IDS As Long = 7
Public Const M_SENT As Long = 8
Public Const M_KEY As Long = 9

' ---- colors ----
Public Const CLR_ERR As Long = 13551615     ' RGB(255,199,206)
Public Const CLR_WARN As Long = 10284031    ' RGB(255,235,156)
Public Const CLR_CONFLICT As Long = 11389944 ' RGB(248,203,173)
Public Const CLR_GRAY As Long = 15921906    ' RGB(242,242,242)
Public Const CLR_TODAY As Long = 16247773   ' RGB(221,235,247)

Public gBusy As Boolean
Private gTxt As Collection
Private gEvents As clsAppEvents
Private gOldAddr As String
Private gOldVals As Variant
Private gRoWarned As Boolean
Private gSavingCopy As Boolean

'==============================================================================
' Text / settings helpers
'==============================================================================
Public Function T(ByVal k As String) As String
    If gTxt Is Nothing Then LoadTexts
    On Error GoTo missing
    T = gTxt(k)
    Exit Function
missing:
    T = "[" & k & "]"
End Function

Public Function F(ByVal k As String, ParamArray a() As Variant) As String
    Dim s As String, i As Long
    s = T(k)
    For i = LBound(a) To UBound(a)
        s = Replace(s, "{" & i & "}", CStr(a(i)))
    Next i
    F = s
End Function

Private Sub LoadTexts()
    Dim ws As Worksheet, v As Variant, i As Long
    Set gTxt = New Collection
    Set ws = ThisWorkbook.Worksheets("_text")
    v = ws.Range("A1:B" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Value
    On Error Resume Next
    For i = 2 To UBound(v, 1)
        If Len(CStr(v(i, 1))) > 0 Then gTxt.Add Replace(CStr(v(i, 2)), "\n", vbLf), CStr(v(i, 1))
    Next i
    On Error GoTo 0
End Sub

Public Sub ReloadTexts()
    Set gTxt = Nothing
    LoadTexts
End Sub

Public Function SheetOf(ByVal k As String) As Worksheet
    Set SheetOf = ThisWorkbook.Worksheets(T("sh." & k))
End Function

Public Function ListWs() As Worksheet
    Set ListWs = SheetOf("list")
End Function

Public Function SetV(ByVal nm As String) As Variant
    On Error GoTo none
    SetV = ThisWorkbook.Names(nm).RefersToRange.Cells(1, 1).Value
    Exit Function
none:
    SetV = Empty
End Function

Public Function SetN(ByVal nm As String) As Double
    Dim v As Variant
    v = SetV(nm)
    If IsNumeric(v) And Not IsEmpty(v) Then SetN = CDbl(v) Else SetN = 0
End Function

Public Function SetYes(ByVal nm As String) As Boolean
    SetYes = (Nz(SetV(nm)) = T("yes"))
End Function

Public Function Nz(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsArray(v) Then Exit Function
    Nz = Trim$(CStr(v))
End Function

Public Function IsRealDate(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then IsRealDate = True: Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        If v > 36526 And v < 73051 Then IsRealDate = True
    End If
End Function

Public Function DateOf(ByVal v As Variant) As Date
    DateOf = CDate(Int(CDbl(v)))
End Function

Public Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function ListLast() As Long
    Dim ws As Worksheet, a As Long, b As Long, c As Long
    Set ws = ListWs
    a = LastRow(ws, C_ID): b = LastRow(ws, C_VENDOR): c = LastRow(ws, C_ITEM)
    If b > a Then a = b
    If c > a Then a = c
    If a < FIRST Then a = FIRST - 1
    ListLast = a
End Function

Public Function FindRowById(ByVal id As String) As Long
    Dim ws As Worksheet, v As Variant, i As Long, last As Long
    id = UCase$(Trim$(id))
    If id = "" Then Exit Function
    Set ws = ListWs
    last = ListLast
    If last < FIRST Then Exit Function
    v = ws.Range(ws.Cells(FIRST, C_ID), ws.Cells(last + 1, C_ID)).Value
    For i = 1 To UBound(v, 1)
        If UCase$(Nz(v(i, 1))) = id Then FindRowById = FIRST + i - 1: Exit Function
    Next i
End Function

Public Function WeekdayKo(ByVal d As Date) As String
    WeekdayKo = Split(T("weekdays"), ",")(Weekday(d) - 1)
End Function

Public Function YMD(ByVal d As Variant) As String
    If IsRealDate(d) Then YMD = Format$(DateOf(d), "yyyy-mm-dd")
End Function

Public Function DateKo(ByVal d As Variant) As String
    If IsRealDate(d) Then DateKo = Format$(DateOf(d), "yyyy-mm-dd") & "(" & WeekdayKo(DateOf(d)) & ")" Else DateKo = "-"
End Function

Public Function MDKo(ByVal d As Variant) As String
    If IsRealDate(d) Then MDKo = Format$(DateOf(d), "mm-dd") & "(" & WeekdayKo(DateOf(d)) & ")" Else MDKo = "-"
End Function

Public Function ValText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        If CDbl(v) = Int(CDbl(v)) Then ValText = Format$(v, "yyyy-mm-dd") Else ValText = Format$(v, "yyyy-mm-dd hh:nn")
        Exit Function
    End If
    ValText = Nz(v)
End Function

Public Function CurrentUser() As String
    Dim acct As String, v As Variant, i As Long
    acct = LCase$(Environ$("USERNAME"))
    On Error Resume Next
    v = ThisWorkbook.Names("set_users").RefersToRange.Value
    On Error GoTo 0
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If LCase$(Nz(v(i, 1))) = acct And acct <> "" And Nz(v(i, 2)) <> "" Then CurrentUser = Nz(v(i, 2)): Exit Function
        Next i
    End If
    CurrentUser = Trim$(Application.UserName)
    If CurrentUser = "" Then CurrentUser = Environ$("USERNAME")
End Function

Public Function BaseDir() As String
    BaseDir = ThisWorkbook.Path
End Function

Public Function EnsureDir(ByVal p As String) As Boolean
    Dim parts() As String, cur As String, i As Long
    On Error Resume Next
    If Len(Dir(p, vbDirectory)) > 0 Then EnsureDir = True: Exit Function
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
    EnsureDir = (Len(Dir(p, vbDirectory)) > 0)
End Function

Public Function SafeName(ByVal s As String) As String
    Dim bad As Variant, b As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each b In bad
        s = Replace(s, CStr(b), "_")
    Next b
    SafeName = Trim$(s)
End Function

Public Function InList(ByVal s As String, ByVal csv As String) As Boolean
    InList = (InStr(1, "," & csv & ",", "," & s & ",", vbBinaryCompare) > 0)
End Function

Public Function JoinCol(ByVal c As Collection, ByVal sep As String) As String
    Dim x As Variant, s As String
    For Each x In c
        If s <> "" Then s = s & sep
        s = s & CStr(x)
    Next x
    JoinCol = s
End Function

Public Function CHas(ByVal c As Collection, ByVal k As String) As Boolean
    Dim b As Boolean
    On Error Resume Next
    Err.Clear
    b = IsObject(c.Item(k))
    CHas = (Err.Number = 0)
    Err.Clear
End Function

Public Function Msg(ByVal text As String, Optional ByVal style As Long = 64) As Long
    Msg = MsgBox(text, style, T("app.title"))
End Function

Public Function AskReason(ByVal prompt As String) As String
    Dim s As String
    s = Trim$(InputBox(prompt, T("app.title")))
    If s = "" Then Msg T("msg.reason.required"), vbExclamation
    AskReason = s
End Function

'==============================================================================
' Status helpers
'==============================================================================
Public Function StatusKey(ByVal s As String) As String
    Select Case Trim$(s)
        Case T("st.registered"): StatusKey = "registered"
        Case T("st.nego"): StatusKey = "nego"
        Case T("st.confirmed"): StatusKey = "confirmed"
        Case T("st.changed"): StatusKey = "changed"
        Case T("st.done"): StatusKey = "done"
        Case T("st.canceled"): StatusKey = "canceled"
        Case Else: StatusKey = ""
    End Select
End Function

Public Function StatusLabel(ByVal k As String) As String
    StatusLabel = T("st." & k)
End Function

Private Function AllowedNext(ByVal fromKey As String) As String
    Select Case fromKey
        Case "registered", "": AllowedNext = "nego,confirmed,canceled"
        Case "nego": AllowedNext = "registered,confirmed,canceled"
        Case "confirmed": AllowedNext = "changed,done,canceled"
        Case "changed": AllowedNext = "confirmed,done,canceled"
        Case "done": AllowedNext = ""
        Case "canceled": AllowedNext = ""
    End Select
End Function

Public Function ActiveInspector(ByVal nm As String) As Boolean
    Dim ws As Worksheet, v As Variant, i As Long
    nm = Trim$(nm)
    If nm = "" Then Exit Function
    Set ws = SheetOf("insp")
    v = ws.Range("A2:E" & (LastRow(ws, 1) + 1)).Value
    For i = 1 To UBound(v, 1)
        If Nz(v(i, 1)) = nm And Nz(v(i, 5)) <> T("no") Then ActiveInspector = True: Exit Function
    Next i
End Function

Public Function InspectorEmail(ByVal nm As String) As String
    Dim ws As Worksheet, v As Variant, i As Long
    Set ws = SheetOf("insp")
    v = ws.Range("A2:E" & (LastRow(ws, 1) + 1)).Value
    For i = 1 To UBound(v, 1)
        If Nz(v(i, 1)) = Trim$(nm) Then InspectorEmail = Nz(v(i, 3)): Exit Function
    Next i
End Function

Public Function NextId() As String
    Dim ws As Worksheet, v As Variant, i As Long, pre As String, n As Long, mx As Long, s As String
    pre = Nz(SetV("set_prefix"))
    If pre = "" Then pre = "IN"
    pre = pre & "-" & Format$(Date, "yyyy") & "-"
    Set ws = ListWs
    If ListLast >= FIRST Then
        v = ws.Range(ws.Cells(FIRST, C_ID), ws.Cells(ListLast + 1, C_ID)).Value
        For i = 1 To UBound(v, 1)
            s = UCase$(Nz(v(i, 1)))
            If Left$(s, Len(pre)) = UCase$(pre) Then
                n = Val(Mid$(s, Len(pre) + 1))
                If n > mx Then mx = n
            End If
        Next i
    End If
    NextId = pre & Format$(mx + 1, "0000")
End Function

Public Sub AddHist(ByVal id As String, ByVal field As String, ByVal oldV As Variant, ByVal newV As Variant, Optional ByVal reason As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = SheetOf("hist")
    r = LastRow(ws, 1) + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = CurrentUser
    ws.Cells(r, 3).Value = id
    ws.Cells(r, 4).Value = field
    ws.Cells(r, 5).Value = "'" & ValText(oldV)
    ws.Cells(r, 6).Value = "'" & ValText(newV)
    ws.Cells(r, 7).Value = reason
End Sub

Public Function ColLabel(ByVal col As Long) As String
    ColLabel = Replace(Nz(ListWs.Cells(HDR, col).Value), " *", "")
End Function

Private Sub Unprot(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
End Sub

Public Sub ProtectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True
    ws.EnableAutoFilter = True
End Sub

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

'==============================================================================
' Vendor request form (new workbook)
'==============================================================================
Public Sub MakeVendorTemplate()
    Dim wb As Workbook, ws As Worksheet, g As Worksheet, m As Worksheet, heads As Variant, ex As Variant
    Dim widths As Variant, i As Long, items As String, v As Variant, path As Variant, req As Variant
    EnsureEvents
    heads = Split(T("tpl.heads"), "|")
    req = Split(T("tpl.required"), "|")
    widths = Array(15, 14, 18, 12, 30, 8, 13, 10, 14, 11, 15, 24, 24)
    On Error Resume Next
    v = ThisWorkbook.Names("set_items").RefersToRange.Value
    On Error GoTo 0
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Nz(v(i, 1)) <> "" Then items = items & IIf(items = "", "", ",") & Nz(v(i, 1))
        Next i
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = T("tpl.sheet")
    For i = 0 To UBound(heads)
        ws.Cells(1, i + 1).Value = heads(i) & IIf(req(i) = "1", " *", "")
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(heads) + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 47, 99)
        .HorizontalAlignment = xlCenter
    End With
    ' example row
    ex = Split(T("tpl.example"), "|")
    For i = 0 To UBound(ex)
        ws.Cells(2, i + 1).Value = ex(i)
    Next i
    If items <> "" Then ws.Cells(2, 4).Value = Split(items, ",")(0)
    ws.Cells(2, 7).Value = Date + 7
    ws.Cells(2, 7).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(heads) + 1)).Font.Italic = True
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(heads) + 1)).Font.Color = RGB(85, 96, 122)
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(heads) + 1)).Interior.Color = RGB(243, 245, 249)
    ' input area
    ws.Range("A3:M302").Locked = False
    ws.Range("G3:G302").NumberFormat = "yyyy-mm-dd"
    ws.Range("H3:H302").NumberFormat = "@"
    ws.Range("K3:K302").NumberFormat = "@"
    On Error Resume Next
    If items <> "" Then
        With ws.Range("D3:D302").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:=items
            .ErrorMessage = T("tpl.v.category")
        End With
    End If
    With ws.Range("G3:G302").Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=CStr(CLng(DateSerial(2020, 1, 1)))
        .ErrorMessage = T("tpl.v.date")
        .InputMessage = T("tpl.v.dateHint")
    End With
    With ws.Range("F3:F302").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1"
        .ErrorMessage = T("tpl.v.qty")
    End With
    With ws.Range("H3:H302").Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputMessage = T("tpl.v.timeHint")
    End With
    On Error GoTo 0
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Protect Password:="", AllowFormattingColumns:=True, AllowFormattingRows:=True

    Set g = wb.Worksheets.Add(After:=ws)
    g.Name = T("tpl.guideSheet")
    g.Columns(1).ColumnWidth = 110
    v = Split(T("tpl.guide"), vbLf)
    For i = 0 To UBound(v)
        g.Cells(i + 1, 1).Value = v(i)
    Next i
    g.Cells(1, 1).Font.Bold = True
    g.Cells(1, 1).Font.Size = 14
    Set m = wb.Worksheets.Add(After:=g)
    m.Name = "_meta"
    m.Range("A1").Value = "TEMPLATE_VERSION"
    m.Range("A2").Value = "v2-excel"
    m.Visible = xlSheetVeryHidden
    ws.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    path = Application.GetSaveAsFilename(InitialFileName:=T("tpl.fileName"), FileFilter:=T("filter.xlsx"))
    If VarType(path) = vbBoolean Then
        Msg T("msg.tpl.notSaved"), vbInformation
        Exit Sub
    End If
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(path), FileFormat:=51
    Application.DisplayAlerts = True
    Msg F("msg.tpl.saved", path)
End Sub

'==============================================================================
' Vendor schedule confirmation (new workbook)
'==============================================================================
Public Sub MakeVendorSchedule()
    Dim vendor As String, d0 As Variant, d1 As Variant, ok As Boolean, L As Worksheet, v As Variant, last As Long
    Dim wb As Workbook, ws As Worksheet, i As Long, r As Long, heads As Variant, st As String, d As Long
    Dim remark As String, n As Long, path As Variant, defV As String
    EnsureEvents
    Set L = ListWs
    If ActiveSheet.Name = T("sh.list") And ActiveCell.Row >= FIRST Then defV = Nz(L.Cells(ActiveCell.Row, C_VENDOR).Value)
    vendor = Trim$(InputBox(T("ask.vs.vendor"), T("app.title"), defV))
    If vendor = "" Then Exit Sub
    d0 = ParseDateV(InputBox(T("ask.vs.from"), T("app.title"), Format$(DateSerial(Year(Date), Month(Date), 1), "yyyy-mm-dd")), ok)
    If Not ok Or IsEmpty(d0) Then Msg T("msg.date.invalid"), vbExclamation: Exit Sub
    d1 = ParseDateV(InputBox(T("ask.vs.to"), T("app.title"), Format$(DateSerial(Year(Date), Month(Date) + 1, 0), "yyyy-mm-dd")), ok)
    If Not ok Or IsEmpty(d1) Then Msg T("msg.date.invalid"), vbExclamation: Exit Sub
    last = ListLast
    If last < FIRST Then Msg T("msg.none"): Exit Sub
    v = L.Range(L.Cells(FIRST, 1), L.Cells(last, C_LAST)).Value

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = T("vs.sheet")
    heads = Split(T("vs.heads"), "|")
    ws.Range("A1").Value = T("vs.title")
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = T("vs.vendor"): ws.Range("B3").Value = vendor
    ws.Range("A4").Value = T("vs.period"): ws.Range("B4").Value = DateKo(d0) & " ~ " & DateKo(d1)
    ws.Range("A5").Value = T("vs.made"): ws.Range("B5").Value = DateKo(Date)
    ws.Range("A6").Value = T("vs.by"): ws.Range("B6").Value = CurrentUser
    ws.Range("A3:A6").Font.Bold = True
    For i = 0 To UBound(heads)
        ws.Cells(8, i + 1).Value = heads(i)
    Next i
    With ws.Range(ws.Cells(8, 1), ws.Cells(8, UBound(heads) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 47, 99)
        .Font.Color = RGB(255, 255, 255)
    End With
    r = 9
    For i = 1 To UBound(v, 1)
        If Nz(v(i, C_VENDOR)) = vendor And IsRealDate(v(i, C_SDATE)) Then
            d = CLng(DateOf(v(i, C_SDATE)))
            If d >= CLng(d0) And d <= CLng(d1) Then
                st = StatusKey(Nz(v(i, C_STATUS)))
                remark = ""
                If st = "canceled" Then remark = T("vs.canceled")
                If st = "changed" Then remark = T("vs.changed")
                If st = "registered" Or st = "nego" Then remark = T("vs.unconfirmed")
                If Nz(v(i, C_RESULT)) = T("res.fail") Then remark = T("res.fail") & ": " & Nz(v(i, C_FAIL))
                ws.Cells(r, 1).Value = Nz(v(i, C_ID))
                ws.Cells(r, 2).Value = Nz(v(i, C_PROJECT))
                ws.Cells(r, 3).Value = Nz(v(i, C_ITEM))
                ws.Cells(r, 4).Value = v(i, C_QTY)
                ws.Cells(r, 5).Value = DateKo(v(i, C_SDATE))
                ws.Cells(r, 6).Value = Nz(v(i, C_STIME))
                ws.Cells(r, 7).Value = Nz(v(i, C_SLOC))
                ws.Cells(r, 8).Value = Nz(v(i, C_INSP))
                ws.Cells(r, 9).Value = Nz(v(i, C_STATUS)) & IIf(Nz(v(i, C_RESULT)) <> "", T("sep") & Nz(v(i, C_RESULT)), "")
                ws.Cells(r, 10).Value = remark
                r = r + 1: n = n + 1
            End If
        End If
    Next i
    If n = 0 Then ws.Cells(9, 1).Value = T("w.none")
    ws.Range(ws.Cells(8, 1), ws.Cells(IIf(r > 9, r - 1, 9), UBound(heads) + 1)).Borders.LineStyle = xlContinuous
    ws.Columns("A:J").AutoFit
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    path = Application.GetSaveAsFilename(InitialFileName:=SafeName(F("vs.fileName", vendor, Format$(Date, "yyyymmdd"))), FileFilter:=T("filter.xlsx"))
    If VarType(path) = vbBoolean Then Exit Sub
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(path), FileFormat:=51
    Application.DisplayAlerts = True
    Msg F("msg.vs.saved", n, path)
End Sub

'==============================================================================
' Export the currently filtered list
'==============================================================================
Public Sub ExportVisible()
    Dim L As Worksheet, wb As Workbook, rng As Range, path As Variant, last As Long
    EnsureEvents
    Set L = ListWs
    last = ListLast
    If last < FIRST Then Msg T("msg.none"): Exit Sub
    Application.ScreenUpdating = False
    Set rng = L.Range(L.Cells(HDR, 1), L.Cells(last, C_LAST)).SpecialCells(xlCellTypeVisible)
    Set wb = Workbooks.Add(xlWBATWorksheet)
    rng.Copy
    wb.Worksheets(1).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    wb.Worksheets(1).Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    wb.Worksheets(1).Rows(1).Font.Bold = True
    wb.Worksheets(1).Name = T("sh.list")
    Application.ScreenUpdating = True
    path = Application.GetSaveAsFilename(InitialFileName:=F("export.fileName", Format$(Now, "yyyymmdd_hhnn")), FileFilter:=T("filter.xlsx"))
    If VarType(path) = vbBoolean Then Exit Sub
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(path), FileFormat:=51
    Application.DisplayAlerts = True
    Msg F("msg.export.saved", path)
End Sub

'==============================================================================
' Backups (a copy after every save, oldest removed)
'==============================================================================
Public Sub MakeBackup()
    Dim dirp As String, nm As String, keep As Long, fl() As String, n As Long, fv As String, i As Long, j As Long, tv As String
    If ThisWorkbook.ReadOnly Or BaseDir = "" Then Exit Sub
    dirp = BaseDir & "\" & T("dir.backup")
    If Not EnsureDir(dirp) Then Exit Sub
    nm = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    On Error GoTo fin
    gSavingCopy = True
    ThisWorkbook.SaveCopyAs dirp & "\" & nm & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    gSavingCopy = False
    keep = SetN("set_backupKeep"): If keep < 5 Then keep = 5
    ReDim fl(1 To 1)
    fv = Dir(dirp & "\" & nm & "_*.xlsm")
    Do While fv <> ""
        n = n + 1
        ReDim Preserve fl(1 To n)
        fl(n) = fv
        fv = Dir()
    Loop
    If n <= keep Then Exit Sub
    For i = 1 To n - 1
        For j = i + 1 To n
            If fl(j) < fl(i) Then tv = fl(i): fl(i) = fl(j): fl(j) = tv
        Next j
    Next i
    For i = 1 To n - keep
        Kill dirp & "\" & fl(i)
    Next i
fin:
    gSavingCopy = False
End Sub

Public Sub OpenBackupFolder(): OpenFolder T("dir.backup"): End Sub
Public Sub OpenOrigFolder(): OpenFolder T("dir.orig"): End Sub

Private Sub OpenFolder(ByVal sub_ As String)
    Dim p As String
    EnsureEvents
    If BaseDir = "" Then Msg T("msg.save.first"), vbExclamation: Exit Sub
    p = BaseDir & "\" & sub_
    EnsureDir p
    On Error GoTo fail
    ThisWorkbook.FollowHyperlink p
    Exit Sub
fail:
    Msg F("msg.open.fail", p), vbExclamation
End Sub

'==============================================================================
' Setup: folders, buttons, protection
'==============================================================================
Public Sub EnsureSetup(Optional ByVal force As Boolean = False)
    Dim p As String
    p = BaseDir
    If p <> "" And Not ThisWorkbook.ReadOnly Then
        EnsureDir p & "\" & T("dir.orig")
        EnsureDir p & "\" & T("dir.att")
        EnsureDir p & "\" & T("dir.backup")
    End If
    If force Or Not HasShape(SheetOf("home"), "btn_home_refresh") Then CreateButtons
    ApplyProtection
End Sub

Public Sub SetupNow()
    EnsureEvents
    ReloadTexts
    EnsureSetup True
    RefreshAll
    Msg T("msg.setup.done")
End Sub

Private Function HasShape(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Object
    On Error GoTo no
    Set s = ws.Shapes(nm)
    HasShape = True
no:
End Function

Public Sub ApplyProtection()
    Dim k As Variant
    For Each k In Array("list", "hist", "maillog", "orig", "nego", "att", "home", "cal", "detail")
        Unprot SheetOf(CStr(k))
        ProtectSheet SheetOf(CStr(k))
    Next k
End Sub

Private Sub CreateButtons()
    Dim specs As Variant, s As Variant, p() As String, ws As Worksheet, shp As Object
    Dim i As Long, k As Variant, lefts As New Collection, key As String, x As Double, top As Double, h As Double
    ' sheetKey|name|row|widthPt|textKey|macro|style(1=primary)
    Dim sp As String
    sp = sp & "home|refresh|3|90|btn.refresh|RefreshAll|1;"
    sp = sp & "home|import|3|110|btn.import|GoImport|1;"
    sp = sp & "home|template|3|130|btn.template|MakeVendorTemplate|0;"
    sp = sp & "home|list|3|90|btn.list|GoList|0;"
    sp = sp & "home|cal|3|80|btn.cal|GoCalendar|0;"
    sp = sp & "home|mail|3|80|btn.mail|GoMail|0;"
    sp = sp & "home|remind|4|140|btn.remind|DraftInspectorReminders|0;"
    sp = sp & "home|resrem|4|150|btn.resultRemind|DraftResultReminders|0;"
    sp = sp & "home|vs|4|130|btn.vendorSchedule|MakeVendorSchedule|0;"
    sp = sp & "home|backup|4|110|btn.backupFolder|OpenBackupFolder|0;"
    sp = sp & "home|orig|4|120|btn.origFolder|OpenOrigFolder|0;"
    sp = sp & "list|mail|1|130|btn.mailSel|MailForSelection|1;"
    sp = sp & "list|detail|1|80|btn.detail|ShowDetailForActive|0;"
    sp = sp & "list|nego|1|110|btn.nego|AddNegotiation|0;"
    sp = sp & "list|att|1|90|btn.att|AddAttachment|0;"
    sp = sp & "list|clear|1|100|btn.clearFilter|ClearFilters|0;"
    sp = sp & "list|export|1|130|btn.export|ExportVisible|0;"
    sp = sp & "list|recalc|1|80|btn.recalc|RefreshAll|0;"
    sp = sp & "list|home|1|70|btn.home|GoHome|0;"
    sp = sp & "import|pick|1|110|btn.pickFiles|ImportStart|1;"
    sp = sp & "import|recheck|1|90|btn.recheck|ImportRecheckBtn|0;"
    sp = sp & "import|commit|1|120|btn.commit|ImportCommit|1;"
    sp = sp & "import|clear|1|80|btn.clear|ImportClear|0;"
    sp = sp & "import|home|1|70|btn.home|GoHome|0;"
    sp = sp & "mail|body|1|90|btn.copyBody|MailCopyBody|1;"
    sp = sp & "mail|subj|1|90|btn.copySubj|MailCopySubject|0;"
    sp = sp & "mail|to|1|110|btn.copyTo|MailCopyTo|0;"
    sp = sp & "mail|web|1|100|btn.webmail|OpenWebmail|0;"
    sp = sp & "mail|sent|1|120|btn.markSent|MailMarkSent|1;"
    sp = sp & "mail|clear|1|80|btn.clear|MailClearDraftsBtn|0;"
    sp = sp & "mail|home|1|70|btn.home|GoHome|0;"
    sp = sp & "cal|prev|3|80|btn.prev|CalPrev|0;"
    sp = sp & "cal|today|3|70|btn.today|CalToday|0;"
    sp = sp & "cal|next|3|80|btn.next|CalNext|0;"
    sp = sp & "cal|list|3|80|btn.list|GoList|0;"
    sp = sp & "cal|home|3|70|btn.home|GoHome|0;"
    sp = sp & "detail|list|3|110|btn.toList|DetailGoList|0;"
    sp = sp & "detail|nego|3|110|btn.nego|AddNegotiation|0;"
    sp = sp & "detail|att|3|90|btn.att|AddAttachment|0;"
    sp = sp & "detail|mail|3|100|btn.mailOne|MailForSelection|1;"
    sp = sp & "detail|home|3|70|btn.home|GoHome|0;"
    specs = Split(Left$(sp, Len(sp) - 1), ";")
    ' remove old buttons
    For Each k In Array("home", "list", "import", "mail", "cal", "detail")
        Set ws = SheetOf(CStr(k))
        Unprot ws
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
        Next i
    Next k
    For Each s In specs
        p = Split(CStr(s), "|")
        Set ws = SheetOf(p(0))
        key = p(0) & "|" & p(2)
        If CHas(lefts, key) Then
            x = lefts(key)
            lefts.Remove key
        Else
            x = 6
        End If
        top = ws.Rows(CLng(p(2))).Top + 3
        h = ws.Rows(CLng(p(2))).Height - 6
        If h < 16 Then h = 16
        Set shp = ws.Shapes.AddShape(5, x, top, CDbl(p(3)), h)
        lefts.Add x + CDbl(p(3)) + 6, key
        shp.Name = "btn_" & p(0) & "_" & p(1)
        shp.OnAction = p(5)
        shp.Placement = 3
        With shp.TextFrame2
            .TextRange.Text = T(p(4))
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = True
            .VerticalAnchor = 3
            .TextRange.ParagraphFormat.Alignment = 2
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
        If p(6) = "1" Then
            shp.Fill.ForeColor.RGB = RGB(31, 47, 99)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(31, 47, 99)
        End If
        shp.Line.ForeColor.RGB = RGB(31, 47, 99)
        shp.Line.Weight = 1
    Next s
End Sub

Public Sub ImportRecheckBtn()
    Dim n As Long
    n = ImportRecheck
    Msg F("msg.imp.rechecked", n)
End Sub

Public Sub MailClearDraftsBtn(): MailClearDrafts True: End Sub

Public Sub DetailGoList()
    Dim id As String, r As Long
    EnsureEvents
    id = Nz(SheetOf("detail").Range("C2").Value)
    r = FindRowById(id)
    ClearFilters
    ListWs.Activate
    If r > 0 Then ListWs.Cells(r, C_ID).Select
End Sub
