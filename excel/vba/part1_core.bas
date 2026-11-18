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
