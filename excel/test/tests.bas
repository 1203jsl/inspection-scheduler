' ---- test stubs (replace workbook-backed helpers) ----
Private gFails As String
Private gPass As Long

Public Function T(ByVal k As String) As String
    Select Case k
        Case "time.pm": T = ChrW(50724) & ChrW(54980)
        Case "time.am": T = ChrW(50724) & ChrW(51204)
        Case "time.half": T = ChrW(48152)
        Case "yes": T = "Y"
        Case "weekdays": T = "Su,Mo,Tu,We,Th,Fr,Sa"
        Case "cf.day": T = "DAY>{0} {1}"
        Case "cf.insp": T = "INSP>{0} {1} {2}"
        Case "cf.vendor": T = "VEND>{0} {1} {2}"
        Case "cf.slot": T = "SLOT{0} {1} {2}"
        Case "cf.qty": T = "QTY>{0} {1}"
        Case "cf.holiday": T = "HOL {0} {1}"
        Case "cf.weekend": T = "WKND {0}"
        Case "w.count": T = "{0}c"
        Case "w.countItems": T = "{0}c({1}i)"
        Case Else: T = "[" & k & "]"
    End Select
End Function

Public Function SetV(ByVal nm As String) As Variant
    Select Case nm
        Case "set_maxPerDay": SetV = 3
        Case "set_maxPerInsp": SetV = 2
        Case "set_slotMin": SetV = 30
        Case "set_maxPerVendor": SetV = 2
        Case "set_maxQty": SetV = 0
        Case "set_weekend": SetV = "N"
        Case Else: SetV = Empty
    End Select
End Function

Private Sub Eq(ByVal name As String, ByVal got As Variant, ByVal want As Variant)
    If CStr(got) = CStr(want) Then
        gPass = gPass + 1
    Else
        gFails = gFails & " | " & name & ": got [" & CStr(got) & "] want [" & CStr(want) & "]"
    End If
End Sub

Public Function RunTests() As String
    Dim ok As Boolean, v As Variant, s As String
    On Error GoTo eh
    ' dates
    v = ParseDateV("2026-11-14", ok): Eq "date1", ok & Format$(v, "yyyy-mm-dd"), "True2026-11-14"
    v = ParseDateV("2026.11.14", ok): Eq "date2", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    v = ParseDateV("2026/11/14", ok): Eq "date3", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    v = ParseDateV("2026. 11. 14.", ok): Eq "date4", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    v = ParseDateV("20261114", ok): Eq "date5", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    v = ParseDateV("2026.13.40", ok): Eq "date6 bad", ok, False
    v = ParseDateV("2026-02-30", ok): Eq "date7 bad", ok, False
    v = ParseDateV("next week", ok): Eq "date8 bad", ok, False
    v = ParseDateV(Empty, ok): Eq "date9 empty", ok & IsEmpty(v), "TrueTrue"
    v = ParseDateV(46340, ok): Eq "date10 serial", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    v = ParseDateV(DateSerial(2026, 11, 14), ok): Eq "date11 date", Format$(v, "yyyy-mm-dd"), "2026-11-14"
    ' times
    Eq "time1", ParseTimeV("13:30", ok), "13:30"
    Eq "time2", ParseTimeV("9:05", ok), "09:05"
    Eq "time3", ParseTimeV(0.5625, ok), "13:30"
    Eq "time4", ParseTimeV("1330", ok), "13:30"
    Eq "time5", ParseTimeV(ChrW(50724) & ChrW(54980) & " 2" & ChrW(49884), ok), "14:00"
    Eq "time6", ParseTimeV(ChrW(50724) & ChrW(51204) & " 9" & ChrW(49884) & " 30" & ChrW(48516), ok), "09:30"
    Eq "time7", ParseTimeV("14" & ChrW(49884), ok), "14:00"
    s = ParseTimeV("25:00", ok): Eq "time8 bad", ok, False
    Eq "time9 empty", ParseTimeV("", ok) & ok, "True"
    Eq "time10 pm", ParseTimeV("2:30 PM", ok), "14:30"
    Eq "time11 num", ParseTimeV(930, ok), "09:30"
    Eq "time12 half", ParseTimeV("2" & ChrW(49884) & ChrW(48152), ok), "02:30"
    ' quantities
    Eq "qty1", ParseQtyV(10, ok), 10
    Eq "qty2", ParseQtyV("1,200", ok), 1200
    Eq "qty3", ParseQtyV("20" & ChrW(44060), ok), 20
    Eq "qty4", ParseQtyV("20 EA", ok), 20
    v = ParseQtyV(0, ok): Eq "qty5 bad", ok, False
    v = ParseQtyV(2.5, ok): Eq "qty6 bad", ok, False
    v = ParseQtyV("ten", ok): Eq "qty7 bad", ok, False
    v = ParseQtyV("2.5", ok): Eq "qty8 bad", ok, False
    ' email
    Eq "mail1", IsEmailOk("kim@hansol.example.com"), True
    Eq "mail2", IsEmailOk("han@seowon"), False
    Eq "mail3", IsEmailOk("a b@c.com"), False
    Eq "mail4", IsEmailOk("@c.com"), False
    ' time minutes
    Eq "min1", TimeMinutes("13:30"), 810
    Eq "min2", TimeMinutes(""), -1
    ' conflicts
    TestConflicts
    RunTests = gPass & " passed" & IIf(gFails = "", "", "; FAIL" & gFails)
    Exit Function
eh:
    RunTests = "ERR " & Err.Number & " " & Err.Description & " (line " & Erl & ") after " & gPass & " passed" & gFails
End Function

Private Sub TestConflicts()
    Dim n As Long, sDate() As Long, sTime() As String, sLoc() As String, sInsp() As String, sVendor() As String
    Dim sQty() As Double, sId() As String, msgs As Variant, d As Long, i As Long
    On Error GoTo tceh
    n = 6
    ReDim sDate(1 To n): ReDim sTime(1 To n): ReDim sLoc(1 To n): ReDim sInsp(1 To n)
    ReDim sVendor(1 To n): ReDim sQty(1 To n): ReDim sId(1 To n)
    d = CLng(DateSerial(2026, 12, 1)) ' Tuesday
    For i = 1 To n
        sId(i) = "ID" & i: sDate(i) = d: sQty(i) = 1
    Next i
    ' same vendor + same place twice = 1 visit
    sVendor(1) = "A": sLoc(1) = "L1": sTime(1) = "10:00": sInsp(1) = "P"
    sVendor(2) = "A": sLoc(2) = "L1": sTime(2) = "10:10": sInsp(2) = "P"
    sVendor(3) = "B": sLoc(3) = "L2": sTime(3) = "11:00": sInsp(3) = "Q"
    sVendor(4) = "C": sLoc(4) = "L3": sTime(4) = "13:00": sInsp(4) = "Q"
    sVendor(5) = "D": sLoc(5) = "L1": sTime(5) = "15:00": sInsp(5) = "R"
    sDate(6) = 0
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)
    ' visits: A@L1, B@L2, C@L3, D@L1 = 4 > 3
    Eq "cf day", InStr(msgs(1), "DAY>3 4c(5i)") > 0, True
    Eq "cf same vendor no slot", InStr(msgs(1), "SLOT") = 0, True
    Eq "cf none for undated", msgs(6), ""
    ' slot: different vendor same place within 30 min
    sTime(5) = "10:20"
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)
    Eq "cf slot", InStr(msgs(5), "SLOT30 10:00 ID1") > 0, True
    ' inspector Q has 2 visits (limit 2) -> fine; add one more
    sInsp(5) = "Q"
    msgs = ComputeConflicts(n, sDate, sTime, sLoc, sInsp, sVendor, sQty, sId)
    Eq "cf insp", InStr(msgs(5), "INSP>2 Q 3c") > 0, True
    Eq "cf insp not P", InStr(msgs(1), "INSP>") = 0, True
    Exit Sub
tceh:
    gFails = gFails & " | TestConflicts ERR " & Err.Description & " line " & Erl
End Sub
