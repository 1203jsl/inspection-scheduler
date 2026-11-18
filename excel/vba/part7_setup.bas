
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
