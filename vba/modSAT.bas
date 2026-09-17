' SAT Smart Mistake Sheet
' Made by Muhammad Salar Khan
' Instagram: @salars_catalogue  https://www.instagram.com/salars_catalogue/
' Updates:   https://github.com/dev-flowstate/SAT-Smart-Mistake-Sheet

Option Explicit

Public Const HDR_ROW As Long = 4
Public Const FIRST_ROW As Long = 5
Public Const BTN_NAME As String = "PhotoBtn"
Public Const PICK_NAME As String = "PickBtn"
Public Const ZOOM_NAME As String = "ZoomPic"
Public Const CELLPIC As String = "CellPic_"

Private gItems() As String
Private gCount As Long
Private gTarget As Range
Private gPhotoRow As Long
Private gLastShapeCount As Long
Private gSeq As Long
Private gHook As CUIHook
Private gLastAdoptRow As Long

' =====================================================================
'  helpers
' =====================================================================
Public Function LogSheet() As Worksheet
    Set LogSheet = ThisWorkbook.Worksheets("Mistake Log")
End Function

Public Function ShotCol(ws As Worksheet) As Long
    Dim c As Long
    For c = 1 To 40
        If InStr(1, CStr(ws.Cells(HDR_ROW, c).Value), "Question Screenshot", vbTextCompare) = 1 Then
            ShotCol = c
            Exit Function
        End If
    Next c
End Function

' Fully qualified name for menus and keys built at run time. They are never
' saved, so renaming the file cannot break them, and a second open copy of
' the log can never answer a click meant for this one.
Private Function MacroRef(ByVal procName As String) As String
    MacroRef = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

Public Function HasList(c As Range) As Boolean
    Dim t As Long
    On Error Resume Next
    t = 0
    t = c.Validation.Type
    HasList = (t = 3)
    On Error GoTo 0
End Function

Private Function IsHelperShape(ByVal nm As String) As Boolean
    IsHelperShape = (nm = ZOOM_NAME Or nm = BTN_NAME Or nm = PICK_NAME Or _
                     Left$(nm, Len(CELLPIC)) = CELLPIC)
End Function

Public Function ZoomOpen() As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = LogSheet().Shapes(ZOOM_NAME)
    ZoomOpen = Not (s Is Nothing)
    On Error GoTo 0
End Function

' A row's photo is found by WHERE it sits, not by its name, so photos stay
' matched to their row after sorting or filtering.
Public Function PhotoShapeForRow(ws As Worksheet, ByVal r As Long) As Shape
    Dim sh As Shape, sc As Long, c As Range, midX As Double, midY As Double
    sc = ShotCol(ws)
    If sc = 0 Then Exit Function
    Set c = ws.Cells(r, sc)
    For Each sh In ws.Shapes
        If Left$(sh.Name, Len(CELLPIC)) = CELLPIC Then
            midX = sh.Left + sh.Width / 2
            midY = sh.Top + sh.Height / 2
            If midY >= c.Top And midY < c.Top + c.Height And _
               midX >= c.Left And midX < c.Left + c.Width Then
                Set PhotoShapeForRow = sh
                Exit Function
            End If
        End If
    Next sh
End Function

Private Function RowOfPhoto(ws As Worksheet, sh As Shape) As Long
    Dim r As Long, midY As Double
    midY = sh.Top + sh.Height / 2
    r = sh.TopLeftCell.Row
    Do While r < sh.BottomRightCell.Row
        If ws.Rows(r).Top + ws.Rows(r).Height > midY Then Exit Do
        r = r + 1
    Loop
    RowOfPhoto = r
End Function

Public Function PhotoCount() As Long
    Dim sh As Shape, n As Long
    For Each sh In LogSheet().Shapes
        If Left$(sh.Name, Len(CELLPIC)) = CELLPIC Then n = n + 1
    Next sh
    PhotoCount = n
End Function

' =====================================================================
'  dropdown list, centred in the cell
' =====================================================================
Public Function BuildPickList(c As Range) As Long
    Dim f As String, rng As Range, v As Variant, i As Long, parts() As String
    gCount = 0
    Erase gItems
    Set gTarget = c
    On Error Resume Next
    f = c.Validation.Formula1
    On Error GoTo 0
    If Len(f) = 0 Then Exit Function
    If Left$(f, 1) = "=" Then f = Mid$(f, 2)
    On Error Resume Next
    Set rng = Application.Evaluate(f)
    On Error GoTo 0
    If rng Is Nothing Then
        If InStr(f, ",") > 0 Then
            parts = Split(Replace(f, Chr$(34), ""), ",")
            ReDim gItems(0 To UBound(parts))
            For i = 0 To UBound(parts)
                gItems(i) = Trim$(parts(i))
            Next i
            gCount = UBound(parts) + 1
        End If
        BuildPickList = gCount
        Exit Function
    End If
    ReDim gItems(0 To rng.Cells.Count - 1)
    i = 0
    For Each v In rng.Cells
        If Len(CStr(v.Value)) > 0 Then
            gItems(i) = CStr(v.Value)
            i = i + 1
        End If
    Next v
    gCount = i
    BuildPickList = gCount
End Function

' Builds the popup for the active cell. Kept separate from ShowListPopup so a
' test can build it and fire each entry with .Execute() - ShowPopup is modal.
Public Function BuildPopupBar() As Long
    Dim cb As CommandBar, ci As CommandBarButton, i As Long, n As Long
    On Error GoTo Oops
    n = BuildPickList(ActiveCell)
    If n = 0 Then Exit Function
    On Error Resume Next
    Application.CommandBars("SATPick").Delete
    On Error GoTo Oops
    Set cb = Application.CommandBars.Add("SATPick", msoBarPopup, False, True)
    For i = 0 To n - 1
        Set ci = cb.Controls.Add(msoControlButton)
        ' "&" marks a menu shortcut key, so a literal & must be doubled
        ci.Caption = Replace(gItems(i), "&", "&&")
        ci.Style = msoButtonCaption
        ci.Tag = CStr(i)
        ci.OnAction = MacroRef("PickFromMenu")
    Next i
    Set ci = cb.Controls.Add(msoControlButton)
    ci.Caption = "(clear this cell)"
    ci.Style = msoButtonCaption
    ci.BeginGroup = True
    ci.Tag = "-1"
    ci.OnAction = MacroRef("PickFromMenu")
    BuildPopupBar = n
    Exit Function
Oops:
    BuildPopupBar = -1
End Function

Public Sub ShowListPopup()
    Dim n As Long
    On Error GoTo Oops
    n = BuildPopupBar()
    If n = 0 Then
        MsgBox "No options to show yet." & vbCrLf & vbCrLf & _
               "If this is the Topic column, choose a Section in this row first.", _
               vbInformation, "Nothing to choose"
        Exit Sub
    End If
    If n < 0 Then GoTo Oops
    Application.CommandBars("SATPick").ShowPopup
    Exit Sub
Oops:
    MsgBox "Could not open the list here." & vbCrLf & _
           "You can still use the small arrow at the right edge of the cell.", _
           vbInformation, "SAT Log"
End Sub

' Menu entries carry no argument - the index rides on the control's Tag - so
' there is no OnAction quoting to get wrong.
Public Sub PickFromMenu()
    Dim idx As Long
    idx = -999
    On Error Resume Next
    idx = CLng(Application.CommandBars.ActionControl.Tag)
    On Error GoTo 0
    If idx = -999 Then Exit Sub
    PickByIndex idx
End Sub

Public Sub PickByIndex(ByVal idx As Long)
    On Error GoTo Oops
    If gTarget Is Nothing Then Exit Sub
    Application.EnableEvents = False
    If idx < 0 Then
        gTarget.ClearContents
    Else
        gTarget.Value = gItems(idx)
    End If
    Application.EnableEvents = True
    PositionButtons LogSheet(), gTarget
    Exit Sub
Oops:
    Application.EnableEvents = True
End Sub

' =====================================================================
'  photos
' =====================================================================
' Fits a picture inside a cell and locks it there. Any photo already on that
' row is replaced. The picture is embedded in the workbook, so it travels
' with the file when it is emailed.
Private Sub PlaceShapeInCell(ws As Worksheet, sh As Shape, ByVal r As Long, ByVal sc As Long)
    Dim c As Range, k As Double, old As Shape, ev As Boolean, guard As Long
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set c = ws.Cells(r, sc)
    gSeq = gSeq + 1
    sh.Name = "SATPlacing" & gSeq          ' so the loop below cannot find it
    Do
        Set old = PhotoShapeForRow(ws, r)
        If old Is Nothing Then Exit Do
        old.Delete
        guard = guard + 1
        If guard > 20 Then Exit Do
    Loop
    On Error Resume Next
    c.ClearContents                        ' also clears any Place-in-Cell picture
    sh.Placement = 3
    sh.LockAspectRatio = msoTrue
    sh.ScaleHeight 1, msoTrue              ' back to natural size before fitting
    sh.ScaleWidth 1, msoTrue
    On Error GoTo 0
    k = Application.Min((c.Width - 6) / sh.Width, (c.Height - 6) / sh.Height)
    If k > 0 Then sh.Width = sh.Width * k
    sh.Left = c.Left + (c.Width - sh.Width) / 2
    sh.Top = c.Top + (c.Height - sh.Height) / 2
    sh.Placement = 1                       ' move and size with cells
    sh.Name = CELLPIC & Format$(Now, "yyyymmddhhnnss") & "_" & gSeq
    sh.OnAction = "ZoomFromPicture"
    On Error Resume Next
    sh.AlternativeText = "Question screenshot"
    On Error GoTo 0
    Application.EnableEvents = ev
End Sub

Public Sub AddPhotoFile(ByVal r As Long, ByVal f As String)
    Dim ws As Worksheet, sc As Long, sh As Shape
    Set ws = LogSheet()
    sc = ShotCol(ws)
    If sc = 0 Or r < FIRST_ROW Then Exit Sub
    Set sh = ws.Shapes.AddPicture(f, msoFalse, msoCTrue, ws.Cells(r, sc).Left, ws.Cells(r, sc).Top, -1, -1)
    PlaceShapeInCell ws, sh, r, sc
    gLastShapeCount = ws.Shapes.Count
End Sub

Public Sub RemovePhotoRow(ByVal r As Long)
    Dim ws As Worksheet, sh As Shape, ev As Boolean, sc As Long, guard As Long
    Set ws = LogSheet()
    sc = ShotCol(ws)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    CloseZoomQuiet
    Do
        Set sh = Nothing
        Set sh = PhotoShapeForRow(ws, r)
        If sh Is Nothing Then Exit Do
        sh.Delete
        guard = guard + 1
        If guard > 20 Then Exit Do
    Loop
    ws.Cells(r, sc).ClearContents
    Err.Clear
    On Error GoTo 0
    Application.EnableEvents = ev
    gLastShapeCount = ws.Shapes.Count
End Sub

' Pictures dropped anywhere on a row (right-click Paste, drag in) are pulled
' into that row's screenshot cell the next time the selection moves.
Public Function AdoptLoosePictures(ws As Worksheet) As Long
    Dim sh As Shape, sc As Long, r As Long, n As Long, i As Long
    Dim names As New Collection, rows As New Collection
    If ws.Shapes.Count = gLastShapeCount Then Exit Function
    sc = ShotCol(ws)
    If sc = 0 Then Exit Function
    For Each sh In ws.Shapes
        If sh.Type = msoPicture Then
            If Not IsHelperShape(sh.Name) Then
                r = sh.TopLeftCell.Row
                If r >= FIRST_ROW Then
                    names.Add sh.Name
                    rows.Add r
                    gLastAdoptRow = r
                End If
            End If
        End If
    Next sh
    On Error Resume Next
    For i = 1 To names.Count
        PlaceShapeInCell ws, ws.Shapes(names(i)), rows(i), sc
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next i
    On Error GoTo 0
    gLastShapeCount = ws.Shapes.Count
    AdoptLoosePictures = n
End Function

' ---------- Ctrl+V ----------
' Excel pastes the picture normally. A listener INSIDE this workbook notices the
' new picture as soon as it appears and fits it into its row's cell.
' No Application.OnKey: that switch lives on Excel itself, so if it were ever
' left on, Ctrl+V in any other workbook would reopen this file.
Public Sub ArmHook()
    On Error Resume Next
    If gHook Is Nothing Then
        Set gHook = New CUIHook
        Set gHook.Bars = Application.CommandBars
    End If
End Sub

' Called on every UI refresh, so it must be cheap: one count comparison.
Public Sub OnUiUpdate()
    Static busy As Boolean
    Dim ws As Worksheet, sc As Long
    If busy Then Exit Sub
    On Error GoTo Done
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Name <> "Mistake Log" Then Exit Sub
    Set ws = ActiveSheet
    If ws.Shapes.Count = gLastShapeCount Then Exit Sub
    busy = True
    If AdoptLoosePictures(ws) > 0 Then
        sc = ShotCol(ws)
        ws.Cells(gLastAdoptRow, sc).Select
        PositionButtons ws, ws.Cells(gLastAdoptRow, sc)
    End If
Done:
    busy = False
End Sub

' ---------- Delete key ----------
' Pressing Delete on a Question Screenshot cell removes its photo. Only for a
' single row, so sorting or a bulk paste can never wipe photos.
Public Sub PhotoCellEdited(ws As Worksheet, Target As Range)
    Dim sc As Long, hit As Range, cell As Range
    sc = ShotCol(ws)
    If sc = 0 Then Exit Sub
    If Target.Rows.Count > 1 Then Exit Sub
    Set hit = Intersect(Target, ws.Columns(sc))
    If hit Is Nothing Then Exit Sub
    Set cell = hit.Cells(1, 1)
    If cell.Row < FIRST_ROW Then Exit Sub
    If Len(CStr(cell.Value)) > 0 Then Exit Sub
    If PhotoShapeForRow(ws, cell.Row) Is Nothing Then Exit Sub
    RemovePhotoRow cell.Row
    PositionButtons ws, Selection
End Sub

' ---------- zoom ----------
Public Sub ZoomRow(ByVal r As Long)
    Dim ws As Worksheet, src As Shape
    Set ws = LogSheet()
    Set src = PhotoShapeForRow(ws, r)
    If src Is Nothing Then
        MsgBox "There is no photo on this row yet." & vbCrLf & vbCrLf & _
               "Click the Question Screenshot cell and use Add Photo, " & _
               "or copy a screenshot and press Ctrl+V.", vbInformation, "Nothing to zoom"
        Exit Sub
    End If
    ZoomShape ws, src
End Sub

' Enlarges a copy of the embedded picture - no file on disk is needed, so zoom
' still works after the log has been emailed to someone else.
Private Sub ZoomShape(ws As Worksheet, src As Shape)
    Dim z As Shape, k As Double, ev As Boolean, btn As Shape, vr As Range
    Dim zf As Double, vw As Double, vh As Double, topOff As Double
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Oops
    CloseZoomQuiet
    Set z = src.Duplicate
    z.Name = ZOOM_NAME
    z.Placement = 3
    z.LockAspectRatio = msoTrue
    On Error Resume Next
    z.ScaleHeight 1, msoTrue
    z.ScaleWidth 1, msoTrue
    On Error GoTo Oops
    zf = ActiveWindow.Zoom / 100
    If zf <= 0 Then zf = 1
    If ActiveWindow.FreezePanes And ActiveWindow.SplitRow > 0 Then
        topOff = ws.Range(ws.Rows(1), ws.Rows(ActiveWindow.SplitRow)).Height
    End If
    vw = ActiveWindow.UsableWidth / zf
    vh = ActiveWindow.UsableHeight / zf - topOff
    k = Application.Min(vw * 0.86 / z.Width, vh * 0.82 / z.Height, 3)
    z.Width = z.Width * k
    Set vr = ActiveWindow.VisibleRange
    z.Left = vr.Left + (vw - z.Width) / 2
    z.Top = vr.Top + (vh - z.Height) / 2
    If z.Top < vr.Top Then z.Top = vr.Top
    If z.Left < vr.Left Then z.Left = vr.Left
    z.Line.Visible = msoTrue
    z.Line.ForeColor.RGB = RGB(37, 99, 235)
    z.Line.Weight = 3
    z.OnAction = "CloseZoom"
    z.ZOrder msoBringToFront
    On Error Resume Next
    Set btn = ws.Shapes(BTN_NAME)
    If Not btn Is Nothing Then
        btn.Visible = msoTrue
        btn.TextFrame2.TextRange.Text = "Close"
        btn.Width = 64
        btn.Height = 20
        btn.Left = z.Left + z.Width - btn.Width - 8
        btn.Top = z.Top + 8
        btn.ZOrder msoBringToFront
    End If
    ws.Shapes(PICK_NAME).Visible = msoFalse
    On Error GoTo 0
    Application.StatusBar = "Click Close, or click the picture, to put it back."
    Application.EnableEvents = ev
    Exit Sub
Oops:
    Application.EnableEvents = ev
    Application.StatusBar = False
End Sub

Public Sub CloseZoomQuiet()
    On Error Resume Next
    LogSheet().Shapes(ZOOM_NAME).Delete
    Application.StatusBar = False
End Sub

Public Sub CloseZoom()
    CloseZoomQuiet
    On Error Resume Next
    If ActiveSheet.Name = LogSheet().Name Then PositionButtons LogSheet(), Selection
End Sub

' Clicking a photo selects its cell (so its button shows) and enlarges it.
Public Sub ZoomFromPicture()
    Dim ws As Worksheet, sh As Shape, r As Long
    On Error GoTo Quiet
    Set ws = LogSheet()
    Set sh = ws.Shapes(CStr(Application.Caller))
    r = RowOfPhoto(ws, sh)
    If r >= FIRST_ROW Then ws.Cells(r, ShotCol(ws)).Select
    ZoomShape ws, sh
Quiet:
End Sub

' ---------- the photo button and its menu ----------
Private Function PickFile() As String
    Dim fd As FileDialog, v As Variant
    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd Is Nothing Then
        v = Application.GetOpenFilename("Images (*.png;*.jpg;*.jpeg;*.gif;*.bmp),*.png;*.jpg;*.jpeg;*.gif;*.bmp")
        If VarType(v) = vbString Then PickFile = v
        Exit Function
    End If
    fd.Title = "Choose the question screenshot"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.png; *.jpg; *.jpeg; *.gif; *.bmp"
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Public Sub PhotoButtonClick()
    Dim ws As Worksheet, r As Long, f As String
    On Error GoTo Oops
    Set ws = LogSheet()
    If ZoomOpen() Then
        CloseZoom
        Exit Sub
    End If
    AdoptLoosePictures ws
    r = ActiveCell.Row
    If r < FIRST_ROW Then Exit Sub
    If Not PhotoShapeForRow(ws, r) Is Nothing Then
        ShowPhotoMenu
        Exit Sub
    End If
    f = PickFile()
    If Len(f) = 0 Then Exit Sub
    AddPhotoFile r, f
    PositionButtons ws, ws.Cells(r, ShotCol(ws))
    Exit Sub
Oops:
    Application.EnableEvents = True
    MsgBox "Could not add that picture: " & Err.Description, vbExclamation, "SAT Log"
End Sub

Private Sub AddMenuItem(cb As CommandBar, ByVal cap As String, ByVal proc As String, ByVal newGroup As Boolean)
    Dim ci As CommandBarButton
    Set ci = cb.Controls.Add(msoControlButton)
    ci.Caption = cap
    ci.Style = msoButtonCaption
    ci.BeginGroup = newGroup
    ci.OnAction = MacroRef(proc)
End Sub

Public Function BuildPhotoMenu() As Long
    Dim cb As CommandBar
    gPhotoRow = ActiveCell.Row
    On Error Resume Next
    Application.CommandBars("SATPhoto").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add("SATPhoto", msoBarPopup, False, True)
    AddMenuItem cb, "Zoom In", "ZoomMenuRow", False
    AddMenuItem cb, "Replace Photo...", "ReplaceMenuRow", False
    AddMenuItem cb, "Remove Photo", "RemoveMenuRow", True
    BuildPhotoMenu = cb.Controls.Count
End Function

Public Sub ShowPhotoMenu()
    On Error GoTo Oops
    If BuildPhotoMenu() > 0 Then Application.CommandBars("SATPhoto").ShowPopup
    Exit Sub
Oops:
    MsgBox "Could not open the photo menu: " & Err.Description, vbExclamation, "SAT Log"
End Sub

Public Sub ZoomMenuRow()
    ZoomRow gPhotoRow
End Sub

Public Sub ReplaceMenuRow()
    Dim f As String
    If gPhotoRow < FIRST_ROW Then Exit Sub
    f = PickFile()
    If Len(f) = 0 Then Exit Sub            ' cancelled: the old photo stays
    AddPhotoFile gPhotoRow, f               ' replaces whatever was on the row
    PositionButtons LogSheet(), LogSheet().Cells(gPhotoRow, ShotCol(LogSheet()))
End Sub

Public Sub RemoveMenuRow()
    If gPhotoRow < FIRST_ROW Then Exit Sub
    RemovePhotoRow gPhotoRow
    PositionButtons LogSheet(), LogSheet().Cells(gPhotoRow, ShotCol(LogSheet()))
End Sub

' =====================================================================
'  keep both buttons centred on the selected cell
' =====================================================================
Public Sub PositionButtons(ws As Worksheet, Target As Object)
    Dim sc As Long, photo As Shape, pick As Shape, c As Range, lbl As String
    Dim ok As Boolean
    On Error GoTo Quiet
    sc = ShotCol(ws)
    On Error Resume Next
    Set photo = ws.Shapes(BTN_NAME)
    Set pick = ws.Shapes(PICK_NAME)
    On Error GoTo Quiet

    If ZoomOpen() Then
        If Not pick Is Nothing Then pick.Visible = msoFalse
        Exit Sub
    End If

    ok = (TypeName(Target) = "Range")
    If ok Then ok = (Target.Cells.Count = 1)
    If ok Then ok = (Target.Row >= FIRST_ROW)

    If Not photo Is Nothing Then
        If ok And Target.Column = sc Then
            Set c = ws.Cells(Target.Row, sc)
            If PhotoShapeForRow(ws, Target.Row) Is Nothing Then
                photo.TextFrame2.TextRange.Text = "Add Photo"
            Else
                photo.TextFrame2.TextRange.Text = "Photo " & ChrW(9660)
            End If
            photo.Visible = msoTrue
            photo.Width = 76
            photo.Height = 19
            photo.Left = c.Left + (c.Width - photo.Width) / 2
            photo.Top = c.Top + (c.Height - photo.Height) / 2
            photo.ZOrder msoBringToFront
        Else
            photo.Visible = msoFalse
        End If
    End If

    If Not pick Is Nothing Then
        If ok And Target.Column <> sc And HasList(Target) Then
            Set c = Target
            lbl = CStr(c.Value)
            If Len(lbl) = 0 Then
                lbl = "Choose " & ChrW(9660)
            Else
                If Len(lbl) > 20 Then lbl = Left$(lbl, 18) & ".."
                lbl = lbl & "  " & ChrW(9660)
            End If
            pick.Visible = msoTrue
            pick.TextFrame2.TextRange.Text = lbl
            pick.Width = Application.Min(c.Width - 6, 120)
            pick.Height = 19
            pick.Left = c.Left + (c.Width - pick.Width) / 2
            pick.Top = c.Top + (c.Height - pick.Height) / 2
            pick.ZOrder msoBringToFront
        Else
            pick.Visible = msoFalse
        End If
    End If
Quiet:
End Sub
