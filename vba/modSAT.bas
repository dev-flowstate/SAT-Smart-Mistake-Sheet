Option Explicit

Public Const HDR_ROW As Long = 4
Public Const FIRST_ROW As Long = 5
Public Const PATH_COL As Long = 17
Public Const BTN_NAME As String = "PhotoBtn"
Public Const PICK_NAME As String = "PickBtn"
Public Const ZOOM_NAME As String = "ZoomPic"
Public Const CELLPIC As String = "CellPic_"

' Set True if in-cell pictures misbehave; forces the classic sized-picture mode.
Public ForceClassicPicture As Boolean
Private gExportBroken As Boolean

Private gItems() As String
Private gCount As Long
Private gTarget As Range

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

' Works on Windows and Mac; falls back to the workbook folder if neither is writable.
Public Function StoreDir() As String
    Dim base As String, sep As String, d As String
    sep = Application.PathSeparator
    base = Environ$("USERPROFILE")
    If Len(base) = 0 Then base = Environ$("HOME")
    If Len(base) = 0 Then base = ThisWorkbook.Path
    d = base & sep & "Pictures" & sep & "SAT Question Screenshots"
    On Error Resume Next
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    On Error GoTo 0
    If Len(Dir(d, vbDirectory)) = 0 Then d = ThisWorkbook.Path
    StoreDir = d
End Function

Public Function HasList(c As Range) As Boolean
    Dim t As Long
    On Error Resume Next
    t = 0
    t = c.Validation.Type
    HasList = (t = 3)
    On Error GoTo 0
End Function

Public Function ZoomOpen() As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = LogSheet().Shapes(ZOOM_NAME)
    ZoomOpen = Not (s Is Nothing)
    On Error GoTo 0
End Function

Private Function IsHelperShape(nm As String) As Boolean
    IsHelperShape = (nm = ZOOM_NAME Or nm = BTN_NAME Or nm = PICK_NAME Or _
                     Left$(nm, Len(CELLPIC)) = CELLPIC)
End Function

' ---------- dropdown list, centred in the cell ----------
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

Public Sub ShowListPopup()
    Dim cb As CommandBar, ci As CommandBarButton, i As Long, n As Long
    On Error GoTo Oops
    n = BuildPickList(ActiveCell)
    If n = 0 Then
        MsgBox "No options to show yet." & vbCrLf & vbCrLf & _
               "If this is the Topic column, choose a Section in this row first.", _
               vbInformation, "Nothing to choose"
        Exit Sub
    End If
    On Error Resume Next
    Application.CommandBars("SATPick").Delete
    On Error GoTo Oops
    Set cb = Application.CommandBars.Add("SATPick", msoBarPopup, False, True)
    For i = 0 To n - 1
        Set ci = cb.Controls.Add(msoControlButton)
        ci.Caption = gItems(i)
        ci.Style = msoButtonCaption
        ci.OnAction = "PickByIndex " & i
    Next i
    Set ci = cb.Controls.Add(msoControlButton)
    ci.Caption = "(clear this cell)"
    ci.Style = msoButtonCaption
    ci.BeginGroup = True
    ci.OnAction = "PickByIndex -1"
    cb.ShowPopup
    Exit Sub
Oops:
    MsgBox "Could not open the list here." & vbCrLf & _
           "You can still type the value, or use the small arrow at the cell edge.", _
           vbInformation, "SAT Log"
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

' ---------- photos ----------
Private Sub RemoveCellPic(ws As Worksheet, r As Long)
    On Error Resume Next
    ws.Shapes(CELLPIC & r).Delete
    On Error GoTo 0
End Sub

' Puts the picture in the cell. Uses the modern in-cell picture where Excel
' supports it, otherwise a normal picture shrunk to fit and locked to the cell.
Public Sub InsertShot(ws As Worksheet, r As Long, sc As Long, f As String)
    Dim rg As Object, c As Range, p As Shape, k As Double, okInCell As Boolean
    Set c = ws.Cells(r, sc)
    On Error Resume Next
    c.ClearContents
    On Error GoTo 0
    RemoveCellPic ws, r

    okInCell = False
    If Not ForceClassicPicture Then
        On Error Resume Next
        Set rg = c
        rg.InsertPictureInCell f
        okInCell = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If Not okInCell Then
        On Error Resume Next
        Set p = ws.Shapes.AddPicture(f, msoFalse, msoCTrue, c.Left + 2, c.Top + 2, -1, -1)
        If Not p Is Nothing Then
            p.Name = CELLPIC & r
            p.LockAspectRatio = msoTrue
            k = Application.Min((c.Width - 4) / p.Width, (c.Height - 4) / p.Height)
            If k > 0 Then p.Width = p.Width * k
            p.Left = c.Left + (c.Width - p.Width) / 2
            p.Top = c.Top + (c.Height - p.Height) / 2
            p.Placement = 1
            p.OnAction = "ZoomFromPicture"
        End If
        On Error GoTo 0
    End If
    ws.Cells(r, PATH_COL).Value = f
End Sub

' clicking a fallback picture zooms it, same as the in-cell version
Public Sub ZoomFromPicture()
    Dim nm As String, r As Long
    On Error Resume Next
    nm = Application.Caller
    If Left$(nm, Len(CELLPIC)) = CELLPIC Then
        r = CLng(Mid$(nm, Len(CELLPIC) + 1))
        ZoomPhoto r
    End If
End Sub

Private Function ExportShape(ws As Worksheet, sh As Shape, f As String) As Boolean
    Dim co As ChartObject
    On Error GoTo Fail
    sh.Copy
    Set co = ws.ChartObjects.Add(0, 0, sh.Width, sh.Height)
    co.Chart.Paste
    co.Chart.Export f, "PNG"
    co.Delete
    ExportShape = True
    Exit Function
Fail:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    ExportShape = False
End Function

Public Function FitFloating(ws As Worksheet) As Long
    Dim i As Long, sh As Shape, r As Long, f As String, sc As Long, n As Long
    If gExportBroken Then Exit Function
    sc = ShotCol(ws)
    If sc = 0 Then Exit Function
    On Error GoTo Done
    Application.EnableEvents = False
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If sh.Type = msoPicture Then
            If Not IsHelperShape(sh.Name) Then
                r = sh.TopLeftCell.Row
                If r < FIRST_ROW Then r = FIRST_ROW
                f = StoreDir() & Application.PathSeparator & "q_" & _
                    Format(Now, "yyyymmdd_hhnnss") & "_" & i & ".png"
                If ExportShape(ws, sh, f) Then
                    sh.Delete
                    InsertShot ws, r, sc, f
                    n = n + 1
                Else
                    gExportBroken = True     ' this Excel cannot export; stop retrying
                    GoTo Done
                End If
            End If
        End If
    Next i
Done:
    Application.EnableEvents = True
    FitFloating = n
End Function

Public Sub ZoomPhoto(r As Long)
    Dim ws As Worksheet, f As String, p As Shape, k As Double, btn As Shape
    On Error GoTo Oops
    Set ws = LogSheet()
    f = CStr(ws.Cells(r, PATH_COL).Value)
    If Len(f) = 0 Then
        MsgBox "No picture saved for this row yet." & vbCrLf & vbCrLf & _
               "Select the Question Screenshot cell and click Add Photo.", _
               vbInformation, "Nothing to zoom"
        Exit Sub
    End If
    If Len(Dir(f)) = 0 Then
        MsgBox "The picture file for this row was moved or deleted:" & vbCrLf & f, _
               vbExclamation, "File not found"
        Exit Sub
    End If
    CloseZoom
    Set p = ws.Shapes.AddPicture(f, msoFalse, msoCTrue, 0, 0, -1, -1)
    p.Name = ZOOM_NAME
    p.LockAspectRatio = msoTrue
    k = Application.Min(ActiveWindow.UsableWidth * 0.88 / p.Width, _
                        ActiveWindow.UsableHeight * 0.88 / p.Height)
    If k < 1 Then p.Width = p.Width * k
    p.Top = ActiveWindow.VisibleRange.Top + (ActiveWindow.UsableHeight - p.Height) / 2
    p.Left = ActiveWindow.VisibleRange.Left + (ActiveWindow.UsableWidth - p.Width) / 2
    p.OnAction = "CloseZoom"
    p.ZOrder msoBringToFront
    On Error Resume Next
    Set btn = ws.Shapes(BTN_NAME)
    If Not btn Is Nothing Then
        btn.Visible = msoTrue
        btn.Width = 60
        btn.Height = 19
        btn.Left = p.Left + p.Width - btn.Width - 6
        btn.Top = p.Top + 6
        btn.TextFrame2.TextRange.Text = "Close"
        btn.ZOrder msoBringToFront
    End If
    ws.Shapes(PICK_NAME).Visible = msoFalse
    On Error GoTo 0
    Application.StatusBar = "Click Close, or the picture itself, to put it back in the cell."
    Exit Sub
Oops:
    Application.StatusBar = False
    MsgBox "Could not open that picture: " & Err.Description, vbExclamation, "SAT Log"
End Sub

Public Sub CloseZoom()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = LogSheet()
    ws.Shapes(ZOOM_NAME).Delete
    Application.StatusBar = False
    PositionButtons ws, Selection
End Sub

Private Function PickFile() As String
    Dim fd As FileDialog
    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd Is Nothing Then
        PickFile = CStr(Application.GetOpenFilename("Images, *.png;*.jpg;*.jpeg;*.gif;*.bmp"))
        If PickFile = "False" Then PickFile = ""
        Exit Function
    End If
    fd.Title = "Choose the question screenshot"
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.png; *.jpg; *.jpeg; *.gif; *.bmp"
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Public Sub PhotoButtonClick()
    Dim ws As Worksheet, sc As Long, r As Long, f As String
    On Error GoTo Oops
    Set ws = LogSheet()
    If ZoomOpen() Then
        CloseZoom
        Exit Sub
    End If
    sc = ShotCol(ws)
    r = ActiveCell.Row
    If r < FIRST_ROW Then Exit Sub
    If Len(CStr(ws.Cells(r, PATH_COL).Value)) > 0 Then
        ZoomPhoto r
        Exit Sub
    End If
    If FitFloating(ws) > 0 Then
        PositionButtons ws, ActiveCell
        Exit Sub
    End If
    f = PickFile()
    If Len(f) = 0 Then Exit Sub
    InsertShot ws, r, sc, f
    PositionButtons ws, ActiveCell
    Exit Sub
Oops:
    Application.EnableEvents = True
    MsgBox "Could not add that picture: " & Err.Description, vbExclamation, "SAT Log"
End Sub

Public Sub PositionButtons(ws As Worksheet, Target As Range)
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

    ok = True
    If Target Is Nothing Then ok = False
    If ok Then If Target.Cells.Count > 1 Then ok = False
    If ok Then If Target.Row < FIRST_ROW Then ok = False

    If Not photo Is Nothing Then
        If ok And Target.Column = sc Then
            Set c = ws.Cells(Target.Row, sc)
            photo.Visible = msoTrue
            If Len(CStr(ws.Cells(Target.Row, PATH_COL).Value)) > 0 Then
                photo.TextFrame2.TextRange.Text = "Zoom In"
            Else
                photo.TextFrame2.TextRange.Text = "Add Photo"
            End If
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

Public Sub SetClassicMode(ByVal onOff As Boolean)
    ' Force the old-Excel picture mode (also used to test that fallback).
    ForceClassicPicture = onOff
End Sub
