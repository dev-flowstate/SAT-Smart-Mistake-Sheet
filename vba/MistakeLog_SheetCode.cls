Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    ArmHook
    AdoptLoosePictures Me
    PositionButtons Me, Target
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    PhotoCellEdited Me, Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    If Target.Column = ShotCol(Me) And Target.Row >= FIRST_ROW Then
        If ZoomOpen() Then
            Cancel = True
            CloseZoom
        ElseIf Not PhotoShapeForRow(Me, Target.Row) Is Nothing Then
            Cancel = True
            ZoomRow Target.Row
        End If
    End If
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    ArmHook
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    CloseZoomQuiet
End Sub
