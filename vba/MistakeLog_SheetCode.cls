Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Me.Shapes.Count > 2 Then FitFloating Me
    PositionButtons Me, Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    If Target.Column = ShotCol(Me) And Target.Row >= FIRST_ROW Then
        Cancel = True
        If ZoomOpen() Then
            CloseZoom
        Else
            ZoomPhoto Target.Row
        End If
    End If
End Sub
