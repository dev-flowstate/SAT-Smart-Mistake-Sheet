Public WithEvents Bars As Office.CommandBars

Private Sub Bars_OnUpdate()
    On Error Resume Next
    OnUiUpdate
End Sub
