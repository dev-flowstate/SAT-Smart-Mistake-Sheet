Private Sub Workbook_Open()
    On Error Resume Next
    Application.EnableEvents = True
    Worksheets("Mistake Log").Activate
    CloseZoomQuiet
    ArmHook
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    ArmHook
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    CloseZoomQuiet
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    CloseZoomQuiet
End Sub
