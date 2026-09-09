Private Sub Workbook_Open()
    On Error Resume Next
    Application.EnableEvents = True
    Worksheets("Mistake Log").Activate
End Sub
