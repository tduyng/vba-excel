Sub DELETE_SHEETS()
    Dim sht As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name <> "Feuil1" Then sht.Delete
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub