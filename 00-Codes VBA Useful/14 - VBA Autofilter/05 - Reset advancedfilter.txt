Sub resetFilters()

Dim sht As Worksheet
Dim LastRow As Range

Application.ScreenUpdating = False

   On Error Resume Next
        If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
  End If

'Range("A3:T3").ClearContents
Application.ScreenUpdating = True
Call GetLastRow

End Sub
