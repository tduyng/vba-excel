'// Kiem tra su ton tai cua mot key trong Collection'
Public Function KeyExists(myCol As Collection, ByVal keyCheck As String) As Boolean
    KeyExists = False
    On Error GoTo EndFunction
    myCol.Item keyCheck
    KeyExists = True
EndFunction:
End Function