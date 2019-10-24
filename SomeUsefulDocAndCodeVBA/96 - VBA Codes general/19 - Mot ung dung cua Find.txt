Sub DeleteUserWithConfirmation()
Dim oFound As Range
Dim oLookin As Range
Dim sLookFor As String
SearchUserName = InputBox("Please input a user")
sLookFor = SearchUserName
Set oLookin = Worksheets("Sheet13").UsedRange
Set oFound = oLookin.Find(what:=sLookFor, _
LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
If Not oFound Is Nothing Then
uChoice = MsgBox("Are you sure you want to delete the user: " & _
vbNewLine & sLookFor, vbQuestion + vbYesNo, "Delete Confirmation")
If uChoice = vbYes Then
oLookin.Range("A" & oFound.Row).EntireRow.Delete
MsgBox "The user: " & vbNewLine & sLookFor & _
vbNewLine & "has been deleted successfully."
End If
Else
MsgBox "Nothing found for : " & vbNewLine & _
sLookFor
End If
End Sub