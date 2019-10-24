Sub SaveNewVersion_Excel()
'PURPOSE: Save file, if already exists add a new version indicator to filename
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim FolderPath As String
Dim myPath As String
Dim SaveName As String
Dim SaveExt As String
Dim VersionExt As String
Dim Saved As Boolean
Dim x As Long

TestStr = ""
Saved = False
x = 2

'Version Indicator (change to liking)
  VersionExt = "_v"

'Pull info about file
  On Error GoTo NotSavedYet
    myPath = ActiveWorkbook.FullName
    myFileName = Mid(myPath, InStrRev(myPath, "\") + 1, InStrRev(myPath, ".") - InStrRev(myPath, "\") - 1)
    FolderPath = Left(myPath, InStrRev(myPath, "\"))
    SaveExt = "." & Right(myPath, Len(myPath) - InStrRev(myPath, "."))
  On Error GoTo 0

'Determine Base File Name
  If InStr(1, myFileName, VersionExt) > 1 Then
    myArray = Split(myFileName, VersionExt)
    SaveName = myArray(0)
  Else
    SaveName = myFileName
  End If
    
'Test to see if file name already exists
  If FileExist(FolderPath & SaveName & SaveExt) = False Then
    ActiveWorkbook.SaveAs FolderPath & SaveName & SaveExt
    Exit Sub
  End If
      
'Need a new version made
  Do While Saved = False
    If FileExist(FolderPath & SaveName & VersionExt & x & SaveExt) = False Then
      ActiveWorkbook.SaveAs FolderPath & SaveName & VersionExt & x & SaveExt
      Saved = True
    Else
      x = x + 1
    End If
  Loop

'New version saved
  MsgBox "New file version saved (version " & x & ")"

Exit Sub

'Error Handler
NotSavedYet:
  MsgBox "This file has not been initially saved. " & _
    "Cannot save a new version!", vbCritical, "Not Saved To Computer"

End Sub