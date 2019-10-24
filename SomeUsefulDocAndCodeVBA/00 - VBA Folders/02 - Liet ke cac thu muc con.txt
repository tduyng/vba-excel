Sub FolderNames()
'Update 20141027
Application.ScreenUpdating = False
Dim xPath As String
Dim xWs As Worksheet
Dim fso As Object, j As Long, folder1 As Object
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose the folder"
    .Show
End With
On Error Resume Next
xPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"
Application.Workbooks.Add
Set xWs = Application.ActiveSheet
xWs.Cells(1, 1).Value = xPath
xWs.Cells(2, 1).Resize(1, 5).Value = Array("Path", "Dir", "Name", "Date Created", "Date Last Modified")
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder1 = fso.getFolder(xPath)
getSubFolder folder1
xWs.Cells(2, 1).Resize(1, 5).Interior.Color = 65535
xWs.Cells(2, 1).Resize(1, 5).EntireColumn.AutoFit
Application.ScreenUpdating = True
End Sub

Sub getSubFolder(ByRef prntfld As Object)
Dim SubFolder As Object
Dim subfld As Object
Dim xRow As Long
For Each SubFolder In prntfld.SubFolders
    xRow = Range("A1").End(xlDown).Row + 1
    Cells(xRow, 1).Resize(1, 5).Value = Array(SubFolder.Path, Left(SubFolder.Path, InStrRev(SubFolder.Path, "\")), SubFolder.Name, SubFolder.DateCreated, SubFolder.DateLastModified)
Next SubFolder
For Each subfld In prntfld.SubFolders
    getSubFolder subfld
Next subfld
End Sub


Option Explicit

Sub GetFileNames()

Dim xRow As Long

Dim xDirect$, xFname$, InitialFoldr$

InitialFoldr$ = "C:\"

With Application.FileDialog(msoFileDialogFolderPicker)

.InitialFileName = Application.DefaultFilePath & "\"

.Title = "Please select a folder to list Files from"

.InitialFileName = InitialFoldr$

.Show

If .SelectedItems.Count <> 0 Then

xDirect$ = .SelectedItems(1) & "\"

xFname$ = Dir(xDirect$, 7)

Do While xFname$ <> ""

ActiveCell.Offset(xRow) = xFname$

xRow = xRow + 1

xFname$ = Dir

Loop

End If

End With

End Sub.





Sub MainList()
'Updateby20150706
Set folder = Application.FileDialog(msoFileDialogFolderPicker)
If folder.Show <> -1 Then Exit Sub
xDir = folder.SelectedItems(1)
Call ListFilesInFolder(xDir, True)
End Sub
Sub ListFilesInFolder(ByVal xFolderName As String, ByVal xIsSubfolders As Boolean)
Dim xFileSystemObject As Object
Dim xFolder As Object
Dim xSubFolder As Object
Dim xFile As Object
Dim rowIndex As Long
Set xFileSystemObject = CreateObject("Scripting.FileSystemObject")
Set xFolder = xFileSystemObject.GetFolder(xFolderName)
rowIndex = Application.ActiveSheet.Range("A65536").End(xlUp).Row + 1
For Each xFile In xFolder.Files
  Application.ActiveSheet.Cells(rowIndex, 1).Formula = xFile.Name
  rowIndex = rowIndex + 1
Next xFile
If xIsSubfolders Then
  For Each xSubFolder In xFolder.SubFolders
    ListFilesInFolder xSubFolder.Path, True
  Next xSubFolder
End If
Set xFile = Nothing
Set xFolder = Nothing
Set xFileSystemObject = Nothing
End Sub
Function GetFileOwner(ByVal xPath As String, ByVal xName As String)
Dim xFolder As Object
Dim xFolderItem As Object
Dim xShell As Object
xName = StrConv(xName, vbUnicode)
xPath = StrConv(xPath, vbUnicode)
Set xShell = CreateObject("Shell.Application")
Set xFolder = xShell.Namespace(StrConv(xPath, vbFromUnicode))
If Not xFolder Is Nothing Then
  Set xFolderItem = xFolder.ParseName(StrConv(xName, vbFromUnicode))
End If
If Not xFolderItem Is Nothing Then
  GetFileOwner = xFolder.GetDetailsOf(xFolderItem, 8)
Else
  GetFileOwner = ""
End If
Set xShell = Nothing
Set xFolder = Nothing
Set xFolderItem = Nothing
End Function