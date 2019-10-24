Option Explicit
'Developed by d.nguyen @ HocExcel.Online
'All right reserved.

Sub RemoveObject()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String

'Open dialog box to select a file
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Select file to remove objects"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "Please select a file!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

sourceFullName = xlChangeType(sourceFullName)
'Get folder path, file type and file name from the sourceFullName
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Use the date and time to create a unique file name
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copy and rename original file to a zip file with a unique name
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "Unable to copy " & sourceFullName & vbNewLine _
        & "Check the file is closed and try again"
    Exit Sub
End If
On Error GoTo 0

'Create folder to unzip to
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extract the files into the newly created folder
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'loop through each file in the \xl\worksheets folder of the unzipped file
xmlSheetFile = Dir(zipFilePath & "\xl\worksheets\*.xml*")
Do While xmlSheetFile <> ""

    'Read text of the file to a variable
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Input As xmlFile
    xmlFileContent = Input(LOF(xmlFile), xmlFile)
    Close xmlFile

'Manipulate the text in the file to remove the workbook protection
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<drawing")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
        xmlFileContent, "/>") + 2 ''"/>" is 2 characters long
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

    'Output the text of the variable to the file
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Output As xmlFile
    Print #xmlFile, xmlFileContent
    Close xmlFile

    'Loop to next xmlFile in directory
    xmlSheetFile = Dir

Loop


'Create empty Zip File
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Move files into the zip file
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Keep script waiting until Compressing is done
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Delete the files & folders created during the sub
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Rename the final file back to an xlsx file
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Show message box
MsgBox "All objects in the workbook have been removed.", _
vbInformation + vbOKOnly, Title:="Clear Objects"
TurnOnOff True
Exit Sub

xError:
MsgBox "Error!", vbExclamation, Title:="Clear Objects"
TurnOnOff True
End Sub

Sub RemoveAllName()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String
 
'Open dialog box to select a file
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Select file to remove all name"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "Please select a file!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

Dim regex As Object, str As String
Set regex = CreateObject("VBScript.RegExp")

sourceFullName = xlChangeType(sourceFullName)
'Get folder path, file type and file name from the sourceFullName
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Use the date and time to create a unique file name
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copy and rename original file to a zip file with a unique name
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "Unable to copy " & sourceFullName & vbNewLine _
        & "Check the file is closed and try again"
    Exit Sub
End If
On Error GoTo 0

'Create folder to unzip to
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extract the files into the newly created folder
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'Read text of the xl\workbook.xml file to a variable
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" For Input As xmlFile
xmlFileContent = Input(LOF(xmlFile), xmlFile)
Close xmlFile

With regex
  .Pattern = "(<definedName\s.*>#REF!<\/definedName>)"
  .Global = True 'If False, would replace only first
End With
xmlFileContent = regex.Replace(xmlFileContent, "")

''Manipulate the text in the file to remove the workbook protection
'xmlStartProtectionCode = 1
'xmlEndProtectionCode = 1
'Do
'xmlStartProtectionCode = InStr(xmlEndProtectionCode, xmlFileContent, "<definedNames")
'If xmlStartProtectionCode > 0 Then
'
'    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
'        xmlFileContent, "/definedNames>") ''"/>" is 2 characters long
'    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode + 14, _
'        xmlEndProtectionCode - xmlStartProtectionCode - 1)
'    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")
'
'End If
'Loop While xmlStartProtectionCode > 0

'Manipulate the text in the file to remove the workbook protection
'xmlStartProtectionCode = 0
'xmlStartProtectionCode = InStr(1, xmlFileContent, "<definedNames")
'If xmlStartProtectionCode > 0 Then
'
'    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
'        xmlFileContent, "/definedNames>") + 2 ''"/>" is 2 characters long
'    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
'        xmlEndProtectionCode - xmlStartProtectionCode)
'    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")
'
'End If

'Output the text of the variable to the file
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" & xmlSheetFile For Output As xmlFile
Print #xmlFile, xmlFileContent
Close xmlFile

'Create empty Zip File
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Move files into the zip file
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Keep script waiting until Compressing is done
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Delete the files & folders created during the sub
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Rename the final file back to an xlsx file
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Show message box
MsgBox "The all define names have been removed.", _
vbInformation + vbOKOnly, Title:="Password protection"
TurnOnOff True
Exit Sub

xError:
MsgBox "Error!", vbExclamation, Title:="Clear Objects"
TurnOnOff True
End Sub

Sub RemoveProtection()

Dim dialogBox As Object
Dim sourceFullName As String
Dim sourceFilePath As String
Dim sourceFileName As String
Dim sourceFileType As String
Dim newFileName As Variant
Dim tempFileName As String
Dim zipFilePath As Variant
Dim oApp As Object
Dim fso As Object
Dim xmlSheetFile As String
Dim xmlFile As Integer
Dim xmlFileContent As String
Dim xmlStartProtectionCode As Double
Dim xmlEndProtectionCode As Double
Dim xmlProtectionString As String

'Open dialog box to select a file
Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
dialogBox.AllowMultiSelect = False
dialogBox.Title = "Select file to remove protection from"

If dialogBox.Show = -1 Then
    sourceFullName = dialogBox.SelectedItems(1)
Else
    MsgBox "Please select a file!"
    Exit Sub
End If

TurnOnOff
On Error GoTo xError

sourceFullName = xlChangeType(sourceFullName)
'Get folder path, file type and file name from the sourceFullName
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)

'Use the date and time to create a unique file name
tempFileName = "Temp" & Format(Now, " dd-mmm-yy h-mm-ss")

'Copy and rename original file to a zip file with a unique name
newFileName = sourceFilePath & tempFileName & ".zip"
On Error Resume Next
FileCopy sourceFullName, newFileName

If Err.Number <> 0 Then
    MsgBox "Unable to copy " & sourceFullName & vbNewLine _
        & "Check the file is closed and try again"
    Exit Sub
End If
On Error GoTo 0

'Create folder to unzip to
zipFilePath = sourceFilePath & tempFileName & "\"
MkDir zipFilePath

'Extract the files into the newly created folder
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items

'loop through each file in the \xl\worksheets folder of the unzipped file
xmlSheetFile = Dir(zipFilePath & "\xl\worksheets\*.xml*")
Do While xmlSheetFile <> ""

    'Read text of the file to a variable
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Input As xmlFile
    xmlFileContent = Input(LOF(xmlFile), xmlFile)
    Close xmlFile

    'Manipulate the text in the file
    xmlStartProtectionCode = 0
    xmlStartProtectionCode = InStr(1, xmlFileContent, "<sheetProtection")

    If xmlStartProtectionCode > 0 Then

        xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
            xmlFileContent, "/>") + 2 '"/>" is 2 characters long
        xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
            xmlEndProtectionCode - xmlStartProtectionCode)
        xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

    End If

    'Output the text of the variable to the file
    xmlFile = FreeFile
    Open zipFilePath & "xl\worksheets\" & xmlSheetFile For Output As xmlFile
    Print #xmlFile, xmlFileContent
    Close xmlFile

    'Loop to next xmlFile in directory
    xmlSheetFile = Dir

Loop

'Read text of the xl\workbook.xml file to a variable
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" For Input As xmlFile
xmlFileContent = Input(LOF(xmlFile), xmlFile)
Close xmlFile

'Manipulate the text in the file to remove the workbook protection
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<workbookProtection")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, _
        xmlFileContent, "/>") + 2 ''"/>" is 2 characters long
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If



'Manipulate the text in the file to remove the modify password
xmlStartProtectionCode = 0
xmlStartProtectionCode = InStr(1, xmlFileContent, "<fileSharing")
If xmlStartProtectionCode > 0 Then

    xmlEndProtectionCode = InStr(xmlStartProtectionCode, xmlFileContent, _
        "/>") + 2 ''"/>" is 2 characters long
    xmlProtectionString = Mid(xmlFileContent, xmlStartProtectionCode, _
        xmlEndProtectionCode - xmlStartProtectionCode)
    xmlFileContent = Replace(xmlFileContent, xmlProtectionString, "")

End If

'Output the text of the variable to the file
xmlFile = FreeFile
Open zipFilePath & "xl\workbook.xml" & xmlSheetFile For Output As xmlFile
Print #xmlFile, xmlFileContent
Close xmlFile

'Create empty Zip File
Open sourceFilePath & tempFileName & ".zip" For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Move files into the zip file
oApp.Namespace(sourceFilePath & tempFileName & ".zip").CopyHere _
oApp.Namespace(zipFilePath).items
'Keep script waiting until Compressing is done
On Error Resume Next
Do Until oApp.Namespace(sourceFilePath & tempFileName & ".zip").items.Count = _
    oApp.Namespace(zipFilePath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

'Delete the files & folders created during the sub
Set fso = CreateObject("scripting.filesystemobject")
fso.deletefolder sourceFilePath & tempFileName

'Rename the final file back to an xlsx file
Name sourceFilePath & tempFileName & ".zip" As sourceFilePath & sourceFileName & Format(Now, " ddmmmyyhmmss") & "." & sourceFileType

'Show message box
MsgBox "The workbook and worksheet protection passwords have been removed.", _
vbInformation + vbOKOnly, Title:="Password protection"
TurnOnOff True
Exit Sub

xError:
MsgBox "Error!", vbExclamation, Title:="Clear Objects"
TurnOnOff True
End Sub

Function xlChangeType(ByVal sourceFullName As String) As String

Dim sourceFilePath As String, sourceFileType As String, sourceFileName As String
Dim strNewPathFix As String
'Get folder path, file type and file name from the sourceFullName
sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)
If Not sourceFileType = "xlsb" Then
    xlChangeType = sourceFullName
    Exit Function
End If
    Dim strCurrentFileExt   As String
    Dim strNewFileExt       As String
    Dim xlFile              As Workbook
    Dim strNewName          As String
    Dim strFolderPath       As String
    Dim strTimeNow As String
    strCurrentFileExt = ".xlsb"
    strNewFileExt = ".xlsm"

    strFolderPath = sourceFilePath
    If Right(strFolderPath, 1) <> "\" Then
        strFolderPath = strFolderPath & "\"
    End If
            strTimeNow = Format(Now, "ddmmmyyhmmss")
            Application.DisplayAlerts = False
            Set xlFile = Workbooks.Open(sourceFullName, , True)
            strNewName = Replace(sourceFileName, strCurrentFileExt, strNewFileExt)
            Select Case strNewFileExt
            Case ".xlsx"
                xlFile.SaveAs strFolderPath & strNewName & strTimeNow, XlFileFormat.xlOpenXMLWorkbook
            Case ".xlsm"
                xlFile.SaveAs strFolderPath & strNewName & strTimeNow, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
            End Select
            xlFile.Close
            Application.DisplayAlerts = True

ClearMemory:
    strNewPathFix = strFolderPath & strNewName & strTimeNow & strNewFileExt
    strCurrentFileExt = vbNullString
    strNewFileExt = vbNullString
    Set xlFile = Nothing
    strNewName = vbNullString
    strFolderPath = vbNullString
    xlChangeType = strNewPathFix
End Function
