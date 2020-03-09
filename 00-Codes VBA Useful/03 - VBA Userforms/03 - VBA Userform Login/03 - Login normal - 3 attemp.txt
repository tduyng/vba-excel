Option Explicit

'https://www.onlinepclearning.com/excel-userform-login/


 Private Trial As Long
 Private Sub cmdCheck_Click()
 'Declare the variables
 Dim AddData As Range, Current As Range
 Dim user As Variant, Code As Variant
 Dim PName As Variant, AName As Variant
 Dim ws As Worksheet, ws2 As Worksheet, ws3 As Worksheet
 Dim result As Integer
 Dim TitleStr As String
 Dim msg As VbMsgBoxResult

'Variables
 user = Me.txtUser.Value
 Code = Me.txtPass.Value
 TitleStr = "Password check"
 result = 0
 Set Current = Sheet2.Range("O8")

'Error handler
 On Error GoTo errHandler:
 'Destination location for login storage
 Set AddData = Sheet2.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0)
 'Check the login and passcode for the administrator
 If user <> "" And Not IsNumeric(user) And Code <> "" And IsNumeric(Code) Then
 For Each AName In Sheet2.Range("T8:T108")
 'If AName = Code Then 'Use this for passcode text
 If AName = CLng(Code) And AName.Offset(0, -1) = user Then ' Use this for passcode numbers only
 MsgBox "Welcome Back: – " & user & " " & Code
 'record user login
 AddData.Value = user
 AddData.Offset(0, 1).Value = Now
 'Add usernmae to the worksheet
 Current.Value = user
 'Change variable if the condition is meet
 result = 1
 'Unload the form
 Sheet2.visible = True
 Sheet2.Select
 Unload Me
 'Show the navigation form
 'frmNavigation.Show
 Exit Sub
 End If
 Next AName
 End If

'Check user login with loop
 If user <> "" And Not IsNumeric(user) And Code <> "" And IsNumeric(Code) Then
 For Each PName In Sheet2.Range("H8:H108")
 'If PName = Code Then 'Use this for passcode text
 If PName = CLng(Code) And PName.Offset(0, -1) = user Then ' Use this for passcode numbers only
 MsgBox "Welcome Back: – " & user & " " & Code
 'record user login
 AddData.Value = user
 AddData.Offset(0, 1).Value = Now
 'Add usernmae to the worksheet
 Current.Value = user

'unhide worksheet for user
 If PName.Offset(0, 1) <> "" Then
 Set ws = Worksheets(PName.Offset(0, 1).Value)
 ws.visible = True
 End If

'unhide worksheet for user
 If PName.Offset(0, 2) <> "" Then
 Set ws2 = Worksheets(PName.Offset(0, 2).Value)
 ws2.visible = True
 End If

'unhide worksheet for user
 If PName.Offset(0, 3) <> "" Then
 Set ws3 = Worksheets(PName.Offset(0, 3).Value)
 ws3.visible = True
 End If

'show sheet tab if hidden
 ActiveWindow.DisplayWorkbookTabs = True

'Change variable if the condition is meet
 result = 1
 'Unload the form
 Unload Me
 'Show the navigation form
 'frmNavigation.Show
 Exit Sub
 End If
 Next PName
 End If

'Check to see if an error occurred
 If result = 0 Then
 'Increment error variable
 Trial = Trial + 1
 'Less then 3 error message
 If Trial < 3 Then msg = MsgBox("Wrong password, please try again", vbExclamation + vbOKOnly, TitleStr)
 Me.txtUser.SetFocus
 'Last chance and close the workbook
 If Trial = 3 Then
 msg = MsgBox("Wrong password, the form will close…", vbCritical + vbOKOnly, TitleStr)
 ActiveWorkbook.Close False
 End If
 End If
 Exit Sub
 'Error block
 errHandler:
 MsgBox "An Error has Occurred " & vbCrLf & "The error number is: " _
 & Err.Number & vbCrLf & Err.Description & vbCrLf & _
 "Please notify the administrator"
 End Sub
