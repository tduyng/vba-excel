 Option Explicit

'Source ref: https://www.onlinepclearning.com/excel-userform-login-form/

Private Trial As Long
Private Sub cmdCheck_Click()
 'Declare the variables
 Dim AddData As Range
 Dim user As Variant
 Dim Code As Variant
 Dim result As Integer
 Dim TitleStr As String
 Dim Current As Range
 Dim PName As Variant
 Dim msg As VbMsgBoxResult
 'Variables
 user = Me.txtUser.Value
 Code = Me.txtPass.Value
 TitleStr = "Password check"
 result = 0
 Set Current = Sheet2.Range("K2")
 'Error handler
 On Error GoTo errHandler:
 'Destination location for login storage
 Set AddData = Sheet2.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0)
 'Check the login and passcode for the administrator
 If user = "Trevor" And Code = 9999 Then
 MsgBox "Welcome Back: –   " & user & "   " & Code & vbCrLf _
 & " You have admistrator priviledges" & vbCrLf _
 & " I will open the control panel for you"
 'record user login
 AddData.Value = user
 AddData.Offset(0, 1).Value = Now
 'send the username to the worksheet
 Current.Value = user
 'unoad this form
 Unload Me
 'Show navigation form
 frmSetup.Show
 'End the procedure if conditions are meet
 Exit Sub
 End If
 'Check user login with loop
 If user <> "" And Code <> "" Then
 For Each PName In Sheet2.Range("H2:H100")
 'If PName = Code Then 'Use this for passcode text
 If PName = CInt(Code) And PName.Offset(0, -1) = user Then ' Use this for passcode numbers only
 MsgBox "Welcome Back: –   " & user & "   " & Code
 'record user login
 AddData.Value = user
 AddData.Offset(0, 1).Value = Now
 'Change variable if the condition is meet
 result = 1
 'Add usernmae to the worksheet
 Current.Value = user
 'Unload the form
 Unload Me
 'Show the navigation form
 frmNavigation.Show
 Exit Sub
 End If
 Next PName
 End If
 ' Next UName
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
 MsgBox "An Error has Occurred  " & vbCrLf & "The error number is:  " _
 & Err.Number & vbCrLf & Err.Description & vbCrLf & _
 "Please notify the administrator"
 End Sub

This code will stop  the x button on the Userform from working.

Private Sub UserForm_QueryClose _
 (Cancel As Integer, CloseMode As Integer)
 '   Prevents use of the Close button
 If CloseMode = vbFormControlMenu Then
 MsgBox "Clicking the Close button does not work."
 Cancel = True
 End If
 End Sub
