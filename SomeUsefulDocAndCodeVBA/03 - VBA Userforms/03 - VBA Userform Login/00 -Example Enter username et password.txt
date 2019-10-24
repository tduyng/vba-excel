Dim temp As String
On Error Resume Next 
temp = WorksheetFunction.VLookup(Me.txtUserNameIn.Value, Range("UserRegister"), 1, 0) 
 
If Username = temp Then 
     '****************
    Err.Clear 
    temp = "" 
    temp = WorksheetFunction.VLookup(Me.txtUserNameIn.Value, Range("UserRegister"), 2, 0) 
    On Error Goto 0 
    If password = temp Then 
        Sheets("ADMIN").Visible = xlSheetVisible 
        MsgBox "Password & Username Accepted" 
        Unload Me 
        Sheets("ADMIN").Select 
        Range("A1").Select 
    Else 
        Sheets("ADMIN").Visible = xlVeryHidden 

        MsgBox "Invalid Password, Workbook will close" 
        Unload Me 
         'ThisWorkbook.Close  (I'm not wanting to go-live with this bit yet, so commented out)
    End If 
Else
        Sheets("ADMIN").Visible = xlVeryHidden 
        MsgBox "Invalid UserName, Workbook will close"
        Unload Me
End If