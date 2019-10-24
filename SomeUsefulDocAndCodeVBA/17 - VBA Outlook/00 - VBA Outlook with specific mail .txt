Option Explicit


Sub Mail_workbook_Outlook_1()

    Dim outApp As Object
    Dim outMail As Object

    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)

    On Error Resume Next
    With outMail
        .To = "tienduy.nguyen2@estp.fr"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .Body = "Hi there"
        .Attachments.Add ActiveWorkbook.FullName
        .Send
    End With
    On Error GoTo 0

    Set outMail = Nothing
    Set outApp = Nothing
End Sub
Sub Which_Account_Number()
'Don't forget to set a reference to Outlook in the VBA editor
    Dim outApp As Object
    Set outApp = CreateObject("Outlook.Application")
    Dim i As Long

    For i = 1 To outApp.session.Accounts.Count
        MsgBox outApp.session.Accounts.Item(i) & " : This is account number " & i
    Next i
End Sub

Sub MAIL_WITH_SPECIFIC_ACCOUNT()

    Dim outApp As Object
    Dim outMail As Object
    Dim outAccount As Object
    Dim strbody As String, StrFrom As String, i As Integer, itemAccount As Integer, signature As String
    Dim foundAccount As Boolean
    
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)

    StrFrom = "tienduy.nguyen2@estp.fr"
    strbody = "Ligne 1 avec 2 retour à la ligne derriere mais en fait ca fait une interligne, à la limite, pas de pb <br>" _
          & "<br>" _
          & "<br>" _
          & "Ligne 2 : OK" _
          & "Ligne 3 : OK<br>" _
          & "Ligne 4 : OK<br>" _
          & "Ligne 5 : OK<br>" _
          & "Ligne 6 avec 2 retour à la ligne derriere mais en fait ca fait une interligne, à la limite, pas de pb <br>" _
          & "<br>" _
          & "<br>" _
          & "Ligne 7 : OK" & "<br>" _
          & "Ligne 8 avec UN SEUL retour à la ligne derriere mais en fait ca fait une interligne. Je pense que c'est dû au fait qu'il y a beaucoup de choses sur cette ligne mais comment y palier ? <br>" _
          & "<br>" _
          & "Ligne 9 : OK" & "<br>" _
          & "Ligne 10 : OK" & "<br>" _
          & "<br>"

    strbody = "genre" & " " & "nom" & ",<BR><BR>" & _
              "comme indiqué par téléphone, je vous confirme votre réservation au restaurant " & _
              "<B>" & "restaurant" & "</B>"
               
    foundAccount = False
'    'Give index item of mail specific
    For i = 1 To outApp.session.Accounts.Count
        If outApp.session.Accounts.Item(i).smtpAddress = StrFrom Then
            itemAccount = i
            foundAccount = True
            Exit For
        End If
    Next i
    
    If foundAccount Then
        On Error Resume Next
        With outMail
            .Display
            .To = "tien-duy.nguyen@vinci-construction.fr"
            .CC = ""
            .BCC = ""
            .Subject = "This is the Subject line"
            .HTMLBody = strbody & "<br>" & .HTMLBody

    '        Set .SendOnBehaftOfName = "tienduy9x@outlook.com" --> no working
    '       'Set mail sender with the mail specific
            Set .SendUsingAccount = outApp.session.Accounts.Item(itemAccount)
'            .Send   'or use .Display
            
    '        'Delete after submit
    '        .DeleteAfterSubmit = True 'Set to True if a copy of the mail message is not saved upon being sent
            
        End With
        On Error GoTo 0
    Else
        MsgBox "No account with this mail founded" & vbNewLine & "You need to verify the mail of sender"
    End If
    
    Set outMail = Nothing
    Set outApp = Nothing
    
End Sub

