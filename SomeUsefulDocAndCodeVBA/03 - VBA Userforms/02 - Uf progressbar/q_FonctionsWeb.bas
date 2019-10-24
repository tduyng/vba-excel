Attribute VB_Name = "q_FonctionsWeb"
Option Explicit
Public Sub LoginAndGetToken(Login As String, password As String, nomChantier As String)
    Dim wsTest As Worksheet
    Dim token As String
    Dim timeExpire As Date
    
    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets("SheetToken")
    On Error GoTo 0
    
    If wsTest Is Nothing Then
        Worksheets.Add.Name = "SheetToken"
    End If
    
    ActiveWorkbook.Worksheets("SheetToken").Visible = xlSheetVeryHidden

    'GenerateAuthentificationToken([login], [mot de passe], [nom du chantier Exacte])
    token = GenerateAuthentificationToken(Login, password, nomChantier)
    Sheets("SheetToken").Range("A1").value = token
    
End Sub

Public Function GetToken() As String
'Get the token dans le sheets : SheetToken
    GetToken = Sheets("SheetToken").Range("A1").value
End Function
Private Function GetBaseUrl() As String
    GetBaseUrl = "https://wizzyou.wizzcad.com:8091/Locataire.svc/web/"
'    GetBaseUrl = "http://dev.wizzyou.wizzcad.com:8081/Locataire.svc/web/"
End Function
Private Function GetResponseRequest(Url As String) As String
'    Dim Req As Object
    Dim Req As New XMLHTTP60
'    Set Req = CreateObject("MSXML2.serverXMLHTTP")
    Req.Open "POST", Url, False
    Req.send
    
    GetResponseRequest = Replace(Req.responseText, """", "")

End Function
Private Function ReplaceSpecialCharactere(value As String)

    value = Replace(value, "%", "%25")
    value = Replace(value, "$", "%24")
    value = Replace(value, "&", "%26")
    value = Replace(value, "+", "%2B")
    value = Replace(value, ",", "%2C")
    value = Replace(value, "/", "%2F")
    value = Replace(value, ":", "%3A")
    value = Replace(value, ";", "%3B")
    value = Replace(value, "=", "%3D")
    value = Replace(value, "?", "%3F")
    value = Replace(value, "@", "%40")
    value = Replace(value, " ", "%20")
    value = Replace(value, """", "%22")
    value = Replace(value, "<", "%3C")
    value = Replace(value, ">", "%3E")
    value = Replace(value, "#", "%23")
    value = Replace(value, "{", "%7B")
    value = Replace(value, "}", "%7D")
    value = Replace(value, "|", "%7C")
    value = Replace(value, "\", "%5C")
    value = Replace(value, "^", "%5E")
    value = Replace(value, "~", "%7E")
    value = Replace(value, "[", "%5B")
    value = Replace(value, "]", "%5D")
    value = Replace(value, "`", "%60")
    value = Replace(value, "è", "%C3%A8")
    value = Replace(value, "é", "%C3%A9")
    'caractère chantier
    value = Replace(value, "à", "%C3%A0")
    
    ReplaceSpecialCharactere = value
End Function
Private Function FormaterTel(tel As String) As String

    tel = Replace(tel, " ", "")
    tel = Replace(tel, ".", "")
    tel = Replace(tel, ",", "")
    tel = Right(tel, 9)
    
    FormaterTel = tel

End Function
Public Function GenerateAuthentificationToken(Login As String, password As String, chantier As String) As String

    Dim Url As String
    
    Login = ReplaceSpecialCharactere(Login)
    password = ReplaceSpecialCharactere(password)
    chantier = ReplaceSpecialCharactere(chantier)
    
    Url = GetBaseUrl() & "GetTokenAuthentification?login=" & Login & "&pwd=" & password & "&nomChantier=" & chantier
    GenerateAuthentificationToken = Replace(GetResponseRequest(Url), "\/", "/")

End Function
Private Function GetInfoLocataire(ByVal token As String, Appartement As String, nomFunction As String)
    Dim Url As String
    
    token = ReplaceSpecialCharactere(token)
    If Not Appartement <> vbNullString Then
        Exit Function
    Else
        Appartement = ReplaceSpecialCharactere(Appartement)
        
        Url = GetBaseUrl() & nomFunction & "?token=" & token & "&appartement=" & Appartement
        GetInfoLocataire = GetResponseRequest(Url)
    End If
End Function

Public Function GetBatimentLocataire(ByVal token As String, Appartement As String)
    GetBatimentLocataire = GetInfoLocataire(token, Appartement, "GetBatiment")

End Function

Public Function GetHallLocataire(ByVal token As String, Appartement As String)
    GetHallLocataire = GetInfoLocataire(token, Appartement, "GetHall")

End Function

Public Function GetAdresseLocataire(ByVal token As String, Appartement As String)
    GetAdresseLocataire = GetInfoLocataire(token, Appartement, "GetAdresse")

End Function

Public Function GetEtageLocataire(ByVal token As String, Appartement As String)
    
    GetEtageLocataire = GetInfoLocataire(token, Appartement, "GetEtage")
End Function

Public Function GetTypeLocataire(ByVal token As String, Appartement As String)
    GetTypeLocataire = GetInfoLocataire(token, Appartement, "GetType")
End Function

Public Function GetAppartementLocataire(ByVal token As String, Appartement As String)
    GetAppartementLocataire = GetInfoLocataire(token, Appartement, "GetAppartement")
End Function

Public Function GetNomLocataire(ByVal token As String, Appartement As String)
    GetNomLocataire = GetInfoLocataire(token, Appartement, "GetNom")
End Function

Public Function GetPrenomLocataire(ByVal token As String, Appartement As String)
    GetPrenomLocataire = GetInfoLocataire(token, Appartement, "GetPrenom")
End Function

Public Function GetEmailLocataire(ByVal token As String, Appartement As String)
    GetEmailLocataire = GetInfoLocataire(token, Appartement, "GetEmail")
End Function

Public Function GetTelephoneLocataire(ByVal token As String, Appartement As String)
    GetTelephoneLocataire = GetInfoLocataire(token, Appartement, "GetTelephoneFixe")
End Function

Public Function GetTelephoneMobileLocataire(token As String, Appartement As String)
    Dim mobile As String
    
    mobile = GetInfoLocataire(ByVal token, Appartement, "GetTelephoneMobile")
    If InStr(1, mobile, "+33") = 1 Then
        mobile = "0" & Right(mobile, 9)
    End If
    GetTelephoneMobileLocataire = mobile
End Function

Public Function GetLoginLocataire(ByVal token As String, Appartement As String)
    GetLoginLocataire = GetInfoLocataire(token, Appartement, "GetLogin")
End Function

Public Function GetDerniereDateConnexionLocataire(ByVal token As String, Appartement As String)
    Dim dateCo As String
    dateCo = GetInfoLocataire(token, Appartement, "GetDerniereDateConnexion")
    GetDerniereDateConnexionLocataire = Replace(dateCo, "\/", "/")
End Function

Private Function ArrayLen(arr As Variant) As Long
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Sub GetCodeImportAndFillTable(ByVal token As String, startColumn As String)
'----------------------------------------------------
'Récupérer tous des valeurs et remplir dans le tableau
'----------------------------------------------------
    Dim Url As String
    Dim listCode As String
    Dim code As Variant
    Dim Codes() As String
    Dim i As Long
    Dim CodeEtNomMetier() As String
    Dim nomMetier As String
    Dim sh As Worksheet
    
    Set sh = Sheets("0 - WIZZCAD")
    
    token = ReplaceSpecialCharactere(token)
    
    Url = GetBaseUrl() & "GetCodeImport?token=" & token
    listCode = GetResponseRequest(Url)
    Codes = Split(listCode, ";")
    
    i = 0
    With sh
        For Each code In Codes
            CodeEtNomMetier = Split(code, "=")
            If (ArrayLen(CodeEtNomMetier) = 2) Then
                .Range(startColumn).Offset(0, i).value = CodeEtNomMetier(0)
                .Range(startColumn).Offset(1, i).value = Replace(CodeEtNomMetier(1), "\/", "/")
            Else
                .Range(startColumn).value = CodeEtNomMetier(0)
            End If
            
            i = i + 1
        Next code
    End With
End Sub

Public Sub GetRendezVousParMetierId(ByVal token As String, Appartement As String, startRange As Range)
    Dim Url As String
    Dim DateRendezVous As String
    Dim NbMetier As Long
    Dim metierId As String
    Dim i As Long
    
    Select Case True
    Case Appartement <> ""
        token = ReplaceSpecialCharactere(token)
        metierId = Cells(1, startRange.Column).value
        NbMetier = Range(Cells(1, startRange.Column), Cells(1, startRange.Column)).End(xlToRight).Column - startRange.Column
        
        For i = 0 To NbMetier
            metierId = Cells(1, startRange.Column + i).value
            
            Url = GetBaseUrl() & "GetRendezVous?token=" & token & "&appartement=" & Appartement & "&metierId=" & metierId
            DateRendezVous = GetResponseRequest(Url)
            DateRendezVous = Replace(DateRendezVous, "\/", "/")
            
            Cells(startRange.Row, startRange.Column + i).value = DateRendezVous
        Next i
    Case Appartement = ""
        Exit Sub
    End Select
End Sub

Public Sub InsertOrUpdateLocataireRendezVous(ByVal token As String, Appartement As String, Batiment As String, _
                                            Hall As String, Adresse As String, Etage As String, TypeLocataire As String, _
                                            Nom As String, Prenom As String, Email As String, TelFixe As String, TelMobile As String, _
                                            startRange As Range, importInfoLocataire As Long)

    '--------------------------------------------------------------------------
    'Cette fonction permet d'envoyer tous les informations de locataires vers le web Wizzcad
    'importInfoLocataire = 0: Envoyer que les rendez-vous
    'importInfoLocataire = 1: import la liste de locataire et leur rdv
    '--------------------------------------------------------------------------
    Dim Url As String
    Dim RendezVousCodeDate As String
    Dim DateRendezVous As String
    Dim NbMetier As Long
    Dim metierId As String
    Dim i As Long
    Dim dataConcat As String
    Dim response As String
    Dim responseParts() As String
    Dim Part As Variant
    
    If Appartement <> "" Then

        token = ReplaceSpecialCharactere(token)
        NbMetier = Range(Cells(1, startRange.Column), Cells(1, startRange.Column)).End(xlToRight).Column - startRange.Column
        
        DateRendezVous = ""

        Select Case True
        Case NbMetier > 0
            For i = 0 To NbMetier
                If Cells(startRange.Row, startRange.Column + i).value <> "" Then
                    metierId = Cells(1, startRange.Column + i).value
                    DateRendezVous = Cells(startRange.Row, startRange.Column + i).value
                    If RendezVousCodeDate <> "" Then
                        RendezVousCodeDate = RendezVousCodeDate & ";" & metierId & "=" & DateRendezVous
                    Else
                        RendezVousCodeDate = metierId & "=" & DateRendezVous
                    End If
                End If
            Next i
        Case NbMetier = 0
            RendezVousCodeDate = ""
        End Select

        
        TelFixe = FormaterTel(TelFixe)
        TelMobile = FormaterTel(TelMobile)
        
        dataConcat = Appartement & "!" & Batiment & "!" & Hall & "!" & Adresse & "!" & Etage & "!" & TypeLocataire & "!" & Nom & "!" & Prenom & "!" & Email & "!" & TelFixe & "!" & TelMobile & "!" & RendezVousCodeDate
        dataConcat = ReplaceSpecialCharactere(dataConcat)
        
        Url = GetBaseUrl() & "InsertOrUpdateLocataireRendezVous?token=" & token & "&dataConcat=" & dataConcat & "&importLocataireInfo=" & importInfoLocataire
        response = GetResponseRequest(Url)
        response = Replace(response, "\/", "/")
        
        responseParts = Split(response, ";")
        response = ""
        For Each Part In responseParts
            If Part <> "" Then
                If response = "" Then
                    response = Part
                Else
                    response = response & Chr(13) & Part
                End If
            End If
        Next Part
        
'        MsgBox response

    Else
        MsgBox "La valeur de numéro de logment est vide"
    End If

End Sub

Public Function VerificationToken(ByVal token As String) As String

    Dim Url As String
    
    token = ReplaceSpecialCharactere(token)
    
    Url = GetBaseUrl() & "GetValidityToken?token=" & token
    VerificationToken = GetResponseRequest(Url)

End Function

