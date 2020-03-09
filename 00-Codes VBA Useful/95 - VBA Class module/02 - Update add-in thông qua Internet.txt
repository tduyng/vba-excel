'-------------------------------------------------------------------------
' Class Module    : clsUpdate
' Company         : JKP Application Development Services (c)
' Author          : Jan Karel Pieterse
' Created         : 19-2-2007
' Purpose         : Class to check for program updates
'-------------------------------------------------------------------------
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private mdtLastUpdate As Date

Private msAppName As String
Private msBuild As String
Private msCheckURL As String
Private msCurrentadd-inName As String
Private msDownloadName As String
Private msTempadd-inName As String

Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
    ' Download the file.
    URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0
End Sub

Public Function IsThereAnUpdate() As Boolean
    Dim oIE As InternetExplorer
    Set oIE = New InternetExplorer
    With oIE
        .Navigate2 CheckURL
        Do
        Loop Until .Busy = False
        If Len(.Document.body.innerhtml) > 0 Then
            If CLng(.Document.body.innerhtml) > CLng(Build) Then
                IsThereAnUpdate = True
            End If
        End If
        .Quit
    End With
    Set oIE = Nothing
End Function

Public Property Get Build() As String
    Build = msBuild
End Property

Public Property Let Build(ByVal sBuild As String)
    msBuild = sBuild
End Property

Public Sub RemoveOldCopy()
    Currentadd-inName = ThisWorkbook.FullName
    Tempadd-inName = Currentadd-inName & "(OldVersion)"
    On Error Resume Next
    Kill Tempadd-inName
End Sub

Public Function GetUpdate() As Boolean
    On Error Resume Next
    ThisWorkbook.SaveAs Tempadd-inName
    DoEvents
    Kill Currentadd-inName
    On Error GoTo 0
    DownloadFile DownloadName, Currentadd-inName
    LastUpdate = Now
    If Err = 0 Then GetUpdate = True
End Function

Private Property Get Currentadd-inName() As String
    Currentadd-inName = msCurrentadd-inName
End Property

Private Property Let Currentadd-inName(ByVal sCurrentadd-inName As String)
    msCurrentadd-inName = sCurrentadd-inName
End Property

Private Property Get Tempadd-inName() As String
    Tempadd-inName = msTempadd-inName
End Property

Private Property Let Tempadd-inName(ByVal sTempadd-inName As String)
    msTempadd-inName = sTempadd-inName
End Property

Public Property Get DownloadName() As String
    DownloadName = msDownloadName
End Property

Public Property Let DownloadName(ByVal sDownloadName As String)
    msDownloadName = sDownloadName
End Property

Public Property Get CheckURL() As String
    CheckURL = msCheckURL
End Property

Public Property Let CheckURL(ByVal sCheckURL As String)
    msCheckURL = sCheckURL
End Property

Public Property Get LastUpdate() As Date
    Dim dtNow As Date
    dtNow = Int(Now)
    mdtLastUpdate = CDate(GetSetting(AppName, "Updates", "LastUpdate", CStr(dtNow)))
    If mdtLastUpdate = dtNow Then
        'Never checked for an update, save today!
        SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(dtNow))
    End If
    LastUpdate = mdtLastUpdate
End Property

Public Property Let LastUpdate(ByVal dtLastUpdate As Date)
    mdtLastUpdate = dtLastUpdate
    SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(mdtLastUpdate))
End Property

Public Property Get AppName() As String
    AppName = msAppName
End Property

Public Property Let AppName(ByVal sAppName As String)
    msAppName = sAppName
End Property




'Và b?n thêm vào do?n code sau dây trong module d? th?c hi?n vi?c c?p nh?t


Option Explicit

Public Sub CheckAndUpdate()
    Dim cUpdate As clsUpdate
    Set cUpdate = New clsUpdate
    With cUpdate
        'Set intial values of class
        'Current build
        .Build = "0"
        'Name of this app, probably a global variable, such as GSAPPNAME
        .AppName = "CheckForUpdate"
        'Get rid of old backup copy
        .RemoveOldCopy
        'URL which contains build # of new version
        .CheckURL = "http://www.jkp-ads.com/downloads/UpdateAnAddinBuild.htm"
        'Check once a week
        If Now - .LastUpdate >= 7 Or Int(Now) = .LastUpdate Then
            If .IsThereAnUpdate Then
                If MsgBox("We have an update, do you wish to download?", _
                            vbQuestion + vbYesNo) = vbYes Then
                    .DownloadName = "http://www.jkp-ads.com/downloads/" & _
                                     ThisWorkbook.Name
                    If .GetUpdate Then
                        MsgBox "Successfully updated the add-in, " & vbNewLine & _
                               "please restart Excel to start using the new version!", _
                                vbOKOnly + vbInformation
                    Else
                        MsgBox "Updating has failed.", vbInformation + vbOKOnly
                    End If
                End If
            End If
        End If
    End With
    Set cUpdate = Nothing
End Sub