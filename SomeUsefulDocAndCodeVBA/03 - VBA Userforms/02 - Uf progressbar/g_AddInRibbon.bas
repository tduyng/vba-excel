Attribute VB_Name = "g_AddInRibbon"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpfile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public bckMEFPlanning As Boolean
Public bckColorCiblePlan As Boolean
Public bckMEFInfoGenComp As Boolean
Public bckDialRecap As Boolean

'Lien entre checkbox of Addin et valeur dans la feuille Paramètrage
Sub CkxValParam_MEFPlanning(Optional bVal As Boolean = True)
     With ThisWorkbook.Sheets("1 - PARAMETRES")
        If bVal = True Then
            .Range("Refresh_Plan").value = "oui"
        Else
            .Range("Refresh_Plan").value = "non"
        End If
    End With
End Sub

Sub CkxValParam_ColorCiblePlan(Optional bVal As Boolean = True)
    With ThisWorkbook.Sheets("1 - PARAMETRES")
        If bVal = True Then
            .Range("Refresh_Color_Plan").value = "oui"
        Else
            .Range("Refresh_Color_Plan").value = "non"
        End If
    End With
End Sub

Sub CkxValParam_MEFInfoGenComp(Optional bVal As Boolean = True)
    With ThisWorkbook.Sheets("1 - PARAMETRES")
        If bVal = True Then
            .Range("Refresh_Infos").value = "oui"
            .Range("Refresh_Compl").value = "oui"
        Else
            .Range("Refresh_Infos").value = "non"
            .Range("Refresh_Compl").value = "non"
        End If
    End With
End Sub


Sub CkxValParam_DialRecap(Optional bVal As Boolean = True)
    With ThisWorkbook.Sheets("1 - PARAMETRES")
        If bVal = True Then
            .Range("Actualise_Recap").value = "oui"
        Else
            .Range("Actualise_Recap").value = "non"
        End If
    End With
End Sub

'*******************************************************************
'===================================================================
'IRibbonControl of AddIN GTM Bâtiment
'===================================================================

'Quand Le ruban est chargé, on lui demande de faire les actions suivantes
Sub RubanCharge(ribbon As IRibbonUI)
    bckMEFPlanning = False
    bckColorCiblePlan = False
    bckMEFInfoGenComp = False
    bckDialRecap = False
    With ThisWorkbook.Sheets("1 - PARAMETRES")
        If .Range("Refresh_Plan").value = "oui" Then bckMEFPlanning = True
        If .Range("Refresh_Color_Plan").value = "oui" Then bckColorCiblePlan = True
        If .Range("Refresh_Infos").value = "oui" Then
            bckMEFInfoGenComp = True
            .Range("Refresh_Compl").value = "oui"
        End If
        If .Range("Actualise_Recap").value = "oui" Then bckDialRecap = True
    End With
End Sub


'********************************************************
'Group WIZZCAD
Sub btAiLogin(Control As IRibbonControl)
    Call SHOW_ufConnection
End Sub

Sub btAiImport(Control As IRibbonControl)
    Call SHOW_ufImportWizzCad
End Sub
Sub btAiExport(Control As IRibbonControl)
    Call SHOW_ufExportWizzCad
End Sub




'********************************************************
'Group PLANNING
Sub btAiPhotoPlanning(Control As IRibbonControl)
    Call PHOTO_PLANNING
End Sub

Sub btAiMAJPointage(Control As IRibbonControl)
    Call MAJ_POINTAGE
End Sub

Sub btAiMAJMetier(Control As IRibbonControl)
    Call MAJ_METIER
End Sub
Sub btAiDataLocataire(Control As IRibbonControl)
    Call TRANSFER_DATA_FROM_INFOS_TO_PLANNING("Yes", "Tous")
End Sub


Sub btAiRefreshComptage(Control As IRibbonControl)
    Call SHOW_COMPTAGE_TRAVAUX
End Sub
Sub btAiAvisPassage(Control As IRibbonControl)
    Call AVIS_PASSAGE_MENU
End Sub
Sub btAiPlanningST(Control As IRibbonControl)
    Call CMD_PLANNING
End Sub
Sub btAixx(Control As IRibbonControl)

End Sub
Sub btAixx2(Control As IRibbonControl)

End Sub



'********************************************************
'Group SYNOPTIQUE
Sub btAiCreateSynoptique(Control As IRibbonControl)
    Call SHOW_CREATE_Synoptique
End Sub




'********************************************************
'Groupe EDITION

Sub btAiCopyLogementToWizzcad(Control As IRibbonControl)
    COPY_LOGEMENTS_TO_WIZZCAD
End Sub

Sub btAiDelAllDataForNewProject(Control As IRibbonControl)
    Call CREATE_NEW_PROJECT
End Sub
Sub btAiDelWizzcad(Control As IRibbonControl)
    Call VIDER_ShtWizzcad
End Sub
Sub btAiDelPlanning(Control As IRibbonControl)
    Call VIDER_ShtPlanning
End Sub
Sub btAiDelInfoGen(Control As IRibbonControl)
    Call VIDER_ShtInfoGen
End Sub
Sub btAiDelInfoComp(Control As IRibbonControl)
    Call VIDER_ShtInfoComp
End Sub
Sub btAiDelSynoptique(Control As IRibbonControl)
    Call VIDER_ShtSynop
End Sub
Sub btAiDeleteAvis(Control As IRibbonControl)
    Call EFFACER_SELECT
End Sub




'********************************************************
'Groupe MISE EN FORME BUTTON
Sub btAiCaptureScreen(Control As IRibbonControl)
    Call prcSave_Picture_Active_Window
End Sub
Sub btAiCaptureScreen5s(Control As IRibbonControl)
    Call prcSave_Picture_Active_Window(, 5000)
End Sub



'********************************************************
'Groupe MISE EN FORME BUTTON
Sub btAiMEFPlanning(Control As IRibbonControl)
    Call COULEUR_GRILLE
End Sub
Sub btAiMEFInfoGen(Control As IRibbonControl)
    Call COULEUR_GRILLE_2("Infos")
End Sub
Sub btAiMEFInfoComp(Control As IRibbonControl)
    Call COULEUR_GRILLE_2("Complément")
End Sub
Sub btAiMEFSynop(Control As IRibbonControl)
    Call SHOW_MEF_Synoptique
End Sub

Sub btAiRefreshColorPlan(Control As IRibbonControl)
    Call SHOW_Color_Planning
End Sub



'********************************************************
'Groupe CHECBOX FONCTION AUTOMATIQUE
Sub ckxAiMEFPlanning(Control As IRibbonControl, pressed As Boolean)
    CkxValParam_MEFPlanning (pressed)
End Sub
Sub ckxAiColorCiblePlan(Control As IRibbonControl, pressed As Boolean)
    CkxValParam_ColorCiblePlan (pressed)
End Sub

Sub ckxAiMEFInfoGenComp(Control As IRibbonControl, pressed As Boolean)
    CkxValParam_MEFInfoGenComp (pressed)
End Sub


Sub ckxAiDialRecap(Control As IRibbonControl, pressed As Boolean)
    CkxValParam_DialRecap (pressed)
End Sub

'group ckeckbox enabled
Sub ckxAiMEFPlanning_Pressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = bckMEFPlanning
    CkxValParam_MEFPlanning (returnedVal)
End Sub
Sub ckxAiColorCiblePlan_Pressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = bckColorCiblePlan
    CkxValParam_ColorCiblePlan (returnedVal)
End Sub

Sub ckxAiMEFInfoGenComp_Pressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = bckMEFInfoGenComp
    CkxValParam_MEFInfoGenComp (returnedVal)
End Sub


Sub ckxAiDialRecap_Pressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = bckDialRecap
    CkxValParam_DialRecap (returnedVal)
End Sub



'********************************************************
'Groupe MAIL
Sub btAiTestMail(Control As IRibbonControl)
    Call TEST_MAIL
End Sub
Sub btAiSendMail(Control As IRibbonControl)
    MsgBox "Cette fonctionnalité est bloquée pour l'instant.", vbOKOnly + vbInformation, "Add-In GTM Bâtiment"
End Sub



'********************************************************
'Groupe AIDE
Sub btAiWebWizzcad(Control As IRibbonControl)
    ShellExecute 0&, vbNullString, Control.Tag, vbNullString, vbNullString, 0
End Sub
Sub btAiAide(Control As IRibbonControl)
    ShellExecute 0&, vbNullString, Control.Tag, vbNullString, vbNullString, 0
End Sub

Sub btAiAPropos(Control As IRibbonControl)
    Load Accueil_UF
    Accueil_UF.Show
End Sub
Sub btAiReportProblem(Control As IRibbonControl)
    Call Open_Outlook
End Sub
