Attribute VB_Name = "d_CallUserForm"
Option Explicit

Sub SHOW_ufConnection()
'==================================================================
'Afficher userform ufConnection
'==================================================================
    Load ufConnection
    ufConnection.StartUpPosition = 2
    ufConnection.Show False
End Sub

Sub SHOW_ufExportWizzCad()
'==================================================================
'Afficher userform ufExportWizzcad
'==================================================================
    Load ufExportWizzCad
    ufExportWizzCad.StartUpPosition = 2
    ufExportWizzCad.Show False
End Sub

Sub SHOW_ufImportWizzCad()
'==================================================================
'Afficher userform ufImportWizzcad
'==================================================================
    Load ufImportWizzCad
    ufImportWizzCad.StartUpPosition = 2
    ufImportWizzCad.Show False
End Sub



Sub SHOW_UFPROGRESSBAR()
'==================================================================
'Pour événément initialisé de userform ufPrgressBar
'Mettre en 0 tous les valeurs d'épasseur et la caption de la barre de progression
'==================================================================
    Load ufProgress
    With ufProgress
        .LabelCaption.Caption = "Traitement en cours... Veuillez patienter."
        .LabelProgress.Width = 0
        .StartUpPosition = 2
        .Show
    End With
End Sub


Sub UF_PROGRESS_BAR_DOEVENTS(iRun As Single)
'==================================================================
'Augmenter l'épaisseur de la barre de progression par rapport au pourcentage de irunning (irunning = 1 to 100)
'==================================================================
        With ufProgress
                .LabelProgress.Width = iRun * (.FrameProgress.Width)
        End With
'       The DoEvents statement is responsible for the form updating
        DoEvents
End Sub


Sub SHOW_COMPTAGE_TRAVAUX()

    On Error Resume Next
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = "Comptage_Travaux"
    On Error GoTo 0
    ThisWorkbook.Worksheets("2 - PLANNING").Activate
    Call SHOW_UFPROGRESSBAR
End Sub

Sub SHOW_CREATE_Synoptique()

    On Error Resume Next
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = "CREATESynoptique"
    On Error GoTo 0
    ThisWorkbook.Worksheets("9 - SYNOPTIQUE").Activate
    Call SHOW_UFPROGRESSBAR
End Sub


Sub SHOW_MEF_Synoptique()

    On Error Resume Next
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = "MEFSynoptique"
    On Error GoTo 0
    ThisWorkbook.Worksheets("9 - SYNOPTIQUE").Activate
    Call SHOW_UFPROGRESSBAR
End Sub

Sub SHOW_Color_Planning()

    On Error Resume Next
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = "Couleur_Planning"
    On Error GoTo 0
    ThisWorkbook.Worksheets("2 - PLANNING").Activate
    Call SHOW_UFPROGRESSBAR
End Sub
