VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Exécution en cours"
   ClientHeight    =   1050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "ufProgress.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim subProgressBar As String


Private Sub Label1_Click()

End Sub

Private Sub cbAnnuler_Click()
    Application.SendKeys ("{BREAK}")
End Sub



Private Sub UserForm_Activate()

    
    On Error Resume Next
    subProgressBar = ThisWorkbook.Worksheets("SheetToken").Range("A10").value
    On Error GoTo 0
    
    Select Case subProgressBar
        Case Is = "import_wz0"
            Call IMPORT_WIZZCAD(0)
                
        Case Is = "import_wz1"
            Call IMPORT_WIZZCAD(1)
                
        Case Is = "export_wz0"
            Call EXPORT_WIZZCAD(0)
                
        Case Is = "export_wz1"
            Call EXPORT_WIZZCAD(1)
    
        Case Is = "Comptage_Travaux"
            Call COMPTAGE_TRAVAUX
            
        Case Is = "MEFSynoptique"
            Call MEF_SYNOPTIQUE
            
        Case Is = "CREATESynoptique"
            Call CREATE_SYNOPTIQUE
            
        Case Is = "Couleur_Planning"
            Call REFRESH_COlOR_PLANNING
            
        Case Else
            Unload Me
    End Select
    
    Unload Me
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = ""
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()
On Error Resume Next
    ThisWorkbook.Worksheets("SheetToken").Range("A10").value = ""
On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
'    'Hide tile bar
'    #If IsMac = False Then
'    'hide the title bar if you're working on a windows machine. Otherwise, just display it as you normally would
'    Me.Height = Me.Height - 10
'    y_HideTitle.HideTitleBar Me
'    #End If
    
    
    On Error Resume Next
    subProgressBar = ThisWorkbook.Worksheets("SheetToken").Range("A10").value
    On Error GoTo 0
    
    With ufProgress
        Select Case subProgressBar
            Case Is = "import_wz0"
                .Caption = "Import WizzCAD sans RDV"
                    
            Case Is = "import_wz1"
                .Caption = "Import WizzCAD avec RDV"
                    
            Case Is = "export_wz0"
                .Caption = "Export WizzCAD sans RDV"
                    
            Case Is = "export_wz1"
                .Caption = "Export WizzCAD avec RDV"
        
            Case Is = "Comptage_Travaux"
                .Caption = "Rafraîchir comptage"
                
            Case Is = "CREATESynoptique"
                .Caption = "Créer Synoptique"
                
            Case Is = "MEFSynoptique"
                .Caption = "Mise en forme Synoptique"
                
            Case Is = "Couleur_Planning"
                .Caption = "Actualiser couleurs Planning"
                
            Case Else
                Unload Me
        End Select
    subRemoveCloseButton Me
        .LabelCaption.Caption = "Traitement en cours... Veuillez patienter."
        .LabelProgress.Width = 0
    End With

End Sub
