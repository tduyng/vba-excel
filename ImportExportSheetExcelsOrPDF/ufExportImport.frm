VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufExportImport 
   Caption         =   "Export -Import"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "ufExportImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufExportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filePathOpen As String



Private Sub Label2_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub tbxNameFile_Change()

End Sub

'----------------------------------------------------
'Event initialize of userform
'----------------------------------------------------
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Set wb = ThisWorkbook
   
    With Me
        .lbxSheet.List = getListSheet(wb, 1)
'        .lbPath.Caption = wb.FullName
        .ckxShowOnly.value = False
        .ckxShowSheetHidden.value = True
        .lbCountSheet.Caption = "Nombre éléments: " & .lbxSheet.ListCount
        .ckxAll.value = False
        .ckxUnCheckAll.value = False
        
        If Not .ckxExSeparate.value Then
            .tbxNameFile.value = getExportName(wb)
        End If
        
        .cbxSaveasType.List = Array("Excel(*.xl*)", "PDF(*.pdf)")
        .cbxSaveasType.ListIndex = 1
    End With
    filePathOpen = ""
End Sub



'--------------------------------------------
'Event change control of userform
'--------------------------------------------


Private Sub lbxSheet_AfterUpdate()
    Me.lbCountSheet.Caption = "Nombre éléments: " & Me.lbxSheet.ListCount
End Sub

Private Sub ckxShowOnly_Change()
    Dim dic As Object, DicHidden As Object
    Dim wb As Workbook
    Dim id As Variant
    Dim iKey As Variant
    Dim i&, countSh&
    Dim iSh As Worksheet
    
    Set wb = ThisWorkbook
    countSh = CountSheet(wb, 0)
    Set dic = CreateObject("Scripting.Dictionary")
    Set DicHidden = CreateObject("Scripting.Dictionary")
    
    
    For i = 0 To Me.lbxSheet.ListCount - 1
                If Me.lbxSheet.Selected(i) And Not dic.exists(Me.lbxSheet.List(i)) Then
                    dic.Add Me.lbxSheet.List(i), i
                End If
    Next i
    For i = 0 To Me.lbxSheet.ListCount - 1
            Set iSh = wb.Sheets(Me.lbxSheet.List(i))
                If iSh.Visible = -1 And Me.lbxSheet.Selected(i) And Not DicHidden.exists(Me.lbxSheet.List(i)) Then
                    DicHidden.Add Me.lbxSheet.List(i), i
                End If
            Set iSh = Nothing
    Next i
    
    If dic Is Nothing Then Exit Sub
    With Me
        .lbxSheet.Clear
        .ckxAll.value = False
        .ckxUnCheckAll.value = False
        'neu show only = value thi nghia la chi can show gia tri cua
        If .ckxShowOnly.value Then
            If .ckxShowSheetHidden.value Then
                .lbxSheet.Clear
                .lbxSheet.List = dic.keys
                For i = 0 To .lbxSheet.ListCount - 1
                    If dic.exists(.lbxSheet.List(i)) Then
                        .lbxSheet.Selected(i) = True
                    End If
                Next i
            Else
                .lbxSheet.Clear
                .lbxSheet.List = DicHidden.keys
                For i = 0 To .lbxSheet.ListCount - 1
                    If DicHidden.exists(.lbxSheet.List(i)) Then
                        .lbxSheet.Selected(i) = True
                    End If
                Next i
            End If
            
        'Neu khong show only thi sao: se say ra 2 truong hop
        'TH1: neu checkbox showhidden = true thi nghia la hien thi tat cac cac sheet hidden
        'Khi do ta se get all sheet trong workbook (ngoai tru sheeteveryhidden luon de an)
        'TH2: neu check box showhidden = false thi chi lay cac sheet dan hien thi roi gan vao listbox
        Else
            If .ckxShowSheetHidden.value Then
                .lbxSheet.Clear
                .lbxSheet.List = getListSheet(wb, 1)
                For i = 0 To .lbxSheet.ListCount - 1
                    If dic.exists(.lbxSheet.List(i)) Then
                    .lbxSheet.Selected(i) = True
                    End If
                Next i
            Else
                .lbxSheet.Clear
                 .lbxSheet.List = getListSheet(wb, 0)
                 For i = 0 To .lbxSheet.ListCount - 1
                    If dic.exists(.lbxSheet.List(i)) Then
                    .lbxSheet.Selected(i) = True
                    End If
                Next i
            End If
        End If
    End With
    Set dic = Nothing
    Set DicHidden = Nothing
End Sub



Private Sub ckxShowSheetHidden_Change()
    Dim dic As Object, DicHidden As Object
    Dim wb As Workbook
    Dim id As Variant, count&
    Dim iKey As Variant
    Dim i&
    Dim iSh As Worksheet

    Set wb = ThisWorkbook
    count = 0
    Set dic = CreateObject("Scripting.Dictionary")
    Set DicHidden = CreateObject("Scripting.Dictionary")
    For i = 0 To Me.lbxSheet.ListCount - 1
                If Me.lbxSheet.Selected(i) And Not dic.exists(Me.lbxSheet.List(i)) Then
                    dic.Add Me.lbxSheet.List(i), i
                End If
    Next i
    
    For i = 0 To Me.lbxSheet.ListCount - 1
            Set iSh = wb.Sheets(Me.lbxSheet.List(i))
                If iSh.Visible = -1 And Me.lbxSheet.Selected(i) And Not DicHidden.exists(Me.lbxSheet.List(i)) Then
                    DicHidden.Add Me.lbxSheet.List(i), i
                End If
            Set iSh = Nothing
    Next i
    
    If dic Is Nothing Then Exit Sub
    
    With Me
        'Tam thoi thay gia tri checkall khong th co lien quan den  checlbox con lai
        .lbxSheet.Clear
        .ckxAll.value = False
        .ckxUnCheckAll.value = False
        'Ta set truong hop neu ckecbox Showonly  =false thi se recuperer truc tiep valu tu sheet
        'Neu truong hop
        If .ckxShowOnly.value = False Then
            If .ckxShowSheetHidden.value Then
                .lbxSheet.List = getListSheet(wb, 1)
                
                 For i = 0 To .lbxSheet.ListCount - 1
                    If dic.exists(.lbxSheet.List(i)) Then
                    .lbxSheet.Selected(i) = True
                    End If
                Next i
    
            Else
                .lbxSheet.List = getListSheet(wb, 0)
                For i = 0 To .lbxSheet.ListCount - 1
                    If dic.exists(.lbxSheet.List(i)) Then
                    .lbxSheet.Selected(i) = True
                    End If
                Next i
            End If
        Else
            'Ta se t truong hop checkbox showOnly = true: nghia la nos dang trong trang thai chi hien thi sheet dang duoc chon
            If .ckxShowSheetHidden.value Then
                .lbxSheet.List = dic.keys
                 For i = 0 To .lbxSheet.ListCount - 1
                        If dic.exists(.lbxSheet.List(i)) Then
                        .lbxSheet.Selected(i) = True
                        End If
                Next i
            Else
                .lbxSheet.List = DicHidden.keys
                 For i = 0 To .lbxSheet.ListCount - 1
                        If DicHidden.exists(.lbxSheet.List(i)) Then
                        .lbxSheet.Selected(i) = True
                        End If
                Next i
            End If
            
            
        End If
    
     End With
End Sub

Private Sub cbxSaveasType_Change()
    If filePathOpen = "" Then
        Me.tbxDirectory.value = GET_PATH_EXPORT(ThisWorkbook)
    Else
        Me.tbxDirectory.value = filePathOpen
    End If
    With Me
        If Not .ckxExSeparate.value Then
            Select Case .cbxSaveasType.ListIndex
            Case 0
            'Excel
                        .lbNameFile.Visible = False
                        .tbxNameFile.value = ""
                    If Not .ckxExSeparate.value Then
                        .lbCreateNewFolder.Visible = False
                        .cbCreateFolder.Visible = False
                        .lbSaveDirectory.Visible = False
                        .tbxDirectory.Visible = False
                        .cbOpenFile.Visible = False
                        .tbxNameFile.Visible = True
                        .cbSelectFileToExport.Visible = True
                    Else
                        .lbCreateNewFolder.Visible = True
                        .cbCreateFolder.Visible = True
                        .lbSaveDirectory.Visible = True
                        .tbxDirectory.Visible = True
                        .cbOpenFile.Visible = True
                        .cbSelectFileToExport.Visible = False
                        .tbxNameFile.Visible = False
                    End If
                    
            Case 1
            'PDF
                    .lbCreateNewFolder.Visible = True
                    .cbCreateFolder.Visible = True
                    .lbSaveDirectory.Visible = True
                    .tbxDirectory.Visible = True
                    .cbOpenFile.Visible = True
                    .cbSelectFileToExport.Visible = False
                    If Not .ckxExSeparate.value Then
                        .tbxNameFile.Visible = True
                        .tbxNameFile.value = getExportName(ThisWorkbook)
                        .lbNameFile.Visible = True
                    Else
                        .tbxNameFile.Visible = False
                        .lbNameFile.Visible = False
                    End If

            End Select
        Else
        
        End If
    End With
End Sub

Private Sub ckxExSeparate_Change()
   With Me
        If filePathOpen = "" Then
            Me.tbxDirectory.value = GET_PATH_EXPORT(ThisWorkbook)
        Else
            Me.tbxDirectory.value = filePathOpen
        End If
    Select Case .cbxSaveasType.ListIndex
        Case 1
            'PDF
            If .cbxSaveasType.ListIndex = 1 And Not .ckxExSeparate.value Then
                .tbxNameFile.Visible = True
                .tbxNameFile.value = getExportName(ThisWorkbook)
                .lbNameFile.Visible = True
            ElseIf .cbxSaveasType.ListIndex = 1 And .ckxExSeparate.value Then
                .tbxNameFile.Visible = False
                .lbNameFile.Visible = False
            End If
         Case 0
            
              .lbNameFile.Visible = False
                        .tbxNameFile.value = ""
                    If Not .ckxExSeparate.value Then
                        .lbCreateNewFolder.Visible = False
                        .cbCreateFolder.Visible = False
                        .lbSaveDirectory.Visible = False
                        .tbxDirectory.Visible = False
                        .cbOpenFile.Visible = False
                        .tbxNameFile.Visible = True
                        .cbSelectFileToExport.Visible = True
                    Else
                        .lbCreateNewFolder.Visible = True
                        .cbCreateFolder.Visible = True
                        .lbSaveDirectory.Visible = True
                        .tbxDirectory.Visible = True
                        .cbOpenFile.Visible = True
                        .cbSelectFileToExport.Visible = False
                        .tbxNameFile.Visible = False
                End If
    End Select
    End With
End Sub





'----------------------------------------------------
'Event click control of userform
'----------------------------------------------------

Private Sub cbCancelExport_Click()
    Unload Me
End Sub

Private Sub cbOkExport_Click()
    With Me
    If Len(Join(GetSelectedSheet)) <= 0 Then
        MsgBox "Il faut sélectionner au moins une feuille à exporter", vbInformation, "Export"
        Exit Sub
    End If
    If Not .ckxExSeparate.value Then
        
        
        Select Case .cbxSaveasType.ListIndex
        Case 1
            If .tbxDirectory.value = "" Then
                MsgBox "Vous devez sélectionner un dossier à exporter!", vbInformation, "Manque dossier de sauvegarde"
                Exit Sub
            Else
                Call EXPORT_PDF_MultipleSheet_ForOne(GetSelectedSheet, .tbxDirectory & "\" & .tbxNameFile.value & ".pdf")
            End If

        Case 0
            If .tbxNameFile.value = "" Then
                MsgBox "Vous devez sélectionner un fichier Excel à exporter", vbInformation, "Export Excel"
                Exit Sub
            Else
                Call Export_Excel_MultipleSheet_ForOne(GetSelectedSheet, .tbxNameFile.value)
            End If
        End Select
    Else
        Select Case .cbxSaveasType.ListIndex
        Case 1
            If .tbxDirectory.value = "" Then
                MsgBox "Vous devez sélectionner un dossier à exporter!", vbInformation, "Manque dossier de sauvegarde"
                Exit Sub
            Else
                Call EXPORTPDFOneByOne(GetSelectedSheet, .tbxDirectory.value)
            End If

        Case 0
            If .tbxDirectory.value = "" Then
                 MsgBox "Vous devez sélectionner un dossier à exporter!", vbInformation, "Manque dossier de sauvegarde"
                Exit Sub
            Else
                Call EXPORTExcelOneByOne(GetSelectedSheet, .tbxDirectory.value)
            End If
        End Select
    End If
    End With
    
    Me.Hide
End Sub
Private Sub cbCreateFolder_Click()
    Call CREATE_FOLDER
    If filePathOpen = "" Then
        Me.tbxDirectory.value = GET_PATH_EXPORT(ThisWorkbook)
    Else
        Me.tbxDirectory.value = filePathOpen
    End If
   
End Sub

Private Sub cbOpenFile_Click()
    Dim fileDial As FileDialog, xFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Sélectionner une destination"
    .Show
    On Error Resume Next
    xFolder = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End With
    If xFolder = "" Then Exit Sub
    filePathOpen = xFolder
    Me.tbxDirectory.value = filePathOpen
            
End Sub
Private Sub cbSelectFileToExport_Click()
    Dim fileDial As FileDialog, xFolder As String
    Dim fso As Object
    Dim selectedFileExtension As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    With Application.FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .InitialFileName = Application.ActiveWorkbook.Path & "\"
            .Filters.Clear
            .Filters.Add "Files Excel", "*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xls"
            .Filters.Add "All Files", "*.*"
            .Show
        On Error Resume Next
        xFolder = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End With
    If xFolder = "" Then Exit Sub
    selectedFileExtension = fso.GetExtensionName(xFolder)
    If InStr(selectedFileExtension, "xl") = 0 Then
        MsgBox "Le fichier sélectionné n'est pas a format Excel.", vbCritical, "BEAB"
        Exit Sub
    End If
    Me.tbxNameFile.value = xFolder
End Sub
Private Sub ckxAll_Click()
    Dim i&
    With Me
        If .ckxAll.value = True Then
            For i = 0 To .lbxSheet.ListCount - 1
                .lbxSheet.Selected(i) = True
            Next i
    
        .ckxUnCheckAll.value = False
        End If
    End With

End Sub

Private Sub ckxUnCheckAll_Click()
    Dim i&
    With Me
    
    If .ckxUnCheckAll.value = True Then
        For i = 0 To .lbxSheet.ListCount - 1
            .lbxSheet.Selected(i) = False
        Next i

        .ckxAll.value = False
    End If
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
