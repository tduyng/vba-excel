Attribute VB_Name = "m04_Export"
Option Explicit


Sub EXPORT_PDF_MultipleSheet_ForOne(arrNamesheet As Variant, nameFilePDF As String)
    '----------------------------------------------------------------------
    'This fucntion is used for exporting all selected sheets in one file PDF
    '----------------------------------------------------------------------
    Dim sFilePDF As String
    Dim wb As Workbook

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    If Len(Join(arrNamesheet)) <= 0 Then Exit Sub
    'Export PDF
    'expression .ExportAsFixedFormat(Type, Filename, Quality, _
    'IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
    On Error Resume Next
    Set wb = ThisWorkbook
    sFilePDF = GetNextAvailableName(nameFilePDF)

    wb.Sheets(arrNamesheet).Select
    wb.Sheets(arrNamesheet(0)).Activate
    ActiveSheet.ExportAsFixedFormat _
                        Type:=xlTypePDF, _
                        FileName:=sFilePDF, _
                        Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, _
                        IgnorePrintAreas:=False, _
                        OpenAfterPublish:=ufExportImport.ckxOpenAfterExport.value
                    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .EnableAnimations = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    wb.Sheets("02-MODEL").Select
End Sub
Sub test_EXPORTPDFOneByOne()
    Dim arrSheet As Variant
    Dim pathFolderPDF As String
    arrSheet = Array("FAC.01", "FAC.02", "FAC.03")
    pathFolderPDF = "E:\00 - BIM\07 - EXCEL\01 - PROJETS EXCEL VINCI\02 - BEAB Facturation\Export PDF"
    Call EXPORTPDFOneByOne(arrSheet, pathFolderPDF)
    
End Sub



Sub EXPORTPDFOneByOne(arrNamesheet As Variant, pathFolderPDF As String)
    '----------------------------------------------------------------------
    'This fucntion is used for exporting all selected sheets in one file PDF
    '----------------------------------------------------------------------
    Dim sFilePDF As String
    Dim wb As Workbook
    Dim itemSheet As Variant

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    If Len(Join(arrNamesheet)) <= 0 Then Exit Sub

    On Error Resume Next
    Set wb = ThisWorkbook
    pathFolderPDF = pathFolderPDF & "\"
    For Each itemSheet In arrNamesheet
         sFilePDF = GetNextAvailableName(pathFolderPDF & itemSheet & ".pdf")

        wb.Sheets(itemSheet).Select
        wb.Sheets(itemSheet).Activate
        ActiveSheet.ExportAsFixedFormat _
                            Type:=xlTypePDF, _
                            FileName:=sFilePDF, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=ufExportImport.ckxOpenAfterExport.value
            
    Next itemSheet
    
   
                    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .EnableAnimations = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    wb.Sheets("02-MODEL").Select
End Sub
Sub Export_Excel_MultipleSheet_ForOne(arrNamesheet As Variant, fullNameFile As String)
    Dim ws As Worksheet
    Dim OriginalWB As Workbook, DestWB As Workbook
    Dim selectedNameSheet As String
    
    On Error Resume Next
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    Set OriginalWB = ThisWorkbook
    Set DestWB = Application.Workbooks.Open(fullNameFile)
    
    If Len(Join(arrNamesheet)) <= 0 Then Exit Sub
    For Each ws In OriginalWB.Sheets(arrNamesheet)
            ws.Copy After:=DestWB.Sheets(DestWB.Sheets.count)
'            DestWB.Sheets(DestWB.Sheets.count).Name = NextNameSheetAvailable(DestWB, ws.Name)
    Next ws

    If DestWB.Sheets(1).Name = "Feuil1" Then DestWB.Sheets(1).Delete

    DestWB.Close SaveChanges:=True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .EnableAnimations = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub


Sub test_ExportExcelOneByOne()
    Dim arrSheet As Variant
    Dim pathFolderExcel As String
    arrSheet = Array("FAC.01", "FAC.02", "FAC.03")
    pathFolderExcel = "E:\00 - BIM\07 - EXCEL\01 - PROJETS EXCEL VINCI\02 - BEAB Facturation\Export Excel"
    Call EXPORTExcelOneByOne(arrSheet, pathFolderExcel)
End Sub
Sub EXPORTExcelOneByOne(arrNamesheet As Variant, pathFolderExcel As String)
    Dim ws As Worksheet
    Dim OriginalWB As Workbook, DestWB As Workbook
    Dim selectedNameSheet As String
    Dim fullNameFileExcel As String
    
    On Error Resume Next
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    pathFolderExcel = pathFolderExcel & "\"
    Set OriginalWB = ThisWorkbook
    
    
    If Len(Join(arrNamesheet)) <= 0 Then Exit Sub
    For Each ws In OriginalWB.Sheets(arrNamesheet)
            Set DestWB = Workbooks.Add
            ws.Copy After:=DestWB.Sheets(DestWB.Sheets.count)
'            DestWB.Sheets(DestWB.Sheets.count).Name = NextNameSheetAvailable(DestWB, ws.Name)
                If DestWB.Sheets(1).Name = "Feuil1" Then DestWB.Sheets(1).Delete
             fullNameFileExcel = pathFolderExcel & ws.Name & ".xlsx"
             fullNameFileExcel = GetNextAvailableName(fullNameFileExcel)
            DestWB.SaveAs FileName:=fullNameFileExcel
            DestWB.Close
    Next ws


'    DestWB.Close SaveChanges:=True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .EnableAnimations = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub
Function GetNextAvailableName(ByVal strPath As String) As String

    With CreateObject("Scripting.FileSystemObject")

        Dim strFolder As String, strBaseName As String, strExt As String, i As Long
        strFolder = .GetParentFolderName(strPath)
        strBaseName = .getbasename(strPath)
        strExt = .GetExtensionName(strPath)

        Do While .FileExists(strPath)
            i = i + 1
            strPath = .BuildPath(strFolder, strBaseName & " (" & i & ")" & "." & strExt)
        Loop

    End With

    GetNextAvailableName = strPath

End Function
Sub test_GetListProjectChecked()
    Dim arr As Variant
    On Error Resume Next
    arr = GetListProjectChecked
End Sub
Function GetListProjectChecked() As Variant
    Dim shProject As Worksheet
    Dim startRowProject&, lastRowProject&
    Dim colCheck&, count&, i&
    Dim listProjectChecked()
    
    Set shProject = ThisWorkbook.Sheets("00-PROJETS")
    With shProject
        startRowProject = .Range("nameProject").Row + 1
        lastRowProject = getLastRowCol(4, shProject, shProject.Range("nameProject").Column)
        colCheck = .Range("colCheck").Column
        ReDim listProjectChecked(1 To lastRowProject - startRowProject + 1)
        For i = startRowProject To lastRowProject
        
            If .Cells(i, colCheck).value = True Then
                count = count + 1
                listProjectChecked(count) = .Cells(i, .Range("nameProject").Column).value

            End If
        Next i
    End With
    If count = 0 Then Exit Function
    ReDim Preserve listProjectChecked(1 To count)
    GetListProjectChecked = Application.Transpose(listProjectChecked)
End Function

Function NextNameSheetAvailable(wb As Workbook, nameSheetInput As String) As String
    Dim dic As Object
    Dim i&
    Dim ws As Worksheet
    Set dic = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Sheets
        If Not dic.exists(ws.Name) Then
            dic.Add ws.Name, vbNullString
        End If
    Next ws
        Do While dic.exists(nameSheetInput)
            i = i + 1
            nameSheetInput = nameSheetInput & " (" & i & ")"
        Loop
    NextNameSheetAvailable = nameSheetInput

End Function
Function GetSelectedSheet() As Variant
    Dim selectedSheets As Variant
    Dim index&
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    With ufExportImport
        For index = 0 To .lbxSheet.ListCount - 1
                If .lbxSheet.Selected(index) And Not dic.exists(.lbxSheet.List(index)) Then
                    dic.Add .lbxSheet.List(index), vbNullString
                End If
        Next index
    End With
    GetSelectedSheet = dic.keys
End Function
Private Function CreateSaveFileName(ByVal FileName As String) As String
  Dim n As Long
  Dim sFolder As String, sFile As String, sExt As String, sTmpFile As String
  sTmpFile = FileName
  With CreateObject("Scripting.FileSystemObject")
    sExt = .GetExtensionName(FileName)
    sFolder = .GetParentFolderName(FileName)
    sFile = .getbasename(FileName)
    Do While .FileExists(sTmpFile) = True
      n = n + 1
      sTmpFile = .BuildPath(sFolder, sFile & "(" & n & ")." & sExt)
    Loop
    CreateSaveFileName = sTmpFile
  End With
End Function




Sub CREATE_FOLDER()
    
'   Get name or directory of file Excel active
'   iOption = 0: getbase name of this file and add date today for exporting
'   iOption = 1 for getting the Directory
    Dim wb As Workbook
    Dim sFolderPDF As String, sFolderPDFChild As String
    Dim sFilePDF As String
    Dim sFolderExcel As String, sFolderExcelChild As String
    Dim baseNameFile As String
    Dim fso As Object

    Set wb = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If wb.Path = "" Then
        Exit Sub
    End If
    'Dect if folder exits
    
    sFolderPDF = wb.Path & "\" & "Export PDF"
    sFolderPDFChild = sFolderPDF & "\" & "PDF_" & Format(Date, "yyyymmdd")
    sFolderExcel = wb.Path & "\" & "Export Excel"
    sFolderExcelChild = sFolderExcel & "\" & "Excel_" & Format(Date, "yyyymmdd")
    
    On Error Resume Next
    If Not fso.Folderexists(sFolderPDF) Then
       fso.createFolder (sFolderPDF)
    End If
    
    If Not fso.Folderexists(sFolderPDFChild) And ufExportImport.ckxExSeparate.value Then
       fso.createFolder (sFolderPDFChild)
    End If
    If Not fso.Folderexists(sFolderExcel) Then
       fso.createFolder (sFolderExcel)
    End If
    If Not fso.Folderexists(sFolderExcelChild) And ufExportImport.ckxExSeparate.value Then
       fso.createFolder (sFolderExcelChild)
    End If
    
    Call Shell("explorer.exe " & wb.Path, vbNormalFocus)
    On Error GoTo 0
End Sub

Public Function getExportName(wb As Workbook)
    Dim fso As Object
    Dim filePath As String
    Dim nameExport As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = wb.Path
    If filePath = "" Then Exit Function
    
    nameExport = fso.getbasename(wb.FullName) & "_" & Format(Date, "yyyymmdd")
    getExportName = nameExport
    
End Function


Public Function GET_PATH_EXPORT(wb As Workbook) As String

    Dim fso As Object
    Dim wbPath As String
    Dim filePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    wbPath = wb.Path
    If wbPath = "" Then Exit Function
    
    With ufExportImport
        Select Case .cbxSaveasType.ListIndex
        Case 1
            If Not .ckxExSeparate.value Then
                    If fso.Folderexists(wbPath & "\" & "Export PDF") Then
                        filePath = wbPath & "\" & "Export PDF"
                    Else
                        filePath = wbPath
                    End If
               
            Else
            'Name PDF avec format name_PDF_DateActuelle
                If fso.Folderexists(wbPath & "\" & "Export PDF" & "\PDF_" & Format(Date, "yyyymmdd")) Then
                    filePath = wbPath & "\" & "Export PDF" & "\PDF_" & Format(Date, "yyyymmdd")
                ElseIf fso.Folderexists(wbPath & "\" & "Export PDF") Then
                    filePath = wbPath & "\" & "Export PDF"
                Else
                    filePath = wbPath
                End If
                
            End If
        Case 0
            If Not .ckxExSeparate.value Then
                    If fso.Folderexists(wbPath & "\" & "Export Excel") Then
                        filePath = wbPath & "\" & "Export Excel"
                    Else
                        filePath = wbPath
                    End If
            Else
                If fso.Folderexists(wbPath & "\" & "Export Excel" & "\Excel_" & Format(Date, "yyyymmdd")) Then
                    filePath = wbPath & "\" & "Export Excel" & "\Excel_" & Format(Date, "yyyymmdd")
                ElseIf fso.Folderexists(wbPath & "\" & "Export Excel") Then
                    filePath = wbPath & "\" & "Export Excel"
                Else
                    filePath = wbPath
                End If
            End If
        Case Else
            Exit Function
        End Select
    End With
    GET_PATH_EXPORT = filePath

End Function
Public Function getListSheet(wb As Workbook, Optional IncludeShHidden As Long = 1) As Variant
    Dim iSh As Worksheet
    Dim arr()
    Dim i&
    Set wb = ThisWorkbook
    ReDim arr(1 To wb.Sheets.count)
    If IncludeShHidden = 1 Then
        i = 0
        For Each iSh In wb.Sheets
            If iSh.Visible = -1 Or iSh.Visible = 0 Then
                i = i + 1
                arr(i) = iSh.Name
            End If
        Next
        ReDim Preserve arr(1 To i)
    Else
        i = 0
        For Each iSh In wb.Sheets
            If iSh.Visible = -1 Then
                i = i + 1
                arr(i) = iSh.Name
            End If
        Next
        ReDim Preserve arr(1 To i)
    End If
    
    getListSheet = arr
End Function
Public Function CountSheet(wb As Workbook, Optional IncludeShHidden As Long = 1) As Variant
    Dim iSh As Worksheet
    Dim i&
    Set wb = ThisWorkbook
    If IncludeShHidden = 1 Then
        CountSheet = wb.Sheets.count
    Else
        i = 0
        For Each iSh In wb.Sheets
            If iSh.Visible <> -1 Then
                i = i + 1
            End If
        Next
        CountSheet = wb.Sheets.count - i
    End If
End Function

Sub testgetpath()
    GET_PATH_EXPORT (ThisWorkbook)
End Sub

