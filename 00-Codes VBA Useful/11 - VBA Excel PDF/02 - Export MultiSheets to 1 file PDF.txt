
Sub EXPORT_PDF()
    Dim wb As Workbook, sh As Worksheet, i As Integer
    Dim sFolder As String, sFile As Variant, sFilePDF As String
    Dim arrSheets()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Export PDF
    'expression .ExportAsFixedFormat(Type, Filename, Quality, _
    'IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
    ReDim arrSheets(1 To Worksheets.Count)
    For Each shi In Worksheets
        i = i + 1
        sFile = ThisWorkbook.Path & "\" & shi.Name & ".pdf"
        sFilePDF = CStr(sFile)
        arrSheets(i) = shi.Name
    Next shi
    ReDim Preserve arrSheets(1 To i)
    Sheets(arrSheets).Select
    ActiveSheet.ExportAsFixedFormat _
                    xlTypePDF, _
                    "test", _
                    xlQualityStandard, _
                    True, _
                    False, , , False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function CreateSaveFileName(ByVal FileName As String) As String
  Dim n As Long
  Dim sFolder As String, sFile As String, sExt As String, sTmpFile As String
  sTmpFile = FileName
  With CreateObject("Scripting.FileSystemObject")
    sExt = .GetExtensionName(FileName)
    sFolder = .GetParentFolderName(FileName)
    sFile = .GetBaseName(FileName)
    Do While .FileExists(sTmpFile) = True
      n = n + 1
      sTmpFile = .BuildPath(sFolder, sFile & "(" & n & ")." & sExt)
    Loop
    CreateSaveFileName = sTmpFile
  End With
End Function
