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


Sub test_create8pdf
	Dim filepdf  as string
	filepdf = CreateSaveFileName(ThisWorkbook.Path & "\" & Range("G10") & ".pdf")
	ActiveSheet.ExportAsFixedFormat xlTypePDF, filepdf
end sub