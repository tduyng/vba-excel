Sub TestFolder()
    Debug.Print XLFileIsOpen("C:\Test")
End Sub

Function XLFileIsOpen(sFolder As String) As Boolean
    For Each Item In EnumerateFiles(sFolder)
        If IsWorkBookOpen(CStr(Item)) = True Then XLFileIsOpen = True
    Next Item
End Function

Function EnumerateFiles(sFolder As String) As Variant
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder As Object: Set objFolder = objFSO.GetFolder(sFolder)
    Dim objFile As Object, V() As String

    For Each objFile In objFolder.Files
        If IsArrayAllocated(V) = False Then
            ReDim V(0)
        Else
            ReDim Preserve V(UBound(V) + 1)
        End If
        V(UBound(V)) = objFile.Name
    Next objFile

    EnumerateFiles = V
End Function

Function IsArrayAllocated(Arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(Arr) And Not IsError(LBound(Arr, 1)) And LBound(Arr, 1) <= UBound(Arr, 1)
End Function

Function IsWorkBookOpen(sFile As String) As Boolean
    On Error Resume Next
    IsWorkBookOpen = Len(Application.Workbooks(sFile).Name) > 0
End Function