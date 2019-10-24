Function GetNextAvailableName(ByVal strPath As String) As String

    With CreateObject("Scripting.FileSystemObject")

        Dim strFolder As String, strBaseName As String, strExt As String, i As Long
        strFolder   = .GetParentFolderName(strPath)
        strBaseName = .GetBaseName(strPath)
        strExt      = .GetExtensionName(strPath)

        Do While .FileExists(strPath)
            i = i + 1
            strPath = .BuildPath(strFolder, strBaseName & " (" & i & ")" & "." & strExt)
        Loop

    End With

    GetNextAvailableName = strPath

End Function