Function GetStringNumber(txt As String, getStr As Boolean) As String
'If getStr = true or 1 --> Get only string
'If getStr = false, or 0 --> Get only nulber

    With CreateObject("VBScript.RegExp")
    .Pattern = IIf(getStr = True, "\d+", "\D+")
    .Global = True
    GetStringNumber = .Replace(txt, "")
    End With
End Function