Function MYMATCH(strValue As String, strPattern As String, Optional blnCase As Boolean = True, Optional blnBoolean = True) As String
    Dim objRegEx As Object
    Dim strPosition As Integer
    Dim RegMC

    ' Create regular expression.
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        .Pattern = strPattern
        .IgnoreCase = blnCase
        If .test(strValue) Then
            Set RegMC = .Execute(strValue)
            MYMATCH = RegMC(0).firstindex + 1
        Else
            MYMATCH = "no match"
        End If
    End With
End Function

Sub TestMe()
    MsgBox MYMATCH("test 1", "\d+")
End Sub