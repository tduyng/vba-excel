Const accent As String = "ÀÁÂÃÄÅàáâãäåÒÓÔÕÖØòóôõöøÈÉÊËèéêëÌÍÎÏìíîïÙÚÛÜùúûüÿÑñÇç"
Const noAccent As String = "AAAAAAaaaaaaOOOOOOooooooEEEEeeeeIIIIiiiiUUUUuuuuyNnCc"


'*****************************************************************************'
'Ici pour exécuter le macro - ne pas besoin d'utiliser l'option "Option Explicite"
' Ce macro est n'utilisé que pour enlever des accents en français, pas pour les accents vietnamiens

Public Sub Remove_Accent()
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeConstants, 23)
If TypeName(c.Value) = "String" Then
    If InStr(1, c.Text, "@") > 0 Then

    c = noneAccents(c.Text)
    Else
    c = noneAccents(c.Text)
    End If
End If
Next

End Sub


'End macro
'*****************************************************************************'



'================================================================================'
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
'Ligne code pour la fonction - Ne pas toucher ici  - Dont touche it, please



Private Function noneAccents(ByRef s As String) As String
Dim i As Integer
Dim lettre As String * 1
noneAccents = s
For i = 1 To Len(accent)
lettre = Mid$(accent, i, 1)
If InStr(noneAccents, lettre) > 0 Then
noneAccents = Replace(noneAccents, lettre, Mid$(noAccent, i, 1))
End If
Next i
End Function


'End function
'================================================================================'