Sub Navigation_1()
'
' Navigation_1 Macro
'
' Touche de raccourci du clavier: Ctrl+l

'
    ActiveSheet.Range("A11").Select
End Sub
Sub Image8_Cliquer()

End Sub
Sub Navigation_2()
'
' Navigation_2 Macro
'
' Touche de raccourci du clavier: Ctrl+p
    Dim shp As Shape, rowTxtBox As Integer, rngToHide As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Image 217")
    rowTxtBox = shp.TopLeftCell.Row + 15
    
'
    Range("B" & rowTxtBox).Select
End Sub

Sub Navigation_F()
'
' Navigation_F Macro
'
' Touche de raccourci du clavier: Ctrl+f
    Dim shp As Shape, rowTxtBox As Integer, rngToHide As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Image 328")
    rowTxtBox = shp.TopLeftCell.Row + 10
'
    Range("B" & rowTxtBox).Select
End Sub

