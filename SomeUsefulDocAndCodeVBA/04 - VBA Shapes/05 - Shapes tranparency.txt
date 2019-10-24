Public sh2 As Worksheet

Sub Emoj_risque1()
'Macro pour l'évaluation des riquessur les intervenants
'======================================================='
    Dim sh As Worksheet
    Set sh = ActiveSheet
    Set sh2 = Frisque
    Dim shp As Shape, rngToHide As Range
    Dim rowTxtBox As Integer, ColTxtBox As Integer, rowsData As Integer
    
    Set shp = sh.Shapes("TextBox 99")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 11), "A" & rowTxtBox)
    
    If sh.Range("o" & (rowTxtBox - 8)) = "Non examiné" Or sh.Range("o" & (rowTxtBox - 8)) = "" Then
        With sh.Shapes("Graphic 199").Fill
            .Transparency = 0.7
        End With
        
        sh2.Shapes("Graphic 2").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 4").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 5").Fill.Transparency = 0.7
        
         With sh.Shapes("Graphic 201").Fill
            .Transparency = 0.7
        End With
         With sh.Shapes("Graphic 202").Fill
            .Transparency = 0.7
        End With
    ElseIf sh.Range("o" & (rowTxtBox - 8)) = "Sans objet" Then
        With sh.Shapes("Graphic 199").Fill
            .Transparency = 0
        End With
         With sh.Shapes("Graphic 201").Fill
            .Transparency = 0.7
        End With
         With sh.Shapes("Graphic 202").Fill
            .Transparency = 0.7
        End With
        sh2.Shapes("Graphic 2").Fill.Transparency = 0
        sh2.Shapes("Graphic 4").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 5").Fill.Transparency = 0.7
        
     ElseIf sh.Range("o" & (rowTxtBox - 8)) = "Courant" Then
        With sh.Shapes("Graphic 199").Fill
            .Transparency = 0.7
        End With
         With sh.Shapes("Graphic 201").Fill
            .Transparency = 0
        End With
         With sh.Shapes("Graphic 202").Fill
            .Transparency = 0.7
        End With
        sh2.Shapes("Graphic 2").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 4").Fill.Transparency = 0
        sh2.Shapes("Graphic 5").Fill.Transparency = 0.7
        
      ElseIf sh.Range("o" & (rowTxtBox - 8)) = "Important" Then
        With sh.Shapes("Graphic 199").Fill
            .Transparency = 0.7
        End With
         With sh.Shapes("Graphic 201").Fill
            .Transparency = 0.7
        End With
         With sh.Shapes("Graphic 202").Fill
            .Transparency = 0
        End With
        sh2.Shapes("Graphic 2").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 4").Fill.Transparency = 0.7
        sh2.Shapes("Graphic 5").Fill.Transparency = 0
        
    End If
    
End Sub