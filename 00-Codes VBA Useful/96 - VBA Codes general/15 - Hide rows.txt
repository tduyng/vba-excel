Option Explicit
Sub hide_Conception()
    Dim shp As Shape, rowTxtBox As Integer, rngToHide As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("TextBox 164")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 2))

    If sh.OLEObjects("OptionButton5").Object.Value = True Then
       rngToHide.EntireRow.Hidden = False
    Else
        rngToHide.EntireRow.Hidden = True
    End If
    
End Sub

Sub hide_Clause()
    Dim shp As Shape, shp2 As Shape, rowTxtBox As Integer
    Dim sh As Worksheet, rngToHide As Range, rngToHide2 As Range
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("Graphic 247")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 2))
    'Debug.Print rowTxtBox
    Select Case sh.Range("L" & (rowTxtBox - 4))
        Case "": rngToHide.EntireRow.Hidden = True
        Case "Non": rngToHide.EntireRow.Hidden = True
        Case "Oui": rngToHide.EntireRow.Hidden = False
        End Select
    Set rngToHide2 = sh.Range("A" & (rowTxtBox - 8))
    Select Case sh.Range("X" & (rowTxtBox - 10))
        Case "": rngToHide2.EntireRow.Hidden = True
        Case "Non": rngToHide2.EntireRow.Hidden = True
        Case "Oui": rngToHide2.EntireRow.Hidden = False
        End Select
End Sub


Sub hide_CES()
    Dim shp As Shape, shp2 As Shape, rowTxtBox As Integer
    Dim sh As Worksheet, rngToHide As Range, rngToHide2 As Range, rngToHide3 As Range
    Dim numRowToHide As Integer
    Call OptimzeCode_Begin
    
    Set sh = ActiveSheet
    Set shp = sh.Shapes("Rectangle 35")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox + 1), "A" & (rowTxtBox + 34))
    Debug.Print rowTxtBox
    numRowToHide = sh.Range("G" & (rowTxtBox + 2)).Value
    Set rngToHide2 = sh.Range("A" & (rowTxtBox + 3 + numRowToHide * 2), "A" & (rowTxtBox + 34))
    Set rngToHide3 = sh.Range("A" & (rowTxtBox + 5 + numRowToHide * 2), "A" & (rowTxtBox + 34))
    If sh.Range("G" & rowTxtBox).Value = "Corps d'état séparés" Then
        rngToHide.EntireRow.Hidden = True
    ElseIf (sh.Range("G" & (rowTxtBox + 2)) = "") Or (numRowToHide = 0) Then
        MsgBox "Nb de Lots doit être supérieur à 0", vbCritical
        rngToHide.EntireRow.Hidden = False
        rngToHide2.EntireRow.Hidden = True
    ElseIf numRowToHide > 0 Then
        rngToHide.EntireRow.Hidden = False
        rngToHide3.EntireRow.Hidden = True
    Else:
        rngToHide2.EntireRow.Hidden = False
    End If
    
    Call OptimizeCode_End
End Sub

Sub hide_Decoupage()
    Dim shp As Shape, shp2 As Shape, rowTxtBox As Integer
    Dim sh As Worksheet, rngToHide As Range, rngToHide2 As Range, rngToHide3 As Range
    Dim numRowToHide As Integer
    Call OptimzeCode_Begin
    Set sh = ActiveSheet
    Set shp = sh.Shapes("Rectangle 16")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox + 1), "A" & (rowTxtBox + 18))
    
    numRowToHide = sh.Range("G" & (rowTxtBox + 2)).Value
    'Set rngToHide2 = sh.Range("A" & (rowTxtBox + 3 + numRowToHide+1), "A" & (rowTxtBox + 7 + numRowToHide))
    
    Set rngToHide3 = sh.Range("A" & (rowTxtBox + 1), "A" & (rowTxtBox + 3 + numRowToHide * 2))
    If sh.Range("G" & rowTxtBox).Value = "Non" Then
        rngToHide.EntireRow.Hidden = True
    ElseIf (numRowToHide <= 0) Then
        MsgBox "Nb de tranches doit être supérieur à 0", vbCritical
        rngToHide.EntireRow.Hidden = False
        rngToHide2.EntireRow.Hidden = True
    Else:
        rngToHide.EntireRow.Hidden = True
        rngToHide3.EntireRow.Hidden = False


    End If
    Call OptimizeCode_End
End Sub


