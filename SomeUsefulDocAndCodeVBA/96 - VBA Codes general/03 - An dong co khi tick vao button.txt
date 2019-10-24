Option Explicit
Sub hide_Conception()
    Dim shp As Shape, rowTxtBox As Integer, rngToHide As Range
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    Set shp = sh.Shapes("TextBox 164")
    rowTxtBox = shp.TopLeftCell.Row
    Set rngToHide = sh.Range("A" & (rowTxtBox - 2))

    If sh.OLEObjects("OptionButton1").Object.Value = True Then
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
    Select Case sh.Range("L" & (rowTxtBox - 3))
        Case "": rngToHide.EntireRow.Hidden = True
        Case "Non": rngToHide.EntireRow.Hidden = True
        Case "Oui": rngToHide.EntireRow.Hidden = False
        End Select
    Set rngToHide2 = sh.Range("A" & (rowTxtBox - 8))
    Select Case sh.Range("X" & (rowTxtBox - 9))
        Case "": rngToHide.EntireRow.Hidden = True
        Case "Non": rngToHide.EntireRow.Hidden = True
        Case "Oui": rngToHide.EntireRow.Hidden = False
        End Select
End Sub
