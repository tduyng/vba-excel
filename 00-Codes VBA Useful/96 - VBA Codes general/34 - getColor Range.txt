Public Function iColor(rng As Range, Optional formatType As String) As Variant
'formatType: Hex for #RRGGBB, RGB for (R, G, B) and IDX for VBA Color Index
    Dim colorVal As Variant
    colorVal = rng.DisplayFormat.Interior.Color
    Select Case UCase(formatType)
        Case "HEX"
            iColor = "#" & Format(Hex(colorVal Mod 256),"00") & _
                           Format(Hex((colorVal \ 256) Mod 256),"00") & _
                           Format(Hex((colorVal \ 65536)),"00")
        Case "RGB"
            iColor = Format((colorVal Mod 256),"00") & ", " & _
                     Format(((colorVal \ 256) Mod 256),"00") & ", " & _
                     Format((colorVal \ 65536),"00")
        Case "IDX"
            iColor = rng.Interior.ColorIndex
        Case Else
            iColor = colorVal
    End Select
End Function

'Example use of the iColor function
Sub Get_Color_Format()
    Dim rng As Range

    For Each rng In Selection.Cells
        rng.Offset(0, 1).Value = iColor(rng, "HEX")
        rng.Offset(0, 2).Value = iColor(rng, "RGB")
    Next
End Sub


iNbLgt = getLastRowCol(4, sht_Planning, colApp)