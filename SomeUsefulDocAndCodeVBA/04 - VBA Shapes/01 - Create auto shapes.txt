Option Explicit


Private Sub cmdCreateAutoshapes_Click()
    Dim i As Integer
    Dim t As Integer
    Dim shp As Shape
    
    If ActiveSheet.Shapes.Count > 2 Then
        DeleteAllAutoShapes
    End If
    
    Randomize
    t = 15
    For i = 1 To 137
        Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
        shp.TextFrame.Characters.Text = i
        If CInt(application.Version) >= 12 Then
            If i = 25 Then
                ' Treat as line
                shp.ShapeStyle = msoLineStylePreset1
            Else
                ' Randomly select a style
                shp.ShapeStyle = Int(Rnd() * 42 + 1)
            End If
        End If
        t = t + 75
    Next
    ' skip 138 - not supported
    If CInt(application.Version) >= 12 Then
        For i = 139 To 179
            Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
            shp.TextFrame.Characters.Text = i
            shp.ShapeStyle = Int(Rnd() * 42 + 1)
            t = t + 75
        Next
        ' These shapes don't have TextFrame's
        For i = 180 To 183
            Set shp = ActiveSheet.Shapes.AddShape(i, 48, t, 96, 60)
            t = t + 75
        Next
    End If
    
End Sub

Private Sub cmdDeleteShapes_Click()
    DeleteAllAutoShapes
End Sub



