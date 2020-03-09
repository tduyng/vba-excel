Option Explicit

Private m_Worksheet As Worksheet

Private Sub cmdClearShapes_Click()
    DeleteAllAutoShapes 'in module MCommon
End Sub

Private Sub cmdRunDemo_Click()

    Dim oShape1 As Shape
    Dim oShape2 As Shape
    Dim oConnector As Shape
    
    Set m_Worksheet = ActiveSheet
    
    If m_Worksheet.Shapes.Count > 2 Then
        DeleteAllAutoShapes
    End If
    
    Set oShape1 = AddShapeToRange(msoShapeFlowchartProcess, "B3:D5")
    AddFormattedTextToShape oShape1, "First Shape"
    
    Set oShape2 = AddShapeToRange(msoShapeFlowchartAlternateProcess, "B9:D11")
    AddFormattedTextToShape oShape2, "Second Shape"
    
    Set oConnector = AddConnectorBetweenShapes(msoConnectorStraight, oShape1, oShape2)
    
    If CInt(application.Version) < 12 Then
        FormatShape2003 oShape1
        FormatShape2003 oShape2
        FormatConnector2003 oConnector
        'IlluminateShapeText2003
    Else
        FormatShape2007 oShape1
        FormatShape2007 oShape2
        FormatConnector2007 oConnector
        'IlluminateShapeText2007
    End If
    
    Set oShape1 = Nothing
    Set oShape2 = Nothing
    Set oConnector = Nothing
    
    Set m_Worksheet = Nothing

End Sub

Private Function AddShapeToRange(ShapeType As MsoAutoShapeType, _
                                 sAddress As String) As Shape
    With m_Worksheet.Range(sAddress)
        Set AddShapeToRange = m_Worksheet.Shapes.AddShape(ShapeType, .Left, .Top, .Width, .Height)
    End With
End Function

Private Sub AddFormattedTextToShape(oShape As Shape, _
                                    sText As String)
    If Len(sText) > 0 Then
        With oShape.TextFrame
            .Characters.Text = sText
            .Characters.Font.Name = "Garamond"
            .Characters.Font.Size = 12
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
End Sub

Private Function AddConnectorBetweenShapes(ConnectorType As MsoConnectorType, _
                                          oBeginShape As Shape, _
                                          oEndShape As Shape) As Shape

    ' NOTE: These connection site constants only work for rectangular shapes with
    '       4 connection points. The call to RerouteConnections below will
    '       automatically reroute the connector to the shortest path between the shapes.
    Const TOP_SIDE As Integer = 1
    Const LEFT_SIDE As Integer = 2
    Const BOTTOM_SIDE As Integer = 3
    Const RIGHT_SIDE As Integer = 4
    
    Dim oConnector As Shape
    Dim x1 As Single
    Dim x2 As Single
    Dim y1 As Single
    Dim y2 As Single
    
    With oBeginShape
        x1 = .Left + .Width / 2
        y1 = .Top + .Height
    End With
    
    With oEndShape
        x2 = .Left + .Width / 2
        y2 = .Top
    End With
    
    ' Excel 2007 uses absolute coordinates for the second point,
    ' of the AddConnector function. Previous versions of Excel
    ' use relative coordinates. But, ... (continued below)
    If CInt(application.Version) < 12 Then
        x2 = x2 - x1
        y2 = y2 - y1
    End If
    
    Set oConnector = m_Worksheet.Shapes.AddConnector(ConnectorType, x1, y1, x2, y2)
    
    ' ... you can use any positive Single values if you connect
    ' the end points with BeginConnect and EndConnect:
    oConnector.ConnectorFormat.BeginConnect oBeginShape, BOTTOM_SIDE
    oConnector.ConnectorFormat.EndConnect oEndShape, TOP_SIDE
    oConnector.RerouteConnections
    
    Set AddConnectorBetweenShapes = oConnector
    
    Set oConnector = Nothing

End Function

Private Sub FormatConnector2003(oConnector As Shape)
    If oConnector.Connector Or oConnector.Type = msoLine Then
        ' rough approximation of the Excel 2007 preset line style #17
        With oConnector
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.Weight = 2
            .Line.ForeColor.RGB = RGB(192, 0, 0)
        End With
    End If
End Sub

Private Sub FormatConnector2007(oConnector As Shape)
    With oConnector
        If .Connector Or .Type = msoLine Then
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .ShapeStyle = msoLineStylePreset17
        End If
    End With
End Sub

Private Sub FormatShape2003(oShape As Shape)
    ' rough approximation of the Excel 2007 preset shape style #2
    With oShape
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(79, 129, 189)
        .Line.Weight = 2
        .Shadow.OffsetX = 0.8
        .Shadow.OffsetY = 0.8
        .Shadow.ForeColor.RGB = RGB(192, 192, 192)
        .Shadow.Transparency = 0.5
        .Shadow.Visible = msoTrue
    End With
End Sub

Private Sub FormatShape2007(oShape As Shape)
    oShape.ShapeStyle = msoShapeStylePreset2
End Sub

Private Sub IlluminateShapeText2003()
    On Error Resume Next

    Dim oShape As Shape
    Dim numChars As Integer
    
    For Each oShape In m_Worksheet.Shapes
        If oShape.Type = msoAutoShape And oShape.AutoShapeType > 0 Then
            ' Throws error when shape doesn't contain text:
            numChars = oShape.TextFrame.Characters.Count
            If Err.Number <> 0 Then
                Debug.Print oShape.Name
                Err.Clear
            ElseIf numChars > 0 Then
                With oShape.TextFrame.Characters(1, 1).Font
                    .Name = "Garamond"
                    ' The following does not work in Excel 2003 and below.
                    ' It rounds color to nearest preset.
                    .Color = RGB(192, 80, 77)
                    .Size = 24
                    .Bold = True
                End With
            End If
        End If
    Next
End Sub

Private Sub IlluminateShapeText2007()
    Dim oShape As Shape
    
    For Each oShape In m_Worksheet.Shapes
        If oShape.Type = msoAutoShape And oShape.AutoShapeType > 0 Then
            ' Does not throw error when shape doesn't contain text:
            If oShape.TextFrame2.HasText Then
                With oShape.TextFrame2.TextRange.Characters(1, 1).Font
                    .Name = "Garamond"
                    ' Sets color properly - i.e. no rounding like
                    ' earlier versions of Excel.:
                    .Fill.ForeColor.RGB = RGB(192, 80, 77)
                    .Size = 24
                    .Bold = True
                    .Reflection.Type = msoReflectionType1
                    .Shadow.Type = msoShadow14
                    .Glow.Color.RGB = RGB(192, 137, 45)
                    .Glow.radius = 5
                End With
            End If
            
        End If
    Next
End Sub


