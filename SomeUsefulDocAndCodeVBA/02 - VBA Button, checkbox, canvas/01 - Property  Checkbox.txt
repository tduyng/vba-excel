Sub CreateFormsCheckBoxes()

    Dim objCtrl As Object
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim shp As Shape
    Dim strCaption As String
    Dim lngCapLen As Long
        
    strCaption = "My Check Box" 'Edit to required caption for checkbox.
    lngCapLen = Len(strCaption)
    
    With Sheets("Sheet1")
        With ActiveCell
            'Position and size with respect to the cell location
            'Size and position adjusted slightly to prevent grid lines being covered
            dblLeft = .Left + 1
            dblWidth = .Width * 3 - 2
            dblHeight = .Height - 2
            dblTop = .Top + 1
        End With
        Set objCtrl = .CheckBoxes.Add(dblLeft, dblTop, dblWidth, dblHeight)
        objCtrl.Name = "chkBox1"
        objCtrl.Characters.Text = ""    'No name within the CheckBox
        
        'Superimpose a TextBox over the CheckBox
        dblLeft = dblLeft + 15      'Move left of TextBox to right so Check is visible
        dblWidth = dblWidth - 15    'Shorten width of TextBox to align right side of CheckBox
        Set shp = .Shapes.AddTextbox(msoTextOrientationHorizontal, dblLeft, dblTop, dblWidth, dblHeight)
    End With
    
    shp.TextFrame2.TextRange.Characters.Text = strCaption       'Insert the caption
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle     'Centre font top to bottom
    shp.Line.Visible = msoFalse              'No visible border
    
    With shp.TextFrame2.TextRange.Characters(1, lngCapLen).Font
        .NameComplexScript = "Times New Roman"
        .Size = 12
        .Bold = msoTrue
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        
        '**************************************************************
        'The following are in regard to the Font; not the shape fill.
        'All untested.
        'If shape is right clicked on the worksheet
        'and the Drawing Tools ribbon that appears is selected
        'then see the Word Art Styes block to which these apply to.
        '.Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        '.Fill.ForeColor.TintAndShade = 0
        '.Fill.ForeColor.Brightness = 0
        '.Fill.Transparency = 0
        '.Fill.Solid
        '***************************************************************
    End With
    
    ActiveCell.Offset(2, 0).Select  'Move selection away from shapes
    
End Sub


Sub ClearShapes()
'Use this sub to delete all of the CheckBoxes and TextBoxes during testing.
Dim shp As Shape

For Each shp In ActiveSheet.Shapes
    shp.Delete
Next shp

End Sub