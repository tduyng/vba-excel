Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub DeleteAllAutoShapes()
    
    'Normally, you can do the following:
    '
    '   Worksheet.Shapes.SelectAll
    '   Selection.Delete
    '
    ' However, that would delete the command buttons as well.
    ' So the code below was written to demonstrate one way to
    ' use the ShapeRange object
    
    Dim arrShapeNames()     As Variant 'must be Variant to work with Shapes.Range()
    Dim shp                 As Excel.Shape
    Dim sr                  As Excel.ShapeRange
    Dim ws                  As Worksheet
    Dim i                   As Integer
    
    Set ws = ActiveSheet
    
    ReDim arrShapeNames(ws.Shapes.Count - 1)
    i = 0
    For Each shp In ws.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoCallout Then 'no buttons!
            arrShapeNames(i) = shp.Name
            i = i + 1
        End If
    Next
    
    If i = 0 Then Exit Sub  ' there were no valid shapes
    
    ' The first ReDim may have included room for invalid
    ' shape types - in this case the buttons. This ReDim
    ' trims the array size down accordingly
    ReDim Preserve arrShapeNames(i - 1)
    Set sr = ws.Shapes.Range(arrShapeNames)
    sr.Select
    sr.Delete
    
    Set sr = Nothing
    Set ws = Nothing
    
End Sub

Public Sub SleepAndDo(NumberOfMilliseconds As Long)
  Sleep NumberOfMilliseconds
  DoEvents
End Sub

