Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$D$2" Or Target.Address = "$F$2" Then Call Report2
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Hight light row and column from activecell:'
    
    'Exit Sub   '<---------Bo dòng này dê chay code'
    
    Dim iRow As Long, iCol As Long, i As Long, j As Long
    Application.ScreenUpdating = False
    
    Cells.Interior.ColorIndex = 0   'Xóa màu nên cu'
    iRow = ActiveCell.Row           'Tra vê dòng cua ô hiên hành'
    iCol = ActiveCell.Column        'Tra vê côt cua ô hiên hành'
    
    For i = 1 To iRow
        Cells(i, iCol).Interior.ColorIndex = 6  'Tô màu ô cùng dòng'
    Next i
    
    For j = 1 To iCol
        Cells(iRow, j).Interior.ColorIndex = 6  'Tô màu ô cùng côt'
    Next j
    
    Application.ScreenUpdating = True
End Sub

