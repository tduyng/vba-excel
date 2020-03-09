    'Annuler toutes les fusions cells
    Dim cell As Range, joinedCells As Range
    For Each cell In rngUsed
        If cell.MergeCells Then
            Set joinedCells = cell.MergeArea
            cell.MergeCells = False
            joinedCells.value = cell.value
        End If
    Next cell


'This Function will return all the merged cells in your range.
Function getMergedCells(rngRow As Range) As Range

        Dim rngCell     As Range
        Dim rngMerged   As Range

        For Each rngCell In rngRow.Cells
            If rngCell.MergeCells Then

              If rngMerged Is Nothing Then
                Set rngMerged = rngCell
              Else
                Set rngMerged = Union(rngMerged, rngCell)
              End If

            End If
        Next

        Set getMergedCells = rngMerged

    End Function
Sub test()

    Dim rngTest As Range
    Set rngTest = getMergedCells(Selection)

    If Not rngTest Is Nothing Then
        MsgBox "Merged cells :" & rngTest.Address
    Else
        MsgBox "No cells are merged"
    End If

    End Sub