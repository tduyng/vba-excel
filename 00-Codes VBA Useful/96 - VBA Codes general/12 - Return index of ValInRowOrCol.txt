Public Function IndexOfValInRowOrCol( _
                                    ByVal searchVal As String, _
                                    Optional ByRef ws As Worksheet = Nothing, _
                                    Optional ByRef rng As Range = Nothing, _
                                    Optional ByRef vertical As Boolean = True, _
                                    Optional ByRef rowOrColNum As Long = 1 _
                                    ) As Long

    'Returns position in Row or Column, or 0 if no matches found

    Dim usedRng As Range, result As Variant, searchRow As Long, searchCol As Long

    result = CVErr(9999) '- generate custom error

    Set usedRng = GetUsedRng(ws, rng)
    If Not usedRng Is Nothing Then
        If rowOrColNum < 1 Then rowOrColNum = 1
        With Application
            If vertical Then
                result = .Match(searchVal, rng.Columns(rowOrColNum), 0)
            Else
                result = .Match(searchVal, rng.Rows(rowOrColNum), 0)
            End If
        End With
    End If
    If IsError(result) Then IndexOfValInRowOrCol = 0 Else IndexOfValInRowOrCol = result
End Function