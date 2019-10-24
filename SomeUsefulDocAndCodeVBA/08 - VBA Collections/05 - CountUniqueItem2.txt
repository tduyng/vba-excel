Public Function CountUnique(ListRange As Range) As Integer

Dim CellValue As Variant

Dim UniqueValues As New Collection

Application.Volatile

On Error Resume Next

For Each CellValue In ListRange

UniqueValues.Add CellValue, CStr(CellValue) ' add the unique item

Next

CountUnique = UniqueValues.count

End Function