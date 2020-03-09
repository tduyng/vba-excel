'// Sort A-Z cac Item trong Collection'
Public Sub SortingCollection(myCol As Collection, firstIndex As Long, lastIndex As Long)
  Dim valCentre As Variant, vTemp As Variant
  Dim valMin As Long
  Dim valMax As Long
  valMin = firstIndex
  valMax = lastIndex
  valCentre = myCol((firstIndex + lastIndex) \ 2)
  Do While valMin <= valMax
    Do While myCol(valMin) < valCentre And valMin < lastIndex
      valMin = valMin + 1
    Loop
    Do While valCentre < myCol(valMax) And valMax > firstIndex
      valMax = valMax - 1
    Loop
    If valMin <= valMax Then
      ' Swap values
      vTemp = myCol(valMin)
      myCol.Add myCol(valMax), After:=valMin
      myCol.Remove valMin
      myCol.Add vTemp, before:=valMax
      myCol.Remove valMax + 1
      ' Move to next positions
      valMin = valMin + 1
      valMax = valMax - 1
    End If
  Loop
  If firstIndex < valMax Then SortingCollection myCol, firstIndex, valMax
  If valMin < lastIndex Then SortingCollection myCol, valMin, lastIndex
End Sub
