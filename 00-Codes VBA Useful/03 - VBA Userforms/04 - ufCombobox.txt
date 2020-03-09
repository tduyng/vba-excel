Sub CALL_UFCOMBOBOX()
    ufCombobox.Show
End Sub


Option Explicit

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbSetValue_Click()
    MsgBox "testList"
    lbTotal.Caption = getTotal
    'Unload Me
End Sub



Private Sub cbxData_Change()
    lbTotal.Caption = getTotal
End Sub

Private Sub UserForm_Initialize()

'    For Each rng In Sheet2.Range("A2:A25")
'        cbxData.AddItem rng.Value
'    Next rng


'
    cbxData.List = getList(Sheet2.Range("A2:B23"), 1)

End Sub

Public Property Get getTotal() As Variant
    Dim arr1 As Variant, arr2 As Variant
    Dim i As Long, rng As Range
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Set rng = Sheet2.Range("A2:B23")
    
    'On Error Resume Next
    
    arr1 = getList(rng, 1)
    arr2 = getList(rng, 2)
    For i = 1 To UBound(arr1)
            Dic.Item(arr1(i, 1)) = arr2(i, 1)
    Next i
    
    If cbxData.value = "" Then
        getTotal = "Total: 00"
    Else
        getTotal = "Total: " & Dic.Item(cbxData.value)
    End If
    
End Property















 Sub TEST_FUNCTION()
    Sheet2.Range("H1").Resize(8, 1) = getList(Sheet2.Range("A2:B23"), 1)
    Sheet2.Range("I1").Resize(8, 1) = getList(Sheet2.Range("A2:B23"), 2)
 End Sub
Public Function getList(ByVal rng As Range, iArr As Integer) As Variant
    'iArr = 1: mang1; iArray = 2: Mang2)
    
    Dim Dic As Object
    Dim arr As Variant
    Dim i As Long, total As Variant
    
    
    Set Dic = CreateObject("Scripting.Dictionary")
    arr = rng.value

    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not Dic.Exists(arr(i, 1)) Then
            Dic.Add arr(i, 1), arr(i, 2)
        Else
            Dic.Item(arr(i, 1)) = Dic.Item(arr(i, 1)) + arr(i, 2)
        End If
    Next i
    
    If iArr = 1 Then
        getList = Application.Transpose(Dic.Keys)
    ElseIf iArr = 2 Then
        getList = Application.Transpose(Dic.items)
    Else
        MsgBox "iArr muse be 1 or 2"
    End If
  
End Function


