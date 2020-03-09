Option Explicit

Function RegExample(inputString As String, _
                    Optional globalRe As Boolean = True, _
                    Optional ignoreCase As Boolean = True) As Boolean
    
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")

    With re
        .Global = globalRe
        .Pattern = "^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$"
        .ignoreCase = ignoreCase
        RegExample = .test(inputString)
    End With

Sub timtheo_vitri()
    Dim str As String, kqua
    str = "Lua nep la lua nep lang,lua len lop lop ,long nang lang lang "
    With CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = True
        .Pattern = "Lua"
            Set kqua = .Execute(str) ' Kqua se tra ve 1 mang gom 3 chu : Lua lua lua
        .Pattern = "^Lua"
            Set kqua = .Execute(str) ' Kqua se tra ve 1 gia tri chu Lua dau tien
    End With
' Tuong tu co the ap dung cho cac truong hop khac ^^
End Sub

Sub Regx1()
    Dim str As String, kqua
    str = " Messi dang co bong, anh sut bong, vaooooo)"
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "[vao]{5}"
        Set kqua = .Execute(str) ' Tra ket qua la chu vaooo
        .IgnoreCase = True
        .Pattern = "[a-z]* +"
        Set kqua = .Execute(str) ' Tach tat ca cac chu cai tu a-z  va dang sau la 1 dau cach
    End With
End sub


Dim rRange As Range
Dim rCell As Range

Dim re As Object
Set re = CreateObject("vbscript.regexp")
With re
  .Pattern = "^\d\d?:\d\d[aApP][mM] - \d\d?:\d\d[aApP][mM]$"
  .Global = False
  .IgnoreCase = False
End With

Set rRange = Range("A2", "G225")

For Each rCell In rRange.Cells
    If re.Test(rCell) Then
        rCell.Interior.Color = RGB(0, 250, 0)
    Else
        rCell.Interior.Color = RGB(250, 0, 0)
    End If
Next rCell