
Declare Function GetTickCount Lib "kernel32" () As Long
Option Explicit



Sub QUERY_NAME_SHEET()

Dim sh As Worksheet
Dim arr As Variant, i As Integer
Dim T1 As Long, T2 As Long
T1 = GetTickCount
ReDim arr(1 To ThisWorkbook.Worksheets.Count)
i = 1
For Each sh In Worksheets
    arr(i) = sh.Name
    i = i + 1
Next sh

If i = 1 Then Exit Sub
Sheet2.Range("H2").Resize(UBound(arr), 1).Value = Application.Transpose(arr)
T2 = GetTickCount
MsgBox (T2 - T1) / 1000, vbInformation, "So giay thuc hien la"
End Sub







Khi can kiem tra su dich chuyrn thoi gian mà moc thay doi lên den 1/1000 giây (1 millisecond - mili giây) thì can dùng các hàm API.
Có the liet kê các hàm do thoi gian nhu sau:
Now, Time, Timer (thu?c VB/VBA)
GetTickCount
TimeGetTime
QueryPerformanceCounter, QueryPerformanceFrequency 




Declare Function QueryPerformanceCounter Lib "Kernel32" _
                        (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                        (X As Currency) As Boolean
Declare Function GetTickCount Lib "Kernel32" () As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long




Ham Timer
Sub Test_Timer()
    '
    ' Timer Function
    '
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = Timer
    Do
      T2 = Timer
      Loops = Loops + 1
    Loop Until T1 <> T2
    
    Debug.Print "Timer minimum resolution: "; _
                (T2 - T1); "second(s)"
    Debug.Print "Took"; Loops; "loops"

End Sub




 [b]Sub DinhThoi()[/b]
  Dim TGian
 
  TGian = Timer
   [My  Code]
   . . . 
   [My code]
   Msgbox Str(Timer -TGian)
 [b] End Sub [/b]




Ham GetTickCount
Sub Test_GetTickCount()
    '
    ' GetTickCount API Function
    '
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = GetTickCount
    Do
      T2 = GetTickCount
      Loops = Loops + 1
    Loop Until T1 <> T2
    
    Debug.Print "GetTickCount minimum resolution: "; _
                (T2 - T1); "millisecond(s)"
    Debug.Print "Took"; Loops; "loops"

End Sub




Ham timeGetTime
Sub Test_timeGetTime()
    '
    ' timeGetTime API Function
    '
    Dim Loops&, T1&, T2&
    Debug.Print
    Loops = 0
    T1 = timeGetTime
    Do
      T2 = timeGetTime
      Loops = Loops + 1
    Loop Until T1 <> T2
    
    Debug.Print "timeGetTime minimum resolution: "; _
                (T2 - T1); "millisecond(s)"
    Debug.Print "Took"; Loops; "loops"

End Sub






Hàm QueryPerformanceCounter
Private Sub Test_QueryPerformanceCounter()
    '
    ' QueryPerformanceFrequency API Function
    ' QueryPerformanceCounter API Function
    '

    Dim Loops&, T1@, T2@, Freq@, Overhead@, I&
   
    QueryPerformanceFrequency Freq
    QueryPerformanceCounter T1
    QueryPerformanceCounter T2
    Overhead = T2 - T1        
    QueryPerformanceCounter T1  
     
    Do
      QueryPerformanceCounter T2
      Loops = Loops + 1
    Loop Until T1 <> T2
     
    Debug.Print (T2 - T1 - Overhead) / Freq * 1000; "milliseconds(ms)"
    Debug.Print "Took"; Loops; "loops"
   
End Sub




'Chia se giaiphapexcel

 Dim StartTime As Double

'   B?t d?u tính th?i gian
   StartTime = Timer

'   Các do?n mã c?a b?n ? dây
'
'

' K?t thúc do?n mã c?a b?n
' Tính và thông báo th?i gian th?c hi?n

    MsgBox Format(Timer - StartTime, "00.00") & " giây."



'--------------------------------------
'cach khac
Public Declare Function QueryPerformanceFrequency _
                         Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter _
                         Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long

Sub CalculateTime()
    Dim Ar(1 To 20, 1 To 4) As Currency, WS As Worksheet
    Dim n As Currency, str As Currency, fin As Currency
    Dim y As Currency
    Dim i As Long, j As Long
    Application.ScreenUpdating = False
    For i = 1 To 4
        For j = 1 To 20
            Set WS = ThisWorkbook.Sheets.Add
            WS.Range(“A1”).Value = 1
            QueryPerformanceFrequency y
            QueryPerformanceCounter str
            Select Case i
            Case 1: Macro1
            Case 2: Macro1_Version2
            Case 3: Macro1_Version3
            Case 4: Macro1_Version4
            End Select
            QueryPerformanceCounter fin
            Application.DisplayAlerts = False
            WS.Delete
            Application.DisplayAlerts = True
            n = (fin - str)
            Ar(j, i) = CCur(Format(n, "##########.############") / y)
        Next j
    Next i
    With Range(“A1”).Resize(1, 4)
        .Value = Array("Macro1", "Macro2", "Macro3", "Macro4")
        .Font.Bold = True
    End With
    Range("A2").Resize(20, 4).Value = Ar
    With Range("A22").Resize(1, 4)
        .FormulaR1C1 = "=AVERAGE(R2C:R21C)"
        .Offset(1).FormulaR1C1 = "=RANK(R22C,R22C1:R22C4,1)"
        .Resize(2).Font.Bold = True
    End With
    Application.ScreenUpdating = True
End Sub

