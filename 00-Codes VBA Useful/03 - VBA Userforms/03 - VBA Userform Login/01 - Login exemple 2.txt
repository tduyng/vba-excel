Public Username As String
   
'Source ref: https://www.c-sharpcorner.com/UploadFile/88d8c0/create-login-application-in-excel-macro-using-visual-basic-f/

02.Public Password As String   
03.Public i As Integer   
04.Public j As Integer   
05.Public u As String   
06.Public p As String    
07.   
08.Private Sub CommandButton1_Click ()   
09.    Application.ScreenUpdating = False   
10.    If Trim (TextBox1.Text) = "" And Trim (TextBox2.Text) = "" Then   
11.        MsgBox "Enter username and password.", vbOKOnly   
12.        Else If Trim (TextBox1.Text) = "" Then   
13.        MsgBox "Enter the username ", vbOKOnly   
14.        Else If Trim(TextBox2.Text) = "" Then   
15.        MsgBox "Enter the Password ", vbOKOnly   
16.    Else   
17.        Username = Trim (TextBox1.Text)   
18.        Password = Trim (TextBox2.Text)   
19.        i = 1   
20.        Do While Cells (1, 1).Value <> ""   
21.            j = 1   
22.            u = Cells (i, j).Value   
23.            j = j + 1   
24.            p = Cells (i, j).Value   
25.            If Username = u And Password = p And Cells (i, 3).Value = "fail" Then   
26.                MsgBox "Your Account temporarily locked", vbCritical   
27.                Exit Do   
28.                Else If Username = u And Password = p Then   
29.                Call clear   
30.                UserForm1.Hide   
31.                UserForm2.Label1.Caption = u   
32.                UserForm2.Label1.ForeColor = &H8000000D   
33.                UserForm2.Show   
34.                Exit Do   
35.                Else If Username <> u And Password = p Then   
36.                MsgBox "Username not matched", vbCritical + vbOKCancel   
37.                Exit Do   
38.                Else If Username = u And Password <> p Then   
39.                If Cells (i, 3).Value = "fail" Then   
40.                    MsgBox "Your account is blocked", vbCritical + vbOKCancel   
41.                    Exit Do   
42.                    Else If Cells (i, 4).Value < 2 Then   
43.                    MsgBox "Invalid password", vbCritical   
44.                    Cells (i, 4).Value = Cells (i, 4) + 1   
45.                    Exit Do   
46.                Else   
47.                    Cells (i, 4).Value = Cells (i, 4) + 1   
48.                    Cells (i, 3).Value = "fail"   
49.                    Cells (i, 2).Interior.ColorIndex = 3   
50.                    Exit Do   
51.                End If   
52.            Else   
53.                i = i + 1   
54.            End If   
55.        Loop   
56.    End If   
57.    Application.ScreenUpdating = True   
58.End Sub   
59.Sub clear ()   
60.    TextBox1.Value = ""   
61.    TextBox2.Value = ""   
62.End Sub   
63.Private Sub TextBox1_Enter ()   
64.    With TextBox1   
65.        .Back Color = &H8000000E   
66.        .Fore Color = &H80000001   
67.        .Border Color = &H8000000D   
68.    End With    
69.    TextBox1.Text = ""   
70.End Sub   
71.Private Sub TextBox1_AfterUpdate ()   
72.    If TextBox1.Value = "" Then   
73.        TextBox1.BorderColor = RGB (255, 102, 0)   
74.    End If   
75.    i = 1   
76.    Do Until Is Empty (Cells (i, 1).Value)   
77.        If TextBox1.Value = Cells (i, 1).Value Then   
78.            With TextBox1   
79.                .Border Color = RGB (186, 214, 150)   
80.                .Back Color = RGB (216, 241, 211)   
81.                .Fore Color = RGB (81, 99, 51)   
82.            End With   
83.        End If   
84.        i = i + 1   
85.    Loop   
86.End Sub   
87.Private Sub TextBox2_Enter ()   
88.    With TextBox2   
89.        .Back Color = &H8000000E   
90.        .Fore Color = &H80000001   
91.        .Border Color = &H8000000D   
92.    End With   
93.    TextBox2.Text = ""   
94.End Sub   
95.Private Sub TextBox2_AfterUpdate ()   
96.    i = 1   
97.    Username = TextBox1.Value   
98.    Password = TextBox2.Value   
99.    If TextBox2.Text = "" Then   
100.        TextBox2.BorderColor = RGB (255, 102, 0)   
101.    End If   
102.    Do Until Is Empty (Cells (i, 1).Value)   
103.        j = 1   
104.        u = Cells (i, j).Value   
105.        j = j + 1   
106.        p = Cells (i, j).Value   
107.        If Username = u and Password = p Then   
108.            With TextBox2   
109.                .Border Color = RGB (186, 214, 150)   
110.                .Back Color = RGB (216, 241, 211)   
111.                .Fore Color = RGB (81, 99, 51)   
112.            End With   
113.            Exit Do   
114.            Else If Username = u and Password <> p Then   
115.            TextBox2.BorderColor = RGB (255, 102, 0)   
116.            Exit Do   
117.        Else   
118.            i = i + 1   
119.        End If   
120.    Loop   
121.End Sub   
122.Sub settings ()   
123.    With UserForm1   
124.        TextBox1.ForeColor = &H8000000C   
125.        TextBox2.ForeColor = &H8000000C   
126.        TextBox1.BackColor = &H80000004   
127.        TextBox2.BackColor = &H80000004   
128.        TextBox1.Text = "Username"   
129.        TextBox2.Text = "Password"   
130.        TextBox1.BorderColor = RGB (0, 191, 255)   
131.        TextBox2.BorderColor = RGB (0, 191, 255)   
132.        CommandButton1.SetFocus   
133.    End With   
134.End Sub   
135.Private Sub UserForm_Initialize ()   
136.    Call settings   
137.End Sub  
