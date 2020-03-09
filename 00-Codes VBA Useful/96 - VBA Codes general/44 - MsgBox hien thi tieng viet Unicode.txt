'****************************************
'Tac gia: Nguyen Duy Tuan
'Tel    : 0904.210.337
'E.Mail : tuanktcdcn@yahoo.com
'Website: www.bluesofts.net
'****************************************
'Khai báo các hàm API trong thu vi?n User32.DLL

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant = vbNullString) As VbMsgBoxResult

'BStrMsg,BStrTitle : Là chu?i Unicode
Dim BStrMsg, BStrTitle
    'Hàm StrConv Chuy?n chu?i v? mã Unicode
    BStrMsg = StrConv(PromptUni, vbUnicode)
    BStrTitle = StrConv(TitleUni, vbUnicode)
    'Hi?n thông báo
    MsgBoxUni = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons)
End Function

'==================================================================
'Hàm TCVN3toUNICODE, VNItoUNICODE du?c vi?t b?i Bình - OverAC
'www.giaiphapexcel.com
Function TCVN3toUNICODE(vnstr As String)
Dim c As String, i As Integer
    For i = 1 To Len(vnstr)
        c = Mid(vnstr, i, 1)
        Select Case c
        Case "a": c = ChrW$(97)
        Case "¸": c = ChrW$(225)
        Case "µ": c = ChrW$(224)
        Case "¶": c = ChrW$(7843)
        Case "·": c = ChrW$(227)
        Case "¹": c = ChrW$(7841)
        Case "¨": c = ChrW$(259)
        Case "¾": c = ChrW$(7855)
        Case "»": c = ChrW$(7857)
        Case "¼": c = ChrW$(7859)
        Case "½": c = ChrW$(7861)
        Case "Æ": c = ChrW$(7863)
        Case "©": c = ChrW$(226)
        Case "Ê": c = ChrW$(7845)
        Case "Ç": c = ChrW$(7847)
        Case "È": c = ChrW$(7849)
        Case "É": c = ChrW$(7851)
        Case "Ë": c = ChrW$(7853)
        Case "e": c = ChrW$(101)
        Case "Ð": c = ChrW$(233)
        Case "Ì": c = ChrW$(232)
        Case "Î": c = ChrW$(7867)
        Case "Ï": c = ChrW$(7869)
        Case "Ñ": c = ChrW$(7865)
        Case "ª": c = ChrW$(234)
        Case "Õ": c = ChrW$(7871)
        Case "Ò": c = ChrW$(7873)
        Case "Ó": c = ChrW$(7875)
        Case "Ô": c = ChrW$(7877)
        Case "Ö": c = ChrW$(7879)
        Case "o": c = ChrW$(111)
        Case "ã": c = ChrW$(243)
        Case "ß": c = ChrW$(242)
        Case "á": c = ChrW$(7887)
        Case "â": c = ChrW$(245)
        Case "ä": c = ChrW$(7885)
        Case "«": c = ChrW$(244)
        Case "è": c = ChrW$(7889)
        Case "å": c = ChrW$(7891)
        Case "æ": c = ChrW$(7893)
        Case "ç": c = ChrW$(7895)
        Case "é": c = ChrW$(7897)
        Case "¬": c = ChrW$(417)
        Case "í": c = ChrW$(7899)
        Case "ê": c = ChrW$(7901)
        Case "ë": c = ChrW$(7903)
        Case "ì": c = ChrW$(7905)
        Case "î": c = ChrW$(7907)
        Case "i": c = ChrW$(105)
        Case "Ý": c = ChrW$(237)
        Case "×": c = ChrW$(236)
        Case "Ø": c = ChrW$(7881)
        Case "Ü": c = ChrW$(297)
        Case "Þ": c = ChrW$(7883)
        Case "u": c = ChrW$(117)
        Case "ó": c = ChrW$(250)
        Case "ï": c = ChrW$(249)
        Case "ñ": c = ChrW$(7911)
        Case "ò": c = ChrW$(361)
        Case "ô": c = ChrW$(7909)
        Case "­": c = ChrW$(432)
        Case "ø": c = ChrW$(7913)
        Case "õ": c = ChrW$(7915)
        Case "ö": c = ChrW$(7917)
        Case "÷": c = ChrW$(7919)
        Case "ù": c = ChrW$(7921)
        Case "y": c = ChrW$(121)
        Case "ý": c = ChrW$(253)
        Case "ú": c = ChrW$(7923)
        Case "û": c = ChrW$(7927)
        Case "ü": c = ChrW$(7929)
        Case "þ": c = ChrW$(7925)
        Case "®": c = ChrW$(273)
        Case "A": c = ChrW$(65)
        Case "¡": c = ChrW$(258)
        Case "¢": c = ChrW$(194)
        Case "E": c = ChrW$(69)
        Case "£": c = ChrW$(202)
        Case "O": c = ChrW$(79)
        Case "¤": c = ChrW$(212)
        Case "¥": c = ChrW$(416)
        Case "I": c = ChrW$(73)
        Case "U": c = ChrW$(85)
        Case "¦": c = ChrW$(431)
        Case "Y": c = ChrW$(89)
        Case "§": c = ChrW$(272)
        End Select
        TCVN3toUNICODE = TCVN3toUNICODE + c
    Next i
End Function



Function VNItoUNICODE(vnstr As String)
Dim c As String, i As Integer
Dim db         As Boolean
    For i = 1 To Len(vnstr)
        db = False
        If i < Len(vnstr) Then
            c = Mid(vnstr, i + 1, 1)
            If c = "ù" Or c = "ø" Or c = "û" Or c = "õ" Or c = "ï" Or _
               c = "ê" Or c = "é" Or c = "è" Or c = "ú" Or c = "ü" Or c = "ë" Or _
               c = "â" Or c = "á" Or c = "à" Or c = "å" Or c = "ã" Or c = "ä" Or _
               c = "Ù" Or c = "Ø" Or c = "Û" Or c = "Õ" Or c = "Ï" Or _
               c = "Ê" Or c = "É" Or c = "È" Or c = "Ú" Or c = "Ü" Or c = "Ë" Or _
               c = "Â" Or c = "Á" Or c = "À" Or c = "Å" Or c = "Ã" Or c = "Ä" Then db = True
        End If
        If db Then
            c = Mid(vnstr, i, 2)
            Select Case c
            Case "aù": c = ChrW$(225)
            Case "aø": c = ChrW$(224)
            Case "aû": c = ChrW$(7843)
            Case "aõ": c = ChrW$(227)
            Case "aï": c = ChrW$(7841)
            Case "aê": c = ChrW$(259)
            Case "aé": c = ChrW$(7855)
            Case "aè": c = ChrW$(7857)
            Case "aú": c = ChrW$(7859)
            Case "aü": c = ChrW$(7861)
            Case "aë": c = ChrW$(7863)
            Case "aâ": c = ChrW$(226)
            Case "aá": c = ChrW$(7845)
            Case "aà": c = ChrW$(7847)
            Case "aå": c = ChrW$(7849)
            Case "aã": c = ChrW$(7851)
            Case "aä": c = ChrW$(7853)
            Case "eù": c = ChrW$(233)
            Case "eø": c = ChrW$(232)
            Case "eû": c = ChrW$(7867)
            Case "eõ": c = ChrW$(7869)
            Case "eï": c = ChrW$(7865)
            Case "eâ": c = ChrW$(234)
            Case "eá": c = ChrW$(7871)
            Case "eà": c = ChrW$(7873)
            Case "eå": c = ChrW$(7875)
            Case "eã": c = ChrW$(7877)
            Case "eä": c = ChrW$(7879)
            Case "où": c = ChrW$(243)
            Case "oø": c = ChrW$(242)
            Case "oû": c = ChrW$(7887)
            Case "oõ": c = ChrW$(245)
            Case "oï": c = ChrW$(7885)
            Case "oâ": c = ChrW$(244)
            Case "oá": c = ChrW$(7889)
            Case "oà": c = ChrW$(7891)
            Case "oå": c = ChrW$(7893)
            Case "oã": c = ChrW$(7895)
            Case "oä": c = ChrW$(7897)
            Case "ôù": c = ChrW$(7899)
            Case "ôø": c = ChrW$(7901)
            Case "ôû": c = ChrW$(7903)
            Case "ôõ": c = ChrW$(7905)
            Case "ôï": c = ChrW$(7907)
            Case "uù": c = ChrW$(250)
            Case "uø": c = ChrW$(249)
            Case "uû": c = ChrW$(7911)
            Case "uõ": c = ChrW$(361)
            Case "uï": c = ChrW$(7909)
            Case "öù": c = ChrW$(7913)
            Case "öø": c = ChrW$(7915)
            Case "öû": c = ChrW$(7917)
            Case "öõ": c = ChrW$(7919)
            Case "öï": c = ChrW$(7921)
            Case "yù": c = ChrW$(253)
            Case "yø": c = ChrW$(7923)
            Case "yû": c = ChrW$(7927)
            Case "yõ": c = ChrW$(7929)
            Case "AÙ": c = ChrW$(193)
            Case "AØ": c = ChrW$(192)
            Case "AÛ": c = ChrW$(7842)
            Case "AÕ": c = ChrW$(195)
            Case "AÏ": c = ChrW$(7840)
            Case "AÊ": c = ChrW$(258)
            Case "AÉ": c = ChrW$(7854)
            Case "AÈ": c = ChrW$(7856)
            Case "AÚ": c = ChrW$(7858)
            Case "AÜ": c = ChrW$(7860)
            Case "AË": c = ChrW$(7862)
            Case "AÂ": c = ChrW$(194)
            Case "AÁ": c = ChrW$(7844)
            Case "AÀ": c = ChrW$(7846)
            Case "AÅ": c = ChrW$(7848)
            Case "AÃ": c = ChrW$(7850)
            Case "AÄ": c = ChrW$(7852)
            Case "EÙ": c = ChrW$(201)
            Case "EØ": c = ChrW$(200)
            Case "EÛ": c = ChrW$(7866)
            Case "EÕ": c = ChrW$(7868)
            Case "EÏ": c = ChrW$(7864)
            Case "EÂ": c = ChrW$(202)
            Case "EÁ": c = ChrW$(7870)
            Case "EÀ": c = ChrW$(7872)
            Case "EÅ": c = ChrW$(7874)
            Case "EÃ": c = ChrW$(7876)
            Case "EÄ": c = ChrW$(7878)
            Case "OÙ": c = ChrW$(211)
            Case "OØ": c = ChrW$(210)
            Case "OÛ": c = ChrW$(7886)
            Case "OÕ": c = ChrW$(213)
            Case "OÏ": c = ChrW$(7884)
            Case "OÂ": c = ChrW$(212)
            Case "OÁ": c = ChrW$(7888)
            Case "OÀ": c = ChrW$(7890)
            Case "OÅ": c = ChrW$(7892)
            Case "OÃ": c = ChrW$(7894)
            Case "OÄ": c = ChrW$(7896)
            Case "ÔÙ": c = ChrW$(7898)
            Case "ÔØ": c = ChrW$(7900)
            Case "ÔÛ": c = ChrW$(7902)
            Case "ÔÕ": c = ChrW$(7904)
            Case "ÔÏ": c = ChrW$(7906)
            Case "UÙ": c = ChrW$(218)
            Case "UØ": c = ChrW$(217)
            Case "UÛ": c = ChrW$(7910)
            Case "UÕ": c = ChrW$(360)
            Case "UÏ": c = ChrW$(7908)
            Case "ÖÙ": c = ChrW$(7912)
            Case "ÖØ": c = ChrW$(7914)
            Case "ÖÛ": c = ChrW$(7916)
            Case "ÖÕ": c = ChrW$(7918)
            Case "ÖÏ": c = ChrW$(7920)
            Case "YÙ": c = ChrW$(221)
            Case "YØ": c = ChrW$(7922)
            Case "YÛ": c = ChrW$(7926)
            Case "YÕ": c = ChrW$(7928)
            End Select
        Else
            c = Mid(vnstr, i, 1)
            Select Case c
            Case "ô": c = ChrW$(417)
            Case "í": c = ChrW$(237)
            Case "ì": c = ChrW$(236)
            Case "æ": c = ChrW$(7881)
            Case "ó": c = ChrW$(297)
            Case "ò": c = ChrW$(7883)
            Case "ö": c = ChrW$(432)
            Case "î": c = ChrW$(7925)
            Case "ñ": c = ChrW$(273)
            Case "Ô": c = ChrW$(416)
            Case "Í": c = ChrW$(205)
            Case "Ì": c = ChrW$(204)
            Case "Æ": c = ChrW$(7880)
            Case "Ó": c = ChrW$(296)
            Case "Ò": c = ChrW$(7882)
            Case "Ö": c = ChrW$(431)
            Case "Î": c = ChrW$(7924)
            Case "Ñ": c = ChrW$(272)
            End Select
        End If
        VNItoUNICODE = VNItoUNICODE + c
        If db Then i = i + 1
    Next i
End Function


Function UNICODEtoVNI(ByVal vnstr As String)
Dim c As String, i As Integer
   For i = 1 To Len(vnstr)
      c = Mid(vnstr, i, 1)
      Select Case c
         Case ChrW$(97): c = "a"
         Case ChrW$(225): c = "aù"
         Case ChrW$(224): c = "aø"
         Case ChrW$(7843): c = "aû"
         Case ChrW$(227): c = "aõ"
         Case ChrW$(7841): c = "aï"
         Case ChrW$(259): c = "aê"
         Case ChrW$(7855): c = "aé"
         Case ChrW$(7857): c = "aè"
         Case ChrW$(7859): c = "aú"
         Case ChrW$(7861): c = "aü"
         Case ChrW$(7863): c = "aë"
         Case ChrW$(226): c = "aâ"
         Case ChrW$(7845): c = "aá"
         Case ChrW$(7847): c = "aà"
         Case ChrW$(7849): c = "aå"
         Case ChrW$(7851): c = "aã"
         Case ChrW$(7853): c = "aä"
         Case ChrW$(101): c = "e"
         Case ChrW$(233): c = "eù"
         Case ChrW$(232): c = "eø"
         Case ChrW$(7867): c = "eû"
         Case ChrW$(7869): c = "eõ"
         Case ChrW$(7865): c = "eï"
         Case ChrW$(234): c = "eâ"
         Case ChrW$(7871): c = "eá"
         Case ChrW$(7873): c = "eà"
         Case ChrW$(7875): c = "eå"
         Case ChrW$(7877): c = "eã"
         Case ChrW$(7879): c = "eä"
         Case ChrW$(111): c = "o"
         Case ChrW$(243): c = "où"
         Case ChrW$(242): c = "oø"
         Case ChrW$(7887): c = "oû"
         Case ChrW$(245): c = "oõ"
         Case ChrW$(7885): c = "oï"
         Case ChrW$(244): c = "oâ"
         Case ChrW$(7889): c = "oá"
         Case ChrW$(7891): c = "oà"
         Case ChrW$(7893): c = "oå"
         Case ChrW$(7895): c = "oã"
         Case ChrW$(7897): c = "oä"
         Case ChrW$(417): c = "ô"
         Case ChrW$(7899): c = "ôù"
         Case ChrW$(7901): c = "ôø"
         Case ChrW$(7903): c = "ôû"
         Case ChrW$(7905): c = "ôõ"
         Case ChrW$(7907): c = "ôï"
         Case ChrW$(105): c = "i"
         Case ChrW$(237): c = "í"
         Case ChrW$(236): c = "ì"
         Case ChrW$(7881): c = "æ"
         Case ChrW$(297): c = "ó"
         Case ChrW$(7883): c = "ò"
         Case ChrW$(117): c = "u"
         Case ChrW$(250): c = "uù"
         Case ChrW$(249): c = "uø"
         Case ChrW$(7911): c = "uû"
         Case ChrW$(361): c = "uõ"
         Case ChrW$(7909): c = "uï"
         Case ChrW$(432): c = "ö"
         Case ChrW$(7913): c = "öù"
         Case ChrW$(7915): c = "uø"
         Case ChrW$(7917): c = "öû"
         Case ChrW$(7919): c = "öõ"
         Case ChrW$(7921): c = "öï"
         Case ChrW$(121): c = "y"
         Case ChrW$(253): c = "yù"
         Case ChrW$(7923): c = "yø"
         Case ChrW$(7927): c = "yû"
         Case ChrW$(7929): c = "yõ"
         Case ChrW$(7925): c = "î"
         Case ChrW$(273): c = "ñ"
         Case ChrW$(65): c = "A"
         Case ChrW$(193): c = "AÙ"
         Case ChrW$(192): c = "AØ"
         Case ChrW$(7842): c = "AÛ"
         Case ChrW$(195): c = "AÕ"
         Case ChrW$(7840): c = "AÏ"
         Case ChrW$(258): c = "AÊ"
         Case ChrW$(7854): c = "AÉ"
         Case ChrW$(7856): c = "AÈ"
         Case ChrW$(7858): c = "AÚ"
         Case ChrW$(7860): c = "AÜ"
         Case ChrW$(7862): c = "AË"
         Case ChrW$(194): c = "AÂ"
         Case ChrW$(7844): c = "AÁ"
         Case ChrW$(7846): c = "AÀ"
         Case ChrW$(7848): c = "AÅ"
         Case ChrW$(7850): c = "AÃ"
         Case ChrW$(7852): c = "AÄ"
         Case ChrW$(69): c = "E"
         Case ChrW$(201): c = "EÙ"
         Case ChrW$(200): c = "EØ"
         Case ChrW$(7866): c = "EÛ"
         Case ChrW$(7868): c = "EÕ"
         Case ChrW$(7864): c = "EÏ"
         Case ChrW$(202): c = "EÂ"
         Case ChrW$(7870): c = "EÁ"
         Case ChrW$(7872): c = "EÀ"
         Case ChrW$(7874): c = "EÅ"
         Case ChrW$(7876): c = "EÃ"
         Case ChrW$(7878): c = "EÄ"
         Case ChrW$(79): c = "O"
         Case ChrW$(211): c = "OÙ"
         Case ChrW$(210): c = "OØ"
         Case ChrW$(7886): c = "OÛ"
         Case ChrW$(213): c = "OÕ"
         Case ChrW$(7884): c = "OÏ"
         Case ChrW$(212): c = "OÂ"
         Case ChrW$(7888): c = "OÁ"
         Case ChrW$(7890): c = "OÀ"
         Case ChrW$(7892): c = "OÅ"
         Case ChrW$(7894): c = "OÃ"
         Case ChrW$(7896): c = "OÄ"
         Case ChrW$(416): c = "Ô"
         Case ChrW$(7898): c = "ÔÙ"
         Case ChrW$(7900): c = "ÔØ"
         Case ChrW$(7902): c = "ÔÛ"
         Case ChrW$(7904): c = "ÔÕ"
         Case ChrW$(7906): c = "ÔÏ"
         Case ChrW$(73): c = "I"
         Case ChrW$(205): c = "Í"
         Case ChrW$(204): c = "Ì"
         Case ChrW$(7880): c = "Æ"
         Case ChrW$(296): c = "Ó"
         Case ChrW$(7882): c = "Ò"
         Case ChrW$(85): c = "U"
         Case ChrW$(218): c = "UÙ"
         Case ChrW$(217): c = "UØ"
         Case ChrW$(7910): c = "UÛ"
         Case ChrW$(360): c = "UÕ"
         Case ChrW$(7908): c = "UÏ"
         Case ChrW$(431): c = "Ö"
         Case ChrW$(7912): c = "ÖÙ"
         Case ChrW$(7914): c = "ÖØ"
         Case ChrW$(7916): c = "ÖÛ"
         Case ChrW$(7918): c = "ÖÕ"
         Case ChrW$(7920): c = "ÖÏ"
         Case ChrW$(89): c = "Y"
         Case ChrW$(221): c = "YÙ"
         Case ChrW$(7922): c = "YØ"
         Case ChrW$(7926): c = "YÛ"
         Case ChrW$(7928): c = "YÕ"
         Case ChrW$(7924): c = "Î"
         Case ChrW$(272): c = "Ñ"
      End Select
      UNICODEtoVNI = UNICODEtoVNI + c
   Next i
End Function
Function UNC(strTCVN3 As String)
    UNC = TCVN3toUNICODE(strTCVN3)
End Function

Function VNI(strVNI As String)
    VNI = VNItoUNICODE(strVNI)
End Function
