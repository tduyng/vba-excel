Option Explicit
 
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const VK_SNAPSHOT = 44
Const VK_LMENU = 164
Const KEYEVENTF_KEYUP = 2
Const KEYEVENTF_EXTENDEDKEY = 1

'Adapted from Kenneth Hobs
'at http://www.ozgrid.com/forum/showthread.php?t=157677

Private Sub CommandButton1_Click()
'change to your button name
    Dim pdfName As String
    Dim newWS As Worksheet
     
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
     
    DoEvents 'Otherwise, all of screen would be pasted as if PrtScn rather than Alt+PrtScn was used for the copy.
     
    Set newWS = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newWS.PasteSpecial Format:="Bitmap", Link:=False, DisplayAsIcon:=False
    pdfName = ActiveWorkbook.Path & "\" & Me.Name & " " & Format(Now, "yyyy-mmm-dd") & ".pdf"
    newWS.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    newWS.Delete
    Unload Me
End Sub