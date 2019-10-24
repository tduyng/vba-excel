Attribute VB_Name = "z_ScreenShot"
Option Explicit

'Capture d'écran
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef PicDesc As PicBmp, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32.dll" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, ByRef lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32.dll" (ByRef lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32.dll" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Const SM_CXSCREEN = 0&
Private Const SM_CYSCREEN = 1&
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Sub prcSave_Picture_Screen(Optional fullNameScreenShot As String = vbNullString) 'ganzer bildschirm
    Dim fso As Object
    Dim pathScreenShot As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Créer un dossier qui s'appelle Captures d'écran
    pathScreenShot = ThisWorkbook.Path & "\" & "Captures d'écran"
    If Not fso.FolderExists(pathScreenShot) Then
        fso.CreateFolder (pathScreenShot)
    End If
    

        
    If fullNameScreenShot = vbNullString Then
    
        fullNameScreenShot = pathScreenShot & "\Screenshot_" & Format(Now(), "yyyymmdd-hmmss") & ".bmp"
        
        stdole.SavePicture hDCToPicture(GetDC(0&), 0&, 0&, _
        GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)), fullNameScreenShot 'anpassen !!!
        
    Else
        stdole.SavePicture hDCToPicture(GetDC(0&), 0&, 0&, _
        GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)), _
        pathScreenShot & "\Screenshot" & Format(Now(), "yyyymmdd_hmmss") & ".bmp" 'anpassen !!!
    End If

End Sub

Public Sub prcSave_Picture_Active_Window(Optional fullNameScreenShot As String = vbNullString, Optional timeWait As Double = 0) 'aktives Fenster
    Dim hWnd As Long
    Dim udtRect As RECT
    Dim fso As Object
    Dim pathScreenShot As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Créer un dossier qui s'appelle Captures d'écran
    pathScreenShot = ThisWorkbook.Path & "\" & "Captures d'écran"
    If Not fso.FolderExists(pathScreenShot) Then
        fso.CreateFolder (pathScreenShot)
    End If
    
    Sleep timeWait 'Le temps en miliseconde attente pour capturer écran
    
    hWnd = GetForegroundWindow
    GetWindowRect hWnd, udtRect
    If fullNameScreenShot = vbNullString Then
    
        fullNameScreenShot = pathScreenShot & "\Screenshot_" & Format(Now(), "yyyymmdd-hmmss") & ".bmp"
        
        stdole.SavePicture hDCToPicture(GetDC(0&), udtRect.Left, udtRect.Top, _
        udtRect.Right - udtRect.Left, udtRect.Bottom - udtRect.Top), fullNameScreenShot 'anpassen !!!
        
    Else
        stdole.SavePicture hDCToPicture(GetDC(0&), udtRect.Left, udtRect.Top, _
        udtRect.Right - udtRect.Left, udtRect.Bottom - udtRect.Top), _
        pathScreenShot & "\Screenshot_" & Format(Now(), "yyyymmdd-hmmss") & ".bmp" 'anpassen !!!
    End If

End Sub

Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Object
    Dim Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        .Size = Len(Pic)
        .Type = 1
        .hBmp = hBmp
        .hPal = hPal
    End With
    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function

Private Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Object
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE
    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        Call GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        Call RealizePalette(hDCMemory)
    End If
    Call BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, 13369376)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    Call DeleteDC(hDCMemory)
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function

