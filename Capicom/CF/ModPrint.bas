Attribute VB_Name = "ModPrint"

Rem < Modulo di Stampa
' Set Picture1.Picture = CaptureForm(Me)
' Set Picture1.Picture = CaptureClient(Me)
' Set Picture1.Picture = CaptureActiveWindow()
' Set Picture1.Picture = CaptureScreen()
'<!

Option Explicit
Option Base 0

Private Type PALETTEENTRY
peRed As Byte
peGreen As Byte
peBlue As Byte
peFlags As Byte
End Type

Private Type LOGPALETTE
palVersion As Integer
palNumEntries As Integer
palPalEntry(255) As PALETTEENTRY ' 256 colori
End Type

Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

#If Win32 Then

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
ByVal hDC As Long, ByVal nWidth As Long, _
ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" ( _
ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
ByVal hDC As Long, ByVal wStartIndex As Long, _
ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
As Long
Private Declare Function CreatePalette Lib "GDI32" ( _
lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" ( _
ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" ( _
ByVal hDCDest As Long, ByVal XDest As Long, _
ByVal YDest As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hDCSrc As Long, _
ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
As Long
Private Declare Function DeleteDC Lib "GDI32" ( _
ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () _
As Long
Private Declare Function SelectPalette Lib "GDI32" ( _
ByVal hDC As Long, ByVal hPalette As Long, _
ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" ( _
ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" ( _
ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "USER32" ( _
ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" ( _
ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32" ( _
ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Type PicBmp
Size As Long
Type As Long
hBmp As Long
hPal As Long
Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect _
Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

#ElseIf Win16 Then

Private Const RASTERCAPS As Integer = 38
Private Const RC_PALETTE As Integer = &H100
Private Const SIZEPALETTE As Integer = 104

Private Type RECT
Left As Integer
Top As Integer
Right As Integer
Bottom As Integer
End Type

Private Declare Function CreateCompatibleDC Lib "GDI" ( _
ByVal hDC As Integer) As Integer
Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
ByVal hDC As Integer, ByVal nWidth As Integer, _
ByVal nHeight As Integer) As Integer
Private Declare Function GetDeviceCaps Lib "GDI" ( _
ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer
Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
ByVal hDC As Integer, ByVal wStartIndex As Integer, _
ByVal wNumEntries As Integer, _
lpPaletteEntries As PALETTEENTRY) As Integer
Private Declare Function CreatePalette Lib "GDI" ( _
lpLogPalette As LOGPALETTE) As Integer
Private Declare Function SelectObject Lib "GDI" ( _
ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Private Declare Function BitBlt Lib "GDI" ( _
ByVal hDCDest As Integer, ByVal XDest As Integer, _
ByVal YDest As Integer, ByVal nWidth As Integer, _
ByVal nHeight As Integer, ByVal hDCSrc As Integer, _
ByVal XSrc As Integer, ByVal YSrc As Integer, _
ByVal dwRop As Long) As Integer
Private Declare Function DeleteDC Lib "GDI" ( _
ByVal hDC As Integer) As Integer
Private Declare Function GetForegroundWindow Lib "USER" _
Alias "GetActiveWindow" () As Integer
Private Declare Function SelectPalette Lib "USER" ( _
ByVal hDC As Integer, ByVal hPalette As Integer, ByVal _
bForceBackground As Integer) As Integer
Private Declare Function RealizePalette Lib "USER" ( _
ByVal hDC As Integer) As Integer
Private Declare Function GetWindowDC Lib "USER" ( _
ByVal hWnd As Integer) As Integer
Private Declare Function GetDC Lib "USER" ( _
ByVal hWnd As Integer) As Integer
Private Declare Function GetWindowRect Lib "USER" ( _
ByVal hWnd As Integer, lpRect As RECT) As Integer
Private Declare Function ReleaseDC Lib "USER" ( _
ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
Private Declare Function GetDesktopWindow Lib "USER" () As Integer

Private Type PicBmp
Size As Integer
Type As Integer
hBmp As Integer
hPal As Integer
Reserved As Integer
End Type

Private Declare Function OleCreatePictureIndirect _
Lib "oc25.dll" (PictDesc As PicBmp, RefIID As GUID, _
ByVal fPictureOwnsHandle As Integer, IPic As IPicture) _
As Integer

#End If

#If Win32 Then
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
Dim r As Long
#ElseIf Win16 Then
Public Function CreateBitmapPicture(ByVal hBmp As Integer, ByVal hPal As Integer) As Picture

Dim r As Integer
#End If
Dim Pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID

With IID_IDispatch
.Data1 = &H20400
.Data4(0) = &HC0
.Data4(7) = &H46
End With

With Pic
.Size = Len(Pic)
.Type = vbPicTypeBitmap
.hBmp = hBmp
.hPal = hPal
End With

r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
Set CreateBitmapPicture = IPic
End Function



#If Win32 Then
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
#ElseIf Win16 Then
Public Function CaptureWindow(ByVal hWndSrc As Integer, ByVal Client As Boolean, ByVal LeftSrc As Integer, ByVal TopSrc As Integer, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Integer
Dim hBmp As Integer
Dim hBmpPrev As Integer
Dim r As Integer
Dim hDCSrc As Integer
Dim hPal As Integer
Dim hPalPrev As Integer
Dim RasterCapsScrn As Integer
Dim HasPaletteScrn As Integer
Dim PaletteSizeScrn As Integer
#End If
Dim LogPal As LOGPALETTE
If Client Then
hDCSrc = GetDC(hWndSrc)
Else
hDCSrc = GetWindowDC(hWndSrc)
End If
hDCMemory = CreateCompatibleDC(hDCSrc)
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
HasPaletteScrn = RasterCapsScrn And RC_PALETTE
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
LogPal.palVersion = &H300
LogPal.palNumEntries = 256
r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
hPal = CreatePalette(LogPal)
hPalPrev = SelectPalette(hDCMemory, hPal, 0)
r = RealizePalette(hDCMemory)
End If
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
hBmp = SelectObject(hDCMemory, hBmpPrev)
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function


Public Function CaptureScreen() As Picture
#If Win32 Then
Dim hWndScreen As Long
#ElseIf Win16 Then
Dim hWndScreen As Integer
#End If
hWndScreen = GetDesktopWindow()
Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

Public Function CaptureForm(frmSrc As Form) As Picture
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
    frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

Public Function CaptureClient(frmSrc As Form) As Picture
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, _
    frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function

Public Function CaptureActiveWindow() As Picture
#If Win32 Then
Dim hWndActive As Long
Dim r As Long
#ElseIf Win16 Then
Dim hWndActive As Integer
Dim r As Integer
#End If
Dim RectActive As RECT
hWndActive = GetForegroundWindow()
r = GetWindowRect(hWndActive, RectActive)
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
    RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function

Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double

If Pic.Height >= Pic.Width Then
Prn.Orientation = vbPRORPortrait
Else
Prn.Orientation = vbPRORLandscape
End If

PicRatio = Pic.Width / Pic.Height

PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
PrnRatio = PrnWidth / PrnHeight

If PicRatio >= PrnRatio Then
PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
Else
PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
End If

Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub
