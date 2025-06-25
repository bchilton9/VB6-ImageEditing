Attribute VB_Name = "API"
' API.bas

Option Explicit
Option Base 1

Public Type POINTAPI
        kX As Long
        kY As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public iREC As RECT
'ar = SetRect(IREC, 0, 0, picWidth - 1, picHeight - 1)
'ar = InvertRect(PIC(1).hdc, IREC)
'------------------------------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Public RGBS As RGBQUAD

Public Type BITMAPINFO
   bmi As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const PM_REMOVE = &H1

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" _
(lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

' To invert picture box
Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function InvertRect Lib "user32" _
(ByVal hDC As Long, lpRect As RECT) As Long

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
 ByVal y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1

Public Declare Function SetPixelV Lib "gdi32" _
(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
(Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

'API for postioning mouse
Public Declare Sub SetCursorPos Lib "user32" (ByVal ix As Long, ByVal iy As Long)
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
    lpPoint As POINTAPI) As Long

Public Declare Function SetDIBitsToDevice Lib "gdi32" _
(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal Scan As Long, ByVal NumScans As Long, _
Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
' EG
'   SetDIBitsToDevice des.hdc, 0, 0, desWidth, desHeight, _
'   xs, ys, 0, BArrayHeight, BArray(1, 1), bArr, DIB_RGB_COLORS
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

' For transferring drawing in an integer array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, _
ByVal x As Long, ByVal y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'eg StretchDIBits PIC.hDC, 0&, 0&, W4, H4, 0&, 0&, W, H, b8(1, 1), BS, DIB_RGB_COLORS, vbSrcCopy
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'StretchBlt Dhdc,xd,yd,dw,dh,Shdc,xs,ys,sw,sh,vbSrcCopy

Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long
'-----------------------------------------------------------------
