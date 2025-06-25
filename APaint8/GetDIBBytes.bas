Attribute VB_Name = "GetDIBBytes"
' GetDIBBytes.bas  By  Robert Rayment

Option Explicit
Option Base 1


' Public BITMAPINFOHEADER

Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Const DIB_PAL_COLORS = 1 '  system colors
' -----------------------------------------------------------
Private Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
(ByVal hDC As Long) As Long
'----------------------------------------------------------------

Public Sub GETBYTES(ByVal PICINP As Long, _
   bA() As Byte, bWidth As Long, bHeight As Long, bppNum As Integer, cPAL() As Long, FileKind As Long)

' In: PICINP picBox.Picture, bA(1,1) bppNum=8, bA(1,1,1) bppNum=32
'     cPal(0 To 255)  R + 256*G + 65536*B   A=0
'     FileKind 0=BMP, 1=GIF  affects sign of bS.bmi.biheight

Dim NewDC As Long
Dim OldH As Long
Dim bS As BITMAPINFO
Dim BytesPerScanLine As Long
Dim k As Long

   If bppNum <> 8 And bppNum <> 32 Then
      MsgBox "Wrong bpp GETBYTES", vbCritical, " Painter8bpp"
      Exit Sub
   End If
   If FileKind <> 0 And FileKind <> 1 Then
      MsgBox "Wrong FileKind GETBYTES", vbCritical, " Painter8bpp"
      Exit Sub
   End If
   
   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PICINP)
   
   If bppNum = 32 Then
      BytesPerScanLine = bWidth * 4&
   Else  ' bppNum=8
      BytesPerScanLine = (bWidth + 3) And &HFFFFFFFC
      ' Set up palette
      CopyMemory bS.Colors(0), cPAL(0), 256
   End If
   
   With bS.bmi
      .biSize = 40
      .biwidth = bWidth
      If FileKind = 0 Then
         .biheight = bHeight  ' As require by BMP
      Else
         .biheight = -bHeight ' As require by GIF
      End If
      .biPlanes = 1
      .biBitCount = bppNum    ' Sets up 8 or 32-bit colors
      .biCompression = 0
      .biSizeImage = BytesPerScanLine * Abs(bHeight)
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   If bppNum = 32 Then
      If GetDIBits(NewDC, PICINP, 0, bHeight, bA(1, 1, 1), bS, DIB_PAL_COLORS) = 0 Then
         MsgBox "DIB Error in GETBYTES 32bpp", vbCritical, " Painter8bpp"
         End
      End If
   Else  ' bppNum = 8
      If GetDIBits(NewDC, PICINP, 0, bHeight, bA(1, 1), bS, DIB_RGB_COLORS) = 0 Then
         MsgBox "DIB Error in GETBYTES 8 bpp", vbCritical, " Painter8bpp"
         End
      End If
   End If
   
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC

End Sub

Public Sub GETLONGS(ByVal PICINP As Long, _
   LA() As Long, bWidth As Long, bHeight As Long)

' Used by ZOOMER
Dim NewDC As Long
Dim OldH As Long
Dim bS As BITMAPINFO

   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PICINP)
   With bS.bmi
      .biSize = 40
      .biwidth = bWidth    ' Always multiple of 4 !!
      .biheight = bHeight
      .biPlanes = 1
      .biBitCount = 32     ' 32-bit colors
      .biCompression = 0
      .biSizeImage = 4 * bWidth * bHeight
   End With
   
   If GetDIBits(NewDC, PICINP, 0, bHeight, LA(1, 1), bS, DIB_PAL_COLORS) = 0 Then
      MsgBox "DIB Error in GETLONGS 32bpp", vbCritical, " ZOOMER"
      End
   End If
   
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC
End Sub

Public Sub GetPICBytes(ByVal PICINP As Long, bA() As Byte, LW As Long, LH As Long)
' Used to get mask in Form1 & Text in frmText
Dim NewDC As Long
Dim OldH As Long
Dim bS As BITMAPINFO
Dim BytesPerScanLine As Long
Dim k As Long
   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PICINP)
'  Not nec since width is already multiple of 4
'  BytesPerScanLine = (bWidth + 3) And &HFFFFFFFC
   
   CopyMemory bS.Colors(0), CulRGB(0), 1024
   
   With bS.bmi
      .biSize = 40
      .biwidth = LW
      .biheight = LH ' As require for text scanning
      .biPlanes = 1
      .biBitCount = 8
      .biCompression = 0
      .biSizeImage = LW * Abs(LH) 'BytesPerScanLine * Abs(LH)
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   If GetDIBits(NewDC, PICINP, 0, LH, bA(1, 1), bS, DIB_PAL_COLORS) = 0 Then
      MsgBox "DIB Error in GETPICBytes 8 bpp", vbCritical, " Get text & Ellipse"
      End
   End If
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC
End Sub

