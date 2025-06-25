Attribute VB_Name = "BMPSave"
'BMPSave.bas  by  Robert Rayment

' For saving 8bpp BMPs only

Option Explicit
Option Base 1

Private Type BITMAPFILEHEADER   ' For 8bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+1024+WxH  4
    bReserved1        As Integer  '           2
    bReserved2        As Integer  '           2
    bOffBits          As Long     ' 1078      4 =54+1024
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' W         4
    bHeight           As Long     ' H         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 8         2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' WxH       4
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 54
End Type


Public Sub MSaveBMP(FSpec$, b8A() As Byte, bWidth As Long, bHeight As Long, cPAL() As Long)
'     b8A(1-bWidth,1-bHeight) As byte from GETDIBBytes
'     cPal() 0-255 B + 256*G + 65536*R   A=0  Len=4x256=1024

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Long
Dim BytesPerScanLine As Long
      
   BytesPerScanLine = (bWidth + 3) And &HFFFFFFFC

   With BFH
      .bType = &H4D42    ' BM
      .bSize = 54 + 1024 + bWidth * bHeight
      .bOffBits = 1078
      .bHeaderSize = 40
      .bWidth = bWidth
      .bHeight = bHeight
      .bNumPlanes = 1
      .bBPP = 8
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   
   '-- Kill previous
   On Error Resume Next
   Kill FSpec$
   On Error GoTo 0
   
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , cPAL()
   Put #fnum, , b8A()
   Close #fnum
   
End Sub



