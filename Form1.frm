VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintPicture 
      Caption         =   "&Print Picture"
      Default         =   -1  'True
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SRCCOPY = &HCC0020
Private Declare Function Escape Lib "gdi32" (ByVal hdc As Long, _
     ByVal nEscape As Long, ByVal nCount As Long, lpInData As Any, _
     lpOutData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
     ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
     ByVal ySrc As Long, ByVal nSrcWidth As Long, _
     ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
     ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" _
     (ByVal hdc As Long) As Long

Private Sub cmdPrintPicture_Click()
Dim hMemoryDC As Long
Dim hOldBitMap As Long
Dim x As Long
Const NEWFRAME = 1

    Picture1.Picture = Picture1.Image
    '* StretchBlt requires pixel coordinates.
    
    Picture1.ScaleMode = vbPixels
    Printer.ScaleMode = vbPixels
    Printer.Print ""; ' init printer object
    
    hMemoryDC = CreateCompatibleDC(Picture1.hdc)
    hOldBitMap = SelectObject(hMemoryDC, Picture1.Picture)
    x = StretchBlt(Printer.hdc, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight, _
         hMemoryDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, SRCCOPY)
    hOldBitMap = SelectObject(hMemoryDC, hOldBitMap)
    x = DeleteDC(hMemoryDC)
    x = Escape(Printer.hdc, NEWFRAME, 0, 0&, 0&)
    
    Printer.EndDoc

End Sub

