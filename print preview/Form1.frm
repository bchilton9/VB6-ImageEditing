VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   9135
      Left            =   240
      ScaleHeight     =   9075
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   720
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   120
      ScaleHeight     =   9075
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Preview first"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()
   ' Setup (Could be done at design time or in form load)
   ' Make printing stick:
   Picture1.AutoRedraw = True
   ' Add a palette for 256 colors:
   Picture1.Picture = LoadPicture(App.Path & "\PASTEL.DIB")
   ' Set up hidden picture:
   Picture2.AutoRedraw = False
   Picture2.ScaleMode = 3 'Pixels
   Picture2.Visible = False
   Picture2.AutoSize = True
   Picture2.Picture = LoadPicture(App.Path & "\PRINTER.WMF")

   ' This print job can go to the printer or the picture box:
   If Check1.Value = 0 Then PrinterFlag = True
   PrintStartDoc Picture1, PrinterFlag, 8.5, 11

   ' All the subs use inches:
   PrintBox 1, 1, 6.5, 9
   PrintLine 1.1, 2, 7.4, 2
   PrintPicture Picture2, 1.1, 1.1, 0.8, 0.8
   PrintFilledBox 2.1, 1.2, 5.2, 0.7, RGB(200, 200, 200)
   PrintFontName "Arial"
   PrintCurrentX 2.3
   PrintCurrentY 1.3
   PrintFontSize 35
   PrintPrint "Visual Basic Printing"
   For x = 3 To 5.5 Step 0.2
      PrintCircle x, 3.5, 0.75
   Next
   PrintFontName "Courier New"
   PrintFontSize 30
   PrintCurrentX 1.5
   PrintCurrentY 5
   PrintPrint "It is possible to do"
   PrintFontSize 24
   PrintCurrentX 1.5
   PrintCurrentY 6.5
   PrintPrint "It is possible to do print"
   PrintFontSize 18
   PrintCurrentX 1.5
   PrintCurrentY 8
   PrintPrint "It is possible to do print preview"
   PrintFontSize 12
   PrintCurrentX 1.5
   PrintCurrentY 9.5
   PrintPrint "It is possible to do print preview with good results."
   PrintEndDoc
End Sub
