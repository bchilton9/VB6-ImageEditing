VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmPrintImage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Print"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   Icon            =   "frmPrintImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer wfm 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   582
      PictureLeft     =   "frmPrintImage.frx":0ECA
      PictureMiddle   =   "frmPrintImage.frx":1934
      PictureRight    =   "frmPrintImage.frx":19D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmPrintImage.frx":1A70
      FormBorderLeft  =   "frmPrintImage.frx":1AD2
      FormBorderBottom=   "frmPrintImage.frx":1B30
      FormBorderRight =   "frmPrintImage.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmPrintImage.frx":1BF0
      AllowMaximize   =   0   'False
      FormIcon        =   "frmPrintImage.frx":2842
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   8388608
      PictureMaximize =   "frmPrintImage.frx":371C
      PictureMinimize =   "frmPrintImage.frx":3AAE
      PictureClose    =   "frmPrintImage.frx":3E40
      PictureMinimizeToTray=   "frmPrintImage.frx":41D2
      PictureShrink   =   "frmPrintImage.frx":4564
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmPrintImage.frx":48F6
      PictureMaximizeDown=   "frmPrintImage.frx":4C88
      PictureMinimizeDown=   "frmPrintImage.frx":501A
      PictureShrinkDown=   "frmPrintImage.frx":53AC
      PictureMinimizeToTrayDown=   "frmPrintImage.frx":573E
      PicturePin      =   "frmPrintImage.frx":5AD0
      PicturePinDown  =   "frmPrintImage.frx":5E62
      PicturePinHover =   "frmPrintImage.frx":61F4
      PictureMinimizeToTrayHover=   "frmPrintImage.frx":6546
      PictureShrinkHover=   "frmPrintImage.frx":6898
      PictureMinimizeHover=   "frmPrintImage.frx":6BEA
      PictureMaximizeHover=   "frmPrintImage.frx":6F3C
      PictureCloseHover=   "frmPrintImage.frx":728E
      TrayTip         =   " Print "
      FormMouseIcon   =   "frmPrintImage.frx":75E0
      TrayIcon        =   "frmPrintImage.frx":7DFA
   End
   Begin prjImageBrowser.chameleonButton cmdFitReset 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Reset"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPrintImage.frx":8CD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.chameleonButton cmdFitCenter 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Center"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPrintImage.frx":8CF0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPrintImage.frx":8D0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.chameleonButton cmdPrint 
      Default         =   -1  'True
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPrintImage.frx":8D28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.chameleonButton cmdFitBest 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Best fit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPrintImage.frx":8D44
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.ctlImagePreview ctlImagePreview 
      Height          =   5415
      Left            =   5880
      TabIndex        =   1
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9551
   End
   Begin prjImageBrowser.ctlImagePrinter ctlImagePrinter 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12726
      ColorMode       =   1
      ImageHeight     =   5
      ImageWidth      =   5
      BackColor       =   16761024
   End
End
Attribute VB_Name = "frmPrintImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowPrint(picImage As StdPicture)
    Set ctlImagePreview.Picture = picImage
    Set ctlImagePrinter.Picture = picImage
    ctlImagePrinter.ImageWidth = Round(picImage.Width / 1440, 2)
    ctlImagePrinter.ImageHeight = Round(picImage.Height / 1440, 2)
    Me.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFitBest_Click()
    ctlImagePrinter.FitBest
End Sub

Private Sub cmdFitCenter_Click()
    ctlImagePrinter.FitCenter
End Sub

Private Sub cmdFitReset_Click()
    ctlImagePrinter.FitReset
End Sub

Private Sub cmdPrint_Click()
    ctlImagePrinter.PrintImage
End Sub
