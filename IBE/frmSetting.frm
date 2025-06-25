VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmSetting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer wfm 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   582
      PictureLeft     =   "frmSetting.frx":0ECA
      PictureMiddle   =   "frmSetting.frx":1934
      PictureRight    =   "frmSetting.frx":19D2
      PictureRightWidth=   46
      FormBorderTop   =   "frmSetting.frx":1A70
      FormBorderLeft  =   "frmSetting.frx":1AD2
      FormBorderBottom=   "frmSetting.frx":1B30
      FormBorderRight =   "frmSetting.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmSetting.frx":1BF0
      AllowMaximize   =   0   'False
      FormIcon        =   "frmSetting.frx":2842
      AllowMinimize   =   0   'False
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
      PictureMaximize =   "frmSetting.frx":371C
      PictureMinimize =   "frmSetting.frx":3AAE
      PictureClose    =   "frmSetting.frx":3E40
      PictureMinimizeToTray=   "frmSetting.frx":41D2
      PictureShrink   =   "frmSetting.frx":4564
      AllowShrink     =   0   'False
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmSetting.frx":48F6
      PictureMaximizeDown=   "frmSetting.frx":4C88
      PictureMinimizeDown=   "frmSetting.frx":501A
      PictureShrinkDown=   "frmSetting.frx":53AC
      PictureMinimizeToTrayDown=   "frmSetting.frx":573E
      PicturePin      =   "frmSetting.frx":5AD0
      PicturePinDown  =   "frmSetting.frx":5E62
      PicturePinHover =   "frmSetting.frx":61F4
      PictureMinimizeToTrayHover=   "frmSetting.frx":6546
      PictureShrinkHover=   "frmSetting.frx":6898
      PictureMinimizeHover=   "frmSetting.frx":6BEA
      PictureMaximizeHover=   "frmSetting.frx":6F3C
      PictureCloseHover=   "frmSetting.frx":728E
      TrayTip         =   " Settings "
      FormMouseIcon   =   "frmSetting.frx":75E0
      TrayIcon        =   "frmSetting.frx":7DFA
   End
   Begin prjImageBrowser.chameleonButton cmdCancel 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "frmSetting.frx":8CD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.chameleonButton cmdOk 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Okay"
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
      MICON           =   "frmSetting.frx":8CF0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjImageBrowser.ctlFileBrowser txtExternalApplicationPrinter 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      Caption         =   " External printer "
      ButtonCaption   =   "..."
      Filter          =   "Application|*.EXE"
      BackColor       =   16761024
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FileFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjImageBrowser.ctlFileBrowser txtExternalApplicationEditor 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      Caption         =   " External editor  "
      ButtonCaption   =   "..."
      Filter          =   "Application|*.EXE"
      BackColor       =   16761024
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FileFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjImageBrowser.ctlFileBrowser txtExternalApplicationViewer 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      Caption         =   " External viewer"
      ButtonCaption   =   "..."
      Filter          =   "Application|*.EXE"
      BackColor       =   16761024
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FileFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblExternalApplicationTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Use %FILENAME% to replace the command line with the filename."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4800
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim New_Setting As SettingType
    With New_Setting
        .ExternalApplication.Viewer = txtExternalApplicationViewer.File
        .ExternalApplication.Editor = txtExternalApplicationEditor.File
        .ExternalApplication.Printer = txtExternalApplicationPrinter.File
        
        SettingSave New_Setting
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    With ApplicationSetting
        txtExternalApplicationViewer.File = .ExternalApplication.Viewer
        txtExternalApplicationEditor.File = .ExternalApplication.Editor
        txtExternalApplicationPrinter.File = .ExternalApplication.Printer
    End With
End Sub
