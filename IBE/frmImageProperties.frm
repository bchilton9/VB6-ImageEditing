VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmImageProperties 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Properties"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   Icon            =   "frmImageProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer wfm 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
      PictureLeft     =   "frmImageProperties.frx":0ECA
      PictureMiddle   =   "frmImageProperties.frx":1934
      PictureRight    =   "frmImageProperties.frx":19D2
      PictureRightWidth=   102
      FormBorderTop   =   "frmImageProperties.frx":1A70
      FormBorderLeft  =   "frmImageProperties.frx":1AD2
      FormBorderBottom=   "frmImageProperties.frx":1B30
      FormBorderRight =   "frmImageProperties.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmImageProperties.frx":1BF0
      AllowMaximize   =   0   'False
      FormIcon        =   "frmImageProperties.frx":2842
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
      PictureMaximize =   "frmImageProperties.frx":371C
      PictureMinimize =   "frmImageProperties.frx":3AAE
      PictureClose    =   "frmImageProperties.frx":3E40
      PictureMinimizeToTray=   "frmImageProperties.frx":41D2
      PictureShrink   =   "frmImageProperties.frx":4564
      PictureCloseDown=   "frmImageProperties.frx":48F6
      PictureMaximizeDown=   "frmImageProperties.frx":4C88
      PictureMinimizeDown=   "frmImageProperties.frx":501A
      PictureShrinkDown=   "frmImageProperties.frx":53AC
      PictureMinimizeToTrayDown=   "frmImageProperties.frx":573E
      PicturePin      =   "frmImageProperties.frx":5AD0
      PicturePinDown  =   "frmImageProperties.frx":5E62
      PicturePinHover =   "frmImageProperties.frx":61F4
      PictureMinimizeToTrayHover=   "frmImageProperties.frx":6546
      PictureShrinkHover=   "frmImageProperties.frx":6898
      PictureMinimizeHover=   "frmImageProperties.frx":6BEA
      PictureMaximizeHover=   "frmImageProperties.frx":6F3C
      PictureCloseHover=   "frmImageProperties.frx":728E
      TrayTip         =   " Properties "
      FormMouseIcon   =   "frmImageProperties.frx":75E0
      TrayIcon        =   "frmImageProperties.frx":7DFA
   End
   Begin prjImageBrowser.ctlImagePreview ctlImagePreview 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5318
      BackColor       =   16777215
   End
   Begin VB.TextBox txtSize 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1515
      Width           =   1935
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1875
      Width           =   4095
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1155
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1155
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   795
      Width           =   4095
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   435
      Width           =   4095
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   195
      Left            =   6480
      TabIndex        =   2
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      Height          =   195
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   330
   End
End
Attribute VB_Name = "frmImageProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowProperties(picPicture As StdPicture, Optional FileName As String)
    Set ctlImagePreview.Picture = picPicture
    txtPath.Text = GetPathOnlyFromFullPath(FileName)
    txtFileName.Text = GetFileNameOnlyFromFullPath(FileName)
    txtWidth.Text = Int(picPicture.Width / K_DotsPerPixel)
    txtHeight.Text = Int(picPicture.Height / K_DotsPerPixel)
    
    Me.Show 'vbModal
End Sub

