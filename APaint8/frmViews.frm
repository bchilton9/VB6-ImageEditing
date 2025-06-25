VERSION 5.00
Begin VB.Form frmViews 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    View stack"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   13
      Left            =   10650
      Picture         =   "frmViews.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   12
      Left            =   9900
      Picture         =   "frmViews.frx":00D2
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   11
      Left            =   9120
      Picture         =   "frmViews.frx":01A4
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   10
      Left            =   8340
      Picture         =   "frmViews.frx":0276
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   9
      Left            =   7530
      Picture         =   "frmViews.frx":0348
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   8
      Left            =   6750
      Picture         =   "frmViews.frx":041A
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   7
      Left            =   5985
      Picture         =   "frmViews.frx":04EC
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   6
      Left            =   5205
      Picture         =   "frmViews.frx":05BE
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   5
      Left            =   4410
      Picture         =   "frmViews.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   4
      Left            =   3645
      Picture         =   "frmViews.frx":0762
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   3
      Left            =   2835
      Picture         =   "frmViews.frx":0834
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   2
      Left            =   2040
      Picture         =   "frmViews.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   1
      Left            =   1260
      Picture         =   "frmViews.frx":09D8
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVAdd 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   0
      Left            =   480
      Picture         =   "frmViews.frx":0AAA
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   " Add "
      Top             =   1125
      Width           =   345
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1665
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.CommandButton cmdVSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-->"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   870
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   " Swap "
      Top             =   1110
      Width           =   330
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   1
      Left            =   1050
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   2
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   14
      Left            =   11205
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   13
      Left            =   10425
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   12
      Left            =   9645
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   13
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   11
      Left            =   8865
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   12
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   10
      Left            =   8085
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   11
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   9
      Left            =   7305
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   10
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   8
      Left            =   6525
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   9
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   7
      Left            =   5745
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   6
      Left            =   4965
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   7
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   5
      Left            =   4185
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   6
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   4
      Left            =   3405
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   5
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   3
      Left            =   2625
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   2
      Left            =   1830
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   12450
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      Top             =   75
      Width           =   735
   End
   Begin VB.PictureBox picTV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   780
      Index           =   0
      Left            =   255
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   45
      Width           =   780
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   14
      Left            =   11715
      TabIndex        =   75
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   13
      Left            =   10965
      TabIndex        =   74
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   12
      Left            =   10155
      TabIndex        =   73
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   11
      Left            =   9405
      TabIndex        =   72
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   10
      Left            =   8640
      TabIndex        =   71
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   9
      Left            =   7845
      TabIndex        =   70
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   8
      Left            =   7035
      TabIndex        =   69
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   7
      Left            =   6270
      TabIndex        =   68
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   5475
      TabIndex        =   67
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   4725
      TabIndex        =   66
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   3945
      TabIndex        =   65
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   3165
      TabIndex        =   64
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   2355
      TabIndex        =   63
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   1545
      TabIndex        =   62
      ToolTipText     =   " Delete "
      Top             =   855
      Width           =   180
   End
   Begin VB.Label LabVX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   765
      TabIndex        =   61
      ToolTipText     =   " Delete "
      Top             =   855
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LabVClose 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C  L O  S  E   "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   1
      Left            =   12030
      TabIndex        =   60
      Top             =   60
      Width           =   210
   End
   Begin VB.Label LabVClose 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C  L O  S  E   "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   15
      TabIndex        =   59
      Top             =   45
      Width           =   210
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   14
      Left            =   11235
      TabIndex        =   30
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   13
      Left            =   10455
      TabIndex        =   29
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   9690
      TabIndex        =   28
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   11
      Left            =   8910
      TabIndex        =   27
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   10
      Left            =   8145
      TabIndex        =   26
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   7335
      TabIndex        =   25
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   6555
      TabIndex        =   24
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   5775
      TabIndex        =   23
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   4995
      TabIndex        =   22
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   4200
      TabIndex        =   21
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   3435
      TabIndex        =   20
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   2655
      TabIndex        =   19
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   1845
      TabIndex        =   18
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1065
      TabIndex        =   17
      ToolTipText     =   " Show "
      Top             =   855
      Width           =   360
   End
   Begin VB.Label LabVN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   285
      TabIndex        =   16
      ToolTipText     =   " Show "
      Top             =   840
      Width           =   360
   End
End
Attribute VB_Name = "frmViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmViews.frm by Robert Rayment

Option Explicit
'  Windows API to make form stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H2

Private Sub Form_Load()
Dim k As Long
   frmViews.Left = frmViewsLeft
   frmViews.Top = frmViewsTop
   aVIEWS = True
   
   ' Size & Make frmZoom stay on top
   k = SetWindowPos(frmViews.hwnd, hWndInsertAfter, frmViewsLeft, frmViewsTop, _
   12360 \ STX, 1740 \ STY, wFlags)

   For k = 0 To 14
      picTV(k).Width = 52
      picTV(k).Height = 52
   Next k
   
   DISPLAY_ALL_VIEWS
   
'   picTemp.Width = 8
'   picTemp.Height = 8
'   picTemp.Picture = LoadPicture
End Sub


Private Sub LabVClose_Click(Index As Integer)
   frmViewsLeft = frmViews.Left
   frmViewsTop = frmViews.Top
   aVIEWS = False
   Unload frmViews
End Sub

Private Sub LabVN_Click(Index As Integer)
' Show pic UndoNum
If ADRAW Then Exit Sub
   If Index + 1 <= TopUndoNum Then
      UndoNum = Index + 1
      Form1.CommonUndo
   End If
End Sub

Private Sub cmdVSwap_Click(Index As Integer)
' Swap with view below
' Index = 0 Swap 1 & 2
' Index = 1 Swap 2 & 3
' etc
If ADRAW Then Exit Sub
   If Index + 2 > 1 And Index + 2 <= TopUndoNum Then
      UndoNum = Index + 2
      Form1.CommonUndo
      Form1.CommonSwap
   End If
End Sub

Private Sub cmdVAdd_Click(Index As Integer)
' Add view above to current view
' Index = 0    1 <-+ 2
' Index = 1    2 <-+ 3
' etc
If ADRAW Then Exit Sub
   If Index + 1 < TopUndoNum Then
      UndoNum = Index + 1
      Form1.CommonUndo
      Form1.CommonAdd
   End If
End Sub

Private Sub LabVX_Click(Index As Integer)
' Index = 1    Del 2
' Index = 2    Del 3
If ADRAW Then Exit Sub
   If Index + 1 <= TopUndoNum Then
      UndoNum = Index + 1
      Form1.CommonUndo
      Form1.CommonDelete
   End If
End Sub
