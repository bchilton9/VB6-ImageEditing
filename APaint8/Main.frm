VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   Caption         =   "Paint8"
   ClientHeight    =   7755
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10365
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0E42
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHairs 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hairs"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4665
      TabIndex        =   135
      Top             =   270
      Value           =   1  'Checked
      Width           =   630
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   133
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   132
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   130
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   129
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   128
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.CommandButton cmdTile 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   127
      ToolTipText     =   " Tile form "
      Top             =   6735
      Width           =   105
   End
   Begin VB.Frame fraSmoothers 
      BackColor       =   &H0080C0FF&
      Caption         =   "Smooth"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   45
      TabIndex        =   122
      Top             =   5715
      Width           =   780
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   40
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   " Large area "
         Top             =   510
         Width           =   330
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   39
         Left            =   435
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   " Medium area "
         Top             =   180
         Width           =   240
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   38
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   " Small area "
         Top             =   210
         Width           =   180
      End
   End
   Begin VB.Frame fraRoller 
      BackColor       =   &H0080C0FF&
      Caption         =   "Roll/Shift"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   45
      TabIndex        =   108
      Top             =   3855
      Width           =   780
      Begin VB.OptionButton optRollShift 
         BackColor       =   &H0080C0FF&
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   119
         Top             =   1545
         Width           =   660
      End
      Begin VB.OptionButton optRollShift 
         BackColor       =   &H0080C0FF&
         Caption         =   "Roll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   118
         Top             =   1335
         Width           =   660
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   7
         Left            =   435
         Picture         =   "Main.frx":12C4
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   1020
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   6
         Left            =   435
         Picture         =   "Main.frx":1396
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   780
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   5
         Left            =   105
         Picture         =   "Main.frx":1468
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   1020
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   4
         Left            =   105
         Picture         =   "Main.frx":153A
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   780
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   3
         Left            =   405
         Picture         =   "Main.frx":160C
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   480
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   2
         Left            =   150
         Picture         =   "Main.frx":16DE
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   480
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   1
         Left            =   405
         Picture         =   "Main.frx":17B0
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   210
         Width           =   244
      End
      Begin VB.CommandButton cmdRoller 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   244
         Index           =   0
         Left            =   150
         Picture         =   "Main.frx":1882
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   210
         Width           =   244
      End
   End
   Begin VB.Frame fraMeas 
      BackColor       =   &H0080C0FF&
      Caption         =   " Measure "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3930
      TabIndex        =   103
      Top             =   5145
      Width           =   1050
      Begin VB.CommandButton cmdCloseMeas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   900
         Width           =   330
      End
      Begin VB.Label LabMeas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   105
         Top             =   585
         Width           =   615
      End
      Begin VB.Label LabMeas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   104
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   36
      Left            =   390
      Picture         =   "Main.frx":1954
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   " Show angle & length "
      Top             =   1425
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   2760
      Picture         =   "Main.frx":1A26
      Style           =   1  'Graphical
      TabIndex        =   98
      ToolTipText     =   " Swap with view below "
      Top             =   15
      Width           =   300
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   37
      Left            =   45
      Picture         =   "Main.frx":1B70
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   " Color picker "
      Top             =   1425
      Width           =   300
   End
   Begin VB.PictureBox picPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   10620
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   96
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   10575
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   95
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   375
      Picture         =   "Main.frx":1CBA
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   " Print selection "
      Top             =   1005
      Width           =   300
   End
   Begin VB.CommandButton cmdFile 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   375
      Picture         =   "Main.frx":1E04
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   " Save selection "
      Top             =   660
      Width           =   300
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   45
      Picture         =   "Main.frx":1F4E
      Style           =   1  'Graphical
      TabIndex        =   86
      ToolTipText     =   " Print "
      Top             =   1005
      Width           =   300
   End
   Begin VB.Frame fraStrip 
      BackColor       =   &H0080C0FF&
      Caption         =   "  Strip "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   45
      TabIndex        =   80
      ToolTipText     =   " Make strip from circular selection "
      Top             =   1770
      Width           =   780
      Begin VB.CommandButton cmdStrip 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         Picture         =   "Main.frx":2098
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   " Make strip from circular selection "
         Top             =   225
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   285
         Picture         =   "Main.frx":216A
         ToolTipText     =   " Make strip from circular selection "
         Top             =   915
         Width           =   240
      End
      Begin VB.Label LabNumStrips 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "99/1024"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   150
         TabIndex        =   85
         ToolTipText     =   " Make strip from circular selection "
         Top             =   570
         Width           =   495
      End
   End
   Begin VB.Frame fraRot 
      BackColor       =   &H0080C0FF&
      Caption         =   " Angle "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   30
      TabIndex        =   76
      ToolTipText     =   " Circular selection rotation "
      Top             =   3000
      Width           =   780
      Begin VB.CommandButton cmdAngle 
         BackColor       =   &H0080C0FF&
         Caption         =   "+-10"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   " Circular selection rotation "
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton cmdAngle 
         BackColor       =   &H0080C0FF&
         Caption         =   "+-1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   " Circular selection rotation "
         Top             =   450
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   465
         Picture         =   "Main.frx":223C
         ToolTipText     =   " Circular selection rotation "
         Top             =   195
         Width           =   240
      End
      Begin VB.Label LabRot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-180"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   77
         ToolTipText     =   " Circular selection rotation "
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   2460
      Picture         =   "Main.frx":230E
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   " Add view above "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdStopUndos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   10905
      Picture         =   "Main.frx":2898
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   " Toggle undos "
      Top             =   7155
      Width           =   300
   End
   Begin VB.CommandButton cmdStopUndos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   10500
      Picture         =   "Main.frx":2E22
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   " Toggle undos "
      Top             =   7155
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   3120
      Picture         =   "Main.frx":33AC
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   " On/Off backups "
      Top             =   15
      Width           =   300
   End
   Begin VB.TextBox txtSyntax 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   6870
      MultiLine       =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Main.frx":3936
      Top             =   420
      Width           =   3285
   End
   Begin VB.Frame fraDrawInstr 
      BackColor       =   &H0080C0FF&
      Caption         =   "Draw Instructions "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6345
      TabIndex        =   45
      Top             =   60
      Width           =   3630
      Begin VB.Label LabDrawInstructions 
         BackColor       =   &H0080C0FF&
         Caption         =   "LC - D - LC - M - LC or RC"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   46
         Top             =   225
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1230
      Picture         =   "Main.frx":3A09
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   " Clear current view "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   2160
      Picture         =   "Main.frx":3B53
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Collapse to top & bottom "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1860
      Picture         =   "Main.frx":3C9D
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Delete views above "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1560
      Picture         =   "Main.frx":3DE7
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   " Delete current view "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   885
      Picture         =   "Main.frx":3F31
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   " Redo "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   585
      Picture         =   "Main.frx":407B
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   " Undo "
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton cmdFile 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   45
      Picture         =   "Main.frx":41C5
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   " Save "
      Top             =   660
      Width           =   300
   End
   Begin VB.CommandButton cmdFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   45
      Picture         =   "Main.frx":430F
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   " Browse "
      Top             =   345
      Width           =   300
   End
   Begin VB.CommandButton cmdFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   60
      Picture         =   "Main.frx":4459
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   " New "
      Top             =   15
      Width           =   300
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   45
      TabIndex        =   15
      Top             =   6975
      Width           =   10260
      Begin VB.Label LabInfo 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   8775
         TabIndex        =   18
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label LabInfo 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " W,H="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   7035
         TabIndex        =   17
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label LabInfo 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File= D:\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   6945
      End
   End
   Begin VB.Frame fraPAL 
      BackColor       =   &H0080C0FF&
      Caption         =   "Palette "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6480
      Left            =   1950
      TabIndex        =   4
      Top             =   495
      Width           =   1410
      Begin VB.CommandButton cmdCPal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rotate"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   " Rotate palette "
         Top             =   690
         Width           =   570
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   " Invert palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.CommandButton cmdCPal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   915
         Picture         =   "Main.frx":45A3
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   " Reset palette "
         Top             =   495
         Width           =   390
      End
      Begin VB.CommandButton cmdCPal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   495
         Picture         =   "Main.frx":46AD
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   " Darken "
         Top             =   480
         Width           =   390
      End
      Begin VB.CommandButton cmdCPal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         Picture         =   "Main.frx":47B7
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   " Brighten "
         Top             =   480
         Width           =   390
      End
      Begin VB.CommandButton cmdSwapLR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   735
         Picture         =   "Main.frx":48C1
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   " Swap Left & Right colors "
         Top             =   930
         Width           =   240
      End
      Begin VB.CommandButton cmdSwapBW 
         Caption         =   "<>"
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
         Left            =   210
         TabIndex        =   56
         ToolTipText     =   " Swap B/W background {B}  "
         Top             =   1200
         Width           =   300
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   900
         Picture         =   "Main.frx":4953
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   " Center banded palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   705
         Picture         =   "Main.frx":4E95
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   " 16  Banded  palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.CommandButton cmdDefPalette 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1095
         TabIndex        =   53
         ToolTipText     =   " Default palette "
         Top             =   720
         Width           =   210
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   510
         Picture         =   "Main.frx":53D7
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " 32 Banded palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         Picture         =   "Main.frx":5919
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   " Grey palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.CommandButton cmdStartPalette 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   " Make/Load a palette "
         Top             =   210
         Width           =   195
      End
      Begin VB.PictureBox picPAL 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   180
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   5
         Top             =   1425
         Width           =   1020
         Begin VB.Shape shpPAL 
            BorderColor     =   &H00808000&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   135
            Index           =   1
            Left            =   450
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape shpPAL 
            BackColor       =   &H00000000&
            BorderColor     =   &H00C0FFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   315
            Width           =   135
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   210
         Left            =   210
         Top             =   945
         Width           =   495
      End
      Begin VB.Label LabSelCul 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   1170
         TabIndex        =   14
         Top             =   1125
         Width           =   150
      End
      Begin VB.Label LabSelCul 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   75
         TabIndex        =   13
         Top             =   915
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "R,G,B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   5895
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Color #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   5655
         Width           =   555
      End
      Begin VB.Label LabSelCul 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255, 255, 255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   195
         TabIndex        =   10
         Top             =   6105
         Width           =   1035
      End
      Begin VB.Label LabSelCul 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "123"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   5655
         Width           =   495
      End
      Begin VB.Label LabSelCul 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Label LabSelCul 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   960
         Width           =   465
      End
      Begin VB.Label LabSelCul 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   690
         TabIndex        =   7
         Top             =   1170
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   240
         Left            =   675
         Top             =   1155
         Width           =   495
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   5805
      Left            =   3525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   240
   End
   Begin VB.HScrollBar HS 
      Height          =   210
      Left            =   3990
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6540
      Width           =   5910
   End
   Begin VB.PictureBox PICC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillStyle       =   6  'Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   3810
      ScaleHeight     =   391
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   585
      Width           =   6360
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   1  'Dash
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   30
         ScaleHeight     =   228
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   1
         Top             =   15
         Width           =   3840
         Begin VB.Line SL 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   0
            X1              =   4
            X2              =   -30
            Y1              =   2
            Y2              =   2
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF80FF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   14
            X2              =   14
            Y1              =   4
            Y2              =   -30
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF80FF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   -41
            X2              =   6
            Y1              =   9
            Y2              =   9
         End
         Begin VB.Line LineMeas 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   8
            X2              =   8
            Y1              =   -29
            Y2              =   3
         End
         Begin VB.Shape shpEllip 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            Height          =   330
            Left            =   225
            Shape           =   2  'Oval
            Top             =   1155
            Width           =   585
         End
         Begin VB.Shape shpCirc 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            Height          =   330
            Left            =   330
            Shape           =   3  'Circle
            Top             =   720
            Width           =   315
         End
         Begin VB.Shape shpRect 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            Height          =   330
            Left            =   240
            Top             =   255
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraTools 
      BackColor       =   &H0080C0FF&
      Caption         =   "Tools "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6480
      Left            =   825
      TabIndex        =   31
      Top             =   495
      Width           =   1065
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   21
         Left            =   195
         Picture         =   "Main.frx":5E5B
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   " Select lasso "
         Top             =   1185
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   35
         Left            =   195
         Picture         =   "Main.frx":5F2D
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   " Replace Left by Right color "
         Top             =   5850
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   23
         Left            =   165
         Picture         =   "Main.frx":6077
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   " Copy & Paste "
         Top             =   1815
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   585
         Picture         =   "Main.frx":6149
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   " Text "
         Top             =   6000
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   34
         Left            =   150
         Picture         =   "Main.frx":621B
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   " Pepper "
         Top             =   5475
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   33
         Left            =   150
         Picture         =   "Main.frx":62ED
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   " Thicken all objects "
         Top             =   5145
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   32
         Left            =   135
         Picture         =   "Main.frx":6437
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   " Mix colors "
         Top             =   4800
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   31
         Left            =   150
         Picture         =   "Main.frx":6509
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   " Rotate by 90 degrees "
         Top             =   4470
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   30
         Left            =   180
         Picture         =   "Main.frx":65DB
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   " Clear selection "
         Top             =   4095
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   29
         Left            =   180
         Picture         =   "Main.frx":6725
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   " Paste "
         Top             =   3840
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   28
         Left            =   150
         Picture         =   "Main.frx":67F7
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   " Copy Rotate "
         Top             =   3555
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   27
         Left            =   195
         Picture         =   "Main.frx":68C9
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   " Copy Reflect UD "
         Top             =   3180
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   26
         Left            =   180
         Picture         =   "Main.frx":699B
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   " Copy Reflect LR "
         Top             =   2835
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   25
         Left            =   165
         Picture         =   "Main.frx":6A6D
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   " Cut "
         Top             =   2505
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   24
         Left            =   150
         Picture         =   "Main.frx":6B3F
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   " Copy "
         Top             =   2145
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   22
         Left            =   165
         Picture         =   "Main.frx":6C11
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   " Deselect "
         Top             =   1485
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   180
         Picture         =   "Main.frx":6D5B
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   " Select ellipse "
         Top             =   930
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   195
         Picture         =   "Main.frx":6E2D
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   " Select circle "
         Top             =   630
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   195
         Picture         =   "Main.frx":6EFF
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   " Select rectangle "
         Top             =   315
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   585
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   " Arrows "
         Top             =   5685
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   " Bushes "
         Top             =   5325
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   " Brushes "
         Top             =   375
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   " Radials "
         Top             =   4635
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   " Arcs "
         Top             =   3990
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   " Shapes "
         Top             =   4305
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   " Fills "
         Top             =   4980
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   " Junctions "
         Top             =   3660
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   " Bullets "
         Top             =   3330
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   " Tubes "
         Top             =   2985
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   " Cones "
         Top             =   2640
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   " Cirllipses "
         Top             =   2325
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Rectangles "
         Top             =   1995
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   " CurvyLines "
         Top             =   1650
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   " PolyLines "
         Top             =   1365
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   " Lines "
         Top             =   1065
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   " Sprays "
         Top             =   705
         Width           =   300
      End
      Begin VB.Label LabQuery2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   90
         Top             =   165
         Width           =   135
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   75
         Top             =   180
         Width           =   390
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         DrawMode        =   2  'Blackness
         Height          =   3930
         Left            =   90
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   134
      Top             =   225
      Width           =   750
   End
   Begin VB.Label LabCentreForm 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3375
      TabIndex        =   117
      ToolTipText     =   " Center form "
      Top             =   6810
      Width           =   135
   End
   Begin VB.Label LabPark 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3360
      TabIndex        =   99
      Top             =   1275
      Width           =   105
   End
   Begin VB.Label LabSTO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " RC to set Tool Options"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1050
      TabIndex        =   91
      Top             =   345
      Width           =   1410
   End
   Begin VB.Label LabLight 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3360
      TabIndex        =   89
      Top             =   585
      Width           =   105
   End
   Begin VB.Label LabInfo 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3435
      TabIndex        =   57
      Top             =   15
      Width           =   915
   End
   Begin VB.Label LabQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9990
      TabIndex        =   47
      Top             =   180
      Width           =   150
   End
   Begin VB.Label LabXY 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabXY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5415
      TabIndex        =   19
      Top             =   225
      Width           =   810
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "&Browser"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveSelection 
         Caption         =   "Sa&ve selection"
      End
      Begin VB.Menu mnuStartPalette 
         Caption         =   "&Palette"
         Begin VB.Menu mnuPal 
            Caption         =   "&Make/Load a palette"
            Index           =   0
         End
         Begin VB.Menu mnuPal 
            Caption         =   "&Greyed palette"
            Index           =   1
         End
         Begin VB.Menu mnuPal 
            Caption         =   "32 &banded palette"
            Index           =   2
         End
         Begin VB.Menu mnuPal 
            Caption         =   "&16 banded palette"
            Index           =   3
         End
         Begin VB.Menu mnuPal 
            Caption         =   "&Centered banded palette"
            Index           =   4
         End
      End
      Begin VB.Menu zbrk0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "P&rint"
      End
      Begin VB.Menu mnuPrintSelection 
         Caption         =   "Pr&int selection"
      End
      Begin VB.Menu zbrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCanvas 
      Caption         =   "&Resizing"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuClearCurrentView 
         Caption         =   "&Clear current view"
      End
      Begin VB.Menu mnuDeleteCurrentView 
         Caption         =   "&Delete current view"
      End
      Begin VB.Menu mnuDeleteAbove 
         Caption         =   "Delete &views above"
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "C&ollapse to top && bottom"
      End
      Begin VB.Menu mnuAddViewAbove 
         Caption         =   "&Add view above"
      End
      Begin VB.Menu mnuSwapViews 
         Caption         =   "&Swap views"
      End
      Begin VB.Menu mnuToggleUndos 
         Caption         =   "On/Off &backups"
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "&Show views"
   End
   Begin VB.Menu mnuTransforms 
      Caption         =   "&Transforms"
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' APaint8  by  Robert Rayment (May 2004) 12

' Form1 (Main.frm)

' Best for screen 1024 x 768 or greater.
' Compile

' Updates
'1.   8/5/04  Preserve image saved with white background
'2.  11/5/04  Flush mouse & avoid DISPLAY_ALL_VIEWS
'             unless view stack full or menu action aMNUACT = True
'3.  12/5/04  Update Browser to avoid optimisation
'             of Gifs when W & H a multiple of 4.


Option Explicit
Option Base 1

' Variables starting with:-
' b      byte
' a      boolean
' x,y,z  single
' else   long
' ___________________________
'| ___________________       |    _______
'||                   |      |   |       |
'||  PIC              |      |   |picMask| bMask()
'||  bArray()         |      |   |_______|
'||                   |      |    _______
'||___________________|      |   |       |
'|                           |   |picPIC | bPic()
'|                     PICC  |   |_______|
'|___________________________|


' Minimum form size
Dim OrgFW As Long
Dim OrgFH As Long
' Maintained resizing values
Dim aResize As Boolean
Dim RightGap As Long  ' PICC
Dim BottomGap As Long ' PICC
Dim fraInfoBottomGap As Long  ' fraInfo
Dim actVWH As Boolean

Private CommonDialog1 As OSDialog

Private Sub Form_Initialize()
Dim k As Long
' All Public
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   MAXWIDTH = 2048 '1024
   MAXHEIGHT = 2048 '1024
   
   canvasW = 256  ' Start Canvas size
   canvasH = 256
   ZoomFactor = 2
   
   ' Original image
   StartWidth = 256
   StartHeight = 256
   ReDim bArray(StartWidth, StartHeight)
   
   ' Zoom image
   ZoomSize = 240
   ZoomWidth = ZoomSize
   ZoomHeight = ZoomSize
   
   ' Long color holders
   ReDim CulRGB(0 To 255)
   ReDim CulBGR(0 To 255)
   
   SetfrmDefaultPositions
   
   FillInstrucLabels
   
   xprev = 0
   yprev = 0
   
   ReDim RadialRep(5)
   For k = 1 To 5
      RadialRep(k) = 8
   Next k
   
   ' Simple full fill
   ReDim bPattern(16, 16)
   FillMemory bPattern(1, 1), 256, 1
   
   ' Tree levels
   SetUpFractalTrees

   SizeUndoPalettes
   ReDim StorePal(0 To 2, 0 To 255)
   ReDim BackUpRGB(0 To 255)
   
   With SVFont
      .FontName = "Arial"
      .FontSize = 8
      .FontBold = False
      .FontItalic = False
   End With
End Sub

Private Sub Form_Load()
   KeyPreview = True
   picMask.Visible = False
   'picPic.Visible = False
   cmdStopUndos(0).Visible = False
   cmdStopUndos(1).Visible = False
   '---------------------------------------
   ' Public AppPathSpec$, OpenPathSpec$, SavePathSpec$
   ' & PaintInfo.txt
   GetSetUpInfo
   '---------------------------------------
   ' Load from RES  Id:="GETINDEX" Type:="CUSTOM"
   GetIndexMC = LoadResData("GETINDEX", "CUSTOM")
   ptMC = VarPtr(GetIndexMC(0))
   ' Check 1st 2 bytes
'   Dim AA$, BB$
'   AA$ = Hex$(GetIndexMC(0))
'   BB$ = Hex$(GetIndexMC(1))
   '---------------------------------------
   
   PIC.Width = 256 'picW
   PIC.Height = 256 'picH
   ' Default PIC container size
   PICCW = 424
   PICCH = 412
   PICC.Width = PICCW
   PICC.Height = PICCH
   FixScrollbars PICC, PIC, HS, VS
   '---------------------------------------
   ' For different OSs
   GetExtras Me.BorderStyle
   ' OUT: ExtraHeight,  ExtraBorder
   aResize = False
   Me.Height = 7760 + ExtraHeight * STY
   aResize = True
   OrgFW = Me.Width
   OrgFH = Me.Height
   RightGap = Me.Width \ STX - PICC.Left - PICC.Width
   BottomGap = Me.Height \ STY - PICC.Top - PICC.Height
   fraInfoBottomGap = Me.Height \ STY - fraInfo.Top - fraInfo.Height
   '---------------------------------------
   LineUpTools
   SetInitialDrawTools
   '---------------------------------------
   NewNum = 0
   NumLassoLines = 1
   mnuNew_Click
   '---------------------------------------
   Show  ' To Form_Resize now
   
   zangRotCSEL = 0
   LabRot = Str$(zangRotCSEL)
   
   aVIEWS = False
   aMNUACTION = False
   
   aHairs = False
   chkHairs = Unchecked
   Line2.Visible = False
   Line3.Visible = False

   ASELECTION = False
   
   LabSTO.Visible = False
   
   aLensCheck = False
   aUseSelectcn = True
   
   optRollShift(0).Value = True
   RollShift = 0

   ' Backup Current Pal
   BackUpRGB() = CulRGB()

   PIC.MousePointer = vbCustom
   PIC.MouseIcon = LoadResPicture(101, vbResCursor)

End Sub

Private Sub mnuNew_Click()
Dim k As Long
   If ADRAW Then Exit Sub
   
   If NewNum > 1 Then
      If MsgBox("Restart", vbQuestion + vbYesNo, "New") = vbNo Then
         ' Offset cursor to avoid click thru
         SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
         Exit Sub
      End If
   End If
   '---------------------------------------

   NewNum = 1
   ADRAW = False
   
   'Undos
   UndoNum = 0
   TopUndoNum = 0
   ReDimUndos
   StopUndos = False
   cmdEdit(8).Picture = cmdStopUndos(1).Picture
   cmdStopUndos(0).Visible = False
   cmdStopUndos(1).Visible = False
   
   ' Menus
   mnuSave.Enabled = False
   cmdFile(2).Enabled = False
   mnuSaveSelection.Enabled = False
   cmdFile(3).Enabled = False
   mnuPrint.Enabled = False
   cmdPrint(0).Enabled = False
   mnuPrintSelection.Enabled = False
   cmdPrint(1).Enabled = False
   DoEvents
   
   FixUndos
   Unload frmCanvas
   Unload frmZoom
   mnuZoom.Enabled = False
   aZoom = False
   Unload frmViews
   mnuViews.Enabled = False
   aVIEWS = False
   mnuTransforms.Enabled = False
   Unload frmHelp
   DoEvents
   
   ' Color
   ' Left & Right color shape markers
   Shape1.BorderColor = vbWhite
   Shape2.BorderColor = vbRed
   shpPAL(0).BorderColor = picPAL.BackColor Xor vbWhite  ' Left color
   shpPAL(1).BorderColor = picPAL.BackColor Xor vbRed    ' Right color
   picPAL.Width = 68 * STX
   picPAL.Height = 260 * STY
   InitDefaultPalette
   AdjustPalette
   ' Set default L/R selected colors
   picPAL_MouseUp 2, 0, 4, 4     ' Right button color
   picPAL_MouseUp 1, 0, 8, 4     ' Left button color
   SelLeftCulNum = 1
   SelRightCulNum = 0
   LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
   LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
   
   fraMeas.Left = 256
   fraMeas.Top = 350
   
   PIC.Picture = LoadPicture
   PIC.BackColor = CulRGB(0)
   
   GETBYTES PIC.Image, bArray(), canvasW, canvasH, 8, CulRGB(), 0
   '---------------------------------------
   FileSpec$ = "New"
   CopyFileSpec$(0) = "New"
   LabInfo(0) = FileSpec$
   
   ' Tool stuff
   ' Tool syntax
   txtSyntax.Visible = False
   ' Mouse click counters
   LCNum = 0
   RCNum = 0
   ReDim XT(12), YT(12)

   '---------------------------------------
   ' No image so:-
   ' disable all tools requiring an image
   ' Default start Tool
   optTools(Brush).Value = True
   ToolType = 0
   
   For k = Measure To Smooth4
      optTools(k).Value = False
   Next k
   LineMeas.Visible = False
   aMeasure = False
   fraMeas.Visible = False
   
   fraRoller.Enabled = False
   fraSmoothers.Enabled = False
   DoDeselect
   cmdStrip.Enabled = False
   aPepper = False
   LabNumStrips = "N = 0"
   For k = SelR To Desel
      optTools(k).Enabled = False
   Next k
   optTools(Rot90).Enabled = False
   optTools(Mix).Enabled = False
   optTools(Thicken).Enabled = False
   optTools(LRColor).Enabled = False
   
   FillLabInfos
End Sub

Private Sub LineUpTools()
Dim k As Long
Dim ttop As Long
Dim toptleft As Long
Dim tvstep As Long
   Shape3.Top = 190
   Shape3.Height = 4170
   Label2.Top = 210
   ttop = 405
   optTools(0).Top = ttop
   toptleft = 630
   tvstep = 330
   For k = 0 To 17
      optTools(k).Left = toptleft
   Next k
   For k = 1 To 17
      optTools(k).Top = optTools(0).Top + k * tvstep
      optTools(k).TabStop = False
   Next k
   
   toptleft = 165
   For k = 18 To 35
      optTools(k).Left = toptleft
   Next k
   For k = 18 To 35
      optTools(k).Top = optTools(0).Top + (k - 18) * tvstep
      optTools(k).TabStop = False
   Next k
   
   fraStrip.Left = 3
   fraRot.Left = 3
End Sub

Private Sub SetInitialDrawTools()
Dim k As Long
   ToolType = -1 'Brush
   frmToolOptions.Show vbModeless ' 0
   Unload frmToolOptions
   ToolType = Brush
   optTools_MouseUp 0, 1, 0, 0, 0
End Sub

' Measure frame
Private Sub fraMeas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Xfra = x
   Yfra = y
End Sub

Private Sub fraMeas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, x As Single, y As Single)
   fraMOVER Form1, fraMeas, Button, x, y
End Sub

Private Sub cmdCloseMeas_Click()
 fraMeas.Visible = False
 optTools(Measure).Value = False
End Sub

' Roller/Shifter
Private Sub optRollShift_Click(Index As Integer)
   If optRollShift(0) Then
      RollShift = 0
   Else
      RollShift = 1
   End If
End Sub

Private Sub cmdRoller_Click(Index As Integer)
'  <-0   ->1  (1)
' <<-2  ->>3  (8)

'   4 (1)   6  (8)
'   ^       ^
'   |       ^
'           |

'   5       7
'   |       |
'   v       v
'           v

   Select Case Index
   Case 0: RollLeft 1, RollShift
   Case 1: RollRight 1, RollShift
   Case 2: RollLeft 8, RollShift
   Case 3: RollRight 8, RollShift
   Case 4: RollUp 1, RollShift
   Case 5: RollDown 1, RollShift
   Case 6: RollUp 8, RollShift
   Case 7: RollDown 8, RollShift
   End Select
   LCNum = -1
   DISPSAVE
   FillLabInfos
End Sub

Private Sub LabCentreForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Center form
   If WindowState <> vbMaximized Then
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Top = (Screen.Height - Me.Height) / 2
   End If
End Sub

Private Sub LabPark_Click()
' Park mouse
   LabPark.BackColor = vbBlack
End Sub


'#### Canvas size ############

Private Sub mnuCanvas_Click()
   If ADRAW Then Exit Sub
   Unload frmZoom
   aZoom = False
   Unload frmToolOptions
   Unload frmPalette
   Unload frmHelp
   If aVIEWS Then
      frmViewsLeft = frmViews.Left
      frmViewsTop = frmViews.Top
   End If
   Unload frmViews
   'aVIEWS = False   ' Leave, if True will reshow frmViews
   DoEvents
   
   frmCanvas.Show vbModal '1
   ' returns new W & H or cancel
   ' & new bArray()

   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
      
   If aCanWH Then
      optTools_MouseUp Desel, 1, 0, 0, 0
      PIC.Width = canvasW
      PIC.Height = canvasH
      If NewNum > 1 Then   ' ie not New
         SAVE_CurrentImage
         FixUndos          ' unless StopUndos = True
      End If
      FillLabInfos
      DISPLAY
   End If
   
   If aVIEWS Then frmViews.Show 0

End Sub

Private Sub mnuEdit_Click()
   If ADRAW Then Exit Sub
End Sub

Private Sub mnuFile_Click()
   If ADRAW Then Exit Sub
End Sub

Private Sub mnuHelp_Click()
   frmHelp.Show 0 '1
End Sub


'######################################################

'Private Sub chkHairs_Click()
Private Sub chkHairs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aHairs = chkHairs.Value
   If Not aHairs Then
      Line2.Visible = False
      Line3.Visible = False
   Else
      'Line2.Visible = True
      'Line3.Visible = True
   End If
End Sub

Private Sub cmdPrint_Click(Index As Integer)
Dim APErr As Boolean
Dim res As Long
   If ADRAW Then Exit Sub
   
   PIC.Picture = LoadPicture  ' Too reset PIC memory
   DISPLAY
   
   res = MsgBox("IS PRINTER LIVE!", vbQuestion + vbYesNo + vbSystemModal, "Printing")
   If res = vbYes Then
      ShowPrinter Me, APErr
      If Not APErr Then
         Select Case Index
         Case 0   ' Print PIC
            Printer.PaintPicture PIC.Image, Printer.Width / 12, Printer.Height / 12
            Printer.EndDoc
         Case 1   ' Print selection picPIC
            If ASELECTION Then
               MakeMask
               ReDim bMask(SSW, SSH)
               GetPICBytes picMask.Image, bMask(), SSW, SSH
               GetbPic
               DISPLAYpicPic
               Printer.PaintPicture picPic.Image, Printer.Width / 12, Printer.Height / 12
               Printer.EndDoc
            End If
         End Select
      End If
      DoEvents
   End If
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight

End Sub

Private Sub cmdStrip_Click()
Dim k As Long
Dim zIncR As Single, zIncrFracReduc As Single
Dim zang As Single
Dim zPepperFrac As Single, zRndLim As Single
Dim ixd As Long, iyd As Long
Dim xc As Single, yc As Single

   Unload frmToolOptions
   Unload frmPalette
   Unload frmZoom
   aZoom = False
   DoEvents
   frmStrip.Show vbModal '1

   If NumStrips > 0 Then
      'Public bDummy() As Byte
      'Public bMask() As Byte
      'Public MaxNumFrames As Long
      'Public NumStrips As Long
      'Public zTotalAng As Single
      'Public zIncrAng As Single
      'Public zFinalPercentReduc As Single
      'Public zIncrPercentReduc As Single
      'Public canvasW As Long ''''
      'Public SSW ' Circular select width
      'Public Sub RotateEllbRect(zANGLE)
      ' Public zangRot as single   ' degrees
   
      MakeMask
      ReDim bMask(SSW, SSH)
      GetPICBytes picMask.Image, bMask(), SSW, SSH
      GetbPic
      xc = SSW / 2 + 0.5
      yc = SSH / 2 + 0.5
      zIncrFracReduc = zIncrPercentReduc / 100
      zPepperFrac = 1 / NumStrips
      canvasW = SSW * NumStrips
      canvasH = SSH
      ReDim bDummy(SSW, SSH)
      ReDim bCopy(SSW, SSH) As Byte
      ReDim bArray(canvasW, canvasH)
      DoEvents
      bCopy() = bPic()
      bDummy() = bPic()
      For k = 0 To NumStrips - 1
         ' Transfer first image unchanged
         For iy = 1 To SSH
         For ix = k * SSW + 1 To (k + 1) * SSW
            bArray(ix, iy) = bDummy(ix - (k * SSW), iy)
         Next ix
         Next iy
         
         ' Rotate & reduce
         If zTotalAng > 0 And zFinalPercentReduc <= 100 Then
            zangRot = zIncrAng * (k + 1)  ' degrees
            RotateEllbRect zangRot    ' bArray to bMask(SSW,SSH)
            zIncR = zIncrFracReduc * (k + 1)
            ReDim bDummy(SSW, SSH)
            zIncR = zIncrFracReduc * (k + 1)
            For iy = 1 To SSH
            For ix = 1 To SSW
               zrad = Sqr((iy - yc) ^ 2 + (ix - xc) ^ 2)
               zang = zATan2(iy - yc, ix - xc)
               ixd = xc + zrad * (1 - zIncR) * Cos(zang)
               iyd = yc + zrad * (1 - zIncR) * Sin(zang)
               bDummy(ixd, iyd) = bPic(ix, iy)
            Next ix
            Next iy
            ' Recover original Pic
            bPic() = bCopy()
         ' Reduce only
         ElseIf zFinalPercentReduc < 100 Then
            ' Move every point along it's radius
            ' towards the center xc,yc
            ReDim bDummy(SSW, SSH)
            zIncR = zIncrFracReduc * (k + 1)
            For iy = 1 To SSH
            For ix = 1 To SSW
               zrad = Sqr((iy - yc) ^ 2 + (ix - xc) ^ 2)
               zang = zATan2(iy - yc, ix - xc)
               ixd = xc + zrad * (1 - zIncR) * Cos(zang)
               iyd = yc + zrad * (1 - zIncR) * Sin(zang)
               bDummy(ixd, iyd) = bPic(ix, iy)
            Next ix
            Next iy
         ' Rotate only
         ElseIf zTotalAng > 0 Then
            zangRot = zIncrAng * (k + 1)  ' degrees
            RotateEllbRect zangRot    ' To bPic
            ' Put result in bDummy for transfer to bArray
            ReDim bDummy(SSW, SSH)
            bDummy() = bPic()
            ' Recover original Pic
            bPic() = bCopy()
         End If
         If aPepper Then
            zRndLim = 1 - zPepperFrac * k
            For iy = 1 To SSH
            For ix = 1 To SSW
               If Rnd - zRndLim > 0 Then
                  bDummy(ix, iy) = 0
               End If
            Next ix
            Next iy
         End If
      Next k
      
      SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
      DISPLAY
      FixUndos           ' unless StopUndos = True
      FillLabInfos
      Erase bDummy(), bCopy()
      
      LabNumStrips = "N=" & Str$(NumStrips) & vbCr & " L=" & Str$(canvasW \ NumStrips)
      LabNumStrips.Refresh
   End If
   optTools_MouseUp Desel, vbLeftButton, 0, 0, 0
End Sub

' Angle for circular selection
Private Sub cmdAngle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   aDone = False
   Do
      Select Case Index
      Case 0   ' +/- 1
         Select Case Button
         Case vbLeftButton    ' +1
            If zangRotCSEL < 180 Then
               zangRotCSEL = zangRotCSEL + 1
            End If
         Case vbRightButton   ' -1
            If zangRotCSEL > -180 Then
               zangRotCSEL = zangRotCSEL - 1
            End If
         End Select
         LabRot = Str$(zangRotCSEL)
         Sleep 150
         DoEvents
      
      Case 1   '+/-10
         Select Case Button
         Case vbLeftButton
            If zangRotCSEL <= 170 Then
               zangRotCSEL = zangRotCSEL + 10
            End If
         Case vbRightButton
            If zangRotCSEL >= -170 Then
               zangRotCSEL = zangRotCSEL - 10
            End If
         End Select
         LabRot = Str$(zangRotCSEL)
         Sleep 100
         DoEvents
      End Select
   Loop Until aDone
End Sub

Private Sub cmdAngle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   aDone = True
End Sub

Private Sub cmdSwapLR_Click()
' Swap Left & Right color selection
Dim NN As Long
Dim XX As Single, YY As Single

   NN = SelLeftCulNum
   SelLeftCulNum = SelRightCulNum
   SelRightCulNum = NN
   
   ' New Left color
   LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
   YY = (8 * (SelLeftCulNum \ 8) + 4)
   XX = (8 * (SelLeftCulNum - 8 * (SelLeftCulNum \ 8)) + 4)
   picPAL_MouseUp 1, 0, XX, YY   ' Left button color
   ' New Right color
   LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
   YY = 8 * (SelRightCulNum \ 8) + 4
   XX = 8 * (SelRightCulNum - 8 * (SelRightCulNum \ 8)) + 4
   picPAL_MouseUp 2, 0, XX, YY   ' Right button color
End Sub

Private Sub LabQuery2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabSTO.Visible = True
End Sub

Private Sub LabQuery2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabSTO.Visible = False
End Sub

Private Sub mnuExit_Click()
   Form_QueryUnload 1, 0
End Sub

Private Sub mnuPal_Click(Index As Integer)
   cmdStartPalette_Click Index
End Sub

Private Sub mnuPrint_Click()
   cmdPrint_Click 0
End Sub

Private Sub mnuPrintSelection_Click()
   cmdPrint_Click 1
End Sub

Private Sub mnuTransforms_Click()
If ADRAW Then Exit Sub
   
   Unload frmZoom
   aZoom = False
   Unload frmToolOptions
   Unload frmPalette
   Unload frmHelp
   If aVIEWS Then
      frmViewsLeft = frmViews.Left
      frmViewsTop = frmViews.Top
   End If
   Unload frmViews
   'aVIEWS = False   ' Leave, if True will reshow frmViews
   DoEvents
   
   frmTransform.Show vbModal '1  ' returns image FileSpec$

   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   
      
   If aVIEWS Then frmViews.Show 0

   ' For Cancel
   ' TransformType = TNone  ' 0
   LabLight.BackColor = vbRed
   DoEvents
   
   If Not aSelRect Then
      WWLO = 1: HHLO = 1
      WWHI = canvasW: HHHI = canvasH
   Else
      WWLO = SSX
      HHHI = canvasH - SSY
      WWHI = WWLO + SSW
      HHLO = HHHI - SSH
   End If
   
   
   Select Case TransformType
   Case TContour:       Contour
   Case TDither:        Dither
   Case TEngraveEmboss: EngraveEmboss
   Case TPosterize:     Posterize
   Case TRelief:        Relief
   Case TSmooth:        Smooth
   Case TShadeV:        ShadeV
   Case TShadeH:        ShadeH
   Case TMelt:          Melt
   Case TOil:           Oil
   Case TSharpen:       Sharpen
   Case TLitho:         Litho
   Case TContrast:      Contrast
   Case TDiffuse, THDiffuse, TVDiffuse: Diffuse
   Case TBlackWhite:    BlackWhite
   Case TSolar:         Solarize
   Case TInvert:        Invert
   Case TFog:           Fog
   Case TSquare:        Pixelize
   
   Case TEllipse:   Elliptic
   Case TFluteH:    FluteH
   Case TFluteV:    FluteV
   Case TRippleH:   RippleH
   Case TRippleV:   RippleV
   Case TRoundRect: RoundRect
   Case TTile:      Tile
   Case TMirrorL:   MirrorLeft
   Case TMirrorR:   MirrorRight
   Case TMirrorT:   MirrorTop
   Case TMirrorB:   MirrorBottom
   Case TMlens:     MirrorLens
   Case TLens:      ALens
   Case TFWindowHorz: FlutedWindowHorz
   Case TFWindowVert: FlutedWindowVert
   Case TFWindowHV:   FlutedWindowHV
   Case TSwirl:     Swirl
   Case TSpokess:   Stars
   Case TMinMag:    MinMag
   Case TBubbly:    Bubbly
   Case TRotate:    Rotate
   Case TTunnel:    Tunnel
   
   Case THLines:      AddHLines
   Case TVLines:      AddVLines
   Case THVLines:     AddHVLines
   Case THWaves:      AddHWaves
   Case TVWaves:      AddVWaves
   Case THVWaves:     AddHVWaves
   Case TCircles:     AddCircles
   Case TEllipses:    AddEllipses
   Case TThickLineH:  AddThickLineH
   Case TThickLineV:  AddThickLineV
   Case TThickLineHV: AddThickLineHV
   Case TBorder:      AddBorder
   Case TSpokes:      AddSpokes
   Case TDNet:        AddDiagNet
   End Select

   If TransformType > 0 Then
      If Not StopUndos Then
         LCNum = -1
         DISPSAVE
      Else
         SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
         LCNum = -1
         DISPLAY
      End If
      FillLabInfos
      LCNum = 0: RCNum = 0
   End If
   LabLight.BackColor = vbGreen

End Sub

Private Sub mnuViews_Click()
If ADRAW Then Exit Sub
   
   Unload frmZoom
   aZoom = False
   Unload frmToolOptions
   Unload frmPalette
   frmViews.Show 0
   
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight

End Sub

Private Sub LabQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtSyntax.Visible = True
End Sub

Private Sub LabQuery_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtSyntax.Visible = False
End Sub

'#### TOOLS  ###############################################

Private Sub PICC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ShowInstructions CInt(ToolType)
End Sub

Private Sub optTools_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'ToolType = Index
   ShowInstructions Index
   TempToolType = Index
End Sub

Private Sub optTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim k As Long
Dim svToolType As Long
Dim xc As Single, yc As Single

If ADRAW Then
   ' In case pressed mid-drawing
   optTools(Measure) = False
   optTools(Pick) = False
   Exit Sub
End If
   
   optTools(Measure) = False
   LineMeas.Visible = False
   optTools(Pick) = False
   optTools(Smooth1) = False
   optTools(Smooth2) = False
   optTools(Smooth4) = False
   ASELECTION = False
   If aSelRect Or aSelCirc Or aSelEllip Or aSelLasso Then _
         ASELECTION = True
         
   If Button = vbRightButton Then
      If Index <= Arrow Then
         ' Selector form
         frmToolOptions.Show vbModeless ' 0
      Else
         optTools(Index) = True
         ToolType = Index
         Select Case ToolType
         Case Pepper
            ImageStart_Enabler
            If ASELECTION Then
               cmdFile(3).Enabled = True  ' Save Selection
               MakeMask
               ReDim bMask(SSW, SSH)
               GetPICBytes picMask.Image, bMask(), SSW, SSH
            End If
            GetColor Button
            PepperEm CulNum
            DISPSAVE
         Case AText
            Unload frmZoom
            aZoom = False
            GetColor Button
            frmText.Show vbModal ' Topform only active
            If NSTOREXY = 0 Then
               optTools(AText).Value = False
               LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
            Else
               PIC_MouseUp 2, 0, 1, canvasH - 3
            End If
         End Select
      End If
   
   ElseIf Button = vbLeftButton Then
      ToolType = Index
      ShowInstructions Index
      optTools(Index) = True
      cmdStrip.Enabled = False
      
      Select Case Index
      Case AText
         Unload frmZoom
         aZoom = False
         GetColor Button
         frmText.Show vbModal ' Topform only active
         If NSTOREXY = 0 Then
            optTools(AText).Value = False
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
         Else
            PIC_MouseUp 1, 0, 1, canvasH - 3
         End If
      '------------------------------------
      Case SelR   ' Sel rect
         shpCirc.Visible = False
         aSelCirc = False
         shpEllip.Visible = False
         aSelEllip = False
         If NumLassoLines > 1 Then
            For k = 1 To NumLassoLines - 1
               Unload SL(k)
            Next k
            NumLassoLines = 1
         End If
         SL(0).Visible = False
      Case SelC   ' Sel circ
         shpRect.Visible = False
         aSelRect = False
         shpEllip.Visible = False
         aSelEllip = False
         If NumLassoLines > 1 Then
            For k = 1 To NumLassoLines - 1
               Unload SL(k)
            Next k
            NumLassoLines = 1
         End If
         SL(0).Visible = False
      Case SelE   ' Sel ellipse
         shpRect.Visible = False
         aSelRect = False
         shpCirc.Visible = False
         aSelCirc = False
         If NumLassoLines > 1 Then
            For k = 1 To NumLassoLines - 1
               Unload SL(k)
            Next k
            NumLassoLines = 1
         End If
         SL(0).Visible = False
      Case SelL   ' Sel Lasso
         shpRect.Visible = False
         aSelRect = False
         shpCirc.Visible = False
         aSelCirc = False
         shpEllip.Visible = False
         aSelEllip = False
         SL(0).Visible = True
         
      Case Desel  ' Deselect
         DoDeselect
      
      Case SCopyPaste To SCut
         MakeMask
         ReDim bMask(SSW, SSH)
         GetPICBytes picMask.Image, bMask(), SSW, SSH
         GetbPic
         DISPLAYpicPic
         If Index = SCut Then
            DISPLAY
            SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
            FixUndos           ' unless StopUndos = True
         End If
      Case SReflectLR
         MakeMask
         ReDim bMask(SSW, SSH)
         GetPICBytes picMask.Image, bMask(), SSW, SSH
         ReflectSelectLR
         DISPLAYpicPic
      Case SReflectUD
         MakeMask
         ReDim bMask(SSW, SSH)
         GetPICBytes picMask.Image, bMask(), SSW, SSH
         ReflectSelectUD
         DISPLAYpicPic
      Case SRotate   ' Circular selection only
         MakeMask
         ReDim bMask(SSW, SSH)
         GetPICBytes picMask.Image, bMask(), SSW, SSH
         GetbPic     ' Needed
         RotateEllbRect zangRotCSEL
         DISPLAYpicPic
         cmdStrip.Enabled = True
      Case SPaste
      
      Case SClear
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
         End If
         ClearEm
         DISPLAY
         SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
         FixUndos           ' unless StopUndos = True
         DoDeselect         ' to avoid a Paste
      
      ' Following depend on whether or not a selection
      Case Rot90
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
            GetbPic        ' Needed
         End If
         Rotator  ' NB Swaps canvasW & canvasH
         If ASELECTION Then DISPLAYpicPic
         If Not ASELECTION Then
            PIC.Width = canvasW
            PIC.Height = canvasH
         
         ElseIf aSelLasso Then
         ' Rotate SL() about selection center
            xc = SSX + SSW / 2
            yc = SSY + SSH / 2
            For k = 0 To NumLassoLines - 1
               ix = SL(k).x1
               SL(k).x1 = xc - (SL(k).y1 - yc)
               SL(k).y1 = yc + (ix - xc)
               ix = SL(k).x2
               SL(k).x2 = xc - (SL(k).y2 - yc)
               SL(k).y2 = yc + (ix - xc)
            Next k
            ix = XSMin
            iy = YSMin
            XSMin = 10000: YSMin = 10000
            XSMax = -10000: YSMax = -10000
            For k = 0 To NumLassoLines - 1
               With SL(k)
                  If .x1 < XSMin Then XSMin = .x1
                  If .y1 < YSMin Then YSMin = .y1
                  If .x1 > XSMax Then XSMax = .x1
                  If .y1 > YSMax Then YSMax = .y1
                  If .x2 < XSMin Then XSMin = .x2
                  If .y2 < YSMin Then YSMin = .y2
                  If .x2 > XSMax Then XSMax = .x2
                  If .y2 > YSMax Then YSMax = .y2
               End With
            Next k
            SSX = XSMin
            SSY = YSMin
            SSW = (XSMax - XSMin + 8) And &HFFFFFFFC
            SSH = (YSMax - YSMin + 8) And &HFFFFFFFC
            IncrX = ix - SSX
            IncrY = iy - SSY
            For k = 0 To NumLassoLines - 1
               SL(k).x1 = SL(k).x1 + IncrX
               SL(k).y1 = SL(k).y1 + IncrY
               SL(k).x2 = SL(k).x2 + IncrX
               SL(k).y2 = SL(k).y2 + IncrY
            Next k
            XSMin = XSMin + IncrX
            YSMin = YSMin + IncrY
            XSMax = XSMax + IncrX
            YSMax = YSMax + IncrY
            SSX = XSMin
            SSY = YSMin
            SSW = (XSMax - XSMin + 8) And &HFFFFFFFC
            SSH = (YSMax - YSMin + 8) And &HFFFFFFFC
         End If
         DISPSAVE
      
      Case Mix    ' Mix up
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
         End If
         MixEm
         DISPSAVE
      
      Case Thicken   ' Thicken
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
         End If
         ThickenPixels
         DISPSAVE
      
      Case Pepper   ' Pepper
         ImageStart_Enabler
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
         End If
         GetColor Button
         PepperEm CulNum
         DISPSAVE
      Case LRColor
         If ASELECTION Then
            MakeMask
            ReDim bMask(SSW, SSH)
            GetPICBytes picMask.Image, bMask(), SSW, SSH
         End If
         ColorReplacer
         DISPSAVE
      Case Measure
         If ADRAW Then Exit Sub
         fraMeas.Visible = True
         LineMeas.Visible = True
         For k = Brush To LRColor
            optTools(k).Value = False
         Next k
         optTools(Pick).Value = False
      Case Pick
         If ADRAW Then
            optTools(Pick).Enabled = False
            Exit Sub
         End If
         For k = Brush To Measure
            optTools(k).Value = False
         Next k
      
      Case Smooth1, Smooth2, Smooth4
         If ADRAW Then Exit Sub
         For k = Brush To Pick
            optTools(k).Value = False
         Next k
      
      End Select
      
      If Not ASELECTION Then
         For k = SCopyPaste To SClear
            optTools(k).Enabled = False
         Next k
         optTools(Desel).Enabled = True
         cmdFile(3).Enabled = False  ' Save Selection
         mnuSaveSelection.Enabled = False
         cmdPrint(1).Enabled = False   ' Print selection
         mnuPrintSelection.Enabled = False
      Else
         cmdFile(3).Enabled = True  ' Save Selection
         mnuSaveSelection.Enabled = True
         cmdPrint(1).Enabled = True
         mnuPrintSelection.Enabled = True
      End If
   End If
   FillLabInfos
   LabInfo(3) = ToolType  ' Test
End Sub

Private Sub MakeMask()
Dim k As Long
Dim xc As Single, yc As Single
Dim LCul As Long
   LCul = vbWhite
   ' Set up mask & Transfer to picMask
   With picMask
      .Width = SSW
      .Height = SSH
      .BackColor = 0
      .Picture = LoadPicture
      .Refresh
   End With
   
   If aSelRect Then
      picMask.Line (0, 0)-(SSW, SSH), LCul, B
   ElseIf aSelCirc Or aSelEllip Then
      ' Cirllipse
      ixc = SSX + SSW \ 2
      iyc = SSY + SSH \ 2
      xc = SSX + 2
      yc = SSY + 2
      EvalZradZratio xc, yc   'ixc,iyc public
      ixc = SSW / 2: iyc = SSH / 2
      picMask.Circle (ixc, iyc), zrad, LCul, , , zratio
   ElseIf aSelLasso Then
      For k = 0 To NumLassoLines - 1
         With SL(k)
            picMask.Line (.x1 - XSMin + 1, .y1 - YSMin + 1)-(.x2 - XSMin + 1, .y2 - YSMin + 1), LCul
         End With
      Next k
   End If
   picMask.Refresh
   
   ' Fill with FillColor = DrawColor at X,Y
   picMask.DrawStyle = vbSolid
   picMask.FillColor = LCul  'vbWhite
   picMask.FillStyle = vbFSSolid
   
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' color = APIC.Point(X, Y)
   
   ExtFloodFill picMask.hDC, 0, 0, picMask.Point(0, 0), FLOODFILLSURFACE
   
   picMask.Refresh
   
   ' Invert
   SetRect iREC, 0, 0, picMask.Width, picMask.Height
   InvertRect picMask.hDC, iREC
   picMask.Refresh
   ' Have mask culnum=255 surrounded by culnum=0   (eg black around red or white arounbd red)
   With picMask
      .DrawWidth = 1
      .FillStyle = vbFSTransparent
   End With
   picMask.Refresh
End Sub

Private Sub DoDeselect()
Dim k As Long
   shpRect.Visible = False
   aSelRect = False
   shpCirc.Visible = False
   aSelCirc = False
   shpEllip.Visible = False
   aSelEllip = False
   If NumLassoLines > 1 Then
      For k = 1 To NumLassoLines - 1
         Unload SL(k)
      Next k
      NumLassoLines = 1
   End If
   SL(0).Visible = False
   aSelLasso = False
   For k = SCopyPaste To SClear 'SPaste
      optTools(k).Enabled = False
   Next k
   ASELECTION = False
   cmdFile(3).Enabled = False
   mnuSaveSelection.Enabled = False
   cmdPrint(1).Enabled = False
   mnuPrintSelection.Enabled = False
End Sub

Private Sub DISPSAVE()
'   DISPLAY
   SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
   FixUndos           ' unless StopUndos = True
   NewNum = NewNum + 1
   DISPLAY
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Button As Integer
Dim x As Single, y As Single
Dim k As Long

   If KeyCode = 18 Then Exit Sub
   
   If KeyCode = 66 Or KeyCode = 98 Then      'B, b
      CommonSwapBW
      Exit Sub
   End If
   If KeyCode = 72 Or KeyCode = 104 Then     'H, h
      chkHairs.Value = 1 - chkHairs.Value
      chkHairs_MouseUp 1, 0, 0, 0
      Exit Sub
   End If
   
   MouseKeys KeyCode, Shift, Button, x, y
   
   If NewNum = 1 Then
      ' Only increment NewNum for LC/RC keys
      If KeyCode = 13 Or KeyCode = 8 Then
         NewNum = 2
         Exit Sub
      End If
   End If
   
   PIC.SetFocus
   DoEvents
   
   If NewNum > 1 Then
      mnuSave.Enabled = True
      cmdFile(2).Enabled = True
      mnuViews.Enabled = True
      mnuZoom.Enabled = True
      mnuTransforms.Enabled = True
   End If
   
   If NewNum > 1 Then
   If LCNum = 0 Then
   If KeyCode = 32 Then ' Spacebar allows copying of some shapes
      Select Case ToolType
      Case Brush
         If BrushType > Dot3 Then
            LCNum = 2
            PIC.DrawMode = 7
            ADRAW = True
            DISABLES
         End If
      Case Spray:
         LCNum = 1: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case ALine:
         LCNum = 2: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case PolyLine, CurvyLine:
         If CopyStartButton = vbLeftButton Then
            StartButton = CopyStartButton
            'LCNum = 2
            RCNum = 1
            AMoveAll = True
            PIC.DrawMode = 7
         ElseIf CopyStartButton = vbRightButton Then
            StartButton = CopyStartButton
            'RCNum = 2
            LCNum = 1
            AMoveAll = True
            PIC.DrawMode = 7
         End If
         ADRAW = True
         DISABLES
      Case Rectangle, Cirllipse:
         LCNum = 2: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case Cone, Tube, Bullet:
         LCNum = 3: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case Junction:
         LCNum = 2: PIC.DrawMode = 7
         For k = 1 To 12
            Getpx1py1 CLng(XT(k)), 0, CLng(YT(k)), 0
            XT(k) = px1
            YT(k) = py1
         Next k
         ADRAW = True
         DISABLES
      Case Arc To Radial:
         LCNum = 2: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case Tree:
         LCNum = 1: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case Arrow:
         LCNum = 2: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      Case AText:
         LCNum = 1: PIC.DrawMode = 7
         ADRAW = True
         DISABLES
      End Select
   End If
   End If
   End If
   
   ' OUT: Button 1 or 2, x,y
   GetExtras Me.BorderStyle
   ' OUT: Public ExtraBorder, ExtraHeight
   If KeyCode = 13 Then ' LC
      ' Get position on PIC picturebox
      ' +2 for PICC border
      x = x - Me.Left / STX - PICC.Left - PIC.Left - ExtraBorder + 2
      y = y - Me.Top / STY - PICC.Top - PIC.Top - ExtraHeight + 2
      PIC_MouseUp Button, Shift, x, y
      Sleep 400   ' Reduce repeat
   ElseIf KeyCode = 8 Then 'RC
      x = x - Me.Left / STX - PICC.Left - PIC.Left - ExtraBorder + 2
      y = y - Me.Top / STY - PICC.Top - PIC.Top - ExtraHeight + 2
      PIC_MouseUp Button, Shift, x, y
      Sleep 400   ' Reduce repeat
   End If
End Sub

Private Sub Form_Resize()
   'RightGap = Me.Width \ STX - PICC.Left - PICC.Width
   'BottomGap = Me.Height \ STY - PICC.Top - PICC.Height
   'fraInfoBottomGap = Me.Height \ STY - fraInfo.Top - fraInfo.Height
   
   If aResize Then
      If WindowState <> vbMinimized Then
         If WindowState <> vbMaximized Then
            If Me.Width < OrgFW Then Me.Width = OrgFW
            If Me.Height < OrgFH Then Me.Height = OrgFH
         End If
         PICC.Width = Me.Width \ STX - PICC.Left - RightGap
         PICCW = PICC.Width
         PICC.Height = Me.Height \ STY - PICC.Top - BottomGap
         PICCH = PICC.Height
         FixScrollbars PICC, PIC, HS, VS
         fraInfo.Top = Me.Height \ STY - fraInfo.Height - fraInfoBottomGap
         fraInfo.Width = Me.Width \ STX - 16
         'LabInfo(3).Width = (fraInfo.Width - 16) * STX - (LabInfo(0).Width + LabInfo(1).Width + LabInfo(2).Width)
         TileForm1
         PICC.Refresh
         picPic.Left = PICC.Left + PICC.Width + 100
         picPic.Top = 176
      End If
   End If
End Sub

Private Sub TileForm1()
   For iy = 0 To 32 Step 8
   For ix = 0 To 32 Step 8
      BitBlt Me.hDC, ix, iy, 8, 8, Me.hDC, 0, 0, vbSrcCopy
   Next ix
   Next iy
   For iy = 0 To Me.Height / STY Step 32
   For ix = 0 To Me.Width / STX Step 32
      BitBlt Me.hDC, ix, iy, 32, 32, Me.hDC, 0, 0, vbSrcCopy
   Next ix
   Next iy
   Me.Refresh
   BitBlt PICC.hDC, 8, 8, 8, 8, PICC.hDC, 0, 0, vbSrcCopy
   For iy = 0 To PICC.Height Step 16
   For ix = 0 To PICC.Width Step 16
      BitBlt PICC.hDC, ix, iy, 16, 16, PICC.hDC, 0, 0, vbSrcCopy
   Next ix
   Next iy
End Sub

Private Sub cmdTile_Click(Index As Integer)
   Select Case Index
   Case 0
      Set Form1.Picture = LoadResPicture("GOLD", vbResBitmap)
      ForeColor = ForeColor Xor vbYellow
   Case 1
      Set Form1.Picture = LoadResPicture("GREEN", vbResBitmap)
      ForeColor = ForeColor Xor vbGreen
   Case 2
      Set Form1.Picture = LoadResPicture("BLUE", vbResBitmap)
      ForeColor = ForeColor Xor vbBlue
   Case 3
      Set Form1.Picture = LoadResPicture("WHITE", vbResBitmap)
      ForeColor = ForeColor Xor vbWhite
   Case 4
      Set Form1.Picture = LoadResPicture("GREY", vbResBitmap)
      ForeColor = ForeColor Xor vbCyan
   Case 5
      Set Form1.Picture = LoadResPicture("BLACK", vbResBitmap)
      ForeColor = ForeColor Xor vbBlack
   Case 6
      Set Form1.Picture = LoadResPicture("RED", vbResBitmap)
      ForeColor = ForeColor Xor vbRed
   End Select
   TileForm1
   CurrentX = 10: CurrentY = 520
   Print " APaint8 ";
   CurrentX = 10: CurrentY = 540
   Print " by ";
   CurrentX = 10: CurrentY = 560
   Print " Robert Rayment ";
   CurrentX = 10: CurrentY = 580
   Print " 2004 ";
End Sub

'#### Menus #######################################################

Private Sub cmdFile_Click(Index As Integer)
   If ADRAW Then Exit Sub
   Select Case Index
   Case 0: mnuNew_Click
   Case 1: mnuBrowser_Click
   Case 2: mnuSave_Click
   Case 3: mnuSaveSelection_Click
   End Select
End Sub

Private Sub cmdEdit_Click(Index As Integer)
   If ADRAW Then Exit Sub
   
   Select Case Index
   Case 0: mnuUndo_Click
   Case 1: mnuRedo_Click
   Case 2: mnuClearCurrentView_Click
   Case 3: mnuDeleteCurrentView_Click
   Case 4: mnuDeleteAbove_Click
   Case 5: mnuCollapse_Click
   Case 6: mnuAddViewAbove_Click
   Case 7: mnuSwapViews_Click
   Case 8: mnuToggleUndos_Click
   End Select
End Sub

Private Sub cmdSwapBW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   CommonSwapBW
End Sub

Public Sub CommonSwapBW()
Dim k As Long
   If ADRAW Then Exit Sub
   
   palRed(0) = 255 - palRed(0)
   palGreen(0) = 255 - palGreen(0)
   palBlue(0) = 255 - palBlue(0)
   palRed(1) = 255 - palRed(0)
   palGreen(1) = 255 - palGreen(0)
   palBlue(1) = 255 - palBlue(0)
   ReDim CulRGB(0 To 255), CulBGR(0 To 255)
   For k = 0 To 255
      CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
      CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
   Next k
   ShowPalette
   CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   DISPLAY
End Sub

Private Sub mnuSave_Click()
Dim Title$, Filt$, InDir$
Dim Ext$
Dim FIndex As Long

If ADRAW Then Exit Sub

   If aVIEWS Then
      frmViewsLeft = frmViews.Left
      frmViewsTop = frmViews.Top
   End If
   Unload frmViews
   'aVIEWS = False   ' Leave, if True will reshow frmViews

Set CommonDialog1 = New OSDialog
   Title$ = "Save As 8bpp BMP or GIF"
   'Filt$ = "jpg|*.jpg|gif|*.gif|png|*.png"
   Filt$ = "Save BMP (*.bmp)|*.bmp|Save GIF (*.gif)|*.gif"
   InDir$ = SavePathSpec$
   Ext$ = "bmp"
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, Ext$, Me.hWnd, FIndex
Set CommonDialog1 = Nothing

   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   
   If Len(FileSpec$) <> 0 Then
      SavePathSpec$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))
      If FIndex = 1 Then Ext$ = ".bmp" Else Ext$ = ".gif"
      If Ext$ = ".gif" Then
         FixExtension FileSpec$, ".gif"
         MSaveGIF FileSpec$, bArray(), CInt(canvasW), CInt(canvasH), CulRGB(), True
      Else
         FixExtension FileSpec$, ".bmp"
         MSaveBMP FileSpec$, bArray(), canvasW, canvasH, CulBGR()
      End If
   End If
   FillLabInfos
   If aVIEWS Then frmViews.Show 0
End Sub

Private Sub mnuSaveSelection_Click()
Dim Title$, Filt$, InDir$
Dim Ext$
Dim FIndex As Long

If ADRAW Then Exit Sub
If Not ASELECTION Then Exit Sub
         
   If aVIEWS Then
      frmViewsLeft = frmViews.Left
      frmViewsTop = frmViews.Top
   End If
   Unload frmViews
   'aVIEWS = False   ' Leave, if True will reshow frmViews

   MakeMask
   ReDim bMask(SSW, SSH)
   GetPICBytes picMask.Image, bMask(), SSW, SSH
   GetbPic
   DISPLAYpicPic

Set CommonDialog1 = New OSDialog
   Title$ = "Save Selection As 8bpp BMP or GIF"
   Filt$ = "Save BMP (*.bmp)|*.bmp|Save GIF (*.gif)|*.gif"
   InDir$ = SavePathSpec$
   Ext$ = "bmp"
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, Ext$, Me.hWnd, FIndex
Set CommonDialog1 = Nothing

   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   
   If Len(FileSpec$) <> 0 Then
      SavePathSpec$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))
      If FIndex = 1 Then Ext$ = ".bmp" Else Ext$ = ".gif"
'      Ext$ = LCase$(Mid$(FileSpec$, InStrRev(FileSpec$, ".")))
      If Ext$ = ".gif" Then
         FixExtension FileSpec$, ".gif"
         MSaveGIF FileSpec$, bPic(), CInt(SSW), CInt(SSH), CulRGB(), True
      Else
         FixExtension FileSpec$, ".bmp"
         MSaveBMP FileSpec$, bPic(), SSW, SSH, CulBGR()
      End If
   End If
   FillLabInfos
   If aVIEWS Then frmViews.Show 0
End Sub

Private Sub mnuUndo_Click()
If ADRAW Then Exit Sub
   If UndoNum > 1 Then
      UndoNum = UndoNum - 1
      CommonUndo
   End If
End Sub

Private Sub mnuRedo_Click()
If ADRAW Then Exit Sub
   If UndoNum < TopUndoNum Then
      UndoNum = UndoNum + 1
      RESTORE_Image
      FixUndos
      FillLabInfos
      ShowPalette
      PIC.Picture = LoadPicture
      LCNum = -1
      DISPLAY
      BackUpRGB() = CulRGB()
   End If
End Sub
Public Sub CommonUndo()
' Called from frmViews
   If UndoNum <= TopUndoNum Then
         RESTORE_Image
         FixUndos
         FillLabInfos
         ShowPalette
         PIC.Picture = LoadPicture
         LCNum = -1
         DISPLAY
         BackUpRGB() = CulRGB()
   End If
End Sub

Private Sub mnuClearCurrentView_Click()
If ADRAW Then Exit Sub
   
   If MsgBox("Continue ?", vbQuestion + vbYesNo, "Clear Current View") = vbYes Then
      ClearUndo ' UndoNum
      ReDim bArray(1 To canvasW, 1 To canvasH)
      PIC.Picture = LoadPicture
      LCNum = -1
      DISPLAY
   End If
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Private Sub mnuDeleteCurrentView_Click()
If ADRAW Then Exit Sub
   CommonDelete
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Public Sub CommonDelete()
' Delete current view
   If MsgBox("Continue ?", vbQuestion + vbYesNo, "Delete Current View") = vbYes Then
      If UndoNum > 1 Then
         DeleteCurrentView
         FixUndos
         FillLabInfos
         PIC.Picture = LoadPicture
         LCNum = -1
         aMNUACTION = True
         DISPLAY
      End If
   End If
End Sub

Private Sub mnuDeleteAbove_Click()
If ADRAW Then Exit Sub
   
   If MsgBox("Continue ?", vbQuestion + vbYesNo, "Delete All Views Above....") = vbYes Then
      If UndoNum >= 1 And TopUndoNum <= MaxUndos Then
         TopUndoNum = UndoNum
         FixUndos
         FillLabInfos
         PIC.Picture = LoadPicture
         LCNum = -1
         aMNUACTION = True
         DISPLAY
      End If
   End If
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Private Sub mnuCollapse_Click()
Dim k As Long
If ADRAW Then Exit Sub
   
   ' Keeps just Undo 1 and Undo TopUndoNum
   If MsgBox("Continue ?", vbQuestion + vbYesNo, "Collapse backups to first & last view....") = vbYes Then
      Collapse  ' Sets TopUndoNum = 2 & UndoNum = 2
      ReDim bArray(1 To canvasW, 1 To canvasH)
      bArray() = bUndo2()
      'DISPLAY
      FixUndos
      FillLabInfos
      PIC.Picture = LoadPicture
      LCNum = -1
      aMNUACTION = True
      DISPLAY
      BackUpRGB() = CulRGB()
   End If
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Private Sub mnuAddViewAbove_Click()
If ADRAW Then Exit Sub
   CommonAdd
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Public Sub CommonAdd()
Dim res As Long

   res = MsgBox("   NB The current palette will be used" & vbCr & vbCr _
              & "   Yes   -   OVERWRITE" & vbCr _
              & "    No   -    ADD TO BACKGROUND" & vbCr _
              & "              or   CANCEL", _
                  vbYesNoCancel + vbSystemModal, "Overwrite current with view above       ")
   
   If res = vbCancel Then Exit Sub
   
   ReDim bDummy(canvasW, canvasH)
   bDummy() = bArray()  '- Save Current view
   Select Case res
      Case vbYes: aOverWrite = True
       AddView
       PIC.Picture = LoadPicture
       LCNum = -1
       DISPLAY
      Case vbNo: aOverWrite = False
       AddView
       PIC.Picture = LoadPicture
       LCNum = -1
       aMNUACTION = True
       DISPLAY
      Case vbCancel: 'Cancel
   End Select
   ' Acceptable?
   res = MsgBox("Acceptable?", vbYesNo + vbQuestion, "Preview adding images")
   If res = vbNo Then
      bArray() = bDummy()           '- Restore saved view
      FillUndoNWithbArray UndoNum   '- Restore current UndoNum
      PIC.Picture = LoadPicture
      LCNum = -1
      DISPLAY
   End If
End Sub

Private Sub mnuSwapViews_Click()
If ADRAW Then Exit Sub
   
   If UndoNum > 1 Then
      CommonSwap
   End If
   ' Offset cursor to avoid click thru
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
End Sub

Public Sub CommonSwap()
      If MsgBox("Continue ?", vbQuestion + vbYesNo, "Swap with View below....") = vbYes Then
         SwapViews
         FixUndos
         FillLabInfos
         PIC.Picture = LoadPicture
         LCNum = -1
         aMNUACTION = True
         DISPLAY
         BackUpRGB() = CulRGB()
      End If
End Sub

Private Sub mnuToggleUndos_Click()
If ADRAW Then Exit Sub
 
 StopUndos = Not StopUndos
 If StopUndos Then
   cmdEdit(8).Picture = cmdStopUndos(0).Picture
   mnuClearCurrentView.Enabled = True
   mnuDeleteCurrentView.Enabled = False
   mnuDeleteAbove.Enabled = False
   mnuCollapse.Enabled = False
   cmdEdit(2).Enabled = True
   cmdEdit(3).Enabled = mnuDeleteCurrentView.Enabled
   cmdEdit(4).Enabled = mnuDeleteAbove.Enabled
   cmdEdit(5).Enabled = mnuCollapse.Enabled
 Else
    cmdEdit(8).Picture = cmdStopUndos(1).Picture
    FixUndos
 End If
 PIC.SetFocus
End Sub

Private Sub FixUndos()
   If UndoNum = TopUndoNum Then
      mnuRedo.Enabled = False
      Select Case UndoNum
      Case 0   ' At Start up
         mnuUndo.Enabled = False
         If Not StopUndos Then
            mnuClearCurrentView.Enabled = False
            mnuDeleteCurrentView.Enabled = False
            mnuDeleteAbove.Enabled = False
            mnuCollapse.Enabled = False
            mnuAddViewAbove.Enabled = False
            mnuSwapViews.Enabled = False
         End If
      Case 1   ' First image
         If Not StopUndos Then
            mnuClearCurrentView.Enabled = True
            mnuDeleteCurrentView.Enabled = False
            mnuDeleteAbove.Enabled = False
            mnuCollapse.Enabled = False
            mnuAddViewAbove.Enabled = False
            mnuSwapViews.Enabled = False
         End If
      Case 2
         mnuUndo.Enabled = True
         If Not StopUndos Then
            mnuDeleteCurrentView.Enabled = True
            mnuDeleteAbove.Enabled = False
            mnuCollapse.Enabled = False
            mnuAddViewAbove.Enabled = False
            mnuSwapViews.Enabled = True
         End If
      Case Else
         mnuUndo.Enabled = True
         If Not StopUndos Then
            mnuDeleteCurrentView.Enabled = True
            mnuDeleteAbove.Enabled = False
            mnuCollapse.Enabled = True
            mnuAddViewAbove.Enabled = False
            mnuSwapViews.Enabled = True
         End If
      End Select
    
    Else ' UndoNum < TopUndoNum
      
      Select Case UndoNum
      Case 1
         mnuSwapViews.Enabled = False
         If TopUndoNum >= 2 Then
            mnuUndo.Enabled = False
            mnuRedo.Enabled = True
            If Not StopUndos Then
               mnuDeleteCurrentView.Enabled = False
               mnuDeleteAbove.Enabled = True
               mnuCollapse.Enabled = True
               mnuAddViewAbove.Enabled = True
            End If
         End If
      Case Else   ' UndoNum >1 & < TopUndoNum
         mnuUndo.Enabled = True
         mnuRedo.Enabled = True
         If Not StopUndos Then
            mnuDeleteCurrentView.Enabled = True
            mnuDeleteAbove.Enabled = True
            mnuCollapse.Enabled = True
            mnuAddViewAbove.Enabled = True
            mnuSwapViews.Enabled = True
         End If
      End Select
    End If
    cmdEdit(0).Enabled = mnuUndo.Enabled
    cmdEdit(1).Enabled = mnuRedo.Enabled
    If Not StopUndos Then
      cmdEdit(2).Enabled = mnuClearCurrentView.Enabled
      cmdEdit(3).Enabled = mnuDeleteCurrentView.Enabled
      cmdEdit(4).Enabled = mnuDeleteAbove.Enabled
      cmdEdit(5).Enabled = mnuCollapse.Enabled
      cmdEdit(6).Enabled = mnuAddViewAbove.Enabled
      cmdEdit(7).Enabled = mnuSwapViews.Enabled
    End If
'LabInfo(3) = UndoNum
End Sub

Private Sub mnuBrowser_Click()
'Public FileSpec$
Dim OldFileSpec$
Dim k As Long
Dim c0 As Long
Dim c1 As Long

On Error GoTo BrowseErr

If ADRAW Then Exit Sub
   
   OldFileSpec$ = FileSpec$
   
   Unload frmZoom
   aZoom = False
   Unload frmToolOptions
   Unload frmPalette
   Unload frmHelp
   If aVIEWS Then
      frmViewsLeft = frmViews.Left
      frmViewsTop = frmViews.Top
   End If
   Unload frmViews
   'aVIEWS = False   ' Leave, if True will reshow frmViews
   DoEvents
   
   frmBrowse.Show vbModal '1  ' returns image FileSpec$

   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   
   If aVIEWS Then frmViews.Show 0
   
   If Len(FileSpec$) <> 0 Then
      NewNum = 2
      
      PIC.Picture = LoadPicture(FileSpec$)
      PIC.Width = canvasW
      PIC.Height = canvasH
      
      ' Default color numbers 0 & 1
      c0 = RGB(palRed(0), palGreen(0), palBlue(0))
      c1 = RGB(palRed(1), palGreen(1), palBlue(1))
      
      ' Set color nums 0 & 1 if not B/W or W/B
      If (c0 = 0 And c1 = vbWhite) Or _
         (c0 = vbWhite And c1 = 0) Then
      Else
         palRed(0) = 0
         palGreen(0) = 0
         palBlue(0) = 0
         CulRGB(0) = 0
         CulBGR(0) = 0
         
         palRed(1) = 255
         palGreen(1) = 255
         palBlue(1) = 255
         CulRGB(1) = vbWhite
         CulBGR(1) = vbWhite
      End If
      
      ' Backup Image's Palette
      BackUpRGB() = CulRGB()
      If UndoNum = 0 Then
         TopUndoNum = 0
         ReDimUndos
      Else  ' Load into current view
         k = MsgBox("Reset view stack ?" & vbCr & vbCr & "NB If No, then new palette" & vbCr & "can affect added views", _
                    vbYesNo + vbQuestion + vbSystemModal, "Loading image into a view")
         
         SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
         If aVIEWS Then LCNum = -1
         If k = vbNo Then
            SAVE_CurrentImage
            FixUndos
         Else
            UndoNum = 0
            TopUndoNum = 0
            ReDimUndos
            StopUndos = False
            cmdEdit(8).Picture = cmdStopUndos(1).Picture
            cmdStopUndos(0).Visible = False
            cmdStopUndos(1).Visible = False
         ''''''''''''''''''''''''''''
            UndoNum = 0
            TopUndoNum = 0
            SAVE_CurrentImage
            FixUndos
         ''''''''''''''
         End If
      End If
      
      If UndoNum > 0 Then  ' NB Palette will be for last image loaded !!
         FillUndoNWithbArray UndoNum
      End If
      
      ShowPalette
      
      ' Set default L/R selected colors
      picPAL_MouseUp 2, 0, 4, 4     ' Right button color
      picPAL_MouseUp 1, 0, 8, 4     ' Left button color
      
      SelLeftCulNum = 1
      SelRightCulNum = 0
      LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
      LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
      
      PIC.Picture = LoadPicture
      DISPLAY
      
      xprev = 1
      yprev = 1
      If UndoNum = 0 Then
         StopUndos = True
         mnuToggleUndos_Click
         SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
         FixUndos          ' unless StopUndos = True
      End If
      
      ImageStart_Enabler
      
      CopyFileSpec$(UndoNum) = FileSpec$
      FillLabInfos
   Else
      FileSpec$ = OldFileSpec$
   End If
   Exit Sub
'==========
BrowseErr:
Close
On Error GoTo 0
MsgBox "Browser error", vbInformation, "Browser"
End Sub

Private Sub mnuZoom_Click()
   aZoom = True
   frmZoom.Show vbModeless ' 0
End Sub

'#### END Menus #######################################################

Private Sub FillLabInfos()
   FileSpec$ = CopyFileSpec$(UndoNum)
   LabInfo(0) = FileSpec$
   LabInfo(1) = "W,H=" & Str$(canvasW) & "," & Str$(canvasH)
   If UndoNum < MaxUndos Then
      LabInfo(2) = Str$(UndoNum) & " of" & Str$(TopUndoNum)
   Else
      LabInfo(2) = Str$(UndoNum) & " of" & Str$(TopUndoNum) & " !!!"
   End If
End Sub

Private Sub DISPLAY()
Dim bS As BITMAPINFO
   
   If aVIEWS Then
      
      
      ' Fix to prevent Views until
      ' dot drawing & smoothing finished
      Select Case ToolType
      Case Brush
         Select Case BrushType
         Case Dot1, Dot2, Dot3
            If LCNum = -1 Then
               If TopUndoNum = MaxUndos Or aMNUACTION Then
                  DISPLAY_ALL_VIEWS
                  aMNUACTION = False
               Else
                  DISPLAY_VIEW UndoNum
               End If
            End If
         Case Else   ' Other Brush types
            If TopUndoNum = MaxUndos Or aMNUACTION Then
               DISPLAY_ALL_VIEWS
               aMNUACTION = False
            Else
               DISPLAY_VIEW UndoNum
            End If
         End Select
      Case Smooth1, Smooth2, Smooth4
            If LCNum = -1 Then
               If TopUndoNum = MaxUndos Or aMNUACTION Then
                  DISPLAY_ALL_VIEWS
                  aMNUACTION = False
               Else
                  DISPLAY_VIEW UndoNum
               End If
            End If
      Case Else   ' Other ToolTypes
         If TopUndoNum = MaxUndos Or aMNUACTION Then
            DISPLAY_ALL_VIEWS
            aMNUACTION = False
         Else
            DISPLAY_VIEW UndoNum
         End If
      End Select
   End If
   
   ' Set up palette
   CopyMemory bS.Colors(0), CulBGR(0), 1024
   PIC.Width = canvasW     ' Always multiple of 4 !!
   PIC.Height = canvasH
   
   With bS.bmi
      .biSize = 40
      .biwidth = canvasW
      .biheight = canvasH
      .biPlanes = 1
      .biBitCount = 8
      .biSizeImage = canvasW * canvasH
   End With
   DoEvents
   
   If SetDIBitsToDevice(PIC.hDC, 0, 0, canvasW, canvasH, _
   0, 0, 0, canvasH, bArray(1, 1), bS, DIB_RGB_COLORS) = 0 Then
      MsgBox "DISPLAY ERROR", vbCritical, "Display"
      End
   End If
   
   PIC.Refresh
   PIC.SetFocus
   
   FixScrollbars PICC, PIC, HS, VS
   
   LCNum = 0: RCNum = 0 ' For any L/RCNums = -1

End Sub

Private Sub DISPLAYpicPic()
' Show bPic
''Redim bPic(SSW,SSH)
Dim bS As BITMAPINFO
Dim BytesPerScanLine As Long
'picPic.Visible = True
   
   ' Set up palette
   CopyMemory bS.Colors(0), CulBGR(0), 1024
   picPic.Width = SSW
   picPic.Height = SSH
   picPic.Picture = LoadPicture
   BytesPerScanLine = (SSW + 3) And &HFFFFFFFC  'OK gives mod 4 byte width
   picPic.Refresh
   
   With bS.bmi
      .biSize = 40
      .biwidth = picPic.Width
      .biheight = picPic.Height
      .biPlanes = 1
      .biBitCount = 8
      .biCompression = 0
      .biSizeImage = BytesPerScanLine * Abs(SSH)   ' but doesn't work??
   End With

   If SetDIBitsToDevice(picPic.hDC, 0, 0, SSW, SSH, _
   0, 0, 0, SSH, bPic(1, 1), bS, DIB_RGB_COLORS) = 0 Then
      MsgBox "DISPLAY picMask ERROR", vbCritical, "Display"
      End
   End If
   picPic.Refresh
End Sub

'#### COLOR SELECTION ###############################################

Private Sub cmdStartPalette_Click(Index As Integer)
' Inside fraPAL
Dim k As Long
If ADRAW Then Exit Sub
   
   Select Case Index
   Case 0: ' Make or load a palette
      Unload frmZoom
      aZoom = False
      Unload frmToolOptions
      If aVIEWS Then
         frmViewsLeft = frmViews.Left
         frmViewsTop = frmViews.Top
      End If
      Unload frmViews
      'aVIEWS = False   ' Leave, if True will reshow frmViews

      frmPalette.Show vbModal '1
      ' Offset cursor to avoid click thru
      SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   
      If aVIEWS Then frmViews.Show 0
   
   Case 1:  GreyedPalette         ' Greyed
   Case 2:  BandedPalette         ' 64 Banded
   Case 3:  ShortBandedPalette    ' 32 Banded
   Case 4:  CenteredPal           ' 16 center banded
   Case 5   ' Invert
      For k = 0 To 255
         palRed(k) = 255 - palRed(k)
         palGreen(k) = 255 - palGreen(k)
         palBlue(k) = 255 - palBlue(k)
         CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
   End Select
   AdjustPalette
   ' Copy to current UndoNum
   CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   DISPLAY
End Sub

Private Sub cmdDefPalette_Click()
Dim k As Long
   For k = 0 To 255
     palRed(k) = (DefaultRGB(k) And &HFF&)
     palGreen(k) = (DefaultRGB(k) And &HFF00&) / &H100&
     palBlue(k) = (DefaultRGB(k) And &HFF0000) / &H10000
   Next k
   AdjustPalette
   ' Copy to current UndoNum
   CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   DISPLAY
End Sub

Private Sub AdjustPalette()
Dim c0 As Long
Dim c1 As Long
Dim k As Long
   ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
   ' Default color numbers 0 & 1
   c0 = RGB(palRed(0), palGreen(0), palBlue(0))
   c1 = RGB(palRed(1), palGreen(1), palBlue(1))
   
   ' Set color nums 0 & 1 if not B/W or W/B
   If (c0 = 0 And c1 = vbWhite) Or _
      (c0 = vbWhite And c1 = 0) Then
   Else
      palRed(0) = 0
      palGreen(0) = 0
      palBlue(0) = 0
      palRed(1) = 255
      palGreen(1) = 255
      palBlue(1) = 255
   End If
   ReDim CulRGB(0 To 255), CulBGR(0 To 255)
   For k = 0 To 255
      CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
      CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
   Next k
   ' Set default L/R selected colors
   picPAL_MouseUp 2, 0, 4, 4     ' Right button color
   picPAL_MouseUp 1, 0, 8, 4     ' Left button color
   SelLeftCulNum = 1
   SelRightCulNum = 0
   LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
   LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
   ShowPalette
End Sub

Private Sub LabSelCul_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cnn As Long
   If Index = 0 Or Index = 1 Then
      ' Show color under cursor
      LabSelCul(2).BackColor = LabSelCul(Index).BackColor
      If Index = 0 Then ' Show color number
         cnn = SelLeftCulNum
         LabSelCul(3) = LTrim$(Str$(SelLeftCulNum))
      Else
         cnn = SelRightCulNum
         LabSelCul(3) = LTrim$(Str$(SelRightCulNum))
      End If
      ' Show RGB
      LngToRGB LabSelCul(Index).BackColor
      LabSelCul(4) = Str$(palRed(cnn)) & "," & Str$(palGreen(cnn)) & "," & Str$(palBlue(cnn))
   End If
End Sub

Private Sub picPAL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Inside fraPAL
Dim cn As Long
   If x >= 0 And x <= picPAL.Width \ STX Then
   If y >= 0 And y <= picPAL.Height \ STY Then
      cn = (y \ 8) * 8 + x \ 8
      If cn < 0 Then cn = 0
      If cn > 255 Then cn = 255
      LabSelCul(2).BackColor = CulRGB(cn)
      LabSelCul(2).Refresh
      LabSelCul(3) = Str$(cn)
      LngToRGB CulRGB(cn)
      LabSelCul(4) = Str$(palRed(cn)) & "," & Str$(palGreen(cn)) & "," & Str$(palBlue(cn))
   End If
   End If
   LabXY = Str$(x) & Str$(y)
End Sub

Private Sub picPAL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Inside fraPAL
Dim cn As Long
LabXY = Str$(x) & Str$(y)

   If x >= 0 And x <= (picPAL.Width \ STX) Then
   If y >= 0 And y <= (picPAL.Height \ STY) Then
      cn = (y \ 8) * 8 + x \ 8
      If cn < 0 Then cn = 0
      If cn > 255 Then cn = 255
      If Button = vbLeftButton Then
         LabSelCul(0).BackColor = CulRGB(cn)
         SelLeftCulNum = cn
         LabSelCul(0).Refresh
         shpPAL(0).Left = (x \ 8) * 8
         shpPAL(0).Top = (y \ 8) * 8
      ElseIf Button = vbRightButton Then
         LabSelCul(1).BackColor = CulRGB(cn)
         SelRightCulNum = cn
         LabSelCul(1).Refresh
         shpPAL(1).Left = (x \ 8) * 8
         shpPAL(1).Top = (y \ 8) * 8
      End If
      LabSelCul(3) = Str$(cn)
      LabSelCul(4) = Str$(palRed(cn)) & "," & Str$(palGreen(cn)) & "," & Str$(palBlue(cn))
   End If
   End If
End Sub

Private Sub ShowPalette()
' Public ix,iy
Dim k As Long
   For k = 0 To 255
    ix = 1 + (k Mod 8) * 8
    iy = 1 + (k \ 8) * 8
    picPAL.Line (ix, iy)-(ix + 5, iy + 5), CulRGB(k), BF
   Next k
   picPAL.Refresh
End Sub

Private Sub cmdCPal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim k As Long
Dim bRO As Byte
Dim bGO As Byte
Dim bBO As Byte
   Select Case Index
   Case 0   ' Brighten
      For k = 2 To 255
         If palRed(k) < 251 Then palRed(k) = palRed(k) + 4
         If palGreen(k) < 251 Then palGreen(k) = palGreen(k) + 4
         If palBlue(k) < 251 Then palBlue(k) = palBlue(k) + 4
      Next k
      'ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
      For k = 0 To 255
         CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
      CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   Case 1   ' Darken
      For k = 2 To 255
         If palRed(k) > 4 Then palRed(k) = palRed(k) - 4
         If palGreen(k) > 4 Then palGreen(k) = palGreen(k) - 4
         If palBlue(k) > 4 Then palBlue(k) = palBlue(k) - 4
      Next k
      'ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
      For k = 0 To 255
         CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
      CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   Case 2   ' Reset
      ' Restore Palette
      CulRGB() = BackUpRGB()
      For k = 0 To 255
         palRed(k) = CulRGB(k) And &HFF&
         palGreen(k) = (CulRGB(k) And &HFF00&) / &H100&
         palBlue(k) = (CulRGB(k) And &HFF0000) / &H10000
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
      CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   Case 3   ' Rotate palette
      If Button = vbLeftButton Then
            bRO = palRed(2)
            bGO = palGreen(2)
            bBO = palBlue(2)
            For k = 3 To 255
                palRed(k - 1) = palRed(k)
                palGreen(k - 1) = palGreen(k)
                palBlue(k - 1) = palBlue(k)
            Next k
            palRed(255) = bRO
            palGreen(255) = bGO
            palBlue(255) = bBO
      ElseIf Button = vbRightButton Then
            bRO = palRed(255)
            bGO = palGreen(255)
            bBO = palBlue(255)
            For k = 255 To 3 Step -1
                palRed(k) = palRed(k - 1)
                palGreen(k) = palGreen(k - 1)
                palBlue(k) = palBlue(k - 1)
            Next k
            palRed(2) = bRO
            palGreen(2) = bGO
            palBlue(2) = bBO
      End If
      ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
      For k = 0 To 255
         CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
      CopyMemory CopyRGB(0, UndoNum), CulRGB(0), 1024
   End Select
   ShowPalette
   LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
   LabSelCul(0).Refresh
   LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
   LabSelCul(1).Refresh
   DISPLAY
End Sub

'#### END COLOR SELECTION ###############################################

'#### PIC CURSOR ACTION #################################################

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NN As Long

   AMouseDown = True
   
   If NewNum > 0 Then NewNum = 2
      
   Select Case ToolType
   Case SelR, SelC, SelE, SelL, Desel
   Case Measure
      LineMeas.Visible = True
      LineMeas.x1 = x
      LineMeas.x2 = x
      LineMeas.y1 = y
      LineMeas.y2 = y
      
   Case Else
      If Not mnuSave.Enabled Then
         ImageStart_Enabler
      End If
      If ASELECTION Then
         mnuSaveSelection.Enabled = True
         cmdFile(3).Enabled = True
         mnuPrintSelection.Enabled = True
         cmdPrint(1).Enabled = True
      End If
   End Select
   
   PIC.SetFocus
   
   xprev = x
   yprev = y

   ' Adjust cords to bArray()
   Getpx1py1 CSng(x), 0, CSng(y), 0
   
   ShowInfoAtCursor px1, py1
   
   Select Case ToolType
   Case Brush
      Select Case BrushType
      Case Dot1, Dot2, Dot3
         GetColor Button
         SetDot px1, py1, CulNum, BrushType
         DISPLAY
      End Select
   Case Smooth1, Smooth2, Smooth4
   End Select
End Sub


Private Sub picPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PIC_MouseMove 0, 0, CSng(picPic.Left - (PICC.Left + PIC.Left) + x), CSng(picPic.Top - (PICC.Top + PIC.Top) + y)
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NN As Long
Dim x1 As Single, y1 As Single

'Public shw As Long, shh As Long
'Public shx As Long, shy As Long

' Rest all variables public
   
   PIC.SetFocus
   
   xprev = x
   yprev = y
      
   ' Adjust ccords to bArray()
   Getpx1py1 CSng(x), 0, CSng(y), 0
   
   ShowInfoAtCursor px1, py1
   If ToolType <> TempToolType Then
      ShowInstructions CInt(ToolType)
      TempToolType = ToolType
   End If
   If aHairs Then
      Line2.Visible = True
      Line3.Visible = True
      Line2.x1 = x
      Line2.x2 = x
      Line2.y1 = 0
      Line2.y2 = PIC.Height
      Line3.x1 = 0
      Line3.x2 = PIC.Width
      Line3.y1 = y
      Line3.y2 = y
   End If

   Select Case ToolType
   Case Brush
      Select Case BrushType
      Case Dot1, Dot2, Dot3
         If AMouseDown Then
            SetDot px1, py1, CulNum, BrushType
            DISPLAY
         End If
      Case FreeDraw1, FreeDraw2, FreeDraw3
         If Button = 0 Then MoveFreeDraws PIC, Button, x, y
      Case BRibbon1, BRibbon2, BRibbon3, FRibbon1, FRibbon2, FRibbon3   ' \ /
         If Button = 0 Then MoveRibbons PIC, Button, x, y
      End Select
   
   Case Measure
      'Public zarrang As Single
      'Public zarrlen As Single

      If AMouseDown Then
         LineMeas.x2 = x
         LineMeas.y2 = y
         zarrang = zATan2(y - LineMeas.y1, x - LineMeas.x1)
         LabMeas(0) = Str$(CInt(zarrang / d2r#)) & " " & Chr$(176)
         LabMeas(0).Refresh
         zarrlen = Sqr((y - LineMeas.y1) ^ 2 + (x - LineMeas.x1) ^ 2)
         LabMeas(1) = Str$(CInt(zarrlen))
         LabMeas(1).Refresh
      End If
      
   Case Smooth1, Smooth2, Smooth4
      If AMouseDown Then
         Select Case ToolType
         Case Smooth1:   SmoothArea px1, py1, 1
         Case Smooth2:   SmoothArea px1, py1, 2
         Case Smooth4:   SmoothArea px1, py1, 4
         End Select
         DISPLAY
      End If
   End Select
      
   If Button = 0 Then
      
      Select Case ToolType
      Case Spray
         Select Case SprayType
         Case Dots1, Dots2, Dots3, _
              Plusses1, Plusses2, Plusses3, _
              Crosses1, Crosses2, Crosses3, _
              Diamonds1, Diamonds2, Diamonds3
            MoveSprays PIC, Button, x, y
            If LCNum > 0 Then
               NN = LCNum
            End If
         End Select
         
      Case ALine
         Select Case LineType
         Case SingleLine1, SingleLine2, SingleLine3
            If Button = 0 Then MoveSingleLines PIC, Button, x, y
         Case DottedLine1, DottedLine2, DottedLine3, _
              DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
            If Button = 0 Then MoveDottedLines PIC, Button, x, y
         Case DoubleLine1, DoubleLine2, DoubleLine3, _
              DoubleLineEnd1, DoubleLineEnd2, DoubleLineEnd3, _
              ShadedLine1, ShadedLine2, ShadedLine3
            If Button = 0 Then MoveDoubleAndShadedLines PIC, Button, x, y
         End Select
      
      Case PolyLine
         Select Case PolyLineType
         Case PolySingleLine1, PolySingleLine2, PolySingleLine3
            MovePolySingleLines PIC, Button, x, y
         
         Case PolyDoubleLine1, PolyDoubleLine2, PolyDoubleLine3, _
              PolyDoubleLineEnd1, PolyDoubleLineEnd2, PolyDoubleLineEnd3, _
              PolyShadedLine1, PolyShadedLine2, PolyShadedLine3
            MovePolyDoubleAndShadedLines PIC, Button, x, y
         End Select
      
      Case CurvyLine
         svPolyLineType = PolyLineType
         PolyLineType = CurvyLineType
         Select Case CurvyLineType
         Case CurvySingleLine1, CurvySingleLine2, CurvySingleLine3
            MovePolySingleLines PIC, Button, x, y
            PolyLineType = svPolyLineType
         Case CurvyDoubleLine1, CurvyDoubleLine2, CurvyDoubleLine3, _
              CurvyDoubleLineEnd1, CurvyDoubleLineEnd2, CurvyDoubleLineEnd3, _
              CurvyShadedLine1, CurvyShadedLine2, CurvyShadedLine3
            MovePolyDoubleAndShadedLines PIC, Button, x, y
            PolyLineType = svPolyLineType
         End Select
      
      Case Rectangle
         svRectangleType = 0
         Select Case RectangleType
         Case RectangleSingle1, RectangleSingle2, RectangleSingle3
            MoveRectangleSingle PIC, Button, x, y
         Case RectangleDotted1, RectangleDotted2, RectangleDotted3
            MoveRectangleDotted PIC, Button, x, y
         Case RectangleDouble1, RectangleDouble2, RectangleDouble3
            MoveRectangleDouble PIC, Button, x, y
         Case RectangleShaded1, RectangleShaded2, RectangleShaded3, _
              RectangleFShade, RectangleBShade, RectangleFilled
            svRectangleType = RectangleType
            RectangleType = 0
            MoveRectangleSingle PIC, Button, x, y
            RectangleType = svRectangleType
         End Select
      
      Case Cirllipse
            MoveCirllipseSDD PIC, Button, x, y
      Case Cone
            MoveConeAndTube PIC, Button, x, y
      Case Tube
            MoveConeAndTube PIC, Button, x, y
      Case Bullet
            MoveBullet PIC, Button, x, y
      Case Junction
            MoveJunction PIC, Button, x, y
      Case Arc
         MoveArc PIC, Button, x, y
      Case Shape
         Select Case ShapeType
         Case TShape6   ' Dumbell
            MoveShapeB PIC, Button, x, y
         Case Else
            MoveShapeA PIC, Button, x, y
         End Select
      Case Radial
         MoveRadial PIC, Button, x, y
      Case AFill
      Case Tree
         MoveTree PIC, Button, x, y
      Case Arrow
         MoveArrow PIC, Button, x, y
      Case AText
         If LCNum = 1 Then
         MoveText PIC, Button, x, y
         End If
      '------------------------------------------
      Case SelR
         If LCNum = 1 Then       ' Size Sel rectangle
            With shpRect
               shw = Abs(shx - x)
               shh = Abs(shy - y)
               If x <= shx Then .Left = x Else .Left = x - shw
               If y <= shy Then .Top = y Else .Top = y - shh
               .Width = shw
               .Height = shh
            End With
         ElseIf LCNum = 2 Then  ' Move Sel rectangle
            With shpRect
               .Left = .Left + (x - shx)
               If .Left < 0 Then .Left = 0
               If .Left + .Width > canvasW - 1 Then .Left = canvasW - 1 - .Width
               .Top = .Top + (y - shy)
               If .Top < 0 Then .Top = 0
               If .Top + .Height > canvasH - 1 Then .Top = canvasH - 1 - .Height
               shx = x
               shy = y
            End With
         End If
      Case SelC
         If LCNum = 1 Then       ' Size Sel circle
            With shpCirc
               shw = Abs(shx - x)
               shh = Abs(shy - y)
               If shh > shw Then shw = shh
               If shw > shh Then shh = shw
               If x <= shx Then .Left = x Else .Left = x - shw
               If y <= shy Then .Top = y Else .Top = y - shh
               .Width = shw
               .Height = shh
               shpRect.Left = .Left
               shpRect.Top = .Top
               shpRect.Width = .Width
               shpRect.Height = .Height
            End With
         ElseIf LCNum = 2 Then   ' Move Sel circle
            With shpCirc
               .Left = .Left + (x - shx)
               If .Left < 0 Then .Left = 0 Else shx = x
               If .Left + .Width > canvasW - 1 Then
                  .Left = canvasW - 1 - .Width
               Else
                  shx = x
               End If
               .Top = .Top + (y - shy)
               If .Top < 0 Then .Top = 0 Else shy = y
               If .Top + .Height > canvasH - 1 Then
                  .Top = canvasH - 1 - .Height
               Else
                  shy = y
               End If
               shpRect.Left = .Left
               shpRect.Top = .Top
               shpRect.Width = .Width
               shpRect.Height = .Height
            End With
         End If
      Case SelE
         If LCNum = 1 Then       ' Size Sel ellipse
            With shpEllip
               shw = Abs(shx - x)
               shh = Abs(shy - y)
               If x <= shx Then .Left = x Else .Left = x - shw
               If y <= shy Then .Top = y Else .Top = y - shh
               .Width = shw
               .Height = shh
               shpRect.Left = .Left
               shpRect.Top = .Top
               shpRect.Width = .Width
               shpRect.Height = .Height
            End With
         ElseIf LCNum = 2 Then   ' Move Sel ellipse
            With shpEllip
               .Left = .Left + (x - shx)
               If .Left < 0 Then .Left = 0
               If .Left + .Width > canvasW - 1 Then .Left = canvasW - 1 - .Width
               .Top = .Top + (y - shy)
               If .Top < 0 Then .Top = 0
               If .Top + .Height > canvasH - 1 Then .Top = canvasH - 1 - .Height
               shx = x
               shy = y
               shpRect.Left = .Left
               shpRect.Top = .Top
               shpRect.Width = .Width
               shpRect.Height = .Height
            End With
         End If
      Case SelL   ' Lasso
         If LCNum = 1 Then
            Draw_Lasso x, y
         ElseIf LCNum = 2 Then ' Move selection
            IncrX = x - shx
            IncrY = y - shy
            For NN = 0 To NumLassoLines - 1
               With SL(NN)
                  If .x1 + IncrX < 0 Then IncrX = 0
                  If .x1 + IncrX > canvasW - 1 Then IncrX = 0
                  If .x2 + IncrX < 0 Then IncrX = 0
                  If .x2 + IncrX > canvasW - 1 Then IncrX = 0

                  If .y1 + IncrY < 0 Then IncrY = 0
                  If .y1 + IncrY > canvasH - 1 Then IncrY = 0
                  If .y2 + IncrY < 0 Then IncrY = 0
                  If .y2 + IncrY > canvasH - 1 Then IncrY = 0
               End With
            Next NN
            For NN = 0 To NumLassoLines - 1
               With SL(NN)
                  .x1 = .x1 + IncrX
                  .x2 = .x2 + IncrX
                  .y1 = .y1 + IncrY
                  .y2 = .y2 + IncrY
               End With
            Next NN
            shx = x
            shy = y
         End If
      
      Case Desel
      
      '------------------------------------------------
      Case SCopyPaste To SPaste
         If LCNum = 1 Then
            If aSelRect Then
               With shpRect
                  If x - .Width >= 0 Then .Left = x - .Width
                  If y - .Height >= 0 Then .Top = y - .Height
                  ' Selected image rect
                  picPic.Left = PICC.Left + PIC.Left + .Left   '//
                  picPic.Top = PICC.Top + PIC.Top + .Top       '//
               End With
            ElseIf aSelCirc Then
               With shpCirc
                  If x - .Width >= 0 Then .Left = x - .Width
                  If y - .Height >= 0 Then .Top = y - .Height
                  ' Selected image rect
                  picPic.Left = PICC.Left + PIC.Left + .Left   '//
                  picPic.Top = PICC.Top + PIC.Top + .Top       '//
               End With
            ElseIf aSelEllip Then
               With shpEllip
                  If x - .Width >= 0 Then .Left = x - .Width
                  If y - .Height >= 0 Then .Top = y - .Height
                  ' Selected image rect
                  picPic.Left = PICC.Left + PIC.Left + .Left   '//
                  picPic.Top = PICC.Top + PIC.Top + .Top       '//
               End With
            ElseIf aSelLasso Then
               IncrX = x - shx
               IncrY = y - shy
               If XSMin + IncrX < 1 Then IncrX = 0
               If YSMin + IncrY < 1 Then IncrY = 0
               XSMin = XSMin + IncrX
               YSMin = YSMin + IncrY
               For NN = 0 To NumLassoLines - 1
                  With SL(NN)
                     .x1 = .x1 + IncrX
                     .x2 = .x2 + IncrX
                     .y1 = .y1 + IncrY
                     .y2 = .y2 + IncrY
                  End With
               Next NN
               SSX = XSMin
               SSY = YSMin
               ' Selected image rect
               picPic.Left = PICC.Left + PIC.Left + SSX
               picPic.Top = PICC.Top + PIC.Top + SSY
               shx = x
               shy = y
            End If
         End If
      
      End Select
   
   End If   ' If Button = 0
   
   If aZoom Then
      frmZoom.LabXY = Str$(px1) & "," & Str$(py1)
      frmZoom.LabXY.Refresh
      ZOOMER
   End If
   
End Sub

Private Sub picPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PIC_MouseUp 0, 0, 0, 0
End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NN As Long
Dim x1 As Single, y1 As Single
Dim cn As Long ' Color num
Dim svFillType As Long

' rest all variables public

   If NewNum > 0 Then NewNum = 2
      
   PIC.SetFocus
   
   xprev = x
   yprev = y
   
   AMouseDown = False
   
   If ToolType <> Pick Then
   If ToolType < Rot90 Then
      ADRAW = True
      DISABLES
   End If
   End If
   
   ' Adjust ccords to bArray()
   Getpx1py1 CSng(x), 0, CSng(y), 0
   
   Select Case ToolType
   Case Brush
      Select Case BrushType
      Case Dot1, Dot2, Dot3
         GetColor Button
         SetDot px1, py1, CulNum, BrushType
         LCNum = -1
      Case FreeDraw1, FreeDraw2, FreeDraw3
         StartEndFreedraws PIC, Button, x, y
      Case BRibbon1, BRibbon2, BRibbon3, FRibbon1, FRibbon2, FRibbon3  ' \ /
         StartEndRibbons PIC, Button, x, y
      End Select
   
   Case Spray
      Select Case SprayType
      Case Dots1, Dots2, Dots3, _
           Plusses1, Plusses2, Plusses3, _
           Crosses1, Crosses2, Crosses3, _
           Diamonds1, Diamonds2, Diamonds3
         StartEndSprays PIC, Button, x, y
      End Select
   
   Case ALine
      Select Case LineType
      Case SingleLine1, SingleLine2, SingleLine3
         StartEndSingleLines PIC, Button, x, y
      Case DottedLine1, DottedLine2, DottedLine3, _
           DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         StartEndDottedLines PIC, Button, x, y
      Case DoubleLine1, DoubleLine2, DoubleLine3, _
           DoubleLineEnd1, DoubleLineEnd2, DoubleLineEnd3, _
           ShadedLine1, ShadedLine2, ShadedLine3
         StartEndDoubleAndShadedLines PIC, Button, x, y
      End Select
   
   Case PolyLine
      Select Case PolyLineType
      Case PolySingleLine1, PolySingleLine2, PolySingleLine3
         StartEndPolySingleLines PIC, Button, x, y
      Case PolyDoubleLine1, PolyDoubleLine2, PolyDoubleLine3, _
           PolyDoubleLineEnd1, PolyDoubleLineEnd2, PolyDoubleLineEnd3, _
           PolyShadedLine1, PolyShadedLine2, PolyShadedLine3
         StartEndPolyDoubleAndShadedLines PIC, Button, x, y
      End Select
   
   Case CurvyLine
      svPolyLineType = PolyLineType
      PolyLineType = CurvyLineType
      Select Case CurvyLineType
      Case CurvySingleLine1, CurvySingleLine2, CurvySingleLine3
         StartEndPolySingleLines PIC, Button, x, y
         PolyLineType = svPolyLineType
      Case CurvyDoubleLine1, CurvyDoubleLine2, CurvyDoubleLine3, _
           CurvyDoubleLineEnd1, CurvyDoubleLineEnd2, CurvyDoubleLineEnd3, _
           CurvyShadedLine1, CurvyShadedLine2, CurvyShadedLine3
         StartEndPolyDoubleAndShadedLines PIC, Button, x, y
         PolyLineType = svPolyLineType
      End Select
   
   Case Rectangle
      svRectangleType = 0
      Select Case RectangleType
      Case RectangleSingle1, RectangleSingle2, RectangleSingle3
         StartEndRectangleSingle PIC, Button, x, y
      Case RectangleDotted1, RectangleDotted2, RectangleDotted3
         StartEndRectangleDotted PIC, Button, x, y
      Case RectangleDouble1, RectangleDouble2, RectangleDouble3
         StartEndRectangleDouble PIC, Button, x, y
      Case RectangleShaded1, RectangleShaded2, RectangleShaded3, _
           RectangleFShade, RectangleBShade, RectangleFilled
         svRectangleType = RectangleType
         RectangleType = 0
         StartEndRectangleSingle PIC, Button, x, y
         RectangleType = svRectangleType
      End Select
   
   Case Cirllipse
         StartEndCirllipseSDD PIC, Button, x, y
   Case Cone
         StartEndConeAndTube PIC, Button, x, y
   Case Tube
         StartEndConeAndTube PIC, Button, x, y
   Case Bullet
         StartEndConeAndTube PIC, Button, x, y
   Case Junction
         StartEndJunction PIC, Button, x, y
   Case Arc
      StartEndArc PIC, Button, x, y
   Case Shape
      Select Case ShapeType
      Case TShape6   ' Dumbell
         StartEndShapeB PIC, Button, x, y
      Case Else
         StartEndShapeA PIC, Button, x, y
      End Select
   Case Radial
      StartEndRadials PIC, Button, x, y
   Case AFill
      GetColor Button
      If FillType = Fill21 Or FillType = Fill22 Then
         svFillType = FillType
         FillType = 0
         ' Get ixpmin/max iypmin/max
         PatternFiller bArray(), px1, py1, CulNum
         FillType = svFillType
         PatternFiller bArray(), px1, py1, CulNum
      Else
         PatternFiller bArray(), px1, py1, CulNum
      End If
      LCNum = -1
   Case Tree
      StartEndTree PIC, Button, x, y
   Case Arrow
      StartEndArrow PIC, Button, x, y
   Case AText
      If LCNum <> -1 Then
         PIC.DrawMode = 7
         LCNum = LCNum + 1
         If LCNum = 1 Then ' Move text
            If NSTOREXY > 0 Then
               For NN = 1 To NSTOREXY
                  STOREX(NN) = STOREX(NN) + x
                  STOREY(NN) = STOREY(NN) + y
                  SetPixelV PIC.hDC, STOREX(NN), STOREY(NN), DCul
               Next NN
               PIC.Refresh
            Else
               LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
            End If
         ElseIf LCNum = 2 Then
            If NSTOREXY > 0 Then
               CompleteText CulNum
            End If
            PIC.DrawWidth = 1
            'NSTOREXY = 0     ' Preserve for copying
            optTools(AText).Value = False
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
         End If
      End If
   '-------------------------------------------
   Case SelR
      LCNum = LCNum + 1
      If LCNum = 1 Then
         shx = x
         shy = y
         With shpRect
            .Left = x
            .Top = y
            .Visible = True
         End With
         aSelRect = True
         For NN = SCopyPaste To SClear
            optTools(NN).Enabled = True
         Next NN
         optTools(SRotate).Enabled = False
      ElseIf LCNum = 2 Then   ' Move
         shx = x
         shy = y
      Else
         With shpRect
            SSW = .Width And &HFFFFFFFC
            SSH = .Height And &HFFFFFFFC
            If SSW < 4 Then .Width = 4
            If SSH < 4 Then .Height = 4
            .Width = SSW
            .Height = SSH
            SSX = .Left
            SSY = .Top
         End With
         'ReDim bMask(SSW, SSH)
         
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
   Case SelC   ' Circular selection
      cmdStrip.Enabled = True
      LCNum = LCNum + 1
      If LCNum = 1 Then
         shpRect.Visible = True
         shx = x
         shy = y
         With shpCirc
            .Left = x
            .Top = y
            .Visible = True
         End With
         aSelCirc = True
         For NN = SCopyPaste To SClear
            optTools(NN).Enabled = True
         Next NN
         cmdStrip.Enabled = True
      ElseIf LCNum = 2 Then   ' Move
         shx = x
         shy = y
      Else
         shpRect.Visible = False
         With shpCirc
            SSW = .Width And &HFFFFFFFC
            SSH = .Height And &HFFFFFFFC
            If SSW < 4 Then .Width = 4
            If SSH < 4 Then .Height = 4
            .Width = SSW
            .Height = SSH
            SSX = .Left
            SSY = .Top
         End With
         'ReDim bMask(SSW, SSH)
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
   Case SelE   ' Elliptical selection
      LCNum = LCNum + 1
      If LCNum = 1 Then
         shpRect.Visible = True
         shx = x
         shy = y
         With shpEllip
            .Left = x
            .Top = y
            .Visible = True
         End With
         aSelEllip = True
         For NN = SCopyPaste To SClear
            optTools(NN).Enabled = True
         Next NN
         optTools(SRotate).Enabled = False
      ElseIf LCNum = 2 Then   ' Move
         shx = x
         shy = y
      Else
         shpRect.Visible = False
         With shpEllip
            SSW = .Width And &HFFFFFFFC
            SSH = .Height And &HFFFFFFFC
            If SSW < 4 Then .Width = 4
            If SSH < 4 Then .Height = 4
            .Width = SSW
            .Height = SSH
            SSX = .Left
            SSY = .Top
         End With
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
   Case SelL   ' Lasso selection
      LCNum = LCNum + 1
      If LCNum = 1 Then
         aSelLasso = True
         For NN = SCopyPaste To SClear
            optTools(NN).Enabled = True
         Next NN
         optTools(SRotate).Enabled = False
         Start_Lasso x, y
      ElseIf LCNum = 2 Then
         If x < 1 Then x = 1
         If y < 1 Then y = 1
         If x > canvasW - 2 Then x = canvasW - 2
         If y > canvasH - 2 Then y = canvasH - 2
         shx = x
         shy = y
         ' Close shape
         NumLassoLines = NumLassoLines + 1
         Load SL(NumLassoLines - 1)
         With SL(NumLassoLines - 1)
            .x1 = x
            .y1 = y
            .x2 = SL(0).x1
            .y2 = SL(0).y1
            .Visible = True
         End With
         FindLassoMaxMin
         shx = SSX + SSW
         shy = SSY + SSH
         ' & move
      Else  ' Find min rectangle surrounding lasso
         FindLassoMaxMin
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
   
   Case Desel
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   '-------------------------------------------
   
   Case SCopy To SRotate
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      picPic.Left = PICC.Left + PICC.Width + 100
      picPic.Top = 176
   Case SCopyPaste, SPaste
      LCNum = LCNum + 1
      LabInfo(3) = LCNum
      If LCNum = 1 Then ' Move
         If aSelLasso Then
            IncrX = x - XSMin
            IncrY = y - YSMin
            If XSMin + IncrX < 1 Then IncrX = 0
            If YSMin + IncrY < 1 Then IncrY = 0
            XSMin = XSMin + IncrX
            YSMin = YSMin + IncrY
            For NN = 0 To NumLassoLines - 1
               With SL(NN)
                  .x1 = .x1 + IncrX
                  .x2 = .x2 + IncrX
                  .y1 = .y1 + IncrY
                  .y2 = .y2 + IncrY
               End With
            Next NN
            SSX = XSMin + SSW
            SSY = YSMin + SSH
            shx = SSX
            shy = SSY
         Else
            shx = x
            shy = y
         End If
      Else  ' LCNum = 2
         If aSelRect Then
            SSX = shpRect.Left
            SSY = shpRect.Top
            InsertbPic
            picPic.Left = PICC.Left + PICC.Width + 100
            picPic.Top = 176
         ElseIf aSelCirc Then
            SSX = shpCirc.Left
            SSY = shpCirc.Top
            InsertbPic
            picPic.Left = PICC.Left + PICC.Width + 100
            picPic.Top = 176
         ElseIf aSelEllip Then
            SSX = shpEllip.Left
            SSY = shpEllip.Top
            InsertbPic
            picPic.Left = PICC.Left + PICC.Width + 100
            picPic.Top = 176
         ElseIf aSelLasso Then
            SSX = XSMin
            SSY = YSMin
            InsertbPic
            picPic.Left = PICC.Left + PICC.Width + 100
            picPic.Top = 176
         End If
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
      
   Case SClear To LRColor
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   Case Measure
      LineMeas.Visible = False
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   Case Pick
      cn = bArray(px1, py1)
      If Button = vbLeftButton Then
         ' New Left color
         SelLeftCulNum = cn
         LabSelCul(0).BackColor = CulRGB(SelLeftCulNum)
         iy = (8 * (SelLeftCulNum \ 8) + 4)
         ix = (8 * (SelLeftCulNum - 8 * (SelLeftCulNum \ 8)) + 4)
         picPAL_MouseUp 1, 0, CSng(ix), CSng(iy)   ' Left button color
      ElseIf Button = vbRightButton Then
         ' New Right color
         SelRightCulNum = cn
         LabSelCul(1).BackColor = CulRGB(SelRightCulNum)
         iy = 8 * (SelRightCulNum \ 8) + 4
         ix = 8 * (SelRightCulNum - 8 * (SelRightCulNum \ 8)) + 4
         picPAL_MouseUp 2, 0, CSng(ix), CSng(iy)   ' Right button color
      End If
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   
      Case Smooth1, Smooth2, Smooth4
         Select Case ToolType
         Case Smooth1:   SmoothArea px1, py1, 1
         Case Smooth2:   SmoothArea px1, py1, 2
         Case Smooth4:   SmoothArea px1, py1, 4
         End Select
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   
   End Select
   '-----------------------------------
   'LabInfo(3) = Str$(LCNum)
   
   If LCNum = -1 Then   ' End & Display from bArray()
      PIC.DrawWidth = 1
      PIC.DrawMode = 13
      
      Select Case ToolType
      Case SelR, SelC, SelE, SelL
         ASELECTION = True
         mnuSaveSelection.Enabled = True
         cmdFile(3).Enabled = True
         mnuPrintSelection.Enabled = True
         cmdPrint(1).Enabled = True
      Case Desel
      Case Measure
      Case Pick
      Case Else
         SAVE_CurrentImage  ' Sets TopUndoNum = UndoNum
         FixUndos           ' unless StopUndos = True
         DISPLAY
      End Select
      
      LCNum = 0
      RCNum = 0
      AMoveAll = False
      CopyStartButton = StartButton
      StartButton = 0
      
      ENABLES
      ADRAW = False
   End If
   
   FillLabInfos

   If aZoom Then
      frmZoom.LabXY = Str$(px1) & "," & Str$(py1)
      frmZoom.LabXY.Refresh
      ZOOMER
   End If
End Sub

Private Sub FindLassoMaxMin()
Dim NN As Long
   XSMin = 10000: YSMin = 10000
   XSMax = -10000: YSMax = -10000
   For NN = 0 To NumLassoLines - 1
      With SL(NN)
         If .x1 < XSMin Then XSMin = .x1
         If .y1 < YSMin Then YSMin = .y1
         If .x1 > XSMax Then XSMax = .x1
         If .y1 > YSMax Then YSMax = .y1
         If .x2 < XSMin Then XSMin = .x2
         If .y2 < YSMin Then YSMin = .y2
         If .x2 > XSMax Then XSMax = .x2
         If .y2 > YSMax Then YSMax = .y2
      End With
   Next NN
   SSX = XSMin
   SSY = YSMin
   SSW = (XSMax - XSMin + 8) And &HFFFFFFFC
   SSH = (YSMax - YSMin + 8) And &HFFFFFFFC
End Sub

Private Sub DISABLES()
   LabLight.BackColor = vbRed
   
   If ToolType <> Brush Or (ToolType = Brush And BrushType > Dot3) Then
      fraTools.Enabled = False
      fraRoller.Enabled = False
      fraPal.Enabled = False
      fraMeas.Enabled = False
   End If
End Sub

Private Sub ENABLES()
   fraTools.Enabled = True
   fraRoller.Enabled = True
   fraPal.Enabled = True
   fraMeas.Enabled = True
   fraRot.Enabled = True
   fraStrip.Enabled = True
   LabLight.BackColor = vbGreen
End Sub

Private Sub ImageStart_Enabler()
Dim k As Long
   mnuSave.Enabled = True
   cmdFile(2).Enabled = True
   mnuPrint.Enabled = True
   cmdPrint(0).Enabled = True
   mnuViews.Enabled = True
   mnuZoom.Enabled = True
   mnuTransforms.Enabled = True
   For k = SelR To Desel
      optTools(k).Enabled = True
   Next k
   optTools(Rot90).Enabled = True
   optTools(Mix).Enabled = True
   optTools(Thicken).Enabled = True
   optTools(LRColor).Enabled = True
   For k = Measure To Smooth4
      optTools(k).Value = False
   Next k
   fraRoller.Enabled = True
   fraSmoothers.Enabled = True
End Sub

Private Sub Start_Lasso(x As Single, y As Single)
Dim k As Long
   If NumLassoLines > 1 Then ' Clear extra lasso lines SL(1)-SL(NumLassoLines-1)
      For k = 1 To NumLassoLines - 1
         Unload SL(k)
      Next k
      NumLassoLines = 1
   End If
   If x < 3 Then x = 3
   If y < 3 Then y = 3
   If x > canvasW - 3 Then x = canvasW - 3
   If y > canvasH - 3 Then y = canvasH - 3
   XS1 = x
   YS1 = y
   XS2 = x
   YS2 = y
   With SL(0)
      .x1 = XS1
      .y1 = YS1
      .x2 = XS2
      .y2 = YS2
   End With
   SL(0).Visible = True
End Sub

Private Sub Draw_Lasso(x As Single, y As Single)
   NumLassoLines = NumLassoLines + 1
   Load SL(NumLassoLines - 1)
   If x < 3 Then x = 3
   If y < 3 Then y = 3
   If x > canvasW - 3 Then x = canvasW - 3
   If y > canvasH - 3 Then y = canvasH - 3
   XS1 = XS2
   YS1 = YS2
   XS2 = x
   YS2 = y
   With SL(NumLassoLines - 1)
      .x1 = XS1
      .y1 = YS1
      .x2 = XS2
      .y2 = YS2
      If ((NumLassoLines - 1) Mod 2) = 0 Then
         .BorderColor = vbWhite
      Else
         .BorderColor = vbBlack
      End If
      .Visible = True
   End With
End Sub

Private Sub ShowInfoAtCursor(px As Long, py As Long)
Dim LCul As Long
Dim cn As Long
   LabXY = Str$(px) & "," & Str$(canvasH - py + 1)
   LabXY.Refresh
   cn = bArray(px, py)
   LabSelCul(2).BackColor = CulRGB(cn)
   LabSelCul(2).Refresh
   LabSelCul(3) = Str$(cn)
   LCul = CulRGB(cn)
   LngToRGB LCul
   LabSelCul(4) = Str$(bred) & "," & Str$(bgreen) & "," & Str$(bblue)
   LabSelCul(4).Refresh
End Sub

'#### CANVAS SCROLL BARS #########################################

Private Sub HS_Change()
   PIC.Left = -HS.Value
End Sub

Private Sub HS_Scroll()
   PIC.Left = -HS.Value
End Sub

Private Sub VS_Change()
   PIC.Top = -VS.Value
End Sub

Private Sub VS_Scroll()
   PIC.Top = -VS.Value
End Sub
'#### END CANVAS SCROLL BARS #########################################


'#### QUITTING ###############################################

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form
Dim res As Long
   If UnloadMode = 0 Then    'Close on Form1 pressed
      res = MsgBox("", vbQuestion + vbYesNo + vbSystemModal, "Quit ?")
      If res = vbNo Then
         Cancel = True
         SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
      Else
         Cancel = False
         Screen.MousePointer = vbDefault
         
         ' Hard name = "PaintInfo.txt"
         ' Save Paths & Info
         PrintPaintInfoTxt
         EraseArrays
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      End If
   End If
End Sub

Private Sub EraseArrays()
   ' Probably not necessary
   Erase bArray()
   Erase bUndo1()
   Erase bUndo2()
   Erase bUndo3()
   Erase bUndo4()
   Erase bUndo5()
   Erase bUndo6()
   Erase bUndo7()
   Erase bUndo8()
   Erase bUndo9()
   Erase bUndo10()
   Erase bUndo11()
   Erase bUndo12()
   Erase bUndo13()
   Erase bUndo14()
   Erase bUndo15()
End Sub

