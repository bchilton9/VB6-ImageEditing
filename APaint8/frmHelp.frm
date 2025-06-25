VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " APaint8  Help"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClosehelp 
      Caption         =   "X"
      Height          =   270
      Index           =   1
      Left            =   975
      TabIndex        =   7
      Top             =   45
      Width           =   4320
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "&Home"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   45
      Width           =   720
   End
   Begin VB.PictureBox picHelpC 
      Height          =   5865
      Left            =   120
      ScaleHeight     =   387
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   334
      TabIndex        =   2
      Top             =   345
      Width           =   5070
      Begin VB.PictureBox picHelp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   52500
         Left            =   0
         ScaleHeight     =   3496
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   326
         TabIndex        =   3
         Top             =   15
         Width           =   4950
         Begin VB.CommandButton Command1 
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
            Height          =   210
            Index           =   11
            Left            =   2805
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1335
            Width           =   510
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
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
            Height          =   255
            Index           =   10
            Left            =   4425
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1170
            Width           =   225
         End
         Begin VB.CommandButton Command1 
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
            Index           =   9
            Left            =   3555
            Picture         =   "frmHelp.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1230
            Width           =   480
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show views"
            Height          =   300
            Index           =   7
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2070
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tools"
            Height          =   285
            Index           =   1
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   330
            Width           =   1110
         End
         Begin VB.CommandButton Command1 
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
            Index           =   8
            Left            =   2580
            Picture         =   "frmHelp.frx":010A
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1125
            Width           =   480
         End
         Begin VB.CommandButton Command1 
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
            Index           =   7
            Left            =   3075
            Picture         =   "frmHelp.frx":0214
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1125
            Width           =   450
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Transformers"
            Height          =   300
            Index           =   11
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3270
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Drawing"
            Height          =   285
            Index           =   2
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   585
            Width           =   1110
         End
         Begin VB.CommandButton Command1 
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
            Index           =   6
            Left            =   4050
            Picture         =   "frmHelp.frx":031E
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1230
            Width           =   315
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00FFFFFF&
            Caption         =   "References"
            Height          =   300
            Index           =   18
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   5385
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PaintInfo.txt"
            Height          =   300
            Index           =   17
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   5085
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Other tools"
            Height          =   300
            Index           =   16
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   4770
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print"
            Height          =   300
            Index           =   15
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4470
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Save"
            Height          =   300
            Index           =   14
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4170
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Zoom"
            Height          =   300
            Index           =   13
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   3870
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Strip"
            Height          =   300
            Index           =   12
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   3570
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Effects"
            Height          =   300
            Index           =   10
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2970
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Copy/Paste"
            Height          =   300
            Index           =   9
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2670
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Selections"
            Height          =   300
            Index           =   8
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2370
            Width           =   1110
         End
         Begin VB.CommandButton Command1 
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
            Height          =   255
            Index           =   5
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   2115
            Picture         =   "frmHelp.frx":03B0
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF80FF&
            Height          =   255
            Index           =   3
            Left            =   1920
            Picture         =   "frmHelp.frx":08F2
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   1725
            Picture         =   "frmHelp.frx":0E34
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   1
            Left            =   1530
            Picture         =   "frmHelp.frx":1376
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H80000007&
            Height          =   255
            Index           =   0
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1170
            Width           =   195
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Edit"
            Height          =   300
            Index           =   6
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1770
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Browser"
            Height          =   300
            Index           =   5
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1470
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Palette"
            Height          =   300
            Index           =   4
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1170
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Resizing"
            Height          =   300
            Index           =   3
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   870
            Width           =   1110
         End
         Begin VB.CommandButton cmdC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "&OVERVIEW"
            Height          =   285
            Index           =   0
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   45
            Width           =   1110
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   34
            Left            =   1290
            Picture         =   "frmHelp.frx":18B8
            Top             =   4800
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   33
            Left            =   1605
            Picture         =   "frmHelp.frx":1A02
            Top             =   4800
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   32
            Left            =   3435
            Picture         =   "frmHelp.frx":1AD4
            Top             =   1785
            Width           =   300
         End
         Begin VB.Label Labswap 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "< >"
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
            Height          =   195
            Left            =   4500
            TabIndex        =   30
            Top             =   405
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   31
            Left            =   2820
            Picture         =   "frmHelp.frx":1C1E
            Top             =   1785
            Width           =   300
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   " + Cross-hairs, Roller, Shifter && Smoother"
            Height          =   255
            Left            =   1950
            TabIndex        =   27
            Top             =   4830
            Width           =   2895
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   30
            Left            =   1605
            Picture         =   "frmHelp.frx":1D68
            Top             =   4485
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   29
            Left            =   1290
            Picture         =   "frmHelp.frx":1EB2
            Top             =   4485
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   28
            Left            =   1590
            Picture         =   "frmHelp.frx":1FFC
            Top             =   4170
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   27
            Left            =   1290
            Picture         =   "frmHelp.frx":2146
            Top             =   4170
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   26
            Left            =   1260
            Picture         =   "frmHelp.frx":2290
            Top             =   3570
            Width           =   540
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   25
            Left            =   2865
            Picture         =   "frmHelp.frx":2362
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   24
            Left            =   2535
            Picture         =   "frmHelp.frx":24AC
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   23
            Left            =   2235
            Picture         =   "frmHelp.frx":257E
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   22
            Left            =   1920
            Picture         =   "frmHelp.frx":26C8
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   21
            Left            =   1605
            Picture         =   "frmHelp.frx":279A
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   20
            Left            =   1290
            Picture         =   "frmHelp.frx":286C
            Top             =   2985
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   19
            Left            =   3165
            Picture         =   "frmHelp.frx":29B6
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   18
            Left            =   2850
            Picture         =   "frmHelp.frx":2A88
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   17
            Left            =   2535
            Picture         =   "frmHelp.frx":2B5A
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   16
            Left            =   2235
            Picture         =   "frmHelp.frx":2C2C
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   15
            Left            =   1920
            Picture         =   "frmHelp.frx":2CFE
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   14
            Left            =   1605
            Picture         =   "frmHelp.frx":2E48
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   13
            Left            =   1290
            Picture         =   "frmHelp.frx":2F1A
            Top             =   2670
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   12
            Left            =   2550
            Picture         =   "frmHelp.frx":2FEC
            Top             =   2370
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   11
            Left            =   2235
            Picture         =   "frmHelp.frx":3136
            Top             =   2370
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   10
            Left            =   1920
            Picture         =   "frmHelp.frx":3208
            Top             =   2370
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   9
            Left            =   1605
            Picture         =   "frmHelp.frx":32DA
            Top             =   2370
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   8
            Left            =   1290
            Picture         =   "frmHelp.frx":33AC
            Top             =   2370
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   7
            Left            =   4095
            Picture         =   "frmHelp.frx":347E
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   6
            Left            =   3780
            Picture         =   "frmHelp.frx":3A08
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   5
            Left            =   3120
            Picture         =   "frmHelp.frx":3F92
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   4
            Left            =   2520
            Picture         =   "frmHelp.frx":451C
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   3
            Left            =   2205
            Picture         =   "frmHelp.frx":4666
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   2
            Left            =   1890
            Picture         =   "frmHelp.frx":47B0
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   1
            Left            =   1590
            Picture         =   "frmHelp.frx":48FA
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   0
            Left            =   1290
            Picture         =   "frmHelp.frx":4A44
            Top             =   1785
            Width           =   300
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   1320
            Picture         =   "frmHelp.frx":4B8E
            Top             =   1515
            Width           =   240
         End
      End
   End
   Begin VB.CommandButton cmdClosehelp 
      Caption         =   "X"
      Height          =   270
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   6270
      Width           =   5205
   End
   Begin VB.VScrollBar VSHelp 
      Height          =   5805
      Left            =   5205
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   225
   End
   Begin VB.Label Label1 
      Height          =   285
      Left            =   180
      TabIndex        =   8
      Top             =   6570
      Width           =   930
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmHelp.frm by Robert Rayment

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H2

'  For TEST ONLY --------
'Dim STX As Long, STY As Long
'Dim frmHelpTop As Long
'Dim frmHelpLeft As Long
'-------------------------

Dim a$

Private Sub cmdC_Click(Index As Integer)
   ' Scroll down to item
   With VSHelp
      Select Case Index
      Case 0:  .Value = 406      ' OVERVIEW
      Case 1:  .Value = 685      ' Tools
      Case 2:  .Value = 809      ' Drawing
      Case 3:  .Value = 1020     ' Resizing
      Case 4:  .Value = 1115     ' Palette
      Case 5:  .Value = 1368     ' Browser
      Case 6:  .Value = 1520     ' Edit
      Case 7:  .Value = 1705     ' Show views
      Case 8:  .Value = 1803     ' Selections
      Case 9:  .Value = 1915     ' Copy/Paste
      Case 10: .Value = 2040     ' Effects
      Case 11: .Value = 2194     ' Transformers
      Case 12: .Value = 2348     ' Strip
      Case 13: .Value = 2486     ' Zoom
      Case 14: .Value = 2598     ' Save
      Case 15: .Value = 2710     ' Print
      Case 16: .Value = 2821     ' Other
      Case 17: .Value = 2961     ' PaintIno.txt
      Case 18: .Value = 3060     ' References
      End Select
   End With
   picHelp.SetFocus
End Sub

Private Sub Command1_Click(Index As Integer)
   ' Palette cmd buttons
   VSHelp.Value = 1115
End Sub

Private Sub Form_Load()
'aHelp = True
'  For TEST ONLY --------
' Public STX, STY
' Public frmHelpTop,frmHelpLeft
'STX = 15: STY = 15
'frmHelpTop = frmHelp.Top
'frmHelpLeft = frmHelp.Left
'-------------------------

   ' Position form at last position
   frmHelp.Top = frmHelpTop
   frmHelp.Left = frmHelpLeft
   
   picHelp.Width = 344
   picHelpC.Width = picHelp.Width + 4
   frmHelp.Width = (12 + picHelpC.Width + VSHelp.Width + 12) * STX
      
      ' Size & Make frmZoom stay on top
   SetWindowPos frmHelp.hWnd, hWndInsertAfter, frmHelpLeft * STX, frmHelpTop * STY, _
   frmHelp.Width / STX, frmHelp.Height / STY, wFlags

'<> label in text
Labswap.Left = 26
Labswap.Top = 1301


a$ = ""
a$ = a$ + ""
a$ = a$ & Space$(33) & "APaint8   HELP  by  Robert Rayment 2004" & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr & vbCr
a$ = a$ & vbCr


a$ = a$ & " OVERVIEW" & vbCr & vbCr
a$ = a$ & " This is a 256 color palette paint program and can work with" & vbCr
a$ = a$ & " single or multiple palettes.  Using palettes allows strict" & vbCr
a$ = a$ & " control over the colors used.  Scenes and sprites produced" & vbCr
a$ = a$ & " with this are ideal for simple games where single byte arrays rather" & vbCr
a$ = a$ & " the 3 or 4 byte arrays are wanted.  Such games are easier to code," & vbCr
a$ = a$ & " particularly where color is significant and generally run faster." & vbCr
a$ = a$ & " It wouldn't be difficult to convert to True Color just laborious!" & vbCr
a$ = a$ & " Shading requires graded colors in the palette and some default" & vbCr
a$ = a$ & " palettes are included as well as a palette maker.  8 bpp BMP" & vbCr
a$ = a$ & " images where the width & height are a multiple of 4 and with a" & vbCr
a$ = a$ & " suitable palette can be used directly.  Others can be converted in" & vbCr
a$ = a$ & " the Browser which also contains an approximate color sorter and" & vbCr
a$ = a$ & " re-mapper.  Color sorting requires a Get Nearest Palette Index" & vbCr
a$ = a$ & " function.  An ASM MMX routine has been written for this which is" & vbCr
a$ = a$ & " faster than a compiled VB version which in turn is faster than the" & vbCr
a$ = a$ & " API for this - at least on my PC.  Some Filters also need this." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Tools" & vbCr & vbCr
a$ = a$ & " The right-hand strip in the Tools frame shows all the drawing tools" & vbCr
a$ = a$ & " and variations can be set by bringing up the Tools options form by " & vbCr
a$ = a$ & " right-clicking on any of the tools apart from the Text tool." & vbCr
a$ = a$ & " The left-hand strip in the Tools frame contains a variety of selection" & vbCr
a$ = a$ & " options and effects.  See also Other tools." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Drawing" & vbCr & vbCr
a$ = a$ & " Most drawing is started with a left (LC) or right click (RC) which" & vbCr
a$ = a$ & " selects the Left or Right color and drawing is done without pressing" & vbCr
a$ = a$ & " the mouse buttons.  A shape is finished again with a LC or RC" & vbCr
a$ = a$ & " though there may be intermediate LCs, mouse Moves or RCs for" & vbCr
a$ = a$ & " sizing, orientation and location.  A green strip show red while" & vbCr
a$ = a$ & " a drawing is in progress.  Brief instructions are shown above the" & vbCr
a$ = a$ & " picture.  The Arrow keys can be used for drawing, {Enter} for LC" & vbCr
a$ = a$ & " and {BackSpace} for RC.  For diagonal drawing use the 7,9,1 & 3" & vbCr
a$ = a$ & " keys on the keypad.  Pressing the spacebar immediately after" & vbCr
a$ = a$ & " drawing a shape and moving the mouse makes a copy.  This" & vbCr
a$ = a$ & " applies to all the tools, Brush to Text tools except the Brush" & vbCr
a$ = a$ & " Dots & Fill."
a$ = a$ & vbCr & vbCr

a$ = a$ & " Resizing" & vbCr & vbCr
a$ = a$ & " This sets the width & height of the canvas or canvas & image. Both" & vbCr
a$ = a$ & " must be a multiple of 4 in the range 4 to 2048. Alternatively," & vbCr
a$ = a$ & " if there is a rectangular selection it can be extracted." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Palette." & vbCr & vbCr
a$ = a$ & " There are 4 default shaded palettes, set by the buttons in the" & vbCr
a$ = a$ & " Palette frame:-  Grey, thick(32 banded), thin(16 banded) & centered." & vbCr
a$ = a$ & " The I button inverts the current palette & the D button brings in the" & vbCr
a$ = a$ & " Default palette.  Instead a palette can be made or loaded with the" & vbCr
a$ = a$ & " Palette Maker, brought up by the black button.  In this the current" & vbCr
a$ = a$ & " palette can be brightened, darkened, rotated, reversed or smoothed." & vbCr
a$ = a$ & " If Used this will become the backup palette.  To recover the original" & vbCr
a$ = a$ & " palette make sure to Store it before testing out changes." & vbCr
a$ = a$ & " Also the displayed palette can be made the Default palette on start-up." & vbCr
a$ = a$ & " The Browser also sets a palette when a picture is loaded and used." & vbCr
a$ = a$ & " Shaded shapes use the Left & Right colors." & vbCr
a$ = a$ & " The  < >   button swaps the background color to be black or white" & vbCr
a$ = a$ & " ie color number 0.  This means that when a selection is cleared" & vbCr
a$ = a$ & " or cut the background shows through." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Browser" & vbCr & vbCr
a$ = a$ & " BMPs, GIFs & JPEGs can be loaded.  However before they can be" & vbCr
a$ = a$ & " used they must either already be 256 colors or be converted to 256" & vbCr
a$ = a$ & " colors.  Also they must have a width & height a multiple of 4." & vbCr
a$ = a$ & " The Browser can do these conversions as well as sorting the" & vbCr
a$ = a$ & " palette into approximate color bands, inserting black & white into" & vbCr
a$ = a$ & " color numbers 0 & 1 and remapping the image.  Sorting of large" & vbCr
a$ = a$ & " images is a bit slow but correct BMPs can be used immediately." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Edit" & vbCr & vbCr
a$ = a$ & " 15 levels of Undo/Redo are catered for.  After this earlier views are" & vbCr
a$ = a$ & " lost from the start.  The view stack can be manipulated :--" & vbCr
a$ = a$ & " cleared,  deleted,  all views above the current deleted,  the stack" & vbCr
a$ = a$ & " reduced to just the first & last views,  a view above added to" & vbCr
a$ = a$ & " the current view,  current view swapped with that below and lastly" & vbCr
a$ = a$ & " the undo action can be switched on or off.  When adding views" & vbCr
a$ = a$ & " Overwrite adds in all colors not equal to the background else if" & vbCr
a$ = a$ & " Not Overwrite only adds colors where the destination has the" & vbCr
a$ = a$ & " background color." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Show views" & vbCr & vbCr
a$ = a$ & " Keeps track of the view stack.  Thumbnail images, so only a few" & vbCr
a$ = a$ & " dots will show for thin lines.  From the stack, images can be" & vbCr
a$ = a$ & " shown, deleted, swapped or added." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Selections" & vbCr & vbCr
a$ = a$ & " Rectangular, circular, elliptical and lasso selections can be made." & vbCr
a$ = a$ & " Note that the lasso selection can get messed up, occasionally, if" & vbCr
a$ = a$ & " too complicated.  Just Undo and re-select." & vbCr
a$ = a$ & " Lastly there is Deselect button." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Copy/Paste" & vbCr & vbCr
a$ = a$ & " Selections can be Copy & Pasted, Copied or Cut." & vbCr
a$ = a$ & " Copied and Cut can be pasted with the Paste [PA} button," & vbCr
a$ = a$ & " into the current or other views.  Similarly reflected selections" & vbCr
a$ = a$ & " are just copied.  Only a circular selection can be " & vbCr
a$ = a$ & " rotated at any angle.  This angle is set at the Angle frame." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Effects" & vbCr & vbCr
a$ = a$ & " The first effects tool clears a selection to the backcolor." & vbCr
a$ = a$ & " All the rest work on a selection, if there is one, else on the" & vbCr
a$ = a$ & " whole picture.  The effects are :--  Rotate by 90 degrees," & vbCr
a$ = a$ & " Mix colors (unpredictable but interesting), Thicken, Random dots" & vbCr
a$ = a$ & " of the left or right color and Replace left color by the right" & vbCr
a$ = a$ & " color.  Note that thickening is into the backcolor and will stop" & vbCr
a$ = a$ & " when other colors start overwriting each other." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Transformers" & vbCr & vbCr
a$ = a$ & " Pressing the Transformers menu button brings up a preview" & vbCr
a$ = a$ & " window showing all the Transformers - Filters, Deformers" & vbCr
a$ = a$ & " and Adders which can operate over the whole picture or a" & vbCr
a$ = a$ & " rectangular selection.  Some of the Transformers make use of" & vbCr
a$ = a$ & " a selected color and these are indicated against the item." & vbCr
a$ = a$ & " Most of the effects can be varied by sliding mouse-down over" & vbCr
a$ = a$ & " the gridded box, below the picture." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Strip" & vbCr & vbCr
a$ = a$ & " A circular selection can be turned into a strip of images where" & vbCr
a$ = a$ & " each can be reduced, rotated and peppered with the backcolor." & vbCr
a$ = a$ & " When done the Strip frame will show the number of elements in the" & vbCr
a$ = a$ & " strip and the size of an element ie image width divided by the" & vbCr
a$ = a$ & " number of elements.  To have the elements increasing in size," & vbCr
a$ = a$ & " rotate the whole picture by 90 degrees twice." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Zoom" & vbCr & vbCr
a$ = a$ & " As soon as any image is present the Zoom menu button will be" & vbCr
a$ = a$ & " enabled.  Zoom sizes are 2,4,6,8 & 12.  When the Zoom is on the" & vbCr
a$ = a$ & " zoom factor can also be changed using the F2, F4, F6, F8 & F12" & vbCr
a$ = a$ & " keys.  Drawing is still done on the canvas, zoom shows the result." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Save" & vbCr & vbCr
a$ = a$ & " Saving can be for the whole image or a selection if present." & vbCr
a$ = a$ & " Non-rectangular selections use color number 0 outside the" & vbCr
a$ = a$ & " selection boundary.  The saved image can be checked in the" & vbCr
a$ = a$ & " Browser.  Images can only be saved as 8bpp BMPs or GIFs." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Print" & vbCr & vbCr
a$ = a$ & " As with saving, the whole image or a selection can be printed." & vbCr
a$ = a$ & " Similarly color number 0 is used outside a selection boundary. " & vbCr
a$ = a$ & " Printing is at actual size, so images widths geater than about" & vbCr
a$ = a$ & " 800 pixels might go off the page." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " Other tools" & vbCr & vbCr
a$ = a$ & " Six other tools are Measure, Color picker, Cross-Hairs, Roller, Shifter" & vbCr
a$ = a$ & " & Smoother.  Measure shows lengths & angles, of any object, on a" & vbCr
a$ = a$ & " frame.  Roller rolls the image left, right, up or down 1 or 8 pixels" & vbCr
a$ = a$ & " and Shifter the same except that vacated strips have the backcolor." & vbCr
a$ = a$ & " Smoothing is over a small, medium or large area, but effectiveness" & vbCr
a$ = a$ & " depends on the palette." & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " PaintInfo.txt" & vbCr & vbCr
a$ = a$ & " On exitting the program a PaintInfo.txt file is made consisting " & vbCr
a$ = a$ & " of :-- Open path,  Save path,  17 tool options, location of " & vbCr
a$ = a$ & " the sub-forms and the Default JASC palette" & vbCr
a$ = a$ & vbCr & vbCr

a$ = a$ & " References" & vbCr & vbCr
a$ = a$ & " VBaccelerator (Steve McHahon),-  Dialogs & bpp conversion," & vbCr
a$ = a$ & " Carles P V,-  Clarifying conversion code," & vbCr
a$ = a$ & " Ron V Tilburg,-  GIF save," & vbCr
a$ = a$ & " allapi.net (Donald Grover),-  ShowPrinter," & vbCr
a$ = a$ & " Stefan Casier,-  Center shaded palette," & vbCr
a$ = a$ & " vbapi.com/ref/s/shfileoperation.html,-  SH File Ops," & vbCr
a$ = a$ & " Manuel Santos, Malcolm Ferris & Johannes B,-  Some filters." & vbCr
a$ = a$ & vbCr & vbCr

picHelp.Cls
picHelp.Print a$;
picHelp.Refresh
a$ = " "

FixVBar picHelpC, picHelp, VSHelp

End Sub

Private Sub cmdHome_Click()
   VSHelp.Value = 0
   picHelp.SetFocus
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' For locating scroll Y positions
   Label1 = Str$(x) & Str$(y)
End Sub

Private Sub VSHelp_Change()
   picHelp.Top = -VSHelp.Value
End Sub

Private Sub VSHelp_Scroll()
   picHelp.Top = -VSHelp.Value
End Sub

Private Sub FixVBar(picCr As PictureBox, picP As PictureBox, VS As VScrollBar)
   ' picCr = Picture Container
   ' picP  = Picture
   VS.Max = picP.Height - picCr.Height + 12 ' +4 to allow for border
   VS.LargeChange = picCr.Height \ 10
   VS.SmallChange = 1
   VS.Top = picCr.Top
   VS.Left = picCr.Left + picCr.Width + 1
   VS.Height = picCr.Height
   If picP.Height < picCr.Height Then
      VS.Visible = False
      VS.Enabled = False
   Else
      VS.Visible = True
      VS.Enabled = True
   End If
End Sub

Private Sub cmdClosehelp_Click(Index As Integer)
picHelp.Cls
   
frmHelpTop = frmHelp.Top
frmHelpLeft = frmHelp.Left
SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
Unload Me
End Sub


