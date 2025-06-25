VERSION 5.00
Begin VB.Form frmPalette 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Make palette"
   ClientHeight    =   3375
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   4320
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "frmPalette.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeDefault 
      Caption         =   "Make this palette the default on start-up"
      Height          =   270
      Left            =   1125
      TabIndex        =   33
      Top             =   2880
      Width           =   3105
   End
   Begin VB.CommandButton cmdDBRS 
      Caption         =   "Smo"
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
      Left            =   2745
      TabIndex        =   32
      ToolTipText     =   " Smooth "
      Top             =   345
      Width           =   420
   End
   Begin VB.CommandButton cmdDBRS 
      Caption         =   "Rev"
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
      Index           =   3
      Left            =   2325
      TabIndex        =   31
      ToolTipText     =   " Reverse "
      Top             =   345
      Width           =   420
   End
   Begin VB.PictureBox picSee 
      Height          =   465
      Left            =   225
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   27
      Top             =   2580
      Width           =   450
   End
   Begin VB.CommandButton cmdDBRS 
      Caption         =   "Rot"
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
      Index           =   2
      Left            =   1905
      TabIndex        =   25
      ToolTipText     =   " Rotate "
      Top             =   345
      Width           =   420
   End
   Begin VB.CommandButton cmdDBRS 
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
      Height          =   240
      Index           =   1
      Left            =   1485
      Picture         =   "frmPalette.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   " Darken "
      Top             =   345
      Width           =   420
   End
   Begin VB.CommandButton cmdDBRS 
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
      Height          =   240
      Index           =   0
      Left            =   1035
      Picture         =   "frmPalette.frx":0116
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   " Brighten "
      Top             =   345
      Width           =   420
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "<                    SWAP                    >"
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
      Left            =   1050
      TabIndex        =   20
      Top             =   120
      Width           =   2115
   End
   Begin VB.HScrollBar HSEnd 
      Height          =   150
      Index           =   2
      LargeChange     =   8
      Left            =   2160
      Max             =   255
      TabIndex        =   13
      Top             =   990
      Width           =   1440
   End
   Begin VB.HScrollBar HSEnd 
      Height          =   150
      Index           =   1
      LargeChange     =   8
      Left            =   2160
      Max             =   255
      TabIndex        =   12
      Top             =   810
      Width           =   1440
   End
   Begin VB.HScrollBar HSStart 
      Height          =   150
      Index           =   2
      LargeChange     =   8
      Left            =   210
      Max             =   255
      TabIndex        =   11
      Top             =   960
      Width           =   1440
   End
   Begin VB.HScrollBar HSStart 
      Height          =   150
      Index           =   1
      LargeChange     =   8
      Left            =   210
      Max             =   255
      TabIndex        =   10
      Top             =   795
      Width           =   1440
   End
   Begin VB.HScrollBar HSEnd 
      Height          =   150
      Index           =   0
      LargeChange     =   8
      Left            =   2160
      Max             =   255
      TabIndex        =   9
      Top             =   645
      Width           =   1440
   End
   Begin VB.PictureBox picEnd 
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   705
      TabIndex        =   8
      Top             =   225
      Width           =   765
   End
   Begin VB.HScrollBar HSStart 
      Height          =   150
      Index           =   0
      LargeChange     =   8
      Left            =   210
      Max             =   255
      TabIndex        =   7
      Top             =   630
      Width           =   1440
   End
   Begin VB.PictureBox picStart 
      Height          =   375
      Left            =   195
      ScaleHeight     =   315
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   210
      Width           =   765
   End
   Begin VB.CommandButton cmd256 
      Appearance      =   0  'Flat
      Caption         =   "256"
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
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1185
      Width           =   3840
   End
   Begin VB.CommandButton cmd128 
      Appearance      =   0  'Flat
      Caption         =   "128"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   3840
   End
   Begin VB.CommandButton cmd64 
      Appearance      =   0  'Flat
      Caption         =   "64"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1920
   End
   Begin VB.CommandButton cmd32 
      Appearance      =   0  'Flat
      Caption         =   "32"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1770
      Width           =   960
   End
   Begin VB.CommandButton cmd16 
      Appearance      =   0  'Flat
      Caption         =   "16"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1950
      Width           =   480
   End
   Begin VB.PictureBox picPalette 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   2145
      Width           =   3900
   End
   Begin VB.Label LabSee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   2
      Left            =   735
      TabIndex        =   30
      Top             =   2910
      Width           =   330
   End
   Begin VB.Label LabSee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   1
      Left            =   735
      TabIndex        =   29
      Top             =   2745
      Width           =   330
   End
   Begin VB.Label LabSee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "255"
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
      Height          =   135
      Index           =   0
      Left            =   735
      TabIndex        =   28
      Top             =   2580
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Each range shades from Start to End color."
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
      Left            =   1275
      TabIndex        =   26
      Top             =   2580
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "End"
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
      Index           =   1
      Left            =   3660
      TabIndex        =   22
      Top             =   45
      Width           =   285
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Start"
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
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   45
      Width           =   435
   End
   Begin VB.Label LabEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   2
      Left            =   3645
      TabIndex        =   19
      Top             =   975
      Width           =   330
   End
   Begin VB.Label LabEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   1
      Left            =   3645
      TabIndex        =   18
      Top             =   810
      Width           =   330
   End
   Begin VB.Label LabEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "255"
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
      Height          =   135
      Index           =   0
      Left            =   3645
      TabIndex        =   17
      Top             =   645
      Width           =   330
   End
   Begin VB.Label LabStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   2
      Left            =   1710
      TabIndex        =   16
      Top             =   960
      Width           =   330
   End
   Begin VB.Label LabStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "255"
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
      Height          =   135
      Index           =   1
      Left            =   1710
      TabIndex        =   15
      Top             =   795
      Width           =   330
   End
   Begin VB.Label LabStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "255"
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
      Height          =   165
      Index           =   0
      Left            =   1710
      TabIndex        =   14
      Top             =   645
      Width           =   330
   End
   Begin VB.Menu mnuExit1 
      Caption         =   "&Cancel"
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "&Load"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "S&ave"
   End
   Begin VB.Menu mnuStore 
      Caption         =   "&Store"
   End
   Begin VB.Menu mnuRestore 
      Caption         =   "&Restore"
   End
   Begin VB.Menu mnuUse 
      Caption         =   "&Use"
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPalette.frm

Option Explicit

Dim k As Long
Dim j As Long
Dim p1 As Long
Dim p2 As Long

'Public MPal() As Byte
'Public StorePal() As Byte
Dim RGBStart() As Long
Dim RGBEnd() As Long
Dim zCulStep() As Single
Dim zCulStart() As Single

Dim PalSpec$
Dim CommonDialog1 As OSDialog

Private Sub Form_Load()
   ReDim RGBStart(0 To 2)
   ReDim RGBEnd(0 To 2)
   ReDim zCulStart(0 To 2)
   ReDim zCulStep(0 To 2)
   ReDim MPal(0 To 2, 0 To 255)
   frmPalette.Width = 4410 '4335
   frmPalette.Height = 4400 '4500
   ' Position form at last position
   frmPalette.Top = frmPaletteTop
   frmPalette.Left = frmPaletteLeft
   
   PalSpec$ = AppPathSpec$

   AlignCmds
   
   For k = 0 To 2
      HSStart(k).Value = 255
      HSEnd(k).Value = 255
   Next k
   ' Show Current palette
   For k = 0 To 255
      MPal(0, k) = palRed(k)
      MPal(1, k) = palGreen(k)
      MPal(2, k) = palBlue(k)
   Next k
   ' Show Palette
   For k = 0 To 255
      picPalette.Line (k, 0)-(k, picPalette.Height), _
         RGB(MPal(0, k), MPal(1, k), MPal(2, k))
   Next k
   'StorePalette  ' User choice
End Sub

'###################################################################

Private Sub cmd16_Click(Index As Integer)
' 0   0 to  15
' 1  16 to  31
' 2  32 to  47
' 3  48 to  63
' 4  64 to  79
' 5  80 to  95
' 6  96 to 111
' 7 112 to 127
' 8 128 to 143
' 9 144 to 155
'10 160 to 175
'11 176 to 191
'12 192 to 207
'13 208 to 223
'14 224 to 239
'15 240 to 255
   p1 = 16 * Index
   p2 = p1 + 15
   For k = 0 To 2
      zCulStart(k) = RGBStart(k)
      zCulStep(k) = (RGBEnd(k) - RGBStart(k)) / 16
   Next k
   For k = p1 To p2
      picPalette.Line (k, 0)-(k, picPalette.Height), RGB(zCulStart(0), zCulStart(1), zCulStart(2))
      For j = 0 To 2
         zCulStart(j) = zCulStart(j) + zCulStep(j)
         If zCulStart(j) > 255 Then zCulStart(j) = 0
         If zCulStart(j) < 0 Then zCulStart(j) = 255
      Next j
   Next k
   Extract2MPal
End Sub

Private Sub cmd32_Click(Index As Integer)
' 0   0 to  31
' 1  32 to  63
' 2  64 to  95
' 3  96 to 127
' 4 128 to 159
' 5 160 to 191
' 6 192 to 223
' 7 224 to 255
   p1 = 32 * Index
   p2 = p1 + 32
   For k = 0 To 2
      zCulStart(k) = RGBStart(k)
      zCulStep(k) = (RGBEnd(k) - RGBStart(k)) / 32
   Next k
   For k = p1 To p2
      picPalette.Line (k, 0)-(k, picPalette.Height), RGB(zCulStart(0), zCulStart(1), zCulStart(2))
      For j = 0 To 2
         zCulStart(j) = zCulStart(j) + zCulStep(j)
         If zCulStart(j) > 255 Then zCulStart(j) = 0
         If zCulStart(j) < 0 Then zCulStart(j) = 255
      Next j
   Next k
   Extract2MPal
End Sub

Private Sub cmd64_Click(Index As Integer)
' 0   0 to  63
' 1  64 to 127
' 2 128 to 191
' 3 192 to 255
   p1 = 64 * Index
   p2 = p1 + 64
   For k = 0 To 2
      zCulStart(k) = RGBStart(k)
      zCulStep(k) = (RGBEnd(k) - RGBStart(k)) / 64
   Next k
   For k = p1 To p2
      picPalette.Line (k, 0)-(k, picPalette.Height), RGB(zCulStart(0), zCulStart(1), zCulStart(2))
      For j = 0 To 2
         zCulStart(j) = zCulStart(j) + zCulStep(j)
         If zCulStart(j) > 255 Then zCulStart(j) = 0
         If zCulStart(j) < 0 Then zCulStart(j) = 255
      Next j
   Next k
   Extract2MPal
End Sub

Private Sub cmd128_Click(Index As Integer)
' 0   0 to 127
' 1 128 to 255
   p1 = 128 * Index
   p2 = p1 + 128
   For k = 0 To 2
      zCulStart(k) = RGBStart(k)
      zCulStep(k) = (RGBEnd(k) - RGBStart(k)) / 128
   Next k
   For k = p1 To p2
      picPalette.Line (k, 0)-(k, picPalette.Height), RGB(zCulStart(0), zCulStart(1), zCulStart(2))
      For j = 0 To 2
         zCulStart(j) = zCulStart(j) + zCulStep(j)
         If zCulStart(j) > 255 Then zCulStart(j) = 0
         If zCulStart(j) < 0 Then zCulStart(j) = 255
      Next j
   Next k
   Extract2MPal
End Sub

Private Sub cmd256_Click()
' 0   0 to 255
   p1 = 0
   p2 = p1 + 256
   For k = 0 To 2
      zCulStart(k) = RGBStart(k)
      zCulStep(k) = (RGBEnd(k) - RGBStart(k)) / 256
   Next k
   For k = p1 To p2
      picPalette.Line (k, 0)-(k, picPalette.Height), RGB(zCulStart(0), zCulStart(1), zCulStart(2))
      For j = 0 To 2
         zCulStart(j) = zCulStart(j) + zCulStep(j)
         If zCulStart(j) > 255 Then zCulStart(j) = 0
         If zCulStart(j) < 0 Then zCulStart(j) = 255
      Next j
   Next k
   Extract2MPal
End Sub

'###################################################################

Private Sub cmdSwap_Click()
Dim T As Long
   For k = 0 To 2
      T = HSStart(k).Value
      HSStart(k).Value = HSEnd(k).Value
      HSEnd(k).Value = T
      RGBStart(k) = HSStart(k).Value
      LabStart(k) = Str$(RGBStart(k))
      RGBEnd(k) = HSEnd(k).Value
      LabEnd(k) = Str$(RGBEnd(k))
   Next k
   picStart.BackColor = RGB(RGBStart(0), RGBStart(1), RGBStart(2))
   picStart.Refresh
   picEnd.BackColor = RGB(RGBEnd(0), RGBEnd(1), RGBEnd(2))
   picEnd.Refresh
End Sub

'###################################################################

Private Sub HSStart_Change(Index As Integer)
   RGBStart(Index) = HSStart(Index).Value
   LabStart(Index) = Str$(RGBStart(Index))
   picStart.BackColor = RGB(RGBStart(0), RGBStart(1), RGBStart(2))
End Sub
Private Sub HSEnd_Change(Index As Integer)
   RGBEnd(Index) = HSEnd(Index).Value
   LabEnd(Index) = Str$(RGBEnd(Index))
   picEnd.BackColor = RGB(RGBEnd(0), RGBEnd(1), RGBEnd(2))
End Sub

'###################################################################

Private Sub cmdDBRS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Darken, Brighten, Rotate
Dim bRO As Byte
Dim bGO As Byte
Dim bBO As Byte
   Select Case Index
   Case 0      ' Brighten
      For k = 0 To 255
         If MPal(0, k) < 251 Then MPal(0, k) = MPal(0, k) + 4
         If MPal(1, k) < 251 Then MPal(1, k) = MPal(1, k) + 4
         If MPal(2, k) < 251 Then MPal(2, k) = MPal(2, k) + 4
      Next k
   Case 1      ' Darken
      For k = 0 To 255
         If MPal(0, k) > 4 Then MPal(0, k) = MPal(0, k) - 4
         If MPal(1, k) > 4 Then MPal(1, k) = MPal(1, k) - 4
         If MPal(2, k) > 4 Then MPal(2, k) = MPal(2, k) - 4
         picPalette.Line (k, 0)-(k, picPalette.Height), _
            RGB(MPal(0, k), MPal(1, k), MPal(2, k))
      Next k
   Case 2      ' Rotate
      If Button = vbLeftButton Then
         For j = 0 To 15
            bRO = MPal(0, 2)
            bGO = MPal(1, 2)
            bBO = MPal(2, 2)
            For k = 2 To 255
                MPal(0, k - 1) = MPal(0, k)
                MPal(1, k - 1) = MPal(1, k)
                MPal(2, k - 1) = MPal(2, k)
            Next k
            MPal(0, 255) = bRO
            MPal(1, 255) = bGO
            MPal(2, 255) = bBO
         Next j
      ElseIf Button = vbRightButton Then
         For j = 0 To 15
            bRO = MPal(0, 255)
            bGO = MPal(1, 255)
            bBO = MPal(2, 255)
            For k = 254 To 2 Step -1
                MPal(0, k + 1) = MPal(0, k)
                MPal(1, k + 1) = MPal(1, k)
                MPal(2, k + 1) = MPal(2, k)
            Next k
            MPal(0, 2) = bRO
            MPal(1, 2) = bGO
            MPal(2, 2) = bBO
         Next j
      End If
   Case 3      ' Reverse
      For k = 2 To 128
         bRO = MPal(0, k)
         bGO = MPal(1, k)
         bBO = MPal(2, k)
         MPal(0, k) = MPal(0, 255 - k)
         MPal(1, k) = MPal(1, 255 - k)
         MPal(2, k) = MPal(2, 255 - k)
         MPal(0, 255 - k) = bRO
         MPal(1, 255 - k) = bGO
         MPal(2, 255 - k) = bBO
      Next k
   Case 4   ' Smooth
      For k = 1 To 254
         MPal(0, k) = (1& * MPal(0, k - 1) + MPal(0, k) + MPal(0, k + 1)) / 3
         MPal(1, k) = (1& * MPal(1, k - 1) + MPal(1, k) + MPal(1, k + 1)) / 3
         MPal(2, k) = (1& * MPal(2, k) + MPal(2, k) + MPal(2, k + 1)) / 3
      Next k
   
      MPal(0, 1) = (1& * MPal(0, 1) + MPal(0, 2)) / 2
      MPal(1, 1) = (1& * MPal(1, 1) + MPal(1, 2)) / 2
      MPal(2, 1) = (1& * MPal(2, 1) + MPal(2, 2)) / 2
      
      MPal(0, 255) = (1& * MPal(0, 255) + MPal(0, 254)) / 2
      MPal(1, 255) = (1& * MPal(1, 255) + MPal(1, 254)) / 2
      MPal(2, 255) = (1& * MPal(2, 255) + MPal(2, 254)) / 2
      ConvLongPalDataTo16Bit MPal()
   End Select
   ' Show Palette
   For k = 0 To 255
      picPalette.Line (k, 0)-(k, picPalette.Height), _
         RGB(MPal(0, k), MPal(1, k), MPal(2, k))
   Next k
End Sub


'###################################################################

Private Sub Extract2MPal()
Dim LongCul As Long
Dim H As Long
   H = picPalette.Height \ 2
   For k = 0 To 255
      LongCul = picPalette.Point(k, H)
      MPal(0, k) = LongCul And &HFF&
      MPal(1, k) = (LongCul And &HFF00&) / &H100&
      MPal(2, k) = (LongCul And &HFF0000) / &H10000
   Next k
   ConvLongPalDataTo16Bit MPal()
End Sub

Private Sub StorePalette()
Dim LongCul As Long
Dim H As Long
   H = picPalette.Height \ 2
   For k = 0 To 255
      LongCul = picPalette.Point(k, H)
      StorePal(0, k) = LongCul And &HFF&
      StorePal(1, k) = (LongCul And &HFF00&) / &H100&
      StorePal(2, k) = (LongCul And &HFF0000) / &H10000
   Next k
End Sub

Private Sub RestorePalette()
   For k = 0 To 255
      MPal(0, k) = StorePal(0, k)
      MPal(1, k) = StorePal(1, k)
      MPal(2, k) = StorePal(2, k)
   Next k
   ConvLongPalDataTo16Bit MPal()
End Sub

Private Sub mnuRestore_Click()
   RestorePalette
End Sub

Private Sub mnuStore_Click()
   StorePalette
End Sub

Private Sub cmdMakeDefault_Click()
Dim k As Long
   For k = 0 To 255
      DefaultRGB(k) = RGB(MPal(0, k), MPal(1, k), MPal(2, k))
   Next k
   DefaultRGB(0) = 0
   DefaultRGB(1) = RGB(255, 255, 255)
End Sub

'###################################################################

Private Sub mnuLoad_Click()
Dim Title$, Filt$, InDir$
Dim PSpec$
Set CommonDialog1 = New OSDialog
   Title$ = "Load JASC PAL File"
   Filt$ = "Load pal (*.pal)|*.pal"
   InDir$ = PalSpec$
   CommonDialog1.ShowOpen PSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   If Len(PSpec$) <> 0 Then
      PalSpec$ = PSpec$
      READ_PAL_FILE PSpec$
      ConvLongPalDataTo16Bit MPal()
   End If
Set CommonDialog1 = Nothing
End Sub

Private Sub READ_PAL_FILE(PSpec$)
Dim fnum As Long
Dim a$
   fnum = FreeFile
   On Error GoTo PalFileError
   Open PSpec$ For Input As #fnum
   Line Input #fnum, a$
   Close #fnum
   Open PSpec$ For Input As #fnum
   If InStr(1, a$, "JASC") <> 0 Then  'JASC-PAL MAP file
                   'JASC-PAL
                   '0100
                   'Skip 3 lines  '256
      Line Input #fnum, a$
      Line Input #fnum, a$
      Line Input #fnum, a$
      For k = 0 To 255
         If EOF(1) Then Exit For
         Input #fnum, MPal(0, k), MPal(1, k), MPal(2, k)
      Next k
      Close #fnum
   Else
      MsgBox "Not a JASC PAL file", vbInformation, "Reading PAL file"
   End If
   Exit Sub
'===========
PalFileError:
Close
MsgBox "Pal file error", vbCritical, "Reading pal file"
End Sub

Private Sub mnuSave_Click()
Dim Title$, Filt$, InDir$
Dim PSpec$
Set CommonDialog1 = New OSDialog
   Title$ = "Save JASC PAL File"
   Filt$ = "Save pal (*.pal)|*.pal"
   InDir$ = PalSpec$
   CommonDialog1.ShowSave PSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   If Len(PSpec$) <> 0 Then
      PalSpec$ = PSpec$
      SAVE_PAL_FILE PSpec$
   End If
Set CommonDialog1 = Nothing
End Sub

Private Sub SAVE_PAL_FILE(PSpec$)
Dim fnum As Long
   ConvLongPalDataTo16Bit MPal()
   On Error GoTo PalSaveError
   fnum = FreeFile
   Open PSpec$ For Output As #fnum
   Print #fnum, "JASC-PAL"
   Print #fnum, "0100"
   Print #fnum, "256"
   For k = 0 To 255
      Print #fnum, Trim$(Str$(MPal(0, k))) & " " _
               ; Trim$(Str$(MPal(1, k))) & " " _
               ; Trim$(Str$(MPal(2, k)))
   Next k
   Close #fnum
   Exit Sub
'=========
PalSaveError:
Close
MsgBox "Pal save error", vbCritical, "Saving PAL file"
End Sub

Private Sub mnuUse_Click()
Dim c0 As Long
Dim c1 As Long
' For linking with a prog that wants the PAL immediately
'Public CulRGB() As Long
'Public CulBGR() As Long
'Public palRed() As Byte, palGreen() As Byte, palBlue() As Byte
   For k = 0 To 255
      CulRGB(k) = RGB(MPal(0, k), MPal(1, k), MPal(2, k))
      CulBGR(k) = RGB(MPal(2, k), MPal(1, k), MPal(0, k))
      palRed(k) = MPal(0, k)
      palGreen(k) = MPal(1, k)
      palBlue(k) = MPal(2, k)
   Next k
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
      CulRGB(0) = 0
      CulRGB(1) = RGB(255, 255, 255)
      CulBGR(0) = 0
      CulBGR(1) = RGB(255, 255, 255)
   End If
   ' New backup palette
   BackUpRGB() = CulRGB()
   
   frmPaletteTop = frmPalette.Top
   frmPaletteLeft = frmPalette.Left
   Unload Me
End Sub

Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim LongCul As Long
ReDim RGBT(0 To 2) As Byte
   LongCul = picPalette.Point(x, y)
   RGBT(0) = LongCul And &HFF&
   RGBT(1) = (LongCul And &HFF00&) / &H100&
   RGBT(2) = (LongCul And &HFF0000) / &H10000
   For k = 0 To 2
      LabSee(k) = Str$(RGBT(k))
   Next k
   picSee.BackColor = RGB(RGBT(0), RGBT(1), RGBT(2))
End Sub

Private Sub picPalette_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim LongCul As Long
ReDim RGBT(0 To 2) As Byte
   LongCul = picPalette.Point(x, y)
   RGBT(0) = LongCul And &HFF&
   RGBT(1) = (LongCul And &HFF00&) / &H100&
   RGBT(2) = (LongCul And &HFF0000) / &H10000
   If Button = vbLeftButton Then
      For k = 0 To 2
         HSStart(k).Value = RGBT(k)
         RGBStart(k) = RGBT(k)
         LabStart(k) = RGBT(k)
      Next k
      picStart.BackColor = RGB(RGBStart(0), RGBStart(1), RGBStart(2))
   ElseIf Button = vbRightButton Then
      For k = 0 To 2
         HSEnd(k).Value = RGBT(k)
         RGBEnd(k) = RGBT(k)
         LabEnd(k) = RGBT(k)
      Next k
      picEnd.BackColor = RGB(RGBEnd(0), RGBEnd(1), RGBEnd(2))
   End If
End Sub

Private Sub ConvLongPalDataTo16Bit(TPal() As Byte)
Dim remainder As Long
Dim CRed As Byte, CGreen As Byte, CBlue As Byte
   ' RED   'Valid 16-bit values 0,16,24,32,,,248,255
   For k = 0 To 255
      CRed = TPal(0, k)
      remainder = CRed Mod 8
      If remainder <> 0 And CRed <> 255 Then
         CRed = CRed - remainder
      End If
      If CRed = 8 Then CRed = 0
      TPal(0, k) = CRed
   Next k
   ' GREEN  'Valid 16-bit values 0,8,12,16,20,,,,252,255
   For k = 0 To 255
      CGreen = TPal(1, k)
      remainder = CGreen Mod 4
      If remainder <> 0 And CGreen <> 255 Then
         CGreen = CGreen - remainder
      End If
      If CGreen = 4 Then CGreen = 0
      TPal(1, k) = CGreen
   Next k
   ' BLUE   'Valid 16-bit values 0,16,24,32,,,248,255
   For k = 0 To 255
      CBlue = TPal(2, k)
      remainder = CBlue Mod 8
      If remainder <> 0 And CBlue <> 255 Then
         CBlue = CBlue - remainder
      End If
      If CBlue = 8 Then CBlue = 0
      TPal(2, k) = CBlue
   Next k
   ' Show Palette
   For k = 0 To 255
      picPalette.Line (k, 0)-(k, picPalette.Height), _
         RGB(MPal(0, k), MPal(1, k), MPal(2, k))
   Next k
End Sub

Private Sub AlignCmds()
cmd16(0).Width = 16
For k = 1 To 15
   Load cmd16(k)
   cmd16(k).Visible = True
Next k
For k = 1 To 15
   cmd16(k).Top = cmd16(0).Top
   cmd16(k).Left = cmd16(k - 1).Left + cmd16(0).Width
Next k

cmd32(0).Width = 32
For k = 1 To 7
   Load cmd32(k)
   cmd32(k).Visible = True
Next k
For k = 1 To 7
   cmd32(k).Top = cmd32(0).Top
   cmd32(k).Left = cmd32(k - 1).Left + cmd32(0).Width
Next k

cmd64(0).Width = 64
For k = 1 To 3
   Load cmd64(k)
   cmd64(k).Visible = True
Next k
For k = 1 To 3
   cmd64(k).Top = cmd64(0).Top
   cmd64(k).Left = cmd64(k - 1).Left + cmd64(0).Width
Next k

cmd128(0).Width = 128
For k = 1 To 1
   Load cmd128(k)
   cmd128(k).Visible = True
Next k
For k = 1 To 1
   cmd128(k).Top = cmd128(0).Top
   cmd128(k).Left = cmd128(k - 1).Left + cmd128(0).Width
Next k
End Sub

Private Sub mnuExit1_Click()
   frmPaletteTop = frmPalette.Top
   frmPaletteLeft = frmPalette.Left
   Unload Me
End Sub


