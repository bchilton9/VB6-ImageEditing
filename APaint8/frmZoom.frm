VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Zoom"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optZoom 
      Caption         =   "x 12"
      Height          =   285
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3930
      Width           =   480
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "x 8  "
      Height          =   285
      Index           =   3
      Left            =   2115
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3930
      Width           =   465
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "x 6"
      Height          =   285
      Index           =   2
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3930
      Width           =   435
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "x 4"
      Height          =   285
      Index           =   1
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3915
      Width           =   465
   End
   Begin VB.OptionButton optZoom 
      Caption         =   "x 2"
      Height          =   285
      Index           =   0
      Left            =   630
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3915
      Width           =   405
   End
   Begin VB.CommandButton cmdExitZoom 
      Caption         =   "X"
      Height          =   210
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   30
      Width           =   195
   End
   Begin VB.CommandButton cmdExitZoom 
      Caption         =   "X"
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   3900
      Width           =   300
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   90
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   255
      Width           =   3600
      Begin VB.Line LZV 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         Index           =   0
         X1              =   33
         X2              =   33
         Y1              =   43
         Y2              =   102
      End
      Begin VB.Line LZH 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         Index           =   0
         X1              =   34
         X2              =   89
         Y1              =   23
         Y2              =   23
      End
      Begin VB.Shape ShapeC 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         Height          =   270
         Left            =   1125
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Label LabXY 
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   30
      Width           =   1440
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmZoom.frm

Option Explicit
Option Base 1

'  Windows API to make form stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H2
Private k As Long

Private Sub cmdExitZoom_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      frmZoomLeft = frmZoom.Left
      frmZoomTop = frmZoom.Top
      aZoom = False
      For k = 1 To 19
         Unload LZH(k)
         Unload LZV(k)
      Next k
      aZoom = False
      SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
      Unload frmZoom
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Button As Integer
Dim X As Single, Y As Single
   MouseKeys KeyCode, Shift, Button, X, Y
End Sub

Private Sub Form_Load()
Dim retval As Long
   frmZoom.Left = frmZoomLeft
   frmZoom.Top = frmZoomTop
   picZoom.Width = ZoomSize
   picZoom.Height = ZoomSize
   
   ' Size & Make frmZoom stay on top
   retval = SetWindowPos(frmZoom.hWnd, hWndInsertAfter, frmZoomLeft, frmZoomTop, _
   picZoom.Width + 17, picZoom.Height + 72, wFlags)
   
   Form1.PIC.SetFocus      ' NB to stop Arrows tabbing

   KeyPreview = True
   
   ' Grid
   LZH(0).Visible = False
   LZV(0).Visible = False
   For k = 1 To 19
      Load LZH(k)
      Load LZV(k)
   Next k
   For k = 0 To 19
      LZV(k).BorderColor = RGB(180, 180, 180)
      LZV(k).x1 = 12 * k
      LZV(k).x2 = 12 * k
      LZV(k).y1 = 0
      LZV(k).y2 = ZoomSize
      
      LZH(k).BorderColor = RGB(180, 180, 180)
      LZH(k).x1 = 0
      LZH(k).x2 = ZoomSize
      LZH(k).y1 = 12 * k
      LZH(k).y2 = 12 * k
   Next k
   If ZoomFactor = 12 Then VisGrid
   
   Select Case ZoomFactor
   Case 2: optZoom(0).Value = True
   Case 4: optZoom(1).Value = True
   Case 6: optZoom(2).Value = True
   Case 8: optZoom(3).Value = True
   Case 12: optZoom(4).Value = True
   End Select
   
   ShapeC.BackStyle = vbTransparent
End Sub

Private Sub VisGrid()
   For k = 0 To 19
      LZV(k).Visible = True
      LZH(k).Visible = True
   Next k
End Sub

Private Sub InVisGrid()
   For k = 0 To 19
      LZV(k).Visible = False
      LZH(k).Visible = False
   Next k
End Sub

Private Sub optZoom_Click(Index As Integer)
   Form1.PIC.SetFocus      ' NB to stop Arrows tabbing
   Select Case Index
   Case 0: ZoomFactor = 2
      InVisGrid
   Case 1: ZoomFactor = 4
      InVisGrid
   Case 2: ZoomFactor = 6
      InVisGrid
   Case 3: ZoomFactor = 8
      InVisGrid
   Case 4: ZoomFactor = 12
      VisGrid
   End Select
End Sub

