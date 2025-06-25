VERSION 5.00
Begin VB.Form frmSelector 
   Caption         =   " Selector"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   210
      Index           =   1
      Left            =   3420
      TabIndex        =   4
      Top             =   60
      Width           =   240
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   210
      Index           =   0
      Left            =   15
      TabIndex        =   3
      Top             =   5055
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Brush"
      Height          =   4980
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   675
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   135
         Picture         =   "frmSelector.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   660
         Width           =   375
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   135
         Picture         =   "frmSelector.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  Windows API to make application stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H2

Private Sub cmdExit_Click(Index As Integer)
   frmSelectorLeft = frmSelector.Left
   frmSelectorTop = frmSelector.Top
   Unload frmSelector
End Sub

Private Sub Form_Load()
Dim retval As Long
   
   frmSelector.Left = frmSelectorLeft
   frmSelector.Top = frmSelectorTop
   
   
   ' Size & Make frmZoom stay on top
   retval = SetWindowPos(frmSelector.hWnd, hWndInsertAfter, frmSelectorLeft, frmSelectorTop, _
   300, 400, wFlags)

End Sub

Private Sub optBrush_Click(Index As Integer)
   Select Case Index
   Case 0
      Form1.optTools(0).Picture = optBrush(0).Picture
   Case 1
      Form1.optTools(0).Picture = optBrush(1).Picture
   
   End Select
End Sub
