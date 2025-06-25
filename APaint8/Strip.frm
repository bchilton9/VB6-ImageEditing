VERSION 5.00
Begin VB.Form frmStrip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Horizontal strip from circular selection"
   ClientHeight    =   2655
   ClientLeft      =   150
   ClientTop       =   105
   ClientWidth     =   5100
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPepper 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pepper"
      Height          =   255
      Left            =   1110
      TabIndex        =   15
      Top             =   1770
      Width           =   1050
   End
   Begin VB.CommandButton cmdStripAC 
      Caption         =   "Cancel"
      Height          =   285
      Index           =   1
      Left            =   2985
      TabIndex        =   11
      Top             =   2160
      Width           =   900
   End
   Begin VB.CommandButton cmdStripAC 
      Caption         =   "Accept"
      Height          =   285
      Index           =   0
      Left            =   930
      TabIndex        =   10
      Top             =   2160
      Width           =   900
   End
   Begin VB.HScrollBar HSStrip 
      Height          =   180
      Index           =   2
      LargeChange     =   10
      Left            =   3915
      Max             =   100
      Min             =   1
      TabIndex        =   9
      Top             =   1170
      Value           =   1
      Width           =   975
   End
   Begin VB.HScrollBar HSStrip 
      Height          =   180
      Index           =   1
      LargeChange     =   10
      Left            =   3915
      Max             =   360
      TabIndex        =   7
      Top             =   570
      Width           =   975
   End
   Begin VB.HScrollBar HSStrip 
      Height          =   180
      Index           =   0
      LargeChange     =   10
      Left            =   3915
      Min             =   1
      TabIndex        =   5
      Top             =   270
      Value           =   1
      Width           =   975
   End
   Begin VB.Label LabStripInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Incr reduction = .888"
      Height          =   225
      Index           =   6
      Left            =   1110
      TabIndex        =   14
      Top             =   1455
      Width           =   2175
   End
   Begin VB.Label LabStripInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Incr angle = 4.555"
      Height          =   225
      Index           =   4
      Left            =   915
      TabIndex        =   13
      Top             =   885
      Width           =   2370
   End
   Begin VB.Label LabStripInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame width =        1024"
      Height          =   225
      Index           =   0
      Left            =   375
      TabIndex        =   12
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label LabStrip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "20"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   3225
      TabIndex        =   8
      Top             =   1155
      Width           =   570
   End
   Begin VB.Label LabStrip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "20"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   6
      Top             =   555
      Width           =   570
   End
   Begin VB.Label LabStrip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "20"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   3225
      TabIndex        =   4
      Top             =   240
      Width           =   570
   End
   Begin VB.Label LabStripInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Frames "
      Height          =   225
      Index           =   2
      Left            =   2415
      TabIndex        =   3
      Top             =   225
      Width           =   555
   End
   Begin VB.Label LabStripInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Final size as % of original  "
      Height          =   225
      Index           =   5
      Left            =   1080
      TabIndex        =   2
      Top             =   1170
      Width           =   1845
   End
   Begin VB.Label LabStripInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Total angular rotation 0 - 360 "
      Height          =   225
      Index           =   3
      Left            =   795
      TabIndex        =   1
      Top             =   585
      Width           =   2175
   End
   Begin VB.Label LabStripInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Max num frames ="
      Height          =   225
      Index           =   1
      Left            =   375
      TabIndex        =   0
      Top             =   255
      Width           =   1815
   End
End
Attribute VB_Name = "frmStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmStrip.frm

Option Explicit
Option Base 1

Private Sub chkPepper_Click()
   If chkPepper Then aPepper = True Else aPepper = False
End Sub


'Public MaxNumFrames As Long
'Public NumStrips As Long
'Public zTotalAng As Single
'Public zIncrAng As Single
'Public zFinalPercentReduc As Single
'Public zIncrPercentReduc As Single
'Public canvasW As Long ''''
'Public SSW ' Circular select width
'Public MAXWIDTH

Private Sub Form_Load()
   
   frmStrip.Left = frmStripLeft
   frmStrip.Top = frmStripTop
   
   LabStripInfo(0) = "Frame width =" & Str$(SSW)
   MaxNumFrames = MAXWIDTH \ SSW
   HSStrip(0).Max = MaxNumFrames
   LabStripInfo(1) = "Max num frames =" & Str$(MaxNumFrames)
   
   NumStrips = 1
   zTotalAng = 0
   If NumStrips > 1 Then
      zIncrAng = zTotalAng / (NumStrips - 1)
   Else
      zIncrAng = 0
   End If
   LabStripInfo(4) = "Incr Ang =" & Str$(zIncrAng)
   zFinalPercentReduc = 100
   If NumStrips > 1 Then
      zIncrPercentReduc = (100 - zFinalPercentReduc) / (NumStrips - 1)
   Else
      zIncrPercentReduc = 0
   End If
   LabStripInfo(6) = "Incr Reduc =" & Str$(zIncrPercentReduc)

   HSStrip(0).Value = NumStrips
   LabStrip(0) = Str$(NumStrips)
   HSStrip(1).Value = zTotalAng
   LabStrip(1) = Str$(zTotalAng)
   HSStrip(2).Value = zFinalPercentReduc
   LabStrip(2) = Str$(zFinalPercentReduc)
   
   chkPepper_Click
End Sub

Private Sub HSStrip_Change(Index As Integer)
   Select Case Index
   Case 0
      NumStrips = HSStrip(Index).Value
      LabStrip(Index) = Str$(NumStrips)
      If NumStrips > 1 Then
         zIncrAng = zTotalAng / (NumStrips - 1)
         zIncrPercentReduc = (100 - zFinalPercentReduc) / (NumStrips - 1)
      Else
         zIncrAng = 0
         zIncrPercentReduc = 0
      End If
      LabStripInfo(4) = "Incr Ang =" & Str$(zIncrAng)
      LabStripInfo(6) = "Incr Reduc =" & Str$(zIncrPercentReduc)
   Case 1
      zTotalAng = HSStrip(Index).Value
      LabStrip(Index) = Str$(zTotalAng)
      If NumStrips > 1 Then
         zIncrAng = zTotalAng / (NumStrips - 1)
      Else
         zIncrAng = 0
      End If
      LabStripInfo(4) = "Incr Ang =" & Str$(zIncrAng)
   Case 2
      zFinalPercentReduc = HSStrip(Index).Value
      LabStrip(Index) = Str$(zFinalPercentReduc)
      If NumStrips > 1 Then
         zIncrPercentReduc = (100 - zFinalPercentReduc) / (NumStrips - 1)
      Else
         zIncrPercentReduc = 0
      End If
      LabStripInfo(6) = "Incr Reduc =" & Str$(zIncrPercentReduc)
   End Select
End Sub

Private Sub cmdStripAC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 1 Then NumStrips = 0
   frmStripLeft = frmStrip.Left
   frmStripTop = frmStrip.Top
   Unload Me
End Sub

