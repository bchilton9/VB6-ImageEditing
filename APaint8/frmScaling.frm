VERSION 5.00
Begin VB.Form frmScaling 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Scaling"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraS 
      Caption         =   "Resize canvas"
      Height          =   2295
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   1515
      Width           =   3435
      Begin VB.VScrollBar VSC 
         Height          =   255
         Index           =   3
         LargeChange     =   25
         Left            =   3105
         Max             =   25
         Min             =   800
         SmallChange     =   25
         TabIndex        =   33
         Top             =   300
         Value           =   100
         Width           =   210
      End
      Begin VB.VScrollBar VSC 
         Height          =   255
         Index           =   2
         LargeChange     =   25
         Left            =   1485
         Max             =   25
         Min             =   800
         SmallChange     =   25
         TabIndex        =   32
         Top             =   300
         Value           =   100
         Width           =   210
      End
      Begin VB.VScrollBar VSC 
         Height          =   255
         Index           =   1
         LargeChange     =   25
         Left            =   2895
         Max             =   25
         Min             =   800
         SmallChange     =   25
         TabIndex        =   24
         Top             =   300
         Value           =   100
         Width           =   210
      End
      Begin VB.VScrollBar VSC 
         Height          =   255
         Index           =   0
         LargeChange     =   25
         Left            =   1260
         Max             =   25
         Min             =   800
         SmallChange     =   25
         TabIndex        =   23
         Top             =   300
         Value           =   100
         Width           =   225
      End
      Begin VB.CommandButton cmdFixCanvasSize 
         Caption         =   "Set Canvas = Image size"
         Height          =   255
         Left            =   630
         TabIndex        =   20
         Top             =   1920
         Width           =   2100
      End
      Begin VB.Frame fraS 
         Height          =   885
         Index           =   3
         Left            =   1305
         TabIndex        =   9
         Top             =   930
         Width           =   780
         Begin VB.OptionButton optImagePos 
            Height          =   210
            Index           =   4
            Left            =   480
            TabIndex        =   14
            Top             =   570
            Width           =   210
         End
         Begin VB.OptionButton optImagePos 
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   570
            Width           =   225
         End
         Begin VB.OptionButton optImagePos 
            Height          =   210
            Index           =   2
            Left            =   285
            TabIndex        =   12
            Top             =   375
            Width           =   240
         End
         Begin VB.OptionButton optImagePos 
            Height          =   180
            Index           =   1
            Left            =   465
            TabIndex        =   11
            Top             =   195
            Width           =   270
         End
         Begin VB.OptionButton optImagePos 
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   195
            Value           =   -1  'True
            Width           =   240
         End
      End
      Begin VB.Label LabC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2415
         TabIndex        =   28
         Top             =   285
         Width           =   900
      End
      Begin VB.Label LabC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   735
         TabIndex        =   27
         Top             =   285
         Width           =   960
      End
      Begin VB.Label LabS 
         Caption         =   "If canvas = image size then position = top left"
         Height          =   765
         Index           =   11
         Left            =   2220
         TabIndex        =   19
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label LabS 
         Caption         =   "Min/Max = 4 / 1024"
         Height          =   270
         Index           =   10
         Left            =   1770
         TabIndex        =   18
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label LabS 
         Caption         =   "Min/Max= 4 / 1024"
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label LabS 
         Caption         =   "Canvas Width"
         Height          =   450
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   600
      End
      Begin VB.Label LabS 
         Caption         =   "Canvas Height"
         Height          =   450
         Index           =   4
         Left            =   1770
         TabIndex        =   7
         Top             =   225
         Width           =   600
      End
      Begin VB.Label LabS 
         Caption         =   "Image position"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1215
         Width           =   1080
      End
   End
   Begin VB.Frame fraS 
      Caption         =   "Resize image"
      Height          =   1365
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   30
      Width           =   3420
      Begin VB.VScrollBar VSI 
         Height          =   255
         Index           =   3
         LargeChange     =   4
         Left            =   3135
         Max             =   4
         Min             =   512
         SmallChange     =   4
         TabIndex        =   31
         Top             =   315
         Value           =   4
         Width           =   195
      End
      Begin VB.VScrollBar VSI 
         Height          =   255
         Index           =   2
         LargeChange     =   4
         Left            =   1440
         Max             =   4
         Min             =   512
         SmallChange     =   4
         TabIndex        =   30
         Top             =   315
         Value           =   4
         Width           =   195
      End
      Begin VB.CommandButton cmdFixImageSize 
         Caption         =   "Set Image = Canvas size"
         Height          =   240
         Left            =   525
         TabIndex        =   29
         Top             =   1005
         Width           =   2430
      End
      Begin VB.VScrollBar VSI 
         Height          =   255
         Index           =   1
         LargeChange     =   4
         Left            =   2925
         Max             =   4
         Min             =   512
         SmallChange     =   4
         TabIndex        =   22
         Top             =   315
         Value           =   4
         Width           =   195
      End
      Begin VB.VScrollBar VSI 
         Height          =   255
         Index           =   0
         LargeChange     =   4
         Left            =   1230
         Max             =   4
         Min             =   768
         SmallChange     =   4
         TabIndex        =   21
         Top             =   300
         Value           =   4
         Width           =   195
      End
      Begin VB.Label LabI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   26
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label LabI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   660
         TabIndex        =   25
         Top             =   285
         Width           =   990
      End
      Begin VB.Label LabS 
         Caption         =   "Min/Max = 4 / 1024"
         Height          =   270
         Index           =   9
         Left            =   1740
         TabIndex        =   17
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label LabS 
         Caption         =   "Min/Max = 4 / 1024"
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   15
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label LabS 
         Caption         =   "Height"
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   4
         Top             =   315
         Width           =   525
      End
      Begin VB.Label LabS 
         Caption         =   "Width"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdAccCan 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   2055
      TabIndex        =   1
      Top             =   3990
      Width           =   1155
   End
   Begin VB.CommandButton cmdAccCan 
      Caption         =   "Accept"
      Height          =   345
      Index           =   0
      Left            =   405
      TabIndex        =   0
      Top             =   3975
      Width           =   1170
   End
End
Attribute VB_Name = "frmScaling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmScalimg (frmScaling.frm)

Option Explicit
Option Base 1

' Input values
'Public picW As Long
'Public picH As Long
'Public canvasW As Long
'Public canvasH As Long
'Public ImagePosition As Long

' Returned values
'Public RpicW As Long
'Public RpicH As Long
'Public RcanvasW As Long
'Public RcanvasH As Long
'Public RImagePosition As Long
'Public aScaleChange As Boolean
Dim aScroll As Boolean
Dim prevVSIW As Long
Dim prevVSIH As Long
Dim prevVSCW As Long
Dim prevVSCH As Long

Private Sub Form_Load()
Dim k As Long

   LabS(6) = "Min/Max = 4/" & Trim$(Str$(MAXWIDTH))
   LabS(9) = "Min/Max = 4/" & Trim$(Str$(MAXHEIGHT))
   LabS(8) = "Min/Max = 4/" & Trim$(Str$(MAXWIDTH))
   LabS(10) = "Min/Max = 4/" & Trim$(Str$(MAXHEIGHT))

   ' Set up limits
   aScroll = False
   ' picW  [+/-4]
   VSI(0).Min = MAXHEIGHT
   VSI(0).Max = 4
   VSI(0).LargeChange = 4
   VSI(0).SmallChange = 4
   ' picH  [+/-4]
   VSI(1).Min = MAXWIDTH
   VSI(1).Max = 4
   VSI(1).LargeChange = 4
   VSI(1).SmallChange = 4
   
   ' picW  [+/-12]
   VSI(2).Min = MAXHEIGHT
   VSI(2).Max = 4
   VSI(2).LargeChange = 12
   VSI(2).SmallChange = 12
   ' picH  [+/-12]
   VSI(3).Min = MAXWIDTH
   VSI(3).Max = 4
   VSI(3).LargeChange = 12
   VSI(3).SmallChange = 12
   
   ' canvasW [+/-4]
   VSC(0).Min = MAXHEIGHT
   VSC(0).Max = 4
   VSC(0).LargeChange = 4
   VSC(0).SmallChange = 4
   ' canvasH  [+/-4]
   VSC(1).Min = MAXWIDTH
   VSC(1).Max = 4
   VSC(1).LargeChange = 4
   VSC(1).SmallChange = 4
   
   ' canvasW  [+/-12]
   VSC(2).Min = MAXHEIGHT
   VSC(2).Max = 4
   VSC(2).LargeChange = 12
   VSC(2).SmallChange = 12
   ' canvasH  [+/-10]
   VSC(3).Min = MAXWIDTH
   VSC(3).Max = 4
   VSC(3).LargeChange = 12
   VSC(3).SmallChange = 12
   
   
   ' Make = to input values
   optImagePos(ImagePosition).Value = True
   RpicW = picW
   RpicH = picH
   RcanvasW = canvasW
   RcanvasH = canvasH
   RImagePosition = ImagePosition
   
   ' Position form at last position
   frmScaling.Top = frmScalingTop
   frmScaling.Left = frmScalingLeft
   
   aScaleChange = True
   
   SetValues
   aScroll = True
End Sub

Private Sub cmdFixCanvasSize_Click()
   RpicW = VSI(0).Value
   RpicH = VSI(1).Value
   RcanvasW = RpicW
   RcanvasH = RpicH
   aScroll = False
   SetValues
   aScroll = True
End Sub

Private Sub cmdFixImageSize_Click()
   RcanvasW = VSC(0).Value
   RcanvasH = VSC(1).Value
   RpicW = RcanvasW
   RpicH = RcanvasH
   aScroll = False
   SetValues
   aScroll = True
End Sub

Private Sub cmdAccCan_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim remainder As Long
Dim iopt As Long
   Select Case Index
   Case 0   ' Accept
      ' Get RImagePosition
      For iopt = 0 To 4
         If optImagePos(iopt).Value = True Then
            RImagePosition = iopt
            Exit For
         End If
      Next iopt
      
      ' Check if Accept pressed with no change
      aScaleChange = False
      If RpicW <> picW Then aScaleChange = True
      If RpicH <> picH Then aScaleChange = True
      If RcanvasW <> canvasW Then aScaleChange = True
      If RcanvasH <> canvasH Then aScaleChange = True
   Case 1   ' Cancel
      aScaleChange = False
   End Select
   
   frmScalingTop = frmScaling.Top
   frmScalingLeft = frmScaling.Left
   
   Unload Me
End Sub

Private Sub SetValues()
   ' Image
   VSI(0).Value = RpicW
   LabI(0) = Str$(RpicW)
   VSI(1).Value = RpicH
   LabI(1) = Str$(RpicH)
   
   VSI(2).Value = RpicW
   prevVSIW = RpicW
   VSI(3).Value = RpicH
   prevVSIH = RpicH
   
   ' Canvas
   VSC(0).Value = RcanvasW
   LabC(0) = Str$(RcanvasW)
   VSC(1).Value = RcanvasH
   LabC(1) = Str$(RcanvasH)
   
   VSC(2).Value = RcanvasW
   prevVSCW = RcanvasW
   VSC(3).Value = RcanvasH
   prevVSCH = RcanvasH
   
   optImagePos(RImagePosition).Value = True
End Sub

Private Sub VSI_Change(Index As Integer)
' Image size
   If aScroll Then
      'VSCROLLER VSI,Index
      Select Case Index
      Case 0
         RpicW = VSI(0).Value
         If RpicW > RcanvasW Then
            RcanvasW = RpicW
            VSC(0).Value = RcanvasW
            LabC(0) = Str$(RcanvasW)
         End If
         LabI(0) = Str$(RpicW)
      Case 1
         RpicH = VSI(1).Value
         If RpicH > RcanvasH Then
            RcanvasH = RpicH
            VSC(1).Value = RcanvasH
         End If
         LabI(1) = Str$(RpicH)
      
      Case 2  ' picW [+/-12]
         If VSI(2).Value > prevVSIW Then
            If VSI(0).Value + VSI(2).SmallChange <= VSI(0).Min Then
               VSI(0).Value = VSI(0).Value + VSI(2).SmallChange '+ 12
            End If
         ElseIf VSI(2).Value < prevVSIW Then
            If VSI(0).Value - VSI(2).SmallChange >= VSI(0).Max Then
               VSI(0).Value = VSI(0).Value - VSI(2).SmallChange '- 12
            End If
         End If
         prevVSIW = VSI(2).Value
         RpicW = VSI(0).Value
         If RpicW > RcanvasW Then
            RcanvasW = RpicW
            VSC(0).Value = RcanvasW
            LabC(0) = Str$(RcanvasW)
         End If
         LabI(0) = Str$(RpicW)
      Case 3   ' picH [+/-10]
         If VSI(3).Value > prevVSIH Then
            If VSI(1).Value + VSI(3).SmallChange <= VSI(1).Min Then
               VSI(1).Value = VSI(1).Value + VSI(3).SmallChange '+ 10
            End If
         ElseIf VSI(3).Value < prevVSIH Then
            If VSI(1).Value - VSI(3).SmallChange >= VSI(1).Max Then
               VSI(1).Value = VSI(1).Value - VSI(3).SmallChange '- 10
            End If
         End If
         prevVSIH = VSI(3).Value
         RpicH = VSI(1).Value
         If RpicH > RcanvasH Then
            RcanvasH = RpicH
            VSC(1).Value = RcanvasH
         End If
         LabI(1) = Str$(RpicH)
      End Select
   End If
End Sub

Private Sub VSC_Change(Index As Integer)
' Canvas size
   If aScroll Then
      'VSCROLLER VSC,Index
      Select Case Index
      Case 0
         RcanvasW = VSC(0).Value
         If RcanvasW < RpicW Then
            RpicW = RcanvasW
            VSI(0).Value = RcanvasW
            LabI(0) = Str$(RcanvasW)
         End If
         LabC(0) = Str$(RcanvasW)
      Case 1
         RcanvasH = VSC(1).Value
         If RcanvasH < RpicH Then
            RpicH = RcanvasH
            VSI(1).Value = RcanvasH
            LabI(1) = Str$(RcanvasH)
         End If
         LabC(1) = Str$(RcanvasH)
      
      Case 2  ' canvasW [+/-12]
         If VSC(2).Value > prevVSCW Then
            If VSC(0).Value + VSC(2).SmallChange <= VSC(0).Min Then
               VSC(0).Value = VSC(0).Value + VSC(2).SmallChange '+ 12
            End If
         ElseIf VSC(2).Value < prevVSCW Then
            If VSC(0).Value - VSC(2).SmallChange >= VSC(0).Max Then
               VSC(0).Value = VSC(0).Value - VSC(2).SmallChange '- 12
            End If
         End If
         prevVSCW = VSC(2).Value
         RcanvasW = VSC(0).Value
         If RcanvasW < RpicW Then
            RpicW = RcanvasW
            VSI(0).Value = RcanvasW
            LabI(0) = Str$(RcanvasW)
         End If
         LabC(0) = Str$(RcanvasW)
      Case 3   ' canvasH [+/-10]
         If VSC(3).Value > prevVSCH Then
            If VSC(1).Value + VSC(3).SmallChange <= VSC(1).Min Then
               VSC(1).Value = VSC(1).Value + VSC(3).SmallChange '+ 10
            End If
         ElseIf VSC(3).Value < prevVSCH Then
            If VSC(1).Value - VSC(3).SmallChange >= VSC(1).Max Then
               VSC(1).Value = VSC(1).Value - VSC(3).SmallChange '- 10
            End If
         End If
         prevVSCH = VSC(3).Value
         RcanvasH = VSC(1).Value
         If RcanvasH < RpicH Then
            RpicH = RcanvasH
            VSI(1).Value = RcanvasH
            LabI(1) = Str$(RcanvasH)
         End If
         LabC(1) = Str$(RcanvasH)
      End Select
   End If
End Sub

