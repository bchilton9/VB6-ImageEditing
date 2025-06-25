VERSION 5.00
Begin VB.Form frmToolOptions 
   Caption         =   " Tool Options"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTO 
      Caption         =   "Arrows"
      Height          =   615
      Index           =   16
      Left            =   3960
      TabIndex        =   182
      Top             =   5985
      Width           =   1230
      Begin VB.OptionButton optArrows 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optArrows 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":00D2
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optArrows 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":01A4
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Bushes "
      Height          =   1350
      Index           =   15
      Left            =   3975
      TabIndex        =   172
      Top             =   4575
      Width           =   1200
      Begin VB.CommandButton cmdSwapAngles 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   181
         ToolTipText     =   " Swap tilt "
         Top             =   945
         Width           =   285
      End
      Begin VB.CommandButton cmdSwapAngles 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   " Swap tilt "
         Top             =   615
         Width           =   285
      End
      Begin VB.CommandButton cmdSwapAngles 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         Height          =   255
         Index           =   0
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   179
         ToolTipText     =   " Swap tilt "
         Top             =   285
         Width           =   285
      End
      Begin VB.OptionButton optTrees 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   135
         Picture         =   "frmToolOptions.frx":0276
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   915
         Width           =   300
      End
      Begin VB.OptionButton optTrees 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   135
         Picture         =   "frmToolOptions.frx":03C0
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   600
         Width           =   300
      End
      Begin VB.OptionButton optTrees 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   135
         Picture         =   "frmToolOptions.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   270
         Width           =   300
      End
      Begin VB.Label LabLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   510
         TabIndex        =   178
         ToolTipText     =   "  Bush size LC/RC "
         Top             =   945
         Width           =   255
      End
      Begin VB.Label LabLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   510
         TabIndex        =   177
         ToolTipText     =   "  Bush size LC/RC "
         Top             =   615
         Width           =   255
      End
      Begin VB.Label LabLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   510
         TabIndex        =   176
         ToolTipText     =   "  Bush size LC/RC "
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Brushes "
      Height          =   1605
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   1230
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   150
         Picture         =   "frmToolOptions.frx":0654
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   270
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   480
         Picture         =   "frmToolOptions.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   270
         Width           =   285
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   795
         Picture         =   "frmToolOptions.frx":08E8
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   270
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   795
         Picture         =   "frmToolOptions.frx":0A32
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   585
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   480
         Picture         =   "frmToolOptions.frx":0B7C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   585
         Width           =   285
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   150
         Picture         =   "frmToolOptions.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   585
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   795
         Picture         =   "frmToolOptions.frx":0E10
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1215
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   480
         Picture         =   "frmToolOptions.frx":0F5A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1215
         Width           =   285
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   135
         Picture         =   "frmToolOptions.frx":10A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1215
         Width           =   315
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   795
         Picture         =   "frmToolOptions.frx":11EE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   900
         Width           =   300
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   480
         Picture         =   "frmToolOptions.frx":1338
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   900
         Width           =   285
      End
      Begin VB.OptionButton optBrushes 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   135
         Picture         =   "frmToolOptions.frx":1482
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   900
         Width           =   315
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Radials "
      Height          =   1920
      Index           =   14
      Left            =   4080
      TabIndex        =   139
      Top             =   2580
      Width           =   975
      Begin VB.OptionButton optRadials 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   105
         Picture         =   "frmToolOptions.frx":15CC
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   1545
         Width           =   300
      End
      Begin VB.OptionButton optRadials 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   105
         Picture         =   "frmToolOptions.frx":169E
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   1230
         Width           =   300
      End
      Begin VB.OptionButton optRadials 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   105
         Picture         =   "frmToolOptions.frx":1770
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   915
         Width           =   300
      End
      Begin VB.OptionButton optRadials 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   105
         Picture         =   "frmToolOptions.frx":1842
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   600
         Width           =   300
      End
      Begin VB.OptionButton optRadials 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   105
         Picture         =   "frmToolOptions.frx":1914
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   285
         Width           =   300
      End
      Begin VB.Label LabRadial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   480
         TabIndex        =   152
         ToolTipText     =   " # teeth/2  LC/RC "
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label LabRadial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   480
         TabIndex        =   150
         ToolTipText     =   " # sides LC/RC "
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label LabRadial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   480
         TabIndex        =   149
         ToolTipText     =   " # circs LC/RC "
         Top             =   945
         Width           =   345
      End
      Begin VB.Label LabRadial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   148
         ToolTipText     =   " # points LC/RC "
         Top             =   630
         Width           =   345
      End
      Begin VB.Label LabRadial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   480
         TabIndex        =   147
         ToolTipText     =   " # spokes LC/RC "
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Arcs"
      Height          =   1260
      Index           =   13
      Left            =   3810
      TabIndex        =   124
      Top             =   315
      Width           =   1545
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   1110
         Picture         =   "frmToolOptions.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   780
         Picture         =   "frmToolOptions.frx":1B30
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   450
         Picture         =   "frmToolOptions.frx":1C02
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   120
         Picture         =   "frmToolOptions.frx":1CD4
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   1110
         Picture         =   "frmToolOptions.frx":1DA6
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   780
         Picture         =   "frmToolOptions.frx":1E78
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   450
         Picture         =   "frmToolOptions.frx":1F4A
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   120
         Picture         =   "frmToolOptions.frx":201C
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   1110
         Picture         =   "frmToolOptions.frx":20EE
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":21C0
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":2292
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optArcs 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":2364
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Fills "
      Height          =   975
      Index           =   12
      Left            =   105
      TabIndex        =   113
      Top             =   5625
      Width           =   3675
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   21
         Left            =   3240
         Picture         =   "frmToolOptions.frx":2436
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   20
         Left            =   2925
         Picture         =   "frmToolOptions.frx":2580
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   19
         Left            =   2610
         Picture         =   "frmToolOptions.frx":26CA
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   18
         Left            =   2295
         Picture         =   "frmToolOptions.frx":279C
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   17
         Left            =   1980
         Picture         =   "frmToolOptions.frx":286E
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   16
         Left            =   1665
         Picture         =   "frmToolOptions.frx":2940
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   15
         Left            =   1350
         Picture         =   "frmToolOptions.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   14
         Left            =   1035
         Picture         =   "frmToolOptions.frx":2AE4
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   13
         Left            =   705
         Picture         =   "frmToolOptions.frx":2BB6
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   12
         Left            =   390
         Picture         =   "frmToolOptions.frx":2C88
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   75
         Picture         =   "frmToolOptions.frx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   3240
         Picture         =   "frmToolOptions.frx":2E2C
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   2925
         Picture         =   "frmToolOptions.frx":2EFE
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   2610
         Picture         =   "frmToolOptions.frx":2FD0
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   2280
         Picture         =   "frmToolOptions.frx":30A2
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   1965
         Picture         =   "frmToolOptions.frx":3174
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   1650
         Picture         =   "frmToolOptions.frx":3246
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   1335
         Picture         =   "frmToolOptions.frx":3318
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   1020
         Picture         =   "frmToolOptions.frx":33EA
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   705
         Picture         =   "frmToolOptions.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   390
         Picture         =   "frmToolOptions.frx":358E
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optFills 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   75
         Picture         =   "frmToolOptions.frx":3660
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Shapes"
      Height          =   930
      Index           =   11
      Left            =   3960
      TabIndex        =   108
      Top             =   1605
      Width           =   1215
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":3732
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   435
         Picture         =   "frmToolOptions.frx":3804
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   105
         Picture         =   "frmToolOptions.frx":38D6
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":39A8
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   435
         Picture         =   "frmToolOptions.frx":3A7A
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optShapes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   105
         Picture         =   "frmToolOptions.frx":3B4C
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Junctions "
      Height          =   1290
      Index           =   10
      Left            =   2595
      TabIndex        =   105
      Top             =   4275
      Width           =   1185
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   765
         Picture         =   "frmToolOptions.frx":3C1E
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   435
         Picture         =   "frmToolOptions.frx":3CF0
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   90
         Picture         =   "frmToolOptions.frx":3DC2
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":3E94
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   435
         Picture         =   "frmToolOptions.frx":3F66
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   105
         Picture         =   "frmToolOptions.frx":4038
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":410A
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":41DC
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optJunctions 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":42AE
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Bullets"
      Height          =   600
      Index           =   9
      Left            =   2595
      TabIndex        =   101
      Top             =   3645
      Width           =   1185
      Begin VB.OptionButton optBullets 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   795
         Picture         =   "frmToolOptions.frx":4380
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optBullets 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   465
         Picture         =   "frmToolOptions.frx":44CA
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optBullets 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   135
         Picture         =   "frmToolOptions.frx":4614
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Tubes "
      Height          =   600
      Index           =   8
      Left            =   2595
      TabIndex        =   97
      Top             =   3030
      Width           =   1185
      Begin VB.OptionButton optTubes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":46E6
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optTubes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":4830
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   225
         Width           =   300
      End
      Begin VB.OptionButton optTubes 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":497A
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Cones "
      Height          =   1035
      Index           =   7
      Left            =   2595
      TabIndex        =   93
      Top             =   1995
      Width           =   1185
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":4A4C
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   450
         Picture         =   "frmToolOptions.frx":4B1E
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   120
         Picture         =   "frmToolOptions.frx":4C68
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":4DB2
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":4EFC
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optCones 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":5046
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Cirllipses "
      Height          =   1605
      Index           =   6
      Left            =   2595
      TabIndex        =   80
      Top             =   330
      Width           =   1170
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   750
         Picture         =   "frmToolOptions.frx":5118
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   420
         Picture         =   "frmToolOptions.frx":5262
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   90
         Picture         =   "frmToolOptions.frx":57EC
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   750
         Picture         =   "frmToolOptions.frx":5936
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   420
         Picture         =   "frmToolOptions.frx":5A08
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   90
         Picture         =   "frmToolOptions.frx":5ADA
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   750
         Picture         =   "frmToolOptions.frx":5BAC
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   420
         Picture         =   "frmToolOptions.frx":5C7E
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   90
         Picture         =   "frmToolOptions.frx":5D50
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   750
         Picture         =   "frmToolOptions.frx":5E22
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   420
         Picture         =   "frmToolOptions.frx":5EF4
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optCirllipses 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   90
         Picture         =   "frmToolOptions.frx":5FC6
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   255
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Rectangles "
      Height          =   1950
      Index           =   5
      Left            =   1350
      TabIndex        =   67
      Top             =   3615
      Width           =   1215
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   14
         Left            =   795
         Picture         =   "frmToolOptions.frx":6098
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   1530
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   13
         Left            =   450
         Picture         =   "frmToolOptions.frx":616A
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   1530
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   12
         Left            =   105
         Picture         =   "frmToolOptions.frx":62B4
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   1530
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   780
         Picture         =   "frmToolOptions.frx":63FE
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   435
         Picture         =   "frmToolOptions.frx":6548
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   105
         Picture         =   "frmToolOptions.frx":6692
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   780
         Picture         =   "frmToolOptions.frx":67DC
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   435
         Picture         =   "frmToolOptions.frx":68AE
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   105
         Picture         =   "frmToolOptions.frx":6980
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   870
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":6A52
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   435
         Picture         =   "frmToolOptions.frx":6B24
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   105
         Picture         =   "frmToolOptions.frx":6BF6
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   555
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":6CC8
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   435
         Picture         =   "frmToolOptions.frx":6D9A
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optRectangles 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   105
         Picture         =   "frmToolOptions.frx":6E6C
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "CurvyLines "
      Height          =   1605
      Index           =   4
      Left            =   1350
      TabIndex        =   54
      Top             =   1995
      Width           =   1215
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   780
         Picture         =   "frmToolOptions.frx":6F3E
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   450
         Picture         =   "frmToolOptions.frx":7088
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   120
         Picture         =   "frmToolOptions.frx":71D2
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   780
         Picture         =   "frmToolOptions.frx":731C
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   450
         Picture         =   "frmToolOptions.frx":73EE
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   120
         Picture         =   "frmToolOptions.frx":74C0
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":7592
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   450
         Picture         =   "frmToolOptions.frx":7664
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   120
         Picture         =   "frmToolOptions.frx":7736
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":7808
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":78DA
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optCurvyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":79AC
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   255
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "PolyLines "
      Height          =   1605
      Index           =   3
      Left            =   1335
      TabIndex        =   41
      Top             =   330
      Width           =   1215
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   11
         Left            =   780
         Picture         =   "frmToolOptions.frx":7A7E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   10
         Left            =   450
         Picture         =   "frmToolOptions.frx":7BC8
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   9
         Left            =   120
         Picture         =   "frmToolOptions.frx":7D12
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1200
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   8
         Left            =   780
         Picture         =   "frmToolOptions.frx":7E5C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   7
         Left            =   450
         Picture         =   "frmToolOptions.frx":7F2E
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   6
         Left            =   120
         Picture         =   "frmToolOptions.frx":8000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   885
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   5
         Left            =   780
         Picture         =   "frmToolOptions.frx":80D2
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   4
         Left            =   450
         Picture         =   "frmToolOptions.frx":81A4
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   3
         Left            =   120
         Picture         =   "frmToolOptions.frx":8276
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   570
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   2
         Left            =   780
         Picture         =   "frmToolOptions.frx":8348
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   1
         Left            =   450
         Picture         =   "frmToolOptions.frx":841A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton optPolyLines 
         BackColor       =   &H0080C0FF&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmToolOptions.frx":84EC
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   255
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Lines "
      Height          =   1950
      Index           =   2
      Left            =   90
      TabIndex        =   22
      Top             =   3615
      Width           =   1230
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   795
         Picture         =   "frmToolOptions.frx":85BE
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1590
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   465
         Picture         =   "frmToolOptions.frx":8708
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1590
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   135
         Picture         =   "frmToolOptions.frx":8852
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1590
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   795
         Picture         =   "frmToolOptions.frx":899C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1320
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   465
         Picture         =   "frmToolOptions.frx":8A6E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1320
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   135
         Picture         =   "frmToolOptions.frx":8B40
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1320
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   795
         Picture         =   "frmToolOptions.frx":8C12
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1050
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   465
         Picture         =   "frmToolOptions.frx":8CE4
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1050
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   135
         Picture         =   "frmToolOptions.frx":8DB6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1050
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   795
         Picture         =   "frmToolOptions.frx":8E88
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   780
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   465
         Picture         =   "frmToolOptions.frx":8F5A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   780
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   135
         Picture         =   "frmToolOptions.frx":902C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   780
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   795
         Picture         =   "frmToolOptions.frx":90FE
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   510
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   465
         Picture         =   "frmToolOptions.frx":91D0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   510
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   135
         Picture         =   "frmToolOptions.frx":92A2
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   510
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   795
         Picture         =   "frmToolOptions.frx":9374
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   465
         Picture         =   "frmToolOptions.frx":94BE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   300
      End
      Begin VB.OptionButton optLines 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   135
         Picture         =   "frmToolOptions.frx":9608
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fraTO 
      Caption         =   "Sprays "
      Height          =   1605
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   1995
      Width           =   1245
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   780
         Picture         =   "frmToolOptions.frx":9752
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   1215
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   465
         Picture         =   "frmToolOptions.frx":9824
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   1215
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   135
         Picture         =   "frmToolOptions.frx":98F6
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   1215
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   780
         Picture         =   "frmToolOptions.frx":99C8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   900
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   465
         Picture         =   "frmToolOptions.frx":9B12
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   900
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   135
         Picture         =   "frmToolOptions.frx":9C5C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   900
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   795
         Picture         =   "frmToolOptions.frx":9DA6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   585
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   465
         Picture         =   "frmToolOptions.frx":9EF0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   585
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   135
         Picture         =   "frmToolOptions.frx":A03A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   585
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   795
         Picture         =   "frmToolOptions.frx":A184
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   465
         Picture         =   "frmToolOptions.frx":A2CE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   300
      End
      Begin VB.OptionButton optSprays 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   135
         Picture         =   "frmToolOptions.frx":A418
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   285
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   5355
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   6690
      Width           =   5340
   End
   Begin VB.Label Label1 
      Caption         =   "NB Shaded Tools use Left && Right colors"
      Height          =   255
      Left            =   165
      TabIndex        =   195
      Top             =   7050
      Width           =   4995
   End
End
Attribute VB_Name = "frmToolOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmToolsOptions.frm

Option Explicit
Option Base 1

'  Windows API to make form stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H2

Private Sub cmdExit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmToolOptionsLeft = frmToolOptions.Left
   frmToolOptionsTop = frmToolOptions.Top
   ShowInstructions Index
   SetCursorPos Form1.Left \ STX + 230, Form1.Top \ STY + 86 + ExtraHeight
   Form1.PIC.SetFocus
   Unload frmToolOptions
End Sub

Private Sub cmdSwapAngles_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim zA As Single
   zA = zAngP(Index + 1)
   zAngP(Index + 1) = zAngN(Index + 1)
   zAngN(Index + 1) = zA
   If zAngP(Index + 1) > 0 Then
      cmdSwapAngles(Index).Caption = ">"
   Else
      cmdSwapAngles(Index).Caption = "<"
   End If
End Sub

Private Sub Form_Load()
Dim k As Long
   frmToolOptions.Left = frmToolOptionsLeft
   frmToolOptions.Top = frmToolOptionsTop
   
   ' Size & Make frmZoom stay on top
   k = SetWindowPos(frmToolOptions.hWnd, hWndInsertAfter, frmToolOptionsLeft, frmToolOptionsTop, _
   5595 \ STX, 7900 \ STY, wFlags)
   
   With Form1
      .optTools(Brush).Picture = optBrushes(BrushType).Picture
      .optTools(Spray).Picture = optSprays(SprayType).Picture
      .optTools(ALine).Picture = optLines(LineType).Picture
      .optTools(PolyLine).Picture = optPolyLines(PolyLineType).Picture
      .optTools(CurvyLine).Picture = optCurvyLines(CurvyLineType).Picture
      .optTools(Rectangle).Picture = optRectangles(RectangleType).Picture
      .optTools(Cirllipse).Picture = optCirllipses(CirllipseType).Picture
      .optTools(Cone).Picture = optCones(ConeType).Picture
      .optTools(Tube).Picture = optTubes(TubeType).Picture
      .optTools(Bullet).Picture = optBullets(BulletType).Picture
      .optTools(Junction).Picture = optJunctions(JunctionType).Picture
      .optTools(Arc).Picture = optArcs(ArcType).Picture
      .optTools(Shape).Picture = optShapes(ShapeType).Picture
      .optTools(Radial).Picture = optRadials(RadialType).Picture
      .optTools(AFill).Picture = optFills(FillType).Picture
      .optTools(Tree).Picture = optTrees(TreeType).Picture
      .optTools(Arrow).Picture = optArrows(ArrowType).Picture
   End With
   ' From Form1 SetUpInitialDrawTools
   If ToolType = -1 Then
      optBrushes_Click CInt(BrushType)
      optSprays_Click CInt(SprayType)
      optLines_Click CInt(LineType)
      optPolyLines_Click CInt(PolyLineType)
      optCurvyLines_Click CInt(CurvyLineType)
      optRectangles_Click CInt(RectangleType)
      optCirllipses_Click CInt(CirllipseType)
      optCones_Click CInt(ConeType)
      optTubes_Click CInt(TubeType)
      optBullets_Click CInt(BulletType)
      optJunctions_Click CInt(JunctionType)
      optArcs_Click CInt(ArcType)
      optShapes_Click CInt(ShapeType)
      optRadials_Click CInt(RadialType)
      optFills_Click CInt(FillType)
      optTrees_Click CInt(TreeType)
      optArrows_Click CInt(ArrowType)
   End If
   
   For k = 0 To 4
      LabRadial(k) = Str$(RadialRep(k + 1))
   Next k
   'For k = 0 To 4
   '   LabRadial(k).ToolTipText = " # segments "
   'Next k
   For k = 0 To 2
      LabLevel(k) = Str$(BushSize(k + 1))
   '   LabLevel(k).ToolTipText = " Bush size "
   Next k
   For k = 0 To 2
      If zAngP(k + 1) > 0 Then
         cmdSwapAngles(k).Caption = ">"
      Else
         cmdSwapAngles(k).Caption = "<"
      End If
   Next k
End Sub

Private Sub LabLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   aDone = False
   Do
      Select Case Button
      Case vbLeftButton
         If BushSize(Index + 1) < 3 Then
            BushSize(Index + 1) = BushSize(Index + 1) + 1
         End If
      Case vbRightButton
         If BushSize(Index + 1) > 1 Then
            BushSize(Index + 1) = BushSize(Index + 1) - 1
         End If
      End Select
      LabLevel(Index) = Str$(BushSize(Index + 1))
      Sleep 150
      DoEvents
   Loop Until aDone
End Sub

Private Sub LabLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   aDone = True
End Sub

Private Sub LabRadial_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Rtype As Long
   Rtype = Index
   aDone = False
   Do
      Select Case Button
      Case vbLeftButton
         Select Case Rtype
         Case RTeeth
            If RadialRep(Index + 1) <= 34 Then
               RadialRep(Index + 1) = RadialRep(Index + 1) + 2
            End If
         Case Else
            If RadialRep(Index + 1) < 36 Then
               RadialRep(Index + 1) = RadialRep(Index + 1) + 1
            End If
         End Select
      Case vbRightButton
         Select Case Rtype
         Case RTeeth
            If RadialRep(Index + 1) >= 4 Then
               RadialRep(Index + 1) = RadialRep(Index + 1) - 2
            End If
         Case Else
            If RadialRep(Index + 1) > 2 Then
               RadialRep(Index + 1) = RadialRep(Index + 1) - 1
            End If
         End Select
      End Select
      LabRadial(Index) = Str$(RadialRep(Index + 1))
      Sleep 150
      DoEvents
   Loop Until aDone
End Sub

Private Sub LabRadial_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   aDone = True
End Sub

Private Sub optBrushes_Click(Index As Integer)
   Form1.optTools(Brush).Picture = optBrushes(Index).Picture
   Form1.optTools(Brush).Value = True
   BrushType = Index
   ToolType = Brush
End Sub

Private Sub optSprays_Click(Index As Integer)
   Form1.optTools(Spray).Picture = optSprays(Index).Picture
   Form1.optTools(Spray).Value = True
   SprayType = Index
   ToolType = Spray
End Sub

Private Sub optLines_Click(Index As Integer)
   Form1.optTools(ALine).Picture = optLines(Index).Picture
   Form1.optTools(ALine).Value = True
   LineType = Index
   ToolType = ALine
End Sub

Private Sub optPolyLines_Click(Index As Integer)
   Form1.optTools(PolyLine).Picture = optPolyLines(Index).Picture
   Form1.optTools(PolyLine).Value = True
   PolyLineType = Index
   ToolType = PolyLine
End Sub

Private Sub optCurvyLines_Click(Index As Integer)
   Form1.optTools(CurvyLine).Picture = optCurvyLines(Index).Picture
   Form1.optTools(CurvyLine).Value = True
   CurvyLineType = Index
   ToolType = CurvyLine
End Sub

Private Sub optRectangles_Click(Index As Integer)
   Form1.optTools(Rectangle).Picture = optRectangles(Index).Picture
   Form1.optTools(Rectangle).Value = True
   RectangleType = Index
   ToolType = Rectangle
End Sub

Private Sub optCirllipses_Click(Index As Integer)
   Form1.optTools(Cirllipse).Picture = optCirllipses(Index).Picture
   Form1.optTools(Cirllipse).Value = True
   CirllipseType = Index
   ToolType = Cirllipse
End Sub

Private Sub optCones_Click(Index As Integer)
   Form1.optTools(Cone).Picture = optCones(Index).Picture
   Form1.optTools(Cone).Value = True
   ConeType = Index
   ToolType = Cone
End Sub

Private Sub optTubes_Click(Index As Integer)
   Form1.optTools(Tube).Picture = optTubes(Index).Picture
   Form1.optTools(Tube).Value = True
   TubeType = Index
   ToolType = Tube
End Sub

Private Sub optBullets_Click(Index As Integer)
   Form1.optTools(Bullet).Picture = optBullets(Index).Picture
   Form1.optTools(Bullet).Value = True
   BulletType = Index
   ToolType = Bullet
End Sub

Private Sub optJunctions_Click(Index As Integer)
   Form1.optTools(Junction).Picture = optJunctions(Index).Picture
   Form1.optTools(Junction).Value = True
   JunctionType = Index
   ToolType = Junction
End Sub

Private Sub optArcs_Click(Index As Integer)
   Form1.optTools(Arc).Picture = optArcs(Index).Picture
   Form1.optTools(Arc).Value = True
   ArcType = Index
   ToolType = Arc
End Sub

Private Sub optShapes_Click(Index As Integer)
   Form1.optTools(Shape).Picture = optShapes(Index).Picture
   Form1.optTools(Shape).Value = True
   ShapeType = Index
   ToolType = Shape
End Sub

Private Sub optRadials_Click(Index As Integer)
   Form1.optTools(Radial).Picture = optRadials(Index).Picture
   Form1.optTools(Radial).Value = True
   RadialType = Index
   ToolType = Radial
End Sub

Private Sub optFills_Click(Index As Integer)
   Form1.optTools(AFill).Picture = optFills(Index).Picture
   Form1.optTools(AFill).Value = True
   FillType = Index
   FillbPattern Index
   ToolType = AFill
End Sub

Private Sub optTrees_Click(Index As Integer)
   Form1.optTools(Tree).Picture = optTrees(Index).Picture
   Form1.optTools(Tree).Value = True
   TreeType = Index
   ToolType = Tree
End Sub

Private Sub optArrows_Click(Index As Integer)
   Form1.optTools(Arrow).Picture = optArrows(Index).Picture
   Form1.optTools(Arrow).Value = True
   ArrowType = Index
   ToolType = Arrow
End Sub

Private Sub FillbPattern(Index As Integer)
' Public bPattern () as Byte 16x16
Dim k As Long
Dim kX As Long, kY As Long
   ReDim bPattern(16, 16)
   Select Case Index
   Case Fill1  ' Simple fill
      FillMemory bPattern(1, 1), 256, 1
   Case Fill2  ' Dense x
      ReDim bDummy(2, 2)
      bDummy(1, 1) = 1
      bDummy(2, 2) = 1
      TilebPattern 2, 2
   Case Fill3  ' Tight H-lines
      ReDim bDummy(1, 2)
      bDummy(1, 1) = 1
      TilebPattern 1, 2
   Case Fill4  ' Wide H-lines
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(k, 1) = 1
      Next k
      TilebPattern 8, 8
   Case Fill5  ' Tight V-lines
      ReDim bDummy(2, 1)
      bDummy(1, 1) = 1
      TilebPattern 2, 1
   Case Fill6  ' Wide V-lines
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(1, k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill7  ' Squares
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(k, 1) = 1
         bDummy(1, k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill8  ' Open X
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(k, k) = 1
         bDummy(9 - k, k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill9  ' Slant lines //
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(k, k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill10 ' Slant lines \\
      ReDim bDummy(8, 8)
      For k = 1 To 8
         bDummy(k, 9 - k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill11 ' Random
      ReDim bPattern(64, 64)
      For kY = 1 To 64
      For kX = 1 To 64
         If (Rnd - 0.25) < 0 Then bPattern(kX, kY) = 1
      Next kX
      Next kY
   Case Fill12 ' Small brick
      ReDim bDummy(8, 8)
      For k = 1 To 4
         bDummy(5, k) = 1
         bDummy(1, k + 4) = 1
      Next k
      For k = 1 To 8
         bDummy(k, 4) = 1
         bDummy(k, 8) = 1
      Next k
      TilebPattern 8, 8
   Case Fill13 ' Large brick
      ReDim bPattern(16, 16)
      For k = 1 To 8
         bPattern(9, k) = 1
         bPattern(1, k + 8) = 1
      Next k
      For k = 1 To 16
         bPattern(k, 8) = 1
         bPattern(k, 16) = 1
      Next k
   Case Fill14 ' Slant bricks  /-/
      ReDim bDummy(8, 8)
      For k = 5 To 8
         bDummy(k, 9 - k) = 1
      Next k
      For k = 1 To 8
         bDummy(k, k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill15 ' Slant bricks  \-\
      ReDim bDummy(8, 8)
      For k = 1 To 4
         bDummy(k, k) = 1
      Next k
      For k = 1 To 8
         bDummy(k, 9 - k) = 1
      Next k
      TilebPattern 8, 8
   Case Fill16 ' Large H-wave
      ReDim bDummy(16, 8)
      bDummy(1, 1) = 1
      bDummy(17 - 1, 1) = 1
      bDummy(2, 1) = 1
      bDummy(17 - 2, 1) = 1
      bDummy(3, 2) = 1
      bDummy(17 - 3, 2) = 1
      bDummy(4, 3) = 1
      bDummy(17 - 4, 3) = 1
      bDummy(4, 4) = 1
      bDummy(17 - 4, 4) = 1
      bDummy(5, 5) = 1
      bDummy(17 - 5, 5) = 1
      bDummy(5, 6) = 1
      bDummy(17 - 5, 6) = 1
      bDummy(6, 7) = 1
      bDummy(17 - 6, 7) = 1
      bDummy(7, 8) = 1
      bDummy(17 - 7, 8) = 1
      bDummy(8, 8) = 1
      bDummy(17 - 8, 8) = 1
      TilebPattern 16, 8
   Case Fill17 ' Large V-wave
      ReDim bDummy(8, 16)
      bDummy(8, 1) = 1
      bDummy(8, 17 - 1) = 1
      bDummy(8, 2) = 1
      bDummy(8, 17 - 2) = 1
      bDummy(7, 3) = 1
      bDummy(7, 17 - 3) = 1
      bDummy(6, 4) = 1
      bDummy(6, 17 - 4) = 1
      bDummy(5, 4) = 1
      bDummy(5, 17 - 4) = 1
      bDummy(4, 5) = 1
      bDummy(4, 17 - 5) = 1
      bDummy(3, 5) = 1
      bDummy(3, 17 - 5) = 1
      bDummy(2, 6) = 1
      bDummy(2, 17 - 6) = 1
      bDummy(1, 7) = 1
      bDummy(1, 17 - 7) = 1
      bDummy(1, 8) = 1
      bDummy(1, 17 - 8) = 1
      TilebPattern 8, 16
   Case Fill18 ' Small H-wave
      ReDim bDummy(16, 8)
      bDummy(1, 5) = 1
      bDummy(17 - 1, 5) = 1
      bDummy(2, 5) = 1
      bDummy(17 - 2, 5) = 1
      bDummy(3, 5) = 1
      bDummy(17 - 3, 5) = 1
      bDummy(4, 6) = 1
      bDummy(17 - 4, 6) = 1
      bDummy(5, 7) = 1
      bDummy(17 - 5, 7) = 1
      bDummy(6, 8) = 1
      bDummy(17 - 6, 8) = 1
      bDummy(7, 8) = 1
      bDummy(17 - 7, 8) = 1
      bDummy(8, 8) = 1
      bDummy(17 - 8, 8) = 1
      TilebPattern 16, 8
   Case Fill19 ' Small V-wave
      ReDim bDummy(8, 16)
      bDummy(4, 1) = 1
      bDummy(4, 17 - 1) = 1
      bDummy(4, 2) = 1
      bDummy(4, 17 - 2) = 1
      bDummy(4, 3) = 1
      bDummy(4, 17 - 3) = 1
      bDummy(3, 4) = 1
      bDummy(3, 17 - 4) = 1
      bDummy(2, 5) = 1
      bDummy(2, 17 - 5) = 1
      bDummy(1, 6) = 1
      bDummy(1, 17 - 6) = 1
      bDummy(1, 7) = 1
      bDummy(1, 17 - 7) = 1
      bDummy(1, 8) = 1
      bDummy(1, 17 - 8) = 1
      TilebPattern 8, 16
   Case Fill20
      ReDim bDummy(8, 8)
      bDummy(1, 1) = 1
      bDummy(1, 2) = 1
      bDummy(1, 4) = 1
      bDummy(1, 6) = 1
      bDummy(1, 7) = 1
      bDummy(1, 8) = 1
      bDummy(2, 4) = 1
      bDummy(3, 4) = 1
      bDummy(3, 8) = 1
      bDummy(4, 8) = 1
      bDummy(5, 2) = 1
      bDummy(5, 3) = 1
      bDummy(5, 4) = 1
      bDummy(5, 5) = 1
      bDummy(5, 6) = 1
      bDummy(5, 8) = 1
      bDummy(6, 8) = 1
      bDummy(7, 4) = 1
      bDummy(7, 8) = 1
      bDummy(8, 4) = 1
      TilebPattern 8, 8
   Case Else
   End Select
End Sub

Private Sub TilebPattern(kX As Long, kY As Long)
' bPattern(16,16)
' kx x ky pattern element
Dim ix As Long, iy As Long
Dim iix As Long, iiy As Long
   For iy = 1 To 16 Step kY
   For ix = 1 To 16 Step kX
      For iiy = 1 To kY
      For iix = 1 To kX
         bPattern(ix + iix - 1, iy + iiy - 1) = bDummy(iix, iiy)
      Next iix
      Next iiy
   Next ix
   Next iy
End Sub

