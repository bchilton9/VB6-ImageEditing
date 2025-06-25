VERSION 5.00
Begin VB.Form frmTransform 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Transformers  -  Preview"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAC 
      Caption         =   "Accept"
      Height          =   315
      Index           =   3
      Left            =   4680
      TabIndex        =   97
      Top             =   6825
      Width           =   660
   End
   Begin VB.CheckBox chkUseSelectcn 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Use selected color on Deformers"
      Height          =   465
      Left            =   3255
      TabIndex        =   88
      Top             =   5340
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.PictureBox picLCul 
      AutoRedraw      =   -1  'True
      Height          =   270
      Left            =   5055
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   63
      Top             =   4980
      Width           =   1050
   End
   Begin VB.PictureBox picTPal 
      AutoRedraw      =   -1  'True
      Height          =   1020
      Left            =   3225
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   62
      Top             =   3870
      Width           =   2940
   End
   Begin VB.Frame fraT 
      BackColor       =   &H0080C0FF&
      Caption         =   "   Filters   "
      Height          =   4890
      Index           =   1
      Left            =   45
      TabIndex        =   41
      Top             =   0
      Width           =   1440
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PIXELIZE"
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
         Index           =   20
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   3150
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FOG"
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
         Index           =   19
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1890
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INVERT"
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
         Index           =   18
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2310
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SOLARIZE"
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
         Index           =   17
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   4620
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BLACK / WHITE"
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
         Index           =   16
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIFFUSE VERT"
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
         Index           =   15
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIFFUSE HORZ"
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
         Index           =   14
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1050
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIFFUSE"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   840
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTRAST"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   630
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LITHO"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2520
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHARPEN"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   4200
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SMOOTH"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4410
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RELIEF"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3570
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "POSTERIZE"
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
         Index           =   3
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3360
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ENGRAVE EMBOSS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1485
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GREY DITHER"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2100
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTOUR"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   420
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHADE VERT"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3990
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHADE HORZ"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3780
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MELT"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2730
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilters 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OIL"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2940
         Width           =   1185
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   28
         Left            =   1275
         TabIndex        =   105
         Top             =   4605
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   54
         Top             =   1560
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   2
         Left            =   1275
         TabIndex        =   53
         Top             =   3375
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   6
         Left            =   1275
         TabIndex        =   52
         Top             =   3570
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdAC 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   33
      Top             =   6825
      Width           =   675
   End
   Begin VB.Frame fraT 
      BackColor       =   &H0080C0FF&
      Caption         =   "   Adders   "
      Height          =   1755
      Index           =   2
      Left            =   45
      TabIndex        =   9
      Top             =   4920
      Width           =   3105
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THICK LINE H && V"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   855
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THICK  LINE V"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   645
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIAG NET"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   645
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SPOKES"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BORDER"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THICK  LINE H"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   435
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ELLIPSES"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   855
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CIRCLES"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   435
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WAVES H && V"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1485
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WAVES VERT"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1275
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WAVES HORZ"
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
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1065
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LINES H && V"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1485
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LINES VERT"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1275
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdders 
         BackColor       =   &H00FFFFFF&
         Caption         =   " LINES HORZ"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   27
         Left            =   2955
         TabIndex        =   102
         Top             =   1485
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   26
         Left            =   2955
         TabIndex        =   101
         Top             =   660
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   25
         Left            =   2955
         TabIndex        =   92
         Top             =   1275
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   24
         Left            =   2955
         TabIndex        =   90
         Top             =   1065
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   13
         Left            =   2955
         TabIndex        =   61
         Top             =   870
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   12
         Left            =   2955
         TabIndex        =   59
         Top             =   450
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   11
         Left            =   2955
         TabIndex        =   57
         Top             =   255
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   10
         Left            =   1440
         TabIndex        =   55
         Top             =   1500
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   9
         Left            =   1440
         TabIndex        =   39
         Top             =   1290
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   8
         Left            =   1440
         TabIndex        =   37
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   7
         Left            =   1440
         TabIndex        =   36
         Top             =   870
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   5
         Left            =   1440
         TabIndex        =   26
         Top             =   660
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   4
         Left            =   1440
         TabIndex        =   25
         Top             =   450
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   24
         Top             =   255
         Width           =   105
      End
   End
   Begin VB.Frame fraT 
      BackColor       =   &H0080C0FF&
      Caption         =   "  Deformers "
      Height          =   4890
      Index           =   0
      Left            =   1485
      TabIndex        =   8
      Top             =   0
      Width           =   1650
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WINFLUTE H && V"
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
         Index           =   21
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   4620
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WINFLUTE HORZ"
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
         Index           =   20
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4200
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BUBBLY"
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
         Index           =   17
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   210
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TUNNEL"
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
         Index           =   19
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3990
         Width           =   1395
      End
      Begin VB.CheckBox chkLens 
         BackColor       =   &H00C0E0FF&
         Caption         =   "o"
         Height          =   195
         Left            =   60
         TabIndex        =   77
         ToolTipText     =   " Circle area "
         Top             =   1050
         Width           =   405
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ROTATE"
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
         Index           =   18
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2940
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIN-MAG"
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
         Index           =   16
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1260
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "STARS"
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
         Index           =   15
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3360
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SWIRL"
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
         Index           =   14
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3570
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WINFLUTE VERT"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4410
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LENS"
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
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1050
         Width           =   840
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIRROR  LENS"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2310
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIRROR  BOTTOM"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2100
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIRROR  TOP"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1890
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIRROR  RIGHT"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIRROR  LEFT"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1470
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TILE"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3780
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ROUND RECT"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3150
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RIPPLE VERT"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2730
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RIPPLE HORZ"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FLUTE VERT"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   630
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FLUTE HORZ"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeformers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ELLIPSE"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   23
         Left            =   1485
         TabIndex        =   76
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   22
         Left            =   1485
         TabIndex        =   74
         Top             =   3570
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   21
         Left            =   1485
         TabIndex        =   72
         Top             =   3360
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   20
         Left            =   1485
         TabIndex        =   71
         Top             =   3165
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   19
         Left            =   1485
         TabIndex        =   70
         Top             =   2955
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   18
         Left            =   1485
         TabIndex        =   69
         Top             =   2520
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   17
         Left            =   1485
         TabIndex        =   68
         Top             =   1275
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   16
         Left            =   1485
         TabIndex        =   67
         Top             =   1065
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   15
         Left            =   1485
         TabIndex        =   66
         Top             =   855
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   14
         Left            =   1485
         TabIndex        =   65
         Top             =   645
         Width           =   105
      End
      Begin VB.Label LabLCul 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   180
         Index           =   0
         Left            =   1485
         TabIndex        =   64
         Top             =   420
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset"
      Height          =   315
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6825
      Width           =   615
   End
   Begin VB.CommandButton cmdAC 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   915
      TabIndex        =   5
      Top             =   6825
      Width           =   675
   End
   Begin VB.CommandButton cmdAC 
      Caption         =   "Accept"
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   6825
      Width           =   660
   End
   Begin VB.PictureBox picSlider 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   330
      Left            =   3180
      MousePointer    =   9  'Size W E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   3
      Top             =   3165
      Width           =   3060
   End
   Begin VB.PictureBox picFC 
      Height          =   2985
      Left            =   3195
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      Begin VB.PictureBox picF 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2880
         Left            =   30
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   1
         Top             =   15
         Width           =   2880
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   195
      Left            =   4920
      Top             =   5970
      Width           =   225
   End
   Begin VB.Label LabONOFF 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   5280
      TabIndex        =   87
      Top             =   5970
      Width           =   405
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rectangular selection"
      Height          =   270
      Left            =   3240
      TabIndex        =   86
      Top             =   5955
      Width           =   1605
   End
   Begin VB.Label LabT 
      BackColor       =   &H00C0E0FF&
      Caption         =   "LabT"
      Height          =   270
      Left            =   3225
      TabIndex        =   75
      Top             =   3555
      Width           =   1995
   End
   Begin VB.Label LabParam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5460
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selected color = 255"
      Height          =   285
      Left            =   3270
      TabIndex        =   2
      Top             =   4980
      Width           =   1620
   End
End
Attribute VB_Name = "frmTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmTransform.frm

' Uses Filter.bas

Option Explicit
Option Base 1
' Saved Publics
Dim bsvArray() As Byte
Dim svRpicW As Long, svRpicH As Long

Dim svSSX As Long
Dim svSSY As Long
Dim svSSW As Long
Dim svSSH As Long

'Public SSX As Long ' Shape left
'Public SSY As Long ' Shape top
'Public SSW As Long ' Shape width
'Public SSH As Long ' Shape Height
'Public aSelRect As Boolean
'Public WWLO As Long, HHLO As Long
'Public WWHI As Long, HHHI As Long

Dim picFW As Long, picFH As Long
Dim iWMax As Long, iHMax As Long
Dim zAspect As Single
Dim xm As Single, ym As Single
Dim ixx As Long, iyy As Long
Dim zPARAMVAL As Single
Dim i As Long
Dim j As Long
Dim bHolder() As Byte
Dim bS As BITMAPINFO



' ____192_____
'| _________  |
'||picF     | |
'||         | 192
'||_________| |
'|picFC       |
'|____________|


Private Sub Form_Load()

   chkLens.Value = -aLensCheck
   aUseSelectcn = -chkUseSelectcn.Value
   
   TransformType = 0
   
   frmTransform.Left = frmTransformLeft
   frmTransform.Top = frmTransformTop

   ShowTPalette
   
   ' Save public vars
   svcanvasW = canvasW
   svcanvasH = canvasH
   svRpicW = RpicW
   svRpicH = RpicH
   ReDim bsvArray(canvasW, canvasH)
   bsvArray() = bArray()
   
   Selectcn = SelLeftCulNum
   ShowLCul
   
   LabT = ""
   
   ' Display image (scaled down if necessary)
   iWMax = 192
   iHMax = 192
   zAspect = canvasW / canvasH
   If canvasW <= iWMax Then
      picFW = canvasW
      picFH = CLng(picFW / zAspect)
   Else  ' canvasW > iWMax
      picFW = iWMax
      picFH = CLng(iWMax / zAspect)
   End If
   If picFH > iHMax Then
      picFH = iHMax
      picFW = CLng(picFH * zAspect)
   End If
   ' Round down to multiple of 4
   picFW = picFW And &HFFFFFFFC
   If picFW < 4 Then picFW = 4
   picFH = picFH And &HFFFFFFFC
   If picFH < 4 Then picFH = 4
   
   picF.Width = picFW
   picF.Height = picFH
   
   ' Scale picture to fit
   ym = (picFH - 1) / (canvasH - 1)
   xm = (picFW - 1) / (canvasW - 1)
   ReDim bHolder(1 To picFW, 1 To picFH)
   For iy = 1 To canvasH
      iyy = ym * (iy - 1) + 1
      If iyy > 0 Then
      If iyy <= picFH Then
         For ix = 1 To canvasW
            ixx = xm * (ix - 1) + 1
            If ixx > 0 Then
            If ixx <= picFW Then
               bHolder(ixx, iyy) = bArray(ix, iy)
            End If
            End If
         Next ix
      End If
      End If
   Next iy
   ReDim bArray(1 To picFW, 1 To picFH)
   bArray() = bHolder()
   canvasW = UBound(bArray(), 1)
   canvasH = UBound(bArray(), 2)
   
'Public SSX As Long ' Shape left
'Public SSY As Long ' Shape top
'Public SSW As Long ' Shape width
'Public SSH As Long ' Shape Height
'Public aSelRect As Boolean
'Public WWLO As Long, HHLO As Long
'Public WWHI As Long, HHHI As Long
   
   If Not aSelRect Then
      WWLO = 1: HHLO = 1
      WWHI = canvasW: HHHI = canvasH
      LabONOFF = "OFF"
   Else
      WWLO = SSX * (picFW / svcanvasW)
      HHHI = picFH - SSY * (picFH / svcanvasH)
      WWHI = WWLO + SSW * (picFW / svcanvasW) - 1
      HHLO = HHHI - SSH * (picFH / svcanvasH) + 1
      LabONOFF = "ON"
   End If
   
   ' Set up palette
   CopyMemory bS.Colors(0), CulBGR(0), 1024
   'picF.Picture = LoadPicture
   picF.Width = canvasW     ' Always multiple of 4 !!
   picF.Height = canvasH

   With bS.bmi
      .biSize = 40
      .biwidth = canvasW
      .biheight = canvasH
      .biPlanes = 1
      .biBitCount = 8
      .biSizeImage = canvasW * canvasH
   End With
   DoEvents
   DISPLAYF
   
   
   ' Default Public parameters
' Public in Filter.bas

   PCONTOUR = 128
   PDITHER = 16
   PENGRAVEMBOSS = 1
   PPOSTERIZE = 128
   PRELIEF = 1 '40
   PSMOOTH = 1
   PSHADE = 32
   PMELT = 2
   POIL = 2
   PSHARPEN = 1
   PLITHO = 96
   PCONTRAST = 10
   PDIFFUSE = 8
   PBLACKWHITE = 128
   PSOLAR = 128
   PFOG = 0
   PSQUARE = 2
   
   zPELLIPSE = 0.5
   PFLUTE = 8
   PRIPPLE = 1
   zPROUNDRECT = 0.25
   PTILE = 2
   PMLENS = 2
   zPLENS = 1
   PFWINDOW = 8
   zPSWIRL = 6
   PKALI = 1
   zPMINMAG = 1
   zPROTATE = 0
   
   PLINES = 4
   zPTHICKLINE = 0.01
   PSPOKES = 12
   PDNET = 12
   
   For i = 0 To picSlider.Width 'Step 2
     j = RGB(i, i, i)
     If i = 128 Then j = RGB(0, 250, 250)
     picSlider.Line (i, 0)-(i, picSlider.Height), j
     If i Mod 8 = 0 Then
        picSlider.Line (i, 0)-(i, picSlider.Height), RGB(255 - i, 255 - i, 255 - i)
     End If
   Next i
   picSlider.Refresh
   picSlider.Visible = False
   LabParam.Visible = False
End Sub

Private Sub cmdAC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Recover Public variables
   canvasW = svcanvasW
   canvasH = svcanvasH
   RpicW = svRpicW
   RpicH = svRpicH
   ReDim bArray(canvasW, canvasH)
   bArray() = bsvArray()
   
   Select Case Index
   Case 0, 3  ' Accept
      Select Case TransformType
      Case TContour: PCONTOUR = zPARAMVAL
      Case TDither: PDITHER = zPARAMVAL
      Case TEngraveEmboss: PENGRAVEMBOSS = zPARAMVAL
      Case TPosterize: PPOSTERIZE = zPARAMVAL
      Case TRelief: PRELIEF = zPARAMVAL
      Case TSmooth: PSMOOTH = zPARAMVAL
      Case TShadeV: PSHADE = zPARAMVAL
      Case TShadeH: PSHADE = zPARAMVAL
      Case TMelt: PMELT = zPARAMVAL
      Case TOil: POIL = zPARAMVAL
      Case TSharpen: PSHARPEN = zPARAMVAL
      Case TLitho: PLITHO = zPARAMVAL
      Case TContrast: PCONTRAST = zPARAMVAL
      Case TDiffuse, THDiffuse, TVDiffuse: PDIFFUSE = zPARAMVAL
      Case TBlackWhite: PBLACKWHITE = zPARAMVAL
      Case TSolar: PSOLAR = zPARAMVAL
      Case TInvert
      Case TFog: PFOG = zPARAMVAL
      Case TSquare: PSQUARE = zPARAMVAL
      
      Case TEllipse: zPELLIPSE = zPARAMVAL
      Case TFluteH: PFLUTE = zPARAMVAL
      Case TFluteV: PFLUTE = zPARAMVAL
      Case TRippleH: PRIPPLE = zPARAMVAL
      Case TRippleV: PRIPPLE = zPARAMVAL
      Case TRoundRect: zPROUNDRECT = zPARAMVAL
      Case TTile: PTILE = zPARAMVAL
      Case TMirrorL
      Case TMirrorR
      Case TMirrorT
      Case TMirrorB
      Case TMlens: PMLENS = zPARAMVAL
      Case TLens: zPLENS = zPARAMVAL
      Case TFWindowHorz: PFWINDOW = zPARAMVAL
      Case TFWindowVert: PFWINDOW = zPARAMVAL
      Case TFWindowHV: PFWINDOW = zPARAMVAL
      Case TSwirl: zPSWIRL = zPARAMVAL
      Case TSpokess: PKALI = zPARAMVAL
      Case TMinMag: zPMINMAG = zPARAMVAL
      Case TBubbly: zPLENS = zPARAMVAL
      Case TRotate: zPROTATE = zPARAMVAL
      Case TTunnel: PTILE = zPARAMVAL
      
      Case THLines: PLINES = zPARAMVAL
      Case TVLines: PLINES = zPARAMVAL
      Case THVLines: PLINES = zPARAMVAL
      Case THWaves: PLINES = zPARAMVAL
      Case TVWaves: PLINES = zPARAMVAL
      Case THVWaves: PLINES = zPARAMVAL
      Case TCircles: zPLENS = zPARAMVAL
      Case TEllipses: zPLENS = zPARAMVAL
      Case TThickLineH: zPTHICKLINE = zPARAMVAL
      Case TThickLineV: zPTHICKLINE = zPARAMVAL
      Case TThickLineHV: zPTHICKLINE = zPARAMVAL
      Case TBorder: zPTHICKLINE = zPARAMVAL
      Case TSpokes: PSPOKES = zPARAMVAL
      Case TDNet: PDNET = zPARAMVAL
      End Select
         
      'SelLeftCulNum = Selectcn
   
   Case 1, 2  ' Cancel
      TransformType = TNone
   End Select
   
   frmTransformLeft = frmTransform.Left
   frmTransformTop = frmTransform.Top

   Unload frmTransform
End Sub

Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'picSlider_MouseMove Button, Shift, x, y

   If Button = vbLeftButton Then
      Select Case TransformType
      Case TContour:       cmdFilters_Click TContour - TContour
         zPARAMVAL = PCONTOUR
      Case TDither:        cmdFilters_Click TDither - TContour
         zPARAMVAL = PDITHER
      Case TEngraveEmboss: cmdFilters_Click TEngraveEmboss - TContour
         zPARAMVAL = PENGRAVEMBOSS
      Case TPosterize:     cmdFilters_Click TPosterize - TContour
         zPARAMVAL = PPOSTERIZE
      Case TRelief:        cmdFilters_Click TRelief - TContour
         zPARAMVAL = PRELIEF
      Case TSmooth:        cmdFilters_Click TSmooth - TContour
         zPARAMVAL = PSMOOTH
      Case TShadeV:        cmdFilters_Click TShadeV - TContour
         zPARAMVAL = PSHADE
      Case TShadeH:        cmdFilters_Click TShadeH - TContour
         zPARAMVAL = PSHADE
      Case TMelt:          cmdFilters_Click TMelt - TContour
         zPARAMVAL = PMELT
      Case TOil:           cmdFilters_Click TOil - TContour
         zPARAMVAL = POIL
      Case TSharpen:       cmdFilters_Click TSharpen - TContour
         zPARAMVAL = PSHARPEN
      Case TLitho:         cmdFilters_Click TLitho - TContour
         zPARAMVAL = PLITHO
      Case TContrast:      cmdFilters_Click TContrast - TContour
         zPARAMVAL = PCONTRAST
      Case TDiffuse:       cmdFilters_Click TDiffuse - TContour
         zPARAMVAL = PDIFFUSE
      Case THDiffuse:      cmdFilters_Click THDiffuse - TContour
         zPARAMVAL = PDIFFUSE
      Case TVDiffuse:      cmdFilters_Click TVDiffuse - TContour
         zPARAMVAL = PDIFFUSE
      Case TBlackWhite:    cmdFilters_Click TBlackWhite - TContour
         zPARAMVAL = PBLACKWHITE
      Case TSolar:         cmdFilters_Click TSolar - TContour
         zPARAMVAL = PSOLAR
      Case TInvert:        cmdFilters_Click TInvert - TContour
      Case TFog:           cmdFilters_Click TFog - TContour
         zPARAMVAL = PFOG
      Case TSquare:        cmdFilters_Click TSquare - TContour
         zPARAMVAL = PSQUARE
      
      '-------------------------------------------------------------
      Case TEllipse:       cmdDeformers_Click TEllipse - TEllipse
         zPARAMVAL = zPELLIPSE
      Case TFluteH:        cmdDeformers_Click TFluteH - TEllipse
         zPARAMVAL = PFLUTE
      Case TFluteV:        cmdDeformers_Click TFluteV - TEllipse
         zPARAMVAL = PFLUTE
      Case TRippleH:       cmdDeformers_Click TRippleH - TEllipse
         zPARAMVAL = PRIPPLE
      Case TRippleV:       cmdDeformers_Click TRippleV - TEllipse
         zPARAMVAL = PRIPPLE
      Case TRoundRect:     cmdDeformers_Click TRoundRect - TEllipse
         zPARAMVAL = zPROUNDRECT
      Case TTile:          cmdDeformers_Click TTile - TEllipse
         zPARAMVAL = PTILE
      Case TMirrorL:       cmdDeformers_Click TMirrorL - TEllipse
      Case TMirrorR:       cmdDeformers_Click TMirrorR - TEllipse
      Case TMirrorT:       cmdDeformers_Click TMirrorT - TEllipse
      Case TMirrorL:       cmdDeformers_Click TMirrorB - TEllipse
      Case TMlens:         cmdDeformers_Click TMlens - TEllipse
         zPARAMVAL = PMLENS
      Case TLens:          cmdDeformers_Click TLens - TEllipse
         zPARAMVAL = zPLENS
      
      Case TFWindowHorz:    cmdDeformers_Click TFWindowHorz - TEllipse
         zPARAMVAL = PFWINDOW
      Case TFWindowVert:     cmdDeformers_Click TFWindowVert - TEllipse
         zPARAMVAL = PFWINDOW
      Case TFWindowHV:      cmdDeformers_Click TFWindowHV - TEllipse
         zPARAMVAL = PFWINDOW
      
      Case TSwirl:         cmdDeformers_Click TSwirl - TEllipse
         zPARAMVAL = zPSWIRL
      Case TSpokess:         cmdDeformers_Click TSpokess - TEllipse
         zPARAMVAL = PKALI
      Case TMinMag:        cmdDeformers_Click TMinMag - TEllipse
         zPARAMVAL = zPMINMAG
      Case TBubbly:        cmdDeformers_Click TBubbly - TEllipse
         zPARAMVAL = zPLENS
      Case TRotate:        cmdDeformers_Click TRotate - TEllipse
         zPARAMVAL = zPROTATE
      Case TTunnel:        cmdDeformers_Click TTunnel - TEllipse
         zPARAMVAL = PTILE
      '-------------------------------------------------------------
      
      Case THLines:        cmdAdders_Click THLines - THLines
         zPARAMVAL = PLINES
      Case TVLines:        cmdAdders_Click TVLines - THLines
         zPARAMVAL = PLINES
      Case THVLines:       cmdAdders_Click THVLines - THLines
         zPARAMVAL = PLINES
      Case THWaves:        cmdAdders_Click THWaves - THLines
         zPARAMVAL = PLINES
      Case TVWaves:        cmdAdders_Click TVWaves - THLines
         zPARAMVAL = PLINES
      Case THVWaves:       cmdAdders_Click THVWaves - THLines
         zPARAMVAL = PLINES
      Case TCircles:       cmdAdders_Click TCircles - THLines
         zPARAMVAL = zPLENS
      Case TEllipses:      cmdAdders_Click TEllipses - THLines
         zPARAMVAL = zPLENS
      Case TThickLineH:     cmdAdders_Click TThickLineH - THLines
         zPARAMVAL = zPTHICKLINE
      Case TThickLineV:     cmdAdders_Click TThickLineV - THLines
         zPARAMVAL = zPTHICKLINE
      Case TThickLineHV:     cmdAdders_Click TThickLineHV - THLines
         zPARAMVAL = zPTHICKLINE
      Case TBorder:        cmdAdders_Click TBorder - THLines
         zPARAMVAL = zPTHICKLINE
      Case TSpokes:        cmdAdders_Click TSpokes - THLines
         zPARAMVAL = PSPOKES
      Case TDNet:          cmdAdders_Click TDNet - THLines
         zPARAMVAL = PDNET
      End Select
   End If

End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xp As Single
' picSlider.Width=200
   xp = x
   If xp < 0 Then xp = 0
   If xp > 200 Then xp = 200
   
   Select Case TransformType
   Case TContour  ' Threshold 32 -> 224
      PCONTOUR = xp + 33
      If PCONTOUR > 224 Then PCONTOUR = 224
      LabParam = Str$(PCONTOUR): LabParam.Refresh
   Case TDither   ' 16 ->  48 'Floyd-Steinberg B
      PDITHER = (32 / 200) * xp + 16
      LabParam = Str$(PDITHER): LabParam.Refresh
   Case TEngraveEmboss  ' -3 -> +3
      PENGRAVEMBOSS = (xp - 100) \ 32
      If PENGRAVEMBOSS = 0 Then
         If xp <= 100 Then PENGRAVEMBOSS = -1 Else PENGRAVEMBOSS = 1
      End If
      LabParam = Str$(PENGRAVEMBOSS): LabParam.Refresh
   Case TPosterize      ' Threshold 32 -> 224
      PPOSTERIZE = xp + 32
      If PPOSTERIZE > 224 Then PPOSTERIZE = 224
      LabParam = Str$(PPOSTERIZE): LabParam.Refresh
   Case TRelief   '1,2,3
      PRELIEF = xp / 60
      If PRELIEF = 0 Then PRELIEF = 1
      LabParam = Str$(PRELIEF): LabParam.Refresh
   Case TSmooth
      PSMOOTH = xp / 50   ' 1-4
      If PSMOOTH = 0 Then PSMOOTH = 1
      LabParam = Str$(PSMOOTH): LabParam.Refresh
   Case TShadeV, TShadeH
      PSHADE = (xp - 100)     ' -100 -> +100
      LabParam = Str$(PSHADE): LabParam.Refresh
   Case TMelt
      PMELT = 3 * xp / 100 + 2 ' 2 -> 8
      LabParam = Str$(PMELT): LabParam.Refresh
   Case TOil   ' 1,2,3,4,5
      POIL = xp \ 40
      If POIL = 0 Then POIL = 1
      LabParam = Str$(POIL): LabParam.Refresh
   Case TSharpen   ' 1,2,3
      PSHARPEN = xp / 60
      If PSHARPEN = 0 Then PSHARPEN = 1
      LabParam = Str$(PSHARPEN): LabParam.Refresh
   Case TLitho   ' 92-100
      PLITHO = 94 + xp / 25
      LabParam = Str$(PLITHO): LabParam.Refresh
   Case TContrast   ' -66 - +66
      PCONTRAST = (2 * xp - 200) \ 3
      LabParam = Str$(PCONTRAST): LabParam.Refresh
   Case TDiffuse, THDiffuse, TVDiffuse
      PDIFFUSE = xp / 13 + 1     ' 1  ->  16
      LabParam = Str$(PDIFFUSE): LabParam.Refresh
   Case TBlackWhite   ' 0 - +255
      PBLACKWHITE = xp * 1.275
      LabParam = Str$(PBLACKWHITE): LabParam.Refresh
   Case TSolar   ' 32 - +224
      PSOLAR = 32 + xp * 1.12
      LabParam = Str$(PSOLAR): LabParam.Refresh
   Case TInvert
   Case TFog   ' -100 to 100
      PFOG = (xp - 100)
      LabParam = Str$(PFOG): LabParam.Refresh
   Case TSquare   '1 to 50
      PSQUARE = 1 + xp \ 4
      LabParam = Str$(PSQUARE): LabParam.Refresh
   
   '------------------------------------------------------
   
   Case TEllipse
      zPELLIPSE = xp / 100      ' 0.005 to 2
      If zPELLIPSE = 0 Then zPELLIPSE = 0.005
      If zPELLIPSE = 1 Then zPELLIPSE = 1.005
      LabParam = Str$(zPELLIPSE): LabParam.Refresh
   Case TFluteH, TFluteV
      PFLUTE = (xp - 100) / 2 '4 ' -50 -> +50 (exc -2 to +2)
      ' -2 to 3  'Non-linear
      Select Case PFLUTE
      Case -2, -1: PFLUTE = -3
      Case 0, 1, 2: PFLUTE = 3
      End Select
      LabParam = Str$(PFLUTE): LabParam.Refresh
   Case TRippleH, TRippleV
      PRIPPLE = 1 + (40 * xp) / 200 ' 0 -> 40
      LabParam = Str$(PRIPPLE): LabParam.Refresh
   Case TRoundRect
      zPROUNDRECT = xp / 400 ' 0 -> .5
      LabParam = Str$(zPROUNDRECT): LabParam.Refresh
   Case TTile
      PTILE = xp / 5 + 1  ' 1 -> 41
      LabParam = Str$(PTILE): LabParam.Refresh
   Case TMirrorL
   Case TMirrorR
   Case TMirrorT
   Case TMirrorB
   Case TMlens       ' 2 -> 40
      PMLENS = 2 + (40 * xp) / 200
      LabParam = Str$(PMLENS): LabParam.Refresh
   Case TLens     ' .1 -> 4
      zPLENS = xp / 50 + 0.1
      LabParam = Str$(Round(zPLENS, 2)): LabParam.Refresh
   Case TFWindowHorz, TFWindowVert, TFWindowHV    ' 4->100
      PFWINDOW = 4 + xp \ 2
      LabParam = Str$(PFWINDOW): LabParam.Refresh
   Case TSwirl       ' -50 -> 50
      zPSWIRL = (xp - 100) / 2
      If zPSWIRL = 0 Then zPSWIRL = 0.1
      LabParam = Str$(zPSWIRL): LabParam.Refresh
   Case TSpokess ' 1 to 5
      PKALI = xp / 10 + 1
      LabParam = Str$(PKALI): LabParam.Refresh
   Case TMinMag      ' 1.01 - 2
      zPMINMAG = 0.01 + xp / 100
      If zPMINMAG = 0 Then zPMINMAG = 0.01
      LabParam = Str$(Round(zPMINMAG, 3)): LabParam.Refresh
   Case TBubbly     ' 1 -> 4
      zPLENS = xp / 50 + 0.5
      LabParam = Str$(zPLENS): LabParam.Refresh
   Case TRotate     ' -1890 to + 180
      zPROTATE = (xp - 100) * 180 / 100
      LabParam = Str$(zPROTATE): LabParam.Refresh
   Case TTunnel
      PTILE = xp / 5 + 1  ' 1 -> 41
      LabParam = Str$(PTILE): LabParam.Refresh
      
   '------------------------------------------------------
   Case THLines      ' 0 -> 40
      PLINES = (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case TVLines      ' 0 -> 40
      PLINES = (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case THVLines      ' 0 -> 40
      PLINES = (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case THWaves      ' 1 -> 40
      PLINES = 1 + (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case TVWaves      ' 1 -> 40
      PLINES = 1 + (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case THVWaves      ' 1 -> 40
      PLINES = 1 + (xp + 1) / 5
      LabParam = Str$(PLINES): LabParam.Refresh
   Case TCircles     ' 1 -> 4
      zPLENS = xp / 50 + 0.5
      LabParam = Str$(zPLENS): LabParam.Refresh
   Case TEllipses     ' 1 -> 4
      zPLENS = xp / 50 + 0.5
      LabParam = Str$(zPLENS): LabParam.Refresh
   Case TThickLineH     '.0 -> 1.0
      zPTHICKLINE = xp / 200
      LabParam = Str$(zPTHICKLINE): LabParam.Refresh
   Case TThickLineV     '.0 -> 1.0
      zPTHICKLINE = xp / 200
      LabParam = Str$(zPTHICKLINE): LabParam.Refresh
   Case TThickLineHV     '.0 -> 1.0
      zPTHICKLINE = xp / 200
      LabParam = Str$(zPTHICKLINE): LabParam.Refresh
   Case TBorder     '1 -> 41
      zPTHICKLINE = 1 + xp \ 5
      LabParam = Str$(zPTHICKLINE): LabParam.Refresh
   Case TSpokes     '1 -> 201
      PSPOKES = 1 + xp
      LabParam = Str$(PSPOKES): LabParam.Refresh
   Case TDNet     '1 -> 41
      PDNET = 1 + xp \ 5
      LabParam = Str$(PDNET): LabParam.Refresh
   
   End Select
   picSlider_MouseDown Button, Shift, x, y
   DoEvents
End Sub

Private Sub cmdFilters_Click(Index As Integer)
   bArray() = bHolder()
   Select Case Index
   Case 0:
      zPARAMVAL = PCONTOUR
      TransformType = TContour
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PCONTOUR)
      LabParam.Refresh
      Contour
   Case 1:
      zPARAMVAL = PDITHER
      TransformType = TDither
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PDITHER)
      LabParam.Refresh
      Dither
   Case 2:
      zPARAMVAL = PENGRAVEMBOSS
      TransformType = TEngraveEmboss
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PENGRAVEMBOSS)
      LabParam.Refresh
      Dim SC As Long
      SC = Selectcn
      EngraveEmboss
   Case 3:
      zPARAMVAL = PPOSTERIZE
      TransformType = TPosterize
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PPOSTERIZE)
      LabParam.Refresh
      Posterize
   Case 4:
      zPARAMVAL = PRELIEF
      TransformType = TRelief
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PRELIEF)
      LabParam.Refresh
      Relief
   Case 5:
      zPARAMVAL = PSMOOTH
      TransformType = TSmooth
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSMOOTH)
      LabParam.Refresh
      Smooth
   Case 6:
      zPARAMVAL = PSHADE
      TransformType = TShadeV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSHADE)
      LabParam.Refresh
      ShadeV
   Case 7:
      zPARAMVAL = PSHADE
      TransformType = TShadeH
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSHADE)
      LabParam.Refresh
      ShadeH
   Case 8:
      zPARAMVAL = PMELT
      TransformType = TMelt
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PMELT)
      LabParam.Refresh
      Melt
   Case 9:
      zPARAMVAL = POIL
      TransformType = TOil
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(POIL)
      LabParam.Refresh
      Oil
   Case 10:
      zPARAMVAL = PSHARPEN
      TransformType = TSharpen
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSHARPEN)
      LabParam.Refresh
      Sharpen
   Case 11:
      zPARAMVAL = PLITHO
      TransformType = TLitho
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLITHO)
      LabParam.Refresh
      Litho
   Case 12:
      zPARAMVAL = PCONTRAST
      TransformType = TContrast
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PCONTRAST)
      LabParam.Refresh
      Contrast
   Case 13:
      zPARAMVAL = PDIFFUSE
      TransformType = TDiffuse
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PDIFFUSE)
      LabParam.Refresh
      Diffuse
   Case 14:
      zPARAMVAL = PDIFFUSE
      TransformType = THDiffuse
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PDIFFUSE)
      LabParam.Refresh
      Diffuse
   Case 15:
      zPARAMVAL = PDIFFUSE
      TransformType = TVDiffuse
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PDIFFUSE)
      LabParam.Refresh
      Diffuse
   Case 16:
      zPARAMVAL = PBLACKWHITE
      TransformType = TBlackWhite
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PBLACKWHITE)
      LabParam.Refresh
      BlackWhite
   Case 17:
      zPARAMVAL = PSOLAR
      TransformType = TSolar
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSOLAR)
      LabParam.Refresh
      Solarize
   Case 18:
      TransformType = TInvert
      picSlider.Visible = False
      LabParam.Visible = False
      Invert
   Case 19:
      zPARAMVAL = PFOG
      TransformType = TFog
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFOG)
      LabParam.Refresh
      Fog
   Case 20:
      zPARAMVAL = PSQUARE
      TransformType = TSquare
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSQUARE)
      LabParam.Refresh
      Pixelize
   
   End Select
   LabT = cmdFilters(Index).Caption
   DISPLAYF
End Sub

Private Sub cmdDeformers_Click(Index As Integer)
   bArray() = bHolder()
   Select Case Index
   Case 0:
      zPARAMVAL = zPELLIPSE
      TransformType = TEllipse
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam.Refresh
      Elliptic
   Case 1:
      zPARAMVAL = PFLUTE
      TransformType = TFluteH
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFLUTE)
      LabParam.Refresh
      FluteH
   Case 2:
      zPARAMVAL = PFLUTE
      TransformType = TFluteV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFLUTE)
      LabParam.Refresh
      FluteV
   Case 3:
      zPARAMVAL = PRIPPLE
      TransformType = TRippleH
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PRIPPLE)
      LabParam.Refresh
      RippleH
   Case 4:
      zPARAMVAL = PRIPPLE
      TransformType = TRippleV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PRIPPLE)
      LabParam.Refresh
      RippleV
   Case 5:
      zPARAMVAL = zPROUNDRECT
      TransformType = TRoundRect
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPROUNDRECT)
      LabParam.Refresh
      RoundRect
   Case 6:
      zPARAMVAL = PTILE
      TransformType = TTile
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PTILE)
      LabParam.Refresh
      Tile
   Case 7:
      TransformType = TMirrorL
      picSlider.Visible = False
      LabParam.Visible = False
      LabParam.Refresh
      MirrorLeft
   Case 8:
      TransformType = TMirrorR
      picSlider.Visible = False
      LabParam.Visible = False
      LabParam.Refresh
      MirrorRight
   Case 9:
      TransformType = TMirrorT
      picSlider.Visible = False
      LabParam.Visible = False
      MirrorTop
   Case 10:
      TransformType = TMirrorB
      picSlider.Visible = False
      LabParam.Visible = False
      MirrorBottom
   Case 11
      zPARAMVAL = PMLENS
      TransformType = TMlens
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PMLENS)
      LabParam.Refresh
      MirrorLens
   Case 12
      zPARAMVAL = zPLENS
      TransformType = TLens
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPLENS)
      LabParam.Refresh
      ALens
   Case 13
      zPARAMVAL = PFWINDOW
      TransformType = TFWindowVert
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFWINDOW)
      LabParam.Refresh
      FlutedWindowVert
   Case 14
      zPARAMVAL = zPSWIRL
      TransformType = TSwirl
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPSWIRL)
      LabParam.Refresh
      Swirl
   Case 15
      zPARAMVAL = PKALI
      TransformType = TSpokess
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PKALI)
      LabParam.Refresh
      Stars
   Case 16
      zPARAMVAL = zPMINMAG
      TransformType = TMinMag
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPMINMAG)
      LabParam.Refresh
      MinMag
   Case 17
      zPARAMVAL = zPLENS
      TransformType = TBubbly
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPLENS)
      LabParam.Refresh
      Bubbly
   Case 18
      zPARAMVAL = zPROTATE
      TransformType = TRotate
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPROTATE)
      LabParam.Refresh
      Rotate
   Case 19:
      zPARAMVAL = PTILE
      TransformType = TTunnel
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PTILE)
      LabParam.Refresh
      Tunnel
   Case 20
      zPARAMVAL = PFWINDOW
      TransformType = TFWindowHorz
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFWINDOW)
      LabParam.Refresh
      FlutedWindowHorz
   Case 21
      zPARAMVAL = PFWINDOW
      TransformType = TFWindowHV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PFWINDOW)
      LabParam.Refresh
      FlutedWindowHV
   End Select
   
   LabT = cmdDeformers(Index).Caption
   
   DISPLAYF
End Sub

Private Sub chkLens_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aLensCheck = -chkLens.Value
End Sub

Private Sub chkUseSelectcn_Click()
   aUseSelectcn = -chkUseSelectcn.Value
End Sub

Private Sub cmdAdders_Click(Index As Integer)
   bArray() = bHolder()
   Select Case Index
   Case 0:
      zPARAMVAL = PLINES
      TransformType = THLines
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddHLines
   Case 1:
      zPARAMVAL = PLINES
      TransformType = TVLines
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddVLines
   Case 2:
      zPARAMVAL = PLINES
      TransformType = THVLines
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddHVLines
   Case 3:
      zPARAMVAL = PLINES
      TransformType = THWaves
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddHWaves
   Case 4:
      zPARAMVAL = PLINES
      TransformType = TVWaves
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddVWaves
   Case 5:
      zPARAMVAL = PLINES
      TransformType = THVWaves
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PLINES)
      LabParam.Refresh
      AddHVWaves
   Case 6:
      zPARAMVAL = zPLENS
      TransformType = TCircles
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPLENS)
      LabParam.Refresh
      AddCircles
   Case 7:
      zPARAMVAL = zPLENS
      TransformType = TEllipses
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPLENS)
      LabParam.Refresh
      AddEllipses
   Case 8:
      zPARAMVAL = zPTHICKLINE
      TransformType = TThickLineH
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPTHICKLINE)
      LabParam.Refresh
      AddThickLineH
   Case 9:
      zPARAMVAL = zPTHICKLINE
      TransformType = TBorder
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPTHICKLINE)
      LabParam.Refresh
      AddBorder
   Case 10:
      zPARAMVAL = PSPOKES
      TransformType = TSpokes
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PSPOKES)
      LabParam.Refresh
      AddSpokes
   Case 11:
      zPARAMVAL = PDNET
      TransformType = TDNet
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(PDNET)
      LabParam.Refresh
      AddDiagNet
   Case 12:
      zPARAMVAL = zPTHICKLINE
      TransformType = TThickLineV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPTHICKLINE)
      LabParam.Refresh
      AddThickLineV
   Case 13:
      zPARAMVAL = zPTHICKLINE
      TransformType = TThickLineHV
      picSlider.Visible = True
      LabParam.Visible = True
      LabParam = Str$(zPTHICKLINE)
      LabParam.Refresh
      AddThickLineHV
   End Select
   
   LabT = cmdAdders(Index).Caption
   
   DISPLAYF
End Sub

' Select color '''''''''''''''''''''''''''

Private Sub ShowTPalette()
' Public ix,iy
Dim k As Long
   For k = 0 To 255
    iy = (k Mod 8) * 8
    ix = (k \ 8) * 6
    picTPal.Line (ix, iy)-(ix + 4, iy + 6), CulRGB(k), BF
   Next k
   picTPal.Refresh
End Sub

Private Sub picTPal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Selectcn
   If x >= 0 And x <= (picTPal.Width) Then
   If y >= 0 And y <= (picTPal.Height) Then
      Selectcn = (x \ 6) * 8 + y \ 8
      If Selectcn < 0 Then Selectcn = 0
      If Selectcn > 255 Then Selectcn = 255
      If Button = vbLeftButton Then
         ShowLCul
         LabParam = Str$(Selectcn)
      End If
   End If
   End If
End Sub

Private Sub ShowLCul()
   picLCul.BackColor = CulRGB(Selectcn)
   picLCul.Refresh
   Label1 = "Selected color =" & Str$(Selectcn)
   For i = 0 To 28
      LabLCul(i).BackColor = CulRGB(Selectcn)
      LabLCul(i).Refresh
   Next i
End Sub
'''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdReset_Click()
   bArray() = bHolder()
   TransformType = TNone
   DISPLAYF
End Sub

Private Sub DISPLAYF()
   If SetDIBitsToDevice(picF.hDC, 0, 0, canvasW, canvasH, _
      0, 0, 0, canvasH, bArray(1, 1), bS, DIB_RGB_COLORS) = 0 Then
      MsgBox "DISPLAY ERROR IN TRANSFORM", vbCritical, "Display"
      End
   End If
   picF.Refresh
End Sub




