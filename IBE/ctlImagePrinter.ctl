VERSION 5.00
Begin VB.UserControl ctlImagePrinter 
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   Begin prjImageBrowser.ctlLabelTextBox txtPictureWidth 
      Height          =   405
      Left            =   2985
      TabIndex        =   15
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "Picture width"
      Text            =   "6"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorCaption    =   12582912
      WidthTextBox    =   40
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtPictureHeight 
      Height          =   405
      Left            =   960
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "Picture height"
      Text            =   "6"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorCaption    =   12582912
      WidthTextBox    =   40
      AlignmentText   =   2
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   705
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   4815
   End
   Begin VB.OptionButton optPaperOrientation 
      Appearance      =   0  'Flat
      Caption         =   "Portrait"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   6510
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optPaperOrientation 
      Appearance      =   0  'Flat
      Caption         =   "Landscape"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   6510
      Width           =   1215
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   1320
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   1980
      Width           =   3015
      Begin VB.Shape shpImage 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         Height          =   735
         Left            =   720
         Top             =   960
         Width           =   735
      End
      Begin VB.Shape shpMargin 
         BorderStyle     =   3  'Dot
         Height          =   1935
         Left            =   480
         Top             =   480
         Width           =   1935
      End
      Begin VB.Shape shpPaper 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2655
         Left            =   120
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape shpPaperShadow 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2655
         Left            =   240
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkPrintColor 
      Appearance      =   0  'Flat
      Caption         =   "Print in color"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   6900
      Width           =   1335
   End
   Begin prjImageBrowser.ctlLabelTextBox txtPrintCopy 
      Height          =   405
      Left            =   4305
      TabIndex        =   2
      Top             =   6420
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      Caption         =   "Copy(s)"
      Text            =   "1"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorCaption    =   12582912
      WidthTextBox    =   30
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtPaperHeight 
      Height          =   405
      Left            =   2985
      TabIndex        =   3
      Top             =   660
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "Paper height"
      Text            =   "11"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorCaption    =   12582912
      WidthTextBox    =   40
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtPaperWidth 
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   660
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "Paper width"
      Text            =   "8.5"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorCaption    =   12582912
      WidthTextBox    =   40
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtMarginRight 
      Height          =   645
      Left            =   4425
      TabIndex        =   5
      Top             =   3165
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1138
      Caption         =   "Right margin"
      Text            =   "1"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   1
      ColorCaption    =   12582912
      WidthTextBox    =   60
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtMarginBottom 
      Height          =   645
      Left            =   2280
      TabIndex        =   6
      Top             =   5085
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1138
      Caption         =   "Bottom margin"
      Text            =   "1"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   1
      ColorCaption    =   12582912
      WidthTextBox    =   60
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtMarginLeft 
      Height          =   645
      Left            =   120
      TabIndex        =   7
      Top             =   3165
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1138
      Caption         =   "Left margin"
      Text            =   "1.5"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   1
      ColorCaption    =   12582912
      WidthTextBox    =   60
      AlignmentText   =   2
   End
   Begin prjImageBrowser.ctlLabelTextBox txtMarginTop 
      Height          =   645
      Left            =   2280
      TabIndex        =   8
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1138
      Caption         =   "Top margin"
      Text            =   "1"
      BackColor       =   16761024
      Border          =   -1  'True
      BorderColor     =   16711680
      CellPadding     =   2
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   1
      ColorCaption    =   12582912
      WidthTextBox    =   60
      AlignmentText   =   2
   End
   Begin VB.Label lblPrinter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   180
      Width           =   450
   End
   Begin VB.Label lblPaperOrientation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orientation"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   6510
      Width           =   765
   End
End
Attribute VB_Name = "ctlImagePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:

'Property Variables:
Dim m_Picture As Picture

Public Sub RefreshPrinterList()
    Dim PRN As Printer
    For Each PRN In Printers
        cboPrinter.AddItem PRN.DeviceName
    Next
    
    cboPrinter.ListIndex = 0
End Sub

Private Sub optPaperOrientation_Click(Index As Integer)
    DrawPaper
End Sub

Private Sub txtMarginBottom_Change()
    DrawPaper
End Sub

Private Sub txtMarginLeft_Change()
    DrawPaper
End Sub

Private Sub txtMarginRight_Change()
    DrawPaper
End Sub

Private Sub txtMarginTop_Change()
    DrawPaper
End Sub

Private Sub txtPaperHeight_Change()
    DrawPaper
End Sub

Private Sub txtPaperWidth_Change()
    DrawPaper
End Sub

Private Sub txtPictureHeight_Change()
    DrawPaper
End Sub

Private Sub txtPictureWidth_Change()
    DrawPaper
End Sub

Private Sub UserControl_Show()
    picCanvas.BackColor = UserControl.BackColor
    chkPrintColor.BackColor = UserControl.BackColor
    optPaperOrientation(0).BackColor = UserControl.BackColor
    optPaperOrientation(1).BackColor = UserControl.BackColor
    
    RefreshPrinterList
    DrawPaper
End Sub

Private Sub SetPrinterAttributes()
    DrawPaper
    
    Dim PRN As Printer
    'Select the printer to use
    For Each PRN In Printers
        If PRN.DeviceName = cboPrinter.Text Then
            Set Printer = PRN
            Exit For
        End If
    Next
    
    With Printer
        .Copies = CLng(txtPrintCopy.Text) 'Print copy(s)
        'Set printing in color or monochrome
        If chkPrintColor.Value = vbChecked Then .ColorMode = vbPRCMColor Else .ColorMode = vbPRCMMonochrome
        
        .Width = Val(txtPaperWidth.Text) * TwipsPerInch 'Paper width
        .Height = Val(txtPaperHeight.Text) * TwipsPerInch 'Paperheight
        
'        Paper Orientation, Portrait Or Landscape
        If optPaperOrientation(0).Value = True Then .Orientation = vbPRORPortrait Else .Orientation = vbPRORLandscape
    End With
End Sub


Private Sub DrawPaper()
    On Error Resume Next
    
    Dim pWidth As Double, pHeight As Double
    If optPaperOrientation(0).Value = True Then 'Portrait
        pWidth = Val(txtPaperWidth.Text)
        pHeight = Val(txtPaperHeight.Text)
    Else 'Landscape
        pWidth = Val(txtPaperHeight.Text)
        pHeight = Val(txtPaperWidth.Text)
    End If
    
    Dim CanvasPadding As Long, PaperPadding As Long, ShadowPadding As Long
    CanvasPadding = 0
    ShadowPadding = 3
    
    Dim sWidth As Double, sHeight As Double
    sWidth = picCanvas.ScaleWidth - (CanvasPadding * 2) - ShadowPadding
    sHeight = pHeight * sWidth / pWidth
    If sHeight > picCanvas.ScaleHeight - (CanvasPadding * 2) - ShadowPadding Then
        sHeight = picCanvas.ScaleHeight - (CanvasPadding * 2) - ShadowPadding
        sWidth = pWidth * sHeight / pHeight
    End If
    
    shpPaper.Move (picCanvas.ScaleWidth - sWidth) / 2, (picCanvas.ScaleHeight - sHeight) / 2, sWidth, sHeight
    shpPaperShadow.Move shpPaper.Left + ShadowPadding, shpPaper.Top + ShadowPadding, shpPaper.Width, shpPaper.Height

    Dim mTop As Double, mBottom As Double, mLeft As Double, mRight As Double

    mLeft = Val(txtMarginLeft.Text) * shpPaper.Width / pWidth
    mTop = Val(txtMarginTop.Text) * shpPaper.Height / pHeight
    mRight = Val(txtMarginRight.Text) * shpPaper.Width / pWidth
    mBottom = Val(txtMarginBottom.Text) * shpPaper.Height / pHeight

    shpMargin.Move shpPaper.Left + mLeft, shpPaper.Top + mTop, shpPaper.Width - mLeft - mRight, shpPaper.Height - mTop - mBottom
    
    Dim iWidth As Single, iHeight As Single
    iWidth = sWidth / pWidth * Val(txtPictureWidth.Text)
    iHeight = sHeight / pHeight * Val(txtPictureHeight.Text)
    
    shpImage.Move shpMargin.Left, shpMargin.Top, iWidth, iHeight
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub PrintImage()
    SetPrinterAttributes
    
    Printer.ScaleMode = vbInches
    
    Dim dX As Single, dY As Single, dWidth As Single, dHeight As Single
    dX = Val(txtMarginLeft.Text)
    dY = Val(txtMarginTop.Text)
    
    If optPaperOrientation(0).Value = True Then 'Portrait
        dWidth = Val(txtPaperWidth.Text) - Val(txtMarginLeft.Text) - Val(txtMarginRight.Text)
        dHeight = Val(txtPaperHeight.Text) - Val(txtMarginTop.Text) - Val(txtMarginBottom.Text)
    Else 'Landscape
        dWidth = Val(txtPaperHeight.Text) - Val(txtMarginLeft.Text) - Val(txtMarginRight.Text)
        dHeight = Val(txtPaperWidth.Text) - Val(txtMarginTop.Text) - Val(txtMarginBottom.Text)
    End If
    
    Printer.PaintPicture m_Picture, dX, dY, Val(txtPictureWidth.Text), Val(txtPictureHeight.Text)
    
    Printer.FontSize = 8
    Printer.ForeColor = vbBlack
    Printer.FontName = "Arial"
    Printer.PSet (Val(txtMarginLeft.Text), 0.1), vbWhite '(Val(txtMarginLeft.Text), Printer.Height - 0.25)
    Printer.Print "Printed by Image Browser 1.0"
    
    
    Printer.EndDoc
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtPaperWidth,txtPaperWidth,-1,Text
Public Property Get paperWidth() As String
Attribute paperWidth.VB_Description = "Returns/sets the text contained in the control."
    paperWidth = txtPaperWidth.Text
End Property

Public Property Let paperWidth(ByVal New_PaperWidth As String)
    txtPaperWidth.Text() = New_PaperWidth
    PropertyChanged "PaperWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtPaperHeight,txtPaperHeight,-1,Text
Public Property Get paperHeight() As String
Attribute paperHeight.VB_Description = "Returns/sets the text contained in the control."
    paperHeight = txtPaperHeight.Text
End Property

Public Property Let paperHeight(ByVal New_PaperHeight As String)
    txtPaperHeight.Text() = New_PaperHeight
    PropertyChanged "PaperHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMarginTop,txtMarginTop,-1,Text
Public Property Get MarginTop() As String
Attribute MarginTop.VB_Description = "Returns/sets the text contained in the control."
    MarginTop = txtMarginTop.Text
End Property

Public Property Let MarginTop(ByVal New_MarginTop As String)
    txtMarginTop.Text() = New_MarginTop
    PropertyChanged "MarginTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMarginBottom,txtMarginBottom,-1,Text
Public Property Get MerginBottom() As String
Attribute MerginBottom.VB_Description = "Returns/sets the text contained in the control."
    MerginBottom = txtMarginBottom.Text
End Property

Public Property Let MerginBottom(ByVal New_MerginBottom As String)
    txtMarginBottom.Text() = New_MerginBottom
    PropertyChanged "MerginBottom"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMarginLeft,txtMarginLeft,-1,Text
Public Property Get MarginLeft() As String
Attribute MarginLeft.VB_Description = "Returns/sets the text contained in the control."
    MarginLeft = txtMarginLeft.Text
End Property

Public Property Let MarginLeft(ByVal New_MarginLeft As String)
    txtMarginLeft.Text() = New_MarginLeft
    PropertyChanged "MarginLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMarginRight,txtMarginRight,-1,Text
Public Property Get MarginRight() As String
Attribute MarginRight.VB_Description = "Returns/sets the text contained in the control."
    MarginRight = txtMarginRight.Text
End Property

Public Property Let MarginRight(ByVal New_MarginRight As String)
    txtMarginRight.Text() = New_MarginRight
    PropertyChanged "MarginRight"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=optPaperOrientation(0),optPaperOrientation,0,Value
'Public Property Get Orientation() As Boolean
'    Orientation = optPaperOrientation(0).Value
'End Property
'
'Public Property Let Orientation(ByVal New_Orientation As Boolean)
'    optPaperOrientation(0).Value() = New_Orientation
'    PropertyChanged "Orientation"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtPrintCopy,txtPrintCopy,-1,Text
Public Property Get Copy() As String
Attribute Copy.VB_Description = "Returns/sets the text contained in the control."
    Copy = txtPrintCopy.Text
End Property

Public Property Let Copy(ByVal New_Copy As String)
    txtPrintCopy.Text() = New_Copy
    PropertyChanged "Copy"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=chkPrintColor,chkPrintColor,-1,Value
Public Property Get ColorMode() As Integer
Attribute ColorMode.VB_Description = "Returns/sets the value of an object."
    ColorMode = chkPrintColor.Value
End Property

Public Property Let ColorMode(ByVal New_ColorMode As Integer)
    chkPrintColor.Value() = New_ColorMode
    PropertyChanged "ColorMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,6
Public Property Get ImageHeight() As Single
    ImageHeight = Val(txtPictureHeight.Text)
End Property

Public Property Let ImageHeight(ByVal New_ImageHeight As Single)
    txtPictureHeight.Text() = New_ImageHeight
    PropertyChanged "ImageHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,6
Public Property Get ImageWidth() As Single
    ImageWidth = Val(txtPictureWidth.Text)
End Property

Public Property Let ImageWidth(ByVal New_ImageWidth As Single)
    txtPictureWidth.Text() = New_ImageWidth
    PropertyChanged "ImageWidth"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Picture = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtPaperWidth.Text = PropBag.ReadProperty("PaperWidth", "8.5")
    txtPaperHeight.Text = PropBag.ReadProperty("PaperHeight", "11")
    txtMarginTop.Text = PropBag.ReadProperty("MarginTop", "1")
    txtMarginBottom.Text = PropBag.ReadProperty("MerginBottom", "1")
    txtMarginLeft.Text = PropBag.ReadProperty("MarginLeft", "1.5")
    txtMarginRight.Text = PropBag.ReadProperty("MarginRight", "1")
    optPaperOrientation(0).Value = PropBag.ReadProperty("Orientation", True)
    txtPrintCopy.Text = PropBag.ReadProperty("Copy", "1")
    chkPrintColor.Value = PropBag.ReadProperty("ColorMode", 0)
    txtPictureHeight.Text = PropBag.ReadProperty("ImageHeight", Val(txtPictureHeight.Text))
    txtPictureWidth.Text = PropBag.ReadProperty("ImageWidth", Val(txtPictureWidth.Text))
    optPaperOrientation(0).Value = PropBag.ReadProperty("OrientationPortrait", True)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PaperWidth", txtPaperWidth.Text, "8.5")
    Call PropBag.WriteProperty("PaperHeight", txtPaperHeight.Text, "11")
    Call PropBag.WriteProperty("MarginTop", txtMarginTop.Text, "1")
    Call PropBag.WriteProperty("MerginBottom", txtMarginBottom.Text, "1")
    Call PropBag.WriteProperty("MarginLeft", txtMarginLeft.Text, "1.5")
    Call PropBag.WriteProperty("MarginRight", txtMarginRight.Text, "1")
    Call PropBag.WriteProperty("Orientation", optPaperOrientation(0).Value, True)
    Call PropBag.WriteProperty("Copy", txtPrintCopy.Text, "1")
    Call PropBag.WriteProperty("ColorMode", chkPrintColor.Value, 0)
    Call PropBag.WriteProperty("ImageHeight", Val(txtPictureHeight.Text), 6)
    Call PropBag.WriteProperty("ImageWidth", Val(txtPictureWidth.Text), 6)
    Call PropBag.WriteProperty("OrientationPortrait", optPaperOrientation(0).Value, True)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get OrientationPortrait() As Boolean
Attribute OrientationPortrait.VB_Description = "Returns/sets the value of an object."
    OrientationPortrait = optPaperOrientation(0).Value
End Property

Public Property Let OrientationPortrait(ByVal New_OrientationPortrait As Boolean)
    optPaperOrientation(0).Value = New_OrientationPortrait
    optPaperOrientation(1).Value = Not New_OrientationPortrait
    PropertyChanged "OrientationPortrait"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Sub FitBest(Optional KeepOrientation As Boolean = False)
    Dim vTemp As Single
    
    Dim pWidth As Single, pHeight As Single
    pWidth = m_Picture.Width
    pHeight = m_Picture.Height
    
    Dim aWidth As Single, aHeight As Single
    aWidth = Val(txtPaperWidth.Text) - Val(txtMarginLeft.Text) - Val(txtMarginRight.Text)
    aHeight = Val(txtPaperHeight.Text) - Val(txtMarginTop.Text) - Val(txtMarginBottom.Text)
    
    If KeepOrientation Then
    
    Else
        If pWidth > pHeight And aHeight > aWidth Then
            optPaperOrientation(0).Value = False
            optPaperOrientation(1).Value = True
            
            aWidth = Val(txtPaperHeight.Text) - Val(txtMarginLeft.Text) - Val(txtMarginRight.Text)
            aHeight = Val(txtPaperWidth.Text) - Val(txtMarginTop.Text) - Val(txtMarginBottom.Text)
        Else
            optPaperOrientation(0).Value = True
            optPaperOrientation(1).Value = False
        End If
    End If
    
    Dim sWidth As Single, sHeight As Single
    sWidth = aWidth
    sHeight = pHeight * aWidth / pWidth
    
    If sHeight > aHeight Then
        sHeight = aHeight
        sWidth = pWidth * aHeight / pHeight
    End If
    
    txtPictureWidth.Text = Round(sWidth, 2)
    txtPictureHeight.Text = Round(sHeight, 2)
    
    DrawPaper
End Sub

Public Sub FitCenter()
    Dim hMargin As Single, vMargin As Single
    Dim picWidth As Single, picHeight As Single
    Dim paperWidth As Single, paperHeight As Single
    
    picWidth = Val(txtPictureWidth.Text)
    picHeight = Val(txtPictureHeight.Text)
    
    If optPaperOrientation(0).Value = True Then 'Portrait
        paperWidth = Val(txtPaperWidth.Text)
        paperHeight = Val(txtPaperHeight.Text)
    Else 'Landscape
        paperWidth = Val(txtPaperHeight.Text)
        paperHeight = Val(txtPaperWidth.Text)
    End If
    
    vMargin = (paperWidth - picWidth) / 2
    hMargin = (paperHeight - picHeight) / 2
    
    txtMarginLeft.Text = Round(vMargin, 2)
    txtMarginRight.Text = Round(vMargin, 2)
    txtMarginTop.Text = Round(hMargin, 2)
    txtMarginBottom.Text = Round(hMargin, 2)
    
    DrawPaper
End Sub

Public Sub FitReset()
    txtPictureWidth.Text = m_Picture.Width / 1440
    txtPictureHeight.Text = m_Picture.Height / 1440
    
    DrawPaper
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    
    picCanvas.BackColor = UserControl.BackColor
    chkPrintColor.BackColor = UserControl.BackColor
    optPaperOrientation(0).BackColor = UserControl.BackColor
    optPaperOrientation(1).BackColor = UserControl.BackColor
End Property
