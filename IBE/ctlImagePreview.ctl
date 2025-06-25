VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlImagePreview 
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   LockControls    =   -1  'True
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   Begin MSComctlLib.Slider sldZoom 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      Min             =   10
      Max             =   5000
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.VScrollBar vscImage 
      Height          =   2775
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar hscImage 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
   End
   Begin VB.PictureBox picImageContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Shape shpBorder 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnuAutoVerb 
      Caption         =   "AutoVerb"
      Begin VB.Menu mnuAutoVerbZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuAutoVerbZoom50 
            Caption         =   "50%"
         End
         Begin VB.Menu mnuAutoVerbZoom100 
            Caption         =   "100%"
         End
         Begin VB.Menu mnuAutoVerbZoom200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuAutoVerbZoom400 
            Caption         =   "400%"
         End
      End
   End
End
Attribute VB_Name = "ctlImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_ShowZoomSlider = True

'Property Variables:
Dim m_ShowZoomSlider As Boolean
Dim m_Picture As StdPicture

'User variables
Dim v_RightClick As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picImage,picImage,-1,Picture
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    
    sldZoom_Scroll
    sldZoom_Change
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sldZoom,sldZoom,-1,Value
Public Property Get Zoom() As Long
Attribute Zoom.VB_Description = "Returns/sets the value of an object."
    Zoom = sldZoom.Value
End Property

Public Property Let Zoom(ByVal New_Zoom As Long)
    sldZoom.Value() = New_Zoom
    PropertyChanged "Zoom"
    
    SetZoom New_Zoom
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowZoomSlider() As Boolean
    ShowZoomSlider = m_ShowZoomSlider
End Property

Public Property Let ShowZoomSlider(ByVal New_ShowZoomSlider As Boolean)
    m_ShowZoomSlider = New_ShowZoomSlider
    PropertyChanged "ShowZoomSlider"
    
    sldZoom.Visible = New_ShowZoomSlider
End Property

Private Sub hscImage_Change()
    hscImage_Scroll
End Sub

Private Sub hscImage_Scroll()
    picImage.Left = 0 - hscImage.Value '- (picImage.Width / 2)
End Sub

Private Sub SetZoom(ZoomValue As Long)
    sldZoom.Value = ZoomValue
    sldZoom_Scroll
End Sub

Private Sub mnuAutoVerbZoom100_Click()
    SetZoom 100
End Sub

Private Sub mnuAutoVerbZoom200_Click()
    SetZoom 200
End Sub

Private Sub mnuAutoVerbZoom400_Click()
    SetZoom 400
End Sub

Private Sub mnuAutoVerbZoom50_Click()
    SetZoom 50
End Sub

Private Sub picImage_Click()
    picImageContainer_Click
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picImageContainer_MouseUp Button, Shift, picImage.Left + X, picImage.Top + Y
End Sub

Private Sub picImageContainer_Click()
    If v_RightClick Then PopupMenu mnuAutoVerb
End Sub

Private Sub picImageContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then v_RightClick = True Else v_RightClick = False
End Sub

Private Sub picImageContainer_Resize()
    On Error Resume Next
    picImage.Move (picImageContainer.ScaleWidth - picImage.Width) / 2, (picImageContainer.ScaleHeight - picImage.Height) / 2 ', picImageContainer.ScaleWidth, picImageContainer.ScaleHeight
    sldZoom_Change
End Sub

Private Sub sldZoom_Change()
    SetScrollBarValues

End Sub

Private Sub sldZoom_Scroll()
    If m_Picture Is Nothing Then Exit Sub
    
    Dim pWidth As Long, pHeight As Long
    pWidth = (m_Picture.Width / K_DotsPerPixel) * sldZoom.Value / 100
    pHeight = (m_Picture.Height / K_DotsPerPixel) * sldZoom.Value / 100
    
    picImage.Move (picImageContainer.ScaleWidth - pWidth) / 2, (picImageContainer.ScaleHeight - pHeight) / 2, pWidth, pHeight

    'picImage.Width = (m_Picture.Width / K_DotsPerPixel) * sldZoom.Value / 100
    'picImage.Height = (m_Picture.Height / K_DotsPerPixel) * sldZoom.Value / 100
    
    LoadPictureToShow m_Picture
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ShowZoomSlider = m_def_ShowZoomSlider
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    sldZoom.Value = PropBag.ReadProperty("Zoom", 100)
    m_ShowZoomSlider = PropBag.ReadProperty("ShowZoomSlider", m_def_ShowZoomSlider)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", &H80000008)
    shpBorder.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    picImageContainer.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    picImageContainer.Move shpBorder.BorderWidth, shpBorder.BorderWidth, UserControl.ScaleWidth - vscImage.Width - (shpBorder.BorderWidth * 2), UserControl.ScaleHeight - hscImage.Height - sldZoom.Height - (shpBorder.BorderWidth * 2)
    vscImage.Move picImageContainer.Left + picImageContainer.Width, shpBorder.BorderWidth, vscImage.Width, picImageContainer.Height
    hscImage.Move shpBorder.BorderWidth, picImageContainer.Top + picImageContainer.Height, picImageContainer.Width, hscImage.Height
    sldZoom.Move (UserControl.ScaleWidth - (UserControl.ScaleWidth * 60 / 100)) / 2, picImageContainer.Top + picImageContainer.Height + hscImage.Height, (UserControl.ScaleWidth * 60 / 100), sldZoom.Height
    
    shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    SetControlAttributesToPropertyVariables
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Zoom", sldZoom.Value, 100)
    Call PropBag.WriteProperty("ShowZoomSlider", m_ShowZoomSlider, m_def_ShowZoomSlider)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, &H80000008)
    Call PropBag.WriteProperty("BorderWidth", shpBorder.BorderWidth, 1)
    Call PropBag.WriteProperty("BackColor", picImageContainer.BackColor, vbButtonFace)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderWidth
Public Property Get BorderWidth() As Long
    BorderWidth = shpBorder.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Long)
    shpBorder.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

Private Sub SetControlAttributesToPropertyVariables()
    sldZoom.Visible = m_ShowZoomSlider
End Sub

Private Sub LoadPictureToShow(PIC As StdPicture)
    If PIC Is Nothing Then Exit Sub
    
    Dim posLeft As Long, posTop As Long, picWidth As Long, picHeight As Long
    
    picWidth = picImage.ScaleWidth
    picHeight = picWidth / PIC.Width * PIC.Height
    
    If picHeight > picImage.ScaleHeight Then
        picHeight = picImage.ScaleHeight
        picWidth = picHeight / PIC.Height * PIC.Width
    End If
    
    posLeft = (picImage.ScaleWidth - picWidth) / 2
    posTop = (picImage.ScaleHeight - picHeight) / 2
    
    CopyImage PIC, picImage, 0, 0, , , posLeft, posTop, picWidth, picHeight
End Sub

Private Sub SetScrollBarValues()
    vscImage.Max = picImage.Height - picImageContainer.ScaleHeight
    hscImage.Max = picImage.Width - picImageContainer.ScaleWidth
    
    vscImage.Value = (picImage.Height - picImageContainer.ScaleHeight) / 2
    hscImage.Value = (picImage.Width - picImageContainer.ScaleWidth) / 2
    
    If vscImage.Max < 1 Then vscImage.Enabled = False Else vscImage.Enabled = True
    If hscImage.Max < 1 Then hscImage.Enabled = False Else hscImage.Enabled = True
End Sub

Private Sub vscImage_Change()
    vscImage_Scroll
End Sub

Private Sub vscImage_Scroll()
    picImage.Top = 0 - vscImage.Value
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = picImageContainer.BackColor
End Property

Public Property Let BackColor(New_BackColor As OLE_COLOR)
    picImageContainer.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    UserControl.BackColor = New_BackColor
End Property


