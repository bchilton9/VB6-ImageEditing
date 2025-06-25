VERSION 5.00
Begin VB.UserControl ctlLabelTextBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox txtTextBoxMultiLine 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "ctlLabelTextBox.ctx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTextBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Text            =   "TextBox"
      Top             =   600
      Width           =   1380
   End
   Begin VB.Shape shpBorder 
      Height          =   615
      Left            =   840
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   390
   End
End
Attribute VB_Name = "ctlLabelTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Border = False
Const m_def_CellPadding = 4
Const m_def_Layout = 0
'Property Variables:
Dim m_Border As Boolean
Dim m_CellPadding As Long
Dim m_Layout As Long
'Event declaration
Public Event Change()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblLabel.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLabel.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtTextBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtTextBox.Text() = New_Text
    txtTextBoxMultiLine.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtTextBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtTextBox.Locked() = New_Locked
    txtTextBoxMultiLine.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = txtTextBox.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtTextBox.PasswordChar() = New_PasswordChar
    txtTextBoxMultiLine.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,BackColor
Public Property Get BackColorTextBox() As OLE_COLOR
Attribute BackColorTextBox.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColorTextBox = txtTextBox.BackColor
End Property

Public Property Let BackColorTextBox(ByVal New_BackColorTextBox As OLE_COLOR)
    txtTextBox.BackColor() = New_BackColorTextBox
    txtTextBoxMultiLine.BackColor() = New_BackColorTextBox
    PropertyChanged "BackColorTextBox"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    
    shpBorder.Visible = New_Border
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = shpBorder.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    shpBorder.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,4
Public Property Get CellPadding() As Long
    CellPadding = m_CellPadding
End Property

Public Property Let CellPadding(ByVal New_CellPadding As Long)
    m_CellPadding = New_CellPadding
    PropertyChanged "CellPadding"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,Font
Public Property Get FontCaption() As Font
Attribute FontCaption.VB_Description = "Returns a Font object."
    Set FontCaption = lblLabel.Font
End Property

Public Property Set FontCaption(ByVal New_FontCaption As Font)
    Set lblLabel.Font = New_FontCaption
    PropertyChanged "FontCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,Font
Public Property Get FontText() As Font
Attribute FontText.VB_Description = "Returns a Font object."
    Set FontText = txtTextBox.Font
End Property

Public Property Set FontText(ByVal New_FontText As Font)
    Set txtTextBox.Font = New_FontText
    Set txtTextBoxMultiLine.Font = New_FontText
    PropertyChanged "FontText"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Layout() As Long
    Layout = m_Layout
End Property

Public Property Let Layout(ByVal New_Layout As Long)
    m_Layout = New_Layout
    PropertyChanged "Layout"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,ForeColor
Public Property Get ColorCaption() As OLE_COLOR
Attribute ColorCaption.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ColorCaption = lblLabel.ForeColor
End Property

Public Property Let ColorCaption(ByVal New_ColorCaption As OLE_COLOR)
    lblLabel.ForeColor() = New_ColorCaption
    PropertyChanged "ColorCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTextBox,txtTextBox,-1,ForeColor
Public Property Get ColorText() As OLE_COLOR
Attribute ColorText.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ColorText = txtTextBox.ForeColor
End Property

Public Property Let ColorText(ByVal New_ColorText As OLE_COLOR)
    txtTextBox.ForeColor() = New_ColorText
    txtTextBoxMultiLine.ForeColor() = New_ColorText
    PropertyChanged "ColorText"
End Property

Private Sub txtTextBox_Change()
    RaiseEvent Change
End Sub

Private Sub txtTextBoxMultiLine_Change()
    txtTextBox_Change
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Border = m_def_Border
    m_CellPadding = m_def_CellPadding
    m_Layout = m_def_Layout
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblLabel.Caption = PropBag.ReadProperty("Caption", "Label")
    txtTextBox.Text = PropBag.ReadProperty("Text", "TextBox")
    txtTextBox.Locked = PropBag.ReadProperty("Locked", False)
    txtTextBox.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    txtTextBox.BackColor = PropBag.ReadProperty("BackColorTextBox", &H80000005)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    shpBorder.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    m_CellPadding = PropBag.ReadProperty("CellPadding", m_def_CellPadding)
    Set lblLabel.Font = PropBag.ReadProperty("FontCaption", Ambient.Font)
    Set txtTextBox.Font = PropBag.ReadProperty("FontText", Ambient.Font)
    m_Layout = PropBag.ReadProperty("Layout", m_def_Layout)
    lblLabel.ForeColor = PropBag.ReadProperty("ColorCaption", &H80000012)
    txtTextBox.ForeColor = PropBag.ReadProperty("ColorText", &H80000008)
    txtTextBox.Width = PropBag.ReadProperty("WidthTextBox", 100)
    txtTextBox.Alignment = PropBag.ReadProperty("AlignmentText", txtTextBox.Alignment)
    txtTextBox.BorderStyle = PropBag.ReadProperty("BorderStyleTextBox", vbFixedSingle)
    lblLabel.BorderStyle = PropBag.ReadProperty("BorderStyleLabel", vbBSNone)
'    txtTextBoxMultiLine.ScrollBars = PropBag.ReadProperty("ScrollBars", vbSBNone)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    If m_Layout = 0 Then 'Horizontal
        lblLabel.Move shpBorder.BorderWidth + m_CellPadding, _
                      (UserControl.ScaleHeight - lblLabel.Height) / 2, _
                      UserControl.ScaleWidth - (shpBorder.BorderWidth * 2) - (m_CellPadding * 3) - txtTextBox.Width, _
                      lblLabel.Height
        
        txtTextBox.Move shpBorder.BorderWidth + (m_CellPadding * 2) + lblLabel.Width, _
                        shpBorder.BorderWidth + m_CellPadding, _
                        txtTextBox.Width, _
                        UserControl.ScaleHeight - (shpBorder.BorderWidth * 2) - (m_CellPadding * 2)
                        
        shpBorder.Move 0, _
                       0, _
                       UserControl.ScaleWidth, _
                       UserControl.ScaleHeight
    Else 'Vertical
        lblLabel.AutoSize = True
        
        lblLabel.Move (UserControl.ScaleWidth - lblLabel.Width) / 2, _
                      shpBorder.BorderWidth + m_CellPadding, _
                      lblLabel.Width, _
                      lblLabel.Height
        
        txtTextBox.Move (UserControl.ScaleWidth - txtTextBox.Width) / 2, _
                        shpBorder.BorderWidth + (m_CellPadding * 2) + lblLabel.Height, _
                        txtTextBox.Width, _
                        UserControl.ScaleHeight - (shpBorder.BorderWidth * 2) - (m_CellPadding * 3) - lblLabel.Height
    End If
                    
    shpBorder.Move 0, _
                   0, _
                   UserControl.ScaleWidth, _
                   UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    shpBorder.Visible = m_Border
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", lblLabel.Caption, "Label")
    Call PropBag.WriteProperty("Text", txtTextBox.Text, "TextBox")
    Call PropBag.WriteProperty("Locked", txtTextBox.Locked, False)
    Call PropBag.WriteProperty("PasswordChar", txtTextBox.PasswordChar, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColorTextBox", txtTextBox.BackColor, &H80000005)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderWidth", shpBorder.BorderWidth, 1)
    Call PropBag.WriteProperty("CellPadding", m_CellPadding, m_def_CellPadding)
    Call PropBag.WriteProperty("FontCaption", lblLabel.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontText", txtTextBox.Font, Ambient.Font)
    Call PropBag.WriteProperty("Layout", m_Layout, m_def_Layout)
    Call PropBag.WriteProperty("ColorCaption", lblLabel.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ColorText", txtTextBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("WidthTextBox", txtTextBox.Width, 100)
    Call PropBag.WriteProperty("AlignmentText", txtTextBox.Alignment, vbAlignLeft)
    Call PropBag.WriteProperty("BorderStyleTextBox", txtTextBox.BorderStyle, vbFixedSingle)
    Call PropBag.WriteProperty("BorderStyleLabel", lblLabel.BorderStyle, vbBSNone)
'    Call PropBag.WriteProperty("ScrollBars", txtTextBoxMultiLine.ScrollBars, vbSBNone)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,20
Public Property Get WidthTextBox() As Long
    WidthTextBox = txtTextBox.Width
End Property

Public Property Let WidthTextBox(ByVal New_WidthTextBox As Long)
    txtTextBox.Width() = New_WidthTextBox
    PropertyChanged "WidthTextBox"
    
    UserControl_Resize
End Property

Public Property Get AlignmentText() As AlignmentConstants
    AlignmentText = txtTextBox.Alignment
End Property

Public Property Let AlignmentText(New_AlignmentText As AlignmentConstants)
    txtTextBox.Alignment() = New_AlignmentText
    txtTextBoxMultiLine.Alignment() = New_AlignmentText
    PropertyChanged "AlignmentText"
End Property

Public Property Get BorderStyleTextBox() As Long
    BorderStyleTextBox = txtTextBox.BorderStyle
End Property

Public Property Let BorderStyleTextBox(New_BorderStyleTextBox As Long)
    txtTextBox.BorderStyle() = New_BorderStyleTextBox
    txtTextBoxMultiLine.BorderStyle() = New_BorderStyleTextBox
    PropertyChanged "BorderStyleTextBox"
End Property

Public Property Get BorderStyleLabel() As Long
    BorderStyleLabel = lblLabel.BorderStyle
End Property

Public Property Let BorderStyleLabel(New_BorderStyleLabel As Long)
    lblLabel.BorderStyle() = New_BorderStyleLabel
    PropertyChanged "BorderStyleLabel"
End Property
'
'Public Property Get ScrollBars() As ScrollBarConstants
'    ScrollBars = txtTextBoxMultiLine.ScrollBars
'End Property
'
'Public Property Let ScrollBars(New_ScrollBars As ScrollBarConstants)
'    txtTextBoxMultiLine.ScrollBars = New_ScrollBars
'    PropertyChanged "ScrollBars"
'End Property
