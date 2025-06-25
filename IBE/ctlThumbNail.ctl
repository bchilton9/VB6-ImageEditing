VERSION 5.00
Begin VB.UserControl ctlThumbNail 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkChecked 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   195
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape shpFocus 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   1095
      Left            =   240
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnuAutoVerb 
      Caption         =   "AutoVerb"
      Begin VB.Menu mnuAutoVerbCopy 
         Caption         =   "Copy image to clipboard"
      End
      Begin VB.Menu mnuAutoVerbCopyThumbNail 
         Caption         =   "Copy thumbnail to clipboard"
      End
      Begin VB.Menu mnuAutoVerbView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuAutoVerbPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuAutoVerbProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuAutoVerbExternalViewer 
         Caption         =   "Open with external Viewer"
      End
      Begin VB.Menu mnuAutoVerbExternalEditor 
         Caption         =   "Open with external Editor"
      End
      Begin VB.Menu mnuAutoVerbExternalPrinter 
         Caption         =   "Open with external Printer"
      End
   End
End
Attribute VB_Name = "ctlThumbNail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_CellPadding = 4
Const m_def_AutoCheckUncheck = True
Const m_def_AutoVerbMenu = False
Const m_def_FileNameBox = True
Const m_def_CheckBox = True
Const m_def_ExternalViewer = ""
Const m_def_ExternalEditor = ""
Const m_def_ExternalPrinter = ""

'Property Variables:
Dim m_CellPadding As Long
Dim m_ActualFileName As String
Dim m_AutoCheckUncheck As Boolean
Dim m_AutoVerbMenu As Boolean
Dim m_FileNameBox As Boolean
Dim m_CheckBox As Boolean
Dim m_Picture As StdPicture
Dim m_ExternalViewer As String
Dim m_ExternalEditor As String
Dim m_ExternalPrinter As String

'Events
Event Click()
Event RightClick()
Event DblClick()

'Private variables
Private v_RightClick As Boolean

Private Sub mnuAutoVerbCopyThumbNail_Click()
    CopyThumbNailToClipBoard
End Sub

Public Sub ViewImage()
    Dim PreviewForm As Form
    Set PreviewForm = New frmImagePreview
    
    PreviewForm.ShowPreview m_Picture
End Sub

Private Sub mnuAutoVerbExternalEditor_Click()
    OpenWithExternalEditor
End Sub

Private Sub mnuAutoVerbExternalPrinter_Click()
    OpenWithExternalPrinter
End Sub

Private Sub mnuAutoVerbExternalViewer_Click()
    OpenWithExternalViewer
End Sub

Private Sub mnuAutoVerbPrint_Click()
    ShowPrint
End Sub

Public Sub ShowPrint()
    frmPrintImage.ShowPrint m_Picture
End Sub

Private Sub mnuAutoVerbView_Click()
    ViewImage
End Sub

Private Sub picPicture_Click()
    UserControl_Click
End Sub

Private Sub picPicture_DblClick()
    UserControl_DblClick
End Sub

Private Sub picPicture_GotFocus()
    UserControl_GotFocus
End Sub

Private Sub picPicture_LostFocus()
    UserControl_LostFocus
End Sub

Private Sub picPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, picPicture.Left + X, picPicture.Top + Y
End Sub

Public Sub CopyImageToClipBoard()
    Clipboard.Clear
    Clipboard.SetData m_Picture
End Sub

Public Sub CopyThumbNailToClipBoard()
    Clipboard.Clear
    Clipboard.SetData picPicture.Image
End Sub

Private Sub mnuEditCopy_Click()
    CopyImageToClipBoard
End Sub

Private Sub mnuAutoVerbCopy_Click()
    CopyImageToClipBoard
End Sub

Private Sub mnuAutoVerbProperties_Click()
    ShowImageProperties
End Sub

Public Sub ShowImageProperties()
    Dim ImagePropertiesForm As Form
    Set ImagePropertiesForm = New frmImageProperties
    
    ImagePropertiesForm.ShowProperties m_Picture, m_ActualFileName
End Sub

Private Sub picPicture_Resize()
    LoadPictureToShow m_Picture
End Sub

Private Sub txtFileName_GotFocus()
    UserControl_GotFocus
End Sub

Private Sub txtFileName_LostFocus()
    UserControl_LostFocus
End Sub

Private Sub UserControl_Click()
    If v_RightClick Then
        If m_AutoVerbMenu Then PopupMenu mnuAutoVerb Else RaiseEvent RightClick
    Else
        If m_AutoCheckUncheck Then
            If chkChecked.Value = vbChecked Then chkChecked.Value = vbUnchecked Else chkChecked.Value = vbChecked
        End If
        
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    shpFocus.Visible = True
End Sub

Private Sub UserControl_LostFocus()
    shpFocus.Visible = False
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then v_RightClick = True Else v_RightClick = False
End Sub

Private Sub UserControl_Resize()
'    On Error Resume Next

    If m_FileNameBox Or m_CheckBox Then
        picPicture.Move m_CellPadding, m_CellPadding, UserControl.ScaleWidth - (m_CellPadding * 2), UserControl.ScaleHeight - txtFileName.Height - (m_CellPadding * 3)
        
        If Not m_CheckBox Then 'FileNameBox visible
            txtFileName.Move m_CellPadding, UserControl.ScaleHeight - txtFileName.Height - m_CellPadding, UserControl.ScaleWidth - (m_CellPadding * 2), txtFileName.Height
        Else
            If Not m_FileNameBox Then 'CheckBox visible
                chkChecked.Move (UserControl.ScaleWidth - chkChecked.Width) / 2, (UserControl.ScaleHeight - txtFileName.Height - m_CellPadding) + ((txtFileName.Height - chkChecked.Height) / 2), chkChecked.Width, chkChecked.Height
            Else 'Both FileNameBox & CheckBox are visible
                chkChecked.Move m_CellPadding, (UserControl.ScaleHeight - txtFileName.Height - m_CellPadding) + ((txtFileName.Height - chkChecked.Height) / 2), chkChecked.Width, chkChecked.Height
                txtFileName.Move chkChecked.Width + (m_CellPadding * 2), UserControl.ScaleHeight - txtFileName.Height - m_CellPadding, UserControl.ScaleWidth - (m_CellPadding * 3) - chkChecked.Width, txtFileName.Height
            End If
        End If
    Else 'None of FileNameBox or CheckBox is visible
        picPicture.Move m_CellPadding, m_CellPadding, UserControl.ScaleWidth - (m_CellPadding * 2), UserControl.ScaleHeight - (m_CellPadding * 2)
    End If
    
    shpFocus.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,Picture
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    LoadPictureToShow m_Picture
    
    PropertyChanged "Picture"
    
    m_ActualFileName = ""
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFileName,txtFileName,-1,Text
Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the text contained in the control."
    FileName = txtFileName.Text
End Property

Public Property Let FileName(ByVal New_FileName As String)
    If InStr(New_FileName, "\") > 0 Then
        txtFileName.Text() = GetFileNameOnlyFromFullPath(New_FileName)
        m_ActualFileName = New_FileName
    Else
        txtFileName.Text() = New_FileName
        m_ActualFileName = ""
    End If
    
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub LoadPictureFromFile(FileName As String)
Attribute LoadPictureFromFile.VB_Description = "Load picture from a file."
    Set m_Picture = LoadPicture(FileName)
    LoadPictureToShow m_Picture
    
    txtFileName.Text = GetFileNameOnlyFromFullPath(FileName)
    m_ActualFileName = FileName
End Sub

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
'MappingInfo=txtFileName,txtFileName,-1,BackColor
Public Property Get FileNameBackColor() As OLE_COLOR
Attribute FileNameBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    FileNameBackColor = txtFileName.BackColor
End Property

Public Property Let FileNameBackColor(ByVal New_FileNameBackColor As OLE_COLOR)
    txtFileName.BackColor() = New_FileNameBackColor
    PropertyChanged "FileNameBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFileName,txtFileName,-1,Font
Public Property Get FileNameFont() As Font
Attribute FileNameFont.VB_Description = "Returns a Font object."
    Set FileNameFont = txtFileName.Font
End Property

Public Property Set FileNameFont(ByVal New_FileNameFont As Font)
    Set txtFileName.Font = New_FileNameFont
    PropertyChanged "FileNameFont"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,BorderStyle
Public Property Get PictureBorder() As Integer
Attribute PictureBorder.VB_Description = "Returns/sets the border style for an object."
    PictureBorder = picPicture.BorderStyle
End Property

Public Property Let PictureBorder(ByVal New_PictureBorder As Integer)
    picPicture.BorderStyle() = New_PictureBorder
    
    PropertyChanged "PictureBorder"
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
'MemberInfo=5
Public Sub Clear()
Attribute Clear.VB_Description = "Clear the picture."
    Set picPicture.Picture = Nothing
    txtFileName.Text = ""
    
    m_ActualFileName = ""
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CellPadding = m_def_CellPadding
    m_AutoCheckUncheck = m_def_AutoCheckUncheck
    m_AutoVerbMenu = m_def_AutoVerbMenu
    m_FileNameBox = m_def_FileNameBox
    m_CheckBox = m_def_CheckBox
    
    SetControlAttributesToPropertyVariables
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    txtFileName.Text = PropBag.ReadProperty("FileName", "")
    m_CellPadding = PropBag.ReadProperty("CellPadding", m_def_CellPadding)
    txtFileName.BackColor = PropBag.ReadProperty("FileNameBackColor", &HC0E0FF)
    Set txtFileName.Font = PropBag.ReadProperty("FileNameFont", Ambient.Font)
    picPicture.BorderStyle = PropBag.ReadProperty("PictureBorder", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    txtFileName.Locked = PropBag.ReadProperty("NoRename", True)
    m_AutoCheckUncheck = PropBag.ReadProperty("AutoCheckUncheck", m_def_AutoCheckUncheck)
    m_AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", m_def_AutoVerbMenu)
    m_FileNameBox = PropBag.ReadProperty("FileNameBox", True)
    m_CheckBox = PropBag.ReadProperty("CheckBox", True)
    shpFocus.BorderColor = PropBag.ReadProperty("FocusColor", vbBlue)
    shpFocus.BackColor = PropBag.ReadProperty("FocusBackColor", &H8000000D)
    txtFileName.BorderStyle = PropBag.ReadProperty("BorderStyleFileName", 0)
    m_ExternalViewer = PropBag.ReadProperty("ExternalViewer", "")
    m_ExternalEditor = PropBag.ReadProperty("ExternalEditor", "")
    m_ExternalPrinter = PropBag.ReadProperty("ExternalPrinter", "")
End Sub

Private Sub UserControl_Show()
    SetControlAttributesToPropertyVariables
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("FileName", txtFileName.Text, "")
    Call PropBag.WriteProperty("CellPadding", m_CellPadding, m_def_CellPadding)
    Call PropBag.WriteProperty("FileNameBackColor", txtFileName.BackColor, &HC0E0FF)
    Call PropBag.WriteProperty("FileNameFont", txtFileName.Font, Ambient.Font)
    Call PropBag.WriteProperty("PictureBorder", picPicture.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("NoRename", txtFileName.Locked, True)
    Call PropBag.WriteProperty("AutoCheckUncheck", m_AutoCheckUncheck, m_def_AutoCheckUncheck)
    Call PropBag.WriteProperty("AutoVerbMenu", m_AutoVerbMenu, m_def_AutoVerbMenu)
    Call PropBag.WriteProperty("FileNameBox", m_FileNameBox, True)
    Call PropBag.WriteProperty("CheckBox", m_CheckBox, True)
    Call PropBag.WriteProperty("FocusColor", shpFocus.BorderColor, vbBlue)
    Call PropBag.WriteProperty("FocusBackColor", shpFocus.BackColor, &H8000000D)
    Call PropBag.WriteProperty("BorderStyleFileName", txtFileName.BorderStyle, 0)
    Call PropBag.WriteProperty("ExternalViewer", m_ExternalViewer, "")
    Call PropBag.WriteProperty("ExternalEditor", m_ExternalEditor, "")
    Call PropBag.WriteProperty("ExternalPrinter", m_ExternalPrinter, "")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFileName,txtFileName,-1,Locked
Public Property Get NoRename() As Boolean
Attribute NoRename.VB_Description = "Determines whether a control can be edited."
    NoRename = txtFileName.Locked
End Property

Public Property Let NoRename(ByVal New_NoRename As Boolean)
    txtFileName.Locked() = New_NoRename
    PropertyChanged "NoRename"
End Property

Public Property Get Checked() As CheckBoxConstants
    Checked = chkChecked.Value
End Property

Public Property Let Checked(New_Checked As CheckBoxConstants)
    chkChecked.Value = New_Checked
End Property

Public Property Get AutoCheckUncheck() As Boolean
    AutoCheckUncheck = m_AutoCheckUncheck
End Property

Public Property Let AutoCheckUncheck(New_AutoCheckUncheck As Boolean)
    m_AutoCheckUncheck = New_AutoCheckUncheck
    PropertyChanged "AutoCheckUncheck"
End Property

Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = m_AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(New_AutoVerbMenu As Boolean)
    m_AutoVerbMenu = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

Public Property Get FileNameBox() As Boolean
    FileNameBox = m_FileNameBox
End Property

Public Property Let FileNameBox(New_FileNameBox As Boolean)
    m_FileNameBox = New_FileNameBox
    PropertyChanged "FileNameBox"
    
    SetControlAttributesToPropertyVariables
    UserControl_Resize
End Property

Public Property Get CheckBox() As Boolean
    CheckBox = m_CheckBox
End Property

Public Property Let CheckBox(New_CheckBox As Boolean)
    m_CheckBox = New_CheckBox
    PropertyChanged "CheckBox"
    
    SetControlAttributesToPropertyVariables
    UserControl_Resize
End Property

Private Sub SetControlAttributesToPropertyVariables()
    txtFileName.Visible = m_FileNameBox
    chkChecked.Visible = m_CheckBox
End Sub

Private Sub LoadPictureToShow(PIC As StdPicture)
    If PIC Is Nothing Then Exit Sub
    
    Dim posLeft As Long, posTop As Long, picWidth As Long, picHeight As Long
    
    picWidth = picPicture.ScaleWidth
    picHeight = picWidth / PIC.Width * PIC.Height
    
    If picHeight > picPicture.ScaleHeight Then
        picHeight = picPicture.ScaleHeight
        picWidth = picHeight / PIC.Height * PIC.Width
    End If
    
    posLeft = (picPicture.ScaleWidth - picWidth) / 2
    posTop = (picPicture.ScaleHeight - picHeight) / 2
    
    CopyImage PIC, picPicture, 0, 0, , , posLeft, posTop, picWidth, picHeight
End Sub

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = shpFocus.BorderColor
End Property

Public Property Let FocusColor(New_FocustColor As OLE_COLOR)
    shpFocus.BorderColor = New_FocustColor
End Property

Public Property Get FocusBackColor() As OLE_COLOR
    FocusBackColor = shpFocus.BackColor
End Property

Public Property Let FocusBackColor(New_FocustBackColor As OLE_COLOR)
    shpFocus.BackColor = New_FocustBackColor
End Property

Public Property Get BorderStyleFileName() As Integer
    BorderStyleFileName = txtFileName.BorderStyle
End Property

Public Property Let BorderStyleFileName(New_BorderStyleFileName As Integer)
    txtFileName.BorderStyle = New_BorderStyleFileName
    PropertyChanged "BorderStyleFileName"
End Property

Public Property Get ExternalViewer() As String
    ExternalViewer = m_ExternalViewer
End Property

Public Property Let ExternalViewer(New_ExternalViewer As String)
    m_ExternalViewer = New_ExternalViewer
    PropertyChanged "ExternalViewer"
End Property

Public Property Get ExternalEditor() As String
    ExternalEditor = m_ExternalEditor
End Property

Public Property Let ExternalEditor(New_ExternalEditor As String)
    m_ExternalEditor = New_ExternalEditor
    PropertyChanged "ExternalEditor"
End Property

Public Property Get ExternalPrinter() As String
    ExternalPrinter = m_ExternalPrinter
End Property

Public Property Let ExternalPrinter(New_ExternalPrinter As String)
    m_ExternalPrinter = New_ExternalPrinter
    PropertyChanged "ExternalPrinter"
End Property

Public Sub OpenWithExternalViewer()
    If m_ActualFileName <> "" And m_ExternalViewer <> "" Then Shell Replace(m_ExternalViewer, "%FILENAME%", m_ActualFileName), vbNormalFocus
End Sub

Public Sub OpenWithExternalEditor()
    If m_ActualFileName <> "" And m_ExternalEditor <> "" Then Shell Replace(m_ExternalEditor, "%FILENAME%", m_ActualFileName), vbNormalFocus
End Sub

Public Sub OpenWithExternalPrinter()
    If m_ActualFileName <> "" And m_ExternalPrinter <> "" Then Shell Replace(m_ExternalPrinter, "%FILENAME%", m_ActualFileName), vbNormalFocus
End Sub

