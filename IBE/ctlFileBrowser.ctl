VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ctlFileBrowser 
   Alignable       =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   Begin prjImageBrowser.chameleonButton cmdBrowse 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16761024
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "ctlFileBrowser.ctx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open..."
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   630
   End
End
Attribute VB_Name = "ctlFileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_ButtonWidth = 25
'Property Variables:
Dim m_ButtonWidth As Long
'Event Declarations:
Event Change() 'MappingInfo=txtFile,txtFile,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cdlFile,cdlFile,-1,FileName
Public Property Get File() As String
Attribute File.VB_Description = "Returns/sets the path and filename of a selected file."
    File = txtFile.Text 'cdlFile.FileName
End Property

Public Property Let File(ByVal New_File As String)
    txtFile.Text() = New_File
    PropertyChanged "File"
    
    cdlFile.FileName() = New_File
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdBrowse,cmdBrowse,-1,Caption
Public Property Get ButtonCaption() As String
Attribute ButtonCaption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    ButtonCaption = cmdBrowse.Caption
End Property

Public Property Let ButtonCaption(ByVal New_ButtonCaption As String)
    cmdBrowse.Caption() = New_ButtonCaption
    PropertyChanged "ButtonCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cdlFile,cdlFile,-1,Filter
Public Property Get Filter() As String
Attribute Filter.VB_Description = "Returns/sets the filters that are displayed in the Type list box of a dialog box."
    Filter = cdlFile.Filter
End Property

Public Property Let Filter(ByVal New_Filter As String)
    cdlFile.Filter() = New_Filter
    PropertyChanged "Filter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cdlFile,cdlFile,-1,FilterIndex
Public Property Get FilterIndex() As Integer
Attribute FilterIndex.VB_Description = "Returns/sets a default filter for an Open or Save As dialog box."
    FilterIndex = cdlFile.FilterIndex
End Property

Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
    cdlFile.FilterIndex() = New_FilterIndex
    PropertyChanged "FilterIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cdlFile,cdlFile,-1,InitDir
Public Property Get InitialPath() As String
Attribute InitialPath.VB_Description = "Returns/sets the initial file directory."
    InitialPath = cdlFile.InitDir
End Property

Public Property Let InitialPath(ByVal New_InitialPath As String)
    cdlFile.InitDir() = New_InitialPath
    PropertyChanged "InitialPath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFile,txtFile,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtFile.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtFile.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cdlFile,cdlFile,-1,DialogTitle
Public Property Get DialogCaption() As String
Attribute DialogCaption.VB_Description = "Sets the string displayed in the title bar of the dialog box."
    DialogCaption = cdlFile.DialogTitle
End Property

Public Property Let DialogCaption(ByVal New_DialogCaption As String)
    cdlFile.DialogTitle() = New_DialogCaption
    PropertyChanged "DialogCaption"
End Property

Private Sub cmdBrowse_Click()
    On Error GoTo ERROR_HANDLER_cmdBrowse_Click

    cdlFile.ShowOpen
    
    txtFile = cdlFile.FileName
    txtFile.SelStart = Len(txtFile) + 1

EXIT_cmdBrowse_Click:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_cmdBrowse_Click:
    Select Case Err.Number
    Case 20477 'Invalid file name
        cdlFile.FileName = ""
    Case 32755 'Cancel was selected
    Case Else
        If MsgBox("Error in Sub cmdBrowse_Click() of User Control ctlFileBrowser[ctlFileBrowser.ctl] of Project prjImageBrowser[prjImageBrowser.vbp]" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "prjImageBrowser: Application error!") = vbNo Then Resume EXIT_cmdBrowse_Click
    End Select
    
    Resume Next
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Caption = PropBag.ReadProperty("Caption", "Caption")
    txtFile.Text = PropBag.ReadProperty("File", "")
    cmdBrowse.Caption = PropBag.ReadProperty("ButtonCaption", "Caption")
    cdlFile.Filter = PropBag.ReadProperty("Filter", "")
    cdlFile.FilterIndex = PropBag.ReadProperty("FilterIndex", 0)
    cdlFile.InitDir = PropBag.ReadProperty("InitialPath", "")
    txtFile.Locked = PropBag.ReadProperty("Locked", False)
    cdlFile.DialogTitle = PropBag.ReadProperty("DialogCaption", "Open...")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    txtFile.BackColor = PropBag.ReadProperty("FileBackColor", &H80000005)
    lblCaption.ForeColor = PropBag.ReadProperty("CaptionColor", &H80000012)
    txtFile.ForeColor = PropBag.ReadProperty("FileColor", &H80000008)
    m_ButtonWidth = PropBag.ReadProperty("ButtonWidth", m_def_ButtonWidth)
    Set lblCaption.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    Set txtFile.Font = PropBag.ReadProperty("FileFont", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

Dim CaptionPadding As Long
CaptionPadding = 5

txtFile.Move lblCaption.Width + CaptionPadding, 0, UserControl.ScaleWidth - lblCaption.Width - m_ButtonWidth - CaptionPadding, UserControl.ScaleHeight
cmdBrowse.Move UserControl.ScaleWidth - m_ButtonWidth, 0, m_ButtonWidth, UserControl.ScaleHeight
lblCaption.Move 0, (txtFile.Height - lblCaption.Height) / 2, lblCaption.Width, lblCaption.Height
End Sub

Private Sub UserControl_Show()
    
    EnableSet UserControl.Enabled
    UserControl_Resize

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("File", txtFile.Text, "")
    Call PropBag.WriteProperty("ButtonCaption", cmdBrowse.Caption, "Caption")
    Call PropBag.WriteProperty("Filter", cdlFile.Filter, "")
    Call PropBag.WriteProperty("FilterIndex", cdlFile.FilterIndex, 0)
    Call PropBag.WriteProperty("InitialPath", cdlFile.InitDir, "")
    Call PropBag.WriteProperty("Locked", txtFile.Locked, False)
    Call PropBag.WriteProperty("DialogCaption", cdlFile.DialogTitle, "Open...")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FileBackColor", txtFile.BackColor, &H80000005)
    Call PropBag.WriteProperty("CaptionColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("FileColor", txtFile.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ButtonWidth", m_ButtonWidth, m_def_ButtonWidth)
    Call PropBag.WriteProperty("CaptionFont", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FileFont", txtFile.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

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
'MappingInfo=txtFile,txtFile,-1,BackColor
Public Property Get FileBackColor() As OLE_COLOR
Attribute FileBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    FileBackColor = txtFile.BackColor
End Property

Public Property Let FileBackColor(ByVal New_FileBackColor As OLE_COLOR)
    txtFile.BackColor() = New_FileBackColor
    PropertyChanged "FileBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get CaptionColor() As OLE_COLOR
Attribute CaptionColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    CaptionColor = lblCaption.ForeColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
    lblCaption.ForeColor() = New_CaptionColor
    PropertyChanged "CaptionColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFile,txtFile,-1,ForeColor
Public Property Get FileColor() As OLE_COLOR
Attribute FileColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    FileColor = txtFile.ForeColor
End Property

Public Property Let FileColor(ByVal New_FileColor As OLE_COLOR)
    txtFile.ForeColor() = New_FileColor
    PropertyChanged "FileColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonWidth() As Long
Attribute ButtonWidth.VB_Description = "Width of the browse button."
    ButtonWidth = m_ButtonWidth
End Property

Public Property Let ButtonWidth(ByVal New_ButtonWidth As Long)
    m_ButtonWidth = New_ButtonWidth
    PropertyChanged "ButtonWidth"
    
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ButtonWidth = m_def_ButtonWidth
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get CaptionFont() As Font
Attribute CaptionFont.VB_Description = "Returns a Font object."
    Set CaptionFont = lblCaption.Font
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
    Set lblCaption.Font = New_CaptionFont
    PropertyChanged "CaptionFont"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFile,txtFile,-1,Font
Public Property Get FileFont() As Font
Attribute FileFont.VB_Description = "Returns a Font object."
    Set FileFont = txtFile.Font
End Property

Public Property Set FileFont(ByVal New_FileFont As Font)
    Set txtFile.Font = New_FileFont
    PropertyChanged "FileFont"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    EnableSet New_Enabled
End Property

Private Sub EnableSet(Enabled As Boolean)
    lblCaption.Enabled = Enabled
    txtFile.Enabled = Enabled
    cmdBrowse.Enabled = Enabled
End Sub

Private Sub txtFile_Change()
    RaiseEvent Change
End Sub

