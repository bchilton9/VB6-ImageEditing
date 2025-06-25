VERSION 5.00
Begin VB.UserControl ctlThumbNailList 
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   LockControls    =   -1  'True
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   Begin VB.FileListBox flbPictures 
      Height          =   1650
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.VScrollBar vsrThumbList 
      Enabled         =   0   'False
      Height          =   1935
      LargeChange     =   200
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hsrThumbList 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   200
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.PictureBox picThumbListContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox picThumbList 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   1
         Top             =   0
         Width           =   1695
         Begin prjImageBrowser.ctlThumbNail ctlThumbNail 
            Height          =   1695
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2990
            FileNameBackColor=   12648384
            BeginProperty FileNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureBorder   =   0
            BackColor       =   12648447
         End
      End
   End
End
Attribute VB_Name = "ctlThumbNailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Property Variables:
Dim m_CellPadding As Long
Dim m_Columns As Long
Dim m_Rows As Long
Dim m_SelectedItem As ctlThumbNail

'Default Property Values:
Const m_def_CellPadding = 5

'Events
Event Click()
Event RightClick()
Event DblClick()
Event ItemClick(Index As Integer)
Event ItemRightClick(Index As Integer)
Event ItemDblClick(Index As Integer)
Event LoadFromPath(SearchPath As String, FileName As String, ImageIndex As Long, ImageLoaded As Long, ImageTotal As Long)
Event BeforeLoadFromPath(SearchPath As String, ImageTotal As Long)
Event AfterLoadFromPath(SearchPath As String, ImageTotal As Long)

'Private variables
Private v_RightClick As Boolean

Public Property Get HeightThumbNail() As Long
    HeightThumbNail = ctlThumbNail(0).Height
End Property

Public Property Let HeightThumbNail(New_HeightThumbNail As Long)
    ctlThumbNail(0).Height = New_HeightThumbNail
    PropertyChanged "HeightThumbNail"
    
    ArrangeThumbNails
End Property

Public Property Get WidthThumbNail() As Long
    WidthThumbNail = ctlThumbNail(0).Width
End Property

Public Property Let WidthThumbNail(New_WidthThumbNail As Long)
    ctlThumbNail(0).Width = New_WidthThumbNail
    PropertyChanged "WidthThumbNail"
    
    ArrangeThumbNails
End Property

Private Sub ctlThumbNail_Click(Index As Integer)
    RaiseEvent ItemClick(Index)
End Sub

Private Sub ctlThumbNail_DblClick(Index As Integer)
    RaiseEvent ItemDblClick(Index)
End Sub

Private Sub ctlThumbNail_GotFocus(Index As Integer)
    Set m_SelectedItem = ctlThumbNail(Index)
End Sub

Private Sub ctlThumbNail_RightClick(Index As Integer)
    RaiseEvent ItemRightClick(Index)
End Sub

Private Sub hsrThumbList_Change()
    hsrThumbList_Scroll
End Sub

Private Sub picThumbList_Click()
    picThumbListContainer_Click
End Sub

Private Sub picThumbList_DblClick()
    picThumbListContainer_DblClick
End Sub

Private Sub picThumbList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picThumbListContainer_MouseUp Button, Shift, picThumbList.Left + X, picThumbList.Top + Y
End Sub

Private Sub picThumbListContainer_Click()
    UserControl_Click
End Sub

Private Sub picThumbListContainer_DblClick()
    UserControl_DblClick
End Sub

Private Sub picThumbListContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, picThumbListContainer.Left + X, picThumbListContainer.Top + Y
End Sub

Private Sub UserControl_Click()
    If v_RightClick Then
        RaiseEvent RightClick
    Else
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then v_RightClick = True Else v_RightClick = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    picThumbListContainer.Move 0, 0, UserControl.ScaleWidth - vsrThumbList.Width, UserControl.ScaleHeight - hsrThumbList.Height
    
    vsrThumbList.Move UserControl.ScaleWidth - vsrThumbList.Width, 0, vsrThumbList.Width, picThumbListContainer.Height
    hsrThumbList.Move 0, UserControl.ScaleHeight - hsrThumbList.Height, picThumbListContainer.Width, hsrThumbList.Height
    
    If (m_Columns <> ColumnsAvailableForArrangement) Or (m_Rows <> RowsRequiredForArrangement) Then ArrangeThumbNails
End Sub

Private Sub SetScrollBarValues()
    vsrThumbList.Max = picThumbList.Height - picThumbListContainer.ScaleHeight
    hsrThumbList.Max = picThumbList.Width - picThumbListContainer.ScaleWidth
    
    If vsrThumbList.Max < 1 Then vsrThumbList.Enabled = False Else vsrThumbList.Enabled = True
    If hsrThumbList.Max < 1 Then hsrThumbList.Enabled = False Else hsrThumbList.Enabled = True
End Sub

Private Sub vsrThumbList_Change()
    vsrThumbList_Scroll
End Sub

Private Sub vsrThumbList_Scroll()
    picThumbList.Top = 0 - vsrThumbList.Value
End Sub

Private Sub hsrThumbList_Scroll()
    picThumbList.Left = 0 - hsrThumbList.Value
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picThumbListContainer,picThumbListContainer,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picThumbListContainer.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picThumbListContainer.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    
    picThumbList.BackColor = New_BackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picThumbList,picThumbList,-1,BackColor
Public Property Get BackColorList() As OLE_COLOR
Attribute BackColorList.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColorList = picThumbList.BackColor
End Property

Public Property Let BackColorList(ByVal New_BackColorList As OLE_COLOR)
    picThumbList.BackColor() = New_BackColorList
    PropertyChanged "BackColorList"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picThumbListContainer,picThumbListContainer,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = picThumbListContainer.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    picThumbListContainer.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9
Public Function AddItem(Optional picImage As StdPicture, Optional FileName As String, Optional Index As Integer, Optional NoArrange As Boolean = False) As ctlThumbNail
Attribute AddItem.VB_Description = "Add a new item to the collection."
    Load ctlThumbNail(ctlThumbNail.UBound + 1)
    
    If Not picImage Is Nothing Then Set ctlThumbNail(ctlThumbNail.UBound).Picture = picImage
    If FileName <> "" Then ctlThumbNail(ctlThumbNail.UBound).FileName = FileName
    ctlThumbNail(ctlThumbNail.UBound).Visible = True
    If Not NoArrange Then ArrangeThumbNails  'True
'    ArrangeThumbNails True
    
    Set AddItem = ctlThumbNail(ctlThumbNail.UBound)
End Function

Public Sub Clear()
    Dim ItemCounter As Long
    For ItemCounter = ctlThumbNail.UBound To 1 Step -1
        Unload ctlThumbNail(ItemCounter)
    Next
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picThumbListContainer.BackColor = PropBag.ReadProperty("BackColor", &HC0C0FF)
    picThumbList.BackColor = PropBag.ReadProperty("BackColorList", &HC0E0FF)
    ctlThumbNail(0).BackColor = PropBag.ReadProperty("BackColorThumbNail", &HC0FFFF)
    ctlThumbNail(0).FileNameBackColor = PropBag.ReadProperty("BackColorFileName", &HC0FFC0)
    picThumbListContainer.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    ctlThumbNail(0).PictureBorder = PropBag.ReadProperty("BorderStylePicture", 1)
    Set ctlThumbNail(0).FileNameFont = PropBag.ReadProperty("FontFileName", Ambient.Font)
    ctlThumbNail(0).NoRename = PropBag.ReadProperty("NoRename", True)
    ctlThumbNail(0).NoRename = PropBag.ReadProperty("NoRename", True)
    ctlThumbNail(0).BackColor = PropBag.ReadProperty("BackColorThumbNail", &HC0FFFF)
    ctlThumbNail(0).FileNameBackColor = PropBag.ReadProperty("BackColorFileName", &HC0FFC0)
    ctlThumbNail(0).PictureBorder = PropBag.ReadProperty("BorderStylePicture", 1)
    Set ctlThumbNail(0).FileNameFont = PropBag.ReadProperty("FontFileName", Ambient.Font)
    m_CellPadding = PropBag.ReadProperty("CellPadding", m_def_CellPadding)
    ctlThumbNail(0).CellPadding = PropBag.ReadProperty("CellPaddingThumbNail", 4)
    ctlThumbNail(0).Height = PropBag.ReadProperty("HeightThumbNail", 113)
    ctlThumbNail(0).Width = PropBag.ReadProperty("WidthThumbNail", 105)
    ctlThumbNail(0).AutoCheckUncheck = PropBag.ReadProperty("AutoCheckUncheck", True)
    ctlThumbNail(0).AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    ctlThumbNail(0).FileNameBox = PropBag.ReadProperty("FileNameBox", True)
    ctlThumbNail(0).CheckBox = PropBag.ReadProperty("CheckBox", True)
    ctlThumbNail(0).FocusColor = PropBag.ReadProperty("FocusColor", vbBlue)
    ctlThumbNail(0).FocusBackColor = PropBag.ReadProperty("FocusBackColor", &H8000000D)
    ctlThumbNail(0).BorderStyleFileName = PropBag.ReadProperty("BorderStyleFileName", 0)
    ctlThumbNail(0).ExternalViewer = PropBag.ReadProperty("ExternalViewer", "")
    ctlThumbNail(0).ExternalEditor = PropBag.ReadProperty("ExternalEditor", "")
    ctlThumbNail(0).ExternalPrinter = PropBag.ReadProperty("ExternalPrinter", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", picThumbListContainer.BackColor, &HC0C0FF)
    Call PropBag.WriteProperty("BackColorList", picThumbList.BackColor, &HC0E0FF)
    Call PropBag.WriteProperty("BackColorThumbNail", ctlThumbNail(0).BackColor, &HC0FFFF)
    Call PropBag.WriteProperty("BackColorFileName", ctlThumbNail(0).FileNameBackColor, &HC0FFC0)
    Call PropBag.WriteProperty("BorderStyle", picThumbListContainer.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderStylePicture", ctlThumbNail(0).PictureBorder, 1)
    Call PropBag.WriteProperty("FontFileName", ctlThumbNail(0).FileNameFont, Ambient.Font)
    Call PropBag.WriteProperty("NoRename", ctlThumbNail(0).NoRename, True)
    Call PropBag.WriteProperty("NoRename", ctlThumbNail(0).NoRename, True)
    Call PropBag.WriteProperty("BackColorThumbNail", ctlThumbNail(0).BackColor, &HC0FFFF)
    Call PropBag.WriteProperty("BackColorFileName", ctlThumbNail(0).FileNameBackColor, &HC0FFC0)
    Call PropBag.WriteProperty("BorderStylePicture", ctlThumbNail(0).PictureBorder, 1)
    Call PropBag.WriteProperty("FontFileName", ctlThumbNail(0).FileNameFont, Ambient.Font)
    Call PropBag.WriteProperty("CellPadding", m_CellPadding, m_def_CellPadding)
    Call PropBag.WriteProperty("CellPaddingThumbNail", ctlThumbNail(0).CellPadding, 4)
    Call PropBag.WriteProperty("HeightThumbNail", ctlThumbNail(0).Height, 4)
    Call PropBag.WriteProperty("WidthThumbNail", ctlThumbNail(0).Width, 4)
    Call PropBag.WriteProperty("AutoCheckUncheck", ctlThumbNail(0).AutoCheckUncheck, True)
    Call PropBag.WriteProperty("AutoVerbMenu", ctlThumbNail(0).AutoVerbMenu, False)
    Call PropBag.WriteProperty("FileNameBox", ctlThumbNail(0).FileNameBox, True)
    Call PropBag.WriteProperty("CheckBox", ctlThumbNail(0).CheckBox, True)
    Call PropBag.WriteProperty("FocusColor", ctlThumbNail(0).FocusColor, vbBlue)
    Call PropBag.WriteProperty("FocusBackColor", ctlThumbNail(0).FocusBackColor, &H8000000D)
    Call PropBag.WriteProperty("BorderStyleFileName", ctlThumbNail(0).BorderStyleFileName, 0)
    Call PropBag.WriteProperty("ExternalViewer", ctlThumbNail(0).ExternalViewer, "")
    Call PropBag.WriteProperty("ExternalEditor", ctlThumbNail(0).ExternalEditor, "")
    Call PropBag.WriteProperty("ExternalPrinter", ctlThumbNail(0).ExternalPrinter, "")
End Sub

Private Sub ArrangeThumbNails(Optional AddLastItemOnly As Boolean = False)
    Dim Rows As Long, Columns As Long
    Dim RowCounter As Long, ColumnCounter As Long
'    Dim Remainings As Long
    Dim CurrentItem As Long
    
    If ctlThumbNail.Count < 2 Then Exit Sub
    
    Columns = ColumnsAvailableForArrangement 'Int((picThumbListContainer.ScaleWidth - m_CellPadding) / (ctlThumbNail(0).Width + CellPadding))
    Rows = RowsRequiredForArrangement 'Int(ctlThumbNail.UBound / Columns)
    
'    Remainings = ctlThumbNail.UBound - (Columns * Rows)
'    If Remainings > 0 Then Rows = Rows + 1
    
    For RowCounter = 1 To Rows
        For ColumnCounter = 1 To Columns
            CurrentItem = ((RowCounter - 1) * Columns) + ColumnCounter
            
            If CurrentItem <= ctlThumbNail.UBound Then
'                ctlThumbNail(CurrentItem).Move m_CellPadding + (m_CellPadding + ctlThumbNail(0).Width) * (ColumnCounter - 1), m_CellPadding + (m_CellPadding + ctlThumbNail(0).Height) * (RowCounter - 1)
                If AddLastItemOnly Then
                    If CurrentItem = ctlThumbNail.UBound Then
                        ctlThumbNail(CurrentItem).Move m_CellPadding + (m_CellPadding + ctlThumbNail(0).Width) * (ColumnCounter - 1), m_CellPadding + (m_CellPadding + ctlThumbNail(0).Height) * (RowCounter - 1)
                    End If
                Else
                    ctlThumbNail(CurrentItem).Move m_CellPadding + (m_CellPadding + ctlThumbNail(0).Width) * (ColumnCounter - 1), m_CellPadding + (m_CellPadding + ctlThumbNail(0).Height) * (RowCounter - 1)
                End If
            End If
            If CurrentItem = ctlThumbNail.UBound Then Exit For
        Next
        If CurrentItem = ctlThumbNail.UBound Then Exit For
    Next
    
    picThumbList.Move 0, 0, (m_CellPadding + ctlThumbNail(0).Width) * Columns + m_CellPadding, (m_CellPadding + ctlThumbNail(0).Height) * Rows + m_CellPadding
    
    SetScrollBarValues
    
    m_Columns = Columns
    m_Rows = Rows
End Sub

Private Function RowsRequiredForArrangement() As Long
    RowsRequiredForArrangement = Int(ctlThumbNail.UBound / ColumnsAvailableForArrangement)
    If (ctlThumbNail.UBound - (Columns * Rows)) > 0 Then RowsRequiredForArrangement = RowsRequiredForArrangement + 1
End Function

Private Function ColumnsAvailableForArrangement() As Long
    ColumnsAvailableForArrangement = Int((picThumbListContainer.ScaleWidth - m_CellPadding) / (ctlThumbNail(0).Width + CellPadding))
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,NoRename
Public Property Get NoRename() As Boolean
Attribute NoRename.VB_Description = "Determines whether a control can be edited."
    NoRename = ctlThumbNail(0).NoRename
End Property

Public Property Let NoRename(ByVal New_NoRename As Boolean)
    ctlThumbNail(0).NoRename() = New_NoRename
    PropertyChanged "NoRename"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,BackColor
Public Property Get BackColorThumbNail() As OLE_COLOR
Attribute BackColorThumbNail.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColorThumbNail = ctlThumbNail(0).BackColor
End Property

Public Property Let BackColorThumbNail(ByVal New_BackColorThumbNail As OLE_COLOR)
    ctlThumbNail(0).BackColor() = New_BackColorThumbNail
    PropertyChanged "BackColorThumbNail"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,FileNameBackColor
Public Property Get BackColorFileName() As OLE_COLOR
Attribute BackColorFileName.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColorFileName = ctlThumbNail(0).FileNameBackColor
End Property

Public Property Let BackColorFileName(ByVal New_BackColorFileName As OLE_COLOR)
    ctlThumbNail(0).FileNameBackColor() = New_BackColorFileName
    PropertyChanged "BackColorFileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,PictureBorder
Public Property Get BorderStylePicture() As Integer
Attribute BorderStylePicture.VB_Description = "Returns/sets the border style for an object."
    BorderStylePicture = ctlThumbNail(0).PictureBorder
End Property

Public Property Let BorderStylePicture(ByVal New_BorderStylePicture As Integer)
    ctlThumbNail(0).PictureBorder() = New_BorderStylePicture
    PropertyChanged "BorderStylePicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,FileNameFont
Public Property Get FontFileName() As Font
Attribute FontFileName.VB_Description = "Returns a Font object."
    Set FontFileName = ctlThumbNail(0).FileNameFont
End Property

Public Property Set FontFileName(ByVal New_FontFileName As Font)
    Set ctlThumbNail(0).FileNameFont = New_FontFileName
    PropertyChanged "FontFileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub RemoveItem(Index As Integer)
Attribute RemoveItem.VB_Description = "Remove an item off the collection."
     Unload ctlThumbNail(Index)
     
     ArrangeThumbNails
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,1,2,0
Public Property Get Item(Index As Integer) As ctlThumbNail
Attribute Item.VB_Description = "Returns an item."
Attribute Item.VB_MemberFlags = "400"
    Set Item = ctlThumbNail(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get CellPadding() As Long
    CellPadding = m_CellPadding
End Property

Public Property Let CellPadding(ByVal New_CellPadding As Long)
    m_CellPadding = New_CellPadding
    PropertyChanged "CellPadding"
    
    ArrangeThumbNails
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlThumbNail(0),ctlThumbNail,0,CellPadding
Public Property Get CellPaddingThumbNail() As Long
    CellPaddingThumbNail = ctlThumbNail(0).CellPadding
End Property

Public Property Let CellPaddingThumbNail(ByVal New_CellPaddingThumbNail As Long)
    ctlThumbNail(0).CellPadding() = New_CellPaddingThumbNail
    PropertyChanged "CellPaddingThumbNail"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CellPadding = m_def_CellPadding
    
End Sub

Public Property Get AutoCheckUncheck() As Boolean
    AutoCheckUncheck = ctlThumbNail(0).AutoCheckUncheck
End Property

Public Property Let AutoCheckUncheck(New_AutoCheckUncheck As Boolean)
    ctlThumbNail(0).AutoCheckUncheck = New_AutoCheckUncheck
    PropertyChanged "AutoCheckUncheck"
    
'    If ctlThumbNail.Count > 1 Then
'        Dim CTL As Control
'        For Each CTL In Controls
'            If TypeOf CTL Is ctlThumbNail Then CTL.AutoCheckUncheck = New_AutoCheckUncheck
'        Next
'    End If
End Property

Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = ctlThumbNail(0).AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(New_AutoVerbMenu As Boolean)
    ctlThumbNail(0).AutoVerbMenu = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
    
'    If ctlThumbNail.Count > 1 Then
'        Dim CTL As Control
'        For Each CTL In Controls
'            If TypeOf CTL Is ctlThumbNail Then CTL.AutoVerbMenu = New_AutoVerbMenu
'        Next
'    End If
End Property

Public Sub LoadPicturesFromPath(SearchPath As String, Optional FileFilter As String = "")
    Clear
    
    flbPictures.Path = SearchPath
    flbPictures.Pattern = FileFilter
    
    If flbPictures.ListCount > 0 Then
        RaiseEvent BeforeLoadFromPath(SearchPath, flbPictures.ListCount)
    
        Dim PictureCounter As Long
        For PictureCounter = 0 To flbPictures.ListCount - 1
            AddItem LoadPicture(SearchPath & "\" & flbPictures.List(PictureCounter)), SearchPath & "\" & flbPictures.List(PictureCounter), , True
            
            RaiseEvent LoadFromPath(SearchPath, flbPictures.List(PictureCounter), PictureCounter, PictureCounter + 1, flbPictures.ListCount)
        Next

        RaiseEvent AfterLoadFromPath(SearchPath, flbPictures.ListCount)
    
        ArrangeThumbNails
    End If
End Sub

Public Property Get FileNameBox() As Boolean
    FileNameBox = ctlThumbNail(0).FileNameBox
End Property

Public Property Let FileNameBox(New_FileNameBox As Boolean)
    ctlThumbNail(0).FileNameBox = New_FileNameBox
    PropertyChanged "FileNameBox"
End Property

Public Property Get CheckBox() As Boolean
    CheckBox = ctlThumbNail(0).CheckBox
End Property

Public Property Let CheckBox(New_CheckBox As Boolean)
    ctlThumbNail(0).CheckBox = New_CheckBox
    PropertyChanged "CheckBox"
End Property

Public Property Get Columns() As Long
    Columns = m_Columns
End Property

Public Property Get Rows() As Long
    Rows = m_Rows
End Property

Public Property Get SelectedItem() As ctlThumbNail
    Set SelectedItem = m_SelectedItem
End Property

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = ctlThumbNail(0).FocusColor
End Property

Public Property Let FocusColor(New_FocustColor As OLE_COLOR)
    ctlThumbNail(0).FocusColor = New_FocustColor
End Property

Public Property Get FocusBackColor() As OLE_COLOR
    FocusBackColor = ctlThumbNail(0).FocusBackColor
End Property

Public Property Let FocusBackColor(New_FocustBackColor As OLE_COLOR)
    ctlThumbNail(0).FocusBackColor = New_FocustBackColor
End Property

Public Property Get BorderStyleFileName() As Integer
    BorderStyleFileName = ctlThumbNail(0).BorderStyleFileName
End Property

Public Property Let BorderStyleFileName(New_BorderStyleFileName As Integer)
    ctlThumbNail(0).BorderStyleFileName = New_BorderStyleFileName
    PropertyChanged "BorderStyleFileName"
End Property

Public Property Get ExternalViewer() As String
    ExternalViewer = ctlThumbNail(0).ExternalViewer
End Property

Public Property Let ExternalViewer(New_ExternalViewer As String)
    ctlThumbNail(0).ExternalViewer = New_ExternalViewer
    PropertyChanged "ExternalViewer"
End Property

Public Property Get ExternalEditor() As String
    ExternalEditor = ctlThumbNail(0).ExternalEditor
End Property

Public Property Let ExternalEditor(New_ExternalEditor As String)
    ctlThumbNail(0).ExternalEditor = New_ExternalEditor
    PropertyChanged "ExternalEditor"
End Property

Public Property Get ExternalPrinter() As String
    ExternalPrinter = ctlThumbNail(0).ExternalPrinter
End Property

Public Property Let ExternalPrinter(New_ExternalPrinter As String)
    ctlThumbNail(0).ExternalPrinter = New_ExternalPrinter
    PropertyChanged "ExternalPrinter"
End Property

Public Property Get ItemCount() As Long
    ItemCount = ctlThumbNail.UBound
End Property
