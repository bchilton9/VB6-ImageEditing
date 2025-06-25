VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Object = "{0145C507-DA70-4335-9BA5-1F01A5461BE8}#4.0#0"; "WOWHPBar.ocx"
Begin VB.Form frmBrowser 
   Caption         =   "Image Browser"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12105
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picExplorer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   5
      Top             =   720
      Width           =   2655
      Begin CCRPFolderTV6.FolderTreeview ftvExplorer 
         Height          =   2880
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5080
         Appearance      =   0
         IntegralHeight  =   0   'False
      End
      Begin prjImageBrowser.ctlThumbNail tmbMiniPreview 
         Height          =   1455
         Left            =   480
         TabIndex        =   7
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         CellPadding     =   0
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
         AutoCheckUncheck=   0   'False
         AutoVerbMenu    =   -1  'True
         FileNameBox     =   0   'False
         CheckBox        =   0   'False
      End
      Begin VB.Image sptExplorerMiniPreview 
         Appearance      =   0  'Flat
         Height          =   75
         Left            =   240
         MousePointer    =   7  'Size N S
         Picture         =   "frmBrowser.frx":0ECA
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1695
      End
   End
   Begin VB.PictureBox picBrowser 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   3120
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   960
      Width           =   5175
      Begin prjWOWHPBar.ctlWOWHProgress prgImageLoading 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   661
         SlotPicture     =   "frmBrowser.frx":1AC4
         Value           =   100
         SlotHeight      =   17
         TitleOffsetY    =   "3"
         Title           =   "Total of 0 image(s) found."
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Status          =   "Complete"
         StatusForeColor =   12582912
         BeginProperty StatusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusOffsetX   =   195
         StatusOffsetY   =   4
      End
      Begin prjImageBrowser.ctlThumbNailList ctlThumbNailList 
         Height          =   3015
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5318
         BackColor       =   16777215
         BackColorList   =   16777215
         BackColorThumbNail=   16777215
         BackColorFileName=   16761024
         BorderStyle     =   0
         BorderStylePicture=   0
         BeginProperty FontFileName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorThumbNail=   16777215
         BackColorFileName=   16761024
         BorderStylePicture=   0
         BeginProperty FontFileName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeightThumbNail =   113
         WidthThumbNail  =   105
         AutoVerbMenu    =   -1  'True
         BorderStyleFileName=   1
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up One Level"
            Object.ToolTipText     =   "Up One Level"
            ImageKey        =   "Up One Level"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19579
            Text            =   "Ready..."
            TextSave        =   "Ready..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   26
            TextSave        =   "2:21 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1B42
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1C54
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1D66
            Key             =   "Up One Level"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1E78
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1F8A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":209C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":21AE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":22C0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":23D2
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":24E4
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":25F6
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2708
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":281A
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":292C
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Image sptExplorerBrowser 
      Height          =   4815
      Left            =   2880
      MousePointer    =   9  'Size W E
      Picture         =   "frmBrowser.frx":2A3E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu MBARFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu MBAREdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyImage 
         Caption         =   "Copy image"
      End
      Begin VB.Menu mnuEditCopyThumbNail 
         Caption         =   "Copy thumbnail"
      End
      Begin VB.Menu MBAREdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnuEditSelectNone 
         Caption         =   "Select none"
      End
      Begin VB.Menu mnuEditSelectInverse 
         Caption         =   "Select inverse"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsSettings 
         Caption         =   "Settings"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlThumbNailList_AfterLoadFromPath(SearchPath As String, ImageTotal As Long)
    With prgImageLoading
        .Title = "Total of " & ImageTotal & " image(s) found."
        .Status = "Complete"
    End With
End Sub

Private Sub ctlThumbNailList_BeforeLoadFromPath(SearchPath As String, ImageTotal As Long)
    With prgImageLoading
        .Max = ImageTotal
        .Title = "Searching in " & SearchPath
    End With
End Sub

Private Sub ctlThumbNailList_ItemClick(Index As Integer)
    Set tmbMiniPreview.Picture = ctlThumbNailList.Item(Index).Picture
End Sub

Private Sub ctlThumbNailList_ItemDblClick(Index As Integer)
    ctlThumbNailList.SelectedItem.ViewImage
End Sub

Private Sub ctlThumbNailList_LoadFromPath(SearchPath As String, FileName As String, ImageIndex As Long, ImageLoaded As Long, ImageTotal As Long)
    With prgImageLoading
        .Value = ImageLoaded
        .Status = Int(ImageLoaded / ImageTotal * 100) & " %"
        
        DoEvents
    End With
End Sub

Private Sub ctlThumbNailList_RightClick()
    PopupMenu mnuEdit
End Sub

Private Sub Form_Load()
    'Load splitter bar position
    sptExplorerBrowser.Left = INIGetLong(Me.Name & "\ " & sptExplorerBrowser.Name, "Left", sptExplorerBrowser.Left)
    sptExplorerMiniPreview.Top = INIGetLong(Me.Name & "\ " & sptExplorerMiniPreview.Name, "Top", sptExplorerMiniPreview.Top)
    
    With ApplicationSetting.ExternalApplication
        ctlThumbNailList.ExternalViewer = .Viewer
        ctlThumbNailList.ExternalEditor = .Editor
        ctlThumbNailList.ExternalPrinter = .Printer
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save the splitter bar positions
    INISet Me.Name & "\ " & sptExplorerBrowser.Name, "Left", sptExplorerBrowser.Left
    INISet Me.Name & "\ " & sptExplorerMiniPreview.Name, "Top", sptExplorerMiniPreview.Top

    Dim FRM As Form
    For Each FRM In Forms
        Unload FRM
    Next
End Sub

Private Sub ftvExplorer_SelectionChange(Folder As CCRPFolderTV6.Folder, PreChange As Boolean, Cancel As Boolean)
    On Error GoTo ERROR_HANDLER_ftvExplorer_SelectionChange

    If Not PreChange Then
        ctlThumbNailList.LoadPicturesFromPath Folder.FullPath, "*.BMP;*.EMF;*.JPG;*.GIF;*.PCX"
        Me.Caption = CheckPath(Folder.FullPath, False) & " - Image Browser"
    End If

EXIT_ftvExplorer_SelectionChange:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_ftvExplorer_SelectionChange:
    Select Case Err.Number
    Case 53 'File not found!
    Case Else
        If MsgBox("Error in Sub ftvExplorer_SelectionChange() of Form frmBrowser[frmBrowser.frm] of Project prjImageBrowser[prjImageBrowser.vbp]" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "prjImageBrowser: Application error!") = vbNo Then Resume EXIT_ftvExplorer_SelectionChange
    End Select
    
    Resume Next
End Sub

Private Sub mnuEditCopyImage_Click()
    ctlThumbNailList.SelectedItem.CopyImageToClipBoard
End Sub

Private Sub mnuEditCopyThumbNail_Click()
    ctlThumbNailList.SelectedItem.CopyThumbNailToClipBoard
End Sub

Private Sub mnuEditSelectAll_Click()
    Dim ItemCounter As Integer
    For ItemCounter = 1 To ctlThumbNailList.ItemCount
        ctlThumbNailList.Item(ItemCounter).Checked = vbChecked
    Next
End Sub

Private Sub mnuEditSelectInverse_Click()
    Dim ItemCounter As Integer
    For ItemCounter = 1 To ctlThumbNailList.ItemCount
        If ctlThumbNailList.Item(ItemCounter).Checked = vbChecked Then ctlThumbNailList.Item(ItemCounter).Checked = vbUnchecked Else ctlThumbNailList.Item(ItemCounter).Checked = vbChecked
    Next
End Sub

Private Sub mnuEditSelectNone_Click()
    Dim ItemCounter As Integer
    For ItemCounter = 1 To ctlThumbNailList.ItemCount
        ctlThumbNailList.Item(ItemCounter).Checked = vbUnchecked
    Next
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileProperties_Click()
    ctlThumbNailList.SelectedItem.ShowImageProperties
End Sub

Private Sub mnuToolsSettings_Click()
    frmSetting.Show vbModal
End Sub

Private Sub picExplorer_Resize()
    On Error Resume Next

    sptExplorerMiniPreview.Move 0, sptExplorerMiniPreview.Top, picExplorer.ScaleWidth, sptExplorerMiniPreview.Height
    ftvExplorer.Move 0, 0, picExplorer.ScaleWidth, sptExplorerMiniPreview.Top
    tmbMiniPreview.Move 0, sptExplorerMiniPreview.Top + sptExplorerMiniPreview.Height, picExplorer.ScaleWidth, picExplorer.ScaleHeight - ftvExplorer.Height - sptExplorerMiniPreview.Height
    
End Sub

Private Sub sptExplorerMiniPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sptExplorerMiniPreview.Top = sptExplorerMiniPreview.Top + (Y / Screen.TwipsPerPixelX) - 2
        picExplorer_Resize
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Up One Level"
            'ToDo: Add 'Up One Level' button code.
            MsgBox "Add 'Up One Level' button code."
        Case "Find"
            'ToDo: Add 'Find' button code.
            MsgBox "Add 'Find' button code."
        Case "Print"
            ctlThumbNailList.SelectedItem.ShowPrint
        Case "Cut"
            'ToDo: Add 'Cut' button code.
            MsgBox "Add 'Cut' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Paste"
            'ToDo: Add 'Paste' button code.
            MsgBox "Add 'Paste' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
        Case "Properties"
            mnuFileProperties_Click
        Case "Help"
            'ToDo: Add 'Help' button code.
            MsgBox "Add 'Help' button code."
    End Select
End Sub


Private Sub sptExplorerBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sptExplorerBrowser.Left = sptExplorerBrowser.Left + (X / Screen.TwipsPerPixelX) - 2
        Form_Resize
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    sptExplorerBrowser.Move sptExplorerBrowser.Left, tbrMain.Height, sptExplorerBrowser.Width, Me.ScaleHeight - sbrMain.Height - tbrMain.Height
    picExplorer.Move 0, tbrMain.Height, sptExplorerBrowser.Left, Me.ScaleHeight - sbrMain.Height - tbrMain.Height
    picBrowser.Move sptExplorerBrowser.Left + sptExplorerBrowser.Width, tbrMain.Height, Me.ScaleWidth - (sptExplorerBrowser.Left + sptExplorerBrowser.Width), Me.ScaleHeight - sbrMain.Height - tbrMain.Height
End Sub

Private Sub picBrowser_Resize()
    prgImageLoading.Move 0, 0, picBrowser.ScaleWidth, prgImageLoading.Height
    prgImageLoading.StatusOffsetX = prgImageLoading.Width - 100
    
    ctlThumbNailList.Move 0, prgImageLoading.Height, picBrowser.ScaleWidth, picBrowser.ScaleHeight - prgImageLoading.Height
End Sub
