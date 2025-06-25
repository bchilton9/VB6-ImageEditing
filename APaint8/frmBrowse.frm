VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Browse"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10230
   ControlBox      =   0   'False
   HelpContextID   =   30
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProg 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   7215
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   21
      Top             =   5535
      Width           =   2940
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   390
      Left            =   6345
      TabIndex        =   20
      Top             =   5550
      Width           =   690
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5550
      Width           =   825
   End
   Begin VB.CommandButton cmdCancelSave 
      Caption         =   "Cancel save"
      Height          =   345
      Left            =   7335
      TabIndex        =   17
      Top             =   5055
      Width           =   1170
   End
   Begin VB.CommandButton cmdAcceptName 
      Caption         =   "Accept Name"
      Height          =   360
      Left            =   7335
      TabIndex        =   16
      Top             =   4650
      Width           =   1170
   End
   Begin VB.CommandButton cmdSaveGIF 
      Caption         =   "Save As GIF"
      Height          =   375
      Left            =   7335
      TabIndex        =   14
      Top             =   4230
      Width           =   1170
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Save As BMP"
      Height          =   375
      Left            =   7335
      TabIndex        =   13
      Top             =   3810
      Width           =   1170
   End
   Begin VB.PictureBox PICIN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   10395
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   12
      Top             =   165
      Width           =   360
   End
   Begin VB.Frame fraPal 
      Caption         =   " 256 Palette "
      Height          =   4875
      Left            =   8670
      TabIndex        =   8
      Top             =   570
      Width           =   1440
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort PAL && Remap"
         Height          =   495
         Left            =   180
         TabIndex        =   15
         Top             =   4245
         Width           =   1095
      End
      Begin VB.PictureBox picPalette 
         AutoRedraw      =   -1  'True
         Height          =   3900
         Left            =   195
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         Top             =   255
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdVIEW 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5415
      TabIndex        =   7
      Top             =   5550
      Width           =   660
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2655
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5490
      Width           =   2445
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   9195
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   840
   End
   Begin VB.CommandButton cmdUse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Picture         =   "frmBrowse.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5595
      Width           =   600
   End
   Begin VB.DirListBox Dir1 
      Height          =   5040
      Left            =   45
      TabIndex        =   2
      Top             =   375
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   2550
   End
   Begin VB.FileListBox File1 
      Height          =   5355
      Left            =   2640
      Pattern         =   "*.bmp;*.gif;*.jpg"
      TabIndex        =   0
      Top             =   60
      Width           =   2460
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   5235
      ScaleHeight     =   83
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   6
      Top             =   3765
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2415
      TabIndex        =   19
      Top             =   5595
      Width           =   210
   End
   Begin VB.Label Label2 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   5145
      TabIndex        =   11
      Top             =   5580
      Width           =   270
   End
   Begin VB.Shape Shape1 
      Height          =   3330
      Left            =   5160
      Top             =   75
      Width           =   3360
   End
   Begin VB.Label LabScale 
      Caption         =   "Scaled image"
      Height          =   180
      Left            =   5220
      TabIndex        =   10
      Top             =   3450
      Width           =   1290
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   3240
      Left            =   5205
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmBrowse.frm  by  Robert Rayment

Option Explicit
Option Base 1

'---------------------------------------------------
' For sending files to Recycle bin
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Byte) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_NOCONFIRMATION = &H10
'---------------------------------------------------
Private Declare Function SetStretchBltMode Lib "gdi32" _
   (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Const COLORONCOLOR = 3
Const HALFTONE = 4

Private Type PALETTEENTRY
    peR     As Byte
    peG     As Byte
    peB     As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(0 To 255) As PALETTEENTRY
End Type
Private logpal256 As LOGPALETTE256

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)
'--------------------------------------------------------------------------
'FileSpec$            ' Public
'OpenPathSpec$        ' Public
'SavePathSpec$        ' Public

Private SaveFileSpec$

' Image view
Private iW As Long
Private iH As Long
Private iWMax As Long
Private iHMax As Long
Private zAspect As Single

Private W As Long
Private H As Long
Private W4 As Long
Private H4 As Long

' File & Palette stuff
Private FPath$, FName$
Private PrevFName$
Private Dr$
Private Ext$
Private bpp As Integer
Private pCulRGB() As Long
Private pCulBGR() As Long
Private pRed() As Byte, pGreen() As Byte, pBlue() As Byte
Private iiW As Integer       ' Picture1 width
Private iiH As Integer       ' Picture1 height

Private NumColors As Long     ' Num of palette entries in BMP & GIF
Dim ibpp As Integer
Private BYTEGIF As Byte       ' Packed color byte

Private aBMPGIF As Boolean    ' True Save As BMP  False Save As GIF
Private aSaveMode As Boolean
Private FSize As Long

' Conversion & remapping
Private b8() As Byte
Private b32() As Byte
Private pCulRGBCopy() As Long
Private NEnt As Long ' Number of optimal palette entries
' General
Private fnum As Long
Private BB As Byte
Private k As Long
Private ix As Long
Private iy As Long
Private pdot As Long
Private a$


Private Sub Form_Load()
Dim p As Long
Dim ThePath$
   On Error GoTo PERR:
   
   frmBrowse.Top = frmBrowseTop
   frmBrowse.Left = frmBrowseLeft
   
   PathSpec$ = OpenPathSpec$
   
   Drive1.Drive = Left$(PathSpec$, 2)
   Dr$ = Drive1.Drive
   p = InStrRev(PathSpec$, "\")
   If p <> 0 Then
      ThePath$ = Left$(PathSpec$, p)
   Else
      ThePath$ = App.Path
      If Right$(ThePath$, 1) <> "\" Then ThePath$ = ThePath$ & "\"
   End If
   
   Dir1.Path = ThePath$
   File1.Path = ThePath$
   
   Caption = "  View(*.bmp,*.gif,*.jpg,*.pal), Convert, Save 8bpp(*.bmp, *.gif)"
   
   File1.Pattern = "*.bmp;*.gif;*.jpg"
   
   Dir1_Change
   Drive1_Change
   
   Caption = "  View(*.bmp,*.gif,*.jpg), Convert, Save 8bpp(*.bmp, *.gif)"
   
   ' Image preview max size
   iWMax = 216
   iHMax = 216
   
   Text1.Text = ""
   Text1.BackColor = vbWhite
   
   LabScale.Visible = False
   FalseAll
   aSaveMode = False
   
   ReDim b8(1)
   ReDim b32(1)

   ' Progress
   picProg.DrawWidth = 2
   picProg.Cls
   On Error GoTo 0

Exit Sub
'=============
PERR:
PathSpec$ = App.Path
Resume
End Sub

Private Sub cmdVIEW_Click()
'Public CulRGB() As Long
'Public CulBGR() As Long
'Public palRed() As Byte, palGreen() As Byte, palBlue() As Byte

   On Error GoTo FERR
   
   Screen.MousePointer = vbHourglass
   FalseAll
   
   'FName$ = File1.FileName 'Text1.Text(image)
   FName$ = Text1.Text
   
   If Len(FName$) <> 0 Then
      pdot = InStrRev(FName$, ".")
      Ext$ = UCase$(Mid$(FName$, pdot + 1)) ' Extension
      FPath$ = File1.Path
      If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
      FileSpec$ = FPath$ & FName$
   Else
      FileSpec$ = ""
      Ext$ = ""
      TrueAll
      Screen.MousePointer = vbDefault
      DoEvents
      Exit Sub
   End If
      
   ' Needed for ALL pictures
   PICIN.Picture = LoadPicture
   PICIN.Refresh
   PICIN.Picture = LoadPicture(FileSpec$)
   PICIN.Refresh
   W = PICIN.Width
   H = PICIN.Height
   
   ' Test if too big
   If W > MAXWIDTH Or H > MAXHEIGHT Then
      picInfo.Cls
      If W > MAXWIDTH Then
         picInfo.Print " Width=" & Str$(W) & "  >" & Str$(MAXHEIGHT)
      End If
      If H > MAXHEIGHT Then
         picInfo.Print " Height=" & Str$(W) & "  >" & Str$(MAXHEIGHT)
      End If
      
      MsgBox " TOO BIG !" & vbCr & " Max WxH =" & Str$(MAXWIDTH) & " x" & Str$(MAXHEIGHT), vbCritical, "Loading file"
      TrueAll
      Screen.MousePointer = vbDefault
      DoEvents
      Exit Sub
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Image preview max size
   ' Display image (scaled down if necessary)
   imgPreview.Picture = LoadPicture()
   ' Show width & height
   zAspect = W / H
   If W <= iWMax Then
      iW = W
      iH = CLng(iW / zAspect)
   Else  ' Picture1.Width > iWMax
      iW = iWMax
      iH = CLng(iWMax / zAspect)
   End If
   If iH > iHMax Then
      iH = iHMax
      iW = CLng(iH * zAspect)
   End If
   imgPreview.Width = iW
   imgPreview.Height = iH
   imgPreview.Picture = PICIN.Picture  'LoadPicture(FileSpec$)
   imgPreview.Refresh
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Show W & H
   picInfo.Cls
   picInfo.Print " Width=" & Str$(W)
   picInfo.Print " Height=" & Str$(H)
   picInfo.Refresh
   
   W4 = (W + 3) And &HFFFFFFFC
   H4 = (H + 3) And &HFFFFFFFC
      
   ibpp = 0
   Ext$ = LCase$(Right$(FileSpec$, 3))
   ibpp = 777
   
   Select Case Ext$
   Case "bmp"
      fnum = FreeFile
      Open FileSpec$ For Binary As fnum
      FSize = LOF(fnum)
      Seek #fnum, 29
      Get #fnum, , ibpp
      Close
      NumColors = 2 ^ ibpp
   Case "gif"
      fnum = FreeFile
      Open FileSpec$ For Binary As fnum
      a$ = Space$(3)
      Get #1, , a$
      Seek #1, 7
      Get #1, , iiW
      Get #1, , iiH
      Get #1, , BYTEGIF
      Close
      NumColors = 2 ^ ((BYTEGIF And &H7) + 1)    'eg 196 = 11000100  100=4 +1 = 5 2^5=32
   Case Else
      NumColors = 0
   End Select
   
   If Ext$ = "bmp" And ibpp <= 8 _
         And (FSize - 54 - 4 * NumColors) = W * H _
         And W4 = W And H4 = H _
         And NumColors = 256 Then
         
      GetBMPPalette     ' Open file doesn't close
      ' Get b8() indexes from file
      ReDim b8(W, H)
      Get #1, , b8()
      Close
      
   ElseIf Ext$ = "gif" And W4 = W And H4 = H _
          And BYTEGIF >= 128 Then   ' Global palette should follow
            GetGIFPalette  ' Open & close file
            Getb8IndexesFromPICIN
   Else  'BMP or GIF <> mult of 4 or jpg
         ' Progress
         picProg.Cls
         picProg.Print "     CONVERTING"
         picProg.Refresh
         
         ' Make W & H mult of 4
         W4 = (W + 3) And &HFFFFFFFC
         H4 = (H + 3) And &HFFFFFFFC
         '---------------------------------------------------------------
         If W <> W4 Or H <> H4 Then    ' Resize to multiples of 4
            ReDim b32(4, W, H)
            ReDim pCulRGB(0 To 255)
            GETBYTES PICIN.Picture, b32(), W, H, 32, pCulRGB(), 0
            SetStretchBltMode PICIN.hDC, HALFTONE
            PICIN.Width = W4
            PICIN.Height = H4
            SetStretchBltMode PICIN.hDC, HALFTONE
            Dim SS As BITMAPINFO
            With SS.bmi
               .biSize = 40
               .biwidth = W
               .biheight = H
               .biPlanes = 1
               .biBitCount = 32    ' Sets up 32-bit colors
               .biCompression = 0
               .biSizeImage = W * H
            End With
            PICIN.Picture = LoadPicture
            StretchDIBits PICIN.hDC, 0&, 0&, W4, H4, 0&, 0&, W, H, b32(1, 1, 1), SS, DIB_PAL_COLORS, vbSrcCopy
            PICIN.Refresh
         End If
         '---------------------------------------------------------------
         W = W4   ' SET ANY NEW W & H
         H = H4
         ' Show W & H (new)
         picInfo.Cls
         picInfo.Print " Width=" & Str$(W)
         picInfo.Print " Height=" & Str$(H)
         picInfo.Refresh
         
         If Ext$ = "bmp" And ibpp <= 8 And NumColors = 256 Then
            GetBMPPalette      ' Opens file doesn't close
            Close
            Getb8IndexesFromPICIN
         ElseIf Ext$ = "gif" And BYTEGIF >= 128 Then   ' Global palette should follow
            GetGIFPalette  ' Opens & closes file
            Getb8IndexesFromPICIN
         Else  ' 24bpp BMP or jpg ie no palette
            ReDim b32(4, W, H)
            GETBYTES PICIN.Image, b32(), W, H, 32, pCulRGB(), 0
            ' 4x4 image now setup in b32()
            ReDim b8(W, H)   ' To get b8() & partial palette
            ReDim pRed(0 To 255), pGreen(0 To 255), pBlue(0 To 255)
            ReDim pCulRGB(0 To 255)
            CreateOptimal b32(), pRed(), pGreen(), pBlue(), pCulRGB(), b8(), True, NEnt
            Erase b32()
         End If
   End If

   ' Set full BGR palette as well
   ReDim pCulBGR(0 To 255)
   For k = 0 To 255
      pCulBGR(k) = RGB(pBlue(k), pGreen(k), pRed(k))
   Next k
   ShowPalette

   TrueAll
   ' Progress
   picProg.Cls
   picProg.Print "            DONE"
   picProg.Refresh
   
Exit Sub
'==============
FERR:
picInfo.Cls
imgPreview.Picture = LoadPicture()
Text1.Text = ""
FalseAll
Screen.MousePointer = vbDefault
Close
Beep
MsgBox "File error - " & FileSpec$, vbExclamation, "Opening file"
On Error GoTo 0
End Sub

Private Sub GetBMPPalette()
   ReDim pRed(0 To 255), pGreen(0 To 255), pBlue(0 To 255)
   ReDim pCulRGB(0 To 255)
   ' Extract bmp palette
   fnum = FreeFile
   Open FileSpec$ For Binary As fnum
   Seek #fnum, 55
   For k = 0 To NumColors - 1
      Get #1, , pBlue(k)
      Get #1, , pGreen(k)
      Get #1, , pRed(k)
      Get #1, , BB
      pCulRGB(k) = pRed(k) + 256& * pGreen(k) + 65536 * pBlue(k)
   Next k
End Sub

Private Sub GetGIFPalette()
   ReDim pRed(0 To 255), pGreen(0 To 255), pBlue(0 To 255)
   ReDim pCulRGB(0 To 255)
   ' Extract gif palette
   fnum = FreeFile
   Open FileSpec$ For Binary As fnum
   Seek #1, 14
   For k = 0 To NumColors - 1
         Get #1, , pRed(k)
         Get #1, , pGreen(k)
         Get #1, , pBlue(k)
         pCulRGB(k) = pRed(k) + 256& * pGreen(k) + 65536 * pBlue(k)
   Next k
   Close
End Sub

Private Sub Getb8IndexesFromPICIN()
   ReDim b8(W, H)   ' To get b8() & partial palette
   ptStanPal = VarPtr(pCulRGB(0))    ' Standard
   For iy = 1 To H
   For ix = 1 To W
      LongDerived = PICIN.Point(ix - 1, iy - 1)
      k = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      b8(ix, H - iy + 1) = k
   Next ix
   Next iy
End Sub

Private Sub cmdDelete_Click()

'From http://www.vbapi.com/ref/s/shfileoperation.html

Dim fos As SHFILEOPSTRUCT  ' structure to pass to the function
Dim sa(1 To 32) As Byte    ' byte array to make structure properly sized
Dim res As Long         ' return value
Dim KillSpec$
   
   FPath$ = File1.Path
   If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
   KillSpec$ = FPath$ & Text1.Text
    
   If aFileExists(KillSpec$) Then
   
      
      res = MsgBox("   Yes   -   Permanent" & vbCr _
                & "    No   -   To Recycle bin" & vbCr _
                & "          or Cancel", vbYesNoCancel, _
                "Delete file permanently")
      If res = vbYes Then
         Kill KillSpec$
         Text1.Text = ""
      ElseIf res = vbNo Then
         With fos
             .hwnd = Me.hwnd
             .wFunc = FO_DELETE
             .pFrom = KillSpec$ & vbNullChar & vbNullChar
             .pTo = vbNullChar & vbNullChar
             .fFlags = FOF_ALLOWUNDO Or FOF_FILESONLY Or FOF_NOCONFIRMATION
             .fAnyOperationsAborted = 0
             .hNameMappings = 0
             .lpszProgressTitle = vbNullChar
         End With
         ' Necessary sometimes for byte alignment?
         CopyMemory sa(1), fos, LenB(fos)
         CopyMemory sa(19), sa(21), 12
         res = SHFileOperation(sa(1))
         Text1.Text = ""
      End If
   
   End If
   File1.Refresh
End Sub

Private Sub cmdSaveBMP_Click()
   PrevFName$ = FName$
   Text1.BackColor = vbYellow
   If Len(Text1.Text) = 0 Then
      MsgBox "Enter a file name", vbInformation, "Saving file"
      Exit Sub
   End If
   FPath$ = File1.Path
   If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
   SaveFileSpec$ = FPath$ & Text1.Text
   
   FixFileExtension SaveFileSpec$, "bmp"
   
   pdot = InStrRev(SaveFileSpec$, "\")
   If pdot > 0 Then
     FName$ = Mid$(SaveFileSpec$, pdot + 1)
     Text1.Text = FName$
   Else
     MsgBox "Save FileSpec error", vbCritical, "Saving file"
     Exit Sub
   End If
   
   cmdAcceptName.Enabled = True
   cmdCancelSave.Enabled = True
   aBMPGIF = True
   aSaveMode = True
   cmdVIEW.Enabled = False
   cmdUse.Enabled = False
End Sub

Private Sub cmdSaveGIF_Click()
   PrevFName$ = FName$
   Text1.BackColor = vbYellow
   If Len(Text1.Text) = 0 Then
      MsgBox "Enter a file name", vbInformation, "Saving file"
      Exit Sub
   End If
   FPath$ = File1.Path
   If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
   SaveFileSpec$ = FPath$ & Text1.Text
   
   FixFileExtension SaveFileSpec$, "gif"
   
   pdot = InStrRev(SaveFileSpec$, "\")
   If pdot > 0 Then
     FName$ = Mid$(SaveFileSpec$, pdot + 1)
     Text1.Text = FName$
   Else
     MsgBox "Save FileSpec error", vbCritical, "Saving file"
     Exit Sub
   End If
   
   cmdAcceptName.Enabled = True
   cmdCancelSave.Enabled = True
   aBMPGIF = False
   aSaveMode = True
   cmdVIEW.Enabled = False
   cmdUse.Enabled = False
End Sub

Private Sub cmdAcceptName_Click()
Dim resp As Long
   FPath$ = File1.Path
   If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
   SaveFileSpec$ = FPath$ & Text1.Text
   If aBMPGIF Then
      FixFileExtension SaveFileSpec$, "bmp"
   Else
      FixFileExtension SaveFileSpec$, "gif"
   End If

   pdot = InStrRev(SaveFileSpec$, "\")
   If pdot > 0 Then
     FName$ = Mid$(SaveFileSpec$, pdot + 1)
     Text1.Text = FName$
   Else
     MsgBox "Save FileSpec error", vbCritical, "Saving file"
     Exit Sub
   End If

   If aFileExists(SaveFileSpec$) Then
      resp = MsgBox("File exists, overwrite?", vbQuestion + vbYesNo, "Saving file")
      If resp = vbNo Then Exit Sub
   End If

   If aBMPGIF Then   'Save BMP
      MSaveBMP SaveFileSpec$, b8(), W, H, pCulBGR()
   Else  ' Save GIF
      MSaveGIF SaveFileSpec$, b8(), CInt(W), CInt(H), pCulRGB(), True
   End If

   PrevFName$ = FName$
   File1.Refresh
   Text1.BackColor = vbWhite
   cmdAcceptName.Enabled = False
   cmdCancelSave.Enabled = False
   aSaveMode = False
   cmdDelete.Enabled = True
   cmdVIEW.Enabled = True
   cmdUse.Enabled = True
   FileSpec$ = SaveFileSpec
   SavePathSpec$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))

End Sub

Private Sub cmdCancelSave_Click()
   FName$ = PrevFName$
   Text1.Text = FName$
   File1.Refresh
   Text1.BackColor = vbWhite
   cmdAcceptName.Enabled = False
   cmdCancelSave.Enabled = False
   aSaveMode = False
   cmdDelete.Enabled = True
   cmdVIEW.Enabled = True
   cmdUse.Enabled = True
End Sub

Private Sub cmdSort_Click()
Dim hPal As Long
Dim LP256 As LOGPALETTE256
Dim Culr As Long
Dim pIndex As Long
Dim bIndex() As Byte
Dim k As Long
   ' Progress
   picProg.Cls
   picProg.Print "        SORTING"
   picProg.Refresh
   
   ReDim bIndex(0 To 255)

   
   Screen.MousePointer = vbHourglass
   FalseAll

   ReDim pCulRGBCopy(0 To 255)
   pCulRGBCopy() = pCulRGB()
'   If NumColors < 128 Then
'      For k = 2 To 255
'         pRed(k) = pRed(k - 2)
'         pGreen(k) = pGreen(k - 2)
'         pBlue(k) = pBlue(k - 2)
'      Next k
'   End If
      
      PaletteGrader pRed(), pGreen(), pBlue(), pCulRGBCopy()
   
   ' Changes pRed,G,B & pCulRGBCopy()
   ' NB If only a few colors are in original palette then
   '    this will put colors to end leaving black at the
   '    start of the palette:-

   
   If pCulRGBCopy(0) = 0 And pCulRGBCopy(1) = 0 Then
      ' Make 1 white
      pRed(1) = 255
      pGreen(1) = 255
      pBlue(1) = 255
      pCulRGBCopy(1) = vbWhite
   ElseIf pCulRGBCopy(0) <> 0 Or pCulRGBCopy(1) <> vbWhite Then
      ' Shift palette up by 2, leaving 0 & 1 for black & white
      Dim bbdummy(0 To 255) As Byte
      Dim dummyL(0 To 255) As Long
      CopyMemory bbdummy(2), pRed(0), 254
      pRed() = bbdummy()
      pRed(0) = 0
      pRed(1) = 255
      CopyMemory bbdummy(2), pGreen(0), 254
      pGreen() = bbdummy()
      pGreen(0) = 0
      pGreen(1) = 255
      CopyMemory bbdummy(2), pBlue(0), 254
      pBlue() = bbdummy()
      pBlue(0) = 0
      pBlue(1) = 255

      CopyMemory dummyL(2), pCulRGBCopy(0), (1024 - 8)
      pCulRGBCopy() = dummyL()
      pCulRGBCopy(0) = 0
      pCulRGBCopy(1) = vbWhite
End If
'''''''''''''''''''''''''''''''''''''''''''
' Very slow in IDE, unless ASM, BUT faster than GetNearestPaletteIndex
' when compiled!!
Dim MinD As Long
Dim pR As Long
Dim pG As Long
Dim pB As Long
Dim LongVal As Long

   For iy = 1 To H
   For ix = 1 To W
      Culr = pCulRGB(b8(ix, iy))
      ' Get Index
      ptStanPal = VarPtr(pCulRGBCopy(0))    ' Standard
      pIndex = CallWindowProc(ptMC, Culr, ptStanPal, 3&, 4&)
      b8(ix, iy) = pIndex
   Next ix
   Next iy
   
   ' Copy new palettes
   pCulRGB() = pCulRGBCopy()
   For k = 0 To 255
         pCulBGR(k) = RGB(pBlue(k), pGreen(k), pRed(k))
   Next k
   Erase pCulRGBCopy()

   ShowPalette  ' pCulRGB(k)
   
   TrueAll

   Erase bbdummy(), dummyL()
   
   ' Progress
   picProg.Cls
   picProg.Print "            DONE"
   picProg.Refresh
End Sub

Private Sub Dir1_Click()
   FileSpec$ = ""
   FalseAll
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   FileSpec$ = ""
   FalseAll
End Sub

Private Sub Drive1_Change()
   On Error GoTo NoDrive
   Dir1.Path = Drive1.Drive
   Dr$ = Drive1.Drive
   FileSpec$ = ""
   FalseAll
Exit Sub
'==========
NoDrive:
Beep
Drive1.Drive = Dr$
Resume
End Sub

Private Sub File1_DblClick()
   If Not aSaveMode Then
      pdot = InStrRev(File1.FileName, ".")
      Ext$ = UCase$(Mid$(File1.FileName, pdot + 1)) ' Extension
      Text1.Text = File1.FileName
      Text1.BackColor = vbWhite
      cmdVIEW_Click
   End If
End Sub

Private Sub File1_Click()
   pdot = InStrRev(File1.FileName, ".")
   Ext$ = UCase$(Mid$(File1.FileName, pdot + 1)) ' Extension
   Text1.Text = File1.FileName
   If Not aSaveMode Then
      Text1.BackColor = vbWhite
   Else
      Text1.BackColor = vbYellow
   End If
   cmdVIEW.Enabled = True
   cmdDelete.Enabled = True
End Sub

Private Sub cmdUse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Have b8(), palette pCulRGB() & pCulBGR(), & (W & H) mod 4
Dim k As Long
' Return Public FileSpec$ to calling program

   FName$ = Text1.Text

   If Len(FName$) <> 0 Then
      FPath$ = File1.Path
      If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
      FileSpec$ = FPath$ & FName$
      OpenPathSpec$ = FileSpec$
   Else
      FileSpec$ = ""
   End If

   ' Reduce picbox
   PICIN.Picture = LoadPicture
   PICIN.Width = 8
   PICIN.Height = 8
   
   ' TRANSFER TO PUBLICS ----------------
   'Public CulRGB() As Long
   'Public CulBGR() As Long
   'Public palRed() As Byte, palGreen() As Byte, palBlue() As Byte
   ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   For k = 0 To 255
      palRed(k) = pRed(k)
      palGreen(k) = pGreen(k)
      palBlue(k) = pBlue(k)
   Next k
   CulRGB() = pCulRGB()
   CulBGR() = pCulBGR()
   ReDim bArray(W, H)
   bArray() = b8()
   canvasW = W
   canvasH = H
   '-------------------------------------
   
   Erase b8(), b32()
   
   frmBrowseTop = frmBrowse.Top
   frmBrowseLeft = frmBrowse.Left

   Unload frmBrowse
End Sub

Private Sub cmdClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   FileSpec$ = ""
   ' Reduce picbox
   PICIN.Picture = LoadPicture
   PICIN.Width = 8
   PICIN.Height = 8

   Erase b8(), b32()

   frmBrowseTop = frmBrowse.Top
   frmBrowseLeft = frmBrowse.Left

   Unload frmBrowse
End Sub

Private Sub FixFileExtension(FSpec$, Ext$)
   Dim Exten$
   Dim E$
   Exten$ = "." + LCase$(Ext$)
   
   pdot = InStrRev(FSpec$, ".")
   
   If pdot = 0 Then
      FSpec$ = FSpec$ + Exten$
   Else
      E$ = LCase$(Mid$(FSpec$, pdot))
      If E$ <> Exten$ Then
         FSpec$ = Left$(FSpec$, pdot - 1) + Exten$
      End If
   End If
End Sub

Private Sub ShowPalette()
   For k = 0 To 255
    ix = 1 + (k Mod 8) * 8
    iy = 1 + (k \ 8) * 8
    picPalette.Line (ix, iy)-(ix + 5, iy + 5), pCulRGB(k), BF  ' pCulRGB() is RGBA
   Next k
   picPalette.Refresh
End Sub

Private Sub TrueAll()
   Screen.MousePointer = vbDefault
   cmdUse.Enabled = True
   cmdSaveBMP.Enabled = True
   cmdSaveGIF.Enabled = True
   cmdSort.Enabled = True
   cmdVIEW.Enabled = True
   cmdDelete.Enabled = True
   DoEvents
End Sub

Private Sub FalseAll()
   cmdUse.Enabled = False
   cmdSaveBMP.Enabled = False
   cmdSaveGIF.Enabled = False
   cmdAcceptName.Enabled = False
   cmdCancelSave.Enabled = False
   cmdSort.Enabled = False
   cmdVIEW.Enabled = False
   cmdDelete.Enabled = False
   DoEvents
End Sub

