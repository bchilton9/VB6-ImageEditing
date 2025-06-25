VERSION 5.00
Begin VB.Form frmText 
   Caption         =   " GET TEXT"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   315
      Left            =   3150
      TabIndex        =   6
      Top             =   1395
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelText 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3150
      TabIndex        =   5
      Top             =   465
      Width           =   930
   End
   Begin VB.CommandButton cmdAcceptText 
      Caption         =   "Accept"
      Height          =   300
      Left            =   3150
      TabIndex        =   4
      Top             =   105
      Width           =   930
   End
   Begin VB.HScrollBar HSAngle 
      Height          =   240
      LargeChange     =   10
      Left            =   3075
      Max             =   180
      Min             =   -180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   135
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   1980
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   1605
      Left            =   105
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   90
      Width           =   2715
   End
   Begin VB.Label LabSlant 
      Height          =   180
      Left            =   135
      TabIndex        =   9
      Top             =   1725
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "o"
      Height          =   165
      Left            =   3855
      TabIndex        =   8
      Top             =   780
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "NB. Not all fonts can be rotated"
      Height          =   210
      Left            =   1830
      TabIndex        =   7
      Top             =   1755
      Width           =   2340
   End
   Begin VB.Label LabAngle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3405
      TabIndex        =   3
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmText.frm

Option Explicit
Option Base 1

'------------------------------------------------------------------------------

'API's & Structure for Rotating Text
'Logical Font
Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
'        lfFaceName(LF_FACESIZE - 1) As Byte
        lfFaceName As String * LF_FACESIZE
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
(lpLogFont As LOGFONT) As Long
Private RotateFont As LOGFONT

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long
'----------------------------------------------------------------

' Public Type BITMAPINFOHEADER
' Public Sub GetPICBytes
' Public NSTOREXY As Long
' Public STOREX() As Long, STOREY() As Long
' Public frmTextLeft As Long
' Public frmTextTop As Long
' Public STX As Long, STY As Long
' Public Const pi# = 3.1415927
' Public CulNum, CulRGB()

' Public Type FontStuff
'   FontName As String
'   FontSize As Long
'   FontItalic As Boolean
'   FontBold As Boolean
'End Type
'Public SVFont As FontStuff

Dim zangtext As Single
Dim W As Long, H As Long
Dim Slant As Long
Dim bT() As Byte
Dim aTErr As Boolean
Dim Curfont As StdFont

Private Sub Form_Load()
   Me.Left = frmTextLeft
   Me.Top = frmTextTop
   Me.Caption = " GET TEXT "
   TextColor = CulRGB(CulNum)
   With picText
      .ForeColor = TextColor
      .FontName = SVFont.FontName
      .FontSize = SVFont.FontSize
      .FontItalic = SVFont.FontItalic
      .FontBold = SVFont.FontBold
   End With
   With Text1
      .FontName = SVFont.FontName
      .FontSize = SVFont.FontSize
      .FontItalic = SVFont.FontItalic
      .FontBold = SVFont.FontBold
   End With
   
   Text1.Text = vbNullString
   zangtext = 0
   HSAngle.Value = zangtext
   LabAngle = zangtext
   
   Set Curfont = New StdFont
   With Curfont
      .Name = SVFont.FontName
      .Size = SVFont.FontSize
      .Italic = SVFont.FontItalic
      .Bold = SVFont.FontBold
   End With
   
   picText.BackColor = CulRGB(0)
   picText.ForeColor = TextColor
End Sub

Private Sub cmdCancelText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Set Curfont = Nothing
   NSTOREXY = 0
   TextLine$ = vbNullChar
   frmTextLeft = Me.Left
   frmTextTop = Me.Top
   Unload Me
End Sub

Private Sub ShowText(aTextErr As Boolean)
Dim xp As Single, yp As Single
Dim zang As Single
Dim zIB As Single

   aTextErr = False
   TextLine$ = Text1.Text
   If Len(TextLine$) = 0 Then
      picText.Picture = LoadPicture
      Exit Sub
   End If
   With picText
      .Picture = LoadPicture
      .Cls
      .ForeColor = TextColor
      If Curfont.Size < 8 Then
         .FontSize = Curfont.Size
         .FontName = Curfont.Name
         .FontItalic = Curfont.Italic
         .FontBold = Curfont.Bold
      Else
         .FontName = Curfont.Name
         .FontSize = Curfont.Size
         .FontItalic = Curfont.Italic
         .FontBold = Curfont.Bold
      End If
   End With
   zIB = 1
   If Curfont.Italic And Curfont.Bold Then
      zIB = 1.2
   ElseIf Curfont.Italic Or Curfont.Bold Then
      zIB = 1.1
   End If
   W = CLng(zIB * Len(TextLine$) * picText.TextWidth("W"))
   H = picText.TextHeight("_")
   
   Slant = Sqr(H ^ 2 + W ^ 2)
   Slant = (Slant + 3) And &HFFFFFFFC
   LabSlant = Str$(Slant) & " ( Max = 1024)"
   ' Ensure rotated string within the slant*slant picbox
   zang = zangtext * pi# / 180
   xp = Slant / 2 - W / 2 * Cos(zang) + H / 2 * Sin(zang)
   yp = Slant / 2 - H / 2 * Cos(zang) - W / 2 * Sin(zang)
   If Slant > 1024 Then
      MsgBox "TOO BIG!!", vbInformation, "Text"
      TextLine$ = Left$(TextLine$, 1)
      Text1.Text = TextLine$
      aTextErr = True
      Exit Sub
   End If
   With picText
      .Width = Slant
      .Height = Slant
      .CurrentX = xp
      .CurrentY = yp
   End With
   RotateText picText, zangtext, TextLine$
End Sub

Private Sub cmdAcceptText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Get pixels from
Dim ix As Long, iy As Long
Dim ixoff As Long, iyoff As Long
Dim BacNum As Long
   If Slant = 0 Then Exit Sub
   If Len(TextLine$) = 0 Then
      Set Curfont = Nothing
      NSTOREXY = 0
      TextLine$ = vbNullChar
      frmTextLeft = Me.Left
      frmTextTop = Me.Top
      Unload Me
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bT(Slant, Slant)
   
   ' Public
   GetPICBytes picText.Image, bT(), Slant, Slant
   
   BacNum = bT(1, 1)
   NSTOREXY = 0
   ReDim STOREX(1), STOREY(1)
   For iy = 1 To Slant
   For ix = 1 To Slant
      If bT(ix, iy) <> BacNum Then
         NSTOREXY = NSTOREXY + 1
         ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
         If NSTOREXY = 1 Then
            ixoff = ix
            iyoff = iy
         End If
         STOREX(NSTOREXY) = ix - ixoff
         STOREY(NSTOREXY) = canvasH - (iy - iyoff) - 1
      End If
   Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
   Erase bT()
   picText.Picture = LoadPicture
   picText.Width = 4
   picText.Height = 4
   With SVFont
      .FontName = Curfont.Name
      .FontSize = Curfont.Size
      .FontItalic = Curfont.Italic
      .FontBold = Curfont.Bold
   End With

   Set Curfont = Nothing
   frmTextLeft = Me.Left
   frmTextTop = Me.Top
   Unload Me
End Sub

Private Sub cmdFont_Click()
Dim cc As FDialog
   Set cc = New FDialog
   ' Initial Font
   With Curfont
      If SVFont.FontSize < 8 Then
         .Size = SVFont.FontSize
         .Name = SVFont.FontName
         .Italic = SVFont.FontItalic
         .Bold = SVFont.FontBold
         .Weight = 1
         .Strikethrough = False
         .Underline = False
      Else
         .Name = SVFont.FontName
         .Size = SVFont.FontSize
         .Italic = SVFont.FontItalic
         .Bold = SVFont.FontBold
         .Weight = 1
         .Strikethrough = False
         .Underline = False
      End If
   End With
   
   If cc.VBChooseFont(Curfont, , Me.hWnd) Then
      
      With picText
         .Picture = LoadPicture
         .Cls
         .ForeColor = TextColor
         If Curfont.Size < 8 Then
            .FontSize = Curfont.Size
            .FontName = Curfont.Name
            .FontItalic = Curfont.Italic
            .FontBold = Curfont.Bold
         Else
            .FontName = Curfont.Name
            .FontSize = Curfont.Size
            .FontItalic = Curfont.Italic
            .FontBold = Curfont.Bold
         End If
      End With
         
      With Text1
         If Curfont.Size < 8 Then
            .FontSize = Curfont.Size
            .FontName = Curfont.Name
            .FontItalic = Curfont.Italic
            .FontBold = Curfont.Bold
            '.ForeColor = TextColor
         Else
            .FontName = Curfont.Name
            .FontSize = Curfont.Size
            .FontItalic = Curfont.Italic
            .FontBold = Curfont.Bold
         End If
      End With
         
      With SVFont
         .FontName = Curfont.Name
         .FontSize = Curfont.Size
         .FontItalic = Curfont.Italic
         .FontBold = Curfont.Bold
      End With
   
   End If
   Set cc = Nothing
   
   If TextLine$ <> "" Then
      ShowText aTErr
      If aTErr Then ShowText aTErr
      Text1.SetFocus
   Else
      TextLine$ = ""
      Text1.SetFocus
      Text1.Text = ""
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Erase bT()
   picText.Picture = LoadPicture
   picText.Width = 4
   picText.Height = 4
   Set Curfont = Nothing
   frmTextLeft = Me.Left
   frmTextTop = Me.Top
   Unload frmText
End Sub

Private Sub HSAngle_Change()
   zangtext = HSAngle.Value
   LabAngle = zangtext
   ShowText aTErr
   If aTErr Then ShowText aTErr
End Sub

Private Sub Text1_Change()
   ShowText aTErr
   If aTErr Then ShowText aTErr
End Sub


Private Sub RotateText(PIC As PictureBox, zangle As Single, TextLine$)
'NB NOT ALL FONTS CAN BE ROTATED
'eg MS Sans Serif !!
'PIC CurrentX & Y set
'Set rotation in tenths of a degree, i.e., 1800 = 180 degrees
'TextLine$ = Text2.Text
Dim rfont As Long
Dim Curfont As Long
   'Make +ve angles rotate clockwise
   RotateFont.lfEscapement = -zangle * 10
   'RotateFont.lfWidth = 10
   With RotateFont
      If PIC.FontBold Then .lfWeight = 1200 / 2 Else .lfWeight = 400 / 2
      If PIC.FontItalic Then .lfItalic = 1 Else .lfItalic = 0
      .lfStrikeOut = 0
      .lfUnderline = 0
      .lfCharSet = 0   '0,1
      .lfFaceName = PIC.FontName & Chr$(0)
      .lfHeight = (PIC.FontSize * -20) / STY
   End With
   '------------------------------------
   rfont = CreateFontIndirect(RotateFont)
   Curfont = SelectObject(PIC.hDC, rfont)
   PIC.Print TextLine$;
   PIC.Refresh
   'Restore CurFont
   SelectObject PIC.hDC, Curfont
   DeleteObject rfont
   '------------------------------------
End Sub



