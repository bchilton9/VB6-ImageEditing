VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CFDialog  FontDlg.cls   Font Dialog
' Adapted from vbAccelerator.com

Option Explicit

'Private Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Sub CopyMemoryStr Lib "Kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hwndOwner As Long           ' Caller's window handle
    hDC As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    Flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" _
    Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000&  ' Must also have ScreenFonts And PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000&
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Private m_lApiReturn As Long
Private m_lExtendedError As Long

' Font wrapper
Function VBChooseFont(Curfont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional Flags As Long = 0) As Boolean

    m_lApiReturn = 0
    m_lExtendedError = 0

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If Flags And CF_PrinterFonts Then PrinterDC = Printer.hDC
    Else
        Flags = Flags Or CF_PrinterFonts
    End If
    ' Must have some fonts
    If (Flags And CF_PrinterFonts) = 0 Then Flags = Flags Or CF_ScreenFonts
    ' Color can take initial color, receive chosen color
    'If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    'flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then Flags = Flags Or CF_LimitSize
    ' MaxSize can be maximum size accepted
    If MaxSize Then Flags = Flags Or CF_LimitSize

    ' Put in required internal flags and remove unsupported
    Flags = (Flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(Curfont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = Curfont.Weight
    fnt.lfItalic = Curfont.Italic
    fnt.lfUnderline = Curfont.Underline
    fnt.lfStrikeOut = Curfont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, Curfont.Name

    ' Initialize TCHOOSEFONT variable
    Dim CF As TCHOOSEFONT
    CF.lStructSize = Len(CF)
    If owner <> -1 Then CF.hwndOwner = owner
    CF.hDC = PrinterDC
    CF.lpLogFont = VarPtr(fnt)
    CF.iPointSize = Curfont.Size * 10
    CF.Flags = Flags
    CF.rgbColors = Color
    CF.nSizeMin = MinSize
    CF.nSizeMax = MaxSize
    
    ' All other fields zero
    m_lApiReturn = ChooseFont(CF)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseFont = True
        Flags = CF.Flags
        Color = CF.rgbColors
        Curfont.Bold = CF.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType And Italic_FontType
        Curfont.Italic = fnt.lfItalic
        Curfont.Strikethrough = fnt.lfStrikeOut
        Curfont.Underline = fnt.lfUnderline
        Curfont.Weight = fnt.lfWeight
        Curfont.Size = CF.iPointSize / 10
        Curfont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        VBChooseFont = False
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseFont = False
    End Select
        
End Function

Private Sub StrToBytes(ab() As Byte, S As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(S, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(S) < cab Then S = S & String$(cab - Len(S), 0)
        'If UnicodeTypeLib Then
        '    Dim st As String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), S, cab
        'End If
    End If
End Sub

Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function





