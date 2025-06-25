Attribute VB_Name = "Utils"
' Utils.bas  by  Robert Rayment

Option Explicit
Option Base 1

Public Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long

'Public Const SM_CXSCREEN = 0  ' Screen Width
'Public Const SM_CYSCREEN = 1  ' Screen Height
Public Const SM_CYCAPTION = 4 ' Height of window caption
Public Const SM_CYMENU = 15   ' Height of menu
Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)

Public ExtraBorder As Long
Public ExtraHeight  As Long

Public Xfra As Single
Public Yfra As Single

'#### GENERAL FRAME MOVER ####################################

Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, x As Single, y As Single)
Dim fraLeft As Long
Dim fraTop As Long

   If Button = vbLeftButton Then

      fraLeft = fra.Left + (x - Xfra) \ STX
      If fraLeft < 0 Then fraLeft = 0
      If fraLeft + fra.Width > frm.Width \ STX + fra.Width \ 2 Then
         fraLeft = frm.Width \ STX - fra.Width \ 2
      End If
      fra.Left = fraLeft

      fraTop = fra.Top + (y - Yfra) \ STY
      If fraTop < 8 Then fraTop = 8
      If fraTop + fra.Height > frm.Height \ STY + fra.Height \ 2 Then
         fraTop = frm.Height \ STY - fra.Height \ 2
      End If
      fra.Top = fraTop

   End If
End Sub
'#### END GENERAL FRAME MOVER ####################################

'#### POSITION SCROLL BARS AS picbox picP & piccontainer picC ##########

Public Sub FixScrollbars(picCr As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picCr = Picture Container
   ' picP  = Picture
   HS.Max = picP.Width - picCr.Width + 12   ' +4 to allow for border
   VS.Max = picP.Height - picCr.Height + 12 ' +4 to allow for border
   HS.LargeChange = picCr.Width \ 10
   HS.SmallChange = 1
   VS.LargeChange = picCr.Height \ 10
   VS.SmallChange = 1
   HS.Top = picCr.Top + picCr.Height + 1
   HS.Left = picCr.Left
   HS.Width = picCr.Width
   If picP.Width < picCr.Width Then
      HS.Visible = False
      HS.Enabled = False
   Else
      HS.Visible = True
      HS.Enabled = True
   End If
   VS.Top = picCr.Top
   VS.Left = picCr.Left - VS.Width - 1
   VS.Height = picCr.Height
   If picP.Height < picCr.Height Then
      VS.Visible = False
      VS.Enabled = False
   Else
      VS.Visible = True
      VS.Enabled = True
   End If
End Sub
'#### END POSITION SCROLL BARS AS picbox picP & piccontainer picC ##########

'#### FIX FILE EXTENSION #####################################
Public Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub
'#### END FIX FILE EXTENSION #####################################

Public Sub ExtractPath(Path$)
Dim p As Long
   If Right$(Path$, 1) = "\" Then Exit Sub
   p = InStrRev(Path$, "\")
   If p > 0 Then Path$ = Left$(Path$, p)
End Sub


Public Sub LngToRGB(LCul As Long)
'Public bred As Byte, bgreen As Byte, bblue As Byte
'Convert Long Colors() to RGB components
bred = (LCul And &HFF&)
bgreen = (LCul And &HFF00&) / &H100&
bblue = (LCul And &HFF0000) / &H10000
End Sub

Public Sub PaletteGrader(pRed() As Byte, pGreen() As Byte, pBlue() As Byte, pCul() As Long)

' Assumes:-
'pCul(0 To 255) As Long
'pRed(0 To 255) As Byte, pGreen(0 To 255) As Byte, pBlue(0 To 255) As Byte

ReDim RR(0 To 255) As Byte
ReDim GG(0 To 255) As Byte
ReDim BB(0 To 255) As Byte
'
ReDim RR1(0 To 255) As Byte
ReDim GG1(0 To 255) As Byte
ReDim BB1(0 To 255) As Byte

ReDim RR2(0 To 255) As Byte
ReDim GG2(0 To 255) As Byte
ReDim BB2(0 To 255) As Byte
'
ReDim RR3(0 To 255) As Byte
ReDim GG3(0 To 255) As Byte
ReDim BB3(0 To 255) As Byte
'
Dim NRR As Long
Dim NGG As Long
Dim NBB As Long
Dim NWW As Long
Dim ICC As Long
Dim ii As Long
Dim jj As Long

   ' Color sort
   ICC = 16 '20
   NRR = -1
   NGG = -1
   NBB = -1
   NWW = -1

   For ii = 0 To 255
      If pRed(ii) > pGreen(ii) Then
         If pRed(ii) - pGreen(ii) > ICC Then
            If pRed(ii) > pBlue(ii) Then
               If pRed(ii) - pBlue(ii) > ICC Then
                  NRR = NRR + 1
                  RR1(NRR) = pRed(ii)
                  GG1(NRR) = pGreen(ii)
                  BB1(NRR) = pBlue(ii)
                  GoTo nexi
               End If
             End If
         End If
      End If

      If pGreen(ii) > pBlue(ii) Then
         If pGreen(ii) - pBlue(ii) > ICC Then
            If pGreen(ii) > pRed(ii) Then
               If pGreen(ii) - pRed(ii) > ICC Then
                  NGG = NGG + 1
                  RR2(NGG) = pRed(ii)
                  GG2(NGG) = pGreen(ii)
                  BB2(NGG) = pBlue(ii)
                  GoTo nexi
               End If
             End If
         End If
      End If

      If pBlue(ii) > pRed(ii) Then
         If pBlue(ii) - pRed(ii) > ICC Then
            If pBlue(ii) > pGreen(ii) Then
               If pBlue(ii) - pGreen(ii) > ICC Then
                  NBB = NBB + 1
                  RR3(NBB) = pRed(ii)
                  GG3(NBB) = pGreen(ii)
                  BB3(NBB) = pBlue(ii)
                  GoTo nexi
               End If
             End If
         End If
      End If

      NWW = NWW + 1
      RR(NWW) = pRed(ii)
      GG(NWW) = pGreen(ii)
      BB(NWW) = pBlue(ii)

nexi:
   Next ii

''''''''''''''''''''''''''''''''''''''
   If NWW = -1 Then NWW = 0
   If NRR = -1 Then NRR = 0
   If NGG = -1 Then NGG = 0
   If NBB = -1 Then NBB = 0
   ' Redim Preserve for Quicksorting
   ReDim Preserve RR(0 To NWW), GG(0 To NWW), BB(0 To NWW)
   ReDim Preserve RR1(0 To NRR), GG1(0 To NRR), BB1(0 To NRR)
   ReDim Preserve RR2(0 To NGG), GG2(0 To NGG), BB2(0 To NGG)
   ReDim Preserve RR3(0 To NBB), GG3(0 To NBB), BB3(0 To NBB)
   ii = 0

   If NWW > 0 Then
      ReDim GreyCul(0 To NWW) As Long
      For jj = 0 To NWW
         GreyCul(jj) = RGB(RR(NWW), GG(NWW), BB(NWW))
      Next jj
      Quicksort GreyCul(), RR(), GG(), BB(), 1

      For jj = 0 To NWW
            pRed(ii) = RR(jj)
            pGreen(ii) = GG(jj)
            pBlue(ii) = BB(jj)
            ii = ii + 1
      Next jj
   End If

   If NRR > 0 Then
      Quicksort3 RR1(), GG1(), BB1()
      For jj = 0 To NRR
            pRed(ii) = RR1(jj)
            pGreen(ii) = GG1(jj)
            pBlue(ii) = BB1(jj)
            ii = ii + 1
      Next jj
   End If

   If NGG > 0 Then
      Quicksort3 GG2(), BB2(), RR2()
      For jj = 0 To NGG
            pRed(ii) = RR2(jj)
            pGreen(ii) = GG2(jj)
            pBlue(ii) = BB2(jj)
            ii = ii + 1
      Next jj
   End If

   If NBB > 0 Then
      Quicksort3 BB3(), RR3(), GG3()
      For jj = 0 To NBB
            pRed(ii) = RR3(jj)
            pGreen(ii) = GG3(jj)
            pBlue(ii) = BB3(jj)
            ii = ii + 1
            If ii > 255 Then Exit For
      Next jj
   End If

   ConvPalDataTo16Bit pRed(), pGreen(), pBlue()

''''''''''''''''''''''''''''''''''''''
   ' Needed for Browser
   For ii = 0 To 255
      pCul(ii) = RGB(pRed(ii), pGreen(ii), pBlue(ii))
   Next ii
Erase RR(), GG(), BB(), RR1(), GG1(), BB1()
Erase RR2(), GG2(), BB2(), RR3(), GG3(), BB3()
End Sub

Public Sub GreyedPalette()
Dim k As Long
ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   ' Greyed palette
   For k = 0 To 255
      palRed(k) = k
      palGreen(k) = k
      palBlue(k) = k
   Next k
End Sub

Public Sub BandedPalette()
Dim k As Long
ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   ' Banded palette
   For k = 0 To 39
      palRed(k) = 60 + 5 * k
   Next k
   For k = 40 To 79
      palGreen(k) = 60 + 5 * (k - 40)
   Next k
   For k = 80 To 119
      palBlue(k) = 60 + 5 * (k - 80)
   Next k
   For k = 120 To 159
      palGreen(k) = 60 + 5 * (k - 120)
      palBlue(k) = 60 + 5 * (k - 120)
   Next k
   
   For k = 160 To 199
      palRed(k) = 60 + 5 * (k - 160)
      palBlue(k) = 60 + 5 * (k - 160)
   Next k
   
   For k = 200 To 239
      palRed(k) = 60 + 5 * (k - 200)
      palGreen(k) = 60 + 5 * (k - 200)
      palBlue(k) = 5 * (k - 200)
   Next k
   
   For k = 240 To 255
      palRed(k) = 60 + 5 * (k - 240) + 100
      palGreen(k) = 60 + 5 * (k - 240) + 100
      palBlue(k) = 60 + 5 * (k - 240) + 100
   Next k
End Sub

Public Sub ShortBandedPalette()
Dim k As Long
Dim RR As Byte
Dim GG As Byte
Dim BB As Byte

ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   ' Banded palette
   RR = 248: GG = 0: BB = 0
   For k = 0 To 31
      palRed(k) = RR
      If RR > 7 Then RR = RR - 8
   Next k
   RR = 0: GG = 244: BB = 0
   For k = 32 To 63
      palGreen(k) = GG
      If GG > 7 Then GG = GG - 8
   Next k
   RR = 0: GG = 0: BB = 248
   For k = 64 To 95
      palBlue(k) = BB
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 255: GG = 255: BB = 0
   For k = 96 To 127
      palRed(k) = RR
      palGreen(k) = GG
      If RR > 7 Then RR = RR - 8
      If GG > 7 Then GG = GG - 8
   Next k
   RR = 248: GG = 0: BB = 248
   For k = 128 To 159
      palRed(k) = RR
      palBlue(k) = BB
      If RR > 7 Then RR = RR - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 0: GG = 240: BB = 240
   For k = 160 To 191
      palGreen(k) = GG
      palBlue(k) = BB
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 240: GG = 244: BB = 240
   For k = 192 To 223
      palRed(k) = GG
      palGreen(k) = GG
      palBlue(k) = BB
      If RR > 7 Then RR = RR - 8
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 248: GG = 248: BB = 248
   For k = 224 To 255
      palRed(k) = RR
      palGreen(k) = GG
      palBlue(k) = BB
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
End Sub

Public Sub CenteredPal()
' Method from Stefan Casier Paint256
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim zR As Single, zG As Single, zB As Single
Dim k As Long, k2 As Long, k3 As Long
Dim j As Long
ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
ReDim pCul(0 To 2, 0 To 15) As Byte
   pCul(0, 0) = 0:    pCul(1, 0) = 0:    pCul(2, 0) = 254
   pCul(0, 1) = 86:   pCul(1, 1) = 254:  pCul(2, 1) = 86
   pCul(0, 2) = 254:  pCul(1, 2) = 168:  pCul(2, 2) = 86
   pCul(0, 3) = 92:   pCul(1, 3) = 0:    pCul(2, 3) = 0
   pCul(0, 4) = 254:  pCul(1, 4) = 254:  pCul(2, 4) = 0
   pCul(0, 5) = 0:    pCul(1, 5) = 112:  pCul(2, 5) = 142
   pCul(0, 6) = 254:  pCul(1, 6) = 254:  pCul(2, 6) = 254
   pCul(0, 7) = 174:  pCul(1, 7) = 174:  pCul(2, 7) = 174
   pCul(0, 8) = 138:  pCul(1, 8) = 110:  pCul(2, 8) = 233
   pCul(0, 9) = 0:    pCul(1, 9) = 101:  pCul(2, 9) = 0
   pCul(0, 10) = 0:   pCul(1, 10) = 254: pCul(2, 10) = 254
   pCul(0, 11) = 254: pCul(1, 11) = 0:   pCul(2, 11) = 0
   pCul(0, 12) = 254: pCul(1, 12) = 254: pCul(2, 12) = 0
   pCul(0, 13) = 178: pCul(1, 13) = 0:   pCul(2, 13) = 178
   pCul(0, 14) = 254: pCul(1, 14) = 254: pCul(2, 14) = 254
   pCul(0, 15) = 90:  pCul(1, 15) = 90:  pCul(2, 15) = 90
   k3 = 0
   For k = 0 To 15
      k2 = k + 1
      If k = 15 Then k2 = 0
      zR = (1& * pCul(0, k) - pCul(0, k2)) / 16
      zG = (1& * pCul(1, k) - pCul(1, k2)) / 16
      zB = (1& * pCul(2, k) - pCul(2, k2)) / 16
      R2 = pCul(0, k)
      G2 = pCul(1, k)
      B2 = pCul(2, k)
      For j = 0 To 14
         R1 = R2 - zR
         G1 = G2 - zG
         B1 = B2 - zB
         If R1 < 0 Then R1 = 255
         If R1 > 255 Then R1 = 0
         R2 = R1
         If G1 < 0 Then G1 = 255
         If G1 > 255 Then G1 = 0
         G2 = G1
         If B1 < 0 Then B1 = 255
         If B1 > 255 Then B1 = 0
         B2 = B1
         palRed(k3) = R1
         palGreen(k3) = G1
         palBlue(k3) = B1
         k3 = k3 + 1
         If k3 > 255 Then Exit For
      Next j
      If k3 > 255 Then Exit For
   Next k
End Sub

Public Sub ConvPalDataTo16Bit(pR() As Byte, pG() As Byte, pB() As Byte)
'Public pR() As Byte, pG() As Byte, pB() As Byte
Dim remainder As Long
Dim i As Long
   ' RED   'Valid 16-bit values 0,16,24,32,,,248,255
   For i = 0 To 255
      remainder = pR(i) Mod 8
      If remainder <> 0 And pR(i) <> 255 Then
         pR(i) = pR(i) - remainder
      End If
      If pR(i) = 8 Then pR(i) = 0
   Next i
   ' GREEN  'Valid 16-bit values 0,8,12,16,20,,,,252,255
   For i = 0 To 255
      remainder = pG(i) Mod 4
      If remainder <> 0 And pG(i) <> 255 Then
         pG(i) = pG(i) - remainder
      End If
      If pG(i) = 4 Then pG(i) = 0
   Next i
   ' BLUE   'Valid 16-bit values 0,16,24,32,,,248,255
   For i = 0 To 255
      remainder = pB(i) Mod 8
      If remainder <> 0 And pB(i) <> 255 Then
         pB(i) = pB(i) - remainder
      End If
      If pB(i) = 8 Then pB(i) = 0
   Next i
End Sub

Public Sub Quicksort(marr() As Long, pRed() As Byte, pGreen() As Byte, pBlue() As Byte, Param As Long)
'1 dimensional long array sorted in ascending order from k to max
Dim Max As Long
Dim i As Long, j As Long, k As Long, LL As Long, mm As Long
Dim S As Long, ip As Long
Dim im As Long, IT As Long

   Max = UBound(marr)
   If Max = 1 Then Exit Sub
   k = LBound(marr)
   If k = Max Then Exit Sub
   ReDim sortl(Max \ 2) As Long, sortr(Max \ 2) As Long
   S = 1: sortl(1) = k: sortr(1) = Max
   Do While S <> 0
      LL = sortl(S): mm = sortr(S): S = S - 1
      
      Do While LL < mm
         i = LL: j = mm
         ip = (LL + mm) \ 2
         im = marr(ip)
         
         Do While i <= j
            Do While marr(i) < im: i = i + 1: Loop
            Do While im < marr(j): j = j - 1: Loop
            If i <= j Then
               'SWAP marr(i), marr(j)
               IT = marr(i): marr(i) = marr(j): marr(j) = IT
             
               If Param = 2 Then  ' Swap RGBs as well
                  IT = pRed(i): pRed(i) = pRed(j): pRed(j) = IT
                  IT = pGreen(i): pGreen(i) = pGreen(j): pGreen(j) = IT
                  IT = pBlue(i): pBlue(i) = pBlue(j): pBlue(j) = IT
               End If
             
               i = i + 1
               j = j - 1
            End If
         Loop
         If i < mm Then
            S = S + 1: sortl(S) = i: sortr(S) = mm
         End If
         mm = j
      Loop
   Loop
Erase sortl, sortr
End Sub

Public Sub Quicksort3(Arr1() As Byte, Arr2() As Byte, Arr3() As Byte)
'3 1 dimensional long arrays sorted in ascending order from k to max
'based on Arr1()
Dim Max As Long
Dim i As Long, j As Long, k As Long, LL As Long, mm As Long
Dim S As Long, ip As Long
Dim six$, siy$
Dim IT As Long
   Max = UBound(Arr1)
   If Max = 1 Then Exit Sub
   k = LBound(Arr1)
   If k = Max Then Exit Sub
   ReDim sortl(Max \ 2) As Long, sortr(Max \ 2) As Long
   S = 1: sortl(1) = k: sortr(1) = Max
   Do While S <> 0
      LL = sortl(S): mm = sortr(S): S = S - 1
      Do While LL < mm
         i = LL: j = mm
         ip = (LL + mm) \ 2
         six$ = Chr$(Arr1(ip)) + Chr$(Arr2(ip)) + Chr$(Arr3(ip))
         Do While i <= j
            Do
               siy$ = Chr$(Arr1(i)) + Chr$(Arr2(i)) + Chr$(Arr3(i))
               If siy$ >= six$ Then Exit Do
               i = i + 1
            Loop
            Do
               siy$ = Chr$(Arr1(j)) + Chr$(Arr2(j)) + Chr$(Arr3(j))
               If six$ >= siy$ Then Exit Do
               j = j - 1
            Loop
            If i <= j Then
               'SWAP Arr123(i), Arr123(j)
               IT = Arr1(i): Arr1(i) = Arr1(j): Arr1(j) = IT
               IT = Arr2(i): Arr2(i) = Arr2(j): Arr2(j) = IT
               IT = Arr3(i): Arr3(i) = Arr3(j): Arr3(j) = IT
               i = i + 1
               j = j - 1
            End If
         Loop
         If i < mm Then
            S = S + 1: sortl(S) = i: sortr(S) = mm
         End If
         mm = j
      Loop
   Loop
Erase sortl, sortr
End Sub

Public Function aFileExists(spath As String) As Boolean
    On Error GoTo NoFerr
    aFileExists = True
    If FileLen(spath) <> 0 Then Exit Function
NoFerr:
    aFileExists = False
End Function

Public Sub MouseKeys(KeyCode As Integer, Shift As Integer, Button As Integer, x As Single, y As Single)
Dim retval As Long
Dim lp As POINTAPI
Dim ix As Long, iy As Long
Dim imul As Long
Dim ixm As Long, iym As Long
Dim ixp As Long, iyp As Long
' +  107
' -  109
' =  187
' Shift = 16
' Ctrl  = 17

   Form1.PIC.SetFocus
   If aZoom And KeyCode = 17 Then  ' Ctrl
      ZOOMER
      Exit Sub
   End If
   
   lp.kX = 0
   lp.kY = 0
   retval = ClientToScreen(Form1.PIC.hwnd, lp)
   ix = lp.kX   'PIC postion from left of Screen
   iy = lp.kY   'PIC position from top of Screen
   ix = ix + xprev
   iy = iy + yprev
   
   imul = 1
   If Shift = 1 Then imul = 8
   
   If (xprev - imul) > 0 Then ixm = (xprev - imul) + lp.kX Else ixm = lp.kX
   If (yprev - imul) > 0 Then iym = (yprev - imul) + lp.kY Else iym = lp.kY
   If (xprev + imul) <= canvasW Then ixp = (xprev + imul) + lp.kX Else ixp = (canvasW - 1) + lp.kX
   If (yprev + imul) <= canvasH Then iyp = (yprev + imul) + lp.kY Else iyp = (canvasH - 1) + lp.kY
   
   Select Case KeyCode
   ' Arrows & Keypad
   Case 37, 100: SetCursorPos ixm, iy  ' LeftA, 4
   Case 38, 104: SetCursorPos ix, iym  ' UpA, 8
   Case 39, 102: SetCursorPos ixp, iy  ' RightA, 6
   Case 40, 98: SetCursorPos ix, iyp   ' DownA, 2
   
   ' Keypad only
   Case 36, 103 ' TL
         SetCursorPos ixm, iym   ' 7 XLeft, YUp
         x = ixm: y = iym
   Case 33, 105 ' TR
         SetCursorPos ixp, iym   ' 9 XRight, YUp
         x = ixp: y = iym
   Case 34, 99  ' BR
         SetCursorPos ixp, iyp   ' 3 XRight, YDown
         x = ixp: y = iyp
   Case 35, 97  ' BL
         SetCursorPos ixm, iyp   ' 1 XLeft, YDown
         x = ixm: y = iyp
   
   Case 13: Button = 1: SetCursorPos ix, iy  ' Enter
       x = ix: y = iy   'LC
   Case 8: Button = 2: SetCursorPos ix, iy  ' BackSpace or  -
       x = ix: y = iy   'RC
   
   ' F-keys ZOOM
   Case 113: ZoomFactor = 2   ' F2
     If aZoom Then
      frmZoom.optZoom(0).Value = True
      ZOOMER
     End If
   Case 115: ZoomFactor = 4   ' F4
     If aZoom Then
      frmZoom.optZoom(1).Value = True
      ZOOMER
     End If
   Case 117: ZoomFactor = 6   ' F6
     If aZoom Then
      frmZoom.optZoom(2).Value = True
      ZOOMER
     End If
   Case 119: ZoomFactor = 8   ' F8
     If aZoom Then
      frmZoom.optZoom(3).Value = True
      ZOOMER
     End If
   Case 123: ZoomFactor = 12  ' F12
     If aZoom Then
      frmZoom.optZoom(4).Value = True
      ZOOMER
     End If
   
   Case 16  ' Shift
   Case 17  ' Ctrl
   
   End Select
End Sub

Public Sub GetExtras(BStyle As Byte)

' IN:  BStyle = Me.BorderStyle
' OUT: Public ExtraBorder, ExtraHeight

''------------------------------------------------------------------------------
''  This required instead of Screen.Height & Width for resizing
'Public Declare Function GetSystemMetrics Lib "user32" _
'(ByVal nIndex As Long) As Long
'
'Public Const SM_CXSCREEN = 0  ' Screen Width
'Public Const SM_CYSCREEN = 1  ' Screen Height
'Public Const SM_CYCAPTION = 4 ' Height of window caption
'Public Const SM_CYMENU = 15   ' Height of menu
'Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
'Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)
'
''------------------------------------------------------------
Dim Border As Long
Dim CapHeight As Long
Dim MenuHeight As Long
' BStyle 1 to 5 (not 0)
' BStyle = Form1.BorderStyle

Border = GetSystemMetrics(SM_CXDLGFRAME)
If BStyle = 2 Or BStyle = 5 Then Border = Border + 1 ' Sizable
If BStyle > 3 Then
   CapHeight = GetSystemMetrics(SM_CYSMCAPTION) ' Small cap - ToolWindow
Else
   CapHeight = GetSystemMetrics(SM_CYCAPTION)   ' Standard cap
End If
ExtraBorder = 2 * Border
ExtraHeight = CapHeight + ExtraBorder

MenuHeight = GetSystemMetrics(SM_CYMENU)
ExtraHeight = CapHeight + MenuHeight + ExtraBorder

' Win98  ExtraBorder=6 or 8, ExtraHeight= 41 - 46
' WinXP  ExtraBorder=6 or 8, ExtraHeight= 44 - 54
End Sub

Public Sub GetSetUpInfo()
Dim a$
Dim k As Long
Dim N As Long
Dim OOpenPathSpec$
Dim OSavePathSpec$
'Public AppPathSpec$, OpenPathSpec$, SavePathSpec$
On Error GoTo InfoTxtError
   AppPathSpec$ = App.Path
   If Right$(AppPathSpec$, 1) <> "\" Then AppPathSpec$ = AppPathSpec$ & "\"

   'Or pick up from PaintInfo.txt file
   OpenPathSpec$ = AppPathSpec$
   SavePathSpec$ = AppPathSpec$
   OOpenPathSpec$ = AppPathSpec$
   OSavePathSpec$ = AppPathSpec$
   
   DefaultTools
   
   ' Hard name = "PaintInfo.txt"
   If aFileExists(AppPathSpec$ & "PaintInfo.txt") Then
      Open AppPathSpec$ & "PaintInfo.txt" For Input As #1
      Line Input #1, OpenPathSpec$
      Line Input #1, SavePathSpec$
      Input #1, BrushType
      Input #1, SprayType
      Input #1, LineType
      Input #1, PolyLineType
      Input #1, CurvyLineType
      Input #1, RectangleType
      Input #1, CirllipseType
      Input #1, ConeType
      Input #1, TubeType
      Input #1, BulletType
      Input #1, JunctionType
      Input #1, ArcType
      Input #1, ShapeType
      Input #1, RadialType
      Input #1, FillType
      Input #1, TreeType
      Input #1, ArrowType
      
      Input #1, frmBrowseLeft
      Input #1, frmBrowseTop
      Input #1, frmZoomLeft
      Input #1, frmZoomTop
      Input #1, frmToolOptionsLeft
      Input #1, frmToolOptionsTop
      Input #1, frmTextLeft
      Input #1, frmTextTop
      Input #1, frmPaletteLeft
      Input #1, frmPaletteTop
      Input #1, frmStripLeft
      Input #1, frmStripTop
      Input #1, frmHelpLeft
      Input #1, frmHelpTop
      Input #1, frmTransformLeft
      Input #1, frmTransformTop
      Input #1, frmViewsLeft
      Input #1, frmViewsTop
      Input #1, frmCanvasLeft
      Input #1, frmCanvasTop
      
      Line Input #1, a$ ' JASC-PAL
      Line Input #1, a$ ' 0100
      Input #1, N       ' 256
      ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
      For k = 0 To N - 1
         Input #1, palRed(k), palGreen(k), palBlue(k)
      Next k
      Close #1
      ReDim DefaultRGB(0 To 255)
      ReDim CulRGB(0 To 255)
      ReDim CulBGR(0 To 255)
      ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
      palRed(0) = 0
      palGreen(0) = 0
      palBlue(0) = 0
      palRed(1) = 255
      palGreen(1) = 255
      palBlue(1) = 255
      ReDim CulRGB(0 To 255), CulBGR(0 To 255)
      For k = 0 To 255
         CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
         CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
      Next k
      DefaultRGB() = CulRGB()
      
      If ArcType = 0 Then ArcType = 1
      
      If OpenPathSpec$ = "" Then OpenPathSpec$ = AppPathSpec$
'      If UCase$(Left$(OpenPathSpec$, 1)) <> UCase$(Left$(AppPathSpec$, 1)) Then
'         OpenPathSpec$ = AppPathSpec$
'      Else
         ExtractPath OpenPathSpec$
'      End If
      
      If SavePathSpec$ = "" Then SavePathSpec$ = AppPathSpec$
'      If UCase$(Left$(SavePathSpec$, 1)) <> UCase$(Left$(AppPathSpec$, 1)) Then
'         SavePathSpec$ = AppPathSpec$
'      Else
         ExtractPath SavePathSpec$
'      End If
   
   Else
      DefaultTools
      SetfrmDefaultPositions
      SetDefaultPalette
      PrintPaintInfoTxt
   End If
   '---------------------------------------
   On Error GoTo OpenPathError
   ChDir OpenPathSpec$
   On Error GoTo SavePathError
   ChDir SavePathSpec$
   ChDir AppPathSpec$
   On Error GoTo 0
   Exit Sub
'============
OpenPathError:
OpenPathSpec$ = AppPathSpec$
Resume Next
'============
SavePathError:
SavePathSpec$ = AppPathSpec$
Resume Next
'============
InfoTxtError:  ' Wrong entires in "PaintInfo.txt"
Close
DefaultTools
SetfrmDefaultPositions
SetDefaultPalette
SavePathSpec$ = AppPathSpec$
PrintPaintInfoTxt
On Error GoTo 0
End Sub

Public Sub PrintPaintInfoTxt()
Dim k As Long
   Open AppPathSpec$ & "PaintInfo.txt" For Output As #1
   Print #1, OpenPathSpec$
   Print #1, SavePathSpec$
   Print #1, BrushType
   Print #1, SprayType
   Print #1, LineType
   Print #1, PolyLineType
   Print #1, CurvyLineType
   Print #1, RectangleType
   Print #1, CirllipseType
   Print #1, ConeType
   Print #1, TubeType
   Print #1, BulletType
   Print #1, JunctionType
   Print #1, ArcType
   Print #1, ShapeType
   Print #1, RadialType
   Print #1, FillType
   Print #1, TreeType
   Print #1, ArrowType
   
   Print #1, frmBrowseLeft
   Print #1, frmBrowseTop
   Print #1, frmZoomLeft
   Print #1, frmZoomTop
   Print #1, frmToolOptionsLeft
   Print #1, frmToolOptionsTop
   Print #1, frmTextLeft
   Print #1, frmTextTop
   Print #1, frmPaletteLeft
   Print #1, frmPaletteTop
   Print #1, frmStripLeft
   Print #1, frmStripTop
   Print #1, frmHelpLeft
   Print #1, frmHelpTop
   Print #1, frmTransformLeft
   Print #1, frmTransformTop
   Print #1, frmViewsLeft
   Print #1, frmViewsTop
   Print #1, frmCanvasLeft
   Print #1, frmCanvasTop
   
   Print #1, "JASC-PAL"
   Print #1, "0100"
   Print #1, UBound(DefaultRGB(), 1) + 1
   For k = 0 To UBound(DefaultRGB(), 1)
      palRed(k) = (DefaultRGB(k) And &HFF&)
      palGreen(k) = (DefaultRGB(k) And &HFF00&) / &H100&
      palBlue(k) = (DefaultRGB(k) And &HFF0000) / &H10000
      Print #1, Trim$(Str$(palRed(k))); " "; Trim$(Str$(palGreen(k))); " "; Trim$(Str$(palBlue(k)))
   Next k
   Print #1,
   Close #1
End Sub

Public Sub DefaultTools()
   BrushType = Dot1
   SprayType = Dots1
   LineType = SingleLine1
   PolyLineType = PolySingleLine1
   CurvyLineType = CurvySingleLine1
   RectangleType = RectangleSingle1
   CirllipseType = CirllipseSingle1
   ConeType = ConeOutline
   TubeType = TubeOutLine
   BulletType = BulletOutLine
   JunctionType = TPiece1
   ArcType = ArcTL
   ShapeType = TShape1
   RadialType = RSpokes
   FillType = Fill1
   TreeType = Tree1
   ArrowType = ArrSingle
End Sub

Public Sub SetfrmDefaultPositions()
   frmBrowseLeft = Screen.Width \ 3
   frmBrowseTop = Screen.Height \ 6
   frmZoomLeft = 2 * Screen.Width \ 3
   frmZoomTop = Screen.Height \ 6
   frmToolOptionsLeft = Screen.Width \ 3
   frmToolOptionsTop = Screen.Height \ 6
   frmTextLeft = 2 * Screen.Width \ 3
   frmTextTop = Screen.Height \ 6
   frmPaletteLeft = Screen.Width \ 3
   frmPaletteTop = Screen.Height \ 6
   frmStripLeft = Screen.Width \ 2
   frmStripTop = Screen.Height \ 6
   frmHelpLeft = 2 * Screen.Width \ 3
   frmHelpTop = Screen.Height \ 6
   frmTransformLeft = Screen.Width \ 3
   frmTransformTop = Screen.Height \ 6
   frmViewsLeft = Screen.Width \ 3
   frmViewsTop = Screen.Height \ 6
   frmCanvasLeft = Screen.Width \ 3
   frmCanvasTop = Screen.Height \ 3
End Sub

Public Sub SetDefaultPalette()
Dim k As Long
ReDim DefaultRGB(0 To 255)
ReDim CulRGB(0 To 255)
ReDim CulBGR(0 To 255)
   'ShortBandedPalette
   CenteredPal
   ConvPalDataTo16Bit palRed(), palGreen(), palBlue()
   palRed(0) = 0
   palGreen(0) = 0
   palBlue(0) = 0
   palRed(1) = 255
   palGreen(1) = 255
   palBlue(1) = 255
   ReDim CulRGB(0 To 255), CulBGR(0 To 255)
   For k = 0 To 255
      CulRGB(k) = RGB(palRed(k), palGreen(k), palBlue(k))
      CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
   Next k
   DefaultRGB() = CulRGB()
End Sub

Public Sub InitDefaultPalette()
Dim k As Long
   ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   For k = 0 To UBound(DefaultRGB(), 1)
      palRed(k) = (DefaultRGB(k) And &HFF&)
      palGreen(k) = (DefaultRGB(k) And &HFF00&) / &H100&
      palBlue(k) = (DefaultRGB(k) And &HFF0000) / &H10000
   Next k
End Sub

Public Function zATan2(ByVal zY As Single, ByVal zx As Single) As Single
' Public pi# = Const
' 0 degrees to right
' Find angle arctan from -pi# to +pi#
   If zx <> 0 Then
      zATan2 = Atn(zY / zx)
      If (zx < 0) Then  ' Want +pi# when zY=0
         If (zY < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
      End If
   Else  ' zx=0
      zATan2 = Sgn(zY) * pi# / 2
   End If
End Function

Public Sub SetUpFractalTrees()
   ReDim BushSize(3)
   ReDim zAngP(3), zAngN(3)
   ReDim xstep(3), ystep(3)
   ReDim xmul(3), ymul(3)
   ReDim Axiom$(3)
   ReDim PAxiom$(3)
   
   BushSize(1) = 1
   Axiom$(1) = "FG"
   PAxiom$(1) = "FG+[+F-F-F]-[-G+F+F]"
   zAngP(1) = 12
   zAngN(1) = -12
   xstep(1) = 0
   ystep(1) = -4
   xmul(1) = 1.01
   ymul(1) = 1.05

   BushSize(2) = 1
   Axiom$(2) = "F"
   PAxiom$(2) = "FF+[+F-F-F]-[-F+F+F]"
   zAngP(2) = 25
   zAngN(2) = -25
   xstep(2) = 0
   ystep(2) = -5
   xmul(2) = 1
   ymul(2) = 1

   BushSize(3) = 1
   Axiom$(3) = "F"
   PAxiom$(3) = "FF-[-F+F+F]+[+F-F-F+FFF]"
   zAngP(3) = 30
   zAngN(3) = -30
   xstep(3) = 0
   ystep(3) = -2
   xmul(3) = 1
   ymul(3) = 1
End Sub

Public Sub ZOOMER()
' Public px1,py1,ZoomFactor
Dim sx As Long
Dim sy As Long
Dim dx As Long
Dim dy As Long
Dim offset As Long
Dim iz As Long
Dim ipx As Long
Dim ipy As Long
Dim shapeleft As Long
Dim shapetop As Long
Dim bS As BITMAPINFO
   ReDim ZARR(canvasW, canvasH)
   GETLONGS Form1.PIC.Image, ZARR(), canvasW, canvasH
   iz = ZoomSize * ZoomSize
   ReDim ZoomArr(ZoomSize, ZoomSize)
   FillMemory ZoomArr(1, 1), 4 * iz, 208 ' Grey edge
   iz = ZoomFactor
   offset = ZoomSize \ (2 * iz)
   For sy = py1 - (offset - 1) To py1 + offset
      If sy > 0 Then
      If sy <= canvasH Then
         dy = (offset - 1 + sy - py1) * iz + 1
         For sx = px1 - (offset - 1) To px1 + offset
            If sx > 0 Then
            If sx <= canvasW Then
               dx = (offset - 1 + sx - px1) * iz + 1
               For ipy = 0 To iz - 1
               For ipx = 0 To iz - 1
                  If dx + ipx > 0 Then
                  If dx + ipx <= ZoomSize Then
                  If dy + ipy > 0 Then
                  If dy + ipy <= ZoomSize Then
                        ZoomArr(dx + ipx, dy + ipy) = ZARR(sx, sy)
                        If sx = px1 Then
                        If sy = py1 Then
                              If ipx = 0 Then
                              If ipy = iz - 1 Then
                                 shapeleft = dx
                                 shapetop = ZoomSize - dy
                              End If
                              End If
                        End If
                        End If
                  End If
                  End If
                  End If
                  End If
               Next ipx
               Next ipy
            End If
            End If
         Next sx
      End If
      End If
   Next sy
   
   ' Set up palette
   CopyMemory bS.Colors(0), CulBGR(0), 1024
   With bS.bmi
      .biSize = 40
      .biwidth = ZoomSize
      .biheight = ZoomSize
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 4 * canvasW * ZoomSize
   End With
   'frmZoom.picZoom.Cls
   SetDIBitsToDevice frmZoom.picZoom.hDC, 0, 0, ZoomSize, ZoomSize, _
   0, 0, 0, ZoomSize, ZoomArr(1, 1), bS, DIB_RGB_COLORS
   ' Locate & Size Shape marker
   frmZoom.ShapeC.Left = shapeleft - 1
   frmZoom.ShapeC.Top = shapetop - iz + 1
   frmZoom.ShapeC.Width = iz
   frmZoom.ShapeC.Height = iz
   frmZoom.picZoom.Refresh
End Sub

'#### ROLLERS & SHIFTERS ################################

Public Sub RollLeft(NS As Long, RS As Long)
Dim N As Long
Dim BB As Byte
   For N = 1 To NS
      For iy = 1 To canvasH
         BB = bArray(1, iy)
         CopyMemory bArray(1, iy), bArray(2, iy), canvasW - 1
         If RS = 0 Then
            bArray(canvasW, iy) = BB
         Else
            bArray(canvasW, iy) = 0
         End If
      Next iy
   Next N
End Sub

Public Sub RollRight(NS As Long, RS As Long)
Dim N As Long
Dim BB As Byte
   ReDim AA(1 To canvasW - 1) As Byte
   For N = 1 To NS
      For iy = 1 To canvasH
         BB = bArray(canvasW, iy)
         CopyMemory AA(1), bArray(1, iy), canvasW - 1
         CopyMemory bArray(2, iy), AA(1), canvasW - 1
         If RS = 0 Then
            bArray(1, iy) = BB
         Else
            bArray(1, iy) = 0
         End If
      Next iy
   Next N
   Erase AA()
End Sub

Public Sub RollUp(NS As Long, RS As Long)
Dim N As Long
   ReDim BB(1 To canvasW) As Byte
   For N = 1 To NS
      If RS = 0 Then
         CopyMemory BB(1), bArray(1, canvasH), canvasW
      End If
      For iy = canvasH To 2 Step -1
         CopyMemory bArray(1, iy), bArray(1, iy - 1), canvasW
      Next iy
      CopyMemory bArray(1, 1), BB(1), canvasW
   Next N
   Erase BB()
End Sub

Public Sub RollDown(NS As Long, RS As Long)
Dim N As Long
   ReDim BB(1 To canvasW) As Byte
   For N = 1 To NS
      If RS = 0 Then
         CopyMemory BB(1), bArray(1, 1), canvasW
      End If
      For iy = 2 To canvasH
         CopyMemory bArray(1, iy - 1), bArray(1, iy), canvasW
      Next iy
      CopyMemory bArray(1, canvasH), BB(1), canvasW
   Next N
   Erase BB()
End Sub

'Public Sub Loadmcode(InFile$, MCCode() As Byte)
''Load machine code into InCode() byte array
'Dim MCSize As Long
'On Error GoTo InFileErr
'If Dir$(InFile$) = "" Then
'   MsgBox (InFile$ & " missing")
'   DoEvents
'   End
'End If
'Open InFile$ For Binary As #1
'MCSize& = LOF(1)
'If MCSize& = 0 Then
'InFileErr:
'   MsgBox (InFile$ & " missing")
'   DoEvents
'   End
'End If
'ReDim MCCode(1 To MCSize)
'Get #1, , MCCode
'Close #1
'On Error GoTo 0
'End Sub


Public Sub FlushMouseEvents(ByVal hwnd As Long)
Dim msg_info As MSG

    Do While PeekMessage(msg_info, hwnd, WM_MOUSEFIRST, _
        WM_MOUSELAST, PM_REMOVE) <> 0
            ' Fetch messages until there are no more.
    Loop
End Sub



