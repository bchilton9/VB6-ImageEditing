Attribute VB_Name = "Filters"
' Filters.bas by Robert Rayment

Option Explicit
Option Base 1
'
' Filter formulae:
'  some adapted from Manuel Santos, Malcolm Ferris & Johannes B

' Parameters
' Set at frmTransform
' Filters
Public PCONTOUR As Long       ' Threshold 32 -> 224
Public PDITHER As Long        ' 16 ->  48 'Floyd-Steinberg B
Public PENGRAVEMBOSS As Long  ' -3 -> +3
Public PPOSTERIZE As Long     ' Threshold 32 -> 224
Public PRELIEF As Long        ' 1,2,3
Public PSMOOTH As Long        ' 1,2,3,4
Public PSHADE As Long         ' -100 -> +100
Public PMELT As Long          ' 2 -> 8
Public POIL As Long           ' 1 -> 5
Public PSHARPEN As Long       ' 1,2,3
Public PLITHO As Long         ' 92 -> 100
Public PCONTRAST              ' -66 -> +66
Public PDIFFUSE As Long       ' 1  ->  16
Public PBLACKWHITE As Long    ' 1 -> 255
Public PSOLAR As Long         ' 32 -> 224
Public PFOG As Long           ' -100 -> +100
Public PSQUARE As Long         ' 1 to 16

' Deformers
Public zPELLIPSE As Single    ' 0.005 -> 1
Public PFLUTE As Long         ' -50 -> +50 (exc -2 to +2)
Public PRIPPLE As Long        ' 0 -> 40
Public zPROUNDRECT As Single  ' 0 -> .5
Public PTILE As Long          ' 1 -> 21
Public PMLENS As Long         ' 2 -> 40
Public zPLENS As Single       ' 0.1 -> 4
Public PFWINDOW As Long       ' 4 -> 100
Public zPSWIRL As Single      ' -50 -> +50
Public zPMINMAG As Single     ' .01 -> 2 not 0
Public zPROTATE As Single     ' -180 -> +180
Public PKALI As Long          ' 1 -> 21

' Adders
Public PLINES As Long         ' 0 -> 40
Public zPTHICKLINE As Single  ' .0 -> 1.0
Public PSPOKES As Long        ' 1 -> 201
Public PDNET As Long          ' 1 -> 41

Public WWLO As Long, HHLO As Long
Public WWHI As Long, HHHI As Long

Dim SR As Long
Dim SG As Long
Dim SB As Long
Dim pR As Long
Dim pG As Long
Dim pB As Long
Dim Cul As Long
Dim Factor As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim MinD As Long
Dim LongVal As Long
Dim Intens() As Long
Dim zA As Single
Dim zB As Single
Dim zD As Single
Dim ixd As Long
Dim iyd As Long
Dim zix As Single
Dim ziy As Single
Dim istep As Long
Dim HH As Long
Dim WW As Long
Dim ii As Long
Dim jj As Long
Dim Hist() As Long
Dim IndexK() As Long
Dim TX As Long
Dim TY As Long

Dim CH As Long
Dim CW As Long
Dim aUseSelectcnSV As Boolean

Dim Speed(0 To 765) As Long

Public Sub ZeroSelRect()
   bDummy() = bArray()
   For iy = HHLO + 1 To HHHI - 1
   For ix = WWLO + 1 To WWHI - 1
      bDummy(ix, iy) = 0
   Next ix
   Next iy
End Sub

'#### FILTERS ###############################################

Public Sub Contour()
'Public PCONTOUR As Long       ' Threshold 32 -> 224
'1 1 1
'1 0 1
'1 1 1   8*(i,j)-Sum
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO + 1 To HHHI - 1
   For ix = WWLO + 1 To WWHI - 1
      SR = 0: SG = 0: SB = 0
      For j = iy - 1 To iy + 1
      For i = ix - 1 To ix + 1
         'If Not (j = iy And i = ix) Then
         If (j <> iy Or i <> ix) Then
            k = bArray(i, j)
            SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         End If
      Next i
      Next j
      k = bArray(ix, iy)
      SR = 8 * palRed(k) - SR: SG = 8 * palGreen(k) - SG: SB = 8 * palBlue(k) - SB
      FixSRGB
      ' Force B/W unless unsorted palette
      If (SR + SG + SB) \ 3 > PCONTOUR Then bDummy(ix, iy) = 1
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Dither()
'Public PDITHER As Long        ' 16 ->  48 'Floyd-Steinberg B
' Spreader
' 0 7 0
' 3 5 1 /16
Dim BB As Byte
Dim k As Long
Dim greysum As Long
Dim greycount As Long
Dim zDiv As Single
Dim zMul As Single
Dim zErr As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim Intens(-1 To canvasW + 2, -1 To canvasH + 2)
   greysum = 0
   greycount = canvasW * canvasH
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      k = bArray(ix, iy)
      Intens(ix, iy) = (1& * palRed(k) + palGreen(k) + palBlue(k)) \ 3
      greysum = greysum + Intens(ix, iy)
   Next ix
   Next iy
   greysum = greysum \ greycount
   
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   zDiv = PDITHER  ' 16
   zMul = 1 / zDiv
   For iy = HHLO + 1 To HHHI - 1
   For ix = WWLO + 1 To WWHI - 1
      k = bArray(ix, iy)
      If Intens(ix, iy) > greysum Then
         bDummy(ix, iy) = 1
         zErr = (Intens(ix, iy) - 255) * zMul
      Else
         zErr = Intens(ix, iy) * zMul
      End If
      ' Spread error
      Intens(ix - 1, iy + 1) = Intens(ix - 1, iy + 1) + 3 * zErr
      Intens(ix, iy + 1) = Intens(ix, iy + 1) + 5 * zErr
      Intens(ix + 1, iy + 1) = Intens(ix + 1, iy + 1) + zErr
      Intens(ix + 1, iy) = Intens(ix + 1, iy) + 7 * zErr
   Next ix
   Next iy
   Erase Intens()
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub EngraveEmboss()
'Public PENGRAVEMBOSS As Long  ' -3 -> +3
'+1 0 -1    -1  0  +1
'+1 0 -1    -1  0  +1
'+1 0 -1    -1  0  +1

   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO + 1 To HHHI - 1
   For ix = WWLO + 1 To WWHI - 1
      SR = 0: SG = 0: SB = 0
      For j = iy - 1 To iy + 1
      For i = ix - 1 To ix + 1
         If i <> ix Then
            k = bArray(i, j)
            If i = ix + 1 Then
               SR = SR - PENGRAVEMBOSS * palRed(k)
               SG = SG - PENGRAVEMBOSS * palGreen(k)
               SB = SB - PENGRAVEMBOSS * palBlue(k)
            Else  ' i=ix-1
               SR = SR + PENGRAVEMBOSS * palRed(k)
               SG = SG + PENGRAVEMBOSS * palGreen(k)
               SB = SB + PENGRAVEMBOSS * palBlue(k)
            End If
         End If
      Next i
      Next j
      SR = SR + palRed(Selectcn)
      SG = SG + palGreen(Selectcn)
      SB = SB + palBlue(Selectcn)
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Posterize()
'Public PPOSTERIZE As Long     ' Threshold 32 -> 224
'255*(1 + (R1 - 128)/abs(R1 - 128))/2
'255*(1 + (G1 - 128)/abs(G1 - 128))/2
'255*(1 + (B1 - 128)/abs(B1 - 128))/2
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO + 1 To HHHI - 1
   For ix = WWLO + 1 To WWHI - 1
      k = bArray(ix, iy)
      pR = (1& * palRed(k) - PPOSTERIZE)
      If pR = 0 Then
         SR = palRed(Selectcn) '255
      Else
         SR = palRed(Selectcn) * (1 + pR / Abs(pR)) / 2
      End If
      pG = (1& * palGreen(k) - PPOSTERIZE)
      If pG = 0 Then
         SG = palGreen(Selectcn) '255
      Else
         SG = palGreen(Selectcn) * (1 + pG / Abs(pG)) / 2
      End If
      pB = (1& * palBlue(k) - PPOSTERIZE)
      If pB = 0 Then
         SB = palBlue(Selectcn) '255
      Else
         SB = palBlue(Selectcn) * (1 + pB / Abs(pB)) / 2
      End If
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Relief()
'Public PRELIEF As Long        ' 1,2,3
'+2 +1  0
'+1  0 -1
'0  -1 -2
Dim N  As Long
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   For N = 1 To PRELIEF
      ReDim bDummy(canvasW, canvasH)
      If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         SR = 0: SG = 0: SB = 0
         k = bArray(ix - 1, iy + 1)
         SR = SR + 2& * palRed(k): SG = SG + 2& * palGreen(k): SB = SB + 2& * palBlue(k)
         k = bArray(ix, iy + 1)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         k = bArray(ix - 1, iy)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         k = bArray(ix + 1, iy - 1)
         SR = SR - 2& * palRed(k): SG = SG - 2& * palGreen(k): SB = SB - 2& * palBlue(k)
         k = bArray(ix, iy - 1)
         SR = SR - palRed(k): SG = SG - palGreen(k): SB = SB - palBlue(k)
         k = bArray(ix + 1, iy)
         SR = SR - palRed(k): SG = SG - palGreen(k): SB = SB - palBlue(k)
         k = bArray(ix, iy)
         SR = (SR + palRed(k) + palRed(Selectcn)) \ 3
         SG = (SG + palGreen(k) + palGreen(Selectcn)) \ 3
         SB = (SB + palBlue(k) + palBlue(Selectcn)) \ 3
         FixSRGB
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
      bArray() = bDummy()
   Next N
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Smooth()
'Public PSMOOTH As Long        ' 1,2,3,4
'  1      2      3        4
'                     1 1 1 1 1
'0 0 0  0 1 0  1 1 1  1 1 1 1 1
'1 0 1  1 0 1  1 0 1  1 1 0 1 1
'0 0 0  0 1 0  1 1 1  1 1 1 1 1
'                     1 1 1 1 1
' PSMOOTH 1,2,3,4
' Edges not done
Dim N As Long

   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   If PSMOOTH = 5 Then PSMOOTH = 4
   Select Case PSMOOTH
   Case 1
      '0 0 0
      '1 0 1
      '0 0 0
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         SR = 0: SG = 0: SB = 0
         j = iy
         
         k = bArray(ix - 1, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         k = bArray(ix + 1, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         SR = SR \ 2
         SG = SG \ 2
         SB = SB \ 2
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
   Case 2
      '0 1 0
      '1 0 1
      '0 1 0
      ReDim bDummy(canvasW, canvasH)
      If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         SR = 0: SG = 0: SB = 0
         j = iy - 1
         i = ix
         k = bArray(i, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         j = iy
         k = bArray(ix - 1, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         k = bArray(ix + 1, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         j = iy + 1
         i = ix
         k = bArray(i, j)
         SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
         
         SR = SR \ 4
         SG = SG \ 4
         SB = SB \ 4
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
   
   Case 3
      '1 1 1
      '1 0 1
      '1 1 1
      ReDim bDummy(canvasW, canvasH)
      If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         SR = 0: SG = 0: SB = 0
         For j = iy - 1 To iy + 1
         For i = ix - 1 To ix + 1
            'If Not (j = iy And i = ix) Then
            If (j <> iy Or i <> ix) Then
               k = bArray(i, j)
               SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
            End If
         Next i
         Next j
         SR = SR \ 8
         SG = SG \ 8
         SB = SB \ 8
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
   Case 4
      '1 1 1 1 1
      '1 1 1 1 1
      '1 1 0 1 1
      '1 1 1 1 1
      '1 1 1 1 1
      ReDim bDummy(canvasW, canvasH)
      If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         N = 0
         SR = 0: SG = 0: SB = 0
         For j = iy - 2 To iy + 2
            If j >= HHLO Then
            If j <= HHHI Then
               For i = ix - 2 To ix + 2
                  If i >= WWLO Then
                  If i <= WWHI Then
                     If (j <> iy Or i <> ix) Then
                        N = N + 1
                        k = bArray(i, j)
                        SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
                     End If
                  End If
                  End If
               Next i
            End If
            End If
         Next j
         SR = SR \ N '24
         SG = SG \ N '24
         SB = SB \ N '24
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
   
   End Select
   
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub ShadeV()
'Public PSHADE As Long         ' -100 -> +100
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   zA = PSHADE
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      zB = 256 * Abs(1 - 2 * ix / (WWHI - WWLO + 1))
      k = bArray(ix, iy)
      SR = zA * (zB - palRed(k)) \ 256 + palRed(k)
      SG = zA * (zB - palGreen(k)) \ 256 + palGreen(k)
      SB = zA * (zB - palBlue(k)) \ 256 + palBlue(k)
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub ShadeH()
'Public PSHADE As Long         ' -100 -> +100
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   zA = PSHADE
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For ix = WWLO To WWHI
   For iy = HHLO To HHHI
      zB = 256 * Abs(1 - 2 * iy / (HHHI - HHLO + 1))
      k = bArray(ix, iy)
      SR = zA * (zB - palRed(k)) \ 256 + palRed(k)
      SG = zA * (zB - palGreen(k)) \ 256 + palGreen(k)
      SB = zA * (zB - palBlue(k)) \ 256 + palBlue(k)
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Melt()
'Public PMELT As Long          ' 2 -> 8
Dim NN  As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   bDummy() = bArray()
   For NN = 1 To PMELT
      For iy = HHLO To HHHI - NN
      For ix = WWLO To WWHI - 1
         k = bArray(ix, iy)
         SR = palRed(k): SG = palGreen(k): SB = palBlue(k)
         k = bArray(ix, iy + NN)
         SR = SR - palRed(k): SG = SG - palGreen(k): SB = SB - palBlue(k)
         SR = SR + SG + SB
         If SR < 0 Then
            bDummy(ix, iy) = bArray(ix, iy + NN)
         Else
            bDummy(ix, iy) = bArray(ix, iy)
         End If
      Next ix
      Next iy
      bArray() = bDummy()
   Next NN
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Oil()
'Public POIL As Long           ' 1 -> 5
Dim SGrey  As Long
Dim NN As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   bDummy() = bArray()
      HH = (POIL + 1) \ 2
      If HH < 1 Then HH = 1
      WW = HH
      For iy = HHLO + HH To HHHI '- 1
      For ix = WWLO + WW To WWHI '- 1
         ReDim Hist(0 To 255)
         ReDim IndexK(0 To 255)
         For jj = iy - POIL To iy + POIL
            If jj > 0 Then
            If jj <= canvasH Then
               For ii = ix - POIL To ix + POIL
                  If ii > 0 Then
                  If ii <= canvasW Then
                     k = bArray(ii, jj)
                     SR = palRed(k): SG = palGreen(k): SB = palBlue(k)
                     SGrey = (1& * 0.3 * SG + 0.6 * SG + 0.1 * SB) \ 3
                     Hist(SGrey) = Hist(SGrey) + 1
                     IndexK(SGrey) = k
                  End If
                  End If
               Next ii
            End If
            End If
         Next jj
         NN = 0
         For k = 0 To 255
            If Hist(k) > NN Then
               NN = Hist(k)
               i = IndexK(k)
            End If
         Next k
         bDummy(ix, iy) = i
      Next ix
      Next iy
      bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Sharpen()
'Public PSHARPEN As Long       ' 1,2,3
'-2 -2 -2
'-2 26 -2
'-2 -2 -2
Dim N As Long

   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   For N = 1 To PSHARPEN
      ReDim bDummy(canvasW, canvasH)
      If aSelRect Then ZeroSelRect
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         SR = 0: SG = 0: SB = 0
         For j = iy - 1 To iy + 1
         For i = ix - 1 To ix + 1
            'If Not (j = iy And i = ix) Then
            If (j <> iy Or i <> ix) Then
               k = bArray(i, j)
               SR = SR - 2 * palRed(k): SG = SG - 2 * palGreen(k): SB = SB - 2 * palBlue(k)
            Else
               SR = SR + 26 * palRed(k): SG = SG + 26 * palGreen(k): SB = SB + 26 * palBlue(k)
            End If
         Next i
         Next j
         SR = SR \ 10
         SG = SG \ 10
         SB = SB \ 10
         FixSRGB
         'GetIndex
         LongDerived = RGB(SR, SG, SB)
         i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
         bDummy(ix, iy) = i
      Next ix
      Next iy
      bArray() = bDummy()
   Next N
'   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Litho()
'Public PLITHO As Long         ' 92 -> 100
'-1  -1  -1  -1 -1
'-1 -10 -10 -10 -1
'-1 -10  96 -10 -1
'-1 -10 -10 -10 -1
'-1  -1  -1  -1 -1
   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO + 2 To HHHI - 2
   For ix = WWLO + 2 To WWHI - 3
      SR = 0: SG = 0: SB = 0
      For j = iy - 2 To iy + 2
      For i = ix - 2 To ix + 2
         'If Not (j = iy And i = ix) Then
         k = bArray(i, j)
         If (j <> iy Or i <> ix) Then
            If j = iy - 2 Or j = iy + 2 Then
               SR = SR - palRed(k): SG = SG - palGreen(k): SB = SB - palBlue(k)
            ElseIf i = ix - 2 Or i = ix + 2 Then
               SR = SR - palRed(k): SG = SG - palGreen(k): SB = SB - palBlue(k)
            Else
               SR = SR - 10 * palRed(k): SG = SG - 10 * palGreen(k): SB = SB - 10 * palBlue(k)
            End If
         Else
            SR = SR + PLITHO * palRed(k): SG = SG + PLITHO * palGreen(k): SB = SB + PLITHO * palBlue(k)
         End If
      Next i
      Next j
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Contrast()
'Public PCONTRAST              ' -66 -> +66
'1 1 1
'1 0 1
'1 1 1   8*(i,j)-Sum
Dim N As Long
Dim sF As Single
Dim mCol As Long, nCol As Long
Dim i As Long
Dim Factor As Long
   
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   Screen.MousePointer = vbHourglass
   DoEvents
   
   For N = 0 To 765
      Speed(N) = N \ 3
   Next N
   
   mCol = 0
   nCol = 0
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SB = palBlue(i): SG = palGreen(i): SR = palRed(i)
      mCol = mCol + Speed(SB + SG + SR)
      nCol = nCol + 1
   Next ix
   Next iy
   mCol = mCol \ nCol
   
   sF = (PCONTRAST + 100) / 100
   For N = 0 To 255
      Speed(N) = (N - mCol) * sF + mCol
   Next N
   
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SB = Speed(palBlue(i))
      SG = Speed(palGreen(i))
      SR = Speed(palRed(i))
      Do While (SB < 0) Or (SB > 255) Or (SG < 0) Or (SG > 255) Or (SR < 0) Or (SR > 255)
         If (SB <= 0) And (SG <= 0) And (SR <= 0) Then
            SB = 0: SG = 0: SR = 0
         End If
         If (SB >= 255) And (SG >= 255) And (SR >= 255) Then
            SB = 255: SG = 255: SR = 255
         End If
         If SB < 0 Then
            SG = SG + SB \ 2: SR = SR + SB \ 2: SB = 0
         End If
         If SB > 255 Then
            SG = SG + (SB - 255) \ 2
            SR = SR + (SB - 255) \ 2
            SB = 255
         End If
         If SG < 0 Then
            SB = SB + SG \ 2: SR = SR + SG \ 2: SG = 0
         End If
         If SG > 255 Then
            SB = SB + (SG - 255) \ 2
            SR = SR + (SG - 255) \ 2
            SG = 255
         End If
   
         If SR < 0 Then
            SG = SG + SR \ 2: SB = SB + SR \ 2: SR = 0
         End If
         If SR > 255 Then
            SG = SG + (SR - 255) \ 2
            SB = SB + (SR - 255) \ 2
            SR = 255
         End If
      Loop
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   DoEvents
   ' Erase bDummy()   ' Causes bDummy(ix, iy) = i error
                      ' Subscript out of range ??
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Diffuse()
'Public PDIFFUSE As Long       ' 1  ->  16
   Screen.MousePointer = vbHourglass
   DoEvents
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      Select Case TransformType
      Case TDiffuse
         j = Rnd * PDIFFUSE - PDIFFUSE \ 2
         i = Rnd * PDIFFUSE - PDIFFUSE \ 2
      Case THDiffuse
         j = 0
         i = Rnd * PDIFFUSE - PDIFFUSE \ 2
      Case TVDiffuse
         j = Rnd * PDIFFUSE - PDIFFUSE \ 2
         i = 0
      End Select
      If ix + i < 1 Then i = 0
      If ix + i > canvasW Then i = 0
      If iy + j < 1 Then j = 0
      If iy + j > canvasH Then j = 0
      bArray(ix, iy) = bArray(ix + i, iy + j)
   Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub BlackWhite()
'Public PBLACKWHITE As Long    ' 1 -> 255
Dim i As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SG = (1& * palBlue(i) + palGreen(i) + palRed(i)) \ 3
      If SG < PBLACKWHITE Then
         bArray(ix, iy) = 0
      Else
         bArray(ix, iy) = 1
      End If
   Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Solarize()
'Public PSOLAR As Long         ' 32 -> 224
' Invert threshold
Dim i As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SR = palRed(i): SG = palGreen(i): SB = palBlue(i)
      If SR < PSOLAR Then SR = 255 - SR
      If SG < PSOLAR Then SG = 255 - SG
      If SB < PSOLAR Then SB = 255 - SB
      If Selectcn > 1 Then
         SR = (SR + palRed(Selectcn)) \ 2
         SG = (SG + palGreen(Selectcn)) \ 2
         SB = (SB + palBlue(Selectcn)) \ 2
      End If
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Invert()
' Invert all
Dim i As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SR = 255 - palRed(i): SG = 255 - palGreen(i): SB = 255 - palBlue(i)
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Fog()
'Public PFOG As Long           ' -100 -> +100
Dim N As Long
Dim i As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   N = 144
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      i = bArray(ix, iy)
      SR = palRed(i): SG = palGreen(i): SB = palBlue(i)
      If SR > N Then SR = SR - PFOG
      If SR < N Then SR = N
      If SG > N Then SG = SG - PFOG
      If SG < N Then SG = N
      If SB > N Then SB = SB - PFOG
      If SB < N Then SB = N
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Pixelize()
'Public PSQUARE As Long         ' 1 to 16
Dim i As Long
Dim sx As Long, sy As Long
Dim iix As Long, iiy As Long
Dim N As Long
Dim PSQ As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   ReDim bDummy(canvasW, canvasH)
   If aSelRect Then ZeroSelRect
   SR = 0: SG = 0: SB = 0
   PSQ = PSQUARE * Sqr((canvasW ^ 2 + canvasH ^ 2) / (svcanvasW ^ 2 + svcanvasH ^ 2))
   If PSQ = 0 Then PSQ = 1
   For iy = HHLO To HHHI
   sy = (iy \ PSQ) * PSQ + 1
   For ix = WWLO To WWHI
      sx = (ix \ PSQ) * PSQ + 1
      If ((ix - WWLO) Mod PSQ) = 0 Then ' At start of block
         SR = 0: SG = 0: SB = 0
         N = 0
         For iix = sx To sx + PSQ - 1
            For iiy = sy To sy + PSQ - 1
               If iix <= WWHI Then
               If iiy <= HHHI Then
                  i = bArray(iix, iiy)
                  SR = SR + palRed(i): SG = SG + palGreen(i): SB = SB + palBlue(i)
                  N = N + 1
               End If
               End If
            Next iiy
         Next iix
         If N > 0 Then
            SR = SR \ N
            SG = SG \ N
            SB = SB \ N
         End If
      End If
      FixSRGB
      'GetIndex
      LongDerived = RGB(SR, SG, SB)
      i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
      bDummy(ix, iy) = i
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FixSRGB()
   If SR < 0 Then SR = 0
   If SG < 0 Then SG = 0
   If SB < 0 Then SB = 0
   If SR > 255 Then SR = 255
   If SG > 255 Then SG = 255
   If SB > 255 Then SB = 255
End Sub

'#### DEFORMERS ###############################################

Public Sub FixLimits()
   If aSelRect Then
      bDummy() = bArray()
      For iy = HHLO + 1 To HHHI - 1
      For ix = WWLO + 1 To WWHI - 1
         If aUseSelectcn Then
            bDummy(ix, iy) = Selectcn
         Else
            bDummy(ix, iy) = bArray(ix, iy)
         End If
      Next ix
      Next iy
      CH = HHHI - HHLO + 1
      CW = WWHI - WWLO + 1
   Else
      If aUseSelectcn Then
         FillMemory bDummy(1, 1), canvasW * canvasH, Selectcn
      Else
         bDummy() = bArray()
      End If
      CH = canvasH
      CW = canvasW
   End If
End Sub

Public Sub Elliptic()
'Public zPELLIPSE As Single    ' 0.005 -> 1
Dim xc As Single
Dim yc As Single
Dim zH As Single
Dim xi As Single
Dim xw As Single

   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   If zPELLIPSE < 1 Then
      zB = zPELLIPSE * (CH - 1) / 2
      zA = CW / 2
   Else
      zB = CH / 2
      zA = (2 - zPELLIPSE) * (CW - 1) / 2
   End If
   xc = WWLO - 1 + CW / 2
   yc = HHLO - 1 + CH / 2
   zH = (CH - 1) / 2
   For iy = HHLO To HHHI
      zix = (1 - ((iy - yc) / zB) ^ 2)
      If zix >= 0 Then
         xw = zA * Sqr(zix)     ' half width
         xi = WWLO - 1 + CW / 2 - xw ' indent
         xw = 2 * xw / (CW - 1)  ' ellipse W/full W
         iyd = 0.5 + yc + (((iy - (HHLO - 1)) - 1) - zH) * zB / zH
         If iyd = 0 Then iyd = HHLO
         For ix = WWLO To WWHI
            ixd = 1 + xi + ((ix - (WWLO - 1)) - 1) * xw + 0.5
            If ixd = 0 Then ixd = WWLO
            If ixd > WWLO - 1 Then
            If ixd <= WWHI Then
               bDummy(ixd, iyd) = bArray(ix, iy)
            End If
            End If
         Next ix
      End If
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FluteH()
'Public PFLUTE As Long         ' -50 -> +50 (exc -2 to +2)
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zA = CH / Abs(PFLUTE)
   'For ix = 1 To canvasW
   For ix = WWLO To WWHI
      If PFLUTE > 0 Then
         i = Int(3 * zA * (ix - WWLO + 1) / CW)
      Else
         i = Int(3 * zA * (1 - (ix - WWLO + 1) / CW))
      End If
      zB = (CH - 2 * i) / (CH)
      For iy = HHLO To HHHI
         iyd = zB * (iy - HHLO + 1) + i + HHLO
         If iyd > HHLO - 1 Then
         If iyd <= HHHI Then
            bDummy(ix, iyd) = bArray(ix, iy)
         End If
         End If
      Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FluteV()
'Public PFLUTE As Long         ' -50 -> +50 (exc -2 to +2)
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zA = CW / Abs(PFLUTE)
   For iy = HHLO To HHHI
      If PFLUTE > 0 Then
         i = Int(3 * zA * (iy - HHLO + 1) / CH)
      Else
         i = Int(3 * zA * (1 - (iy - HHLO + 1) / CH))
      End If
      zB = (CW - 2 * i) / (CW)
      For ix = WWLO To WWHI
         ixd = zB * (ix - WWLO + 1) + i + WWLO
         If ixd > WWLO - 1 Then
         If ixd <= WWHI Then
            bDummy(ixd, iy) = bArray(ix, iy)
         End If
         End If
      Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub RippleH()
'Public PRIPPLE As Long        ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zA = CW / Abs(PRIPPLE)
   'For iy = 1 To canvasH
   For iy = HHLO To HHHI
      zB = CH \ 20
      i = Int(zA * (1 + Sin(pi# * (1 + PRIPPLE * (iy - HHLO + 1) / CH))))
      zB = (CW - 2 * i) / (CW)
      For ix = WWLO To WWHI
         'ixd = zB * (ix - 1) + i + WWLO
         ixd = zB * (ix - WWLO) + i + WWLO
         If ixd > WWLO - 1 Then
         If ixd <= WWHI Then
            bDummy(ixd, iy) = bArray(ix, iy)
         End If
         End If
      Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub RippleV()
'Public PRIPPLE As Long        ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zA = CH / Abs(PRIPPLE)
   'For ix = 1 To canvasW
   For ix = WWLO To WWHI
      zB = CW \ 2 '0
      i = Int(zA * (1 + Sin(pi# * (1 + PRIPPLE * (ix - WWLO) / CW))))
      zB = (CH - 2 * i) / (CH)
      For iy = HHLO To HHHI
         iyd = zB * ((iy - HHLO) - 1) + i + HHLO
         If iyd > HHLO - 1 Then
         If iyd <= HHHI Then
            bDummy(ix, iyd) = bArray(ix, iy)
         End If
         End If
      Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub RoundRect()
'Public zPROUNDRECT As Single  ' 0 -> .5
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zA = zPROUNDRECT * CW ' Corner radius
   If zA > CH \ 2 Then zA = CH \ 2
   For iy = HHLO To HHHI
      If iy <= zA + (HHLO - 1) Then
         i = zA - Sqr(zA * zA - (zA - (iy - HHLO + 1)) * (zA - (iy - HHLO + 1)))
      ElseIf iy >= HHHI - zA Then
         i = zA - Sqr(zA * zA - (zA - (HHHI - iy)) * (zA - (HHHI - iy))) '- 1
      Else
         i = 0
      End If
      zB = (CW - 2 * i) / (CW)
      For ix = WWLO To WWHI
         ixd = zB * (ix - WWLO + 1) + i + WWLO
         If ixd > WWLO - 1 Then
         If ixd <= WWHI Then
            bDummy(ixd, iy) = bArray(ix, iy)
         End If
         End If
      Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Tile()
'Public PTILE As Long          ' 1 -> 21
Dim iyy As Long
Dim ixx As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   WW = CW \ PTILE
   HH = CH \ PTILE
   j = HHLO
   For iy = HHLO To HHHI Step PTILE
      i = WWLO
      For ix = WWLO To WWHI Step PTILE
         bDummy(i, j) = bArray(ix, iy)
         i = i + 1
      Next ix
      j = j + 1
   Next iy
   j = HHLO
   For jj = HHLO To (PTILE + 10) * HH - 1 + HHLO - 1
      If jj <= (HHLO - 1) + CH Then
         i = WWLO
         For ii = WWLO To (PTILE + 10) * WW - 1 + WWLO - 1
            If ii <= (WWLO - 1) + CW Then
               bDummy(ii, jj) = bDummy(i, j)
            End If
            i = i + 1
            If i > (WWLO - 1) + WW Then i = WWLO
         Next ii
      End If
      j = j + 1
      If j > (HHLO - 1) + HH Then j = HHLO
   Next jj
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub MirrorLeft()
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   For iy = HHLO To (HHLO - 1) + CH
   For ix = WWLO To (WWLO - 1) + CW \ 2
      bDummy(WWHI - (ix - WWLO + 1), iy) = bArray(ix, iy)
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub MirrorRight()
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   For iy = HHLO To (HHLO - 1) + CH
   For ix = (WWLO - 1) + CW \ 2 + 1 To (WWLO - 1) + CW
      bDummy(WWHI - (ix - WWLO - 1), iy) = bArray(ix, iy)
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub MirrorTop()
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   For ix = WWLO To (WWLO - 1) + CW
   For iy = (HHLO - 1) + CH \ 2 To (HHLO - 1) + CH
      bDummy(ix, HHHI - (iy - HHLO - 1)) = bArray(ix, iy)
   Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub MirrorBottom()
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   For ix = WWLO To (WWLO - 1) + CW
   For iy = HHLO To (HHLO - 1) + CH \ 2
      bDummy(ix, HHHI - (iy - HHLO + 1)) = bArray(ix, iy)
   Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub MirrorLens()
'Public PMLENS As Long         ' 2 -> 40
' Mirror top, cylinder horz stretch
Dim zA As Single
Dim zR As Single
Dim xc As Single
Dim yc As Single
Dim xd As Single
Dim yd As Single
Dim ixv As Long
Dim iyv As Long
Dim zRMax As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   xc = CW / 2: yc = CH / 2
   zRMax = Sqr((CW - xc) ^ 2 + (CH - yc) ^ 2)
   For iy = HHLO To (HHLO - 1) + CH
   For ix = WWLO To (WWLO - 1) + CW
      xd = ix - (WWLO + xc)
      yd = iy - (HHLO + yc)
      zA = zATan2(yd, xd)
      zR = Sqr(yd * yd + xd * xd)
      zA = zA * (ix - (WWLO - 1 + xc)) / (PMLENS * pi#)
      zR = zR * (iy - (HHLO - 1 + yc)) / zRMax
      ixv = (WWLO - 1 + xc) + zR * Cos(zA)
      iyv = (HHLO - 1 + yc) - zR * Sin(zA)
      If ixv > WWLO - 1 Then
      If ixv <= WWHI Then
      If iyv > HHLO - 1 Then
      If iyv <= HHHI Then
        bDummy(ix, iy) = bArray(ixv, iyv)
      End If
      End If
      End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub ALens()
'Public zPLENS As Single       ' 0.1 -> 4
Dim zA As Single
Dim zR As Single
Dim xc As Single
Dim yc As Single
Dim xd As Single
Dim yd As Single
Dim ixv As Long
Dim iyv As Long
Dim zRMax As Single
Dim zLim As Single

   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   xc = CW / 2: yc = CH / 2
   zRMax = Sqr((CW - xc) ^ 2 + (CH - yc) ^ 2) / 2
   zLim = 2
   For iy = HHLO To (HHLO - 1) + CH
   For ix = WWLO To (WWLO - 1) + CW
      yd = iy - yc - 1 - (HHLO - 1)
      xd = ix - xc - 1 - (WWLO - 1)
      zA = zATan2(yd, xd)
      zR = Sqr((yd * yd + xd * xd))
      If zR <= zRMax * zPLENS Then
         zR = zR * (1 + 4 * zR) / (1 + 4 * zRMax * zPLENS)
         ixv = xc + zR * Cos(zA) + (WWLO - 1)
         iyv = yc + zR * Sin(zA) + (HHLO - 1)
         If ixv > WWLO - 1 Then
         If ixv <= WWHI Then
         If iyv > HHLO - 1 Then
         If iyv <= HHHI Then
            If zR < (zRMax * zPLENS - zLim) Then
              bDummy(ix, iy) = bArray(ixv, iyv)
            ElseIf aLensCheck Then  ' Draw circle around lens
               If zR < (zRMax * zPLENS + zLim) Then
                 bDummy(ix, iy) = Selectcn
               End If
            End If
         End If
         End If
         End If
         End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub Bubbly()
'Public zPLENS As Single       ' 0.1 -> 4
Dim zA As Single
Dim zR As Single
Dim ixv As Long
Dim iyv As Long
Dim ixdr As Long
Dim iydr As Long
Dim zRM As Single
   
   ' zPLENS .1 to 4
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   zRM = zPLENS * Sqr(CW ^ 2 + CH ^ 2) / 50
   Rnd (-0.5)
   Randomize (1)
   For iy = zRM To HHHI + zRM Step 2 * zRM
   For ix = zRM To WWHI + zRM Step 2 * zRM
      
      If Rnd < 0.4 Then
         
         For jj = -zRM To zRM
         For ii = -zRM To zRM
            zA = zATan2(CSng(jj), CSng(ii))
            zR = Sqr(ii * ii + jj * jj)
            If zR <= zRM Then
                  zR = zR * (1 + 16 * zR) / (1 + 16 * zRM * zPLENS)
                  ixv = ix + zR * Cos(zA)
                  iyv = iy + zR * Sin(zA)
                  If ixv > WWLO - 1 Then
                  If ixv <= WWHI Then
                  If iyv > HHLO - 1 Then
                  If iyv <= HHHI Then
                     ixd = ix + ii
                     iyd = iy + jj
                     ixdr = ix + ii + zRM '* Rnd
                     iydr = iy + jj + zRM '* Rnd
                     If ixd > 1 Then
                     'If ixd < CW Then
                     If ixd < WWHI Then
                     'If iyd > 1 Then
                     If iyd > HHLO Then
                     'If iyd < CH Then
                     If iyd < HHHI Then
                           bDummy(ixd, iyd) = bArray(ixv, iyv)  'Selectcn filled
                           If ixdr > WWLO Then
                           If ixdr < WWHI Then
                           If iydr > HHLO Then
                           If iydr < HHHI Then
                              bDummy(ixdr, iydr) = bArray(ixv, iyv)
                           End If
                           End If
                           End If
                           End If
                     End If
                     End If
                     End If
                     End If
                  End If
                  End If
                  End If
                  End If
            End If
         Next ii
         Next jj
      
      End If

   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FlutedWindowHorz()
'Public PFWINDOW As Long       ' 8,10,12,, 32
' Horz fluted window
Dim NN As Long
Dim iyv As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   NN = CH \ PFWINDOW
   For ix = WWLO To (WWLO - 1) + CW
   For iy = HHLO To (HHLO - 1) + CH
      iyv = iy + (iy Mod NN) - NN \ 2
      If iyv > HHLO - 1 Then
      If iyv <= HHHI Then
         bDummy(ix, iy) = bArray(ix, iyv)
      End If
      End If
   Next iy
   Next ix
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FlutedWindowVert()
'Public PFWINDOW As Long       ' 8,10,12,, 32
' Vert fluted window
Dim NN As Long
Dim ixv As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   NN = CW \ PFWINDOW
   For iy = HHLO To (HHLO - 1) + CH
   For ix = WWLO To (WWLO - 1) + CW
      ixv = ix + (ix Mod NN) - NN \ 2
      If ixv > WWLO - 1 Then
      If ixv <= WWHI Then
         bDummy(ix, iy) = bArray(ixv, iy)
      End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub FlutedWindowHV()
'Public PFWINDOW As Long       ' 8,10,12,, 32
   FlutedWindowHorz
   FlutedWindowVert
End Sub
   
Public Sub Swirl()
'Public zPSWIRL As Single      ' -50 -> +50
Dim zA As Single
Dim zR As Single
Dim xc As Single
Dim yc As Single
Dim xd As Single
Dim yd As Single
Dim ixv As Long
Dim iyv As Long
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   zD = Sqr(CW * CW + CH * CH)
   
   xc = WWLO - 1 + CW / 2: yc = HHLO - 1 + CH / 2
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      xd = (ix - xc)
      yd = (iy - yc)
      zA = zATan2(yd, xd)
      zR = Sqr(yd * yd + xd * xd)
      zA = zA + (zR / zD) * zPSWIRL
      ixv = xc + zR * Cos(zA)
      iyv = yc + Sgn(zPSWIRL) * zR * Sin(zA)
      If ixv > WWLO - 1 Then
      If ixv <= WWHI Then
      If iyv > HHLO - 1 Then
      If iyv <= HHHI Then
        bDummy(ix, iy) = bArray(ixv, iyv)
      End If
      End If
      End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub Stars()
'Public PKALI As Long          ' 1 -> 21
Dim zA As Single
Dim zR As Single
Dim zRR As Single
Dim xc As Single
Dim yc As Single
Dim xd As Single
Dim yd As Single
Dim ixv As Long
Dim iyv As Long
Dim zRMax As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   xc = CW / 2: yc = CH / 2
   zRMax = Sqr((CW - xc) ^ 2 + (CH - yc) ^ 2)
   xc = WWLO - 1 + CW / 2: yc = HHLO - 1 + CH / 2
   
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      xd = ix - xc
      yd = iy - yc
      zA = zATan2(yd, xd) * PKALI
      zR = Sqr(yd * yd + xd * xd)
      zRR = Sqr(zR * zRMax)
      ixv = xc + zRR * Cos(zA)
      iyv = yc + zRR * Sin(zA)
      If ixv > WWLO - 1 Then
      If ixv <= WWHI Then
      If iyv > HHLO - 1 Then
      If iyv <= HHHI Then
        bDummy(ix, iy) = bArray(ixv, iyv)
      End If
      End If
      End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub
   
Public Sub MinMag()
'Public zPMINMAG As Single     ' .01 -> 2 not 0
Dim zA As Single
Dim zR As Single
Dim xc As Single
Dim yc As Single
Dim xd As Single
Dim yd As Single
Dim ixv As Long
Dim iyv As Long
Dim zRMax As Single
Dim zMR As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   FixLimits
   
   xc = CW / 2: yc = CH / 2
   zRMax = Sqr((CW - xc) ^ 2 + (CH - yc) ^ 2)
   
   xc = WWLO - 1 + CW / 2: yc = HHLO - 1 + CH / 2
   zMR = zRMax * zPMINMAG
   If zMR <= 0 Then Exit Sub
   
   For iy = HHLO To HHHI
   For ix = WWLO To WWHI
      xd = ix - xc
      yd = iy - yc
      zA = zATan2(yd, xd)
      zR = Sqr(yd * yd + xd * xd)
      zR = (zR * zRMax) / zMR
      ixv = xc + zR * Cos(zA)
      iyv = yc + zR * Sin(zA)
      If ixv > WWLO - 1 Then
      If ixv <= WWHI Then
      If iyv > HHLO - 1 Then
      If iyv <= HHHI Then
        bDummy(ix, iy) = bArray(ixv, iyv)
      End If
      End If
      End If
      End If
   Next ix
   Next iy
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Rotate()
' Public zPROTATE as single   ' degrees  -180 - + 180
' Public ixc As Long, iyc As Long
'Dim xc As Single, yc As Single
Dim ixd As Long, iyd As Long
Dim Xs As Single, Ys As Single
Dim ixs As Long, iys As Long
Dim idx As Single 'Long
Dim zrad As Single
Dim zang As Single
Dim zcos As Single, zsin As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH) As Byte
   
   FixLimits
   
   zang = zPROTATE * d2r#
   zcos = Cos(zang)
   zsin = Sin(-zang)
   ixc = WWLO - 1 + CW / 2: iyc = HHLO - 1 + CH / 2
   zrad = CW / 2
   For iyd = HHLO To HHHI
      idx = CW / 2
      For ixd = ixc - idx To ixc + idx - 1
         If ixd >= WWLO Then
         If ixd <= WWHI Then
            Xs = ixc + CSng(ixd - ixc) * zcos + CSng(iyd - iyc) * zsin
            Ys = iyc + CSng(iyd - iyc) * zcos - CSng(ixd - ixc) * zsin
            ixs = CLng(Xs)
            iys = CLng(Ys)
            If ixs > WWLO - 1 Then
            If ixs <= WWHI Then
            If iys > HHLO - 1 Then
            If iys <= HHHI Then
               bDummy(ixd, iyd) = bArray(ixs, iys)
            End If
            End If
            End If
            End If
         End If
         End If
      Next ixd
   Next iyd
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub Tunnel()
Dim ziy As Single
Dim zix As Single
Dim zNN As Single
Dim zP As Single

   Screen.MousePointer = vbHourglass
   DoEvents
   ReDim bDummy(canvasW, canvasH)
   
   aUseSelectcnSV = aUseSelectcn
   aUseSelectcn = False
   FixLimits
   aUseSelectcn = aUseSelectcnSV
   
   zP = 1 + (PTILE - 1) / 5
   For zNN = 1 To zP Step 0.2
      WW = CW / zNN
      HH = CH / zNN
      j = HHLO - 1 + (CH - HH) / 2 - 1
      For ziy = HHLO To HHHI Step zNN
         i = WWLO - 1 + (CW - WW) / 2 - 1
         For zix = WWLO To WWHI Step zNN
            If j > HHLO - 1 Then
            If j <= HHHI Then
            If i > WWLO - 1 Then
            If i <= WWHI Then
               bDummy(i, j) = bArray(zix, ziy)
            End If
            End If
            End If
            End If
            i = i + 1
         Next zix
         j = j + 1
      Next ziy
   Next zNN
   bArray() = bDummy()
   Erase bDummy()
   Screen.MousePointer = vbDefault
   DoEvents
End Sub


'#### ADDERS ##################################################

Public Sub FixLimits2()
   If aSelRect Then
      CH = HHHI - HHLO + 1
      CW = WWHI - WWLO + 1
   Else
      CH = canvasH
      CW = canvasW
   End If
End Sub

Public Sub AddHLines()
'Public PLINES As Long         ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   For iy = HHLO To HHHI Step (CH / PLINES)
      For ix = WWLO To WWHI
         bArray(ix, iy) = Selectcn
      Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddVLines()
'Public PLINES As Long         ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   
   For ix = WWLO To WWHI Step (CW / PLINES)
      For iy = HHLO To HHHI
         bArray(ix, iy) = Selectcn
      Next iy
   Next ix
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddHVLines()
'Public PLINES As Long         ' 0 -> 40
   AddHLines
   AddVLines
End Sub

Public Sub AddHWaves()
'Public PLINES As Long         ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   zA = (10 + PLINES) / 2
   For iy = HHLO To HHHI Step CH / (PLINES + 1)
      For zix = WWLO To WWHI + zA Step 1 / zA
         If zix > WWLO - 1 Then
         If zix <= WWHI Then
            jj = iy + zA * Sin(zix * 2 * zA / CW)
            If jj > HHLO - 1 Then
            If jj <= HHHI Then
               bArray(zix, jj) = Selectcn
            End If
            End If
         End If
         End If
      Next zix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddVWaves()
'Public PLINES As Long         ' 0 -> 40
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   zA = (10 + PLINES) / 2
   For ix = WWLO To WWHI Step CW / (PLINES + 1)
      For ziy = HHLO To HHHI + 2 * zA Step 1 / zA
         If ziy > HHLO - 1 Then
         If ziy <= HHHI Then
            ii = ix + zA * Sin(ziy * 2 * zA / canvasH)
            If ii > WWLO - 1 Then
            If ii <= WWHI Then
               bArray(ii, ziy) = Selectcn
            End If
            End If
         End If
         End If
      Next ziy
   Next ix
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddHVWaves()
'Public PLINES As Long         ' 0 -> 40
   AddHWaves
   AddVWaves
End Sub

Public Sub AddCircles()
'Public zPLENS As Single       ' 0.1 -> 4
Dim zA As Single
Dim ixv As Long
Dim iyv As Long
Dim zRM As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   zD = Sqr(CW * CW + CH * CH)
   zRM = zD / (16 * zPLENS)
   For iy = HHLO To HHHI + zRM Step 2 * zRM
   For ix = WWLO To WWHI + zRM Step 2 * zRM
         For zA = 0 To 2 * pi# Step 0.01
            ixv = ix + zRM * Cos(zA)
            iyv = iy + zRM * Sin(zA)
            If ixv > WWLO - 1 Then
            If ixv <= WWHI Then
            If iyv > HHLO - 1 Then
            If iyv <= HHHI Then
                bArray(ixv, iyv) = Selectcn
            End If
            End If
            End If
            End If
         Next zA
   Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddEllipses()
'Public zPLENS As Single       ' 0.1 -> 4
Dim zA As Single
Dim ixv As Long
Dim iyv As Long
Dim zRM As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   zD = Sqr(CW * CW + CH * CH)
   zRM = zD / (8 * zPLENS)
   For iy = HHLO To HHHI + zRM Step zRM
   For ix = WWLO To WWHI + zRM Step 2 * zRM
         For zA = 0 To 2 * pi# Step 0.01
            ixv = ix + zRM * Cos(zA)
            iyv = iy + zRM * Sin(zA) / 2
            If ixv > WWLO - 1 Then
            If ixv <= WWHI Then
            If iyv > HHLO - 1 Then
            If iyv <= HHHI Then
                bArray(ixv, iyv) = Selectcn
            End If
            End If
            End If
            End If
         Next zA
   Next ix
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddThickLineH()
'Public zPTHICKLINE As Single  ' .0 -> 1.0
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   TY = zPTHICKLINE * canvasH
   If TY < 1 Then TY = 1
   For iy = HHLO - 1 + CH / 2 - TY / 2 To HHLO - 1 + CH / 2 + TY / 2
   If iy > HHLO - 1 Then
   If iy <= HHHI Then
      For ix = WWLO To WWHI
         bArray(ix, iy) = Selectcn
      Next ix
   End If
   End If
   Next iy
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddThickLineV()
'Public zPTHICKLINE As Single  ' .0 -> 1.0
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   TX = zPTHICKLINE * canvasW
   If TX < 1 Then TX = 1
   For ix = WWLO - 1 + CW / 2 - TX / 2 To WWLO - 1 + CW / 2 + TX / 2
   If ix > WWLO - 1 Then
   If ix <= WWHI Then
      For iy = HHLO To HHHI
         bArray(ix, iy) = Selectcn
      Next iy
   End If
   End If
   Next ix
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddThickLineHV()
'Public zPTHICKLINE As Single  ' .0 -> 1.0
   AddThickLineH
   AddThickLineV
End Sub

Public Sub AddBorder()
'Public zPTHICKLINE As Single  ' .0 -> 1.0
   Screen.MousePointer = vbHourglass
   DoEvents
   If aSelRect Then
      CH = HHHI - HHLO + 1
      CW = WWHI - WWLO + 1
      TY = zPTHICKLINE * (CH / SSH)
      If TY < 1 Then TY = 1
      TX = zPTHICKLINE * (CW / SSW)
      If TX < 1 Then TY = 1
   Else
      CH = canvasH
      CW = canvasW
      TY = zPTHICKLINE * (CH / svcanvasH)
      If TY < 1 Then TY = 1
      TX = zPTHICKLINE * (CW / svcanvasW)
      If TX < 1 Then TY = 1
   End If
   ' Bottom
   For iy = HHLO To HHLO - 1 + TY
      If iy > HHLO - 1 Then
      If iy <= HHHI Then
         For ix = WWLO To WWHI
            bArray(ix, iy) = Selectcn
         Next ix
      End If
      End If
   Next iy
   ' Top
   For iy = HHHI - TY + 1 To HHHI
      If iy > HHLO - 1 Then
      If iy <= HHHI Then
         For ix = WWLO To WWHI
            bArray(ix, iy) = Selectcn
         Next ix
      End If
      End If
   Next iy
   ' Left
   For ix = WWLO To WWLO - 1 + TX
      If ix > WWLO - 1 Then
      If ix <= WWHI Then
         For iy = HHLO To HHLO - 1 + CH
            bArray(ix, iy) = Selectcn
         Next iy
      End If
      End If
   Next ix
   ' Right
   For ix = WWHI - TX + 1 To WWHI
      If ix > WWLO - 1 Then
      If ix <= WWHI Then
         For iy = HHLO To HHHI
            bArray(ix, iy) = Selectcn
         Next iy
      End If
      End If
   Next ix
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddSpokes()
'Public PSPOKES As Long        ' 1 -> 201
Dim zradi As Single
Dim zphi As Single
Dim xc As Single, yc As Single
Dim zRMax As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   xc = CW / 2: yc = CH / 2
   zRMax = Abs(CW - xc)
   If Abs(CH - yc) < Abs(CW - xc) Then
      zRMax = Abs(CH - yc)
   End If
   xc = WWLO - 1 + CW / 2: yc = HHLO - 1 + CH / 2
   For zradi = 1 To zRMax
   For zphi = 0 To 360 Step 360 / PSPOKES
      ix = xc + zradi * Sin(zphi * d2r#)
      iy = yc + zradi * Cos(zphi * d2r#)
      If ix > WWLO - 1 Then
      If ix <= WWHI Then
      If iy > HHLO - 1 Then
      If iy <= HHHI Then
         bArray(ix, iy) = Selectcn
      End If
      End If
      End If
      End If
   Next zphi
   Next zradi
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

Public Sub AddDiagNet()
'Public PDNET As Long          ' 1 -> 41
Dim N As Long
Dim zix As Single, ziy As Single
   Screen.MousePointer = vbHourglass
   DoEvents
   FixLimits2
   '/ TL
   N = 0
   zix = WWHI
   For ziy = HHLO To HHHI Step (CH / PDNET)
      BresLine WWLO, CLng(ziy), CLng(zix), HHHI, Selectcn, 0
      N = N + 1
      zix = WWHI - N * (CW / PDNET)
   Next ziy
   ' / BR
   N = 0
   ziy = HHHI
   For zix = WWLO To WWHI Step (CW / PDNET)
      BresLine CLng(zix), HHLO, WWHI, ziy, Selectcn, 0
      N = N + 1
      ziy = HHHI - N * (CH / PDNET)
   Next zix
   ' \ BL
   N = 0
   zix = WWLO
   For ziy = HHLO To HHHI Step (CH / PDNET)
      BresLine WWLO, CLng(ziy), CLng(zix), HHLO, Selectcn, 0
      N = N + 1
      zix = WWLO + N * (CW / PDNET)
   Next ziy
   ' \ TR
   N = 0
   ziy = HHLO
   For zix = WWLO To WWHI Step (CW / PDNET)
      BresLine CLng(zix), HHHI, WWHI, ziy, Selectcn, 0
      N = N + 1
      ziy = HHLO + N * (CH / PDNET)
   Next zix
   Screen.MousePointer = vbDefault
   DoEvents
End Sub

