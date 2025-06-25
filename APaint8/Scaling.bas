Attribute VB_Name = "Scaling"
' Scaling.bas

Option Explicit
Option Base 1

Public Sub PictureResize()
' Canvas & picture changed
' NB pic size cannot be > canvas size
Dim xmul As Single
Dim ymul As Single
Dim ixx As Long, iyy As Long
   ' Resize whole pic first
   ReDim bDummy(1 To RpicW, 1 To RpicH) As Byte
   
   xmul = picW / RpicW
   ymul = picH / RpicH
   
   For iyy = 1 To RpicH ' Dest
      iy = iyy * ymul   ' Source
      If iy > 0 And iy <= picH Then
         For ixx = 1 To RpicW ' Dest
            ix = ixx * xmul   ' Source
            If ix > 0 And ix <= picW Then
               ' Dest  <--- Source
               bDummy(ixx, iyy) = bArray(ix, iy)
            End If
         Next ixx
      End If
   Next iyy
   ' Have new pic size in bDummy(RpicW, RpicH)
   ' Transfer to resized bArray
   ReDim bArray(1 To RpicW, 1 To RpicH)
   bArray() = bDummy()
   Erase bDummy()
   
   picW = RpicW
   picH = RpicH
   
   'RelocatePicture
   
End Sub

Public Sub RelocatePicture()
' RpicW = picW & RpicH = picH
' picW,H unchanged
' Canvas changed to RcanvasW,H
' RelocatePicture to new canvas size picW,H
' using ImagePosition ie black border
Dim ixx As Long, iyy As Long
Dim ixd As Long, iyd As Long
Dim ixdd As Long
   ReDim bDummy(1 To RcanvasW, 1 To RcanvasH) As Byte
'   Select Case ImagePosition
'   Case 0: ixd = 1: iyd = RcanvasH - picH + 1  ' TL
'   Case 1: ixd = RcanvasW - picW + 1: iyd = RcanvasH - picH + 1 ' TR
'   Case 2: ixd = (RcanvasW - picW) \ 2 + 1: iyd = (RcanvasH - picH) \ 2 + 1 ' Centre
'   Case 3: ixd = 1: iyd = 1 ' BL
'   Case 4: ixd = RcanvasW - picW + 1: iyd = 1   ' BR
'   End Select
   
'   ixdd = ixd
'   For iy = 1 To picH
'      If iyd > 0 And iyd <= RcanvasH Then
'         ixd = ixdd
'         For ix = 1 To picW
'            If ixd > 0 And ixd <= RcanvasW Then
'               bDummy(ixd, iyd) = bArray(ix, iy)
'            Else
'               Exit For
'            End If
'            ixd = ixd + 1
'         Next ix
'      Else
'         Exit For
'      End If
'      iyd = iyd + 1
'   Next iy
   
   Select Case ImagePosition
   Case 0: ixd = 1: iyd = RcanvasH - canvasH + 1  ' TL
   Case 1: ixd = RcanvasW - canvasW + 1: iyd = RcanvasH - canvasH + 1 ' TR
   Case 2: ixd = (RcanvasW - canvasW) \ 2 + 1: iyd = (RcanvasH - canvasH) \ 2 + 1 ' Centre
   Case 3: ixd = 1: iyd = 1 ' BL
   Case 4: ixd = RcanvasW - canvasW + 1: iyd = 1   ' BR
   End Select
   
   ixdd = ixd
   For iy = 1 To canvasH
      If iyd > 0 And iyd <= RcanvasH Then
         ixd = ixdd
         For ix = 1 To canvasW
            If ixd > 0 And ixd <= RcanvasW Then
               bDummy(ixd, iyd) = bArray(ix, iy)
            Else
               Exit For
            End If
            ixd = ixd + 1
         Next ix
      Else
         Exit For
      End If
      iyd = iyd + 1
   Next iy
   picW = RcanvasW
   picH = RcanvasH
   canvasW = picW
   canvasH = picH
   ReDim bArray(1 To canvasW, 1 To canvasH)
   ' Copy back
   bArray() = bDummy()
   Erase bDummy()
 
End Sub

