Attribute VB_Name = "PICDraw"
'PICDraw.bas

' From Form1.PIC_MouseUp & .PIC_MouseMove

Option Explicit
Option Base 1
Private bTreeArray() As Byte

Public Sub GetColor(Button As Integer)
   If Button = vbLeftButton Then
      DCul = CulRGB(SelLeftCulNum) Xor CulRGB(0)
      CulNum = SelLeftCulNum
   ElseIf Button = vbRightButton Then
      DCul = CulRGB(SelRightCulNum) Xor CulRGB(0)
      CulNum = SelRightCulNum
   End If
End Sub

'Public Sub StartEndDots(PIC As PictureBox, Button As Integer, px As Long, py As Long)
'   GetColor Button
'   SetDot px, py, CulNum, BrushType
''   LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
'End Sub

Public Sub StartEndFreedraws(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   Select Case BrushType
   Case FreeDraw1: PIC.DrawWidth = 1
   Case FreeDraw2: PIC.DrawWidth = 2
   Case FreeDraw3: PIC.DrawWidth = 4
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 1
      ReDim STOREX(NSTOREXY), STOREY(NSTOREXY)
      STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
      PIC.PSet (x, y), DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteFreeDraw CulNum
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndRibbons(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   If BrushType <= 8 Then
      RibIncrX = ((BrushType - 6) * 2) + 1   '+/-  1,3,5
   Else
      RibIncrX = ((BrushType - 9) * 2) + 1   '+/- 1,3,5
   End If
   RibIncrY = RibIncrX
   '      RY1,RX1
   '          \
   ' BRibbon   .
   '            \
   '          RY2,RX2
   RX1 = -RibIncrX
   RY1 = -RibIncrY
   RX2 = RibIncrX
   RY2 = RibIncrY
   If BrushType >= 9 Then
      '          RY1,RX1
      '            /
      ' FRibbon   .
      '          /
      '      RY2,RX2
      RX1 = -RibIncrX
      RY1 = RibIncrY
      RX2 = RibIncrX
      RY2 = -RibIncrY
   End If
   
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 1
      ReDim STOREX(NSTOREXY), STOREY(NSTOREXY)
      STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
      PIC.Line (x + RX1, y + RY1)-(x + RX2, y + RY2), DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteRibbon CulNum
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndSprays(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
Dim iix As Long, iiy As Long
   PIC.DrawMode = 7
   Select Case SprayType
   Case Dots1, Plusses1, Crosses1, Diamonds1: zradmax = 8: sprayn = 8
   Case Dots2, Plusses2, Crosses2, Diamonds2: zradmax = 16: sprayn = 16
   Case Dots3, Plusses3, Crosses3, Diamonds3: zradmax = 20: sprayn = 32
   End Select
   
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      ReDim STOREX(1), STOREY(1)
      NSTOREXY = 0
      
      Select Case SprayType
      Case Dots1, Dots2, Dots3
         For NN = 1 To sprayn
            iix = x - zradmax * Rnd
            iiy = y - zradmax * Rnd
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
         Next NN
      Case Plusses1, Plusses2, Plusses3   ' Changed to n shape but not renamed
         For NN = 1 To sprayn
            iix = x - zradmax * Rnd
            iiy = y - zradmax * Rnd
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
            NSTOREXY = NSTOREXY + 4
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            iix = iix - 1
            STOREX(NSTOREXY - 3) = iix: STOREY(NSTOREXY - 3) = iiy
            PIC.PSet (iix, iiy), DCul
            iiy = iiy + 1
            STOREX(NSTOREXY - 2) = iix: STOREY(NSTOREXY - 2) = iiy
            PIC.PSet (iix, iiy), DCul
            iix = iix + 2
            STOREX(NSTOREXY - 1) = iix: STOREY(NSTOREXY - 1) = iiy
            PIC.PSet (iix, iiy), DCul
            iiy = iiy - 1
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
         Next NN
      Case Crosses1, Crosses2, Crosses3
         For NN = 1 To sprayn
            iix = x - zradmax * Rnd
            iiy = y - zradmax * Rnd
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
            NSTOREXY = NSTOREXY + 4
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            iiy = iiy - 1
            iix = iix - 1
            STOREX(NSTOREXY - 3) = iix: STOREY(NSTOREXY - 3) = iiy
            PIC.PSet (iix, iiy), DCul
            iix = iix + 2
            STOREX(NSTOREXY - 2) = iix: STOREY(NSTOREXY - 2) = iiy
            PIC.PSet (iix, iiy), DCul
            iiy = iiy + 2
            STOREX(NSTOREXY - 1) = iix: STOREY(NSTOREXY - 1) = iiy
            PIC.PSet (iix, iiy), DCul
            iix = iix - 2
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
         Next NN
      Case Diamonds1, Diamonds2, Diamonds3
         For NN = 1 To sprayn
            iix = x - zradmax * Rnd
            iiy = y - zradmax * Rnd
            NSTOREXY = NSTOREXY + 4
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            iiy = iiy - 1
            iix = iix - 1
            STOREX(NSTOREXY - 3) = iix: STOREY(NSTOREXY - 3) = iiy
            PIC.PSet (iix, iiy), DCul
            iix = iix + 2
            STOREX(NSTOREXY - 2) = iix: STOREY(NSTOREXY - 2) = iiy
            PIC.PSet (iix, iiy), DCul
            iiy = iiy + 2
            STOREX(NSTOREXY - 1) = iix: STOREY(NSTOREXY - 1) = iiy
            PIC.PSet (iix, iiy), DCul
            iix = iix - 2
            STOREX(NSTOREXY) = iix: STOREY(NSTOREXY) = iiy
            PIC.PSet (iix, iiy), DCul
         Next NN
      End Select
      PIC.Refresh
   ElseIf LCNum = 2 Then
      CompleteSpray CulNum
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub
   
Public Sub StartEndSingleLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   Select Case LineType
   Case SingleLine1: PIC.DrawWidth = 1
   Case SingleLine2: PIC.DrawWidth = 2
   Case SingleLine3: PIC.DrawWidth = 4
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      PIC.Line (x, y)-(x, y), DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteSingleLine CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndDottedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
   PIC.DrawMode = 7
   Select Case LineType
   Case DottedLine1: PIC.DrawWidth = 1
   Case DottedLine2: PIC.DrawWidth = 2
   Case DottedLine3: PIC.DrawWidth = 4
   End Select
   Select Case LineType
   Case DoubleDottedLine1: zspace = 4: PIC.DrawWidth = 1
   Case DoubleDottedLine2: zspace = 8: PIC.DrawWidth = 1
   Case DoubleDottedLine3: zspace = 16: PIC.DrawWidth = 1
   End Select
   PIC.DrawStyle = vbDot   ' Doesn't work for DrawWidth > 1
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      PIC.Line (x, y)-(x, y), DCul
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         NSTOREXY = 4
         ReDim Preserve STOREX(4), STOREY(4)
         GetParallelCoords zspace, x, y, x, y, ixa, iya, ixb, iyb
         STOREX(3) = ixa: STOREY(3) = iya
         STOREX(4) = ixb: STOREY(4) = iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
      End Select
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteDottedLine CulNum
      PIC.DrawWidth = 1
      PIC.DrawStyle = vbSolid
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndDoubleAndShadedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
   PIC.DrawMode = 7
   Select Case LineType
   Case DoubleLine1, DoubleLineEnd1, ShadedLine1: zspace = 4
   Case DoubleLine2, DoubleLineEnd2, ShadedLine2: zspace = 8
   Case DoubleLine3, DoubleLineEnd3, ShadedLine3: zspace = 16
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 4
      ReDim STOREX(4), STOREY(4)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      PIC.Line (x, y)-(x, y), DCul
      ' line2
      GetParallelCoords zspace, x, y, x, y, ixa, iya, ixb, iyb
      STOREX(3) = ixa: STOREY(3) = iya
      STOREX(4) = ixb: STOREY(4) = iyb
      PIC.Line (ixa, iya)-(ixb, iyb), DCul
      If LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
         PIC.Line (x, y)-(ixa, iya), DCul
         PIC.Line (x, y)-(ixb, iyb), DCul
      End If
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteDoubleLine CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndPolySingleLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
      If StartButton = 0 Then
         StartButton = Button
         GetColor Button
         Select Case PolyLineType
         Case PolySingleLine1: PIC.DrawWidth = 1
         Case PolySingleLine2: PIC.DrawWidth = 2
         Case PolySingleLine3: PIC.DrawWidth = 4
         End Select
         PIC.DrawMode = 7
      End If
      
      If Button = vbLeftButton Then LCNum = LCNum + 1
      If Button = vbRightButton Then RCNum = RCNum + 1
      
      If AMoveAll Then
         If StartButton = vbLeftButton Then
            If RCNum = 1 Then
               CompletePolySingleLine CulNum
               PIC.DrawWidth = 1
               LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
               Exit Sub
            End If
         End If
         If StartButton = vbRightButton Then
            If LCNum = 1 Then
               CompletePolySingleLine CulNum
               PIC.DrawWidth = 1
               LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
               Exit Sub
            End If
         End If
       End If
      
      If StartButton = vbLeftButton Then
         If LCNum = 1 And RCNum = 0 Then
            AMoveAll = False
            NSTOREXY = 2
            ReDim STOREX(2), STOREY(2)
            STOREX(1) = x: STOREY(1) = y
            STOREX(2) = x: STOREY(2) = y
            PIC.Line (x, y)-(x, y), DCul
            PIC.Refresh
         ElseIf LCNum > 1 And RCNum = 0 Then
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
            PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(x, y), DCul
            PIC.Refresh
         ElseIf RCNum = 1 Then  ' Move all
            AMoveAll = True
         ElseIf RCNum = 2 Then
            CompletePolySingleLine CulNum
            PIC.DrawWidth = 1
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
         End If
      
      ElseIf StartButton = vbRightButton Then
         If RCNum = 1 And LCNum = 0 Then
            AMoveAll = False
            NSTOREXY = 2
            ReDim STOREX(2), STOREY(2)
            STOREX(1) = x: STOREY(1) = y
            STOREX(2) = x: STOREY(2) = y
            PIC.Line (x, y)-(x, y), DCul
            PIC.Refresh
         ElseIf RCNum > 1 And LCNum = 0 Then
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
            PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(x, y), DCul
            PIC.Refresh
         ElseIf LCNum = 1 Then  ' Move all
            AMoveAll = True
         ElseIf LCNum = 2 Then
            CompletePolySingleLine CulNum
            PIC.DrawWidth = 1
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
         End If
         
      End If
End Sub

Public Sub StartEndPolyDoubleAndShadedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim x1 As Single, y1 As Single
   If StartButton = 0 Then
      StartButton = Button
      GetColor Button
      Select Case PolyLineType
      Case PolyDoubleLine1, PolyDoubleLineEnd1, PolyShadedLine1: zspace = 4
      Case PolyDoubleLine2, PolyDoubleLineEnd2, PolyShadedLine2: zspace = 8
      Case PolyDoubleLine3, PolyDoubleLineEnd3, PolyShadedLine3: zspace = 16
      End Select
      PIC.DrawMode = 7
   End If
   
   If Button = vbLeftButton Then LCNum = LCNum + 1
   If Button = vbRightButton Then RCNum = RCNum + 1
   
   If AMoveAll Then
      If StartButton = vbLeftButton Then
         If RCNum = 1 Then
            CompletePolyDoubleLine CulNum
            PIC.DrawWidth = 1
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
            Exit Sub
         End If
      End If
      If StartButton = vbRightButton Then
         If LCNum = 1 Then
            CompletePolyDoubleLine CulNum
            PIC.DrawWidth = 1
            LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
            Exit Sub
         End If
      End If
    End If
   
   If StartButton = vbLeftButton Then
      If LCNum = 1 And RCNum = 0 Then
         AMoveAll = False
         NSTOREXY = 2
         ReDim STOREX(2), STOREY(2)
         STOREX(1) = x: STOREY(1) = y
         STOREX(2) = x: STOREY(2) = y
         PIC.Line (x, y)-(x, y), DCul
         GetParallelCoords zspace, x, y, x, y, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
         PIC.Refresh
      ElseIf LCNum > 1 And RCNum = 0 Then
         NSTOREXY = NSTOREXY + 1
         ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
         STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
         x1 = CSng(STOREX(NSTOREXY - 1))
         y1 = CSng(STOREY(NSTOREXY - 1))
         PIC.Line (x1, y1)-(x, y), DCul
         GetParallelCoords zspace, x1, y1, x, y, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
         PIC.Refresh
      ElseIf RCNum = 1 Then  ' Move all
         AMoveAll = True
      ElseIf RCNum = 2 Then
         CompletePolyDoubleLine CulNum
         PIC.DrawWidth = 1
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
      
   ElseIf StartButton = vbRightButton Then
      If RCNum = 1 And LCNum = 0 Then
         AMoveAll = False
         NSTOREXY = 2
         ReDim STOREX(2), STOREY(2)
         STOREX(1) = x: STOREY(1) = y
         STOREX(2) = x: STOREY(2) = y
         PIC.Line (x, y)-(x, y), DCul
         PIC.Refresh
      ElseIf RCNum > 1 And LCNum = 0 Then
         NSTOREXY = NSTOREXY + 1
         ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
         STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
         PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(x, y), DCul
         PIC.Refresh
      ElseIf LCNum = 1 Then  ' Move all
         AMoveAll = True
      ElseIf LCNum = 2 Then
         CompletePolyDoubleLine CulNum
         PIC.DrawWidth = 1
         LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
      End If
   End If
End Sub

Public Sub StartEndRectangleSingle(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   PIC.DrawWidth = 1
   Select Case RectangleType
   Case RectangleSingle1: PIC.DrawWidth = 1
   Case RectangleSingle2: PIC.DrawWidth = 2
   Case RectangleSingle3: PIC.DrawWidth = 4
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      PIC.Line (x, y)-(x, y), DCul, B
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      Select Case svRectangleType
      Case RectangleShaded1, RectangleShaded2, RectangleShaded3, _
           RectangleFShade, RectangleBShade, RectangleFilled
         CompleteRectangleShaded CulNum
      Case Else
         CompleteRectangleSingle CulNum
      End Select
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If

End Sub
         
Public Sub StartEndRectangleDotted(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   Select Case RectangleType
   Case RectangleDotted1: PIC.DrawWidth = 1
   Case RectangleDotted2: PIC.DrawWidth = 2
   Case RectangleDotted3: PIC.DrawWidth = 4
   End Select
   PIC.DrawStyle = vbDot   ' Doesn't work for DrawWidth > 1
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      PIC.Line (x, y)-(x, y), DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteRectangleDotted CulNum
      PIC.DrawWidth = 1
      PIC.DrawStyle = vbSolid
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndRectangleDouble(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim ND As Long
   PIC.DrawMode = 7
   Select Case RectangleType
   Case RectangleDouble1: ND = 2
   Case RectangleDouble2: ND = 4
   Case RectangleDouble3: ND = 8
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 4
      ReDim STOREX(4), STOREY(4)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      STOREX(3) = x + ND: STOREY(3) = y + ND
      STOREX(4) = x - ND: STOREY(4) = y - ND
      PIC.Line (x, y)-(x, y), DCul, B
      PIC.Line (x + ND, y + ND)-(x - ND, y - ND), DCul, B
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteRectangleDouble CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndCirllipseSDD(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Single, Dotted, Double
Dim NSPACE As Long
   PIC.DrawMode = 7
   PIC.DrawWidth = 1
   Select Case CirllipseType
   Case CirllipseSingle1, CirllipseDotted1: PIC.DrawWidth = 1
   Case CirllipseSingle2, CirllipseDotted2: PIC.DrawWidth = 2
   Case CirllipseSingle3, CirllipseDotted3: PIC.DrawWidth = 3
   End Select
   Select Case CirllipseType
   Case CirllipseDotted1, CirllipseDotted2, CirllipseDotted3
      PIC.DrawStyle = vbDot   ' Doesn't work for DrawWidth > 1
   End Select
   NSPACE = 0
   Select Case CirllipseType
   Case CirllipseDouble1: NSPACE = 2
   Case CirllipseDouble2: NSPACE = 4
   Case CirllipseDouble3: NSPACE = 8
   End Select
   
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      ixc = x: iyc = y
      EvalZradZratio x, y  'ixc,iyc public
      PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
      If NSPACE > 0 Then
         If zrad - NSPACE > 0 Then
            PIC.Circle (ixc, iyc), zrad - NSPACE, DCul, , , zratio
         End If
      End If
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      Select Case CirllipseType
      Case CirllipseShaded1, CirllipseShaded2, CirllipseShaded3
         CompleteCirllipseShaded CulNum
      Case Else
         CompleteCirllipseSDD CulNum
      End Select
      PIC.DrawWidth = 1
      PIC.DrawStyle = vbSolid
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndConeAndTube(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   PIC.DrawWidth = 1
   NTube = 0
   LCNum = LCNum + 1
   If LCNum = 1 Then ' Draw base circle
      GetColor Button
      NSTOREXY = 2
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = STOREX(1): STOREY(2) = STOREY(1)
      zrad = 0
      PIC.Refresh
   ElseIf LCNum = 2 Then  ' Draw axial line
      ReDim Preserve STOREX(3), STOREY(3)
      NSTOREXY = 3
      STOREX(3) = x: STOREY(3) = y
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      PIC.Refresh
   ElseIf LCNum = 3 Then  ' Move cone
   ElseIf LCNum = 4 Then
      If ToolType = Cone Then
         CompleteCone CulNum
      ElseIf ToolType = Tube Then
         CompleteTube CulNum
      ElseIf ToolType = Bullet Then
         CompleteBullet CulNum
      End If
      PIC.DrawWidth = 1
      PIC.DrawStyle = vbSolid
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndJunction(PIC As PictureBox, Button As Integer, x As Single, y As Single)
'Public zspace,zTL
   zTL = 10 ' Default side piece length
   PIC.DrawMode = 7
   Select Case JunctionType
   Case TPiece1, Cross1, Corner1: zspace = 4
   Case TPiece2, Cross2, Corner2: zspace = 8
   Case TPiece3, Cross3, Corner3: zspace = 16
   End Select
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 2 '8
      ReDim STOREX(NSTOREXY), STOREY(NSTOREXY)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x + 10: STOREY(2) = y
      XT(1) = x: YT(1) = y
      XT(2) = x + 10: YT(2) = y
      Get12TPieces XT(1), YT(1), XT(2), YT(2), XT(), YT()
      Select Case JunctionType
      Case TPiece1, TPiece2, TPiece3: PicDrawTPiece PIC
      Case Cross1, Cross2, Cross3: PicDrawCross PIC
      Case Corner1, Corner2, Corner3: PicDrawCorner PIC
      End Select
   ElseIf LCNum = 2 Then ' Move all
   ElseIf LCNum = 3 Then
      CompleteJunction CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndArc(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Public zSA As Single, zEA As Single
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      ReDim STOREX(2), STOREY(2)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      ixc = x: iyc = y
      EvalZradZratio x, y  'ixc,iyc public
      Select Case ArcType
      'Case 0: zSA = 0: zEA = 2 * pi#
      Case 0: zSA = pi# / 2: zEA = pi#          ' TL
      Case 1: zSA = pi#: zEA = 3 * pi# / 2      ' BL
      Case 2: zSA = 0: zEA = pi# / 2            ' TR
      Case 3: zSA = 3 * pi# / 2: zEA = 2 * pi#  ' BR
      Case 4: zSA = pi# / 2: zEA = 3 * pi# / 2  ' LS
      Case 5: zSA = 3 * pi# / 2: zEA = pi# / 2  ' RS
      Case 6: zSA = 0: zEA = pi#                ' TS
      Case 7: zSA = pi#: zEA = 2 * pi#          ' BS
      Case 8: zSA = 0: zEA = 3 * pi# / 2        ' BRX
      Case 9: zSA = pi# / 2: zEA = 2 * pi#     ' TRX
      Case 10: zSA = 3 * pi# / 2: zEA = pi#     ' BLX
      Case 11: zSA = pi#: zEA = pi# / 2         ' TLX
      End Select
      PIC.Circle (ixc, iyc), zrad, DCul, zSA, zEA, zratio
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move all
   ElseIf LCNum = 3 Then
      CompleteArc CulNum     ' NB:  NQ = ArcType + 1
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndShapeA(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 4
      ReDim STOREX(4), STOREY(4)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      STOREX(3) = x: STOREY(3) = y
      STOREX(4) = x: STOREY(4) = y
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteShapeA CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndShapeB(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Dumbell
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      NSTOREXY = 4
      ReDim STOREX(4), STOREY(4)
      STOREX(1) = x: STOREY(1) = y
      STOREX(2) = x: STOREY(2) = y
      STOREX(3) = x + 1: STOREY(3) = y
      STOREX(4) = x + 1: STOREY(4) = y
      zrad = 1
      PIC.Line (x, y)-(x, y), DCul
      PIC.Circle (x, y), zrad, DCul
      PIC.Circle (x + 1, y), zrad, DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteShapeB CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndRadials(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim k As Long
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      Select Case RadialType
      Case RSpokes: NSTOREXY = RadialRep(RadialType + 1) + 1
      Case RStars: NSTOREXY = 2 * RadialRep(RadialType + 1) + 1
      Case RRadCircs: NSTOREXY = RadialRep(RadialType + 1) + 1
      Case RPolygons: NSTOREXY = RadialRep(RadialType + 1) + 1
      Case RTeeth: NSTOREXY = RadialRep(RadialType + 1) + 1
      End Select
      zangle = 2 * pi# / (NSTOREXY - 1)
      ' ie 2*pi# / number of spokes
      
      ReDim STOREX(NSTOREXY), STOREY(NSTOREXY)
      For k = 1 To NSTOREXY
         STOREX(k) = x: STOREY(k) = y
      Next k
      zrad = 0
      zrad2 = 0
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteRadials CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub StartEndTree(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim k As Long
Dim Level As Long
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      Select Case TreeType
      Case Tree1
         Select Case BushSize(1)
         Case 1: Level = 2
                 ystep(1) = -4
         Case 2: Level = 3
                 ystep(1) = -4
         Case 3: Level = 3
                 ystep(1) = -5
         End Select
      Case Tree2
         Select Case BushSize(2)
         Case 1: Level = 2
                 ystep(2) = -5
                 ymul(2) = 1
         Case 2: Level = 3
                 ystep(2) = -3
                 ymul(2) = 1
         Case 3: Level = 3
                 ystep(2) = -5
                 ymul(2) = 1.01
         End Select
      Case Tree3
         Select Case BushSize(3)
         Case 1: Level = 2
                 ystep(3) = -2
         Case 2: Level = 3
                 ystep(3) = -2
         Case 3: Level = 3
                 ystep(3) = -3
         End Select
      End Select
      
      DisplayTree PIC, Level, TreeType + 1, x, y
   
   ElseIf LCNum = 2 Then
      CompleteBush CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Private Sub DisplayTree(PIC As PictureBox, Level As Long, IT As Long, ByVal Xs As Single, ByVal Ys As Single)
Dim k As Long
Dim LL As Long
Dim zTmp As Single
Dim zCosAngP As Single
Dim zSinAngP As Single
Dim zCosAngN As Single
Dim zSinAngN As Single
Dim XStepSave As Single, YStepSave As Single
Dim NumBrackets As Long
Dim zSaveState(4000) As Single
Dim ExpandedAxiom$
         
   YTreeMin = 10000
   YTreeMax = -10000
   ExpandedAxiom$ = Axiom$(IT)
   For k = 1 To Level
      ExpandedAxiom$ = Replace(ExpandedAxiom$, "F", PAxiom$(IT))
   Next k
   LL = Len(ExpandedAxiom$)
   ' Transfer characters to byte array
   ReDim bTreeArray(LL)
   CopyMemory bTreeArray(1), ByVal ExpandedAxiom$, LL
   ExpandedAxiom$ = ""
   
   NumBrackets = 0
   zCosAngP = Cos(zAngP(IT) * d2r#)
   zSinAngP = Sin(zAngP(IT) * d2r#)
   zCosAngN = Cos(zAngN(IT) * d2r#)
   zSinAngN = Sin(zAngN(IT) * d2r#)
   
   ' Save start steps for further drawing
   XStepSave = xstep(IT)
   YStepSave = ystep(IT)
   
   NSTOREXY = 1
   ReDim STOREX(1), STOREY(1)
   STOREX(1) = CInt(Xs)
   STOREY(1) = CInt(Ys)
   
   For k = 1 To LL
      
      Select Case bTreeArray(k)
      Case 70, 71 ' F Pen Down, Advance:  G Pen Up, Advance
         
         If bTreeArray(k) = 70 Then   ' F draw
            NSTOREXY = NSTOREXY + 1
            ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
            STOREX(NSTOREXY) = CInt(Xs + xstep(IT))
            STOREY(NSTOREXY) = CInt(Ys + ystep(IT))
            PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(STOREX(NSTOREXY), STOREY(NSTOREXY)), DCul
            
            If STOREY(NSTOREXY) < YTreeMin Then YTreeMin = STOREY(NSTOREXY)
            If STOREY(NSTOREXY) > YTreeMax Then YTreeMax = STOREY(NSTOREXY)
            ' Shell bands
            'PIC.PSet (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1)), TreeColor
            
            'PIC.Refresh
         End If
         'Advance
         Xs = Xs + xstep(IT): Ys = Ys + ystep(IT)
         xstep(IT) = xstep(IT) * xmul(IT)
         ystep(IT) = ystep(IT) * ymul(IT)
           
      Case 91  ' [  Push turtle state
         
         NumBrackets = NumBrackets + 1
         zSaveState(NumBrackets) = Xs
         NumBrackets = NumBrackets + 1
         zSaveState(NumBrackets) = Ys
         NumBrackets = NumBrackets + 1
         zSaveState(NumBrackets) = xstep(IT) '''''
         NumBrackets = NumBrackets + 1
         zSaveState(NumBrackets) = ystep(IT) ''''''
      
      Case 93  ' ] Pop turtle state
         
         ystep(IT) = zSaveState(NumBrackets)  ''''''
         NumBrackets = NumBrackets - 1
         xstep(IT) = zSaveState(NumBrackets)  ''''''
         NumBrackets = NumBrackets - 1
         Ys = zSaveState(NumBrackets)
         NumBrackets = NumBrackets - 1
         Xs = zSaveState(NumBrackets)
         NumBrackets = NumBrackets - 1
           
      Case 43   ' + turn left
   
         zTmp = xstep(IT)
         xstep(IT) = zCosAngP * zTmp - zSinAngP * ystep(IT)
         ystep(IT) = zSinAngP * zTmp + zCosAngP * ystep(IT)
           
      Case 45   ' - turn right
       
         zTmp = xstep(IT)
         xstep(IT) = zCosAngN * zTmp - zSinAngN * ystep(IT)
         ystep(IT) = zSinAngN * zTmp + zCosAngN * ystep(IT)
       
      'Case Else   'Ignore
      
      End Select
      
   Next k
   PIC.Refresh
   ' Restore start steps
   xstep(IT) = XStepSave
   ystep(IT) = YStepSave
   Erase zSaveState(), bTreeArray()
End Sub

Public Sub StartEndArrow(PIC As PictureBox, Button As Integer, x As Single, y As Single)
'Public zarrang As Single
'Public zarrlen As Single
Dim zang1 As Single
   zarrang = 0.333
   zarrlen = 10
   PIC.DrawMode = 7
   LCNum = LCNum + 1
   If LCNum = 1 Then
      GetColor Button
      ReDim STOREX(5), STOREY(5)
      STOREX(1) = x
      STOREY(1) = y
      STOREX(2) = x
      STOREY(2) = y
      zang1 = zATan2((STOREY(2) - STOREY(1)), (STOREX(2) - STOREX(1)))
      xd1 = zarrlen * Cos(zang1 - zarrang)
      yd1 = zarrlen * Sin(zang1 - zarrang)
      xd2 = zarrlen * Sin(pi# / 2 - zang1 - zarrang)
      yd2 = zarrlen * Cos(pi# / 2 - zang1 - zarrang)
      STOREX(3) = STOREX(2) - xd1
      STOREY(3) = STOREY(2) - yd1
      STOREX(4) = STOREX(2) - xd2
      STOREY(4) = STOREY(2) - yd2
      STOREX(5) = (STOREX(3) + STOREX(4)) / 2
      STOREY(5) = (STOREY(3) + STOREY(4)) / 2
      DrawArrow PIC
   ElseIf LCNum = 2 Then ' Move
   ElseIf LCNum = 3 Then
      CompleteArrow CulNum
      PIC.DrawWidth = 1
      LCNum = -1   ' Flag to SAVE_CurrentImage & DISPLAY
   End If
End Sub

Public Sub DrawArrow(PIC As PictureBox)
   Select Case ArrowType
   Case ArrSingle
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
   Case ArrFeathered
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(1) - xd1, STOREY(1) - yd1), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(1) - xd2, STOREY(1) - yd2), DCul
   Case ArrTriangle
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(5), STOREY(5)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
   End Select
   PIC.Refresh
End Sub


'>>>>>>>>>>>>  MOVE >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Public Sub MoveFreeDraws(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   If LCNum = 1 Then
      NSTOREXY = NSTOREXY + 1
      ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
      STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
      PIC.Line -(x, y), DCul
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old
      PIC.PSet (STOREX(1), STOREY(1)), DCul
      For NN = 2 To NSTOREXY
         PIC.Line -(STOREX(NN), STOREY(NN)), DCul
      Next NN
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new
      PIC.PSet (STOREX(1), STOREY(1)), DCul
      For NN = 2 To NSTOREXY
         PIC.Line -(STOREX(NN), STOREY(NN)), DCul
      Next NN
   End If
End Sub

Public Sub MoveRibbons(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   If LCNum = 1 Then
      NSTOREXY = NSTOREXY + 1
      ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
      STOREX(NSTOREXY) = x: STOREY(NSTOREXY) = y
      PIC.Line (x + RX1, y + RY1)-(x + RX2, y + RY2), DCul
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old
      For NN = 1 To NSTOREXY
         PIC.Line (STOREX(NN) + RX1, STOREY(NN) + RY1)-(STOREX(NN) + RX2, STOREY(NN) + RY2), DCul
      Next NN
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new
      For NN = 1 To NSTOREXY
         PIC.Line (STOREX(NN) + RX1, STOREY(NN) + RY1)-(STOREX(NN) + RX2, STOREY(NN) + RY2), DCul
      Next NN
   End If
End Sub

Public Sub MoveSprays(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      For NN = 1 To NSTOREXY
         PIC.PSet (STOREX(NN), STOREY(NN)), DCul
      Next NN
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new
      For NN = 1 To NSTOREXY
         PIC.PSet (STOREX(NN), STOREY(NN)), DCul
      Next NN
   End If
End Sub

Public Sub MoveSingleLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new - rotate line about first point
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
   ElseIf LCNum = 2 Then
      ' Move
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      ' Set new position
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new - locate line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
   End If
End Sub

Public Sub MoveDottedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      End Select
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         GetParallelCoords zspace, CSng(STOREX(1)), CSng(STOREY(1)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
         STOREX(3) = ixa: STOREY(3) = iya
         STOREX(4) = ixb: STOREY(4) = iyb
      End Select
      ' Draw new - rotate line about first point
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      End Select
      PIC.Refresh
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      End Select
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         STOREX(3) = STOREX(3) + IncrX
         STOREY(3) = STOREY(3) + IncrY
         STOREX(4) = STOREX(4) + IncrX
         STOREY(4) = STOREY(4) + IncrY
      End Select
      ' Draw new - locate line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      Select Case LineType
      Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
         PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      End Select
   End If
End Sub

Public Sub MoveDoubleAndShadedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      ' & line2
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      If LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
         PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
         PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      End If
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      GetParallelCoords zspace, CSng(STOREX(1)), CSng(STOREY(1)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
      STOREX(3) = ixa: STOREY(3) = iya
      STOREX(4) = ixb: STOREY(4) = iyb
      ' Draw new - rotate both lines about first point (STOREX(1), STOREY(1))
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      If LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
         PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
         PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      End If
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old lines
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      If LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
         PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
         PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      End If
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      For NN = 1 To 4
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new double line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      
      If LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
         PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
         PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      End If
   End If
End Sub

Public Sub MovePolySingleLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   If AMoveAll = False Then
      If LCNum > 0 Or RCNum > 0 Then
         'Rotate last line  NSTOREXY-1, NSTOREXY
         IncrX = x - STOREX(NSTOREXY)
         IncrY = y - STOREY(NSTOREXY)
         ' Clear old
         PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(STOREX(NSTOREXY), STOREY(NSTOREXY)), DCul
         ' Set new position
         STOREX(NSTOREXY) = STOREX(NSTOREXY) + IncrX
         STOREY(NSTOREXY) = STOREY(NSTOREXY) + IncrY
         ' Draw new - rotate line about previous point
         PIC.Line (STOREX(NSTOREXY - 1), STOREY(NSTOREXY - 1))-(STOREX(NSTOREXY), STOREY(NSTOREXY)), DCul
         PIC.Refresh
      End If
   ElseIf AMoveAll And (RCNum = 1 Or LCNum = 1) Then ' MoveAll
      NSTOREXY = UBound(STOREX())
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      ' Clear old
      For NN = 2 To NSTOREXY
         PIC.Line (STOREX(NN - 1), STOREY(NN - 1))-(STOREX(NN), STOREY(NN)), DCul
      Next NN
      ' Set new positions
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new polylines
      For NN = 2 To NSTOREXY
         PIC.Line (STOREX(NN - 1), STOREY(NN - 1))-(STOREX(NN), STOREY(NN)), DCul
      Next NN
      PIC.Refresh
   End If
End Sub

Public Sub MovePolyDoubleAndShadedLines(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long

   If AMoveAll = False Then
      If LCNum > 0 Or RCNum > 0 Then
         'Rotate last line  NSTOREXY-1, NSTOREXY
         IncrX = x - STOREX(NSTOREXY)
         IncrY = y - STOREY(NSTOREXY)
         ' Clear old
         x1 = CSng(STOREX(NSTOREXY - 1))
         y1 = CSng(STOREY(NSTOREXY - 1))
         x2 = CSng(STOREX(NSTOREXY))
         y2 = CSng(STOREY(NSTOREXY))
         PIC.Line (x1, y1)-(x2, y2), DCul
         GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
         ' Set new position
         STOREX(NSTOREXY) = STOREX(NSTOREXY) + IncrX
         STOREY(NSTOREXY) = STOREY(NSTOREXY) + IncrY
         x2 = CSng(STOREX(NSTOREXY))
         y2 = CSng(STOREY(NSTOREXY))
         ' Draw new - rotate line about previous point
         PIC.Line (x1, y1)-(x2, y2), DCul
         GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
         PIC.Refresh
      End If
   'ElseIf AMoveAll And (RCNum = 1 Or LCNum = 1) Then  ' MoveAll
   ElseIf AMoveAll Then  ' MoveAll
      NSTOREXY = UBound(STOREX())
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      ' Clear old
      For NN = 2 To NSTOREXY
         x1 = CSng(STOREX(NN - 1))
         y1 = CSng(STOREY(NN - 1))
         x2 = CSng(STOREX(NN))
         y2 = CSng(STOREY(NN))
         PIC.Line (x1, y1)-(x2, y2), DCul
         GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
      Next NN
      ' Set new positions
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new polylines
      For NN = 2 To NSTOREXY
         x1 = CSng(STOREX(NN - 1))
         y1 = CSng(STOREY(NN - 1))
         x2 = CSng(STOREX(NN))
         y2 = CSng(STOREY(NN))
         PIC.Line (x1, y1)-(x2, y2), DCul
         GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
         PIC.Line (ixa, iya)-(ixb, iyb), DCul
      Next NN
      PIC.Refresh
   End If
End Sub

Public Sub MoveRectangleSingle(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new - resize about first point
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
   ElseIf LCNum = 2 Then
      ' Move
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      ' Set new position
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new - locate line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
   End If
End Sub

Public Sub MoveRectangleDotted(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   MoveRectangleSingle PIC, Button, x, y
End Sub

Public Sub MoveRectangleDouble(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul, B
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      STOREX(4) = STOREX(4) + IncrX
      STOREY(4) = STOREY(4) + IncrY
      ' Draw new - rotate line about first point
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul, B
   ElseIf LCNum = 2 Then
      ' Move
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul, B
      ' Set new position
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      STOREX(3) = STOREX(3) + IncrX
      STOREY(3) = STOREY(3) + IncrY
      STOREX(4) = STOREX(4) + IncrX
      STOREY(4) = STOREY(4) + IncrY
      ' Draw new - locate line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul, B
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul, B
   End If
End Sub

Public Sub MoveCirllipseSDD(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Single, Dotted, Double
Dim NSPACE As Long
   NSPACE = 0
   Select Case CirllipseType
   Case CirllipseDouble1: NSPACE = 2
   Case CirllipseDouble2: NSPACE = 4
   Case CirllipseDouble3: NSPACE = 8
   End Select
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
      If NSPACE <> 0 Then
         If zrad - NSPACE > 0 Then
            PIC.Circle (ixc, iyc), zrad - NSPACE, DCul, , , zratio
         End If
      End If
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Set new cirllipse about ixc,iyc
      EvalZradZratio x, y  'ixc,iyc public
      PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
      If NSPACE <> 0 Then
         If zrad - NSPACE > 0 Then
            PIC.Circle (ixc, iyc), zrad - NSPACE, DCul, , , zratio
         End If
      End If
      PIC.Refresh
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old
      PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
      If NSPACE <> 0 Then
         If zrad - NSPACE > 0 Then
            PIC.Circle (ixc, iyc), zrad - NSPACE, DCul, , , zratio
         End If
      End If
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ixc = STOREX(1)
      iyc = STOREY(1)
      EvalZradZratio x, y  'ixc,iyc public
      ' Draw new cirllipse
      PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
      If NSPACE <> 0 Then
         If zrad - NSPACE > 0 Then
            PIC.Circle (ixc, iyc), zrad - NSPACE, DCul, , , zratio
         End If
      End If
   End If
End Sub

Public Sub MoveConeAndTube(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   If LCNum = 1 Then
      ' Move
      ' Clear old cone or tube base
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new cone or tube base
      zrad = Sqr((STOREX(2) - STOREX(1)) ^ 2 + (STOREY(2) - STOREY(1)) ^ 2)
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then
      NTube = NTube + 1
      ' Clear old cone or tube base & axis
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      If (ToolType = Tube Or ConeType = ConeCross) And NTube > 1 Then
         PIC.Circle (STOREX(3), STOREY(3)), zrad, DCul
      End If
      ' Set new position
      IncrX = x - STOREX(3)
      IncrY = y - STOREY(3)
      STOREX(3) = STOREX(3) + IncrX
      STOREY(3) = STOREY(3) + IncrY
      ' Draw same base cone & new axis
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      If ToolType = Tube Or ConeType = ConeCross Then
         PIC.Circle (STOREX(3), STOREY(3)), zrad, DCul
      End If
   ElseIf LCNum = 3 Then ' Move whole cone or Tube
      ' Clear old cone base & axis
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      If ToolType = Tube Or ConeType = ConeCross Then
         PIC.Circle (STOREX(3), STOREY(3)), zrad, DCul
      End If
      ' Set new position
      IncrX = x - STOREX(3)
      IncrY = y - STOREY(3)
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      STOREX(3) = STOREX(3) + IncrX
      STOREY(3) = STOREY(3) + IncrY
      ' Draw new base cone & new axis
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      If ToolType = Tube Or ConeType = ConeCross Then
         PIC.Circle (STOREX(3), STOREY(3)), zrad, DCul
      End If
   End If
End Sub


Public Sub MoveBullet(PIC As PictureBox, Button As Integer, x As Single, y As Single)
   If LCNum = 1 Then
      ' Move
      ' Clear old bullet base
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw bullet base
      zrad = Sqr((STOREX(2) - STOREX(1)) ^ 2 + (STOREY(2) - STOREY(1)) ^ 2)
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Refresh
   ElseIf LCNum = 2 Then
      ' Clear bullet base & axis
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      ' Set new position
      IncrX = x - STOREX(3)
      IncrY = y - STOREY(3)
      STOREX(3) = STOREX(3) + IncrX
      STOREY(3) = STOREY(3) + IncrY
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
   ElseIf LCNum = 3 Then ' Move bullet
      ' Clear old bullet base
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      ' Set new position
      IncrX = x - STOREX(3)
      IncrY = y - STOREY(3)
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      STOREX(3) = STOREX(3) + IncrX
      STOREY(3) = STOREY(3) + IncrY
      ' Draw new bullet base
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
   End If
End Sub

Public Sub MoveJunction(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim k As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      Select Case JunctionType
      Case TPiece1, TPiece2, TPiece3: PicDrawTPiece PIC
      Case Cross1, Cross2, Cross3: PicDrawCross PIC
      Case Corner1, Corner2, Corner3: PicDrawCorner PIC
      End Select
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      XT(2) = STOREX(2)
      YT(2) = STOREY(2)
      ' Draw new - rotate lines about first point (STOREX(1), STOREY(1))
      Get12TPieces XT(1), YT(1), XT(2), YT(2), XT(), YT()
      Select Case JunctionType
      Case TPiece1, TPiece2, TPiece3: PicDrawTPiece PIC
      Case Cross1, Cross2, Cross3: PicDrawCross PIC
      Case Corner1, Corner2, Corner3: PicDrawCorner PIC
      End Select
   ElseIf LCNum = 2 Then
      ' Move all
      ' Clear old lines
      Select Case JunctionType
      Case TPiece1, TPiece2, TPiece3: PicDrawTPiece PIC
      Case Cross1, Cross2, Cross3: PicDrawCross PIC
      Case Corner1, Corner2, Corner3: PicDrawCorner PIC
      End Select
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      For k = 1 To 12
         XT(k) = XT(k) + IncrX
         YT(k) = YT(k) + IncrY
      Next k
      ' Draw new lines
      Select Case JunctionType
      Case TPiece1, TPiece2, TPiece3: PicDrawTPiece PIC
      Case Cross1, Cross2, Cross3: PicDrawCross PIC
      Case Corner1, Corner2, Corner3: PicDrawCorner PIC
      End Select
   End If
End Sub

Public Sub PicDrawTPiece(PIC As PictureBox)
   PIC.Line (XT(1), YT(1))-(XT(2), YT(2)), DCul
   PIC.Line (XT(3), YT(3))-(XT(7), YT(7)), DCul
   PIC.Line (XT(8), YT(8))-(XT(4), YT(4)), DCul
   PIC.Line (XT(7), YT(7))-(XT(11), YT(11)), DCul
   PIC.Line (XT(8), YT(8))-(XT(12), YT(12)), DCul
   PIC.Refresh
End Sub

Public Sub PicDrawCross(PIC As PictureBox)
   PIC.Line (XT(1), YT(1))-(XT(5), YT(5)), DCul
   PIC.Line (XT(6), YT(6))-(XT(2), YT(2)), DCul
   PIC.Line (XT(3), YT(3))-(XT(7), YT(7)), DCul
   PIC.Line (XT(8), YT(8))-(XT(4), YT(4)), DCul
   PIC.Line (XT(7), YT(7))-(XT(11), YT(11)), DCul
   PIC.Line (XT(8), YT(8))-(XT(12), YT(12)), DCul
   PIC.Line (XT(5), YT(5))-(XT(9), YT(9)), DCul
   PIC.Line (XT(6), YT(6))-(XT(10), YT(10)), DCul
   PIC.Refresh
End Sub

Public Sub PicDrawCorner(PIC As PictureBox)
   PIC.Line (XT(1), YT(1))-(XT(5), YT(5)), DCul
   PIC.Line (XT(5), YT(5))-(XT(9), YT(9)), DCul
   PIC.Line (XT(3), YT(3))-(XT(8), YT(8)), DCul
   PIC.Line (XT(8), YT(8))-(XT(10), YT(10)), DCul
   PIC.Refresh
End Sub

Public Sub MoveArc(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Public zSA, zEA start & end angles
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Circle (ixc, iyc), zrad, DCul, zSA, zEA, zratio
      PIC.Refresh
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Set new cirllipse about ixc,iyc
      EvalZradZratio x, y  'ixc,iyc public
      PIC.Circle (ixc, iyc), zrad, DCul, zSA, zEA, zratio
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old arc
      PIC.Circle (ixc, iyc), zrad, DCul, zSA, zEA, zratio
      PIC.Refresh
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ixc = STOREX(1)
      iyc = STOREY(1)
      EvalZradZratio x, y  'ixc,iyc public
      ' Draw new arc
      PIC.Circle (ixc, iyc), zrad, DCul, zSA, zEA, zratio
   End If
End Sub

Public Sub MoveShapeA(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim dx As Long, dy As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      dx = (STOREX(2) - STOREX(1))
      dy = (STOREY(2) - STOREY(1))
      
      Select Case ShapeType
      Case TShape1 ' Triangle
         zspace = CSng(Sqr(dx * dx + dy * dy))
      Case TShape2 ' Diamond
         zspace = CSng(Sqr(dx * dx + dy * dy))
      Case TShape3 ' Frustrum
         zspace = CSng(Sqr(dx * dx + dy * dy)) / 2
      Case TShape4 ' Wedge1
         zspace = CSng(Sqr(dx * dx + dy * dy)) / 2
      Case TShape5 ' Wedge2
         zspace = -CSng(Sqr(dx * dx + dy * dy)) / 2
      End Select
      
      GetParallelCoords zspace, CSng(STOREX(1)), CSng(STOREY(1)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
      
      Select Case ShapeType
      Case TShape1 ' Triangle
         STOREX(3) = ixa + dx / 2: STOREY(3) = iya + dx / 2
         STOREX(4) = STOREX(3): STOREY(4) = STOREY(3)
      Case TShape2 ' Diamond
         STOREX(3) = ixa: STOREY(3) = iya
         STOREX(4) = ixb: STOREY(4) = iyb
      Case TShape3 ' Frustrum
         STOREX(3) = ixa + dx / 4: STOREY(3) = iya + dy / 4
         STOREX(4) = ixa + 3 * dx / 4: STOREY(4) = iya + 3 * dy / 4
      Case TShape4, TShape5 ' Wedge
         STOREX(3) = ixa + 2 * dx / 3: STOREY(3) = iya + 2 * dy / 3
         STOREX(4) = ixb: STOREY(4) = iyb
      End Select
      
      ' Draw new - rotate both lines about first point (STOREX(1), STOREY(1))
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      PIC.Refresh
      DoEvents
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old lines
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      For NN = 1 To 4
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new double line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Line (STOREX(3), STOREY(3))-(STOREX(4), STOREY(4)), DCul
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(3), STOREY(3)), DCul
      PIC.Line (STOREX(2), STOREY(2))-(STOREX(4), STOREY(4)), DCul
   End If
End Sub

Public Sub MoveShapeB(PIC As PictureBox, Button As Integer, x As Single, y As Single)
' Dumbell
Dim dx As Long, dy As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Circle (STOREX(2), STOREY(2)), zrad, DCul
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      dx = (STOREX(2) - STOREX(1))
      dy = (STOREY(2) - STOREY(1))
      zrad = Sqr(dx * dx + dy * dy) \ 10
      ' Draw new - rotate line about first point
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Circle (STOREX(2), STOREY(2)), zrad, DCul
   ElseIf LCNum = 2 Then
      ' Move
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      ' Clear old
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Circle (STOREX(2), STOREY(2)), zrad, DCul
      ' Set new position
      STOREX(1) = STOREX(1) + IncrX
      STOREY(1) = STOREY(1) + IncrY
      STOREX(2) = STOREX(2) + IncrX
      STOREY(2) = STOREY(2) + IncrY
      ' Draw new - locate line
      PIC.Line (STOREX(1), STOREY(1))-(STOREX(2), STOREY(2)), DCul
      PIC.Circle (STOREX(1), STOREY(1)), zrad, DCul
      PIC.Circle (STOREX(2), STOREY(2)), zrad, DCul
   End If
End Sub

Public Sub MoveRadial(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim k As Long
Dim dx As Long, dy As Long
Dim zA As Single, zrad2 As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long

   If LCNum = 1 Then
      ' Move
      ' Clear old
      
      Select Case RadialType
      Case RSpokes
         ' Clear old
         For k = 2 To NSTOREXY
            PIC.Line (STOREX(1), STOREY(1))-(STOREX(k), STOREY(k)), DCul
         Next k
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         STOREX(2) = STOREX(2) + IncrX
         STOREY(2) = STOREY(2) + IncrY
         dx = (STOREX(2) - STOREX(1))
         dy = (STOREY(2) - STOREY(1))
         zrad = Sqr(dx * dx + dy * dy)
         zA = zATan2(dy, dx)
         ' Draw new
         For k = 2 To NSTOREXY
            STOREX(k) = STOREX(1) + zrad * Cos(zA)
            STOREY(k) = STOREY(1) + zrad * Sin(zA)
            PIC.Line (STOREX(1), STOREY(1))-(STOREX(k), STOREY(k)), DCul
            zA = zA + zangle
         Next k
      
      Case RStars
         ' Clear old
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         STOREX(2) = STOREX(2) + IncrX
         STOREY(2) = STOREY(2) + IncrY
         dx = (STOREX(2) - STOREX(1))
         dy = (STOREY(2) - STOREY(1))
         zrad = Sqr(dx * dx + dy * dy)
         zrad2 = zrad / 5
         zA = zATan2(dy, dx)
         ' Draw new
         STOREX(2) = STOREX(1) + zrad * Cos(zA)
         STOREY(2) = STOREY(1) + zrad * Sin(zA)
         For k = 3 To NSTOREXY
            zA = zA + zangle
            If (k Mod 2) = 1 Then
               STOREX(k) = STOREX(1) + zrad2 * Cos(zA)
               STOREY(k) = STOREY(1) + zrad2 * Sin(zA)
            Else
               STOREX(k) = STOREX(1) + zrad * Cos(zA)
               STOREY(k) = STOREY(1) + zrad * Sin(zA)
            End If
            PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
      
      Case RRadCircs
         ' Clear old
         zrad2 = zrad / 5
         For k = 2 To NSTOREXY
            PIC.Circle (STOREX(k), STOREY(k)), zrad2, DCul
         Next k
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         STOREX(2) = STOREX(2) + IncrX
         STOREY(2) = STOREY(2) + IncrY
         dx = (STOREX(2) - STOREX(1))
         dy = (STOREY(2) - STOREY(1))
         zrad = Sqr(dx * dx + dy * dy)
         zrad2 = zrad / 5
         zA = zATan2(dy, dx)
         ' Draw new
         For k = 2 To NSTOREXY
            STOREX(k) = STOREX(1) + zrad * Cos(zA)
            STOREY(k) = STOREY(1) + zrad * Sin(zA)
            PIC.Circle (STOREX(k), STOREY(k)), zrad2, DCul
            zA = zA + zangle
         Next k
      Case RPolygons
         ' Clear old
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         STOREX(2) = STOREX(2) + IncrX
         STOREY(2) = STOREY(2) + IncrY
         dx = (STOREX(2) - STOREX(1))
         dy = (STOREY(2) - STOREY(1))
         zrad = Sqr(dx * dx + dy * dy)
         zA = zATan2(dy, dx)
         ' Draw new
         STOREX(2) = STOREX(1) + zrad * Cos(zA)
         STOREY(2) = STOREY(1) + zrad * Sin(zA)
         For k = 3 To NSTOREXY
            zA = zA + zangle
            STOREX(k) = STOREX(1) + zrad * Cos(zA)
            STOREY(k) = STOREY(1) + zrad * Sin(zA)
            PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
      
      Case RTeeth
         ' Clear old
         For k = 3 To NSTOREXY
            If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(k - 1)), CSng(STOREY(k - 1)), CSng(STOREX(k)), CSng(STOREY(k)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(k), STOREY(k)), DCul
            Else
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
            End If
         Next k
         If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(NSTOREXY)), CSng(STOREY(NSTOREXY)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(2), STOREY(2)), DCul
         Else
            PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         End If
         
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         STOREX(2) = STOREX(2) + IncrX
         STOREY(2) = STOREY(2) + IncrY
         dx = (STOREX(2) - STOREX(1))
         dy = (STOREY(2) - STOREY(1))
         zrad = Sqr(dx * dx + dy * dy)
         zA = zATan2(dy, dx)
         ' Draw new
         STOREX(2) = STOREX(1) + zrad * Cos(zA)
         STOREY(2) = STOREY(1) + zrad * Sin(zA)
         For k = 3 To NSTOREXY
            zA = zA + zangle
            STOREX(k) = STOREX(1) + zrad * Cos(zA)
            STOREY(k) = STOREY(1) + zrad * Sin(zA)
            If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(k - 1)), CSng(STOREY(k - 1)), CSng(STOREX(k)), CSng(STOREY(k)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(k), STOREY(k)), DCul
            Else
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
            End If
         Next k
         If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(NSTOREXY)), CSng(STOREY(NSTOREXY)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(2), STOREY(2)), DCul
         Else
            PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         End If
      
      End Select

   ElseIf LCNum = 2 Then
      ' Move all
      ' Clear old
      Select Case RadialType
      Case RSpokes
         For k = 1 To NSTOREXY
            PIC.Line (STOREX(1), STOREY(1))-(STOREX(k), STOREY(k)), DCul
         Next k
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         For k = 1 To NSTOREXY
            STOREX(k) = STOREX(k) + IncrX
            STOREY(k) = STOREY(k) + IncrY
         Next k
         For k = 1 To NSTOREXY
            PIC.Line (STOREX(1), STOREY(1))-(STOREX(k), STOREY(k)), DCul
         Next k
      Case RStars
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         For k = 1 To NSTOREXY
            STOREX(k) = STOREX(k) + IncrX
            STOREY(k) = STOREY(k) + IncrY
         Next k
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
      
      Case RRadCircs
         zrad2 = zrad / 5
         For k = 2 To NSTOREXY
            PIC.Circle (STOREX(k), STOREY(k)), zrad2, DCul
         Next k
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         For k = 1 To NSTOREXY
            STOREX(k) = STOREX(k) + IncrX
            STOREY(k) = STOREY(k) + IncrY
         Next k
         For k = 2 To NSTOREXY
            PIC.Circle (STOREX(k), STOREY(k)), zrad2, DCul
         Next k
      Case RPolygons
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         For k = 1 To NSTOREXY
            STOREX(k) = STOREX(k) + IncrX
            STOREY(k) = STOREY(k) + IncrY
         Next k
         For k = 2 To NSTOREXY - 1
            PIC.Line (STOREX(k), STOREY(k))-(STOREX(k + 1), STOREY(k + 1)), DCul
         Next k
         PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
      Case RTeeth
         For k = 3 To NSTOREXY
            If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(k - 1)), CSng(STOREY(k - 1)), CSng(STOREX(k)), CSng(STOREY(k)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(k), STOREY(k)), DCul
            Else
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
            End If
         Next k
         If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(NSTOREXY)), CSng(STOREY(NSTOREXY)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(2), STOREY(2)), DCul
         Else
            PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         End If
         ' Set new position
         IncrX = x - STOREX(2)
         IncrY = y - STOREY(2)
         For k = 1 To NSTOREXY
            STOREX(k) = STOREX(k) + IncrX
            STOREY(k) = STOREY(k) + IncrY
         Next k
         For k = 3 To NSTOREXY
            If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(k - 1)), CSng(STOREY(k - 1)), CSng(STOREX(k)), CSng(STOREY(k)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(k), STOREY(k)), DCul
            Else
               PIC.Line (STOREX(k - 1), STOREY(k - 1))-(STOREX(k), STOREY(k)), DCul
            End If
         Next k
         If (k Mod 2) = 0 Then
               GetParallelCoords zrad / 4, CSng(STOREX(NSTOREXY)), CSng(STOREY(NSTOREXY)), CSng(STOREX(2)), CSng(STOREY(2)), ixa, iya, ixb, iyb
               PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(ixa, iya), DCul
               PIC.Line (ixa, iya)-(ixb, iyb), DCul
               PIC.Line (ixb, iyb)-(STOREX(2), STOREY(2)), DCul
         Else
            PIC.Line (STOREX(NSTOREXY), STOREY(NSTOREXY))-(STOREX(2), STOREY(2)), DCul
         End If
      End Select
   
   End If
End Sub

Public Sub MoveTree(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   PIC.DrawMode = 7
   If LCNum = 1 Then
      ' Move
      ' Clear old
      For NN = 2 To NSTOREXY
         PIC.Line (STOREX(NN - 1), STOREY(NN - 1))- _
         (STOREX(NN), STOREY(NN)), DCul
      Next NN
      ' Set new position
      IncrX = x - STOREX(1)
      IncrY = y - STOREY(1)
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new
      For NN = 2 To NSTOREXY
         PIC.Line (STOREX(NN - 1), STOREY(NN - 1))- _
         (STOREX(NN), STOREY(NN)), DCul
      Next NN
      PIC.Refresh
   End If
End Sub

Public Sub MoveArrow(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim k As Long
Dim zang1 As Single
   If LCNum = 1 Then
      ' Clear old
      DrawArrow PIC
      ' New end point - rotate
      STOREX(2) = x
      STOREY(2) = y
      zang1 = zATan2((STOREY(2) - STOREY(1)), (STOREX(2) - STOREX(1)))
      xd1 = zarrlen * Cos(zang1 - zarrang)
      yd1 = zarrlen * Sin(zang1 - zarrang)
      xd2 = zarrlen * Sin(pi# / 2 - zang1 - zarrang)
      yd2 = zarrlen * Cos(pi# / 2 - zang1 - zarrang)
      STOREX(3) = STOREX(2) - xd1
      STOREY(3) = STOREY(2) - yd1
      STOREX(4) = STOREX(2) - xd2
      STOREY(4) = STOREY(2) - yd2
      STOREX(5) = (STOREX(3) + STOREX(4)) / 2
      STOREY(5) = (STOREY(3) + STOREY(4)) / 2
      DrawArrow PIC
   ElseIf LCNum = 2 Then
      ' Move
      ' Clear old
      DrawArrow PIC
      ' Set new position
      IncrX = x - STOREX(2)
      IncrY = y - STOREY(2)
      For k = 1 To 5
         STOREX(k) = STOREX(k) + IncrX
         STOREY(k) = STOREY(k) + IncrY
      Next k
      ' Draw new location
      DrawArrow PIC
   End If
End Sub


Public Sub MoveText(PIC As PictureBox, Button As Integer, x As Single, y As Single)
Dim NN As Long
   If LCNum = 1 Then
      ' Move
      ' Clear old
      For NN = 1 To NSTOREXY
         SetPixelV PIC.hDC, STOREX(NN), STOREY(NN), DCul
      Next NN
      ' Set new position
      IncrX = x - STOREX(NSTOREXY)
      IncrY = y - STOREY(NSTOREXY)
      For NN = 1 To NSTOREXY
         STOREX(NN) = STOREX(NN) + IncrX
         STOREY(NN) = STOREY(NN) + IncrY
      Next NN
      ' Draw new
      For NN = 1 To NSTOREXY
         SetPixelV PIC.hDC, STOREX(NN), STOREY(NN), DCul
      Next NN
      PIC.Refresh
   End If
End Sub

