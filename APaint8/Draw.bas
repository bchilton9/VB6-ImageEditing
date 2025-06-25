Attribute VB_Name = "Draw"
' Draw.bas

Option Explicit
Option Base 1

Public Instruc$()

' For Fillers:-
Private bMark() As Byte
Private ix2() As Long 'Integer
Private iy2() As Long 'Integer
Private nspan As Long  ' 16 or 64
' For shaded fil1
Public ixpmin As Long, ixpmax As Long
Public iypmin As Long, iypmax As Long
Public zdc As Single

' For double line shading:-
Private Nutt As Long, Nutu As Long
Private ux() As Long, uy() As Long
' For double line shading, shapes & radials:-
Private TX() As Long, TY() As Long

Private cx() As Long, cy() As Long ' Curvy line points
Private dx() As Long, dy() As Long ' Curvy line points

Dim NSX As Long, NSY As Long     ' Adjust PIC & bArray() coords
Dim NN As Long                   ' General counter
Dim NNN As Long                  ' General counter
Dim k As Long                    ' General counter
Dim ND As Long                   ' Temp var
Dim zdeltacn As Single  '= [SelRightCulNum - SelLeftCulNum] / span
Dim zCul As Single      '= SelLeftCulNum or SelRightCulNum
'Dim CulNum As Long

Public Sub FillInstrucLabels()
   ReDim Instruc$(0 To 40, 0 To 22)
   For k = Dot1 To Dot3
      Instruc$(Brush, k) = "LC/RC or [BD - M - BU (No keys)]"
   Next k
   For k = FreeDraw1 To FRibbon3
      Instruc$(Brush, k) = "LC/RC - D - C - M - C"
   Next k
   For k = Dots1 To Diamonds3
      Instruc$(Spray, k) = "LC/RC - M - C"
   Next k
   For k = SingleLine1 To ShadedLine3
      Instruc$(ALine, k) = "LC/RC - D - C - M - C"
   Next k
   For k = PolySingleLine1 To PolyShadedLine3
      Instruc$(PolyLine, k) = "LC/RC - D - LC/RC - D - LC/RC ---- RC//LC - M - C"
   Next k
   For k = CurvySingleLine1 To CurvyShadedLine3
      Instruc$(CurvyLine, k) = "LC/RC - D - LC/RC - D - LC/RC ---- RC//LC - M - C"
   Next k
   For k = RectangleSingle1 To RectangleFilled
      Instruc$(Rectangle, k) = "LC/RC - D - C - M - C"
   Next k
   For k = CirllipseSingle1 To CirllipseShaded3
      Instruc$(Cirllipse, k) = "LC/RC - D - C - M - C"
   Next k
   For k = ConeOutline To ConeCross
      Instruc$(Cone, k) = "LC/RC - D - C - D - C - M - C"
   Next k
   For k = TubeOutLine To TubeCShade
      Instruc$(Tube, k) = "LC/RC - D - C - D - C - M - C"
   Next k
   For k = BulletOutLine To BulletCShade
      Instruc$(Bullet, k) = "LC/RC - D - C - D - C - M - C"
   Next k
   For k = TPiece1 To Corner3
      Instruc$(Junction, k) = "LC/RC - D - C - M - C"
   Next k
   For k = ArcFull To ArcTLX
      Instruc$(Arc, k) = "LC/RC - D - C - M - C"
   Next k
   For k = TShape1 To TShape6
      Instruc$(Shape, k) = "LC/RC - D - C - M - C"
   Next k
   For k = RSpokes To RTeeth
      Instruc$(Radial, k) = "LC/RC - D - C - M - C"
   Next k
   For k = Fill1 To Fill22
      Instruc$(AFill, k) = "LC/RC"
   Next k
   For k = Tree1 To Tree3
      Instruc$(Tree, k) = "LC/RC - M - C"
   Next k
   For k = ArrSingle To ArrTriangle
      Instruc$(Arrow, k) = "LC/RC - D - C - M - C"
   Next k
   
   
   Instruc$(AText, 0) = "LC/RC on button - Get Text - M - C"
   Instruc$(SelR, 0) = "LC - D - C - M - C"
   Instruc$(SelC, 0) = "LC - D - C - M - C"
   Instruc$(SelE, 0) = "LC - D - C - M - C"
   Instruc$(SelL, 0) = "LC - D - C - M - C"
   Instruc$(Desel, 0) = ""
   Instruc$(SCopyPaste, 0) = "C - M - C"
   Instruc$(SCopy, 0) = "Copy selection"
   Instruc$(SCut, 0) = "Copy && clear selection"
   Instruc$(SReflectLR, 0) = "Copy LR reflected selection"
   Instruc$(SReflectUD, 0) = "Copy UD reflected selection"
   Instruc$(SRotate, 0) = "Copy rotated circular selection"
   Instruc$(SPaste, 0) = "LC - M - C"
   Instruc$(SClear, 0) = ""
   Instruc$(Rot90, 0) = "Rotate image or selection by 90 degrees"
   Instruc$(Mix, 0) = "Mix colors over image or selection"
   Instruc$(Thicken, 0) = "Thicken objects for image or selection"
   Instruc$(Pepper, 0) = "LC/RC on button Pepper image or selection"
   Instruc$(LRColor, 0) = "Replace Left by Right color for image or selection"
   Instruc$(Measure, 0) = "BD - M - BU (No keys)"
   Instruc$(Pick, 0) = "LC or RC on image to select color"

   Instruc$(Smooth1, 0) = "BD - M - BU (No keys)"
   Instruc$(Smooth2, 0) = "BD - M - BU (No keys)"
   Instruc$(Smooth4, 0) = "BD - M - BU (No keys)"
End Sub

Public Sub ShowInstructions(Index As Integer)
   With Form1
      Select Case Index  'ToolType
      Case Brush
         .LabDrawInstructions = Instruc$(Brush, BrushType)
         .fraDrawInstr.Caption = "Draw instructions for Brushes"
      Case Spray
         .LabDrawInstructions = Instruc$(Spray, SprayType)
         .fraDrawInstr.Caption = "Draw instructions for Sprays"
      Case ALine
         .LabDrawInstructions = Instruc$(ALine, LineType)
         .fraDrawInstr.Caption = "Draw instructions for Lines"
      Case PolyLine
         .LabDrawInstructions = Instruc$(PolyLine, PolyLineType)
         .fraDrawInstr.Caption = "Draw instructions for PolyLines"
      Case CurvyLine
         .LabDrawInstructions = Instruc$(CurvyLine, CurvyLineType)
         .fraDrawInstr.Caption = "Draw instructions for CurvyLines"
      Case Rectangle
         .LabDrawInstructions = Instruc$(Rectangle, RectangleType)
         .fraDrawInstr.Caption = "Draw instructions for Rectangles"
      Case Cirllipse
         .LabDrawInstructions = Instruc$(Cirllipse, CirllipseType)
         .fraDrawInstr.Caption = "Draw instructions for Cirllipses"
      Case Cone
         .LabDrawInstructions = Instruc$(Cone, ConeType)
         .fraDrawInstr.Caption = "Draw instructions for Cones"
      Case Tube
         .LabDrawInstructions = Instruc$(Tube, TubeType)
         .fraDrawInstr.Caption = "Draw instructions for Tubes"
      Case Bullet
         .LabDrawInstructions = Instruc$(Bullet, BulletType)
         .fraDrawInstr.Caption = "Draw instructions for Bullets"
      Case Junction
         .LabDrawInstructions = Instruc$(Junction, JunctionType)
         .fraDrawInstr.Caption = "Draw instructions for Junctions"
      Case Arc
         .LabDrawInstructions = Instruc$(Arc, ArcType)
         .fraDrawInstr.Caption = "Draw instructions for Arcs"
      Case Shape
         .LabDrawInstructions = Instruc$(Shape, ShapeType)
         .fraDrawInstr.Caption = "Draw instructions for Shapes"
      Case Radial
         .LabDrawInstructions = Instruc$(Radial, RadialType)
         .fraDrawInstr.Caption = "Draw instructions for Radials"
      Case AFill
         .LabDrawInstructions = Instruc$(AFill, FillType)
         .fraDrawInstr.Caption = "Draw instructions for Fills"
      Case Tree
         .LabDrawInstructions = Instruc$(Tree, 0)
         .fraDrawInstr.Caption = "Draw instructions for Bushes"
      Case Arrow
         .LabDrawInstructions = Instruc$(Arrow, ArrowType)
         .fraDrawInstr.Caption = "Draw instructions for Arrows"
      
      Case AText
         .LabDrawInstructions = Instruc$(AText, 0)
         .fraDrawInstr.Caption = "Draw instructions for Text"
      
      Case SelR
         .LabDrawInstructions = Instruc$(SelR, 0)
         .fraDrawInstr.Caption = "Draw instructions for Select rectangle"
      Case SelC
         .LabDrawInstructions = Instruc$(SelC, 0)
         .fraDrawInstr.Caption = "Draw instructions for Select circle"
      Case SelE
         .LabDrawInstructions = Instruc$(SelE, 0)
         .fraDrawInstr.Caption = "Draw instructions for Select ellipse"
      Case SelL
         .LabDrawInstructions = Instruc$(SelL, 0)
         .fraDrawInstr.Caption = "Draw instructions for Select lasso"
      Case Desel
         .LabDrawInstructions = ""
         .fraDrawInstr.Caption = "Deselect"
      Case SCopyPaste
         .LabDrawInstructions = Instruc$(SCopyPaste, 0)
         .fraDrawInstr.Caption = "Copy && Paste selection"
      Case SCopy
         .LabDrawInstructions = Instruc$(SCopy, 0)
         .fraDrawInstr.Caption = "Copy action"
      Case SCut
         .LabDrawInstructions = Instruc$(SCut, 0)
         .fraDrawInstr.Caption = "Cut action"
      Case SReflectLR
         .LabDrawInstructions = Instruc$(SReflectLR, 0)
         .fraDrawInstr.Caption = "Left-Right reflection action"
      Case SReflectUD
         .LabDrawInstructions = Instruc$(SReflectUD, 0)
         .fraDrawInstr.Caption = "Up-Down reflection action"
      Case SRotate
         .LabDrawInstructions = Instruc$(SRotate, 0)
         .fraDrawInstr.Caption = "Rotation action"
      Case SPaste
         .LabDrawInstructions = Instruc$(SPaste, 0)
         .fraDrawInstr.Caption = "Instructions for Pasting selection"
      Case SClear
         .LabDrawInstructions = ""
         .fraDrawInstr.Caption = "Clear selection"
      Case Rot90
         .LabDrawInstructions = Instruc$(Rot90, 0)
         .fraDrawInstr.Caption = "Rotate 90 action"
      Case Mix
         .LabDrawInstructions = Instruc$(Mix, 0)
         .fraDrawInstr.Caption = "Mix up colors"
      Case Thicken
         .LabDrawInstructions = Instruc$(Thicken, 0)
         .fraDrawInstr.Caption = "Thicken pixels"
      Case Pepper
         .LabDrawInstructions = Instruc$(Pepper, 0)
         .fraDrawInstr.Caption = "Pepper with Left or Right color"
      Case LRColor
         .LabDrawInstructions = Instruc$(LRColor, 0)
         .fraDrawInstr.Caption = "Color replace"
      Case Measure
         .LabDrawInstructions = Instruc$(Measure, 0)
         .fraDrawInstr.Caption = "Measure angles && lengths"
      Case Pick
         .LabDrawInstructions = Instruc$(Pick, 0)
         .fraDrawInstr.Caption = "Color picker"
      Case Smooth1
         .LabDrawInstructions = Instruc$(Smooth1, 0)
         .fraDrawInstr.Caption = "Smooth Small area"
      Case Smooth2
         .LabDrawInstructions = Instruc$(Smooth2, 0)
         .fraDrawInstr.Caption = "Smooth Medium area"
      Case Smooth4
         .LabDrawInstructions = Instruc$(Smooth4, 0)
         .fraDrawInstr.Caption = "Smooth Large area"
      
      Case Else
         .LabDrawInstructions = "NOT DONE YET"
      End Select
   End With
End Sub

Public Sub Getpx1py1(sx As Long, nx As Long, sy As Long, ny As Long)
'Public px1 As Long, py1 As Long
   px1 = sx + 1 - nx
   If px1 < 1 Then px1 = 1
   If px1 > canvasW Then px1 = canvasW
   py1 = canvasH - sy - ny
   If py1 < 1 Then py1 = 1
   If py1 > canvasH Then py1 = canvasH
End Sub

Public Sub Getpx2py2(sx As Long, nx As Long, sy As Long, ny As Long)
'Public px2 As Long, py2 As Long
   px2 = sx + 1 - nx
   If px2 < 1 Then px2 = 1
   If px2 > canvasW Then px2 = canvasW
   py2 = canvasH - sy - ny
   If py2 < 1 Then py2 = 1
   If py2 > canvasH Then py2 = canvasH
End Sub

Public Sub CompleteFreeDraw(CulNum As Long)
' NS= 0,1,2 -  1x1 2x2, 4x4 dots
   ' Adjustment to ~ match VB Line()-()
   Select Case BrushType
   Case FreeDraw1: NSX = 0: NSY = 0: ND = 0  ' 1x1
   Case FreeDraw2: NSX = 1: NSY = 0: ND = 1  ' 2x2
   Case FreeDraw3: NSX = 2: NSY = 1: ND = 2  ' 4x4
   End Select
   For NN = 1 To NSTOREXY - 1
      Getpx1py1 STOREX(NN), NSX, STOREY(NN), NSY
      Getpx2py2 STOREX(NN + 1), NSX, STOREY(NN + 1), NSY
      BresLine px1, py1, px2, py2, CulNum, ND
   Next NN
End Sub

Public Sub CompleteRibbon(CulNum As Long)
' Public BrushType, RibIncrX, RibIncrY
' STOREX(),STOREY(),NSTOREXY
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim xd As Single, yd As Single
Dim xstep As Single, ystep As Single
Dim xi As Single, yi As Single
Dim iix As Long, iiy As Long
Dim zlines As Single
Dim zn As Single
' Rest Public
   ' Adjustment to ~ match VB Line()-()
   Select Case BrushType
   Case BRibbon1, FRibbon1: NSX = 0: NSY = 0
   Case BRibbon2, FRibbon2: NSX = 1: NSY = 0
   Case BRibbon3, FRibbon3: NSX = 2: NSY = 1
   End Select
' Rest Public

   '      RY1,RX1
   '          \
   ' BRibbon   .
   '            \
   '          RY2,RX2
   
   RX1 = -RibIncrX
   RY1 = RibIncrY
   RX2 = RibIncrX
   RY2 = -RibIncrY
   
   If BrushType >= 9 Then
   
      '          RY1,RX1
      '            /
      ' FRibbon   .
      '          /
      '      RY2,RX2
      
      RX1 = RibIncrX
      RY1 = RibIncrY
      RX2 = -RibIncrX
      RY2 = -RibIncrY
   End If

   For NN = 1 To NSTOREXY - 1
      
      y1 = STOREY(NN): y2 = STOREY(NN + 1)
      x1 = STOREX(NN): x2 = STOREX(NN + 1)
      
      xd = x2 - x1
      yd = y2 - y1
      
      zlines = Abs(xd)
      If Abs(xd) < Abs(yd) Then zlines = Abs(yd)
      
      If zlines = 0 Then zlines = 1
      xstep = xd / (zlines)
      ystep = yd / (zlines)
      xi = x1
      yi = y1
      For zn = 0 To zlines '+ 1
         iix = CInt(xi)
         iiy = CInt(yi)
         Getpx1py1 iix, -RX1, iiy, -RY1
         Getpx2py2 iix, -RX2, iiy, -RY2
         BresLine px1, py1, px2, py2, CulNum, 0   ' 1 dot
         xi = xi + xstep
         yi = yi + ystep
      Next zn
   Next NN
End Sub

Public Sub CompleteSpray(CulNum As Long)
   For NN = 1 To NSTOREXY
      py1 = canvasH - STOREY(NN)
      px1 = STOREX(NN) + 1
      If px1 > 1 And px1 < canvasW Then
      If py1 > 1 And py1 < canvasH Then
         SetDot px1, py1, CulNum, 0
      End If
      End If
   Next NN
End Sub

Public Sub CompleteSingleLine(CulNum As Long)
   ' Adjustment to ~ match VB Line()-()
   Select Case LineType
   Case SingleLine1: NSX = 0: NSY = 0
   Case SingleLine2: NSX = 1: NSY = 0
   Case SingleLine3: NSX = 2: NSY = 1
   End Select
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   Select Case LineType
   Case SingleLine1: BresLine px1, py1, px2, py2, CulNum, 0
   Case SingleLine2: BresLine px1, py1, px2, py2, CulNum, 1
   Case SingleLine3: BresLine px1, py1, px2, py2, CulNum, 2
   End Select
End Sub

Public Sub CompleteDoubleLine(CulNum As Long)
Dim px(4) As Long, py(4) As Long
Dim nlines As Long
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim zsp As Single
   For NN = 1 To 4
      px(NN) = STOREX(NN) + 1 '- NSX
      If px(NN) < 1 Then px(NN) = 1
      If px(NN) > canvasW Then px(NN) = canvasW
      
      py(NN) = canvasH - STOREY(NN) '- NSY
      If py(NN) < 1 Then py(NN) = 1
      If py(NN) > canvasH Then py(NN) = canvasH
   Next NN
   If LineType = DoubleLine1 Or LineType = DoubleLine2 Or LineType = DoubleLine3 Then
      BresLine px(1), py(1), px(2), py(2), CulNum, 0
      BresLine px(3), py(3), px(4), py(4), CulNum, 0
   ElseIf LineType = DoubleLineEnd1 Or LineType = DoubleLineEnd2 Or LineType = DoubleLineEnd3 Then
      BresLine px(1), py(1), px(2), py(2), CulNum, 0
      BresLine px(3), py(3), px(4), py(4), CulNum, 0
      ' Join ends
      BresLine px(1), py(1), px(3), py(3), CulNum, 0
      BresLine px(2), py(2), px(4), py(4), CulNum, 0
   ElseIf LineType = ShadedLine1 Or LineType = ShadedLine2 Or LineType = ShadedLine3 Then
      Select Case LineType
      Case ShadedLine1: nlines = 5
      Case ShadedLine2: nlines = 9
      Case ShadedLine3: nlines = 17
      End Select
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / nlines
      zCul = SelLeftCulNum
      ixa = px(3)
      iya = py(3)
      ixb = px(4)
      iyb = py(4)
      For NN = 1 To 2 * nlines
         zsp = NN / 2
         BresLine ixa, iya, ixb, iyb, CLng(zCul), 1   ' 2x2
         GetParallelCoords zsp, CSng(px(3)), CSng(py(3)), CSng(px(4)), CSng(py(4)), ixa, iya, ixb, iyb
         zCul = zCul + zdeltacn / 2
         If zCul > 255 Then zCul = 0
         If zCul < 0 Then zCul = 255
      Next NN
   End If
End Sub

Public Sub CompleteDottedLine(CulNum As Long)
Dim zstepx As Single, zstepy As Single
Dim xstepmul As Single, ystepmul As Single
Dim xi As Single, yi As Single
Dim L As Long, i As Long
Dim zalpha As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long

' Rest Public
   ' Adjustment to ~ match VB Line()-()
   Select Case LineType
   Case DottedLine1: NSX = 0: NSY = 0
         xstepmul = 3: ystepmul = 3: ND = 0
   Case DottedLine2: NSX = 1: NSY = 0
         xstepmul = 4: ystepmul = 4: ND = 1
   Case DottedLine3: NSX = 2: NSY = 1
         xstepmul = 8: ystepmul = 8: ND = 2
   End Select
   
   Select Case LineType
   Case DoubleDottedLine1: NSX = 0: NSY = 0
         xstepmul = 3: ystepmul = 3: ND = 0
   Case DoubleDottedLine2: NSX = 1: NSY = 0
         xstepmul = 3: ystepmul = 3: ND = 0
   Case DoubleDottedLine3: NSX = 2: NSY = 1
         xstepmul = 3: ystepmul = 3: ND = 0
   End Select
   
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   zalpha = zATan2((py2 - py1), (px2 - px1))
   zstepy = Sin(zalpha)
   zstepx = Cos(zalpha)
   L = Sqr((py2 - py1) ^ 2 + (px2 - px1) ^ 2)
   If Abs(py2 - py1) > Abs(px2 - px1) Then   ' Steep slope
      If Abs(py2 - py1) <> 0 Then
         yi = py1
         xi = px1
         For i = 1 To L \ ystepmul + 1
            Select Case LineType
            Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
               SetDot CLng(xi), CLng(yi), CulNum, 0 'ND
               GetParallelCoords zspace, xi, yi, CSng(px1), CSng(py1), ixa, iya, ixb, iyb
               SetDot ixa, iya, CulNum, 0 'ND
            Case Else
               SetDot CLng(xi), CLng(yi), CulNum, ND
            End Select
            yi = yi + (ystepmul * zstepy)
            xi = xi + (xstepmul * zstepx)
         Next i
      End If
   Else  ' Abs(py2 - py1) <= Abs(px2 - px1)  ' Shallow slope
      If Abs(px2 - px1) <> 0 Then
         yi = py1
         xi = px1
         For i = 1 To L \ xstepmul + 1
            Select Case LineType
            Case DoubleDottedLine1, DoubleDottedLine2, DoubleDottedLine3
               SetDot CLng(xi), CLng(yi), CulNum, 0 'ND
               GetParallelCoords zspace, xi, yi, CSng(px1), CSng(py1), ixa, iya, ixb, iyb
               SetDot ixa, iya, CulNum, 0 'ND
            Case Else
               SetDot CLng(xi), CLng(yi), CulNum, ND
            End Select
            yi = yi + (ystepmul * zstepy)
            xi = xi + (xstepmul * zstepx)
         Next i
      End If
   End If
End Sub

Public Sub CompletePolySingleLine(CulNum As Long)
If ToolType = CurvyLine And NSTOREXY > 2 Then
   CompleteCurvySingleLine CulNum
   Exit Sub
End If
   
   ' Adjustment to ~ match VB Line()-()
   Select Case PolyLineType
   Case PolySingleLine1: NSX = 0: NSY = 0
   Case PolySingleLine2: NSX = 0: NSY = 0
   Case PolySingleLine3: NSX = 2: NSY = 1
   End Select
   
   For NN = 1 To NSTOREXY - 1
      Getpx1py1 STOREX(NN), NSX, STOREY(NN), NSY
      Getpx2py2 STOREX(NN + 1), NSX, STOREY(NN + 1), NSY
      Select Case PolyLineType
      Case PolySingleLine1: BresLine px1, py1, px2, py2, CulNum, 0
      Case PolySingleLine2: BresLine px1, py1, px2, py2, CulNum, 1
      Case PolySingleLine3: BresLine px1, py1, px2, py2, CulNum, 2
      End Select
   Next NN
End Sub

Public Sub CompletePolyDoubleLine(CulNum As Long)
Dim px3 As Long, py3 As Long
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim x3 As Single, y3 As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim ixc As Long, iyc As Long
Dim ixd As Long, iyd As Long
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ix3 As Long, iy3 As Long
Dim svix As Long, sviy As Long
Dim svixa As Long, sviya As Long
Dim svLineType As Long
   
'Private tx() As Single, ty() As Single
'Private ux() As Single, uy() As Single
ReDim TX(1), TY(1), ux(1), uy(1)
   
   If NSTOREXY = 2 Then
      ' Get 2 parallel points, Set LineType
      ' Call CompleteDoubleLine CulNum
      ' Restore LineType
      ' Exit sub
         x1 = STOREX(1): y1 = STOREY(1)
         x2 = STOREX(2): y2 = STOREY(2)
         GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
         NSTOREXY = NSTOREXY + 2
         ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
         STOREX(3) = ixa: STOREY(3) = iya
         STOREX(4) = ixb: STOREY(4) = iyb
         svLineType = LineType
         Select Case PolyLineType
         Case PolyDoubleLine1: LineType = DoubleLine1
         Case PolyDoubleLine2: LineType = DoubleLine2
         Case PolyDoubleLine3: LineType = DoubleLine3
         Case PolyDoubleLineEnd1: LineType = DoubleLineEnd1
         Case PolyDoubleLineEnd2: LineType = DoubleLineEnd2
         Case PolyDoubleLineEnd3: LineType = DoubleLineEnd3
         Case PolyShadedLine1: LineType = ShadedLine1
         Case PolyShadedLine2: LineType = ShadedLine2
         Case PolyShadedLine3: LineType = ShadedLine3
         End Select
         CompleteDoubleLine CulNum
         LineType = svLineType
         NSTOREXY = 2
         ReDim Preserve STOREX(NSTOREXY), STOREY(NSTOREXY)
         Exit Sub
   End If
   
If ToolType = CurvyLine Then  '  NSTOREXY > 2
   CompleteCurvyDoubleLine CulNum
   Exit Sub
End If
   
   ' Adjustment to ~ match VB Line()-()
   Select Case PolyLineType
   Case PolyDoubleLine1, PolyDoubleLineEnd1, PolyShadedLine1: NSX = 0: NSY = 0
   Case PolyDoubleLine2, PolyDoubleLineEnd2, PolyShadedLine2: NSX = 0: NSY = 0
   Case PolyDoubleLine3, PolyDoubleLineEnd3, PolyShadedLine3: NSX = 2: NSY = 1
   End Select
   
   ' Draw input lines and
   ' find & draw parallel lines.
   ' Backwards to get parallel line on correct side
   ' of input lines
   Nutt = 0
   Nutu = 0
   
   For NN = NSTOREXY To 2 Step -1 ' Backwards to get parallel lines on same side
                                  ' as PIC draw
      Getpx1py1 STOREX(NN), NSX, STOREY(NN), NSY
      Getpx2py2 STOREX(NN - 1), NSX, STOREY(NN - 1), NSY
      BresLine px1, py1, px2, py2, CulNum, 0
      Nutu = Nutu + 1
      ReDim Preserve ux(Nutu), uy(Nutu)
      ux(Nutu) = px1: uy(Nutu) = py1
      If NN = 2 Then
         Nutu = Nutu + 1
         ReDim Preserve ux(Nutu), uy(Nutu)
         ux(Nutu) = px2: uy(Nutu) = py2
      End If
      
      If NN > 2 Then ' 3 points give intersections
         px3 = STOREX(NN - 2) + 1 - NSX
         If px3 < 1 Then px3 = 1
         If px3 > canvasW Then px3 = canvasW
         py3 = canvasH - STOREY(NN - 2) - NSY
         If py3 < 1 Then py3 = 1
         If py3 > canvasH Then py3 = canvasH
         
         x1 = px1: y1 = py1
         x2 = px2: y2 = py2
         x3 = px3: y3 = py3
         ' In: zspace, x1, y1, x2, y2, x3, y3
         GetIntersection zspace, x1, y1, x2, y2, x3, y3, _
                ixa, iya, ixb, iyb, ixc, iyc, ixd, iyd, _
                ix1, iy1, ix2, iy2, ix3, iy3
         ' Out: ixa, iya, ixb, iyb, ixc, iyc, ixd, iyd parallel points
         '      ix1, iy1, ix2, iy2  intersection ponts, same unless out of bounds
         '      ix3, iy3  possible out of bounds intersection point
         If NN = NSTOREXY Then
            Nutt = 3
            ReDim TX(3), TY(3)
            TX(Nutt - 2) = ixa: TY(Nutt - 2) = iya
            TX(Nutt - 1) = ix3: TY(Nutt - 1) = iy3
            TX(Nutt) = ixd: TY(Nutt) = iyd
         Else
            Nutt = Nutt + 1
            ReDim Preserve TX(Nutt), TY(Nutt)
            TX(Nutt - 1) = ix3: TY(Nutt - 1) = iy3
            TX(Nutt) = ixd: TY(Nutt) = iyd
         End If
         
         If NSTOREXY = 3 Then
            BresLine ixa, iya, ix3, iy3, CulNum, 0
            BresLine ix3, iy3, ixd, iyd, CulNum, 0
      
            Select Case PolyLineType
            Case PolyDoubleLineEnd1, PolyDoubleLineEnd2, PolyDoubleLineEnd3
            Getpx1py1 STOREX(NSTOREXY), NSX, STOREY(NSTOREXY), NSY
            BresLine px1, py1, ixa, iya, CulNum, 0
            Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
            BresLine px1, py1, ixd, iyd, CulNum, 0
            End Select
         
         Else  ' NSTOREXY > 3
         
            If NN = NSTOREXY Then
               BresLine ixa, iya, ix3, iy3, CulNum, 0
               svixa = ixa: sviya = iya ' Save start point
               svix = ix3: sviy = iy3
            ElseIf NN < NSTOREXY And NN > 3 Then
                  BresLine svix, sviy, ix3, iy3, CulNum, 0
                  svix = ix3: sviy = iy3
            ElseIf NN = 3 Then
               BresLine svix, sviy, ix3, iy3, CulNum, 0
               BresLine ix3, iy3, ixd, iyd, CulNum, 0
               
               Select Case PolyLineType
               Case PolyDoubleLineEnd1, PolyDoubleLineEnd2, PolyDoubleLineEnd3
                  Getpx1py1 STOREX(NSTOREXY), NSX, STOREY(NSTOREXY), NSY
                  BresLine px1, py1, svixa, sviya, CulNum, 0
                  Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
                  BresLine px1, py1, ixd, iyd, CulNum, 0
               End Select
            End If   ' If NN = NSTOREXY Then, NN < NSTOREXY And NN > 3 Then, ' If NN = 3 Then
         
         End If   ' If NSTOREXY = 3 Then
      
      End If   ' If NN > 2 Then
   
   Next NN
   
   Select Case PolyLineType
   Case PolyShadedLine1, PolyShadedLine2, PolyShadedLine3
      Dim zA1 As Single, zrad1 As Single
      Dim zA2 As Single, zrad2 As Single
      Dim xd1 As Single, yd1 As Single
      Dim xd2 As Single, yd2 As Single
      Dim xx1 As Single, yy1 As Single
      Dim xx2 As Single, yy2 As Single
      
      For NN = 1 To Nutt - 1
         
         xd1 = TX(NN) - ux(NN)
         yd1 = TY(NN) - uy(NN)
         zrad1 = Sqr(xd1 * xd1 + yd1 * yd1)
         If zrad1 = 0 Then zrad1 = 0.01
         If zrad1 > 1000 Then zrad1 = 1000
         zA1 = zATan2(-yd1, xd1) - pi# / 2   ' Take clockwise angle
         xd1 = Sin(zA1): yd1 = Cos(zA1)      ' x & y incr
         
         xd2 = TX(NN + 1) - ux(NN + 1)
         yd2 = TY(NN + 1) - uy(NN + 1)
         zrad2 = Sqr(xd2 * xd2 + yd2 * yd2)
         If zrad2 = 0 Then zrad2 = 0.01
         If zrad2 > 1000 Then zrad2 = 1000
         zA2 = zATan2(-yd2, xd2) - pi# / 2   ' Take clockwise angle
         xd2 = Sin(zA2) * zrad2 / zrad1: yd2 = Cos(zA2) * zrad2 / zrad1
         
         xx1 = TX(NN)
         yy1 = TY(NN)
         xx2 = TX(NN + 1)
         yy2 = TY(NN + 1)
         
         zdeltacn = (SelRightCulNum - SelLeftCulNum) / zrad1
         zCul = SelLeftCulNum
         
         For k = 1 To CInt(zrad1 + 0.5)
            BresLine xx1, yy1, xx2, yy2, CLng(zCul), 1   ' 2x2 dots on line
            zCul = zCul + zdeltacn
            If zCul < 0 Then zCul = 255
            If zCul > 255 Then zCul = 0
            xx1 = xx1 + xd1
            yy1 = yy1 + yd1
            xx2 = xx2 + xd2
            yy2 = yy2 + yd2
         Next k
      Next NN
   End Select
Erase TX(), TY(), ux(), uy()
End Sub

Public Sub CompleteCurvySingleLine(CulNum)
'Coming in from PolyLineType
Dim cx() As Long, cy() As Long
' Rest Public
   
   ' Adjustment to ~ match VB Line()-()
   Select Case PolyLineType
   Case PolySingleLine1: NSX = 0: NSY = 0
   Case PolySingleLine2: NSX = 1: NSY = 0
   Case PolySingleLine3: NSX = 2: NSY = 1
   End Select
' Points in STOREX(NSTOREXY),STOREY(NSTOREXY)
' Convert to curvy points
' Number of curvy points = 4*(NSTOREXY-2) + 2
   NN = 4 * (NSTOREXY - 2) + 2
   'NN = 8 * (NSTOREXY - 2) + 2
   
   ReDim cx(NN), cy(NN)
   GetCurvyPoints STOREX(), STOREY(), cx(), cy()
   
   For NN = 1 To NN - 1
      Getpx1py1 CLng(cx(NN)), NSX, CLng(cy(NN)), NSY
      Getpx2py2 CLng(cx(NN + 1)), NSX, CLng(cy(NN + 1)), NSY
      Select Case PolyLineType
      Case PolySingleLine1: BresLine px1, py1, px2, py2, CulNum, 0
      Case PolySingleLine2: BresLine px1, py1, px2, py2, CulNum, 1
      Case PolySingleLine3: BresLine px1, py1, px2, py2, CulNum, 2
      End Select
   Next NN
Erase cx(), cy()
End Sub

Public Sub GetCurvyPoints(inx() As Long, iny() As Long, outx() As Long, outy() As Long)
Dim NIN As Long
Dim NOUT As Long
Dim ki As Long
Dim ko As Long
Dim xp1 As Single, yp1 As Single
Dim xp2 As Single, yp2 As Single
Dim Sect As Long
   Sect = 5
   NIN = UBound(inx(), 1)
   NOUT = UBound(outx(), 1)
   outx(1) = inx(1): outy(1) = iny(1)
   ko = 2
   For ki = 2 To NIN - 1
      xp1 = inx(ki) - (inx(ki) - inx(ki - 1)) / Sect
      yp1 = iny(ki) - (iny(ki) - iny(ki - 1)) / Sect
      xp2 = inx(ki) + (inx(ki + 1) - inx(ki)) / Sect
      yp2 = iny(ki) + (iny(ki + 1) - iny(ki)) / Sect
      
      outx(ko) = xp1 - (xp1 - inx(ki - 1)) / Sect
      outy(ko) = yp1 - (yp1 - iny(ki - 1)) / Sect
      ko = ko + 1
      outx(ko) = xp1 + (xp2 - xp1) / Sect
      outy(ko) = yp1 + (yp2 - yp1) / Sect
      ko = ko + 1
      outx(ko) = xp2 - (xp2 - xp1) / Sect
      outy(ko) = yp2 - (yp2 - yp1) / Sect
      ko = ko + 1
      outx(ko) = xp2 + (inx(ki + 1) - xp2) / Sect
      outy(ko) = yp2 + (iny(ki + 1) - yp2) / Sect
      ko = ko + 1
   Next ki
   outx(NOUT) = inx(NIN): outy(NOUT) = iny(NIN)
End Sub

Public Sub CompleteCurvyDoubleLine(CulNum As Long)
'Coming in with PolyLineTpes
'Dim px1 As Long, py1 As Long
'Dim px2 As Long, py2 As Long
Dim px3 As Long, py3 As Long
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim x3 As Single, y3 As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim ixc As Long, iyc As Long
Dim ixd As Long, iyd As Long
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ix3 As Long, iy3 As Long
   
   ' Adjustment to ~ match VB Line()-()
   Select Case PolyLineType
   Case PolyDoubleLine1, PolyDoubleLineEnd1, PolyShadedLine1: NSX = 0: NSY = 0
   Case PolyDoubleLine2, PolyDoubleLineEnd2, PolyShadedLine2: NSX = 1: NSY = 0
   Case PolyDoubleLine3, PolyDoubleLineEnd3, PolyShadedLine3: NSX = 2: NSY = 1
   End Select
   
   ' Store input points and
   ' find parallel line points
   ' Backwards to get parallel line on correct side
   ' of input lines
   Nutt = 0
   Nutu = 0
   
   For NN = NSTOREXY To 2 Step -1 ' Backwards to get parallel lines on same side
                                  ' as PIC draw
      Getpx1py1 STOREX(NN), NSX, STOREY(NN), NSY
      Getpx2py2 STOREX(NN - 1), NSX, STOREY(NN - 1), NSY
      Nutu = Nutu + 1
      ReDim Preserve ux(Nutu), uy(Nutu)
      ux(Nutu) = px1: uy(Nutu) = py1
      
      If NN = 2 Then
      Nutu = Nutu + 1
      ReDim Preserve ux(Nutu), uy(Nutu)
      ux(Nutu) = px2: uy(Nutu) = py2
      End If
      
      If NN > 2 Then ' 3 points give intersections
      
         px3 = STOREX(NN - 2) + 1 - NSX
         If px3 < 1 Then px3 = 1
         If px3 > canvasW Then px3 = canvasW
         py3 = canvasH - STOREY(NN - 2) - NSY
         If py3 < 1 Then py3 = 1
         If py3 > canvasH Then py3 = canvasH
         
         x1 = px1: y1 = py1
         x2 = px2: y2 = py2
         x3 = px3: y3 = py3
         ' In: zspace, x1, y1, x2, y2, x3, y3
         GetIntersection zspace, x1, y1, x2, y2, x3, y3, _
                ixa, iya, ixb, iyb, ixc, iyc, ixd, iyd, _
                ix1, iy1, ix2, iy2, ix3, iy3
         ' Out: ixa, iya, ixb, iyb, ixc, iyc, ixd, iyd parallel points
         '      ix1, iy1, ix2, iy2  intersection ponts, same unless out of bounds
         '      ix3, iy3  possible out of bounds intersection point
         If NN = NSTOREXY Then
            Nutt = 3
            ReDim TX(3), TY(3)
            TX(Nutt - 2) = ixa: TY(Nutt - 2) = iya
            TX(Nutt - 1) = ix3: TY(Nutt - 1) = iy3
            TX(Nutt) = ixd: TY(Nutt) = iyd
         Else
            Nutt = Nutt + 1
            ReDim Preserve TX(Nutt), TY(Nutt)
            TX(Nutt - 1) = ix3: TY(Nutt - 1) = iy3
            TX(Nutt) = ixd: TY(Nutt) = iyd
         End If
      End If   ' If NN > 2 Then
   Next NN

   ' Points on DoublePolyLines
   ' tx(Nutt),ty(Nutt)
   ' ux(Nutu),uy(Nutu)
 
   ' Convert to curvy points
   ' Number of curvy points = 4*(NN-2) + 2
   NNN = 4 * (Nutt - 2) + 2
   ReDim cx(NNN), cy(NNN)
   GetCurvyPoints TX(), TY(), cx(), cy()

   NNN = 4 * (Nutu - 2) + 2
   ReDim dx(NNN), dy(NNN)
   GetCurvyPoints ux(), uy(), dx(), dy()

   For NN = 1 To NNN - 1
      px1 = CLng(cx(NN)) + 1 - NSX
      If px1 < 1 Then px1 = 1
      If px1 > canvasW Then px1 = canvasW
      py1 = CLng(cy(NN)) - NSY
      If py1 < 1 Then py1 = 1
      If py1 > canvasH Then py1 = canvasH
      px2 = CLng(cx(NN + 1)) + 1 - NSX
      If px2 < 1 Then px2 = 1
      If px2 > canvasW Then px2 = canvasW
      py2 = CLng(cy(NN + 1)) - NSY
      If py2 < 1 Then py2 = 1
      If py2 > canvasH Then py2 = canvasH
      BresLine px1, py1, px2, py2, CulNum, 0
      
      px1 = CLng(dx(NN)) + 1 - NSX
      If px1 < 1 Then px1 = 1
      If px1 > canvasW Then px1 = canvasW
      py1 = CLng(dy(NN)) - NSY
      If py1 < 1 Then py1 = 1
      If py1 > canvasH Then py1 = canvasH
      px2 = CLng(dx(NN + 1)) + 1 - NSX
      If px2 < 1 Then px2 = 1
      If px2 > canvasW Then px2 = canvasW
      py2 = CLng(dy(NN + 1)) - NSY
      If py2 < 1 Then py2 = 1
      If py2 > canvasH Then py2 = canvasH
      BresLine px1, py1, px2, py2, CulNum, 0
   Next NN
   
   ' Test DoubleEnd or Shaded
   
   Select Case PolyLineType
   Case PolyDoubleLineEnd1, PolyDoubleLineEnd2, PolyDoubleLineEnd3
      px1 = CLng(cx(1)) + 1 - NSX
      If px1 < 1 Then px1 = 1
      If px1 > canvasW Then px1 = canvasW
      py1 = CLng(cy(1)) - NSY
      If py1 < 1 Then py1 = 1
      If py1 > canvasH Then py1 = canvasH
      px2 = CLng(dx(1)) + 1 - NSX
      If px2 < 1 Then px2 = 1
      If px2 > canvasW Then px2 = canvasW
      py2 = CLng(dy(1)) - NSY
      If py2 < 1 Then py2 = 1
      If py2 > canvasH Then py2 = canvasH
      BresLine px1, py1, px2, py2, CulNum, 0
      px1 = CLng(cx(NNN)) + 1 - NSX
      If px1 < 1 Then px1 = 1
      If px1 > canvasW Then px1 = canvasW
      py1 = CLng(cy(NNN)) - NSY
      If py1 < 1 Then py1 = 1
      If py1 > canvasH Then py1 = canvasH
      px2 = CLng(dx(NNN)) + 1 - NSX
      If px2 < 1 Then px2 = 1
      If px2 > canvasW Then px2 = canvasW
      py2 = CLng(dy(NNN)) - NSY
      If py2 < 1 Then py2 = 1
      If py2 > canvasH Then py2 = canvasH
      BresLine px1, py1, px2, py2, CulNum, 0
   End Select

   Select Case PolyLineType
   Case PolyShadedLine1, PolyShadedLine2, PolyShadedLine3
      Dim zA1 As Single, zrad1 As Single
      Dim zA2 As Single, zrad2 As Single
      Dim xd1 As Single, yd1 As Single
      Dim xd2 As Single, yd2 As Single
      Dim xx1 As Single, yy1 As Single
      Dim xx2 As Single, yy2 As Single
      
      For NN = 1 To NNN - 1
         xd1 = cx(NN) - dx(NN)
         yd1 = cy(NN) - dy(NN)
         zrad1 = Sqr(xd1 * xd1 + yd1 * yd1)
         If zrad1 = 0 Then zrad1 = 0.01
         If zrad1 > 1000 Then zrad1 = 1000
         zA1 = zATan2(-yd1, xd1) - pi# / 2   ' Take clockwise angle
         xd1 = Sin(zA1)
         yd1 = Cos(zA1)      ' x & y incr
         
         xd2 = cx(NN + 1) - dx(NN + 1)
         yd2 = cy(NN + 1) - dy(NN + 1)
         zrad2 = Sqr(xd2 * xd2 + yd2 * yd2)
         If zrad2 = 0 Then zrad2 = 0.01
         If zrad2 > 1000 Then zrad2 = 1000
         zA2 = zATan2(-yd2, xd2) - pi# / 2   ' Take clockwise angle
         xd2 = Sin(zA2) * zrad2 / zrad1
         yd2 = Cos(zA2) * zrad2 / zrad1
         
         xx1 = cx(NN)
         yy1 = cy(NN)
         xx2 = cx(NN + 1)
         yy2 = cy(NN + 1)
         
         zdeltacn = (SelRightCulNum - SelLeftCulNum) / zrad1
         zCul = SelLeftCulNum

         For k = 1 To CInt(zrad1 + 0.5)
            BresLine xx1, yy1, xx2, yy2, CLng(zCul), 1   ' 2x2 dots on line
            zCul = zCul + zdeltacn
            If zCul < 0 Then zCul = 255
            If zCul > 255 Then zCul = 0
            xx1 = xx1 + xd1
            yy1 = yy1 + yd1
            xx2 = xx2 + xd2
            yy2 = yy2 + yd2
         Next k
      Next NN
   End Select
Erase cx(), cy()
Erase dx(), dy()
Erase ux(), uy()
End Sub

Public Sub CompleteRectangleSingle(CulNum As Long)
   ' Adjustment to ~ match VB Line()-()
   Select Case RectangleType
   Case RectangleSingle1: NSX = 0: NSY = 0
   Case RectangleSingle2: NSX = 1: NSY = 0
   Case RectangleSingle3: NSX = 2: NSY = 1
   End Select
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   Select Case RectangleType
   Case RectangleSingle1: BresBox px1, py1, px2, py2, CulNum, 0
   Case RectangleSingle2: BresBox px1, py1, px2, py2, CulNum, 1
   Case RectangleSingle3: BresBox px1, py1, px2, py2, CulNum, 2
   End Select
End Sub

Public Sub CompleteRectangleDotted(CulNum As Long)
Dim stepx As Long, stepy As Long
' Rest Public
   ' Adjustment to ~ match VB Line()-()
   Select Case RectangleType
   Case RectangleDotted1: NSX = 0: NSY = 0
   Case RectangleDotted2: NSX = 1: NSY = 0
   Case RectangleDotted3: NSX = 2: NSY = 1
   End Select
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   Select Case RectangleType
   Case RectangleDotted1
      stepx = 2 * Sgn(px2 - px1)
      stepy = 2 * Sgn(py2 - py1)
      ND = 0
   Case RectangleDotted2: NSX = 1: NSY = 0
      stepx = 4 * Sgn(px2 - px1)
      stepy = 4 * Sgn(py2 - py1)
      ND = 1
   Case RectangleDotted3: NSX = 2: NSY = 1
      stepx = 8 * Sgn(px2 - px1)
      stepy = 8 * Sgn(py2 - py1)
      ND = 2
   End Select
   
   ' Top dotted line
   If stepx <> 0 Then
      For k = px1 To px2 Step stepx
         SetDot k, py1, CulNum, ND
      Next k
   End If
   ' Left dotted line
   If stepy <> 0 Then
      For k = py1 To py2 Step stepy
         SetDot px1, k, CulNum, ND
      Next k
   End If
   ' Right dotted line
   If stepy <> 0 Then
      For k = py1 To py2 Step stepy
         SetDot px2, k, CulNum, ND
      Next k
   ' Bottom dotted line
   End If
   If stepx <> 0 Then
      For k = px1 To px2 Step stepx
         SetDot k, py2, CulNum, ND
      Next k
   End If
End Sub

Public Sub CompleteRectangleDouble(CulNum As Long)
   ' Adjustment to ~ match VB Line()-()
   Select Case RectangleType
   Case RectangleDouble1: NSX = 0: NSY = 0: ND = 2
   Case RectangleDouble2: NSX = 1: NSY = 0: ND = 4
   Case RectangleDouble3: NSX = 2: NSY = 1: ND = 8
   End Select
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   BresBox px1, py1, px2, py2, CulNum, 0
   BresBox px1 + ND, py1 - ND, px2 - ND, py2 + ND, CulNum, 0
End Sub

Public Sub CompleteRectangleShaded(CulNum As Long)
Dim xd As Single
Dim yd As Single

Dim nsq As Long
Dim xdx As Single, ydy As Single
Dim N As Long
Dim x1 As Single, x2 As Single
Dim y1 As Single, y2 As Single
   
   ' Adjustment to ~ match VB Line()-()
   Select Case svRectangleType
   Case RectangleShaded1: NSX = 0: NSY = 0
   Case RectangleShaded2: NSX = 1: NSY = 0
   Case RectangleShaded3: NSX = 2: NSY = 1
   
   Case RectangleFShade: NSX = 0: NSY = 0
   Case RectangleBShade: NSX = 0: NSY = 0
   Case RectangleFilled: NSX = 0: NSY = 0
   End Select
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   xd = px2 - px1
   yd = py2 - py1
   zCul = SelLeftCulNum
   
   Select Case svRectangleType
   Case RectangleShaded1   ' Shade ||
      If xd = 0 Then xd = 1
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(xd)
      For k = px1 To px2 Step Sgn(xd)
         BresLine k, py1, k, py2, CLng(zCul), 0
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next k
   
   Case RectangleShaded2   ' Shade =
      If yd = 0 Then yd = 1
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(yd)
      For k = py1 To py2 Step Sgn(yd)
         BresLine px1, k, px2, k, CLng(zCul), 0
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next k
   Case RectangleShaded3   ' Shade O
      nsq = Abs(px2 - px1) / 2
      If nsq > Abs(py2 - py1) / 2 Then nsq = Abs(py2 - py1) / 2
      If nsq = 0 Then nsq = 1
      zdeltacn = ((SelRightCulNum - SelLeftCulNum) + Sgn(SelRightCulNum - SelLeftCulNum)) / nsq
      For k = 1 To nsq
         BresBox px1, py1, px2, py2, CLng(zCul), 0
         px1 = px1 + Sgn(px2 - px1)
         px2 = px2 - Sgn(px2 - px1)
         py1 = py1 + Sgn(py2 - py1)
         py2 = py2 - Sgn(py2 - py1)
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next k
   Case RectangleFShade, RectangleBShade, RectangleFilled
      Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
      Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
      If px2 < px1 And py2 < py1 Then
         k = px1
         px1 = px2
         px2 = k
         k = py1
         py1 = py2
         py2 = k
      ElseIf px1 > px2 Then
         k = px1
         px1 = px2
         px2 = k
      ElseIf py1 > py2 Then
         k = py1
         py1 = py2
         py2 = k
      End If
      xd = px2 - px1
      If xd = 0 Then xd = 100
      yd = py2 - py1
      If yd = 0 Then yd = 100
      zCul = SelLeftCulNum
      Select Case svRectangleType
      Case RectangleFShade    'diagonal shaded  /
         If xd > yd Then
            xdx = 1
            ydy = yd / xd
            zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(xd)
            N = 0
            For k = px1 To px2
               y2 = py2 - N * ydy
               BresLine k, py1, px2, y2, CLng(zCul), 1
               y1 = py1 + N * ydy
               BresLine px1, y1, px2 - N, py2, CLng(zCul), 1
               N = N + 1
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next k
         Else
            ydy = 1
            xdx = xd / yd
            zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(yd)
            N = 0
            For k = py1 To py2
               x2 = px2 - N * xdx
               BresLine px1, k, x2, py2, CLng(zCul), 1
               x1 = px1 + N * xdx
               BresLine x1, py1, px2, py2 - N, CLng(zCul), 0
               N = N + 1
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next k
         End If
      Case RectangleBShade    'diagonal shaded  \
         If xd > yd Then
            xdx = 1
            ydy = yd / xd
            zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(xd)
            N = 0
            For k = px1 To px2
               y2 = py1 + N * ydy
               BresLine k, py2, px2, y2, CLng(zCul), 1
               y1 = py2 - N * ydy
               BresLine px1, y1, px2 - N, py1, CLng(zCul), 1
               N = N + 1
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next k
         Else
            ydy = 1
            xdx = xd / yd
            zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(yd)
            N = 0
            For k = py1 To py2
               x1 = px1 + N * xdx
               BresLine x1, py2, px2, k, CLng(zCul), 1
               x2 = px2 - N * xdx
               BresLine px1, py2 - N, x2, py1, CLng(zCul), 0
               N = N + 1
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next k
         End If
      Case RectangleFilled
         For k = px1 To px2
            BresLine k, py1, k, py2, CulNum, 0
         Next k
      End Select
   End Select
End Sub

Public Sub CompleteCirllipseSDD(CulNum As Long)
Dim zspace As Single, NDSP As Long
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long

   ' Adjustment to ~ match VB Line()-()
   ND = 0
   Select Case CirllipseType
   Case CirllipseSingle1, CirllipseDotted1
      NSX = 0: NSY = 0: ND = 0
   Case CirllipseSingle2, CirllipseDotted2
      NSX = 1: NSY = 0: ND = 1
   Case CirllipseSingle3, CirllipseDotted3
      NSX = 2: NSY = 1: ND = 2
   End Select
   zspace = 0
   Select Case CirllipseType
   Case CirllipseDouble1: zspace = 2: NSX = 0: NSY = 0
   Case CirllipseDouble2: zspace = 4: NSX = 0: NSY = 0
   Case CirllipseDouble3: zspace = 7: NSX = 0: NSY = 0
   End Select
   NDSP = 0
   Select Case CirllipseType
   Case CirllipseDotted1: NDSP = 2
   Case CirllipseDotted2: NDSP = 4
   Case CirllipseDotted3: NDSP = 8
   End Select
   
   ' STOREXY(1) ixc,iyc
   ' STOREXY(2) X,Y
   ' zrad, zratio
'   ixc = STOREX(1)
'   iyc = STOREY(1)
   
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   If Abs(px2 - px1) > Abs(py2 - py1) Then
      zradx = zrad
      zrady = zradx * zratio
   Else
      zrady = zrad
      zradx = zrady / zratio
   End If
   
   If zrad > 0 Then
      ' px1,py1 are ixc,iyc center coords
      ix1 = px1 - zradx
      iy1 = py1 + zrady
      ix2 = ix1 + 2 * zradx
      iy2 = iy1 - 2 * zrady
      Select Case CirllipseType
      Case CirllipseSingle1, CirllipseSingle2, CirllipseSingle3, _
           CirllipseDouble1, CirllipseDouble2, CirllipseDouble3
         BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
         Select Case CirllipseType
         Case CirllipseDouble1, CirllipseDouble2, CirllipseDouble3
            EvalZradZratio CSng(STOREX(2)), CSng(STOREY(2))
            zradx = zradx - zspace
            zrady = zrady - zspace
            If zradx > 0 And zrady > 0 Then
               ' px1,py1 are ixc,iyc center coords
               ix1 = px1 - zradx
               iy1 = py1 + zrady
               ix2 = ix1 + 2 * zradx
               iy2 = iy1 - 2 * zrady
               BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
            End If
         End Select
      Case CirllipseDotted1, CirllipseDotted2, CirllipseDotted3
         BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, NDSP, 0
      End Select
   End If
End Sub

Public Sub CompleteCirllipseShaded(CulNum As Long)
Dim zspace As Single
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim zalpha As Single
Dim xd As Single
Dim yd As Single
Dim nsq As Long
Dim svSTOREX As Long
Dim svSTOREY As Long
   
   'ixc = STOREX(1)
   'iyc = STOREY(1)
   ND = 1
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   If Abs(px2 - px1) > Abs(py2 - py1) Then
      zradx = zrad
      zrady = zradx * zratio
   Else
      zrady = zrad
      zradx = zrady / zratio
   End If
   ' px1,py1 are ixc,iyc center coords
   ix1 = px1 - zradx
   iy1 = py1 + zrady
   ix2 = ix1 + 2 * zradx
   iy2 = iy1 - 2 * zrady
   Select Case CirllipseType
   Case CirllipseShaded1   '||
      EvalZradZratio CSng(STOREX(2)), CSng(STOREY(2))
      ' OUT: zrad, zratio, zradx, zrady
      Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
      Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
      ix1 = px1 - zradx
      ix2 = ix1 + 2 * zradx
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(2 * zradx)
      zCul = SelLeftCulNum
      For k = ix1 To ix2 Step Sgn(ix2 - ix1)
         yd = (zrady / zradx) * Sqr(Abs(zradx ^ 2 - (k - ixc) ^ 2))
         BresLine k, py1 - yd, k, py1 + yd, CLng(zCul), 0   ' 2x2 dots on line
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next k
   Case CirllipseShaded2   ' =
      EvalZradZratio CSng(STOREX(2)), CSng(STOREY(2))
      ' OUT: zrad, zratio, zradx, zrady
      Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
      Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
      iy1 = py1 - zrady
      iy2 = iy1 + 2 * zrady
         zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(2 * zrady)
         zCul = SelLeftCulNum
         For k = iy1 To iy2 Step Sgn(iy2 - iy1)
            xd = (zradx / zrady) * Sqr(Abs(zrady ^ 2 - (k - py1) ^ 2))
   
            ' Segment shading  (())
            Dim iix As Long
            If xd = 0 Then xd = 0.5
            zdeltacn = (SelRightCulNum - SelLeftCulNum) / Abs(2 * xd)
            zCul = SelLeftCulNum
            For iix = px1 - xd To px1 + xd
               SetDot iix, k, CLng(zCul), 0
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next iix
            ' Horz band shading   ' =
            'BresLine px1 - xd, k, px1 + xd, k, CLng(zCul), 0 ' 1x1 dots on line
            'zCul = zCul + zdeltacn
            'If zCul < 0 Then zCul = 255
            'If zCul > 255 Then zCul = 0
         Next k
   
   Case CirllipseShaded3   ' O
      nsq = Sqr((STOREX(2) - STOREX(1)) ^ 2 + (STOREY(2) - STOREY(1)) ^ 2)
      If nsq = 0 Then nsq = 1
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / nsq
      zCul = SelLeftCulNum
      zspace = 1
      svSTOREX = STOREX(2) ' Save for copying
      svSTOREY = STOREY(2)
      For k = 1 To nsq
         xd = (STOREX(2) - STOREX(1))
         yd = (STOREY(2) - STOREY(1))
         zalpha = zATan2(yd, xd)
         STOREX(2) = Int(STOREX(2) - zspace * Cos(zalpha) + 0.5)
         STOREY(2) = Int(STOREY(2) - zspace * Sin(zalpha) + 0.5)
         EvalZradZratio CSng(STOREX(2)), CSng(STOREY(2))
         If zrad < 1 Then Exit For
         Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
         If Abs(px2 - px1) > Abs(py2 - py1) Then
            zradx = zrad
            zrady = zradx * zratio
         Else
            zrady = zrad
            zradx = zrady / zratio
         End If
         ' px1,py1 are ixc,iyc center coords
         ix1 = px1 - zradx
         iy1 = py1 + zrady
         ix2 = ix1 + 2 * zradx
         iy2 = iy1 - 2 * zrady
         BresEllipse ix1, iy1, ix2, iy2, CLng(zCul), ND, 0, 0
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next k
      STOREX(2) = svSTOREX ' Reset for copying
      STOREY(2) = svSTOREY
   End Select
End Sub

Public Sub CompleteCone(CulNum As Long)
Dim zspace As Single
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ixr As Long, iyr As Long
Dim ixp As Long, iyp As Long
Dim zalpha As Single
Dim zalpha1 As Single
Dim zalpha2 As Single
Dim zstep As Single

   If zrad = 0 Then Exit Sub
   
   ' Adjustment to ~ match VB Line()-()
   NSX = 0: NSY = 0: ND = 0
   ' STOREXY(1) ixc,iyc
   ' STOREXY(2) X,Y
   ' STOREXY(3) XDiag,YDiag
   ' zrad
   
   If zrad = 0 Then Exit Sub
   
   If ConeType = ConeCross Then
      ' Adjustment to ~ match VB Line()-()
      NSX = 0: NSY = 0: ND = 0
      ' STOREXY(1) ixc,iyc
      ' STOREXY(2) X,Y
      ' STOREXY(3) XDiag,YDiag
      ' zrad
      ixr = STOREX(3)
      iyr = STOREY(3)
      Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
      ixc = px1
      iyc = py1
      ' ixc,iyc are center coords 1st circle
      ' px1,py1 are ixc,iyc center coords
      ix1 = px1 - zrad
      iy1 = py1 + zrad
      ix2 = ix1 + 2 * zrad
      iy2 = iy1 - 2 * zrad
      BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
      Getpx2py2 STOREX(3), NSX, STOREY(3), NSY
      ixr = px2
      iyr = py2
      ' ixr,iyr are center coords 2nd circle
      ix1 = px2 - zrad
      iy1 = py2 + zrad
      ix2 = ix1 + 2 * zrad
      iy2 = iy1 - 2 * zrad
      BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
      EvalTangents ixc, iyc, zrad, (ixc + ixr) \ 2, (iyc + iyr) \ 2, ix1, iy1, ix2, iy2
      BresLine ix1, iy1, (ixc + ixr) \ 2, (iyc + iyr) \ 2, CulNum, 0   ' 1x1 dots on line
      BresLine ix2, iy2, (ixc + ixr) \ 2, (iyc + iyr) \ 2, CulNum, 0   ' 1x1 dots on line
      
      EvalTangents ixr, iyr, zrad, (ixc + ixr) \ 2, (iyc + iyr) \ 2, ix1, iy1, ix2, iy2
      BresLine ix1, iy1, (ixc + ixr) \ 2, (iyc + iyr) \ 2, CulNum, 0   ' 1x1 dots on line
      BresLine ix2, iy2, (ixc + ixr) \ 2, (iyc + iyr) \ 2, CulNum, 0   ' 1x1 dots on line
      Exit Sub
   
   Else
      ' Cone outline
      ixr = STOREX(3)
      iyr = STOREY(3)
      Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
      ixc = px1
      iyc = py1
      ' Check if cone point inside base circle
         ' px1,py1 are ixc,iyc center coords
         ix1 = px1 - zrad
         iy1 = py1 + zrad
         ix2 = ix1 + 2 * zrad
         iy2 = iy1 - 2 * zrad
         BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
         Getpx2py2 STOREX(3), NSX, STOREY(3), NSY
         ixr = px2
         iyr = py2
      If (iyr - iyc) ^ 2 + (ixr - ixc) ^ 2 >= zrad ^ 2 Then
         zalpha1 = zATan2(iy1 - iyc, ix1 - ixc)
         zalpha2 = zATan2(iy2 - iyc, ix2 - ixc)
         ixp = ixc + zrad * Cos(zalpha1)
         iyp = iyc + zrad * Sin(zalpha2)
         EvalTangents ixc, iyc, zrad, ixr, iyr, ix1, iy1, ix2, iy2
         BresLine ix1, iy1, ixr, iyr, CulNum, 0   ' 1x1 dots on line
         BresLine ix2, iy2, ixr, iyr, CulNum, 0   ' 1x1 dots on line
      End If
         zdeltacn = (SelRightCulNum - SelLeftCulNum) / 720
         zCul = SelLeftCulNum
         zalpha1 = zATan2(iy1 - iyc, ix1 - ixc)
         zalpha2 = zATan2(iy2 - iyc, ix2 - ixc)
   End If
   
   Select Case ConeType
   
   Case ConeHShade1, ConeHSHade2 ' Base overwritten, not overwritten
      
      ' Check if cone point inside base circle
      If (iyr - iyc) ^ 2 + (ixr - ixc) ^ 2 < zrad ^ 2 Then
         zalpha1 = 0
         zalpha2 = 2 * pi#
         zstep = (2 * pi#) / 720
         If zstep = 0 Then zstep = 0.001
         For zalpha = zalpha1 To zalpha2 Step zstep
            ixp = ixc + zrad * Cos(zalpha)
            iyp = iyc + zrad * Sin(zalpha)
            BresLine ixp, iyp, px2, py2, CLng(zCul), 1   ' 1x1 dots on line
            zCul = zCul + zdeltacn
            If zCul < 0 Then zCul = 255
            If zCul > 255 Then zCul = 0
         Next zalpha
      Else
         Select Case ConeType
         Case ConeHShade1, ConeHSHade2 ' Base overwritten, not overwritten
            If zalpha2 > zalpha1 Then  ' Ensures base overwritten
               zalpha1 = zalpha1 + 2 * pi#
            End If
            zstep = (zalpha2 - zalpha1) / 720
            If zstep = 0 Then zstep = 0.001
            For zalpha = zalpha1 To zalpha2 Step zstep
               ixp = ixc + zrad * Cos(zalpha)
               iyp = iyc + zrad * Sin(zalpha)
               BresLine ixp, iyp, px2, py2, CLng(zCul), 1   ' 1x1 dots on line
               zCul = zCul + zdeltacn
               If zCul < 0 Then zCul = 255
               If zCul > 255 Then zCul = 0
            Next zalpha
         
            If ConeType = ConeHSHade2 Then
               ' Reverse color gradient
               zdeltacn = (SelLeftCulNum - SelRightCulNum) / 720
               zCul = SelRightCulNum
               zalpha1 = zATan2(iy1 - iyc, ix1 - ixc)
               zalpha2 = zATan2(iy2 - iyc, ix2 - ixc)
               If zalpha1 > zalpha2 Then  ' Ensures base NOT overwritten
                  zalpha2 = zalpha2 + 2 * pi#
               End If
               zstep = (zalpha2 - zalpha1) / 720
               If zstep = 0 Then zstep = 0.001
               For zalpha = zalpha1 To zalpha2 Step zstep
                  ix1 = ixc + zrad * Cos(zalpha)
                  iy1 = iyc + zrad * Sin(zalpha)
                  BresLine ix1, iy1, px2, py2, CLng(zCul), 1   ' 1x1 dots on line
                  zCul = zCul + zdeltacn
                  If zCul < 0 Then zCul = 255
                  If zCul > 255 Then zCul = 0
               Next zalpha
            End If
         End Select
      End If
         
   Case ConeCShade1, ConeCShade2
      
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / 240
      zCul = SelLeftCulNum
      If zalpha2 > zalpha1 Then  ' Ensures base overwritten
         zalpha1 = zalpha1 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 360
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ixp = ixc + zrad * Cos(zalpha)
         iyp = iyc + zrad * Sin(zalpha)
         BresLine ixp, iyp, px2, py2, CLng(zCul), 1   ' 1x1 dots on line
         If zalpha <= (zalpha1 + zalpha2) / 2 Then
            zCul = zCul - zdeltacn
         Else
            zCul = zCul + zdeltacn
         End If
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha

      If ConeType = ConeCShade2 Then
         ' Reverse color gradient
         zdeltacn = (SelLeftCulNum - SelRightCulNum) / 240
         zCul = SelRightCulNum
         zalpha1 = zATan2(iy1 - iyc, ix1 - ixc)
         zalpha2 = zATan2(iy2 - iyc, ix2 - ixc)
      End If
      
      If zalpha1 > zalpha2 Then  ' Ensures base NOT overwritten
         zalpha2 = zalpha2 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 360
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         BresLine ix1, iy1, px2, py2, CLng(zCul), 1   ' 1x1 dots on line
         If zalpha <= (zalpha1 + zalpha2) / 2 Then
            zCul = zCul + zdeltacn
         Else
            zCul = zCul - zdeltacn
         End If
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
   End Select
End Sub

Public Sub CompleteTube(CulNum As Long)
Dim zspace As Single
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ix1d As Long, iy1d As Long
Dim ix2d As Long, iy2d As Long
Dim ix3 As Long, iy3 As Long
Dim ix4 As Long, iy4 As Long
Dim ixr As Long, iyr As Long
Dim zalpha As Single
Dim zalpha1 As Single
Dim zalpha2 As Single
Dim zstep As Single

   If zrad = 0 Then Exit Sub
   
   ' Adjustment to ~ match VB Line()-()
   NSX = 0: NSY = 0: ND = 0
   ' STOREXY(1) ixc,iyc
   ' STOREXY(2) X,Y
   ' STOREXY(3) XDiag,YDiag
   ' zrad
   ixr = STOREX(3)
   iyr = STOREY(3)
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   ixc = px1
   iyc = py1
   ' ixc,iyc are center coords 1st circle
   ' px1,py1 are ixc,iyc center coords
   ix1 = px1 - zrad
   iy1 = py1 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
   Getpx2py2 STOREX(3), NSX, STOREY(3), NSY
   ixr = px2
   iyr = py2
   ' ixr,iyr are center coords 2nd circle
   ix1 = px2 - zrad
   iy1 = py2 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
   EvalDiameters ixc, iyc, zrad, ixr, iyr, ix1, iy1, ix2, iy2, ix3, iy3, ix4, iy4
   BresLine ix1, iy1, ix3, iy3, CulNum, 0   ' 1x1 dots on line
   BresLine ix2, iy2, ix4, iy4, CulNum, 0   ' 1x1 dots on line
   
   ' Save diameters
   ix1d = ix1
   iy1d = iy1
   ix2d = ix2
   iy2d = iy2
   ' ix1,iy1 ---------------------- ix3,iy3
   '  ixc,iyc                     ixr,iyr
   ' ix2,iy2 ---------------------- ix4,iy4
   
   Select Case TubeType
   Case TubeHShade
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / 720
      zCul = SelLeftCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      If zalpha2 > zalpha1 Then  ' Ensures base overwritten
         zalpha1 = zalpha1 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 720
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         ix3 = ixr + zrad * Cos(2 * zalpha1 - zalpha)
         iy3 = iyr + zrad * Sin(2 * zalpha1 - zalpha)
         BresLine ix1, iy1, ix3, iy3, CLng(zCul), 1   ' 1x1 dots on line
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
      ' Reverse color gradient
      zdeltacn = (SelLeftCulNum - SelRightCulNum) / 720
      zCul = SelRightCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      If zalpha1 > zalpha2 Then  ' Ensures base NOT overwritten
         zalpha2 = zalpha2 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 720
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         ix3 = ixr + zrad * Cos(zalpha)
         iy3 = iyr + zrad * Sin(zalpha)
         BresLine ix1, iy1, ix3, iy3, CLng(zCul), 1   ' 1x1 dots on line
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
   
   Case TubeCShade
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / 240
      zCul = SelLeftCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      If zalpha2 > zalpha1 Then  ' Ensures base overwritten
         zalpha1 = zalpha1 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 360
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         ix3 = ixr + zrad * Cos(2 * zalpha1 - zalpha)
         iy3 = iyr + zrad * Sin(2 * zalpha1 - zalpha)
         BresLine ix1, iy1, ix3, iy3, CLng(zCul), 1   ' 1x1 dots on line
         If zalpha <= (zalpha1 + zalpha2) / 2 Then
            zCul = zCul - zdeltacn
         Else
            zCul = zCul + zdeltacn
         End If
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
      ' Reverse color gradient
      zdeltacn = (SelLeftCulNum - SelRightCulNum) / 240
      zCul = SelRightCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      If zalpha1 > zalpha2 Then  ' Ensures base NOT overwritten
         zalpha2 = zalpha2 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 360
      If zstep = 0 Then zstep = 0.001
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         ix3 = ixr + zrad * Cos(zalpha)
         iy3 = iyr + zrad * Sin(zalpha)
         BresLine ix1, iy1, ix3, iy3, CLng(zCul), 1   ' 1x1 dots on line
         If zalpha <= (zalpha1 + zalpha2) / 2 Then
            zCul = zCul + zdeltacn
         Else
            zCul = zCul - zdeltacn
         End If
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
   End Select
End Sub

Public Sub CompleteBullet(CulNum As Long)
Dim zspace As Single
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ix1d As Long, iy1d As Long
Dim ix2d As Long, iy2d As Long
Dim ix3 As Long, iy3 As Long
Dim ix4 As Long, iy4 As Long
Dim ixr As Long, iyr As Long
Dim x3 As Single, y3 As Single
Dim zalpha As Single
Dim zalpha1 As Single
Dim zalpha2 As Single
Dim zstep As Single

Dim zdx As Single
Dim zdy As Single

   If zrad = 0 Then Exit Sub
   
   ' Adjustment to ~ match VB Line()-()
   NSX = 0: NSY = 0: ND = 0
   ' STOREXY(1) ixc,iyc
   ' STOREXY(2) X,Y
   ' STOREXY(3) XDiag,YDiag
   ' zrad
   ixr = STOREX(3)
   iyr = STOREY(3)
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   ixc = px1
   iyc = py1
   ' ixc,iyc are center coords 1st circle
   ' px1,py1 are ixc,iyc center coords
   ix1 = px1 - zrad
   iy1 = py1 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   BresEllipse ix1, iy1, ix2, iy2, CulNum, ND, 0, 0
   Getpx2py2 STOREX(3), NSX, STOREY(3), NSY
   ixr = px2
   iyr = py2
   ' ixr,iyr are center coords 2nd circle
   ix1 = px2 - zrad
   iy1 = py2 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   EvalDiameters ixc, iyc, zrad, ixr, iyr, ix1, iy1, ix2, iy2, ix3, iy3, ix4, iy4
   BresLine ix1, iy1, ix3, iy3, CulNum, 0   ' 1x1 dots on line
   BresLine ix2, iy2, ix4, iy4, CulNum, 0   ' 1x1 dots on line
   BresLine ix3, iy3, ix4, iy4, CulNum, 0   ' 1x1 dots on line
   ' Save diameters
   ix1d = ix1
   iy1d = iy1
   ix2d = ix2
   iy2d = iy2
   ' ix1,iy1 ---------------------- ix3,iy3
   '  ixc,iyc                     ixr,iyr
   ' ix2,iy2 ---------------------- ix4,iy4
      
   Select Case BulletType
   Case BulletHShade
   
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / 720
      zCul = SelLeftCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      zdx = 2 * zrad * Cos(zalpha2)
      zdy = 2 * zrad * Sin(zalpha2)
      If zalpha2 > zalpha1 Then  ' Ensures base overwritten
         zalpha1 = zalpha1 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 720
      If zstep = 0 Then zstep = 0.001
      zdx = zdx / ((zalpha2 - zalpha1) / zstep)
      zdy = zdy / ((zalpha2 - zalpha1) / zstep)
      x3 = ixr + zrad * Cos(zalpha1)
      y3 = iyr + zrad * Sin(zalpha1)
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         x3 = x3 + zdx
         y3 = y3 + zdy
         BresLine ix1, iy1, CLng(x3), CLng(y3), CLng(zCul), 1   ' 1x1 dots on line
         zCul = zCul + zdeltacn
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
   
   Case BulletCShade
      ' Reverse color gradient
      zdeltacn = (SelLeftCulNum - SelRightCulNum) / 240
      zCul = SelRightCulNum
      zalpha1 = zATan2(iy1d - iyc, ix1d - ixc)
      zalpha2 = zATan2(iy2d - iyc, ix2d - ixc)
      zdx = 2 * zrad * Cos(zalpha2)
      zdy = 2 * zrad * Sin(zalpha2)
      If zalpha2 > zalpha1 Then  ' Ensures base overwritten
         zalpha1 = zalpha1 + 2 * pi#
      End If
      zstep = (zalpha2 - zalpha1) / 360
      If zstep = 0 Then zstep = 0.001
      zdx = zdx / ((zalpha2 - zalpha1) / zstep)
      zdy = zdy / ((zalpha2 - zalpha1) / zstep)
      x3 = ixr + zrad * Cos(zalpha1)
      y3 = iyr + zrad * Sin(zalpha1)
      For zalpha = zalpha1 To zalpha2 Step zstep
         ix1 = ixc + zrad * Cos(zalpha)
         iy1 = iyc + zrad * Sin(zalpha)
         x3 = x3 + zdx
         y3 = y3 + zdy
         BresLine ix1, iy1, CLng(x3), CLng(y3), CLng(zCul), 1   ' 2x2 dots on line
         If zalpha <= (zalpha1 + zalpha2) / 2 Then
            zCul = zCul - zdeltacn
         Else
            zCul = zCul + zdeltacn
         End If
         If zCul < 0 Then zCul = 255
         If zCul > 255 Then zCul = 0
      Next zalpha
   End Select
End Sub

Public Sub CompleteJunction(CulNum As Long)
   For k = 1 To 12
      Getpx1py1 CLng(XT(k)), 0, CLng(YT(k)), 0
      XT(k) = px1
      YT(k) = py1
   Next k
   Select Case JunctionType
   Case TPiece1, TPiece2, TPiece3: CompleteTPiece CulNum
   Case Cross1, Cross2, Cross3: CompleteCross CulNum
   Case Corner1, Corner2, Corner3: CompleteCorner CulNum
   End Select
End Sub

Public Sub CompleteTPiece(CulNum As Long)
   BresLine XT(1), YT(1), XT(2), YT(2), CulNum, 0   ' 1x1 dots on line
   BresLine XT(3), YT(3), XT(7), YT(7), CulNum, 0
   BresLine XT(8), YT(8), XT(4), YT(4), CulNum, 0
   BresLine XT(7), YT(7), XT(11), YT(11), CulNum, 0
   BresLine XT(8), YT(8), XT(12), YT(12), CulNum, 0
End Sub

Public Sub CompleteCross(CulNum As Long)
   BresLine XT(1), YT(1), XT(5), YT(5), CulNum, 0   ' 1x1 dots on line
   BresLine XT(6), YT(6), XT(2), YT(2), CulNum, 0
   BresLine XT(3), YT(3), XT(7), YT(7), CulNum, 0
   BresLine XT(8), YT(8), XT(4), YT(4), CulNum, 0
   BresLine XT(7), YT(7), XT(11), YT(11), CulNum, 0
   BresLine XT(8), YT(8), XT(12), YT(12), CulNum, 0
   BresLine XT(5), YT(5), XT(9), YT(9), CulNum, 0
   BresLine XT(6), YT(6), XT(10), YT(10), CulNum, 0
End Sub

Public Sub CompleteCorner(CulNum As Long)
   BresLine XT(1), YT(1), XT(5), YT(5), CulNum, 0   ' 1x1 dots on line
   BresLine XT(5), YT(5), XT(9), YT(9), CulNum, 0
   BresLine XT(3), YT(3), XT(8), YT(8), CulNum, 0
   BresLine XT(8), YT(8), XT(10), YT(10), CulNum, 0
End Sub

Public Sub CompleteArc(CulNum As Long)
Dim NQ As Long
' NB:  NQ = ArcType + 1
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long

   NSX = 0: NSY = 0
   ixc = STOREX(1)
   iyc = STOREY(1)
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   If Abs(px2 - px1) > Abs(py2 - py1) Then
      zradx = zrad
      zrady = zradx * zratio
   Else
      zrady = zrad
      zradx = zrady / zratio
   End If
   If zrad > 0 Then
      NQ = ArcType + 1
      ' px1,py1 are ixc,iyc center coords
      ix1 = px1 - zradx
      iy1 = py1 + zrady
      ix2 = ix1 + 2 * zradx
      iy2 = iy1 - 2 * zrady
      BresEllipse ix1, iy1, ix2, iy2, CulNum, 0, 0, NQ
   End If
End Sub

Public Sub CompleteShapeA(CulNum As Long)
   ReDim TX(4) As Long, TY(4) As Long
   For k = 1 To 4
      Getpx1py1 STOREX(k), 0, STOREY(k), 0
      TX(k) = px1: TY(k) = py1
   Next k
   BresLine TX(1), TY(1), TX(2), TY(2), CulNum, 0   ' 1x1 dots on line
   BresLine TX(3), TY(3), TX(4), TY(4), CulNum, 0   ' 1x1 dots on line
   BresLine TX(1), TY(1), TX(3), TY(3), CulNum, 0   ' 1x1 dots on line
   BresLine TX(2), TY(2), TX(4), TY(4), CulNum, 0   ' 1x1 dots on line
Erase TX(), TY()
End Sub

Public Sub CompleteShapeB(CulNum As Long)
' Dumbell
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
' Rest Public
   ' Adjustment to ~ match VB Line()-()
   NSX = 0: NSY = 0
   Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
   Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
   BresLine px1, py1, px2, py2, CulNum, 0
   ix1 = px1 - zrad
   iy1 = py1 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   BresEllipse ix1, iy1, ix2, iy2, CulNum, 0, 0, 0
   ix1 = px2 - zrad
   iy1 = py2 + zrad
   ix2 = ix1 + 2 * zrad
   iy2 = iy1 - 2 * zrad
   BresEllipse ix1, iy1, ix2, iy2, CulNum, 0, 0, 0

End Sub

Public Sub CompleteRadials(CulNum As Long)
Dim ix1 As Long, iy1 As Long
Dim ix2 As Long, iy2 As Long
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long

NSX = 0: NSY = 0

   ReDim TX(NSTOREXY) As Long, TY(NSTOREXY) As Long
   
   For k = 1 To NSTOREXY
      Getpx1py1 STOREX(k), 0, STOREY(k), 0
      TX(k) = px1: TY(k) = py1
   Next k
   
   Select Case RadialType
   Case RSpokes
      For k = 2 To NSTOREXY
         BresLine TX(1), TY(1), TX(k), TY(k), CulNum, 0   ' 1x1 dots on line
      Next k
   Case RStars
      For k = 2 To NSTOREXY - 1
         BresLine TX(k), TY(k), TX(k + 1), TY(k + 1), CulNum, 0 ' 1x1 dots on line
      Next k
      BresLine TX(NSTOREXY), TY(NSTOREXY), TX(2), TY(2), CulNum, 0 ' 1x1 dots on line
   Case RRadCircs
      zrad2 = zrad / 5
      For k = 2 To NSTOREXY
         Getpx1py1 STOREX(k), NSX, STOREY(k), NSY
         ix1 = px1 - zrad2
         iy1 = py1 + zrad2
         ix2 = ix1 + 2 * zrad2
         iy2 = iy1 - 2 * zrad2
         BresEllipse ix1, iy1, ix2, iy2, CulNum, 0, 0, 0
      Next k
   Case RPolygons
      For k = 2 To NSTOREXY - 1
         BresLine TX(k), TY(k), TX(k + 1), TY(k + 1), CulNum, 0 ' 1x1 dots on line
      Next k
      BresLine TX(NSTOREXY), TY(NSTOREXY), TX(2), TY(2), CulNum, 0 ' 1x1 dots on line
   Case RTeeth
      For k = 3 To NSTOREXY
         If (k Mod 2) = 0 Then
            GetParallelCoords zrad / 4, CSng(TX(k)), CSng(TY(k)), CSng(TX(k - 1)), CSng(TY(k - 1)), ixa, iya, ixb, iyb
            BresLine TX(k - 1), TY(k - 1), ixb, iyb, CulNum, 0
            BresLine ixa, iya, ixb, iyb, CulNum, 0
            BresLine ixa, iya, TX(k), TY(k), CulNum, 0
         Else
            BresLine TX(k - 1), TY(k - 1), TX(k), TY(k), CulNum, 0 ' 1x1 dots on line
         End If
      Next k
      If (k Mod 2) = 0 Then
            GetParallelCoords zrad / 4, CSng(TX(2)), CSng(TY(2)), CSng(TX(NSTOREXY)), CSng(TY(NSTOREXY)), ixa, iya, ixb, iyb
            BresLine TX(NSTOREXY), TY(NSTOREXY), ixb, iyb, CulNum, 0
            BresLine ixa, iya, ixb, iyb, CulNum, 0
            BresLine ixa, iya, TX(2), TY(2), CulNum, 0
      Else
            BresLine TX(NSTOREXY), TY(NSTOREXY), TX(2), TY(2), CulNum, 0 ' 1x1 dots on line
      End If
   End Select
Erase TX(), TY()
End Sub

Public Sub CompleteBush(CulNum As Long)
Dim yd As Single
   NSX = 0: NSY = 0
   zCul = SelLeftCulNum
   yd = YTreeMax - YTreeMin
   zdeltacn = (SelRightCulNum - SelLeftCulNum) / NSTOREXY 'Abs(yd)
   For NN = 2 To NSTOREXY
      Getpx1py1 STOREX(NN - 1), NSX, STOREY(NN - 1), NSY
      Getpx2py2 STOREX(NN), NSX, STOREY(NN), NSY
      BresLine px1, py1, px2, py2, CLng(zCul), 0
      zCul = zCul + zdeltacn
      If zCul < 0 Then zCul = 255
      If zCul > 255 Then zCul = 0
   Next NN
End Sub

Public Sub CompleteArrow(CulNum As Long)
   ' Adjustment to ~ match VB Line()-()
    NSX = 0: NSY = 0
      Select Case ArrowType
      Case ArrSingle
         Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
         Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
         BresLine px1, py1, px2, py2, CulNum, 0
         
         Getpx1py1 STOREX(3), NSX, STOREY(3), NSY
         BresLine px2, py2, px1, py1, CulNum, 0
         
         Getpx1py1 STOREX(4), NSX, STOREY(4), NSY
         BresLine px2, py2, px1, py1, CulNum, 0
      Case ArrFeathered
         Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
         Getpx2py2 STOREX(2), NSX, STOREY(2), NSY
         BresLine px1, py1, px2, py2, CulNum, 0

         Getpx1py1 STOREX(3), NSX, STOREY(3), NSY
         BresLine px2, py2, px1, py1, CulNum, 0

         Getpx1py1 STOREX(4), NSX, STOREY(4), NSY
         BresLine px2, py2, px1, py1, CulNum, 0

         Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
         BresLine px1, py1, px1 - xd1, py1 + yd1, CulNum, 0
         BresLine px1, py1, px1 - xd2, py1 + yd2, CulNum, 0
      Case ArrTriangle
      
         Getpx1py1 STOREX(1), NSX, STOREY(1), NSY
         Getpx2py2 STOREX(5), NSX, STOREY(5), NSY
         BresLine px1, py1, px2, py2, CulNum, 0

         Getpx1py1 STOREX(2), NSX, STOREY(2), NSY
         Getpx2py2 STOREX(3), NSX, STOREY(3), NSY
         BresLine px1, py1, px2, py2, CulNum, 0

         Getpx2py2 STOREX(4), NSX, STOREY(4), NSY
         BresLine px1, py1, px2, py2, CulNum, 0

         Getpx1py1 STOREX(3), NSX, STOREY(3), NSY
         BresLine px1, py1, px2, py2, CulNum, 0
      End Select
End Sub

Public Sub CompleteText(CulNum As Long)
   For NN = 1 To NSTOREXY
      py1 = canvasH - STOREY(NN)
      px1 = STOREX(NN) + 1
      If px1 > 1 And px1 < canvasW Then
      If py1 > 1 And py1 < canvasH Then
         SetDot px1, py1, CulNum, 0
      End If
      End If
   Next NN
End Sub

Public Sub SmoothArea(ppx As Long, ppy As Long, N As Long)
' Smooth area iiy,iix -N -> +N
Dim NN As Long
Dim SR As Long, SG As Long, SB As Long
Dim iix As Long, iiy As Long
Dim i As Long
'In: N = 1,2,4
'0 1 0
'1 0 1
'0 1 0
   ptStanPal = VarPtr(CulRGB(0))    ' Standard
   ReDim bDummy(canvasW, canvasH)
   bDummy() = bArray()
   For iiy = ppy - N To ppy + N
   If iiy > 0 And iiy <= canvasH Then
      For iix = ppx - N To ppx + N
      NN = 0
      If iix > 0 And iix <= canvasW Then
         SR = 0: SG = 0: SB = 0
         If iiy > 1 Then
            k = bArray(iix, iiy - 1)
            SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
            NN = NN + 1
         End If
         If iix > 1 Then
            k = bArray(iix - 1, iiy)
            SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
            NN = NN + 1
         End If
         If iix < canvasW Then
            k = bArray(iix + 1, iiy)
            SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
            NN = NN + 1
         End If
         If iiy < canvasH Then
            k = bArray(iix, iiy + 1)
            SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
            NN = NN + 1
         End If
         If NN > 0 Then
            SR = SR / (NN)
            SG = SG / (NN)
            SB = SB / (NN)
            'GetIndex
            LongDerived = RGB(SR, SG, SB)
            i = CallWindowProc(ptMC, LongDerived, ptStanPal, 3&, 4&)
            'bArray(iix, iiy) = i
            bDummy(iix, iiy) = i
         End If
      End If
      Next iix
   End If
   Next iiy
   bArray() = bDummy()
   Erase bDummy
End Sub

' Heavier
' 1 1 1
' 1 0 1
' 1 1 1
' Dim j As Long
'         For j = iiy - 1 To iiy + 1
'         If j > 0 And j <= canvasH Then
'            For i = iix - 1 To iix + 1
'            If i > 0 And i <= canvasW Then
'               If (j <> iiy Or i <> iix) Then
'                   k = bArray(i, j)
'                   SR = SR + palRed(k): SG = SG + palGreen(k): SB = SB + palBlue(k)
'                   NN = NN + 1
'               End If
'            End If
'            Next i
'         End If
'         Next j


'#####################################################################################################
'#####################################################################################################

Public Sub SetDot(ppx As Long, ppy As Long, Cul As Long, ND As Long)
' Sets color num Cul at ppx,ppy
' ND = 0,1,2,,    1,2x2,4x4,,  dots
Dim ixx As Long
Dim iyy As Long
   If ppx > 0 Then
   If ppy > 0 Then
      Select Case ND
      Case 0
         If ppx <= canvasW Then
         If ppy <= canvasH Then
            bArray(ppx, ppy) = Cul
         End If
         End If
      Case 1
         If ppx <= canvasW Then
         If ppy <= canvasH Then
            bArray(ppx, ppy) = Cul
            If ppx < canvasW Then
               bArray(ppx + 1, ppy) = Cul
               If ppy < canvasH Then
                  bArray(ppx + 1, ppy + 1) = Cul
                  bArray(ppx, ppy + 1) = Cul
               End If
            End If
         End If
         End If
      Case 2
         If ppx + 3 <= canvasW Then
         If ppy + 3 <= canvasH Then
            For iyy = 0 To 3
            For ixx = 0 To 3
               bArray(ppx + ixx, ppy + iyy) = Cul
            Next ixx
            Next iyy
         End If
         End If
      End Select
   End If
   End If
End Sub

Public Sub BresBox(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long, ND As Long)
   BresLine ix1, iy1, ix2, iy1, Cul, ND
   BresLine ix2, iy1, ix2, iy2, Cul, ND
   BresLine ix2, iy2, ix1, iy2, Cul, ND
   BresLine ix1, iy2, ix1, iy1, Cul, ND
End Sub

Public Sub BresLine(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long, ND As Long)
'** Public bArray(), canvasW, canvasH
'** BASIC Bresenham Line for drawing into a 2D Public
'** Byte Array (bArray()) with a color index Cul (256 palette)
'** ND 0,1,2   1x1 2x2, 4x4 dots
'** Plus clipping on 1->canvasW, 1->canvasH

Dim iix As Long, iiy As Long
Dim idx As Long, idy As Long
Dim jkstep As Long
Dim incx As Long
Dim id As Long
Dim ainc As Long, binc As Long
Dim jj As Long, kk As Long

   ' Reject lines outside bArray
   If ix1 > 0 Or ix2 > 0 Then
   If ix1 <= canvasW Or ix2 <= canvasW Then
   If iy1 > 0 Or iy2 > 0 Then
   If iy1 <= canvasH Or iy2 <= canvasH Then
      
      idx = Abs(ix2 - ix1)
      idy = Abs(iy2 - iy1)
      jkstep = 1
      incx = 1
      If idx < idy Then   '-- Steep slope
         
         If iy1 > iy2 Then jkstep = -1
         If ix2 < ix1 Then incx = -1
         id = 2 * idx - idy
         ainc = 2 * (idx - idy)   '-ve
         binc = 2 * idx
         jj = iy1: kk = iy2: iix = ix1
      
         For iiy = jj To kk Step jkstep
            ' Reject any point outside bArray
            If iix > 0 Then
            If iix <= canvasW Then
            If iiy > 0 Then
            If iiy <= canvasH Then
               SetDot iix, iiy, Cul, ND
               'bArray(iix, iiy) = Cul
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               iix = iix + incx
            Else
               id = id + binc
            End If
         Next iiy
      
      Else                '-- Shallow slope
         
         If ix1 > ix2 Then jkstep = -1
         If iy2 < iy1 Then incx = -1
         id = 2 * idy - idx
         ainc = 2 * (idy - idx)   '-ve
         binc = 2 * idy
         jj = ix1: kk = ix2: iix = iy1
      
         For iiy = jj To kk Step jkstep
            ' Reject any point outside bArray
            If iiy > 0 Then
            If iiy <= canvasW Then
            If iix > 0 Then
            If iix <= canvasH Then
               SetDot iiy, iix, Cul, ND
               'bArray(iiy, iix) = Cul
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               iix = iix + incx
            Else
               id = id + binc
            End If
         Next iiy
      
      End If
   
   End If
   End If
   End If
   End If
End Sub

Public Sub BresEllipse(ixTL As Long, iyTL As Long, ixBR As Long, iyBR As Long, _
   CulNum As Long, ND As Long, NDSP As Long, NQ As Long)
'** BASIC Bresenham Ellipse or Circle
' Y increasing downwards
' Evaluates Bottom-Left (BL) 1/4 ellipse
' & fills in rest by offsets

' ND dot size 0,1,2 (1,2,4), NDSP dot spacing 2,4,8

' NQ 0 (all), 1 TL,     2 BL,     3 TR,     4 BR,
'       semis 5 1&3 T,  6 2&4 B,  7 1&2 L,  8 3&4 R,
'    3/4 circ 9 1&2&3, 10 1&2&4, 11 1&3&4, 12 2&3&4

Dim ixL As Long, iyL As Long
Dim ixH As Long, iyH As Long
Dim iyoff As Long
Dim kza As Long, kzb As Long
Dim ix As Long, iy As Long
Dim asq As Long, bsq As Long
Dim a22 As Long, b22 As Long
Dim a42 As Long, b42 As Long
Dim xslope As Single, yslope As Single
Dim midla As Long, midlb As Long
Dim d As Long
Dim Dotted As Long
Dim moddot As Long

   ' Centre coords
   'ixc = (ixTL + ixBR) / 2
   'iyc = (iyTL + iyBR) / 2

   ' Get BL & BH coords
   ixL = ixTL
   iyL = (iyTL + iyBR) / 2
   ixH = (ixTL + ixBR) / 2
   iyH = iyBR

   ' Y increasing downwards
   'iyoff = iyL
   'If iyH < iyL Then iyoff = iyH
   
   ' Y increasing upwards
   iyoff = iyH
   If iyH < iyL Then iyoff = iyL
   
   kza = Abs(ixL - ixH)
   'NB a circle when kza = kzb
   'kzb = kza 'Abs(iyH - iyL)
   kzb = Abs(iyH - iyL)
   
   ix = 0
   iy = kzb
   asq = kza * 1 * kza * 1
   bsq = (kzb * 1) * (kzb * 1)
   a22 = asq + asq
   b22 = bsq + bsq
   a42 = (a22 + a22)
   b42 = (b22 + b22)
   xslope = b42
   yslope = a42 * (iy - 1)
   midla = asq / 2
   midlb = bsq / 2
   d = b22 - asq - midla - yslope / 2
   Dotted = 0
   'Region 1 dy/dx > -1 ie more horizontal
   Do While d <= yslope
      moddot = 0
      If NDSP <> 0 Then moddot = (Dotted Mod NDSP)
      If moddot = 0 Then
         Set4Dots ix, iy, ixH, ixL, iyoff, CulNum, ND, NQ
      End If
      If d > 0 Then
         d = d - yslope
         iy = iy - 1
         yslope = yslope - a42
      End If
      d = d + b22 + xslope
      ix = ix + 1
      xslope = xslope + b42
      Dotted = Dotted + 1
   Loop
   d = d - (xslope + yslope) / 2 + (bsq - asq) + (midla - midlb)
   Dotted = 0
   'Region 2 dy/dx <= -1   ie more vertical
   Do While iy >= 0
      moddot = 0
      If NDSP <> 0 Then moddot = (Dotted Mod NDSP)
      If moddot = 0 Then
         Set4Dots ix, iy, ixH, ixL, iyoff, CulNum, ND, NQ
      End If
      If d <= 0 Then
         d = d + xslope
         ix = ix + 1
         xslope = xslope + b42
      End If
      d = d + a22 - yslope
      iy = iy - 1
      yslope = yslope - a42
      Dotted = Dotted + 1
   Loop
End Sub

Public Sub Set4Dots(ix As Long, iy As Long, ixH As Long, ixL As Long, iyoff As Long, _
   CulNum As Long, ND As Long, NQ As Long)
' For BresEllipse
' ND dot size 0,1,2 (1x1,2x2,4x4)
   Select Case NQ
   Case 0
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
         SetDot ixH + ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   Case 1
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
      End If
   Case 2
      If ixH > ixL Then
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
      Else
         SetDot ixH + ix, iyoff - iy, CulNum, ND
      End If
   Case 3
      If ixH > ixL Then
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
      Else
         SetDot ixH - ix, iy + iyoff, CulNum, ND
      End If
   Case 4
      If ixH > ixL Then
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   
   Case 5
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
         SetDot ixH + ix, iyoff - iy, CulNum, ND
      End If
   Case 6
      If ixH > ixL Then
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH - ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   Case 7
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
      End If
   Case 8
      If ixH > ixL Then
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH + ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   
   Case 9
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
         SetDot ixH + ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
      End If
   Case 10
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH + ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
      End If
   Case 11
      If ixH > ixL Then
         SetDot ixH - ix, iy + iyoff, CulNum, ND   ' TL
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH + ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   Case 12
      If ixH > ixL Then
         SetDot ixH - ix, iyoff - iy, CulNum, ND   ' BL
         SetDot ixH + ix, iy + iyoff, CulNum, ND   ' TR
         SetDot ixH + ix, iyoff - iy, CulNum, ND   ' BR
      Else
         SetDot ixH + ix, iyoff - iy, CulNum, ND
         SetDot ixH - ix, iy + iyoff, CulNum, ND
         SetDot ixH - ix, iyoff - iy, CulNum, ND
      End If
   End Select
End Sub

Public Sub PatternFiller(bA() As Byte, ixp As Long, iyp As Long, Fillcn As Long)
' Fills points, in a connected area, having the same
' color as the starting color, with color Fillcn.
' Public bPattern(16,16) filled in frmToolBox
' bA() 2D byte array       eg bArray
' ixp,iyp start fill point eg cursor point
' Fillcn fill color        eg Left paint color number
' Public FillType
Dim ix As Long, iy As Long
Dim N As Long
Dim Culp As Long
Dim bW As Long
Dim bH As Long
   
   If FillType = 0 Then
      ' Prep if FillType=Fill21 or 22
      ixpmin = 10000
      ixpmax = -10000
      iypmin = 10000
      iypmax = -10000
   End If
   
'Dim zdeltacn As Single  '= [SelRightCulNum - SelLeftCulNum] / span
   
   If FillType = Fill21 Then        ' Horz shading Vert bands
      zdc = (ixpmax - ixpmin)
      If zdc = 0 Then zdc = 100
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / zdc
      'CulNum= SelLeftCulNum+(ixp-ixpmin)*zdeltacn
   ElseIf FillType = Fill22 Then    ' Vert shading Horz bands
      zdc = (iypmax - iypmin)
      If zdc = 0 Then zdc = 100
      zdeltacn = (SelRightCulNum - SelLeftCulNum) / zdc
      'CulNum= SelLeftCulNum+(iyp-iypmin)*zdeltacn
   End If
   
   bW = UBound(bA(), 1)
   bH = UBound(bA(), 2)

   nspan = 16
   If FillType = Fill11 Then nspan = 64

' To hold filled point coords
   ReDim ix2(bW * bH)
   ReDim iy2(bW * bH)

   ReDim bMark(bW, bH)
   bMark() = bA()
   N = 0
   Culp = bA(ixp, iyp)
   If Culp = Fillcn Then Exit Sub
   
   Do
      PatternFillPixel bA(), ixp, iyp, Culp, Fillcn, N, ix2(), iy2()    ' 0,0
      ixp = ixp + 1
      If ixp <= bW Then
         PatternFillPixel bA(), ixp, iyp, Culp, Fillcn, N, ix2(), iy2() ' +1,0
      End If
      iyp = iyp + 1
      ixp = ixp - 1
      If iyp <= bH Then
      If ixp > 0 Then
         PatternFillPixel bA(), ixp, iyp, Culp, Fillcn, N, ix2(), iy2() ' 0,+1
      End If
      End If
      ixp = ixp - 1
      iyp = iyp - 1
      If ixp > 0 Then
      If iyp > 0 Then
         PatternFillPixel bA(), ixp, iyp, Culp, Fillcn, N, ix2(), iy2() ' -1,0
      End If
      End If
      iyp = iyp - 1
      ixp = ixp + 1
      If iyp > 0 Then
      If ixp <= bW Then
         PatternFillPixel bA(), ixp, iyp, Culp, Fillcn, N, ix2(), iy2() ' 0,-1
      End If
      End If
      ixp = ixp + 1
      ixp = ix2(N): iyp = iy2(N)
      N = N - 1
   Loop Until N = 0
   
   '--------------------------------------------------
Erase ix2(), iy2()
Erase bMark()
End Sub


Private Sub PatternFillPixel(bA() As Byte, ixp As Long, iyp As Long, Culp As Long, _
   Fillcn As Long, N As Long, sx() As Long, sy() As Long)
Dim CNum As Long
   If bMark(ixp, iyp) = Culp Then
      bMark(ixp, iyp) = Fillcn
      N = N + 1
      sx(N) = ixp: sy(N) = iyp
      ix = (ixp Mod nspan) + 1
      iy = (iyp Mod nspan) + 1
      
      If FillType = Fill21 Then        ' Horz shading Vert bands
         'zdeltacn = (SelRightCulNum - SelLeftCulNum) / (ixpmax - ixpmin)
         zdc = (ixp - ixpmin) * zdeltacn
         CNum = SelLeftCulNum + zdc
         If zdc < 0 Then
            If CNum < 0 Then CNum = 255
         Else
            If CNum > 255 Then CNum = 0
         End If
         bA(ixp, iyp) = CNum
      ElseIf FillType = Fill22 Then    ' Vert shading Horz bands
         'zdeltacn = (SelRightCulNum - SelLeftCulNum) / (iypmax - iypmin)
         zdc = (iyp - iypmin) * zdeltacn
         CNum = SelLeftCulNum + zdc
         If zdc < 0 Then
            If CNum < 0 Then CNum = 255
         Else
            If CNum > 255 Then CNum = 0
         End If
         bA(ixp, iyp) = CNum
      Else
      
         If FillType = 0 Then
            If ixp < ixpmin Then ixpmin = ixp
            If ixp > ixpmax Then ixpmax = ixp
            If iyp < iypmin Then iypmin = iyp
            If iyp > iypmax Then iypmax = iyp
         End If
         
         If bPattern(ix, iy) = 1 Then
            bA(ixp, iyp) = Fillcn
            ' Tartan (overflow?)
'            Fillcn = Fillcn - 1
'            If Fillcn < 0 Then Fillcn = 255
         End If
      
      End If
   End If
End Sub


Public Sub GetParallelCoords(zspace As Single, x1 As Single, y1 As Single, x2 As Single, y2 As Single, _
   ixa As Long, iya As Long, ixb As Long, iyb As Long)

'  GetParallelCoords zSpace,X1,Y1,X2,Y2,ixa,iya,ixb,iyb

' IN:  zspace,X1,Y1,X2,Y2
' OUT:        ixa,iya,ixb,iyb   to left of line X1,Y1->X2,Y2

'              X2,Y2
'                /  ixb,iyb
'               /  /
'              /  /
'             /  /
'            /  /
'           /  /
'          /  /
'         /  /
'  X1,Y1 /  /
'          /ixa,iya

Dim xd As Single, yd As Single
Dim xdd As Single, ydd As Single
Dim zalpha As Single
   xd = x2 - x1
   yd = y2 - y1
   If xd = 0 And yd = 0 Then
      ixa = x1: iya = y1
      ixb = x1: iyb = y2
   Else
      zalpha = zATan2(yd, xd)
      xdd = zspace * Sin(zalpha)
      ydd = zspace * Cos(zalpha)
      ixa = x1 + xdd: iya = y1 - ydd
      ixb = x2 + xdd: iyb = y2 - ydd
   End If
End Sub

Public Sub GetIntersection(zspace As Single, x1 As Single, y1 As Single, x2 As Single, y2 As Single, x3 As Single, y3 As Single, _
                           ixa As Long, iya As Long, ixb As Long, iyb As Long, _
                           ixc As Long, iyc As Long, ixd As Long, iyd As Long, _
                           ix1 As Long, iy1 As Long, ix2 As Long, iy2 As Long, _
                           ix3 As Long, iy3 As Long)

' NB using Public canvasW & canvasH ( but not essential)

' GetIntersection zspace,X1,Y1,X2,Y2,X3,Y3,ixa,iya,ixb,iyb,ixc,iyc,ixd,iyd,ix1,iy1,ix2,iy2

'IN:  zspace, X1,Y1 -> X2,Y2 -> X3,Y3
'OUT:         -- (ix1,iy1)(ix2,iy2) -------  of intersection of lines parallel to input lines
'   (ix1,iy1)(ix2,iy2) are the same unless outside 0->canvasW- 1, 0->canvasH-1
'          whence gives intersections with boundary.
'
'              X2,Y2
'        ixc,iyc /\ ixb,iyb
'               \  /
'              / \  \
'             /  /\<-----------(ix1,iy1)(ix2,iy2)
'            /  /  \  \
'           /  /    \  \
'          /  /      \  \X3,Y3
'         /  /     ixd,iyd
'  X1,Y1 /  /
'          /ixa,iya

' Slopes & Cuts
Dim zm1 As Single, zc1 As Single
Dim zm2 As Single, zc2 As Single
Dim xi As Single, yi As Single
   
   GetParallelCoords zspace, x1, y1, x2, y2, ixa, iya, ixb, iyb
   If x2 - x1 <> 0 Then
      zm1 = (y2 - y1) / (x2 - x1)
   Else
      zm1 = 10000 * Sgn(y2 - y1)
   End If
   zc1 = iya - zm1 * ixa
   
   GetParallelCoords zspace, x2, y2, x3, y3, ixc, iyc, ixd, iyd
   If x3 - x2 <> 0 Then
      zm2 = (y3 - y2) / (x3 - x2)
   Else
      zm2 = 10000 * Sgn(y3 - y2)
   End If
   zc2 = iyd - zm2 * ixd
   
   If zm1 = zm2 Then
      If zm1 = 0 Then
         yi = y1
         If x2 > x1 Then
            xi = canvasW
         Else
            xi = -1
         End If
      ElseIf x1 = x2 Then
         xi = x1
         If y2 > y1 Then
            yi = canvasH
         Else
            yi = -1
         End If
      Else
         xi = ixc
         yi = iyc
      End If
   Else
      xi = (zc2 - zc1) / (zm1 - zm2)
      yi = zm1 * xi + zc1
   End If
   
   ix1 = xi: iy1 = yi
   ix2 = xi: iy2 = yi
   ' Possible out of bounds intersection point
   ix3 = xi: iy3 = yi
   
   If ix3 > 10000 Then ix3 = canvasW + 100
   If ix3 < -10000 Then ix3 = -100
   If iy3 > 10000 Then iy3 = canvasH + 100
   If iy3 < -10000 Then iy3 = -100
   
   ' ix1,iy1, ix2,iy2 intersection with boundaries
   ' 0,0 -> W-1,H-1
   ' 1st parallel line y=zm1+zc1
   ' 2nd parallel line y=zm2+zc2
   ' If xi > W-1 then  ' Find intersection of y=zm1,zc1 with x=PICW-1 -> ix1,iy1
   '                     Find intersection of y=zm2,zc2 with x=PICW-1 -> ix2,iy2
      If xi > canvasW - 1 Then
         ix1 = canvasW - 1
         iy1 = zm1 * ix1 + zc1
         ix2 = canvasW - 1
         iy2 = zm2 * ix2 + zc2
      End If
   ' If xi < 0 then    ' Find intersection of y=zm1,zc1 with x=0 -> ix1,iy1
   '                     Find intersection of y=zm2,zc2 with x=0 -> ix2,iy2
      If xi < 0 Then
         ix1 = 0
         iy1 = zc1
         ix2 = 0
         iy2 = zc2
      End If
   ' If yi > H-1 then  ' Find intersection of y=zm1,zc1 with y=PICH-1 -> ix1,iy1
   '                     Find intersection of y=zm2,zc2 with y=PICH-1 -> ix2,iy2
      'If yi > picH - 1 Then
      If yi > canvasH - 1 Then
         iy1 = canvasH - 1
         If zm1 <> 0 Then
            ix1 = (iy1 - zc1) / zm1
         Else
            ix1 = 0
         End If
         iy2 = canvasH - 1
         If zm2 <> 0 Then
            ix2 = (iy2 - zc2) / zm2
         Else
            ix2 = 0
         End If
      End If
   ' If yi < 0 then    ' Find intersection of y=zm1,zc1 with y=0 -> ix1,iy1
   '                     Find intersection of y=zm2,zc2 with y=0 -> ix2,iy2
      If yi < 0 Then
         iy1 = 0
         If zm1 <> 0 Then
            ix1 = (iy1 - zc1) / zm1
         Else
            ix1 = 0
         End If
         iy2 = 0
         If zm2 <> 0 Then
            ix2 = (iy2 - zc2) / zm2
         Else
            ix2 = 0
         End If
   End If
End Sub

Public Sub EvalZradZratio(x As Single, y As Single)
'Public ixc As Single, iyc As Single   ' Center
'Public zrad As Single, zratio As Single
'Public zradx As Single, zrady As Single
' OUT: zrad, zratio, zradx, zrady
' For eg: PIC.Circle (ixc, iyc), zrad, DCul, , , zratio
   zradx = Abs(x - ixc)
   zrady = Abs(y - iyc)
   If zradx = 0 Then
      zrad = zrady
      zratio = 10
   ElseIf zradx >= zrady Then
      zrad = zradx
      zratio = zrady / zradx
   Else   'zradx<zrady
      zrad = zrady
      zratio = zrady / zradx
   End If
End Sub

Public Sub EvalTangents(kxc As Long, kyc As Long, zradius As Single, kxp As Long, kyp As Long, _
                        kx1 As Long, ky1 As Long, kx2 As Long, ky2 As Long)
'IN: Circle centre = kxc,kyc  radius = zradius, point outside = kxp,kyp
'OUT: Tangents kxp,kyp to kx1,ky1 & kx2,ky2
Dim zL As Single, zL2 As Single, zD As Single
zL2 = (kxp - kxc) * (kxp - kxc) + (kyp - kyc) * (kyp - kyc)
zL = Sqr(zL2)
If zL > zradius Then
   zD = Sqr(zL2 - zradius * zradius)
   kx1 = kxp - zD * ((kxp - kxc) * zD - (kyp - kyc) * zradius) / zL2
   ky1 = kyp - zD * ((kyp - kyc) * zD + (kxp - kxc) * zradius) / zL2
   kx2 = kxp - zD * ((kxp - kxc) * zD + (kyp - kyc) * zradius) / zL2
   ky2 = kyp + zD * (-(kyp - kyc) * zD + (kxp - kxc) * zradius) / zL2
Else  'xp,yp inside circle
   kx1 = kxp: ky1 = kyp
   kx2 = kxc: ky2 = kyc
End If
End Sub

Public Sub EvalDiameters(kxs As Long, kys As Long, zradius As Single, kxp As Long, kyp As Long, _
      kx1 As Long, ky1 As Long, kx2 As Long, ky2 As Long, kx3 As Long, ky3 As Long, kx4 As Long, ky4 As Long)
'IN: Circ1: kxs,kys,zradius  Circ2: kxp,kyp,zradius both same radius
'OUT: Circ1 diam cords: kx1,ky1 -> kx3,ky3  Circ2 diam coords kx2,ky2 -> kx4,ky4
Dim ztheta As Single
ztheta = zATan2(kyp - kys, kxp - kxs)
kx1 = kxs + zradius * Sin(ztheta)
ky1 = kys - zradius * Cos(ztheta)
kx2 = kxs - zradius * Sin(ztheta)
ky2 = kys + zradius * Cos(ztheta)

kx3 = kxp + zradius * Sin(ztheta) - 1.5 * Cos(ztheta) 'more accurate diameter coords
ky3 = kyp - zradius * Cos(ztheta) - 1.5 * Sin(ztheta) 'for this point
kx4 = kxp - zradius * Sin(ztheta)
ky4 = kyp + zradius * Cos(ztheta)

End Sub

Public Sub Get12TPieces(x1 As Single, y1 As Single, x2 As Single, y2 As Single, _
                        XT() As Single, YT() As Single)
'Public zspace,zTL
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim xh As Single, yh As Single
Dim zL As Single
Dim zdx1 As Single, zdy1 As Single
Dim zdx2 As Single, zdy2 As Single

' IN: 2 line end points
' OUT:
'     xT(), yT()
'      11| |12
'        | |
'        | |
'       7| |8
'3-------   --------4
'
'1-------   --------2
'       5| |6
'        | |
'        | |
'       9| |10
   xh = (x1 + x2) / 2
   yh = (y1 + y2) / 2
   zL = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
   If zL > 0 Then
      zdx1 = (1 - zspace / zL) * (xh - x1)
      zdy1 = (1 - zspace / zL) * (yh - y1)
      zdx2 = (1 + zspace / zL) * (xh - x1)
      zdy2 = (1 + zspace / zL) * (yh - y1)
   End If
   XT(1) = x1: YT(1) = y1
   XT(2) = x2: YT(2) = y2
   GetParallelCoords zspace, XT(1), YT(1), XT(2), YT(2), ixa, iya, ixb, iyb
   XT(3) = ixa: YT(3) = iya
   XT(4) = ixb: YT(4) = iyb
   XT(5) = XT(1) + zdx1
   YT(5) = YT(1) + zdy1
   XT(6) = XT(1) + zdx2
   YT(6) = YT(1) + zdy2
   XT(7) = XT(3) + zdx1
   YT(7) = YT(3) + zdy1
   XT(8) = XT(3) + zdx2
   YT(8) = YT(3) + zdy2
   GetParallelCoords zspace + zTL, XT(1), YT(1), XT(2), YT(2), ixa, iya, ixb, iyb
   XT(11) = ixa + zdx1
   YT(11) = iya + zdy1
   XT(12) = ixa + zdx2
   YT(12) = iya + zdy2
   GetParallelCoords zTL, XT(2), YT(2), XT(1), YT(1), ixa, iya, ixb, iyb
   XT(9) = ixb + zdx1
   YT(9) = iyb + zdy1
   XT(10) = ixb + zdx2
   YT(10) = iyb + zdy2
End Sub


'####################################################################

Public Sub ReflectSelectLR()
'Public SSX,SSY,SSW,SSH
Dim ixb As Long, iyb As Long
Dim ixa As Long
      ReDim bDummy(SSW, SSH)
      Getpx1py1 SSX, 0, SSY, 0
      iyb = py1 - SSH + 1
      For iy = 1 To SSH
            ixb = px1 + 1
            ixa = SSW
            For ix = 1 To SSW
               If bMask(ix, iy) <> 0 Then
                  If iyb > 0 Then
                  If iyb <= canvasH Then
                  If ixb > 0 Then
                  If ixb <= canvasW Then
                     bDummy(ixa, iy) = bArray(ixb, iyb)
                  End If
                  End If
                  End If
                  End If
               End If
               ixb = ixb + 1
               ixa = ixa - 1
            Next ix
         iyb = iyb + 1
      Next iy
      bPic() = bDummy()
      Erase bDummy()
End Sub

Public Sub ReflectSelectUD()
'Public SSX,SSY,SSW,SSH
Dim ixb As Long, iyb As Long
Dim iya As Long
      ReDim bDummy(SSW, SSH)
      Getpx1py1 SSX, 0, SSY, 0
      iyb = py1 - SSH + 1
      iya = SSH
      For iy = 1 To SSH
            ixb = px1 + 1
            For ix = 1 To SSW
               If bMask(ix, iy) <> 0 Then
                  If iyb > 0 Then
                  If iyb <= canvasH Then
                  If ixb > 0 Then
                  If ixb <= canvasW Then
                     bDummy(ix, iya) = bArray(ixb, iyb)
                  End If
                  End If
                  End If
                  End If
               End If
               ixb = ixb + 1
            Next ix
         iyb = iyb + 1
         iya = iya - 1
      Next iy
      bPic() = bDummy()
      Erase bDummy()
End Sub

Public Sub GetbPic()
' From bMask() mask & bArray()
' extract pic to bPic() where bMask()<>0
' ie  MakeMask
'     ReDim bMask(SSW, SSH)
'     GetPICBytes picMask.Image, bMask(), SSW, SSH
'     GetbPic
Dim ixb As Long, iyb As Long
   Getpx1py1 SSX, 0, SSY, 0
   ReDim bPic(SSW, SSH)
   iyb = 1
   For iy = py1 - SSH + 1 To py1
      If iy > 0 Then
      If iy <= canvasH Then
         ixb = 1
         For ix = px1 To px1 + SSW - 1
            If ix <= canvasW Then
               If bMask(ixb, iyb) <> 0 Then
                  bPic(ixb, iyb) = bArray(ix, iy)
                  If ToolType = SCut Then bArray(ix, iy) = 0
               End If
            End If
            ixb = ixb + 1
         Next ix
      End If
      End If
      iyb = iyb + 1
   Next iy
End Sub
   
Public Sub InsertbPic()   ' Into bArray()
' Insert Rect,Circ, Ellipse & Lasso
' copy all bPic() where bPic() <> culnum = 0
'Public SSX,SSY,SSW,SSH
'ReDim bMask(SSW, SSH)
Dim ixb As Long, iyb As Long
   Getpx1py1 SSX, 0, SSY, 0
   'iyb = py1 - SSH + 1  ' py1-SSH to py1
   iyb = py1 - SSH + 3 ' py1-SSH to py1
   For iy = 1 To SSH
      If iy > UBound(bPic(), 2) Then Exit For
      'ixb = px1
      ixb = px1 - 2
      For ix = 1 To SSW
         If ix > UBound(bPic(), 1) Then Exit For
         If bPic(ix, iy) <> 0 Then
            If ixb > 0 Then
            If ixb <= canvasW Then
            If iyb > 0 Then
            If iyb <= canvasH Then
               
               bArray(ixb, iyb) = bPic(ix, iy)
            
            End If
            End If
            End If
            End If
         End If
         ixb = ixb + 1
      Next ix
      iyb = iyb + 1
   Next iy
End Sub

Public Sub RotateEllbRect(zangle As Single)
' Used by Strip
' Public zangRotCSEL as single   ' degrees
' Public ixc As Long, iyc As Long
Dim ixd As Long, iyd As Long
Dim Xs As Single, Ys As Single
Dim ixs As Long, iys As Long
Dim idx As Single 'Long
Dim zrad As Single
Dim zang As Single
Dim zcos As Single, zsin As Single
   ReDim bDummy(SSW, SSH) As Byte
   'zang = pi# / 4   ' ve Clockwise
   zang = zangle * d2r#
   zcos = Cos(zang)
   zsin = Sin(-zang)
   ixc = SSW \ 2: iyc = SSH \ 2
   zrad = SSW / 2
   For iyd = 1 To SSH 'Step 0.25
      idx = Sqr(Abs(zrad ^ 2 - (iyd - iyc) ^ 2))
      For ixd = ixc - idx To ixc + idx - 1
         If ixd >= 1 Then
         If ixd <= SSW Then
            Xs = ixc + CSng(ixd - ixc) * zcos + CSng(iyd - iyc) * zsin
            Ys = iyc + CSng(iyd - iyc) * zcos - CSng(ixd - ixc) * zsin
            ixs = CLng(Xs)
            iys = CLng(Ys)
            If ixs >= 1 Then
            If ixs <= SSW Then
            If iys >= 1 Then
            If iys <= SSH Then
               If bMask(ixs, iys) <> 0 Then
                  bDummy(ixd, iyd) = bPic(ixs, iys)
                  'If bPic(ixs, iys) <> 0 Then Stop
               End If
            End If
            End If
            End If
            End If
         End If
         End If
      Next ixd
   Next iyd
   bPic() = bDummy()
   Erase bDummy()
End Sub

Public Sub Rotator()
' Rotate 90 deg
Dim iyb As Long
   If ASELECTION Then
      ReDim bDummy(SSH, SSW) As Byte
      For iy = 1 To SSH
      For ix = 1 To SSW
         iyb = SSW - ix + 1
         bDummy(iy, iyb) = bPic(ix, iy)
      Next ix
      Next iy
      ClearEm
      ix = SSH
      SSH = SSW
      SSW = ix
      ReDim bPic(SSW, SSH) As Byte
      bPic() = bDummy()
      InsertbPic
      Erase bDummy()
      If aSelRect Then
         Form1.shpRect.Width = SSW
         Form1.shpRect.Height = SSH
      ElseIf aSelEllip Then
         Form1.shpEllip.Width = SSW
         Form1.shpEllip.Height = SSH
      ElseIf aSelLasso Then
         Form1.SL(0) = Form1.SL(0)
      End If
   
   Else
      ReDim bDummy(canvasH, canvasW) As Byte
      For iy = 1 To canvasH
      For ix = 1 To canvasW
         bDummy(iy, canvasW - ix + 1) = bArray(ix, iy)
      Next ix
      Next iy
      ix = canvasH
      canvasH = canvasW
      canvasW = ix
      ReDim bArray(canvasW, canvasH)
      bArray() = bDummy()
      Erase bDummy()
   End If
End Sub

Public Sub ClearEm()
'Public SSX,SSY,SSW,SSH
Dim ixb As Long, iyb As Long
   Getpx1py1 SSX, 0, SSY, 0
   iyb = py1 - SSH + 1
   For iy = 1 To SSH
      ixb = px1 '- 1
      For ix = 1 To SSW
         If bMask(ix, iy) <> 0 Then
            If iyb > 0 Then
            If iyb <= canvasH Then
            If ixb > 0 Then
            If ixb <= canvasW Then
               bArray(ixb, iyb) = 0
            End If
            End If
            End If
            End If
         End If
         ixb = ixb + 1
      Next ix
   iyb = iyb + 1
   Next iy
End Sub

Public Sub MixEm()
Dim ixb As Long, iyb As Long
Dim CulAv As Long

   If ASELECTION Then
      ReDim bPic(SSW, SSH)
      Getpx1py1 SSX, 0, SSY, 0
      iyb = py1 - SSH + 1
      For iy = 2 To SSH - 1
            ixb = px1 + 1
            For ix = 2 To SSW - 1
               If bMask(ix, iy) <> 0 Then
                  If iyb > 0 Then
                  If iyb <= canvasH Then
                  If ixb > 0 Then
                  If ixb <= canvasW Then
                     CulAv = (1& * bArray(ixb - 1, iyb) + bArray(ixb + 1, iyb) + _
                                   bArray(ixb, iyb + 1) + bArray(ixb, iyb - 1)) / 3.8
                     If CulAv > 0 Then
                        If CulAv >= 255 Then CulAv = 32 + Rnd * 200
                        bPic(ix, iy) = CulAv
                     Else
                        bPic(ix, iy) = bArray(ixb, iyb)
                     End If
                  End If
                  End If
                  End If
                  End If
               End If
               ixb = ixb + 1
            Next ix
         iyb = iyb + 1
      Next iy
      InsertbPic
   Else
      ReDim bDummy(canvasW, canvasH)
      
      For iy = 2 To canvasH - 1
      For ix = 2 To canvasW - 1
         CulAv = (1& * bArray(ix - 1, iy) + bArray(ix + 1, iy) + _
                       bArray(ix, iy + 1) + bArray(ix, iy - 1)) / 3.8
         If CulAv > 0 Then
            If CulAv >= 255 Then CulAv = 32 + Rnd * 200
            bDummy(ix, iy) = CulAv
         Else
            bDummy(ix, iy) = bArray(ix, iy)
         End If
      Next ix
      Next iy
      bArray() = bDummy()
   End If
End Sub

Public Sub PepperEm(CulNum As Long)
Dim ixb As Long, iyb As Long
   If ASELECTION Then
      Getpx1py1 SSX, 0, SSY, 0
      iyb = py1 - SSH + 1
      For iy = 1 To SSH
      ixb = px1
      For ix = 1 To SSW
         If bMask(ix, iy) <> 0 Then
            If (Rnd - 0.01) < 0 Then
               If ixb > 0 Then
               If ixb <= canvasW Then
               If iyb > 0 Then
               If iyb <= canvasH Then
                  bArray(ixb, iyb) = CulNum
               End If
               End If
               End If
               End If
            End If
         End If
         ixb = ixb + 1
      Next ix
      iyb = iyb + 1
      Next iy
   Else
      For iy = 1 To canvasH
      For ix = 1 To canvasW
         If (Rnd - 0.01) < 0 Then
            SetDot ix, iy, CulNum, 0
         End If
      Next ix
      Next iy
   End If
End Sub

Public Sub ThickenPixels()
Dim ixb As Long, iyb As Long
Dim bCul As Byte
   If ASELECTION Then
      ReDim bDummy(SSW, SSH) As Byte
      Getpx1py1 SSX, 0, SSY, 0
      ' (1,1)
      iyb = py1 - SSH + 1
      For iy = 2 To SSH - 1 Step 2
         ixb = px1 + 1
         For ix = 2 To SSW - 1 Step 2
            THICK ixb, iyb ' bCul = bArray(ixb, iyb) if <> 0 put around bDummy(ix,iy)
            ixb = ixb + 2
         Next ix
         iyb = iyb + 2
      Next iy
      ' (2,1)
      ixb = px1 + 2
      For ix = 3 To SSW - 1 Step 2
      iyb = py1 - SSH + 1
      For iy = 2 To SSH - 1 Step 2
            THICK ixb, iyb ' bCul = bArray(ixb, iyb) if <> 0 put around bDummy(ix,iy)
            iyb = iyb + 2
         Next iy
         ixb = ixb + 2
      Next ix
      ' (2,2)
      iyb = py1 - SSH + 2
      For iy = 3 To SSH - 1 Step 2
         ixb = px1 + 2
         For ix = 3 To SSW - 1 Step 2
            THICK ixb, iyb ' bCul = bArray(ixb, iyb) if <> 0 put around bDummy(ix,iy)
            ixb = ixb + 2
         Next ix
         iyb = iyb + 2
      Next iy
      ' (1,2)
      ixb = px1 + 1
      For ix = 2 To SSW - 1 Step 2
      iyb = py1 - SSH + 2
      For iy = 3 To SSH - 1 Step 2
            THICK ixb, iyb ' bCul = bArray(ixb, iyb) if <> 0 put around bDummy(ix,iy)
            iyb = iyb + 2
         Next iy
         ixb = ixb + 2
      Next ix
      ReDim bPic(SSW, SSH)
      bPic() = bDummy()
      InsertbPic
      Erase bDummy()
   
   Else  ' Thicken whole picture
      ReDim bDummy(canvasW, canvasH)
      ' (1,1)
      For iy = 1 To canvasH Step 2
      For ix = 1 To canvasW Step 2
         ThickenWholeImage bDummy(), bArray()
      Next ix
      Next iy
      ' (2,1)
      For ix = 2 To canvasW Step 2
      For iy = 1 To canvasH Step 2
         ThickenWholeImage bDummy(), bArray()
      Next iy
      Next ix
      ' (2,2)
      For iy = 2 To canvasH Step 2
      For ix = 2 To canvasW Step 2
         ThickenWholeImage bDummy(), bArray()
      Next ix
      Next iy
      ' (1,2)
      For ix = 1 To canvasW Step 2
      For iy = 2 To canvasH Step 2
         ThickenWholeImage bDummy(), bArray()
      Next iy
      Next ix
      bArray() = bDummy()
      Erase bDummy()
   End If
End Sub

Public Sub THICK(ixb As Long, iyb As Long)
' bDummy() <- bArray()
' Public ix as Long, iy As Long
Dim bCul As Byte
   If bMask(ix, iy) <> 0 Then
      If iyb > 0 Then
      If iyb <= canvasH Then
      If ixb > 0 Then
      If ixb <= canvasW Then
         
         bCul = bArray(ixb, iyb)
         If bCul <> 0 Then
            bDummy(ix, iy) = bCul
            bDummy(ix - 1, iy) = bCul
            bDummy(ix + 1, iy) = bCul
            bDummy(ix, iy - 1) = bCul
            bDummy(ix, iy + 1) = bCul
         End If
      
      End If
      End If
      End If
      End If
   End If
End Sub

Public Sub ThickenWholeImage(bArr1() As Byte, bArr2() As Byte)
' bDummy() <- bArray()
' Public ix as Long, iy As Long
   If bArr2(ix, iy) > 0 Then
      bArr1(ix, iy) = bArr2(ix, iy)
      If ix > 1 Then bArr1(ix - 1, iy) = bArr2(ix, iy)
      If ix < canvasW Then bArr1(ix + 1, iy) = bArr2(ix, iy)
      If iy < canvasH Then bArr1(ix, iy + 1) = bArr2(ix, iy)
      If iy > 1 Then bArr1(ix, iy - 1) = bArr2(ix, iy)
   End If
End Sub

Public Sub ColorReplacer()
' CulNum done
Dim ixb As Long, iyb As Long
   If ASELECTION Then
      Getpx1py1 SSX, 0, SSY, 0
      For iy = 1 To SSH
      For ix = 1 To SSW
         If bMask(ix, iy) <> 0 Then
            ixb = px1 + ix - 1
            iyb = py1 - iy + 2
            If ixb > 0 Then
            If ixb <= canvasW Then
            If iyb > 0 Then
            If iyb <= canvasH Then
               If bArray(ixb, iyb) = SelLeftCulNum Then bArray(ixb, iyb) = SelRightCulNum
            End If
            End If
            End If
            End If
         End If
      Next ix
      Next iy
   Else
      For iy = 1 To canvasH
      For ix = 1 To canvasW
         If bArray(ix, iy) = SelLeftCulNum Then bArray(ix, iy) = SelRightCulNum
      Next ix
      Next iy
   End If
End Sub
