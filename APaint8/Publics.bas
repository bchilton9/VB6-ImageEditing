Attribute VB_Name = "Publics"
' Publics.bas

Option Explicit
Option Base 1

' Files
Public PathSpec$, FileSpec$
Public AppPathSpec$, OpenPathSpec$, SavePathSpec$
' frmBrowse position
Public frmBrowseTop As Long, frmBrowseLeft As Long

' Sizes
Public NewNum As Long
Public canvasW As Long, canvasH As Long
Public svcanvasW As Long, svcanvasH As Long
Public MAXWIDTH As Long    ' Set in Form_Initialize
Public MAXHEIGHT As Long   ' Set in Form_Initialize
Public WTemp As Long
Public HTemp As Long
Public bDummy() As Byte
Public frmCanvasLeft As Long
Public frmCanvasTop As Long
Public aCanWH As Boolean

'Roller/Shifter
Public RollShift As Long

' Picture byte arrays
Public bArray() As Byte, StartWidth As Long, StartHeight As Long
' PICContainer Width & Height
Public PICCW As Long
Public PICCH As Long

' Color
Public CulRGB() As Long
Public CulBGR() As Long
Public palRed() As Byte, palGreen() As Byte, palBlue() As Byte
Public bred As Byte, bgreen As Byte, bblue As Byte
Public SelLeftCulNum As Long, SelRightCulNum As Long
Public Selectcn As Long
Public MPal() As Byte
Public StorePal() As Byte
Public BackUpRGB() As Long
Public DefaultRGB() As Long
' frmPalette position
Public frmPaletteLeft As Long, frmPaletteTop As Long

' Views
Public frmViewsLeft As Long, frmViewsTop As Long
Public aVIEWS As Boolean
Public aMNUACTION As Boolean

'' Temp dimensions
Public RpicW As Long
Public RpicH As Long

' Zoom
Public ZoomFactor As Long
Public aZoom As Boolean
Public ZoomSize As Long    ' Set in  Set in Form_Initialize
Public ZoomArr() As Long, ZoomWidth As Long, ZoomHeight As Long
Public ZARR() As Long
' frmZoom position
Public frmZoomTop As Long, frmZoomLeft As Long

' frmHelp
Public frmHelpTop As Long, frmHelpLeft As Long

' Hairs
Public aHairs As Boolean

'Undos/Redos
Public UndoNum As Long
Public TopUndoNum As Long
Public StopUndos As Boolean
Public aOverWrite As Boolean  ' True - Overwrite, False - To background
''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Undo/Redo arrays in UnRedo.bas
''''''''''''''''''''''''''''''''''''''''''''''''''

' For key cursor movement
Public xprev As Single
Public yprev As Single

' Text
Public frmTextTop As Long, frmTextLeft As Long
Public TextLine$
Public TextColor As Long
Public Type FontStuff
   FontName As String
   FontSize As Long
   FontItalic As Boolean
   FontBold As Boolean
End Type
Public SVFont As FontStuff

' To confirm Mouse Down
Public AMouseDown As Boolean
Public ADRAW As Boolean  ' To flag drawing in progress

' Tool options
Public frmToolOptionsTop As Long, frmToolOptionsLeft As Long
' Tools
Public LCNum As Long    ' Left button counter
Public RCNum As Long    ' Right button counter
Public CulNum As Long

' For storing points
Public NSTOREXY As Long
Public STOREX() As Long
Public STOREY() As Long
Public DCul As Long ' XOR Draw color

Public IncrX As Long
Public IncrY As Long

Public StartButton As Integer
Public CopyStartButton As Integer
Public ToolType As Long
Public TempToolType As Long
Public px1 As Long, py1 As Long
Public px2 As Long, py2 As Long

' On Form1
Public Enum EToolType
   Brush = 0
   Spray
   ALine
   PolyLine
   CurvyLine
   Rectangle
   Cirllipse
   Cone
   Tube
   Bullet
   Junction
   Arc
   Shape
   Radial
   AFill
   Tree
   Arrow
   AText
   SelR
   SelC
   SelE
   SelL
   Desel
   SCopyPaste
   SCopy
   SCut
   SReflectLR
   SReflectUD
   SRotate
   SPaste
   SClear
   Rot90
   Mix
   Thicken
   Pepper
   LRColor
   Measure
   Pick
   Smooth1
   Smooth2
   Smooth4
End Enum

Public BrushType As Long
Public RibIncrX As Long
Public RibIncrY As Long
' Signed RibIncrX,RibIncrY
Public RY1 As Long, RX1 As Long
Public RY2 As Long, RX2 As Long
Public Enum EBrushType
   Dot1 = 0
   Dot2
   Dot3
   FreeDraw1
   FreeDraw2
   FreeDraw3
   BRibbon1
   BRibbon2
   BRibbon3
   FRibbon1
   FRibbon2
   FRibbon3
End Enum

Public SprayType As Long
Public zradmax As Single
Public sprayn As Long
Public Enum ESprayType
   Dots1 = 0
   Dots2
   Dots3
   Plusses1
   Plusses2
   Plusses3
   Crosses1
   Crosses2
   Crosses3
   Diamonds1
   Diamonds2
   Diamonds3
End Enum

Public LineType As Long
Public zspace As Single   ' DoubleLine spacins
Public Enum ELineType
   SingleLine1 = 0
   SingleLine2
   SingleLine3
   DottedLine1
   DottedLine2
   DottedLine3
   DoubleLine1
   DoubleLine2
   DoubleLine3
   DoubleLineEnd1
   DoubleLineEnd2
   DoubleLineEnd3
   DoubleDottedLine1
   DoubleDottedLine2
   DoubleDottedLine3
   ShadedLine1
   ShadedLine2
   ShadedLine3
End Enum

Public PolyLineType As Long
Public AMoveAll As Boolean
Public svPolyLineType As Long ' For curvy line to use as well
Public Enum EPolyLineType
   PolySingleLine1 = 0
   PolySingleLine2
   PolySingleLine3
   PolyDoubleLine1
   PolyDoubleLine2
   PolyDoubleLine3
   PolyDoubleLineEnd1
   PolyDoubleLineEnd2
   PolyDoubleLineEnd3
   PolyShadedLine1
   PolyShadedLine2
   PolyShadedLine3
End Enum

Public CurvyLineType As Long
Public Enum ECurvyLineType
   CurvySingleLine1 = 0
   CurvySingleLine2
   CurvySingleLine3
   CurvyDoubleLine1
   CurvyDoubleLine2
   CurvyDoubleLine3
   CurvyDoubleLineEnd1
   CurvyDoubleLineEnd2
   CurvyDoubleLineEnd3
   CurvyShadedLine1
   CurvyShadedLine2
   CurvyShadedLine3
End Enum

Public RectangleType As Long
Public svRectangleType As Long
Public Enum ERectangleType
   RectangleSingle1 = 0
   RectangleSingle2
   RectangleSingle3
   RectangleDotted1
   RectangleDotted2
   RectangleDotted3
   RectangleDouble1
   RectangleDouble2
   RectangleDouble3
   RectangleShaded1
   RectangleShaded2
   RectangleShaded3
   RectangleFShade   '/
   RectangleBShade   '\
   RectangleFilled   'with CulNum
End Enum

Public CirllipseType As Long
Public ixc As Long, iyc As Long
Public zrad As Single, zratio As Single
Public zradx As Single, zrady As Single
Public zrad2 As Single
Public Enum ECirllipseType
   CirllipseSingle1 = 0
   CirllipseSingle2
   CirllipseSingle3
   CirllipseDotted1
   CirllipseDotted2
   CirllipseDotted3
   CirllipseDouble1
   CirllipseDouble2
   CirllipseDouble3
   CirllipseShaded1
   CirllipseShaded2
   CirllipseShaded3
End Enum

Public ConeType As Long
Public Enum EConeType
   ConeOutline = 0
   ConeHShade1
   ConeHSHade2
   ConeCShade1
   ConeCShade2
   ConeCross
End Enum

Public TubeType As Long
Public NTube As Long
Public Enum ETubeType
   TubeOutLine = 0
   TubeHShade
   TubeCShade
End Enum

Public BulletType As Long
Public Enum EBulletType
   BulletOutLine = 0
   BulletHShade
   BulletCShade
End Enum

Public JunctionType As Long
Public zTL As Single ' Default length of side piece
Public XT() As Single, YT() As Single
Public Enum EJunctionType
   TPiece1 = 0
   TPiece2
   TPiece3
   Cross1
   Cross2
   Cross3
   Corner1
   Corner2
   Corner3
End Enum
   
Public ArcType As Long
Public zSA As Single, zEA As Single  ' Start/End angles
Public Enum EArcType
   ArcFull = 0
   ArcTL
   ArcBL
   ArcTR
   ArcBR
   ArcLS
   ArcRS
   ArcTS
   ArcBS
   ArcBRX
   ArcTRX
   ArcBLX
   ArcTLX
End Enum

Public ShapeType As Long
Public Enum EShapeType
   TShape1 = 0
   TShape2
   TShape3
   TShape4
   TShape5
   TShape6
End Enum

Public RadialType As Long
Public RadialRep() As Long
Public zangle As Single
Public Enum ERadialType
   RSpokes = 0
   RStars
   RRadCircs
   RPolygons
   RTeeth
End Enum

Public FillType As Long
Public bPattern() As Byte
Public Enum EFillType
   Fill1 = 0
   Fill2
   Fill3
   Fill4
   Fill5
   Fill6
   Fill7
   Fill8
   Fill9
   Fill10
   Fill11
   Fill12
   Fill13
   Fill14
   Fill15
   Fill16
   Fill17
   Fill18
   Fill19
   Fill20
   Fill21
   Fill22
End Enum
   
Public TreeType As Long
Public BushSize() As Long
Public YTreeMin As Single   ' For color gradient
Public YTreeMax As Single   ' For color gradient
Public zAngP() As Single, zAngN() As Single
Public xstep() As Single, ystep() As Single
Public xmul() As Single, ymul() As Single
Public Axiom$(), PAxiom$()
'Public bTreeArray() As Byte
Public Enum ETreeType
   Tree1 = 0
   Tree2
   Tree3
End Enum

Public ArrowType As Long
Public zarrang As Single
Public zarrlen As Single
Public xd1 As Single, yd1 As Single
Public xd2 As Single, yd2 As Single
Public Enum EArrowType
   ArrSingle = 0
   ArrFeathered
   ArrTriangle
End Enum

Public SelectType As Long
Public ASELECTION As Boolean
Public aSelRect As Boolean
Public aSelCirc As Boolean
Public aSelEllip As Boolean
Public aSelLasso As Boolean
Public SSX As Long ' Shape left
Public SSY As Long ' Shape top
Public SSW As Long ' Shape width
Public SSH As Long ' Shape Height
Public shw As Long, shh As Long
Public shx As Long, shy As Long

Public bMask() As Byte  ' Selected rect
Public bPic() As Byte
Public zangRot As Single
Public zangRotCSEL As Single

Public Enum ESelectType
   SelRect = 0
   SelCirc
   SelEllip
   SelLasso
   Deselect
End Enum
' Lasso
Public NumLassoLines As Long  ' SL(0)-SL(NumLassoLines-1)
Public XS1 As Single, YS1 As Single
Public XS2 As Single, YS2 As Single
Public XSMax As Single, YSMax As Single
Public XSMin As Single, YSMin As Single

' Measure
Public aMeasure As Boolean

' Strip
Public MaxNumFrames As Long
Public NumStrips As Long
Public zTotalAng As Single
Public zIncrAng As Single
Public zFinalPercentReduc As Single
Public zIncrPercentReduc As Single
Public frmStripLeft
Public frmStripTop
Public aPepper As Boolean

' Transforms
' frmTransform position
Public frmTransformTop As Long, frmTransformLeft As Long
Public aLensCheck As Boolean
Public aUseSelectcn As Boolean
Public TransformType As Long
Public Enum ETransformType
   TNone = 0
   TContour
   TDither
   TEngraveEmboss
   TPosterize
   TRelief
   TSmooth
   TShadeV
   TShadeH
   TMelt
   TOil
   TSharpen
   TLitho
   TContrast
   TDiffuse
   THDiffuse
   TVDiffuse
   TBlackWhite
   TSolar
   TInvert
   TFog
   TSquare
   
   TEllipse
   TFluteH
   TFluteV
   TRippleH
   TRippleV
   TRoundRect
   TTile
   TMirrorL
   TMirrorR
   TMirrorT
   TMirrorB
   TMlens
   TLens
   TFWindowVert
   TSwirl
   TSpokess
   TMinMag
   TBubbly
   TRotate
   TTunnel
   TFWindowHorz
   TFWindowHV
   
   THLines
   TVLines
   THVLines
   THWaves
   TVWaves
   THVWaves
   TCircles
   TEllipses
   TThickLineH
   TBorder
   TSpokes
   TDNet
   TThickLineV
   TThickLineHV
End Enum

' ASM
Public GetIndexMC() As Byte     ' Array to hold machine code
Public ptMC As Long             ' Ptr to Machine Code
Public LongDerived As Long       ' Input RGBA from Derived pal
Public ptStanPal As Long         '= VarPtr(CulRGB(0)) Standard pal

' Screen.TwipsPerPixelX: Screen.TwipsPerPixelY
Public STX As Long, STY As Long

' General
Public ix As Long
Public iy As Long
Public aDone As Boolean
Public Const pi# = 3.1415927
Public Const d2r# = pi# / 180
