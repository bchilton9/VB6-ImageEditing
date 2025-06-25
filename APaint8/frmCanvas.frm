VERSION 5.00
Begin VB.Form frmCanvas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Resizing"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3435
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAspect 
      Caption         =   "Keep aspect ratio"
      Height          =   210
      Left            =   630
      TabIndex        =   14
      Top             =   1890
      Width           =   1920
   End
   Begin VB.OptionButton optCanOrImage 
      Caption         =   "Extract rectangular selection"
      Height          =   210
      Index           =   2
      Left            =   615
      TabIndex        =   13
      Top             =   1545
      Width           =   2535
   End
   Begin VB.OptionButton optCanOrImage 
      Caption         =   "Resize canvas && image"
      Height          =   210
      Index           =   1
      Left            =   615
      TabIndex        =   12
      Top             =   1230
      Width           =   2160
   End
   Begin VB.OptionButton optCanOrImage 
      Caption         =   "Resize canvas only"
      Height          =   210
      Index           =   0
      Left            =   615
      TabIndex        =   11
      Top             =   930
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton cmdSET4 
      Caption         =   "Press to allow Accept (Mod 4 check)"
      Height          =   285
      Left            =   105
      TabIndex        =   10
      Top             =   2235
      Width           =   3195
   End
   Begin VB.VScrollBar VSW 
      Height          =   390
      Index           =   1
      Left            =   1245
      TabIndex        =   7
      Top             =   330
      Width           =   210
   End
   Begin VB.VScrollBar VSW 
      Height          =   390
      Index           =   0
      Left            =   975
      TabIndex        =   6
      Top             =   330
      Width           =   210
   End
   Begin VB.TextBox txtW 
      Height          =   285
      Left            =   345
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "W"
      Top             =   375
      Width           =   585
   End
   Begin VB.TextBox txtH 
      Height          =   285
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "H"
      Top             =   375
      Width           =   585
   End
   Begin VB.VScrollBar VSH 
      Height          =   375
      Index           =   1
      Left            =   2820
      TabIndex        =   3
      Top             =   345
      Width           =   195
   End
   Begin VB.VScrollBar VSH 
      Height          =   375
      Index           =   0
      Left            =   2565
      TabIndex        =   2
      Top             =   345
      Width           =   195
   End
   Begin VB.CommandButton cmdWHAC 
      Caption         =   "Cancel"
      Height          =   270
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   2730
      Width           =   780
   End
   Begin VB.CommandButton cmdWHAC 
      Caption         =   "Accept"
      Height          =   270
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   2715
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   210
      Index           =   1
      Left            =   1950
      TabIndex        =   9
      Top             =   75
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Width"
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   8
      Top             =   105
      Width           =   795
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmCanvas.frm
Option Explicit
Option Base 1

Dim a$
Dim actVWH As Boolean
Dim aTC As Boolean
Dim aCanImage As Boolean
Dim iix As Long, iiy As Long
Dim zASP As Single
Dim aASP As Boolean
Dim aASPApplied As Boolean
Dim zVW As Single
Dim zVH As Single

Private Sub Form_Load()
   frmCanvas.Top = frmCanvasTop
   frmCanvas.Left = frmCanvasLeft
   
   Label1(0) = "Width [4-" & Trim$(MAXWIDTH) & "]"
   Label1(1) = "Height [4-" & Trim$(MAXHEIGHT) & "]"
   
   zASP = canvasW / canvasH
   chkAspect.Enabled = False
   aASP = False
   aASPApplied = False
   
   '---------------------------------------
   
   actVWH = False
   VSW(0).Min = MAXWIDTH   '4
   VSW(0).Max = 4          'MAXWIDTH
   VSW(0).SmallChange = 4
   VSW(0).LargeChange = 4
   VSW(1).Min = MAXWIDTH   '4
   VSW(1).Max = 4 'MAXWIDTH
   VSW(1).SmallChange = 12
   VSW(1).LargeChange = 12

   VSH(0).Min = MAXHEIGHT  '4
   VSH(0).Max = 4          'MAXHEIGHT
   VSH(0).SmallChange = 4
   VSH(0).LargeChange = 4
   VSH(1).Min = MAXHEIGHT  '4
   VSH(1).Max = 4          'MAXHEIGHT
   VSH(1).SmallChange = 12
   VSH(1).LargeChange = 12
   actVWH = True
   '---------------------------------------

   ' SetWHScrollBars
   WTemp = canvasW
   HTemp = canvasH
   txtW.Text = Str$(canvasW)
   txtW.Refresh
   txtH.Text = Str$(canvasH)
   txtH.Refresh
   actVWH = False
   VSW(0).Value = WTemp
   VSW(1).Value = WTemp
   VSH(0).Value = HTemp
   VSH(1).Value = HTemp
   actVWH = True
   
   cmdWHAC(0).Enabled = False
   aTC = True
   aCanImage = optCanOrImage(0).Value
   optCanOrImage(2).Enabled = aSelRect
End Sub

Private Sub chkAspect_Click()
   aASP = -chkAspect.Value
   cmdWHAC(0).Enabled = False
End Sub

Private Sub optCanOrImage_Click(Index As Integer)
   aCanImage = optCanOrImage(0).Value
   Select Case Index
   Case 0   ' Resize canvas only
      chkAspect.Value = 0
      chkAspect.Enabled = False
      aASP = False
      cmdWHAC(0).Enabled = False
      txtH.Enabled = True
      txtW.Enabled = True
      cmdSET4.Enabled = True
   Case 1   ' Resize canvas & image
      chkAspect.Enabled = True
      cmdWHAC(0).Enabled = False
      txtH.Enabled = True
      txtW.Enabled = True
      cmdSET4.Enabled = True
   Case 2   ' Extract rectangle
      chkAspect.Value = 0
      chkAspect.Enabled = False
      aASP = False
      cmdWHAC(0).Enabled = True
      txtH.Enabled = False
      txtW.Enabled = False
      cmdSET4.Enabled = False
   End Select
End Sub

'#### W & H Changes from Canvas frame ################################
' W & H Scroll bars
Private Sub VSW_Change(Index As Integer)
   If Not actVWH Then Exit Sub
   If optCanOrImage(2).Value Then Exit Sub
   
   txtW.Text = Str$(WTemp)
   txtW.Refresh
   If Not actVWH Then Exit Sub
   Select Case Index
   Case 0   ' +/-  4 W
      txtW.Text = Str$(VSW(0).Value)
      VSW(1).Value = VSW(0).Value
   Case 1   ' +/- 12 W
      txtW.Text = Str$(VSW(1).Value)
      VSW(0).Value = VSW(1).Value
   End Select
   
   If aASP Then
      zVW = VSW(0).Value
      If zVW / zASP >= 4 And zVW / zASP <= MAXHEIGHT Then
         actVWH = False
         aASPApplied = True
         VSH(0).Value = VSW(0).Value / zASP
         VSH(1).Value = VSW(1).Value / zASP
         txtH.Text = Str$(VSH(0).Value)
         actVWH = True
      End If
   End If
   WTemp = VSW(0).Value
End Sub

Private Sub VSH_Change(Index As Integer)
   If Not actVWH Then Exit Sub
   If optCanOrImage(2).Value Then Exit Sub
   
   txtH.Text = Str$(HTemp)
   txtH.Refresh
   If Not actVWH Then Exit Sub
   Select Case Index
   Case 0   ' +/-  4 H
      txtH.Text = Str$(VSH(0).Value)
      VSH(1).Value = VSH(0).Value
   Case 1   ' +/- 12 H
      txtH.Text = Str$(VSH(1).Value)
      VSH(0).Value = VSH(1).Value
   End Select
   
   If aASP Then
      zVH = VSH(0).Value
      If zVH * zASP >= 4 And zVH * zASP <= MAXHEIGHT Then
         actVWH = False
         aASPApplied = True
         VSW(0).Value = VSH(0).Value * zASP
         VSW(1).Value = VSH(1).Value * zASP
         txtW.Text = Str$(VSW(0).Value)
         actVWH = True
      End If
   End If
      
   HTemp = VSH(0).Value
End Sub

Private Sub txtW_Change()
   If Not aTC Then Exit Sub
   cmdWHAC(0).Enabled = False
   a$ = Trim$(txtW.Text)
   If a$ <> "" Then
      WTemp = (Val(a$) + 3) And &HFFFFFFFC
      If WTemp > MAXWIDTH Then WTemp = MAXWIDTH
      actVWH = False
      VSW(0).Value = WTemp
      VSW(1).Value = WTemp
      actVWH = True
   End If
End Sub

Private Sub txtH_Change()
   If Not aTC Then Exit Sub
   cmdWHAC(0).Enabled = False
   a$ = Trim$(txtH.Text)
   If a$ <> "" Then
      HTemp = (Val(a$) + 3) And &HFFFFFFFC
      If HTemp > MAXHEIGHT Then HTemp = MAXHEIGHT
      actVWH = False
      VSH(0).Value = HTemp
      VSH(1).Value = HTemp
      actVWH = True
   End If
End Sub

Private Sub txtW_KeyPress(KeyAscii As Integer)
   Const Numbers$ = "0123456789"
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdSET4_Click
      cmdWHAC(0).Enabled = False
   ElseIf Len(txtW.Text) > 5 Then
      KeyAscii = 0
   ElseIf KeyAscii <> 8 Then   ' Backspace
      If InStr(Numbers, Chr(KeyAscii)) = 0 Then KeyAscii = 0
   End If
End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)
   Const Numbers$ = "0123456789"
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdSET4_Click
      cmdWHAC(0).Enabled = False
   ElseIf Len(txtW.Text) > 5 Then
      KeyAscii = 0
   ElseIf KeyAscii <> 8 Then   ' Backspace
      If InStr(Numbers, Chr(KeyAscii)) = 0 Then KeyAscii = 0
   End If
End Sub

Private Sub cmdSET4_Click()
   WTemp = (WTemp + 3) And &HFFFFFFFC
   If WTemp > MAXWIDTH Then WTemp = MAXWIDTH
   HTemp = (HTemp + 3) And &HFFFFFFFC
   If HTemp > MAXHEIGHT Then HTemp = MAXHEIGHT
   
   ' If Keep aspec but not so far applied then
   ' keep width & adjust height
   If aASP And Not aASPApplied Then
      WTemp = (WTemp + 3) And &HFFFFFFFC
      If WTemp > MAXWIDTH Then WTemp = MAXWIDTH
      HTemp = WTemp / zASP
      HTemp = (HTemp + 3) And &HFFFFFFFC
      If HTemp > MAXHEIGHT Then HTemp = MAXHEIGHT
   End If
   
   aTC = False
   txtW.Text = Str$(WTemp)
   txtW.SelStart = Len(txtW.Text)
   txtH.Text = Str$(HTemp)
   txtH.SelStart = Len(txtH.Text)
   aTC = True
   cmdWHAC(0).Enabled = True
End Sub

Private Sub cmdWHAC_Click(Index As Integer)
' Accept, Extract SelRect or Cancel changed canvasW/H
' WTemp & HTemp  textbox values
Dim iyd As Long
Dim zW As Single
Dim zH As Single
Dim zix As Single, ziy As Single
Dim i As Long, j As Long

   Select Case Index
   Case 0   ' Accept new W & H
      If optCanOrImage(0).Value Then   ' Resize canvas only
         canvasW = WTemp
         canvasH = HTemp
         ReDim bDummy(canvasW, canvasH) As Byte
         iyd = canvasH + 1
         For iy = UBound(bArray(), 2) To 1 Step -1
            iyd = iyd - 1
            If iyd > 0 Then
            If iyd <= canvasH Then
               For ix = 1 To UBound(bArray(), 1)
                  If ix <= canvasW Then
                     bDummy(ix, iyd) = bArray(ix, iy)
                  End If
               Next ix
            End If
            End If
         Next iy
      
      ElseIf optCanOrImage(1).Value Then  ' Resize canvas & image
         
         zW = canvasW / WTemp
         zH = canvasH / HTemp
         ReDim bDummy(WTemp, HTemp) As Byte
         For iy = 1 To HTemp
            j = CLng(zH * iy)
            If j < 1 Then j = 1
               For ix = 1 To WTemp
                  i = CLng(zW * ix)
                  If i < 1 Then i = 1
                  bDummy(ix, iy) = bArray(i, j)
               Next ix
         Next iy
         canvasW = WTemp
         canvasH = HTemp
      
      ElseIf optCanOrImage(2).Value Then  ' Extract rectangular selection
         SSW = (SSW + 3) And &HFFFFFFFC
         SSH = (SSH + 3) And &HFFFFFFFC
         If SSY < 1 Then SSY = 1
         If SSX < 1 Then SSX = 1
         If SSY + SSH - 1 > canvasH Then SSH = canvasH - SSY + 1
         If SSX + SSW - 1 > canvasW Then SSW = canvasW - SSX + 1
         ReDim bDummy(SSW, SSH) As Byte
         For iy = SSY To SSY + SSH - 1
         For ix = SSX To SSX + SSW - 1
            iix = ix - SSX + 1
            iiy = iy - SSY + 1
            iiy = SSH - iiy
            If iiy > 0 Then
            If iiy <= SSY + SSH - 1 Then
            If iix > 0 Then
            If iix <= SSX + SSW - 1 Then
               bDummy(iix, iiy) = bArray(ix, canvasH - iy)
            End If
            End If
            End If
            End If
         Next ix
         Next iy
         canvasW = SSW
         canvasH = SSH
      End If
      
      ReDim bArray(canvasW, canvasH)
      bArray() = bDummy()
      Erase bDummy()
      aCanWH = True
       ' Action on return
'      PIC.Width = canvasW
'      PIC.Height = canvasH
'      If NewNum > 1 Then    ' ie not New
'        SAVE_CurrentImage
'         FixUndos           ' unless StopUndos = True
'      End If
'      FillLabInfos
'      DISPLAY
   Case 1   ' Cancel new H
      aCanWH = False
   End Select

   frmCanvasTop = frmCanvas.Top
   frmCanvasLeft = frmCanvas.Left
   Unload frmCanvas
End Sub

