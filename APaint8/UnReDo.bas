Attribute VB_Name = "UnReDo"
' UnReDo.bas

Option Explicit
Option Base 1

Public Const MaxUndos As Long = 15

' To allow various canvas sizes
' Use: Select Case UndoNum
' canvasW,H = UBound(bUndo#(),1),UBound(bUndo#(),2)
Public bUndo0() As Byte
Public bUndo1() As Byte
Public bUndo2() As Byte
Public bUndo3() As Byte
Public bUndo4() As Byte
Public bUndo5() As Byte
Public bUndo6() As Byte
Public bUndo7() As Byte
Public bUndo8() As Byte
Public bUndo9() As Byte
Public bUndo10() As Byte
Public bUndo11() As Byte
Public bUndo12() As Byte
Public bUndo13() As Byte
Public bUndo14() As Byte
Public bUndo15() As Byte

Public CopyRGB() As Long
Public CopyFileSpec$()

Public bTemp() As Byte

Dim a$

Public Sub SizeUndoPalettes()
   ReDim CopyRGB(0 To 255, 0 To 15)
   ReDim CopyFileSpec$(0 To 15)
End Sub

Public Sub ReDimUndos()
   ReDim bUndo0(1, 1)
   ReDim bUndo1(1, 1)
   ReDim bUndo2(1, 1)
   ReDim bUndo3(1, 1)
   ReDim bUndo4(1, 1)
   ReDim bUndo5(1, 1)
   ReDim bUndo6(1, 1)
   ReDim bUndo7(1, 1)
   ReDim bUndo8(1, 1)
   ReDim bUndo9(1, 1)
   ReDim bUndo10(1, 1)
   ReDim bUndo11(1, 1)
   ReDim bUndo12(1, 1)
   ReDim bUndo13(1, 1)
   ReDim bUndo14(1, 1)
   ReDim bUndo15(1, 1)
End Sub

Public Sub ClearUndo()
' UndoNum
   CopyFileSpec$(UndoNum) = vbNullChar
   
   Select Case UndoNum
   Case 1: ReDim bUndo1(canvasW, canvasH)
   Case 2: ReDim bUndo2(canvasW, canvasH)
   Case 3: ReDim bUndo3(canvasW, canvasH)
   Case 4: ReDim bUndo4(canvasW, canvasH)
   Case 5: ReDim bUndo5(canvasW, canvasH)
   Case 6: ReDim bUndo6(canvasW, canvasH)
   Case 7: ReDim bUndo7(canvasW, canvasH)
   Case 8: ReDim bUndo8(canvasW, canvasH)
   Case 9: ReDim bUndo9(canvasW, canvasH)
   Case 10: ReDim bUndo10(canvasW, canvasH)
   Case 11: ReDim bUndo11(canvasW, canvasH)
   Case 12: ReDim bUndo12(canvasW, canvasH)
   Case 13: ReDim bUndo13(canvasW, canvasH)
   Case 14: ReDim bUndo14(canvasW, canvasH)
   Case 15: ReDim bUndo15(canvasW, canvasH)
   End Select
End Sub

Public Sub RESTORE_Image()
   CopyMemory CulRGB(0), CopyRGB(0, UndoNum), 1024
   CopyToBGR
   FileSpec$ = CopyFileSpec$(UndoNum)
   Select Case UndoNum
   Case 1:  canvasW = UBound(bUndo1(), 1): canvasH = UBound(bUndo1(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo1()
   Case 2:  canvasW = UBound(bUndo2(), 1): canvasH = UBound(bUndo2(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo2()
   Case 3:  canvasW = UBound(bUndo3(), 1): canvasH = UBound(bUndo3(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo3()
   Case 4:  canvasW = UBound(bUndo4(), 1): canvasH = UBound(bUndo4(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo4()
   Case 5:  canvasW = UBound(bUndo5(), 1): canvasH = UBound(bUndo5(), 2)
            ReDim bArray(canvasW, canvasH)
            CopyMemory bArray(1, 1), bUndo5(1, 1), canvasW * canvasH
            bArray() = bUndo5()
   Case 6:  canvasW = UBound(bUndo6(), 1): canvasH = UBound(bUndo6(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo6()
   Case 7:  canvasW = UBound(bUndo7(), 1): canvasH = UBound(bUndo7(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo7()
   Case 8:  canvasW = UBound(bUndo8(), 1): canvasH = UBound(bUndo8(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo8()
   Case 9:  canvasW = UBound(bUndo9(), 1): canvasH = UBound(bUndo9(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo9()
   
   Case 10:  canvasW = UBound(bUndo10(), 1): canvasH = UBound(bUndo10(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo10()
   Case 11:  canvasW = UBound(bUndo11(), 1): canvasH = UBound(bUndo11(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo11()
   Case 12:  canvasW = UBound(bUndo12(), 1): canvasH = UBound(bUndo12(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo12()
   Case 13:  canvasW = UBound(bUndo13(), 1): canvasH = UBound(bUndo13(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo13()
   Case 14:  canvasW = UBound(bUndo14(), 1): canvasH = UBound(bUndo14(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo14()
   Case 15:  canvasW = UBound(bUndo15(), 1): canvasH = UBound(bUndo15(), 2)
            ReDim bArray(canvasW, canvasH)
            bArray() = bUndo15()
   End Select

   'picW = canvasW
   'picH = canvasH
End Sub

Public Sub SAVE_CurrentImage()
   If Not StopUndos Then
      If TopUndoNum = MaxUndos Then
         ' Cycle down leaving MaxUndos free
         CycleUndosDown 1  ' Do all
         TopUndoNum = MaxUndos - 1
      End If
      TopUndoNum = TopUndoNum + 1
      FillUndoNWithbArray TopUndoNum
      UndoNum = TopUndoNum
   Else
      FillUndoNWithbArray UndoNum
   End If
   'FillUndoNWithbArray UndoNum also does:-
   'CopyMemory CopyRGB(0, NUM), CulRGB(0), 1024
   'CopyFileSpec$(NUM) = FileSpec$

End Sub

Public Sub CycleUndosDown(NUM As Long)
' On deleting a view
Dim CWi As Long
Dim CHi As Long
Dim k As Long
   
   For k = NUM To 14
      CopyMemory CopyRGB(0, k), CopyRGB(0, k + 1), 1024
      CopyFileSpec$(k) = CopyFileSpec$(k + 1)
   Next k
   
   On NUM GoTo 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
1:
   ' 2->1, 3-2,,, 9->8
   CWi = UBound(bUndo2(), 1): CHi = UBound(bUndo2(), 2)
   ReDim bUndo1(1 To CWi, 1 To CHi)
   bUndo1() = bUndo2()
2:
   ' 3->2
   CWi = UBound(bUndo3(), 1): CHi = UBound(bUndo3(), 2)
   ReDim bUndo2(1 To CWi, 1 To CHi)
   bUndo2() = bUndo3()
3:
   ' 4->3
   CWi = UBound(bUndo4(), 1): CHi = UBound(bUndo4(), 2)
   ReDim bUndo3(1 To CWi, 1 To CHi)
   bUndo3() = bUndo4()
4:
   ' 5->4
   CWi = UBound(bUndo5(), 1): CHi = UBound(bUndo5(), 2)
   ReDim bUndo4(1 To CWi, 1 To CHi)
   bUndo4() = bUndo5()
5:
   ' 6->5
   CWi = UBound(bUndo6(), 1): CHi = UBound(bUndo6(), 2)
   ReDim bUndo5(1 To CWi, 1 To CHi)
   bUndo5() = bUndo6()
6:
   ' 7->6
   CWi = UBound(bUndo7(), 1): CHi = UBound(bUndo7(), 2)
   ReDim bUndo6(1 To CWi, 1 To CHi)
   bUndo6() = bUndo7()
7:
   ' 8->7
   CWi = UBound(bUndo8(), 1): CHi = UBound(bUndo8(), 2)
   ReDim bUndo7(1 To CWi, 1 To CHi)
   bUndo7() = bUndo8()
8:
   ' 9->8
   CWi = UBound(bUndo9(), 1): CHi = UBound(bUndo9(), 2)
   ReDim bUndo8(1 To CWi, 1 To CHi)
   bUndo8() = bUndo9()
9:
   ' 10->9
   CWi = UBound(bUndo10(), 1): CHi = UBound(bUndo10(), 2)
   ReDim bUndo9(1 To CWi, 1 To CHi)
   bUndo9() = bUndo10()
10:
   ' 11->10
   CWi = UBound(bUndo11(), 1): CHi = UBound(bUndo11(), 2)
   ReDim bUndo10(1 To CWi, 1 To CHi)
   bUndo10() = bUndo11()
11:
   ' 12->11
   CWi = UBound(bUndo12(), 1): CHi = UBound(bUndo12(), 2)
   ReDim bUndo11(1 To CWi, 1 To CHi)
   bUndo11() = bUndo12()
12:
   ' 13->12
   CWi = UBound(bUndo13(), 1): CHi = UBound(bUndo13(), 2)
   ReDim bUndo12(1 To CWi, 1 To CHi)
   bUndo12() = bUndo13()
13:
   ' 14->13
   CWi = UBound(bUndo14(), 1): CHi = UBound(bUndo14(), 2)
   ReDim bUndo13(1 To CWi, 1 To CHi)
   bUndo13() = bUndo14()
14:
   ' 15->14
   CWi = UBound(bUndo15(), 1): CHi = UBound(bUndo15(), 2)
   ReDim bUndo14(1 To CWi, 1 To CHi)
   bUndo14() = bUndo15()

End Sub

Public Sub Collapse()
   ' Collapse Undos to 1 & TopUndoNum
   
   If TopUndoNum <= 2 Then Exit Sub
   
   CopyMemory CopyRGB(0, 2), CopyRGB(0, TopUndoNum), 1024
   CopyFileSpec$(2) = CopyFileSpec$(TopUndoNum)
   
   Select Case TopUndoNum
   Case 3:  canvasW = UBound(bUndo3(), 1): canvasH = UBound(bUndo3(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo3()
   Case 4:  canvasW = UBound(bUndo4(), 1): canvasH = UBound(bUndo4(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo4()
   Case 5:  canvasW = UBound(bUndo5(), 1): canvasH = UBound(bUndo5(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo5()
   Case 6:  canvasW = UBound(bUndo6(), 1): canvasH = UBound(bUndo6(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo6()
   Case 7:  canvasW = UBound(bUndo7(), 1): canvasH = UBound(bUndo7(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo7()
   Case 8:  canvasW = UBound(bUndo8(), 1): canvasH = UBound(bUndo8(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo8()
   Case 9:  canvasW = UBound(bUndo9(), 1): canvasH = UBound(bUndo9(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo9()
   Case 10: canvasW = UBound(bUndo10(), 1): canvasH = UBound(bUndo10(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo10()
   Case 11: canvasW = UBound(bUndo11(), 1): canvasH = UBound(bUndo11(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo11()
   Case 12: canvasW = UBound(bUndo12(), 1): canvasH = UBound(bUndo12(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo12()
   Case 13: canvasW = UBound(bUndo13(), 1): canvasH = UBound(bUndo13(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo13()
   Case 14: canvasW = UBound(bUndo14(), 1): canvasH = UBound(bUndo14(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo14()
   Case 15: canvasW = UBound(bUndo15(), 1): canvasH = UBound(bUndo15(), 2)
            ReDim bUndo2(canvasW, canvasH)
            bUndo2() = bUndo15()
   End Select
   
   TopUndoNum = 2
   UndoNum = 2
End Sub

Public Sub DeleteCurrentView()
   ' Squeeze out current viewed backup
   If UndoNum > 1 And UndoNum < TopUndoNum Then
      CycleUndosDown UndoNum
      RESTORE_Image
      TopUndoNum = TopUndoNum - 1
   ElseIf UndoNum = TopUndoNum Then
      UndoNum = UndoNum - 1
      TopUndoNum = UndoNum
      RESTORE_Image
   End If
End Sub

Public Sub FillUndoNWithbArray(NUM As Long)
' NUM = TopUndoNum or UndoNum
' TopUndoNum from SAVE_CurrentImage
' UndoNum > 0  ' From Form1 Browser & Utils.bas SAVE_CurrentImage
   
   CopyMemory CopyRGB(0, NUM), CulRGB(0), 1024
   CopyFileSpec$(NUM) = FileSpec$
   
   Select Case NUM
   
   Case 1: ReDim bUndo1(1 To canvasW, 1 To canvasH)
           bUndo1() = bArray()
   
   Case 2: ReDim bUndo2(1 To canvasW, 1 To canvasH)
           bUndo2() = bArray()
   
   Case 3: ReDim bUndo3(1 To canvasW, 1 To canvasH)
           bUndo3() = bArray()
   
   Case 4: ReDim bUndo4(1 To canvasW, 1 To canvasH)
           bUndo4() = bArray()
   
   Case 5: ReDim bUndo5(1 To canvasW, 1 To canvasH)
           bUndo5() = bArray()
   
   Case 6: ReDim bUndo6(1 To canvasW, 1 To canvasH)
           bUndo6() = bArray()
   
   Case 7: ReDim bUndo7(1 To canvasW, 1 To canvasH)
           bUndo7() = bArray()
   
   Case 8: ReDim bUndo8(1 To canvasW, 1 To canvasH)
           bUndo8() = bArray()
   
   Case 9: ReDim bUndo9(1 To canvasW, 1 To canvasH)
           bUndo9() = bArray()
   
   Case 10: ReDim bUndo10(1 To canvasW, 1 To canvasH)
           bUndo10() = bArray()
   
   Case 11: ReDim bUndo11(1 To canvasW, 1 To canvasH)
           bUndo11() = bArray()
   
   Case 12: ReDim bUndo12(1 To canvasW, 1 To canvasH)
           bUndo12() = bArray()
   
   Case 13: ReDim bUndo13(1 To canvasW, 1 To canvasH)
           bUndo13() = bArray()
   
   Case 14: ReDim bUndo14(1 To canvasW, 1 To canvasH)
           bUndo14() = bArray()
   
   Case 15: ReDim bUndo15(1 To canvasW, 1 To canvasH)
           bUndo15() = bArray()
   End Select
End Sub

Public Sub AddView()
'Adds undoNum+1 to UndoNum
' Only adds to UndoNum where UndoNum ix,iy = 0
Dim bW As Long
Dim bH As Long
Dim dx As Long
Dim dy As Long
   CopyFileSpec$(UndoNum) = vbNullChar
   Select Case UndoNum
   Case 1: ' bArray() same size as bUndo1()
           ' bUndo1() + bUndo2() -> bArray() NB Sizes
           
       bW = UBound(bUndo2(), 1): bH = UBound(bUndo2(), 2)
      ' Copy  bUndo2(1-bW, 1-bH) to bUndo1(1-bW, 1-bH)
      ' Copy  bbUndo1(1-bW, 1-bH) to bArray(1-bW, 1-bH)
       For iy = 1 To bH
       dy = canvasH - iy + 1  ' Superimpose @ Top Left
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then ' Transfer all non-background <> 0
                  If bUndo2(ix, bH - iy + 1) <> 0 Then
                     bUndo1(ix, dy) = bUndo2(ix, bH - iy + 1)
                  End If
               Else  ' Only transfer where destination is background 0
                  If bUndo1(ix, dy) = 0 Then
                     bUndo1(ix, dy) = bUndo2(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo1()
   Case 2
       bW = UBound(bUndo3(), 1): bH = UBound(bUndo3(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo3(ix, bH - iy + 1) <> 0 Then
                     bUndo2(ix, dy) = bUndo3(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo2(ix, dy) = 0 Then
                     bUndo2(ix, dy) = bUndo3(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo2()
   Case 3
       bW = UBound(bUndo4(), 1): bH = UBound(bUndo4(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo4(ix, bH - iy + 1) <> 0 Then
                     bUndo3(ix, dy) = bUndo4(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo3(ix, dy) = 0 Then
                     bUndo3(ix, dy) = bUndo4(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo3()
   Case 4
       bW = UBound(bUndo5(), 1): bH = UBound(bUndo5(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo5(ix, bH - iy + 1) <> 0 Then
                     bUndo4(ix, dy) = bUndo5(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo4(ix, dy) = 0 Then
                     bUndo4(ix, dy) = bUndo5(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo4()
   Case 5
       bW = UBound(bUndo6(), 1): bH = UBound(bUndo6(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo6(ix, bH - iy + 1) <> 0 Then
                     bUndo5(ix, dy) = bUndo6(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo5(ix, dy) = 0 Then
                     bUndo5(ix, dy) = bUndo6(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo5()
   Case 6
       bW = UBound(bUndo7(), 1): bH = UBound(bUndo7(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo7(ix, bH - iy + 1) <> 0 Then
                     bUndo6(ix, dy) = bUndo7(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo6(ix, dy) = 0 Then
                     bUndo6(ix, dy) = bUndo7(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo6()
   Case 7
       bW = UBound(bUndo8(), 1): bH = UBound(bUndo8(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo8(ix, bH - iy + 1) <> 0 Then
                     bUndo7(ix, dy) = bUndo8(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo7(ix, dy) = 0 Then
                     bUndo7(ix, dy) = bUndo8(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo7()
   Case 8
       bW = UBound(bUndo9(), 1): bH = UBound(bUndo9(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo9(ix, bH - iy + 1) <> 0 Then
                     bUndo8(ix, dy) = bUndo9(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo8(ix, dy) = 0 Then
                     bUndo8(ix, dy) = bUndo9(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo8()
   Case 9
       bW = UBound(bUndo10(), 1): bH = UBound(bUndo10(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo10(ix, bH - iy + 1) <> 0 Then
                     bUndo9(ix, dy) = bUndo10(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo9(ix, dy) = 0 Then
                     bUndo9(ix, dy) = bUndo10(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo9()
   Case 10
       bW = UBound(bUndo11(), 1): bH = UBound(bUndo11(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo11(ix, bH - iy + 1) <> 0 Then
                     bUndo10(ix, dy) = bUndo11(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo10(ix, dy) = 0 Then
                     bUndo10(ix, dy) = bUndo11(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo10()
   Case 11
       bW = UBound(bUndo12(), 1): bH = UBound(bUndo12(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo12(ix, bH - iy + 1) <> 0 Then
                     bUndo11(ix, dy) = bUndo12(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo11(ix, dy) = 0 Then
                     bUndo11(ix, dy) = bUndo12(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo11()
   Case 12
       bW = UBound(bUndo13(), 1): bH = UBound(bUndo13(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo13(ix, bH - iy + 1) <> 0 Then
                     bUndo12(ix, dy) = bUndo13(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo12(ix, dy) = 0 Then
                     bUndo12(ix, dy) = bUndo13(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo12()
   Case 13
       bW = UBound(bUndo14(), 1): bH = UBound(bUndo14(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo14(ix, bH - iy + 1) <> 0 Then
                     bUndo13(ix, dy) = bUndo14(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo13(ix, dy) = 0 Then
                     bUndo13(ix, dy) = bUndo14(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo13()
   Case 14
       bW = UBound(bUndo15(), 1): bH = UBound(bUndo15(), 2)
       For iy = 1 To bH
       dy = canvasH - iy + 1
       If dy > 0 And dy <= canvasH Then
         For ix = 1 To bW
            If ix <= canvasW Then
               If aOverWrite Then
                  If bUndo15(ix, bH - iy + 1) <> 0 Then
                     bUndo14(ix, dy) = bUndo15(ix, bH - iy + 1)
                  End If
               Else
                  If bUndo14(ix, dy) = 0 Then
                     bUndo14(ix, dy) = bUndo15(ix, bH - iy + 1)
                  End If
               End If
            End If
         Next ix
       End If
       Next iy
       bArray() = bUndo14()
   Case 15 ' No
   End Select
End Sub

Public Sub SwapViews()
' Undonum >1 Swap UndoNum with UndoNum-1
' Dim a$
Dim prevcanvasW As Long
Dim prevcanvasH As Long
   If UndoNum > 1 Then
      ' Swap filspecs
      a$ = CopyFileSpec$(UndoNum - 1)
      CopyFileSpec$(UndoNum - 1) = CopyFileSpec$(UndoNum)
      CopyFileSpec$(UndoNum) = a$
      ' Swap colors
      CopyMemory CulBGR(0), CopyRGB(0, UndoNum - 1), 1024
      CopyMemory CopyRGB(0, UndoNum - 1), CopyRGB(0, UndoNum), 1024
      CopyMemory CopyRGB(0, UndoNum), CulBGR(0), 1024
      CulRGB() = CulBGR()
      CopyToBGR
      ' Swap bUndo[UndoNum-1],bUndo[UndoNum]
      ' bUndo[UndoNum] = current bArray()
      Select Case UndoNum
      Case 2
         prevcanvasW = UBound(bUndo1(), 1)
         prevcanvasH = UBound(bUndo1(), 2)
         ReDim bUndo2(prevcanvasW, prevcanvasH)
         bUndo2() = bUndo1()
         ReDim bUndo1(canvasW, canvasH)
         bUndo1() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo2()
      Case 3
         prevcanvasW = UBound(bUndo2(), 1)
         prevcanvasH = UBound(bUndo2(), 2)
         ReDim bUndo3(prevcanvasW, prevcanvasH)
         bUndo3() = bUndo2()
         ReDim bUndo2(canvasW, canvasH)
         bUndo2() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo3()
      Case 4
         prevcanvasW = UBound(bUndo3(), 1)
         prevcanvasH = UBound(bUndo3(), 2)
         ReDim bUndo4(prevcanvasW, prevcanvasH)
         bUndo4() = bUndo3()
         ReDim bUndo3(canvasW, canvasH)
         bUndo3() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo4()
      Case 5
         prevcanvasW = UBound(bUndo4(), 1)
         prevcanvasH = UBound(bUndo4(), 2)
         ReDim bUndo5(prevcanvasW, prevcanvasH)
         bUndo5() = bUndo4()
         ReDim bUndo4(canvasW, canvasH)
         bUndo4() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo5()
      Case 6
         prevcanvasW = UBound(bUndo5(), 1)
         prevcanvasH = UBound(bUndo5(), 2)
         ReDim bUndo6(prevcanvasW, prevcanvasH)
         bUndo6() = bUndo5()
         ReDim bUndo5(canvasW, canvasH)
         bUndo5() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo6()
      Case 7
         prevcanvasW = UBound(bUndo6(), 1)
         prevcanvasH = UBound(bUndo6(), 2)
         ReDim bUndo7(prevcanvasW, prevcanvasH)
         bUndo7() = bUndo6()
         ReDim bUndo6(canvasW, canvasH)
         bUndo6() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo7()
      Case 8
         prevcanvasW = UBound(bUndo7(), 1)
         prevcanvasH = UBound(bUndo7(), 2)
         ReDim bUndo8(prevcanvasW, prevcanvasH)
         bUndo8() = bUndo7()
         ReDim bUndo7(canvasW, canvasH)
         bUndo7() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo8()
      Case 9
         prevcanvasW = UBound(bUndo8(), 1)
         prevcanvasH = UBound(bUndo8(), 2)
         ReDim bUndo9(prevcanvasW, prevcanvasH)
         bUndo9() = bUndo8()
         ReDim bUndo8(canvasW, canvasH)
         bUndo8() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo9()
      Case 10
         prevcanvasW = UBound(bUndo9(), 1)
         prevcanvasH = UBound(bUndo9(), 2)
         ReDim bUndo10(prevcanvasW, prevcanvasH)
         bUndo10() = bUndo9()
         ReDim bUndo9(canvasW, canvasH)
         bUndo9() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo10()
      Case 11
         prevcanvasW = UBound(bUndo10(), 1)
         prevcanvasH = UBound(bUndo10(), 2)
         ReDim bUndo11(prevcanvasW, prevcanvasH)
         bUndo11() = bUndo10()
         ReDim bUndo10(canvasW, canvasH)
         bUndo10() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo11()
      Case 12
         prevcanvasW = UBound(bUndo11(), 1)
         prevcanvasH = UBound(bUndo11(), 2)
         ReDim bUndo12(prevcanvasW, prevcanvasH)
         bUndo12() = bUndo11()
         ReDim bUndo11(canvasW, canvasH)
         bUndo11() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo12()
      Case 13
         prevcanvasW = UBound(bUndo12(), 1)
         prevcanvasH = UBound(bUndo12(), 2)
         ReDim bUndo13(prevcanvasW, prevcanvasH)
         bUndo13() = bUndo12()
         ReDim bUndo12(canvasW, canvasH)
         bUndo12() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo13()
      Case 14
         prevcanvasW = UBound(bUndo13(), 1)
         prevcanvasH = UBound(bUndo13(), 2)
         ReDim bUndo14(prevcanvasW, prevcanvasH)
         bUndo14() = bUndo13()
         ReDim bUndo13(canvasW, canvasH)
         bUndo13() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo14()
      Case 15
         prevcanvasW = UBound(bUndo14(), 1)
         prevcanvasH = UBound(bUndo14(), 2)
         ReDim bUndo15(prevcanvasW, prevcanvasH)
         bUndo15() = bUndo14()
         ReDim bUndo14(canvasW, canvasH)
         bUndo14() = bArray()
         canvasW = prevcanvasW
         canvasH = prevcanvasH
         ReDim bArray(canvasW, canvasH)
         bArray() = bUndo15()
      End Select
   End If
End Sub

Public Sub CopyToBGR()
' Make BGR from RGB
Dim k As Long
Dim Culr As Long
   For k = 0 To 255
      Culr = CulRGB(k)
      palRed(k) = (Culr And &HFF&)
      palGreen(k) = (Culr And &HFF00&) / &H100&
      palBlue(k) = (Culr And &HFF0000) / &H10000
      CulBGR(k) = RGB(palBlue(k), palGreen(k), palRed(k))
   Next k
End Sub

'#### Display thumbnails in frmViews ###############

Public Sub DISPLAY_ALL_VIEWS()
Dim k As Long
   For k = 1 To 15
      frmViews.picTV(k - 1).Picture = LoadPicture
      frmViews.LabVN(k - 1).BackColor = vbWhite
   Next k
   
   For k = 1 To TopUndoNum
      DisplayUndoNum k
   Next k
   Erase bTemp()
   frmViews.picTemp.Width = 8
   frmViews.picTemp.Height = 8
   frmViews.picTemp.Picture = LoadPicture
   frmViews.LabVN(UndoNum - 1).BackColor = vbYellow
End Sub

Public Sub DISPLAY_VIEW(NUM As Long)
' NUM = UndoNum
Dim k As Long
   For k = 1 To 15
      frmViews.LabVN(k - 1).BackColor = vbWhite
   Next k
      
   frmViews.picTV(NUM - 1).Picture = LoadPicture
   frmViews.LabVN(NUM - 1).BackColor = vbWhite
   k = NUM
   
   DisplayUndoNum k
   
   Erase bTemp()
   frmViews.picTemp.Width = 8
   frmViews.picTemp.Height = 8
   frmViews.picTemp.Picture = LoadPicture
   frmViews.LabVN(UndoNum - 1).BackColor = vbYellow
End Sub

Private Sub DisplayUndoNum(UND As Long)
Dim WTV As Long, HTV As Long
   Select Case UND
   Case 1
      WTV = UBound(bUndo1(), 1): HTV = UBound(bUndo1(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo1()
      DISPVIEW UND, WTV, HTV
   Case 2
      WTV = UBound(bUndo2(), 1): HTV = UBound(bUndo2(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo2()
      DISPVIEW UND, WTV, HTV
   Case 3
      WTV = UBound(bUndo3(), 1): HTV = UBound(bUndo3(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo3()
      DISPVIEW UND, WTV, HTV
   Case 4
      WTV = UBound(bUndo4(), 1): HTV = UBound(bUndo4(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo4()
      DISPVIEW UND, WTV, HTV
   Case 5
      WTV = UBound(bUndo5(), 1): HTV = UBound(bUndo5(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo5()
      DISPVIEW UND, WTV, HTV
   Case 6
      WTV = UBound(bUndo6(), 1): HTV = UBound(bUndo6(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo6()
      DISPVIEW UND, WTV, HTV
   Case 7
      WTV = UBound(bUndo7(), 1): HTV = UBound(bUndo7(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo7()
      DISPVIEW UND, WTV, HTV
   Case 8
      WTV = UBound(bUndo8(), 1): HTV = UBound(bUndo8(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo8()
      DISPVIEW UND, WTV, HTV
   Case 9
      WTV = UBound(bUndo9(), 1): HTV = UBound(bUndo9(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo9()
      DISPVIEW UND, WTV, HTV
   Case 10
      WTV = UBound(bUndo10(), 1): HTV = UBound(bUndo10(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo10()
      DISPVIEW UND, WTV, HTV
   Case 11
      WTV = UBound(bUndo11(), 1): HTV = UBound(bUndo11(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo11()
      DISPVIEW UND, WTV, HTV
   Case 12
      WTV = UBound(bUndo12(), 1): HTV = UBound(bUndo12(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo12()
      DISPVIEW UND, WTV, HTV
   Case 13
      WTV = UBound(bUndo13(), 1): HTV = UBound(bUndo13(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo13()
      DISPVIEW UND, WTV, HTV
   Case 14
      WTV = UBound(bUndo14(), 1): HTV = UBound(bUndo14(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo14()
      DISPVIEW UND, WTV, HTV
   Case 15
      WTV = UBound(bUndo15(), 1): HTV = UBound(bUndo15(), 2)
      ReDim bTemp(WTV, HTV): bTemp() = bUndo15()
      DISPVIEW UND, WTV, HTV
   End Select
End Sub

Private Sub DISPVIEW(UND As Long, WTV As Long, HTV As Long)
Dim i As Long
Dim Culr As Long
Dim WThumb As Long, HThumb As Long
Dim WD As Long, HD As Long
Dim bS As BITMAPINFO
   
   FlushMouseEvents Form1.PIC.hWnd
   
   'RGBQUAD is B G R Res
   For i = 0 To 255
      Culr = CopyRGB(i, UND)
      bS.Colors(i).rgbRed = (Culr And &HFF&)
      bS.Colors(i).rgbGreen = (Culr And &HFF00&) / &H100&
      bS.Colors(i).rgbBlue = (Culr And &HFF0000) / &H10000
'      bS.Colors(k).rgbReserved = 0
   Next i
   frmViews.picTemp.Width = WTV 'UBound(bUndo1(), 1)
   frmViews.picTemp.Height = HTV 'UBound(bUndo1(), 2)
   frmViews.picTemp.Picture = LoadPicture
   With bS.bmi
      .biSize = 40
      .biwidth = WTV
      .biheight = HTV
      .biPlanes = 1
      .biBitCount = 8
      .biSizeImage = WTV * HTV 'picTemp.Width * picTemp.Height
   End With
   'DoEvents
 ' EG
'   SetDIBitsToDevice Dhdc, dx,dy dW, dHt, _
'   sx, sy, 0, bArrHt, bArr(1, 1), BS, DIB_RGB_COLORS
  If SetDIBitsToDevice(frmViews.picTemp.hDC, 0, 0, WTV, HTV, _
      0, 0, 0, HTV, bTemp(1, 1), bS, DIB_RGB_COLORS) = 0 Then
      MsgBox "VIEW ERROR", vbCritical, "frmView"
      End
   End If
   frmViews.picTemp.Refresh
   WThumb = frmViews.picTV(UND - 1).Width - 4
   HThumb = frmViews.picTV(UND - 1).Height - 4
   frmViews.picTV(UND - 1).Picture = LoadPicture
   
   ' Maintain aspect ratio
   If WTV <= WThumb And HTV <= HThumb Then
      WD = WTV: HD = HTV
   ElseIf WTV >= HTV Then
      WD = WThumb
      HD = HTV * (WThumb / WTV)
   ElseIf WTV < HTV Then
      HD = HThumb
      WD = WTV * (HThumb / HTV)
   End If
   
   'SetStretchBltMode frmViews.picTV(UND - 1).hDC, HALFTONE     ' thin lines can disappear
   SetStretchBltMode frmViews.picTV(UND - 1).hDC, COLORONCOLOR  ' less quality but shows something of thin lines
   'StretchBlt Dhdc,xd,yd,dw,dh,Shdc,xs,ys,sw,sh,vbSrcCopy
   If StretchBlt(frmViews.picTV(UND - 1).hDC, 0, 0, WD, HD, _
      frmViews.picTemp.hDC, 0, 0, WTV, HTV, vbSrcCopy) = 0 Then
      MsgBox "VIEW STRETCH ERROR", vbCritical, "frmView"
      End
   End If
   frmViews.picTV(UND - 1).Refresh
End Sub


