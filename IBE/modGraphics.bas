Attribute VB_Name = "modGraphics"
Option Explicit

'The following constant is used to convert the picture dimensions to pixel measurement.
'This constant is myterious! I used Screen.TwipsPerPixelX & Screen.TwipsPerPixelY but
'they didn't work as supposed! This constant has been calculated.
Public Const K_DotsPerPixel = 26.4375

Public Const TwipsPerInch = 1440

Public Sub CopyImage(picSource As Picture, objDestination As Object, Optional sX As Long = 0, Optional sY As Long = 0, Optional sWidth As Long = 0, Optional sHeight As Long = 0, Optional dX As Long = 0, Optional dY As Long = 0, Optional dWidth As Long = 0, Optional dHeight As Long = 0, Optional Tile As Boolean = False, Optional TileWidth As Long = 0, Optional TileHeight As Long = 0, Optional ClearDestination As Boolean = True, Optional DoInvisibly As Boolean = False, Optional ShowProgress As Boolean = False, Optional DotsPerPixel As Single = K_DotsPerPixel)
    Dim cRow As Long, cCol As Long
    Dim ObjectVisible As Boolean, ObjectScaleMode As Long

    On Error GoTo ERROR_HANDLER_CopyImage

    ObjectScaleMode = objDestination.ScaleMode
    ObjectVisible = objDestination.Visible

    If DoInvisibly Then objDestination.Visible = False
    objDestination.ScaleMode = 3 'Pixel

    If sWidth = 0 Then sWidth = picSource.Width / DotsPerPixel
    If sHeight = 0 Then sHeight = picSource.Height / DotsPerPixel

'    If dWidth = 0 Then dWidth = objDestination.ScaleWidth - dX
'    If dHeight = 0 Then dHeight = objDestination.ScaleHeight - dY

    If TileWidth = 0 Then TileWidth = sWidth 'dWidth
    If TileHeight = 0 Then TileHeight = sHeight 'dHeight

    If ClearDestination Then objDestination.Cls

    If Tile Then
        For cRow = dY To dY + dHeight - 1 Step TileHeight 'Crawl through rows
            For cCol = dX To dX + dWidth - 1 Step TileWidth 'Crawl through columns
                objDestination.PaintPicture picSource, cCol, cRow, TileWidth, TileHeight, sX, sY, sWidth, sHeight
                If ShowProgress Then DoEvents
            Next
        Next
    Else
        objDestination.PaintPicture picSource, dX, dY, dWidth, dHeight, sX, sY, sWidth, sHeight
    End If

    objDestination.ScaleMode = ObjectScaleMode
    If DoInvisibly Then objDestination.Visible = ObjectVisible

EXIT_CopyImage:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_CopyImage:
    Select Case Err.Number
    Case 5 'No description!
    Case 91 'Object variable or With block variable not set (Source image is not supplied)
    Case Else
        If MsgBox("Error in Sub CopyImage() of Module modGraphics[modGraphics.bas] of Project prjWOWFrame[prjWOWFrame.vbp]" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "prjWOWFrame: Application error!") = vbNo Then Resume EXIT_CopyImage
    End Select
    
    Resume Next
End Sub
