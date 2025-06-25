VERSION 5.00
Begin VB.Form frmImagePreview 
   Caption         =   "Preview"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   Icon            =   "frmImagePreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   StartUpPosition =   3  'Windows Default
   Begin prjImageBrowser.ctlImagePreview ctlImagePreview 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6165
      BackColor       =   16777215
   End
End
Attribute VB_Name = "frmImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    
    ctlImagePreview.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub ShowPreview(picImage As StdPicture)
    Set ctlImagePreview.Picture = picImage
    
    Me.Show
End Sub
