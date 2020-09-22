VERSION 5.00
Begin VB.Form frmViewPreview 
   Caption         =   "Preview"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmViewPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picScreenShot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      Height          =   2535
      Left            =   0
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmViewPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   On Error Resume Next
   Me.picScreenShot.Move 1, 1, ScaleWidth, ScaleHeight
End Sub



