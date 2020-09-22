VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewFavorites 
   Caption         =   "Favorites"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   Icon            =   "frmViewFavorites.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
      _Version        =   393217
      Style           =   7
      Appearance      =   0
   End
End
Attribute VB_Name = "frmViewFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   On Error Resume Next
   Me.TreeView1.Move 1, 1, ScaleWidth, ScaleHeight
End Sub
