VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewBookmarks 
   Caption         =   "Bookmarks"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   Icon            =   "frmViewBookmarks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView lstBookmarks 
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmViewBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
   On Error Resume Next
   
   Me.lstBookmarks.Move 1, 1, ScaleWidth, ScaleHeight
   With Me.lstBookmarks.ColumnHeaders
      .Item(2).Width = ScaleWidth - .Item(1).Width - .Item(3).Width
   End With
End Sub
