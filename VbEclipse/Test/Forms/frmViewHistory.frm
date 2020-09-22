VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewHistory 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "History"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmViewHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmViewHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   lvwHistory.ListItems.Add , , "http://www.pscode.com"
   lvwHistory.ListItems.Add , , "http://www.oscl.de"
End Sub

Private Sub Form_Resize()
   lvwHistory.Move 1, 1, ScaleWidth, ScaleHeight
End Sub

Private Sub lvwHistory_DblClick()
  
  Dim l As ListItem
  Dim v As View
  
  Set l = Me.lvwHistory.SelectedItem
  
  If Not l Is Nothing Then
  
     Set v = frmMain.Perspective1.ActiveEditor
  
     If Not v Is Nothing Then
     
        If StrComp(TypeName(v.View), "frmBrowser", vbBinaryCompare) = 0 Then
           Dim frm As frmEditorBrowser
           Set frm = v.View
           frm.WebBrowser1.Navigate l.Text
           Set frm = Nothing
        End If

     End If
  End If
  
  Set l = Nothing
  Set v = Nothing
  
End Sub
