VERSION 5.00
Begin VB.Form frmEditorCode 
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   Icon            =   "frmEditorCode.frx":0000
   ScaleHeight     =   10290
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmEditorCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadFile(ByVal file As String)
   
   Dim str As String
   Dim strLine As String
   Dim f As Long
   
   If file <> "" Then
      f = FreeFile
      Open file For Input As #f
         Do While Not EOF(f)
            Line Input #f, strLine
            str = str & vbNewLine & strLine
         Loop
      Close #f
      
      Me.txtCode.Text = str
      
   End If
   
End Sub

Private Sub Form_Resize()
   Me.txtCode.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
