VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewProject 
   Caption         =   "Project Explorer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmViewProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":058A
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":0B24
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":10BE
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":1458
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":19F2
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":1F8C
            Key             =   "UserControl"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":2326
            Key             =   "PropertyPage"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwProject 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   459
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":26C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":2C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":31F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewProject.frx":378E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProject 
      Align           =   1  'Oben ausrichten
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmViewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   On Error Resume Next
   
   tvwProject.Move 1, tbrProject.Height, ScaleWidth, ScaleHeight - tbrProject.Height
End Sub

Public Sub LoadProject(ByVal file As String)
   
   Dim strType As String
   Dim strName As String
   Dim strPath As String
   Dim strLine As String
   Dim pos As Long
   
   Dim f As Long
   
   tvwProject.Nodes.Clear
   
   f = FreeFile
   Open file For Input As #f
       Do While Not EOF(f)
          Line Input #f, strLine
          
          pos = InStr(1, strLine, "=", vbBinaryCompare) - 1
          
          If (pos > 0) Then
             
             strName = vbNullString
             strPath = vbNullString
             strType = Left$(strLine, pos)
             strLine = Mid$(strLine, pos + 2, Len(strLine))
             
             pos = InStr(1, strLine, ";", vbBinaryCompare) - 1
             
             Select Case strType
                Case "Form", "Class", "Module", "UserControl":
                   If pos > 0 Then
                      strName = Left$(strLine, pos)
                      strPath = Right$(strLine, Len(strLine) - (pos + 2))
                   Else
                      strPath = strLine
                   End If
                   
                   AddFile strType, strName, strPath
             
             End Select
             
          End If
          
       Loop
   Close #f
End Sub

Private Sub AddFile(ByVal strType As String, ByVal strName As String, ByVal strPath As String)
    
    On Error Resume Next
    
    Dim n As Node
    
    Set n = Nothing
    Set n = tvwProject.Nodes("Project")
    
    If n Is Nothing Then
       tvwProject.Nodes.Add , , "Project", "Project", "Project", "Project"
       tvwProject.Nodes("Project").Expanded = True
    End If
    
    Set n = Nothing
    Set n = tvwProject.Nodes(strType)
    If n Is Nothing Then
       tvwProject.Nodes.Add "Project", tvwChild, strType, strType, "Folder", "Folder"
    End If
        
    If Len(strName) = 0 Then
       Dim strLine As String
       Dim f As Long
       f = FreeFile
       
       Open strPath For Input As #f
          
          Do While (Not EOF(f) And Len(strName) = 0)
              Line Input #f, strLine
              
              If Left$(strLine, Len("Attribute VB_Name = ")) = "Attribute VB_Name = " Then
                 strName = Right$(strLine, Len(strLine) - Len("Attribute VB_Name = "))
                 strName = Replace(strName, """", "")
              End If
          Loop
          
       Close #f
       
    End If
        
    tvwProject.Nodes.Add strType, tvwChild, strPath, strName, strType, strType
    
End Sub

Private Sub tvwProject_DblClick()
   
   On Error Resume Next
    
   Static ce As Long
   Dim n As Node
    
   Set n = tvwProject.SelectedItem
   
   If Not n Is Nothing Then
      If n.Image <> "Project" And _
         n.Image <> "Folder" Then
         
         Dim frmEdit As New frmEditorCode
         frmEdit.Caption = n.Text
         frmEdit.LoadFile n.Key
         Set frmEdit.Icon = Me.ImageList2.ListImages(n.Image).Picture
         frmMain.Perspective1.OpenEditor "CodeEditor" & ce, frmEdit
         ce = ce + 1
      End If
   End If
   
End Sub
