VERSION 5.00
Object = "{EC05EDA3-2E90-432E-8F99-E99EB90A5C96}#1.0#0"; "VbEclipse.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Browser Example"
   ClientHeight    =   10680
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Oben ausrichten
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   582
      ButtonWidth     =   609
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin absVbEclipse.vbePerspective Perspective1 
      Height          =   5295
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9340
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenProject 
         Caption         =   "Open Project"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuFileNewBrowser 
         Caption         =   "New Browser"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuViewSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFavorites 
         Caption         =   "Favorites"
      End
      Begin VB.Menu mnuViewBookmarks 
         Caption         =   "Bookmarks"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBrowsePerspective 
         Caption         =   "Browser Perspective"
      End
      Begin VB.Menu mnuViewEditorPerspective 
         Caption         =   "Editor Perspective"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      
   Dim Folder As Folder
   Dim Perspective As Perspective
   
   Me.Perspective1.Visible = False
   
   ' Page1
   Set Perspective = New Perspective
   
   Perspective.EditorAreaVisible = True
             
   Set Folder = Perspective.CreateViewFolder("RightFolder", REL_RIGHT, 0.75, Perspective.ID_EDITOR_AREA)
       Folder.Add "Project", frmViewProject
   
   Set Folder = Perspective.CreateViewFolder("RightBottomFolder", REL_BOTTOM, 0.4, "RightFolder")
       Folder.Add "Properties", frmViewProperties
   
   Set Folder = Perspective.CreateViewFolder("BottomFolder", REL_BOTTOM, 0.75, Perspective.ID_EDITOR_AREA)
       Folder.Add "Console", frmViewConsole
       Folder.Add "Tasks", frmViewTasks
   
   Set Folder = Perspective.CreateViewFolder("LeftTopFolder", REL_LEFT, 0.1, Perspective.ID_EDITOR_AREA)
       Folder.Add "ToolBox", frmViewToolBox
   
   Me.Perspective1.AddPerspective "Editor", Perspective
       
   ' Page2
   Set Perspective = New Perspective
   
   Perspective.EditorAreaVisible = True
   
   Set Folder = Perspective.CreateViewFolder("LeftFolder", REL_LEFT, 0.3, Perspective.ID_EDITOR_AREA)
       Folder.Add "Favorites", frmViewFavorites
       Folder.Add "History", frmViewHistory
   
   Set Folder = Perspective.CreateViewFolder("Bookmarks", REL_BOTTOM, 0.7, Perspective.ID_EDITOR_AREA)
       Folder.Add "Bookmarks", frmViewBookmarks
       Folder.Add "Console", frmViewConsole
   
   Me.Perspective1.AddPerspective "Browser", Perspective
       
   Me.Perspective1.ParentHwnd = Me.hwnd
   Me.Perspective1.ShowPerspective "Editor"
   
   Me.Perspective1.Visible = True
   
   Me.Perspective1.OpenEditor "Welcome", New frmEditorWelcome
   
   Set Perspective = Nothing
   Set Folder = Nothing
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Perspective1.Move 40, Toolbar.Height + 40, ScaleWidth - 80, ScaleHeight - 80 - Toolbar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Unload frmEditorBrowser
   Unload frmEditorWelcome
   Unload frmViewConsole
   Unload frmViewBookmarks
   Unload frmViewHelp
   Unload frmViewFavorites
   Unload frmViewProperties
   Unload frmViewHistory
   Unload frmViewPreview
   
   Set frmMain = Nothing
   Set frmEditorBrowser = Nothing
   Set frmEditorWelcome = Nothing
   Set frmViewBookmarks = Nothing
   Set frmViewConsole = Nothing
   Set frmViewFavorites = Nothing
   Set frmViewHelp = Nothing
   Set frmViewHistory = Nothing
   Set frmViewPreview = Nothing
   Set frmViewProject = Nothing
   Set frmViewProperties = Nothing
   
   Dim frm As Variant
   
   For Each frm In Forms
      Unload frm
      Set frm = Nothing
   Next
   
   End
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
End Sub

Private Sub mnuFileNewBrowser_Click()
   
   Static idx As Integer
   Dim browser As New frmEditorBrowser
   
   browser.Caption = "Planet Souce Code"
   Me.Perspective1.OpenEditor "Browser" & idx, browser
   
   Set browser = Nothing
   
   idx = idx + 1
End Sub

Private Sub mnuFoldersRemoveFolder_Click()
   Me.Perspective1.RemoveFolder "Console"
End Sub

Private Sub mnuFileOpenProject_Click()
   On Error GoTo ErrorHandle
   With DLG
      .Filter = "Projectfiles (*.vbp)|*.vbp|All Files (*.*)|*.*"
      .ShowOpen
      
      frmViewProject.LoadProject .FileName
   End With
   
ErrorHandle:
   
End Sub

Private Sub mnuViewBookmarks_Click()
   Me.Perspective1.ShowView "Bookmarks"
End Sub

Private Sub mnuViewBrowsePerspective_Click()
   Me.Perspective1.ShowPerspective "Browser"
End Sub

Private Sub mnuViewEditorPerspective_Click()
   Me.Perspective1.ShowPerspective "Editor"
End Sub

Private Sub mnuViewFavorites_Click()
   Me.Perspective1.ShowView "Favorites"
End Sub

Private Sub mnuPanelsShowPanel_Click()
   Me.Perspective1.ShowView "Console"
End Sub
Private Sub mnuPanelsRemovePanel_Click()
   Me.Perspective1.CloseView Me.Perspective1.ActiveViewId
End Sub

