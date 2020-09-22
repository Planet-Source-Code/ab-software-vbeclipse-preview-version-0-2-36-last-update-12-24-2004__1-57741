VERSION 5.00
Begin VB.UserControl vbeViewFolder 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ScaleHeight     =   3600
   ScaleWidth      =   6390
   ToolboxBitmap   =   "vbeViewFolder.ctx":0000
   Begin absVbEclipse.vbeTabStrip TabStrip 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   635
   End
End
Attribute VB_Name = "vbeViewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MoveWindow Lib "user32" ( _
   ByVal hWnd As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long _
) As Long

Private Declare Function ShowWindow Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   ByVal nCmdShow As Long _
) As Long

Private Declare Function SetParent Lib "user32" ( _
   ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long _
) As Long

Private m_Active As Boolean
Private m_Maximized As Boolean
Private m_Theme As ITheme
Private m_ParentHwnd As Long
Private m_LastRefFolderId As String
Private m_FolderLayout As Folder

Public Property Get LastRefFolderId() As String
   LastRefFolderId = m_LastRefFolderId
End Property
Public Property Let LastRefFolderId(ByVal NewLastRefFolderId As String)
   m_LastRefFolderId = NewLastRefFolderId
End Property

Public Property Get ParentHwnd() As Long
   ParentHwnd = m_ParentHwnd
End Property
Public Property Let ParentHwnd(ByVal NewParentHwnd As Long)
   m_ParentHwnd = NewParentHwnd
End Property

Public Property Get FolderId() As String
   FolderId = m_FolderLayout.FolderId
End Property

Public Property Get FolderLayout() As Folder
   Set FolderLayout = m_FolderLayout
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Maximized() As Boolean
   Maximized = m_Maximized
End Property
Public Property Let Maximized(ByVal NewMaximized As Boolean)
   m_Maximized = NewMaximized
End Property

Public Property Get Active() As Boolean
   Active = m_Active
End Property
Public Property Let Active(ByVal NewActive As Boolean)
   
   
   If m_Active <> NewActive Then
      m_Active = NewActive
      
      TabStrip.Active = NewActive
      
      If Active Then
         UserControl.BackColor = Theme.ActiveFrameColor
      Else
         UserControl.BackColor = vbButtonFace 'Theme.InactiveFrameColor
      End If
      
      Refresh
      
   End If
   
End Property

Public Property Get ActiveViewId() As String
   ActiveViewId = TabStrip.ActiveTabKey
End Property

Public Property Get ActiveView() As View
   
   Dim View As Variant
   Dim ViewId As String
   
   ViewId = ActiveViewId
   
   If Len(ViewId) > 0 Then
      For Each View In Views.Items
         If StrComp(View.ViewId, ViewId, vbBinaryCompare) = 0 Then
           Set ActiveView = View
           Exit Property
         End If
      Next
   End If

End Property

Public Property Get Views() As ViewList
   If Not m_FolderLayout Is Nothing Then
      Set Views = m_FolderLayout.Views
   Else
      Set Views = New ViewList
   End If
End Property

Public Property Get Theme() As ITheme

   If m_Theme Is Nothing Then
      Set m_Theme = New ThemeOffice2003
   End If
   
   Set Theme = m_Theme
   
End Property
Public Property Set Theme(ByVal NewTheme As ITheme)
   Set m_Theme = NewTheme
End Property

Public Sub Initialize(FolderLayout As Folder)
   
   Dim View As Variant
   
   Set m_FolderLayout = FolderLayout
   
   If Not Views.IsEmpty Then
      For Each View In Views.Items
 
         modSubClass.Hook View.View.hWnd
         TabStrip.Add View.ViewId, View.Caption, View.View.Icon
      
         SetWindowStyle View.View.hWnd, False
         SetParent View.View.hWnd, UserControl.hWnd
      
      Next
   End If
   
   Refresh
   
End Sub
Public Sub Add(ByVal ViewId As String, ByRef ViewControl As Object)
       
   If m_FolderLayout Is Nothing Then
      Set m_FolderLayout = New Folder
   End If
   
   m_FolderLayout.Add ViewId, ViewControl
 
   TabStrip.Add ViewId, ViewControl.Caption, ViewControl.Icon
      
   SetWindowStyle ViewControl.hWnd, False
   SetParent ViewControl.hWnd, UserControl.hWnd
   ViewControl.Visible = True
   
   Refresh

End Sub

Public Sub Remove(ByVal ViewId As String)
         
   Dim View As View
       
   If Not m_FolderLayout Is Nothing Then
      Set View = m_FolderLayout.Item(ViewId)
          View.View.Visible = False
      
      SetParent View.View.hWnd, Parent.hWnd
   
      m_FolderLayout.Remove ViewId
      
      If Not Views.IsEmpty Then
         Parent.ShowView Views.Item(0).ViewId
      Else
         Active = False
      End If
   End If
   
   TabStrip.Remove ViewId
      
   Refresh

End Sub


Public Sub Show(ByVal ViewId As String, Optional ByVal Activate As Boolean = True)
   Active = Activate
   TabStrip.Show ViewId
End Sub

Public Sub Move(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
   
   With UserControl
      .Width = Width
      .Height = Height
   End With
   
End Sub

Public Sub Refresh()
 
   Dim View As Variant

   If Not Views.IsEmpty Then
      For Each View In Views.Items
         View.View.Visible = (StrComp(ActiveViewId, View.ViewId, vbBinaryCompare) = 0)
      Next
   End If
   
   TabStrip.Active = Me.Active
   
   UserControl_Resize
   
End Sub

Private Sub TabStrip_TabClose(ByVal TabKey As String)
   Parent.CloseView TabKey
End Sub

Private Sub TabStrip_TabMouseDown(ByVal TabKey As String, ByVal Button As Single)
   
   Parent.ShowView ActiveViewId
   Parent.BeginDrag
   
   Active = True
   Refresh
   
End Sub
Private Sub TabStrip_TabMouseUp(ByVal TabKey As String, ByVal Button As Single)
         
   Parent.EndDrag
      
End Sub
Private Sub TabStrip_TabDblClick(ByVal TabKey As String)
   
   Maximized = Not Maximized
   
   If Maximized Then
      Parent.MaximizedFolderId = m_FolderLayout.FolderId
   Else
      Parent.MaximizedFolderId = vbNullString
   End If
   
   Parent.Refresh
   
End Sub

Private Sub UserControl_Initialize()
   Active = Not Active
End Sub

Private Sub UserControl_Resize()
   
   On Error Resume Next
   
   Dim View As Variant
   
   UserControl.AutoRedraw = True
   UserControl.Cls
   UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight - 10), Theme.InactiveFrameColor, B
   
   If Not Views.IsEmpty Then
      For Each View In Views.Items
         If View.View.Visible Then
            View.View.Move 40, TabStrip.Height + 30, ScaleWidth - 90, ScaleHeight - TabStrip.Height - 80
            Exit For
         End If
      Next
   End If
   
   UserControl.AutoRedraw = False
   
End Sub

Private Sub UserControl_Terminate()
   
   Dim View As Variant
   
   If Not Views.IsEmpty Then
      ' Don't unload view controls; just hide them
      For Each View In Views.Items
         View.View.Visible = False
         SetParent View.View.hWnd, m_ParentHwnd
      Next
   End If
   
End Sub
