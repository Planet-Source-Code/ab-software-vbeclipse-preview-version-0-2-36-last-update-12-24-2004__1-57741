VERSION 5.00
Begin VB.UserControl vbePerspective 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vbePerspective.ctx":0000
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "vbePerspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type PointAPI
   x As Long
   y As Long
End Type
                          
Private Declare Function PtInRect Lib "user32.dll" ( _
   ByRef lpRect As RECT, _
   ByVal x As Long, _
   ByVal y As Long _
) As Long

Private Declare Function GetWindowRect Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   ByRef lpRect As RECT _
) As Long

Private Declare Function GetCursorPos Lib "user32.dll" ( _
   ByRef lpPoint As PointAPI _
) As Long

Private Declare Function SetParent Lib "user32.dll" ( _
   ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long _
) As Long

Private m_FolderCount As Long
Private m_ParentHwnd As Long
Private m_ActivePerspectiveId As String
Private m_ActiveViewId As String
Private m_ActiveFolderId As String
Private m_ActiveEditorId As String
Private m_MaximizedFolderId As String
Private m_FolderRects As FolderRectList
Private m_Perspectives As PerspectiveList
Private m_Splitbars As HashTable
Private m_Folders As FolderList
Private m_IsMaximized As Boolean
Private m_DropFolderId As String
Private m_DropRelation As vbRelationship
Private m_Theme As ITheme

Private WithEvents m_ViewSubClass As SubClass
Attribute m_ViewSubClass.VB_VarHelpID = -1

Public Event ActivateView(ByVal ViewId As String)
Public Event ShowView(ByVal ViewId As String)
Public Event CloseView(ByVal ViewId As String)
Public Event ShowContextMenu(ByVal ViewId As String, ByVal Button As Single)

Public Property Get Theme() As ITheme
   
   If m_Theme Is Nothing Then
      Set m_Theme = New ThemeOffice2003
   End If
   
   Set Theme = m_Theme
   
End Property
Public Property Set Theme(ByVal NewTheme As ITheme)
   Set m_Theme = NewTheme
End Property

Public Sub BeginDrag()
      
   Dim Perspective As Perspective
   Set Perspective = New Perspective
      
   If StrComp(ActiveFolderId, Perspective.ID_EDITOR_AREA, vbBinaryCompare) = 0 Then
      'EndDrag
      Exit Sub
   End If
   
   If Not m_IsMaximized Then
      
      Refresh True
   
      tmrDrag.Enabled = True
   
   End If
     
   Set Perspective = Nothing
   
End Sub

Public Sub EndDrag()
   
   Dim l_Perspective As Perspective
   Dim l_DragView As View
   Dim l_DragViewId As String
   Dim l_DragFolder As vbeViewFolder
   Dim l_DropFolder As vbeViewFolder
   Dim l_Folder As Folder
   
   Dim l_DropRc As RECT
   Dim l_DropRcRel As RECT
   Dim l_DropFoRect As FolderRect
   
   With tmrDrag
      .Enabled = False
      .Interval = 500
   End With
   
   ' ----------------------------------------------------------------------
   ' Undraw focus rect
   ' ----------------------------------------------------------------------
   Set l_DropFoRect = m_FolderRects.Item(m_DropFolderId)
         
   If Not l_DropFoRect Is Nothing Then
      With l_DropFoRect
         l_DropRc.Left = .RectLeft
         l_DropRc.Right = .RectRight
         l_DropRc.Top = .RectTop
         l_DropRc.Bottom = .RectBottom
      End With
     
      l_DropRcRel = GetDragRelRect(m_DropRelation, l_DropRc)
      DrawDragRect l_DropRcRel
   End If
      
   ' ----------------------------------------------------------------------
   ' Move view
   ' ----------------------------------------------------------------------
   l_DragViewId = ActiveViewId
   
   Set l_DragFolder = ActiveFolder
   Set l_DropFolder = Folder(m_DropFolderId)
   
   If m_DropRelation = REL_WINDOW Then
   
      SetWindowStyle l_DragFolder.hWnd, True
      SetParent l_DragFolder.hWnd, 0
      GoTo Finally
   
   ElseIf Len(m_DropFolderId) = 0 Then
      ' tool window
      GoTo Finally
   End If
      
   If m_DropRelation = REL_FOLDER Then
      If StrComp(l_DragFolder.FolderId, l_DropFolder.FolderId, vbBinaryCompare) = 0 Then
         ' Don't drop view into its own folder!
         GoTo Finally
      End If
   End If
   
   Set l_DragView = l_DragFolder.FolderLayout.Item(l_DragViewId)
      
   If Not l_DropFolder Is Nothing Then
      
      CloseView l_DragViewId
      
      If m_DropRelation = REL_FOLDER Then ' Add view to a existing folder
         
         l_DropFolder.Add l_DragView.ViewId, l_DragView.View
         
      Else ' Add view to a new folder
      
         m_FolderCount = m_FolderCount + 1
         
         Set l_Perspective = Perspective(m_ActivePerspectiveId)
         Set l_Folder = l_Perspective.CreateViewFolder("NewFolder" & m_FolderCount, m_DropRelation, 0.5, m_DropFolderId)
             l_Folder.Add l_DragView.ViewId, l_DragView.View
             
         l_DropFolder.LastRefFolderId = l_Folder.FolderId
             
         AddFolder l_Folder
         
         Refresh
         
      End If
   End If
   
   ShowView l_DragViewId
        
Finally:
   
   m_DropFolderId = vbNullString
   Screen.MousePointer = vbDefault

   Set l_DragView = Nothing
   Set l_DragFolder = Nothing
   Set l_DropFolder = Nothing
   Set l_Folder = Nothing
   Set l_Perspective = Nothing
   Set l_DropFoRect = Nothing
        
End Sub


Public Property Get MaximizedFolderId() As String
   MaximizedFolderId = m_MaximizedFolderId
End Property
Public Property Let MaximizedFolderId(ByVal NewMaximizedFolderId As String)
   m_MaximizedFolderId = NewMaximizedFolderId
End Property
Public Property Get ActiveViewId() As String
   
   Dim Folder As vbeViewFolder
   Set Folder = m_Folders.Item(ActiveFolderId)
   
   If Not Folder Is Nothing Then
      ActiveViewId = Folder.ActiveViewId
   End If
   
   Set Folder = Nothing
   
End Property
Private Property Get ActiveFolderId() As String
   
   Dim Folder As vbeViewFolder
   Set Folder = ActiveFolder
         
   If Not Folder Is Nothing Then
      ActiveFolderId = Folder.FolderId
   End If
   
   Set Folder = Nothing
   
End Property
Public Property Get ActiveView() As Object
   
   On Error GoTo ErrorHandle
   
   Dim Perspective As Perspective
   Dim Folder As vbeViewFolder
   
   Set Perspective = New Perspective
   Set Folder = m_Folders.Item(ActiveFolderId)
   
   If StrComp(ActiveFolderId, Perspective.ID_EDITOR_AREA, vbBinaryCompare) <> 0 Then
      If Not Folder Is Nothing Then
         Set ActiveView = Folder.ActiveView
      End If
   End If
   
Finally:

   Set Perspective = Nothing
   Set Folder = Nothing
   
   Exit Property
   
ErrorHandle:
   
   GoTo Finally
   
End Property
Public Property Get ActiveEditor() As Object
   
   On Error GoTo ErrorHandle

   Dim Perspective As Perspective
   Dim Folder As vbeViewFolder
   
   Set Perspective = New Perspective
   Set Folder = m_Folders.Item(Perspective.ID_EDITOR_AREA)
   
   If Not Folder Is Nothing Then
      Set ActiveEditor = Folder.ActiveView
   End If
   
Finally:
   
   Set Perspective = Nothing
   Set Folder = Nothing
   
   Exit Property
   
ErrorHandle:

   GoTo Finally

End Property

Private Property Get ActiveFolder() As vbeViewFolder
   
   Dim Folder As Variant
   
   For Each Folder In Folders.Items
      If Folder.Active Then
         Set ActiveFolder = Folder
         GoTo Finally
      End If
   Next
   
Finally:
   
   Set Folder = Nothing
   
End Property

Private Property Get Folder(ByVal FolderId As String) As vbeViewFolder
   
   Dim F As Variant
   
   For Each F In Folders.Items
      If StrComp(F.FolderId, FolderId, vbBinaryCompare) = 0 Then
         Set Folder = F
         Exit Property
      End If
   Next
   
End Property

'Public Property Let ActiveFolderId(ByVal NewActiveFolderId As String)
'   m_ActiveFolderId = NewActiveFolderId
'End Property
Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property
Public Sub ShowView(ByVal ViewId As String)

   Dim View As Variant
   Dim Folder As Variant
   
   If StrComp(ActiveViewId, ViewId, vbBinaryCompare) = 0 Then
      ' The view is allready active!
      GoTo Finally
   End If
   
   For Each Folder In Folders.Items
      If Not Folder.Views.IsEmpty Then
         For Each View In Folder.Views.Items
            If StrComp(View.ViewId, ViewId, vbBinaryCompare) = 0 Then
               Folder.Show ViewId
               RaiseEvent ActivateView(ViewId)
               Exit For
            Else
               Folder.Active = False
            End If
         Next
      End If
   Next
   
Finally:

   Set View = Nothing
   Set Folder = Nothing
   
End Sub

Public Sub CloseView(ByVal ViewId As String)

   Dim View As Variant
   Dim Folder As Variant
   
   For Each Folder In Folders.Items
      If Not Folder.Views.IsEmpty Then
         For Each View In Folder.Views.Items
            If StrComp(View.ViewId, ViewId, vbBinaryCompare) = 0 Then
               
               Folder.Remove ViewId
               
               If Folder.Views.IsEmpty Then
                  RemoveFolder Folder.FolderId
               End If
               
               RaiseEvent CloseView(ViewId)
               GoTo Finally
            End If
         Next
      End If
   Next
   
Finally:
   
   Set View = Nothing
   Set Folder = Nothing
   
End Sub
Public Sub RemoveFolder(ByVal FolderId As String)
   
   Dim i As Long
   Dim l_LastRefRelation As vbRelationship
   Dim l_LastRefRatio As Double
   Dim l_LastRefFolderId As String
   Dim l_IndexOfRemFolderId As Long
   Dim l_IndexOfRefFolderId As Long
   Dim l_FolderControl As Variant
   Dim l_Folder As Folder
   Dim l_Perspective As Perspective
   
   Set l_Perspective = Perspective(m_ActivePerspectiveId)
   
   If StrComp(FolderId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) = 0 Then
      GoTo Finally
   End If
    
   l_LastRefFolderId = Folder(FolderId).LastRefFolderId
   
   'l_IndexOfRemFolderId = Folders.IndexOf(l_LastRefFolderId)
   With l_Perspective.Folders
            
      For i = 1 To .Count
         Set l_Folder = .Item(i)

         If StrComp(l_Folder.FolderId, l_LastRefFolderId, vbBinaryCompare) = 0 Then
            l_IndexOfRefFolderId = i
         ElseIf StrComp(l_Folder.FolderId, FolderId, vbBinaryCompare) = 0 Then
            l_IndexOfRemFolderId = i
         End If
      Next
      
      ' Replace folder
      If l_IndexOfRefFolderId > 0 Then
         With .Item(l_IndexOfRemFolderId)
            .FolderId = l_Perspective.Folders.Item(l_IndexOfRefFolderId).FolderId
'           .Relationship = l_Perspective.Folders.Item(l_IndexOfRefFolderId).Relationship
         End With
         
         .Remove l_IndexOfRefFolderId
      End If
      
      
      If Len(l_LastRefFolderId) = 0 Then
         For i = 1 To .Count
            Set l_Folder = .Item(i)

            If StrComp(l_Folder.FolderId, FolderId, vbBinaryCompare) = 0 Then
               l_LastRefFolderId = l_Folder.RefId
               l_LastRefRelation = l_Folder.Relationship
               l_LastRefRatio = l_Folder.Ratio
               Exit For
            End If
         Next
         
         ' Set Ref ids
         For i = 1 To .Count
            Set l_Folder = .Item(i)

            If StrComp(l_Folder.RefId, FolderId, vbBinaryCompare) = 0 Then
               l_Folder.RefId = l_LastRefFolderId
               l_Folder.Relationship = l_LastRefRelation
               l_Folder.Ratio = l_LastRefRatio
            End If
         Next
         
         
      Else
         
         ' Set Ref ids
         For i = 1 To .Count
            Set l_Folder = .Item(i)

            If StrComp(l_Folder.RefId, FolderId, vbBinaryCompare) = 0 Then
               l_Folder.RefId = l_LastRefFolderId
            End If
         Next
      End If
      
   End With
   
   If Not m_Folders.Item(FolderId) Is Nothing Then
      Controls.Remove FolderId
      m_Folders.Remove FolderId
   End If
            
   If Not m_Splitbars.Item(FolderId) Is Nothing Then
      Controls.Remove "SplitBar_" & FolderId
      m_Splitbars.Remove FolderId
   End If
                     
   l_Perspective.Remove FolderId
           
   m_IsMaximized = False
           
   Refresh True

Finally:

   Set l_Perspective = Nothing
   
End Sub
Public Sub Clear()

   Dim Ctrl As Variant
   
   For Each Ctrl In Controls
      
      If StrComp(TypeName(Ctrl), "vbeViewFolder", vbBinaryCompare) = 0 Then
         Dim Perspective As Perspective
         Set Perspective = New Perspective
         
         If StrComp(Ctrl.FolderId, Perspective.ID_EDITOR_AREA, vbBinaryCompare) <> 0 Then
            Controls.Remove Ctrl
            Set Ctrl = Nothing
         Else
            Ctrl.Visible = False
         End If
         
         Set Perspective = Nothing
      ElseIf StrComp(TypeName(Ctrl), "vbeSplitBar", vbBinaryCompare) = 0 Then
      
         Controls.Remove Ctrl
         Set Ctrl = Nothing
         
      End If
   Next
   
   Set m_Folders = New FolderList
   Set m_Splitbars = New HashTable
   Set m_FolderRects = New FolderRectList
   Set m_Splitbars = New HashTable
   
   m_IsMaximized = False
   m_MaximizedFolderId = vbNullString
   m_ActivePerspectiveId = vbNullString
   m_ActiveViewId = vbNullString
   m_ActiveEditorId = vbNullString
   
End Sub

Public Sub OpenEditor(ByVal EditorId As String, ByRef EditorControl As Object)
   
   Dim FolderLayout As Folder
   Dim FolderControl As Object
   Dim Perspective As Perspective
   
   Set Perspective = New Perspective
   
   Set FolderControl = m_Folders.Item(Perspective.ID_EDITOR_AREA)
          
   If Not FolderControl Is Nothing Then
      modSubClass.Hook EditorControl.hWnd
      FolderControl.Add EditorId, EditorControl
      FolderControl.Show EditorId, True
   End If
   
   Set Perspective = Nothing
   Set FolderControl = Nothing
   Set FolderLayout = Nothing
   
End Sub

Private Sub m_ViewSubClass_ShowView(ByVal hWnd As Long)
   
   Dim Folder As Variant
   Dim View As Variant
   
   For Each Folder In Me.Folders.Items
      For Each View In Folder.Views.Items
         If View.View.hWnd = hWnd Then
            ShowView View.ViewId
         End If
      Next
   Next

End Sub

Private Sub tmrDrag_Timer()
   
   Dim Perspective As Perspective
   Dim FolderFound As Boolean
   Dim rcNew As RECT
   Dim RcOld As RECT
   Dim RcRel As RECT
   Dim Rel As vbRelationship
   Dim NewFoRect As Variant
   Dim OldFoRect As FolderRect
   Dim ViewCount As Long
   Dim P As PointAPI
   
   Set Perspective = New Perspective
     
   tmrDrag.Interval = 100
   
   GetCursorPos P
   
   Screen.MousePointer = 99

   For Each NewFoRect In m_FolderRects.Items
            
      With NewFoRect
         rcNew.Left = .RectLeft
         rcNew.Right = .RectRight
         rcNew.Top = .RectTop
         rcNew.Bottom = .RectBottom
      End With
         
      If PtInRect(rcNew, P.x, P.y) = 1 Then
            
         FolderFound = True
                  
         Rel = REL_FOLDER
         
         If StrComp(NewFoRect.FolderId, ActiveFolderId, vbBinaryCompare) <> 0 Then
            
            Rel = GetRelationship(P, rcNew)
            
         Else
         
            If Not ActiveFolder.Views.IsEmpty Then
               ViewCount = ActiveFolder.Views.Count
            End If
         
            If ViewCount > 0 Then
               Rel = GetRelationship(P, rcNew)
            End If
            
         End If
         
         If StrComp(m_DropFolderId, NewFoRect.FolderId, vbBinaryCompare) <> 0 Or _
            m_DropRelation <> Rel Then
         
            If Len(m_DropFolderId) > 0 Then
                
               Set OldFoRect = m_FolderRects.Item(m_DropFolderId)
                 
               If Not OldFoRect Is Nothing Then
                  
                  ' Redraw old rect
                  With OldFoRect
                     RcOld.Left = .RectLeft
                     RcOld.Right = .RectRight
                     RcOld.Top = .RectTop
                     RcOld.Bottom = .RectBottom
                  End With
                  
                  RcRel = GetDragRelRect(m_DropRelation, RcOld)
                  DrawDragRect RcRel
               End If
            End If
               
            ' Draw new drag rect
            With NewFoRect
               rcNew.Left = .RectLeft
               rcNew.Right = .RectRight
               rcNew.Top = .RectTop
               rcNew.Bottom = .RectBottom
            End With
                       
            If Rel > REL_FOLDER Then
               
               If Rel = REL_BOTTOM Then
                  Screen.MouseIcon = LoadResPicture("ARROW_BOTTOM", vbResCursor)
               ElseIf Rel = REL_TOP Then
                  Screen.MouseIcon = LoadResPicture("ARROW_TOP", vbResCursor)
               ElseIf Rel = REL_LEFT Then
                  Screen.MouseIcon = LoadResPicture("ARROW_LEFT", vbResCursor)
               ElseIf Rel = REL_RIGHT Then
                  Screen.MouseIcon = LoadResPicture("ARROW_RIGHT", vbResCursor)
               End If
               
               RcRel = GetDragRelRect(Rel, rcNew)
               DrawDragRect RcRel
               
            Else
               Screen.MouseIcon = LoadResPicture("ARROW_FOLDER", vbResCursor)
              
               DrawDragRect rcNew
            End If
            
            m_DropFolderId = NewFoRect.FolderId
            m_DropRelation = Rel
           
            Exit For
            
         End If
         
      End If
   Next
   
   If Not FolderFound Then
      
      Set OldFoRect = m_FolderRects.Item(m_DropFolderId)
     
      If Not OldFoRect Is Nothing Then
       
         ' Redraw old rect
         With OldFoRect
            RcOld.Left = .RectLeft
            RcOld.Right = .RectRight
            RcOld.Top = .RectTop
            RcOld.Bottom = .RectBottom
         End With
         
         RcRel = GetDragRelRect(m_DropRelation, RcOld)
         DrawDragRect RcRel
         
         m_DropRelation = REL_WINDOW
         
         Screen.MouseIcon = LoadResPicture("ARROW_OUT", vbResCursor)
         
      End If
      
      m_DropFolderId = vbNullString
      
   End If
   
Finally:

   Set Perspective = Nothing
   Set NewFoRect = Nothing
   Set OldFoRect = Nothing
   
End Sub

Private Function GetRelationship(Pt As PointAPI, Rc As RECT) As vbRelationship
   
   Dim RcRel As RECT
  
   ' Left
   RcRel = GetDragRelRect(REL_LEFT, Rc)
   
   If PtInRect(RcRel, Pt.x, Pt.y) = 1 Then
      GetRelationship = REL_LEFT
      Exit Function
   End If
   
   ' Right
   RcRel = GetDragRelRect(REL_RIGHT, Rc)

   If PtInRect(RcRel, Pt.x, Pt.y) = 1 Then
      GetRelationship = REL_RIGHT
      Exit Function
   End If

   ' Top
   RcRel = GetDragRelRect(REL_TOP, Rc)

   If PtInRect(RcRel, Pt.x, Pt.y) = 1 Then
      GetRelationship = REL_TOP
      Exit Function
   End If

   ' Bottom
   RcRel = GetDragRelRect(REL_BOTTOM, Rc)

   If PtInRect(RcRel, Pt.x, Pt.y) = 1 Then
      GetRelationship = REL_BOTTOM
      Exit Function
   End If
   
   GetRelationship = REL_FOLDER
   
End Function

Private Function GetDragRelRect(Rel As vbRelationship, Rc As RECT) As RECT
   
   Const REL_SIZE As Long = 100
   
   With GetDragRelRect
      .Left = Rc.Left
      .Right = Rc.Right
      .Top = Rc.Top
      .Bottom = Rc.Bottom
   End With
 
   With GetDragRelRect
      If Rel = REL_LEFT Then
         If .Left + REL_SIZE < (.Left + (.Right - .Left) * 0.4) Then
            .Right = .Left + REL_SIZE
         Else
            .Right = .Left + ((.Right - .Left) * 0.4)
         End If
      ElseIf Rel = REL_RIGHT Then
         If .Right - REL_SIZE > (.Right - (.Right - .Left) * 0.4) Then
            .Left = .Right - REL_SIZE
         Else
            .Left = (.Right - (.Right - .Left) * 0.4)
         End If
      ElseIf Rel = REL_TOP Then
         If .Top + REL_SIZE < (.Top + (.Bottom - .Top) * 0.4) Then
            .Bottom = .Top + REL_SIZE
         Else
            .Bottom = (.Top + (.Bottom - .Top) * 0.4)
         End If
      ElseIf Rel = REL_BOTTOM Then
         If .Bottom - REL_SIZE > (.Bottom - (.Bottom - .Top) * 0.4) Then
            .Top = .Bottom - REL_SIZE
         Else
            .Top = (.Bottom - (.Bottom - .Top) * 0.4)
         End If
      End If
   End With
   
End Function

Private Sub UserControl_Initialize()
   Set m_Folders = New FolderList
   Set m_Perspectives = New PerspectiveList
   Set m_Splitbars = New HashTable
   Set m_FolderRects = New FolderRectList
   Set m_ViewSubClass = New SubClass
   Set m_Theme = New ThemeOffice2003
   Set modSubClass.SetSubClass = m_ViewSubClass
End Sub

Private Sub UserControl_Terminate()
   
   Dim i As Long
   
   Set m_Folders = Nothing
   Set m_Perspectives = Nothing
   Set m_Splitbars = Nothing
   Set m_FolderRects = Nothing
   Set m_Theme = Nothing
   Set m_ViewSubClass = Nothing
   Set modSubClass.SetSubClass = Nothing
   
   ' Unhook views + editors
   For i = 0 To m_SubClassList.Count
      modSubClass.UnHook m_SubClassList.Item(i)
   Next i
   
End Sub

Private Sub UserControl_Resize()
   Refresh
End Sub

Public Property Get Perspectives() As IPerspectiveList
   Set Perspectives = m_Perspectives
End Property
Public Property Get Folders() As IFolderList
   Set Folders = m_Folders
End Property

Private Function Perspective(ByVal PerspectiveId As String) As Perspective
   Set Perspective = m_Perspectives.Item(PerspectiveId)
End Function

Public Property Get ParentHwnd() As Long
   ParentHwnd = m_ParentHwnd
End Property

Public Property Let ParentHwnd(ByVal NewParentHwnd As Long)
   m_ParentHwnd = NewParentHwnd
End Property

' Adds a new perspective.
'
' @PerspectiveId
' @Perspective
Public Sub AddPerspective(ByVal PerspectiveId As String, ByVal Perspective As Perspective)
   m_Perspectives.Add PerspectiveId, Perspective
End Sub

' Adds a new folder
'
' @FolderLayout
Private Sub AddFolder(FolderLayout As Folder, Optional ByVal Visible As Boolean = True)

   Dim SplitBar As Object
   Dim FolderControl As Object
   Dim RefFolderControl As Object
  
   ' ------------------------------------------------------------------------------------------
   ' Create view folder
   ' ------------------------------------------------------------------------------------------
   
   On Error Resume Next
   Set FolderControl = Controls(FolderLayout.FolderId)
   On Error GoTo 0
   
   If FolderControl Is Nothing Then
      Set FolderControl = Controls.Add("absVbEclipse.vbeViewFolder", FolderLayout.FolderId)
   End If
   
   With FolderControl
      .Initialize FolderLayout
      .ParentHwnd = UserControl.hWnd
      .TabStop = False
      .Active = False
      Set .Theme = Me.Theme
      .Visible = Visible
   End With
                         
   m_Folders.Add FolderLayout.FolderId, FolderControl
      
   ' Save the last added folder
   If Len(FolderLayout.RefId) > 0 Then
      
      Set RefFolderControl = Controls(FolderLayout.RefId)
      
      If Not RefFolderControl Is Nothing Then
         RefFolderControl.LastRefFolderId = FolderLayout.FolderId
      End If
      
   End If
      
   ' ------------------------------------------------------------------------------------------
   ' Create split bar
   ' ------------------------------------------------------------------------------------------
   If Len(FolderLayout.RefId) > 0 And Not m_Folders.IsEmpty Then
   
   
      On Error Resume Next
      Set SplitBar = Controls("SplitBar_" & FolderLayout.FolderId)
      On Error GoTo 0
   
      If SplitBar Is Nothing Then
         Set SplitBar = Controls.Add("absVbEclipse.vbeSplitBar", "SplitBar_" & FolderLayout.FolderId)
      End If
      
      With SplitBar
         If FolderLayout.Relationship = REL_LEFT Or _
            FolderLayout.Relationship = REL_RIGHT Then
            .Orientation = espVertical
         Else
            .Orientation = espHorizontal
         End If
         
         Set .FolderControl = FolderControl
        
         .TabStop = False
         .Visible = Visible
         
      End With
      
      m_Splitbars.Add FolderLayout.FolderId, SplitBar

   End If

   If Not FolderControl Is Nothing Then
      FolderControl.ZOrder
   End If
End Sub

Public Sub ShowPerspective(ByVal PerspectiveId As String)
      
   Dim l_Editor As Variant
   Dim l_Folder As Folder
   'Dim l_FolderControl As Object
   'Dim l_SplitBar As Object
   Dim l_Perspective As Perspective
   
   Clear
   
   m_ActivePerspectiveId = PerspectiveId
   
   Set l_Perspective = Perspective(PerspectiveId)
   
   If l_Perspective Is Nothing Then
      Err.Raise -1, , "Perspective '" & PerspectiveId & "' not found!"
      Exit Sub
   End If
   
   ' ------------------------------------------------------------------------
   ' Create editor area if visible
   ' ------------------------------------------------------------------------
   If l_Perspective.EditorAreaVisible Then
  
      Set l_Folder = New Folder
          l_Folder.FolderId = l_Perspective.ID_EDITOR_AREA
          l_Folder.Ratio = RATIO_MAX

      AddFolder l_Folder
  
   End If
   
   ' ------------------------------------------------------------------------
   ' Create view folders
   ' ------------------------------------------------------------------------
   For Each l_Folder In l_Perspective.Folders
      AddFolder l_Folder
   Next
   
   Refresh
   
   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   
End Sub

Public Sub Refresh(Optional ByVal SaveRects As Boolean = False)
    
   Dim l_RefControl As Variant
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_FolderControl As Variant
   Dim l_SplitBar As Variant
   
   If Len(m_ActivePerspectiveId) = 0 Then
      Exit Sub
   End If
   
   Set l_Perspective = Perspective(m_ActivePerspectiveId)
   
   If SaveRects Then
      m_FolderRects.Clear
   End If
   
   If Not l_Perspective Is Nothing Then
      
      LockWindow m_ParentHwnd
      
      ' Maximized
      If Len(MaximizedFolderId) > 0 Then
         
         On Error Resume Next
         Set l_FolderControl = Controls(MaximizedFolderId)
         On Error GoTo 0
         
         If Not IsEmpty(l_FolderControl) Then
             l_FolderControl.Move 1, 1, ScaleWidth, ScaleHeight
         End If
         
         If Not m_IsMaximized Then
            
            If Not IsEmpty(m_Splitbars.Values) Then
               For Each l_SplitBar In m_Splitbars.Values
                  l_SplitBar.Visible = False
               Next
            End If
            
            For Each l_Folder In l_Perspective.Folders
               
               On Error Resume Next
               Set l_FolderControl = Controls(l_Folder.FolderId)
               On Error GoTo 0
               
               If StrComp(l_Folder.FolderId, MaximizedFolderId, vbBinaryCompare) = 0 Then
                  l_FolderControl.Visible = True
               Else
                  l_FolderControl.Visible = False
               End If
            Next
         
            If l_Perspective.EditorAreaVisible Then
               Set l_FolderControl = Controls(l_Perspective.ID_EDITOR_AREA)
            
               If StrComp(l_Perspective.ID_EDITOR_AREA, MaximizedFolderId, vbBinaryCompare) = 0 Then
                  l_FolderControl.Visible = True
               Else
                  l_FolderControl.Visible = False
               End If
            End If
            
            m_IsMaximized = True
            
         End If
         
      Else
         
         If m_IsMaximized Then
            If Not IsEmpty(m_Splitbars.Values) Then
               For Each l_SplitBar In m_Splitbars.Values
                  l_SplitBar.Visible = True
               Next
            End If
         End If
         
         m_IsMaximized = False
         
         If l_Perspective.EditorAreaVisible Then
            
            On Error Resume Next
            Set l_FolderControl = Controls(l_Perspective.ID_EDITOR_AREA)
            On Error GoTo 0
                l_FolderControl.Move 1, 1, ScaleWidth, ScaleHeight
                l_FolderControl.Visible = True
      
         End If
      
         For Each l_Folder In l_Perspective.Folders
       
            On Error Resume Next
            Set l_FolderControl = Controls(l_Folder.FolderId)
            On Error GoTo 0
                l_FolderControl.Visible = True
               
            If Len(l_Folder.RefId) > 0 Then
               If l_Perspective.EditorAreaVisible Or _
                  StrComp(l_Folder.RefId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) <> 0 Then
                  On Error Resume Next
                  Set l_RefControl = Controls(l_Folder.RefId)
                  On Error GoTo 0
               End If
            End If
         
            MoveFolderControls l_FolderControl, l_RefControl, l_Folder, SaveRects
            
         Next
      
      End If
   End If
   
   UnLockWindow m_ParentHwnd
   
   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   
End Sub

Private Sub MoveFolderControls(FolderControl As Variant, _
                               RefControl As Variant, _
                               Layout As Folder, _
                Optional ByVal SaveRects As Boolean = False)
   
   Const SPLITBAR_WIDTH As Long = 20
   Dim IsMaxControl As Boolean
   Dim SplitBar As Object
   Dim RatioHeight As Single
   Dim RatioWidth As Single
   Dim RefRect As RECT
   Dim FolderRect As RECT
   Dim RefFolderRect As FolderRect
   Dim NewFolderRect As FolderRect
      
   Set RefFolderRect = New FolderRect
   Set NewFolderRect = New FolderRect
      
   If IsEmpty(RefControl) Then
      IsMaxControl = True
   ElseIf RefControl Is Nothing Then
      IsMaxControl = True
   End If
   
   If Not IsMaxControl Then
      Select Case Layout.Relationship
      
      Case REL_LEFT:
           
           RatioWidth = RefControl.Width * Layout.Ratio
                     
           With RefControl
              FolderControl.Move .Left, .Top, RatioWidth, .Height
           End With
                      
           With FolderControl
              RefControl.Left = FolderControl.Left + RatioWidth
              RefControl.Width = RefControl.Width - RatioWidth
           End With
                      
           FolderControl.Width = FolderControl.Width - SPLITBAR_WIDTH
           RefControl.Width = RefControl.Width - SPLITBAR_WIDTH
           RefControl.Left = RefControl.Left + SPLITBAR_WIDTH
                      
           ' Move SplitBar
           Set SplitBar = m_Splitbars.Item(FolderControl.FolderId, Nothing)
                      
           If Not SplitBar Is Nothing Then
              With FolderControl
                 
                 SplitBar.Move .Left + .Width, .Top, SPLITBAR_WIDTH * 2, .Height
                 SplitBar.Orientation = espVertical
                 
                 If SaveRects Then
                    
                    GetWindowRect RefControl.hWnd, RefRect
                    GetWindowRect FolderControl.hWnd, FolderRect
                 
                    m_FolderRects.Add RefControl.FolderId, CreateFolderRect(RefRect, RefControl.FolderId)
                    m_FolderRects.Add FolderControl.FolderId, CreateFolderRect(FolderRect, FolderControl.FolderId)
                 
                    With SplitBar
                       .RectLeft = FolderRect.Left + 50
                       .RectTop = FolderRect.Top
                       .RectRight = RefRect.Right - 50
                       .RectBottom = RefRect.Bottom
                    End With
                 End If
                 
              End With
           End If
                      
      Case REL_RIGHT:
      
           RatioWidth = RefControl.Width * Layout.Ratio
                     
           With RefControl
              FolderControl.Move .Left + RatioWidth, .Top, .Width - RatioWidth, .Height
              FolderControl.ZOrder
           End With
           
           With FolderControl
              RefControl.Width = RatioWidth
           End With
      
           RefControl.Width = RefControl.Width - SPLITBAR_WIDTH
           FolderControl.Width = FolderControl.Width - SPLITBAR_WIDTH
           FolderControl.Left = FolderControl.Left + SPLITBAR_WIDTH
      
           ' Move SplitBar
           Set SplitBar = m_Splitbars.Item(FolderControl.FolderId, Nothing)
                      
           If Not SplitBar Is Nothing Then
              With FolderControl
                 
                 SplitBar.Move .Left - (SPLITBAR_WIDTH * 2), .Top, SPLITBAR_WIDTH * 2, .Height
                 SplitBar.Orientation = espVertical
              
                 If SaveRects Then
                    
                    GetWindowRect RefControl.hWnd, RefRect
                    GetWindowRect FolderControl.hWnd, FolderRect
                    
                    m_FolderRects.Add RefControl.FolderId, CreateFolderRect(RefRect, RefControl.FolderId)
                    m_FolderRects.Add FolderControl.FolderId, CreateFolderRect(FolderRect, FolderControl.FolderId)
                    
                    With SplitBar
                       .RectLeft = RefRect.Left + 50
                       .RectTop = RefRect.Top
                       .RectRight = FolderRect.Right - 50
                       .RectBottom = FolderRect.Bottom
                    End With
                 End If
                 
              End With
           End If
      
      Case REL_TOP:
      
           RatioHeight = RefControl.Height * Layout.Ratio
           
           With RefControl
              FolderControl.Move .Left, .Top, .Width, .Height - RatioHeight
              .Top = .Top + FolderControl.Height
              .Height = .Height - FolderControl.Height
           End With
      
           RefControl.Height = RefControl.Height - SPLITBAR_WIDTH
           FolderControl.Height = FolderControl.Height - SPLITBAR_WIDTH
           RefControl.Top = RefControl.Top + SPLITBAR_WIDTH
      
           ' Move SplitBar
           Set SplitBar = m_Splitbars.Item(FolderControl.FolderId, Nothing)
                      
           If Not SplitBar Is Nothing Then
              With FolderControl
                 
                 SplitBar.Move .Left, .Top + .Height, .Width, SPLITBAR_WIDTH * 2
                 SplitBar.Orientation = espHorizontal
                 
                 If SaveRects Then
                    
                    GetWindowRect RefControl.hWnd, RefRect
                    GetWindowRect FolderControl.hWnd, FolderRect
                 
                    m_FolderRects.Add RefControl.FolderId, CreateFolderRect(RefRect, RefControl.FolderId)
                    m_FolderRects.Add FolderControl.FolderId, CreateFolderRect(FolderRect, FolderControl.FolderId)
                    
                    With SplitBar
                       .RectLeft = FolderRect.Left
                       .RectTop = FolderRect.Top + 50
                       .RectRight = RefRect.Right
                       .RectBottom = RefRect.Bottom - 50
                    End With
                 End If
                 
              End With
           End If
      
      Case REL_BOTTOM:
           
           RatioHeight = RefControl.Height * Layout.Ratio
           
           With RefControl
              FolderControl.Move .Left, .Top + RatioHeight, .Width, .Height - RatioHeight
              .Height = RatioHeight
           End With
                                 
           RefControl.Height = RefControl.Height - SPLITBAR_WIDTH
           FolderControl.Height = FolderControl.Height - SPLITBAR_WIDTH
           FolderControl.Top = FolderControl.Top + SPLITBAR_WIDTH
           
           ' Move SplitBar
           Set SplitBar = m_Splitbars.Item(FolderControl.FolderId, Nothing)
                      
           If Not SplitBar Is Nothing Then
              With FolderControl
                 
                 SplitBar.Move .Left, .Top - (SPLITBAR_WIDTH * 2), .Width, SPLITBAR_WIDTH * 2
                 SplitBar.Orientation = espHorizontal
                 
                 If SaveRects Then
                    
                    GetWindowRect RefControl.hWnd, RefRect
                    GetWindowRect FolderControl.hWnd, FolderRect
                    
                    m_FolderRects.Add RefControl.FolderId, CreateFolderRect(RefRect, RefControl.FolderId)
                    m_FolderRects.Add FolderControl.FolderId, CreateFolderRect(FolderRect, FolderControl.FolderId)
                    
                    With SplitBar
                       .RectLeft = RefRect.Left
                       .RectTop = RefRect.Top + 50
                       .RectRight = FolderRect.Right
                       .RectBottom = FolderRect.Bottom - 50
                    End With
                 End If
                 
              End With
           End If
        
      End Select
   Else
      FolderControl.Move 1, 1, ScaleWidth, ScaleHeight
   End If
      
   Set RefFolderRect = Nothing
   Set NewFolderRect = Nothing
   Set SplitBar = Nothing
   Set FolderControl = Nothing
   Set RefControl = Nothing
   
End Sub

Private Function CreateFolderRect(rec As RECT, ByVal FolderId As String) As FolderRect
   
   Set CreateFolderRect = New FolderRect
   
       CreateFolderRect.RectLeft = rec.Left
       CreateFolderRect.RectRight = rec.Right
       CreateFolderRect.RectTop = rec.Top
       CreateFolderRect.RectBottom = rec.Bottom
       CreateFolderRect.FolderId = FolderId
       
End Function
