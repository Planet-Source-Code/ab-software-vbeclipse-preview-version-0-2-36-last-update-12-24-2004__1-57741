VERSION 5.00
Begin VB.UserControl vbeSplitBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   MousePointer    =   6  'Größenänderung NO SW
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "vbeSplitBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private m_Rect As RECT
Private m_RectLeft As Long
Private m_RectRight As Long
Private m_RectTop As Long
Private m_RectBottom As Long
Private m_FolderControl As vbeViewFolder
Private m_Orientation As eOrientationConstants

Private WithEvents SplitBar As SplitBar
Attribute SplitBar.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
   Set SplitBar = New SplitBar
   Orientation = Orientation
End Sub

Private Sub UserControl_Terminate()
   Set SplitBar = Nothing
End Sub

Public Property Get FolderControl() As vbeViewFolder
   Set FolderControl = m_FolderControl
End Property
Public Property Set FolderControl(ByVal NewFolderControl As vbeViewFolder)
   Set m_FolderControl = NewFolderControl
End Property

Public Property Get RectLeft() As Long
   RectLeft = m_RectLeft
End Property
Public Property Let RectLeft(ByVal NewRectLeft As Long)
   m_RectLeft = NewRectLeft
End Property

Public Property Get RectRight() As Long
   RectRight = m_RectRight
End Property
Public Property Let RectRight(ByVal NewRectRight As Long)
   m_RectRight = NewRectRight
End Property

Public Property Get RectTop() As Long
   RectTop = m_RectTop
End Property
Public Property Let RectTop(ByVal NewRectTop As Long)
   m_RectTop = NewRectTop
End Property

Public Property Get RectBottom() As Long
   RectBottom = m_RectBottom
End Property
Public Property Let RectBottom(ByVal NewRectBottom As Long)
   m_RectBottom = NewRectBottom
End Property

Public Property Get Orientation() As eOrientationConstants
   Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal NewOrientation As eOrientationConstants)
   
   m_Orientation = NewOrientation
   
   If NewOrientation = espHorizontal Then
      UserControl.MousePointer = 7
   Else
      UserControl.MousePointer = 9
   End If
   
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   Parent.Refresh True
   
   With m_Rect
      .Left = m_RectLeft
      .Top = m_RectTop
      .Bottom = m_RectBottom
      .Right = m_RectRight
   End With
   
   With SplitBar
      .Orientation = m_Orientation
      .SplitterMouseDown UserControl.hwnd, m_Rect, x, y
   End With
  
End Sub

Private Sub SplitBar_AfterResize(ByVal NewSize As Long)
   
   Dim FolderLayout As Folder
   Dim FolderRect As RECT
   Dim OldRect As RECT
   Dim NewRatio As Double
      
   OldRect = m_Rect
   GetWindowRect FolderControl.hwnd, FolderRect
   
   With FolderControl.FolderLayout
      
      Select Case .Relationship
         Case REL_LEFT, REL_RIGHT:
                           
              OldRect.Left = OldRect.Left - 30
              OldRect.Right = OldRect.Right + 30
              
         Case Else:
         
              OldRect.Top = OldRect.Top - 30
              OldRect.Bottom = OldRect.Bottom + 30
         
      End Select
      
      Select Case .Relationship
         Case REL_LEFT:     NewRatio = (NewSize - OldRect.Left) / (OldRect.Right - OldRect.Left)
         Case REL_RIGHT:    NewRatio = (NewSize - OldRect.Left) / (OldRect.Right - OldRect.Left)
         Case REL_TOP:      NewRatio = 1 - ((NewSize - OldRect.Top) / (OldRect.Bottom - OldRect.Top))
         Case REL_BOTTOM:   NewRatio = (NewSize - OldRect.Top) / (OldRect.Bottom - OldRect.Top)
      End Select
      
      .Ratio = NewRatio
      
   End With
   
   Parent.Refresh
   
End Sub

