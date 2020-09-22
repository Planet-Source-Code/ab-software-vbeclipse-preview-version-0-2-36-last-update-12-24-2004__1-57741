VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_RectLeft As Long
Private m_RectTop As Long
Private m_RectBottom As Long
Private m_RectRight As Long
Private m_FolderId As String

' Sets the rectangle's folder id.
'
' @NewFolderId The new folder id.
Public Property Let FolderId(ByVal NewFolderId As String)
   m_FolderId = NewFolderId
End Property
' Returns the folder id of this rectangle.
Public Property Get FolderId() As String
   FolderId = m_FolderId
End Property

' Sets the rectangle's right coordinate.
'
' @NewRectRight The right coordinate.
Public Property Let RectRight(ByVal NewRectRight As Long)
   m_RectRight = NewRectRight
End Property
' Returns the rectangle's right coordinate.
Public Property Get RectRight() As Long
   RectRight = m_RectRight
End Property

' Sets the rectangle's bottom coordinate.
'
' @NewRectBottom The bottom coordinate.
Public Property Let RectBottom(ByVal NewRectBottom As Long)
   m_RectBottom = NewRectBottom
End Property
' Returns the rectangle's bottom coordinate.
Public Property Get RectBottom() As Long
   RectBottom = m_RectBottom
End Property

' Sets the rectangle's top coordinate.
'
' @NewRectTop The top coordinate.
Public Property Let RectTop(ByVal NewRectTop As Long)
    m_RectTop = NewRectTop
End Property
' Returns the rectangle's top coordinate.
Public Property Get RectTop() As Long
   RectTop = m_RectTop
End Property

' Sets the rectangle's left coordinate.
'
' @NewRectLeft The left coordinate.
Public Property Let RectLeft(ByVal NewRectLeft As Long)
   m_RectLeft = NewRectLeft
End Property
' Returns the rectangle's left coordinate.
Public Property Get RectLeft() As Long
   RectLeft = m_RectLeft
End Property
