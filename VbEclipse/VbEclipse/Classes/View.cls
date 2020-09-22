VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Panel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Icon As Long
Private m_Caption As String
Private m_ViewId As String
Private m_View As Object

' Get the view.
'
' @View
Public Property Get View() As Object
   Set View = m_View
End Property
' Set the view.
'
' @NewView
Public Property Set View(ByRef NewView As Object)
   Set m_View = NewView
End Property

' Get the view id.
Public Property Get ViewId() As String
   ViewId = m_ViewId
End Property
' Set the view id.
Public Property Let ViewId(ByVal NewViewId As String)
   m_ViewId = NewViewId
End Property

' Get the caption.
Public Property Get Caption() As String
   Caption = m_Caption
End Property
' Set the caption.
Public Property Let Caption(ByRef NewCaption As String)
   m_Caption = NewCaption
End Property

' Get the icon.
Public Property Get Icon() As Long
   Icon = m_Icon
End Property
' Set the icon.
Public Property Let Icon(ByRef NewIcon As Long)
   m_Icon = NewIcon
End Property


