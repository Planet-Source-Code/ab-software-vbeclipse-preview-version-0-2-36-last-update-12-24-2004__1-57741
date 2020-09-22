VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_FolderId As String
Private m_RefId As String
Private m_Ratio As Double
Private m_Relationship As Long
Private m_Views As ViewList

' Initialize class variables
Private Sub Class_Initialize()
   Set m_Views = New ViewList
End Sub

' Terminate class variables
Private Sub Class_Terminate()
   Set m_Views = Nothing
End Sub

' Returns the id of this folder.
'
' @FolderId The folder id.
Public Property Get FolderId() As String
   FolderId = m_FolderId
End Property

' Sets the id for this folder.
'
' @NewFolderId The new folder id.
Public Property Let FolderId(ByVal NewFolderId As String)
   m_FolderId = NewFolderId
End Property

' Returns the reference id of this folder.
'
' @RefId The reference id.
Public Property Get RefId() As String
   RefId = m_RefId
End Property

' Sets the reference id for this folder.
'
' @NewRefId The new reference id.
Public Property Let RefId(ByVal NewRefId As String)
   m_RefId = NewRefId
End Property

' Returns the ratio.
'
' @Ratio
Public Property Get Ratio() As String
   Ratio = m_Ratio
End Property

' Sets the ratio.
'
' @NewRatio
Public Property Let Ratio(ByVal NewRatio As String)
   m_Ratio = NewRatio
End Property

' Returns the relationship.
'
' @Relationship The relationship.
Public Property Get Relationship() As Long
   Relationship = m_Relationship
End Property

' Sets the relationship.
'
' @NewRelationship
Public Property Let Relationship(ByVal NewRelationship As Long)
   m_Relationship = NewRelationship
End Property

' Returns all added views.
'
' @Views
Public Property Get Views() As ViewList
   Set Views = m_Views
End Property

' Adds a new view to this folder.
'
' @ViewId
' @ViewControl
' @Index
Public Function Add(ByVal ViewId As String, ByRef ViewControl As Object, Optional ByVal Index As Long)
   
   Dim View As View
   Set View = New View
   
   Set View.View = ViewControl
       View.Caption = ViewControl.Caption
       View.ViewId = ViewId
       
   m_Views.Add ViewId, View
   
   Set View = Nothing
   
End Function

' Returns a view from this folder.
'
' @ViewId
Public Function Item(ByVal ViewId As String) As View
   
   Set Item = m_Views.Item(ViewId)
   
End Function
' Removes a view from this folder.
'
' @ViewId
Public Function Remove(ByVal ViewId As String)
   
   m_Views.Remove ViewId
   
End Function

' Clears the view list.
Public Function Clear()
   Set m_Views = New ViewList
End Function
