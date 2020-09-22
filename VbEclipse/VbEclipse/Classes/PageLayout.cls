VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerspectiveLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum vbRelationship
   REL_LEFT = 1
   REL_RIGHT = 2
   REL_TOP = 3
   REL_BOTTOM = 4
End Enum

Public Enum vbRatio
   RATIO_MIN = 0.05
   RATIO_HALF = 0.5
   RATIO_QUATER = 0.25
   RATIO_MAX = 0.95
   RATIO_DEFAULT = RATIO_HALF
End Enum

Private m_PageId As String
Private m_PageCaption As String
Private m_EditorAreaVisible As Boolean
Private m_Folders As Collection

' Initialize
Private Sub Class_Initialize()
   Set m_Folders = New Collection
End Sub

' Terminate
Private Sub Class_Terminate()
   Set m_Folders = Nothing
End Sub

Public Function ID_EDITOR_AREA() As String
   ID_EDITOR_AREA = "com_absoftware_ui_editorarea"
End Function

Public Property Get Folders() As Collection
   Set Folders = m_Folders
End Property

Public Property Get EditorAreaVisible() As Boolean
   EditorAreaVisible = m_EditorAreaVisible
End Property
Public Property Let EditorAreaVisible(ByVal NewEditorAreaVisible As Boolean)
   m_EditorAreaVisible = NewEditorAreaVisible
End Property

Public Function CreateViewFolder(ByVal FolderId As String, _
                                  ByVal Relationship As vbRelationship, _
                                  ByVal Ratio As Double, _
                                  ByVal RefId As String) As Folder
   
   On Error GoTo ErrorHandle
   
   Dim Folder As Folder
   Set Folder = New Folder
   
   Folder.FolderId = FolderId
   Folder.Ratio = Ratio
   Folder.RefId = RefId
   Folder.Relationship = Relationship
   
   Set CreateViewFolder = Folder
   
   m_Folders.Add Folder, FolderId
   
Finally:
   
   Set Folder = Nothing
   
   Exit Function
   
ErrorHandle:
   
   GoTo Finally
   
End Function

Public Function Remove(ByVal FolderId As String)
   
   m_Folders.Remove FolderId
   
End Function
