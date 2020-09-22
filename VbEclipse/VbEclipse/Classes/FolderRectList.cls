VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderRectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFolderRectList

Dim m_FolderId() As String
Dim m_FolderRect() As FolderRect

Public Function Add(ByVal FolderId As String, FolderRect As FolderRect)
   
   Dim i As Long
   Dim ListCount As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_FolderId)
         If StrComp(m_FolderId(i), FolderId, vbBinaryCompare) = 0 Then
            Set m_FolderRect(i) = FolderRect
            Exit Function
         End If
      Next
   End If
   
   ListCount = Count + 1
   ReDim Preserve m_FolderId(ListCount) As String
   ReDim Preserve m_FolderRect(ListCount) As FolderRect
   m_FolderId(ListCount) = FolderId
   Set m_FolderRect(ListCount) = FolderRect
   
End Function

Public Sub Remove(ByVal FolderId As String)
   
   Dim ListCount As Long
   Dim i As Long
   Dim c As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_FolderId)
         If StrComp(m_FolderId(i), FolderId, vbBinaryCompare) <> 0 Then
            
            m_FolderId(c) = m_FolderId(i)
            Set m_FolderRect(c) = m_FolderRect(i)
            c = c + 1
         End If
      Next
   End If
      
   ListCount = Count
   
   If ListCount > 0 Then
   
      ListCount = ListCount - 1

      ReDim Preserve m_FolderId(ListCount) As String
      ReDim Preserve m_FolderRect(ListCount) As FolderRect
   ElseIf Not IsEmpty Then
      Clear
   End If
   
End Sub

Public Function Item(ByVal Index As Variant) As FolderRect
   
   Dim i As Long
   
   If Not IsEmpty Then
      If IsNumeric(Index) Then
         Set Item = m_FolderRect(Index)
      Else
         For i = 0 To UBound(m_FolderId)
            If StrComp(m_FolderId(i), Index, vbBinaryCompare) = 0 Then
               Set Item = m_FolderRect(i)
               Exit Function
            End If
         Next
      End If
   End If
   
   'Set Item = Nothing
   
End Function

Public Function Items() As Variant
   
   Items = m_FolderRect
   
End Function

Public Function Count() As Long
   
   On Error GoTo ErrorHandle
   
   Count = UBound(m_FolderId)
   
   Exit Function
   
ErrorHandle:
   
   Count = -1

End Function

Public Function IsEmpty() As Boolean
   IsEmpty = (Count = -1)
End Function
Public Sub Clear()
   Erase m_FolderRect
   Erase m_FolderId
End Sub

Private Function IFolderRectList_Count() As Long
   IFolderRectList_Count = Count
End Function

Private Function IFolderRectList_IsEmpty() As Boolean
   IFolderRectList_IsEmpty = IsEmpty
End Function

Private Function IFolderRectList_Item(ByVal Index As Variant) As FolderRect
   Set IFolderRectList_Item = Item(Index)
End Function

Private Function IFolderRectList_Items() As Variant
   Set IFolderRectList_Items = Items()
End Function

