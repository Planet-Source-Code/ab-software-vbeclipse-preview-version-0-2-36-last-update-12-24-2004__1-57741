VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFolderList

Dim m_FolderId() As String
Dim m_Folder() As Object

' Adds a new folder to the folder list.
'
' @FolderId The folder id.
' @Folder The folder control.
Public Function Add(ByVal FolderId As String, Folder As Object)
   
   Dim i As Long
   Dim ListCount As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_FolderId)
         If StrComp(m_FolderId(i), FolderId, vbBinaryCompare) = 0 Then
            Set m_Folder(i) = Folder
            Exit Function
         End If
      Next
   End If
   
   ListCount = Count + 1
   ReDim Preserve m_FolderId(ListCount) As String
   ReDim Preserve m_Folder(ListCount) As Object
   m_FolderId(ListCount) = FolderId
   Set m_Folder(ListCount) = Folder
   
End Function

' Adds a new folder to the folder list.
'
' @FolderId The folder id.
' @Folder The folder control.
Public Function Replace(ByVal ReplaceFolderId As String, ByVal WithFolderId As String)
   
   Dim i As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_FolderId)
         If StrComp(m_FolderId(i), ReplaceFolderId, vbBinaryCompare) = 0 Then
            Set m_Folder(i) = Item(WithFolderId)
            Exit Function
         End If
      Next
   End If
   
End Function

' Removes a folder from the folder list.
'
' @FolderId The folder id.
Public Sub Remove(ByVal FolderId As String)
   
   Dim ListCount As Long
   Dim i As Long
   Dim c As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_FolderId)
         If StrComp(m_FolderId(i), FolderId, vbBinaryCompare) <> 0 Then
            
            m_FolderId(c) = m_FolderId(i)
            Set m_Folder(c) = m_Folder(i)
            c = c + 1
         End If
      Next
   End If
   
   ListCount = Count
   
   If ListCount > 0 Then
      ListCount = ListCount - 1
   
      ReDim Preserve m_FolderId(ListCount) As String
      ReDim Preserve m_Folder(ListCount) As Object
   Else
      Clear
   End If
   
End Sub

' Returns a folder of the folder list.
'
' @Index A numeric index or the folder id.
' @Item The folder control for the numeric index or the folder id.
Public Function Item(ByVal Index As Variant) As Object
   
   Dim i As Long
   
   If Not IsEmpty Then
      If IsNumeric(Index) Then
         Set Item = m_Folder(Index)
      Else
         For i = 0 To UBound(m_FolderId)
            If StrComp(m_FolderId(i), Index, vbBinaryCompare) = 0 Then
               Set Item = m_Folder(i)
               Exit Function
            End If
         Next
      End If
   End If
   
   Set Item = Nothing
   
End Function

' Returns a list of all folder controls.
'
' @Items List of all folder controls.
Public Function Items() As Variant
   
   Items = m_Folder
   
End Function

' The count of folders.
'
' @Count The folder count.
Public Function Count() As Long
   
   On Error GoTo ErrorHandle
   
   Count = UBound(m_FolderId)
   
   Exit Function
   
ErrorHandle:
   
   Count = -1

End Function

Public Function IndexOf(ByVal FolderId As String) As Long
   
   Dim i As Long
   
   For i = 0 To UBound(m_FolderId)
      If StrComp(m_FolderId(i), FolderId, vbBinaryCompare) = 0 Then
         IndexOf = i
         Exit Function
      End If
   Next i
   
End Function
' Returns true if there are no folders in this folder list; false otherwise.
'
' @IsEmpty  True if folder list is empty; false otherwise.
Public Function IsEmpty() As Boolean
   IsEmpty = (Count = -1)
End Function

' Clears the folder list.
Public Sub Clear()
   Erase m_Folder
   Erase m_FolderId
End Sub

' The count of folders.
'
' @Count The folder count.
Private Function IFolderList_Count() As Long
   IFolderList_Count = Count
End Function

Private Function IFolderList_IndexOf(ByVal FolderId As String) As Long
   IFolderList_IndexOf = IndexOf(FolderId)
End Function

Private Function IFolderList_IsEmpty() As Boolean
   IFolderList_IsEmpty = IsEmpty
End Function

Private Function IFolderList_Item(ByVal Index As Variant) As Object
   Set IFolderList_Item = Item(Index)
End Function

Private Function IFolderList_Items() As Variant
   IFolderList_Items = Items
End Function
