VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerspectiveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPerspectiveList

Dim m_PerspectiveId() As String
Dim m_Perspective() As Perspective

Public Function Add(ByVal PerspectiveId As String, Perspective As Perspective)
   
   Dim i As Long
   Dim ListCount As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_PerspectiveId)
         If StrComp(m_PerspectiveId(i), PerspectiveId, vbBinaryCompare) = 0 Then
            Set m_Perspective(i) = Perspective
            Exit Function
         End If
      Next
   End If
   
   ListCount = Count + 1
   ReDim Preserve m_PerspectiveId(ListCount) As String
   ReDim Preserve m_Perspective(ListCount) As Perspective
   m_PerspectiveId(ListCount) = PerspectiveId
   Set m_Perspective(ListCount) = Perspective
   
End Function

Public Sub Remove(ByVal PerspectiveId As String)
   
   Dim ListCount As Long
   Dim i As Long
   Dim c As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_PerspectiveId)
         If StrComp(m_PerspectiveId(i), PerspectiveId, vbBinaryCompare) <> 0 Then
            
            m_PerspectiveId(c) = m_PerspectiveId(i)
            Set m_Perspective(c) = m_Perspective(i)
            c = c + 1
         End If
      Next
   End If
      
   ListCount = Count
   
   If ListCount > 0 Then
   
      ListCount = ListCount - 1

      ReDim Preserve m_PerspectiveId(ListCount) As String
      ReDim Preserve m_Perspective(ListCount) As Perspective
   ElseIf Not IsEmpty Then
      Clear
   End If
   
End Sub

Public Function Item(ByVal Index As Variant) As Perspective
   
   Dim i As Long
   
   If Not IsEmpty Then
      If IsNumeric(Index) Then
         Set Item = m_Perspective(Index)
      Else
         For i = 0 To UBound(m_PerspectiveId)
            If StrComp(m_PerspectiveId(i), Index, vbBinaryCompare) = 0 Then
               Set Item = m_Perspective(i)
               Exit Function
            End If
         Next
      End If
   End If
   
   'Set Item = Nothing
   
End Function

Public Function Items() As Perspective
   
   Items = m_Perspective
   
End Function

Public Function Count() As Long
   
   On Error GoTo ErrorHandle
   
   Count = UBound(m_PerspectiveId)
   
   Exit Function
   
ErrorHandle:
   
   Count = -1

End Function

Public Function IsEmpty() As Boolean
   IsEmpty = (Count = -1)
End Function
Public Sub Clear()
   Erase m_Perspective
   Erase m_PerspectiveId
End Sub

Private Function IPerspectiveList_Count() As Long
   IPerspectiveList_Count = Count
End Function

Private Function IPerspectiveList_IsEmpty() As Boolean
   IPerspectiveList_IsEmpty = IsEmpty
End Function

Private Function IPerspectiveList_Item(ByVal Index As Variant) As Perspective
   Set IPerspectiveList_Item = Item(Index)
End Function

Private Function IPerspectiveList_Items() As Variant
   Set IPerspectiveList_Items = Items()
End Function
