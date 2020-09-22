VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_ViewId() As String
Dim m_View() As View

Public Function Add(ByVal ViewId As String, View As View)
   
   Dim i As Long
   Dim ListCount As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_ViewId)
         If StrComp(m_ViewId(i), ViewId, vbBinaryCompare) = 0 Then
            Set m_View(i) = View
            Exit Function
         End If
      Next
   End If
   
   ListCount = Count + 1
   ReDim Preserve m_ViewId(ListCount) As String
   ReDim Preserve m_View(ListCount) As View
   m_ViewId(ListCount) = ViewId
   Set m_View(ListCount) = View
   
End Function

Public Sub Remove(ByVal ViewId As String)
   
   Dim ListCount As Long
   Dim i As Long
   Dim c As Long
   
   If Not IsEmpty Then
      For i = 0 To UBound(m_ViewId)
         If StrComp(m_ViewId(i), ViewId, vbBinaryCompare) <> 0 Then
            
            m_ViewId(c) = m_ViewId(i)
            Set m_View(c) = m_View(i)
            c = c + 1
         End If
      Next
   End If
      
   ListCount = Count
   
   If ListCount > 0 Then
   
      ListCount = ListCount - 1

      ReDim Preserve m_ViewId(ListCount) As String
      ReDim Preserve m_View(ListCount) As View
   ElseIf Not IsEmpty Then
      Clear
   End If
   
End Sub

Public Function Item(ByVal Index As Variant) As View
   
   Dim i As Long
   
   If Not IsEmpty Then
      If IsNumeric(Index) Then
         Set Item = m_View(Index)
      Else
         For i = 0 To UBound(m_ViewId)
            If StrComp(m_ViewId(i), Index, vbBinaryCompare) = 0 Then
               Set Item = m_View(i)
               Exit Function
            End If
         Next
      End If
   End If
   
   'Set Item = Nothing
   
End Function

Public Function Items() As Variant
   
   Items = m_View
   
End Function

Public Function Count() As Long
   
   On Error GoTo ErrorHandle
   
   Count = UBound(m_ViewId)
   
   Exit Function
   
ErrorHandle:
   
   Count = -1

End Function

Public Function IsEmpty() As Boolean
   IsEmpty = (Count = -1)
End Function
Public Sub Clear()
   Erase m_View
   Erase m_ViewId
End Sub
