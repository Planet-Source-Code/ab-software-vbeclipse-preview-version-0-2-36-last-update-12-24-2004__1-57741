VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IListReader
Implements IListWriter

Dim m_ListKeys() As Variant
Dim m_ListValues() As Variant

Private Function IListReader_Count() As Long
   IListReader_Count = IListWriter_Count
End Function

Private Function IListReader_IsEmpty() As Boolean
   IListReader_IsEmpty = IListWriter_IsEmpty
End Function

Private Function IListReader_Item(Key As Variant) As Variant
   If IsObject(IListWriter_Item(Key)) Then
      Set IListReader_Item = IListWriter_Item(Key)
   Else
      IListReader_Item = IListWriter_Item(Key)
   End If
End Function

Private Function IListReader_Items() As Variant
   IListReader_Items = IListWriter_Items
End Function

Private Sub IListWriter_Add(Key As Variant, Value As Variant)
   
   Dim i As Long
   Dim ListCount As Long
   
   ListCount = IListWriter_Count
   
   If ListCount > -1 Then
      For i = 0 To UBound(m_ListKeys)
         If m_ListKeys(i) Is Key Then
            If IsObject(Value) Then
               Set m_ListValues(i) = Value
            Else
               m_ListValues(i) = Value
            End If
            Exit Sub
         End If
      Next i
   End If
   
   ListCount = ListCount + 1
   
   ReDim Preserve m_ListKeys(ListCount) As Variant
   If IsObject(Key) Then
      Set m_ListKeys(ListCount) = Key
   Else
      m_ListKeys(ListCount) = Key
   End If
   
   ReDim Preserve m_ListValues(ListCount) As Variant
   If IsObject(Value) Then
      Set m_ListValues(ListCount) = Value
   Else
      m_ListValues(ListCount) = Value
   End If
   
End Sub

Private Sub IListWriter_Clear()
   Erase m_ListKeys
   Erase m_ListValues
End Sub

Private Function IListWriter_Count() As Long
   
   On Error GoTo ErrorHandle
   
   IListWriter_Count = UBound(m_ListKeys)
   
   Exit Function
   
ErrorHandle:
   
   IListWriter_Count = -1
   
End Function

Private Function IListWriter_IsEmpty() As Boolean
    IListWriter_IsEmpty = (IListWriter_Count = -1)
End Function

Private Function IListWriter_Item(Key As Variant) As Variant
   
   Dim i As Long
   Dim ListCount As Long
   
   ListCount = IListWriter_Count
   
   If ListCount > -1 Then
      For i = 0 To UBound(m_ListKeys)
         If m_ListKeys(i) = Key Then
            If IsObject(m_ListValues(i)) Then
               Set IListWriter_Item = m_ListValues(i)
            Else
               IListWriter_Item = m_ListValues(i)
            End If
            Exit Function
         End If
      Next i
   End If

End Function

Private Function IListWriter_Items() As Variant
   IListWriter_Items = m_ListValues
End Function

Private Sub IListWriter_Remove(Key As Variant)

   Dim i As Long
   Dim c As Long
   Dim ListCount As Long
   
   ListCount = IListWriter_Count
   
   If ListCount > -1 Then
      
      For i = 0 To UBound(m_ListKeys)
         If m_ListKeys(i) <> Key Then
            
            
            If IsObject(m_ListKeys(i)) Then
               Set m_ListKeys(c) = m_ListKeys(i)
            Else
               m_ListKeys(c) = m_ListKeys(i)
            End If
            
            If IsObject(m_ListValues(i)) Then
               Set m_ListValues(c) = m_ListValues(i)
            Else
               m_ListValues(c) = m_ListValues(i)
            End If
            
            c = c + 1
         End If
      Next i
      
      ReDim Preserve m_ListKeys(ListCount) As Variant
      ReDim Preserve m_ListValues(ListCount) As Variant
   End If

End Sub

Public Function Reader() As IListReader
   Set Reader = Me
End Function

Public Function Writer() As IListWriter
   Set Writer = Me
End Function
