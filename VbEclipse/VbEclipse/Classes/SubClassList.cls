VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubClassList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WinHwnds() As Long
Private WinProcs() As Long

Public Sub Add(ByVal hWnd As Long, ByVal Proc As Long)
   
   Dim Idx As Long
   Idx = Count + 1
   
   ReDim Preserve WinHwnds(Idx) As Long
   ReDim Preserve WinProcs(Idx) As Long
   
   WinHwnds(Idx) = hWnd
   WinProcs(Idx) = Proc
   
End Sub

Public Function Item(ByVal hWnd As Long) As Long
   
   Dim i As Long
   
   For i = 0 To Count
      If WinHwnds(i) = hWnd Then
         Item = WinProcs(i)
         Exit Function
      End If
   Next i
      
End Function

Public Sub Remove(ByVal hWnd As Long)
   
   Dim c As Long
   Dim i As Long
   
   If Count = 0 Then
      Erase WinHwnds
      Erase WinProcs
   Else
      For i = 0 To Count
         If WinHwnds(i) <> hWnd Then
            WinHwnds(i) = WinHwnds(c)
            WinProcs(i) = WinProcs(c)
            c = c + 1
         End If
      Next i
      
      ReDim Preserve WinProcs(Count - 1) As Long
      ReDim Preserve WinHwnds(Count - 1) As Long
   End If
   
End Sub

Public Function Count() As Long
   
   On Error GoTo ErrorHandle
   
   Count = UBound(WinHwnds)
    
   Exit Function
    
ErrorHandle:

   Count = -1
   
End Function
