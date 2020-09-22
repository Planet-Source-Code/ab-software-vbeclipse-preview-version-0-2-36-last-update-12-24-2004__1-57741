Attribute VB_Name = "modSubClass"
Option Explicit

Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long _
) As Long

Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" ( _
   ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, _
   ByVal msg As Long, _
   ByVal Wparam As Long, _
   ByVal Lparam As Long _
) As Long

Public Const GWL_WNDPROC = (-4)

Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_ACTIVATE = &H6
Private Const WM_SETFOCUS = &H7
Private Const WM_ENABLE As Long = &HA

Dim m_SubClass As SubClass
Public m_SubClassList As New SubClassList

Public Property Set SetSubClass(Instance As SubClass)
   Set m_SubClass = Instance
End Property

Public Sub Hook(lngHwnd)
   Dim PrevProc As Long
   
   If m_SubClassList.Item(PrevProc) = 0 Then
      PrevProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WindowProc)
   
      m_SubClassList.Add lngHwnd, PrevProc
   End If
   
End Sub
Public Sub UnHook(lngHwnd)
    
    Dim PrevProc As Long
    
    PrevProc = m_SubClassList.Item(lngHwnd)
    
    SetWindowLong lngHwnd, GWL_WNDPROC, PrevProc
    
    m_SubClassList.Remove lngHwnd
End Sub
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long
    
    Dim PrevProc As Long
    
    PrevProc = m_SubClassList.Item(hWnd)
    
    WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, Wparam, Lparam)
    
    Select Case uMsg
       Case 4110, 513, 33
            
            If Not m_SubClass Is Nothing Then
               m_SubClass.ShowView hWnd
            End If
            
    End Select
    
       
End Function


