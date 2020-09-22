Attribute VB_Name = "modPublic"
Option Explicit

' Registrierungsschlüssel-Stammtypen...
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS As Long = &H80000003
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_CURRENT_CONFIG As Long = &H80000005

Private Const KEY_READ = &H20019 'Lese zugriff
Private Const REG_SZ = 1 'Ein VBNullChar-Zeichen Terminierter String

Public Const REL_FOLDER As Long = -1
Public Const REL_WINDOW As Long = -2

Private Const WS_EX_TOOLWINDOW As Long = &H80&

Private Const WS_EX_MDICHILD As Long = &H40&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_ACTIVECAPTION As Long = &H1
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_SYSMENU As Long = &H80000

Private Const WS_EX_PALETTEWINDOW As Long = &H188

Private Const WS_EX_ACCEPTFILES As Long = &H10&
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_CONTEXTHELP As Long = &H400&
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const WS_EX_LEFT As Long = &H0&
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000&
Private Const WS_EX_LTRREADING As Long = &H0&
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Private Const WS_EX_RIGHTSCROLLBAR As Long = &H0&
Private Const WS_EX_RTLREADING As Long = &H2000&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_EX_RIGHT As Long = &H1000&

Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)

Private Declare Function SetParent Lib "user32.dll" ( _
   ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long _
) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" _
   Alias "CreateDCA" ( _
   ByVal lpDriverName As String, _
   lpDeviceName As Any, _
   lpOutput As Any, _
   lpInitData As Any _
) As Long

Private Declare Function GetWindowText Lib "user32.dll" _
   Alias "GetWindowTextA" ( _
   ByVal hWnd As Long, _
   ByVal lpString As String, _
   ByVal cch As Long _
) As Long

Private Declare Function SetWindowText Lib "user32.dll" _
   Alias "SetWindowTextA" ( _
   ByVal hWnd As Long, _
   ByVal lpString As String _
) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
   ByVal hDc As Long _
) As Long

Private Declare Function DrawFocusRect Lib "user32" ( _
   ByVal hDc As Long, _
   lpRect As RECT _
) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
 ) As Long
 
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
   ByVal hWnd As Long, _
   ByVal hRgn As Long, _
   ByVal bRedraw As Boolean _
) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
   ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal X2 As Long, _
   ByVal Y2 As Long, _
   ByVal X3 As Long, _
   ByVal Y3 As Long _
) As Long

Private Declare Function GetWindowRect Lib "user32" ( _
   ByVal hWnd As Long, _
   lpRect As RECT _
) As Long

Private Declare Function MoveWindow Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long _
) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" _
   Alias "RegOpenKeyExA" ( _
   ByVal HKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   ByRef phkResult As Long _
) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" _
   Alias "RegQueryValueExA" ( _
   ByVal HKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   ByRef lpType As Long, _
   ByVal lpData As String, _
   ByRef lpcbData As Long _
) As Long

Private Declare Function RegCloseKey Lib "advapi32" ( _
   ByVal HKey As Long _
) As Long

Private Declare Function LockWindowUpdate Lib "user32.dll" ( _
   ByVal hWndLock As Long _
) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum VbHKey
   VbHKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
   VbHKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
   VbHKEY_USERS = HKEY_USERS
   VbHKEY_CURRENT_USER = HKEY_CURRENT_USER
   VbHKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
End Enum

Public Enum VbWindowsScheme
   VbClassic = 0
   VbNormalColor = 1
   VbMetallic = 2
   VbHomeStead = 3
End Enum

Public lngLockedWindowHwnd As Long

Public Function LockWindow(ByVal hWnd As Long) As Long
   If lngLockedWindowHwnd = 0 Then
      lngLockedWindowHwnd = hWnd
      LockWindow = LockWindowUpdate(hWnd)
   End If
End Function
Public Function UnLockWindow(ByVal hWnd As Long) As Long
   If lngLockedWindowHwnd = hWnd Then
      UnLockWindow = LockWindowUpdate(0)
      lngLockedWindowHwnd = 0
   End If
End Function

Public Function GetKeyValue(ByVal MainKey As VbHKey, ByVal SubKey As String, ByVal Value As String)
   
   Dim RetVal As Long
   Dim HKey As Long
   Dim TmpSNum As String * 255
   
   RetVal = RegOpenKeyEx(MainKey, SubKey, 0&, KEY_READ, HKey)
   
   If RetVal <> 0 Then
      GetKeyValue = "Can't open the registry."
      Exit Function
   End If
   
   RetVal = RegQueryValueEx(HKey, Value, 0, REG_SZ, ByVal TmpSNum, Len(TmpSNum))
   
   If RetVal <> 0 Then
      GetKeyValue = "Can't read or find the registry."
      Exit Function
   End If
   
   GetKeyValue = Left$(TmpSNum, InStr(1, TmpSNum, vbNullChar) - 1)
   
   RetVal = RegCloseKey(HKey)
   
End Function

'Global Const ID_EDITOR_AREA As String = "abs.dockpanel.editorarea"

'Public DragDropMode As Boolean

Public Sub SetWindowStyle(ByVal hWnd As Long, ByVal ToolWindow As Boolean)
   
'   Dim toolWin As frmToolWindow
   Dim Style As Long
             
   If ToolWindow Then
      
'      'strCaption = GetWindowText(hwnd, strBuffer, Len(strBuffer))
'
'      Set toolWin = New frmToolWindow
'
'      ' Set the caption...
'      SetWindowText hwnd, "strCaption"
'      SetParent hwnd, toolWin.hwnd
'      ' Resize the new "FORM" so the "FORM" will display properly...
'      GetWindowRect hwnd, l_Rect
'      'MoveWindow hwnd, l_Rect.Left, l_Rect.Top, l_Rect.Right - l_Rect.Left + 30, l_Rect.Bottom - l_Rect.Top + 30, True
'      toolWin.Show
   Else
      ' get the current window style
      Style = GetWindowLong(hWnd, GWL_STYLE)
      ' set new window style
      Style = Style And Not WS_DLGFRAME And Not WS_EX_APPWINDOW And Not WS_BORDER And Not WS_EX_WINDOWEDGE Or WS_EX_MDICHILD Or WS_CHILDWINDOW And Not WS_EX_NOPARENTNOTIFY
   End If
   
   SetWindowLong hWnd, GWL_STYLE, Style 'WS_EX_MDICHILD And Not WS_EX_WINDOWEDGE And Not WS_CAPTION

End Sub

Public Sub DrawDragRect(Rc As RECT, Optional ByVal Size As Long = 2)
        
   Dim DrawRect As RECT
   Dim hDc As Long
   Dim i As Long
   Dim Width As Long
        
   hDc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

   For i = 0 To Size
           
      With DrawRect
         .Top = Rc.Top + i
         .Bottom = Rc.Bottom - i
         .Left = Rc.Left + i
         .Right = Rc.Right - i
      End With
           
      DrawFocusRect hDc, DrawRect
           
   Next i
        
   DeleteDC hDc
        
End Sub

' Returns the style of the selected theme scheme.
'
' @GetSchemeStyle Display name of the selected theme scheme.
Public Function Scheme() As VbWindowsScheme
   
   On Error GoTo Error_Handle
   
   Dim SchemeName As String
   Dim RegKeyTheme As String
   
   RegKeyTheme = "Software\Microsoft\Windows\CurrentVersion\ThemeManager"
   
   SchemeName = GetKeyValue(VbHKEY_CURRENT_USER, RegKeyTheme, "ColorName")
   
   Select Case SchemeName
      Case "NormalColor":    Scheme = VbNormalColor
      Case "HomeStead":      Scheme = VbHomeStead
      Case "Metallic":       Scheme = VbMetallic
      Case Else:             Scheme = VbClassic
   End Select
   'Scheme = VbHomeStead
Finally:
      
   Exit Function
   
Error_Handle:

   Scheme = "classic"
   
   GoTo Finally

End Function
