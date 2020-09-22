VERSION 5.00
Begin VB.UserControl vbeTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   ScaleHeight     =   465
   ScaleWidth      =   1410
   ToolboxBitmap   =   "vbeTab.ctx":0000
   Begin VB.Image tabImage 
      Height          =   240
      Left            =   60
      Picture         =   "vbeTab.ctx":0312
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label tabCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "vbeTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum vbTabState
   STATE_INACTIVE = 0
   STATE_ACTIVE = 1
   STATE_ACTIVE_HOT = 2
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   ByRef lpRect As RECT _
) As Long

Private Declare Function DrawFocusRect Lib "user32.dll" ( _
   ByVal hDc As Long, _
   ByRef lpRect As RECT _
) As Long

Private m_Key As String
Private m_State As vbTabState
Private m_Theme As ITheme

Private Sub UserControl_Show()
   Refresh
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   
   tabCaption.Top = (ScaleHeight - tabCaption.Height) * 0.5
   tabImage.Top = (ScaleHeight - tabImage.Height) * 0.5
   
End Sub

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property
Public Property Get HasDC() As Boolean
   HasDC = UserControl.HasDC
End Property
Public Property Get hDc() As Long
   hDc = UserControl.hDc
End Property
Public Property Get State() As vbTabState
   State = m_State
End Property
Public Property Let State(ByVal NewState As vbTabState)
   m_State = NewState
End Property

Public Property Get Image() As Picture
   Set Image = tabImage.Picture
End Property
Public Property Set Image(ByVal NewImage As Picture)
   Set tabImage.Picture = NewImage
End Property

Public Property Get Key() As String
   Key = m_Key
End Property
Public Property Let Key(ByVal NewKey As String)
   m_Key = NewKey
End Property

Public Property Get Caption() As String
   Caption = tabCaption.Caption
End Property
Public Property Let Caption(ByVal NewCaption As String)
   tabCaption.Caption = NewCaption
   tabCaption.ToolTipText = NewCaption
   tabImage.ToolTipText = NewCaption
End Property

Public Property Get Theme() As ITheme
   Set Theme = Parent.Theme
End Property

Public Sub Refresh()
   
   Dim Gradient As Gradient
   Set Gradient = New Gradient
   
   UserControl.AutoRedraw = True
   
   tabImage.Visible = (State > STATE_INACTIVE)
   
   Select Case State
      
      Case STATE_INACTIVE:
      
           tabCaption.ForeColor = Theme.InactiveForeColor
           tabCaption.Left = 120
      
      Case STATE_ACTIVE:
      
           tabCaption.ForeColor = Theme.ActiveForeColor
           tabCaption.Left = 360
      
      Case STATE_ACTIVE_HOT:
           
           tabCaption.ForeColor = Theme.ActiveCaptionForeColor
           tabCaption.Left = 360
      
   End Select
   
   With tabCaption
      UserControl.Width = .Left + .Width + 120
   End With
   
   Select Case State
      
      Case STATE_INACTIVE:
           
           'tabCaption.ForeColor = Theme.InactiveCaptionForeColor
           
'           With Gradient
'              If Parent.Active Then
'                 .Color1 = Theme.ActiveBackColor1
'                 .Color2 = Theme.ActiveBackColor2
'              Else
'                 .Color1 = Theme.InactiveBackColor1
'                 .Color2 = Theme.InactiveBackColor2
'              End If
'              .Angle = Theme.ActiveBackColorGradientAngle
'              .Draw Me
'           End With
           UserControl.Cls
           UserControl.Line (1, ScaleHeight - 10)-(ScaleWidth, ScaleHeight - 10), Theme.InactiveFrameColor
           UserControl.Line (ScaleWidth - 10, 30)-(ScaleWidth - 10, ScaleHeight - 40), Theme.InactiveFrameColor
           
      Case STATE_ACTIVE:
           
           'tabCaption.ForeColor = Theme.ActiveCaptionForeColor

           With Gradient
              .Color1 = Theme.InactiveCaptionBackColor1
              .Color2 = Theme.InactiveCaptionBackColor2
              .Angle = Theme.InactiveCaptionBackGradientAngle
              .Draw Me
           End With
           
           UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight), Theme.InactiveFrameColor, B
            
      Case STATE_ACTIVE_HOT:
           
           'tabCaption.ForeColor = Theme.ActiveCaptionForeColor
           
           With Gradient
              .Color1 = Theme.ActiveCaptionBackColor1
              .Color2 = Theme.ActiveCaptionBackColor2
              .Angle = Theme.ActiveCaptionBackGradientAngle
              .Draw Me
           End With
           
           UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight), Theme.InactiveFrameColor, B
           
           UserControl.AutoRedraw = False
           
           Dim R As RECT
           GetWindowRect Me.hWnd, R
           R.Left = R.Left + 20
           R.Right = R.Right - 20
           R.Top = R.Top + 20
           R.Bottom = R.Bottom - 20

          'DrawFocusRect Me.hDc, R
   End Select
   
   UserControl.AutoRedraw = False
   
   Set Gradient = Nothing
   
End Sub

Private Sub tabImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub tabCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Parent.EventRaise "MouseDown", Me.Key, Button
End Sub

Private Sub tabImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   On Error Resume Next
   
   UserControl_MouseUp Button, Shift, x, y
   Parent.EventRaise "ImageClick", Me.Key, Button
   
End Sub
Private Sub tabCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Parent.EventRaise "MouseUp", Me.Key, Button
End Sub

Private Sub tabCaption_DblClick()
   UserControl_DblClick
End Sub
Private Sub tabImage_DblClick()
   UserControl_DblClick
End Sub
Private Sub UserControl_DblClick()
   Parent.EventRaise "TabDblClick", Me.Key
End Sub

