VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl vbeTabStrip 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   ScaleHeight     =   375
   ScaleWidth      =   10185
   ToolboxBitmap   =   "vbeTabStrip.ctx":0000
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbeTabStrip.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbeTabStrip.ctx":046C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   9000
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            ImageIndex      =   2
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "vbeTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ActiveTabKey As String
Private m_Tabs As HashTable
Private m_Theme As ITheme
Private m_Active As Boolean

Public Event TabMouseDown(ByVal TabKey As String, ByVal Button As Single)
Public Event TabMouseUp(ByVal TabKey As String, ByVal Button As Single)
Public Event TabDblClick(ByVal TabKey As String)
Public Event TabImageClick(ByVal TabKey As String, ByVal Button As Single)
Public Event TabClose(ByVal TabKey As String)

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   
   Select Case Button.Key
      Case "close":   EventRaise "Close", ActiveTabKey
   End Select
   
End Sub

Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
   Dim i As Long
   Dim Views As Variant
   Dim View As View
   
   Select Case Button.Key
      Case "close":
      
           Button.ButtonMenus.Clear
                      
           Views = Parent.Views.Items
           
           For i = 0 To UBound(Views)
              Set View = Views(i)
              Button.ButtonMenus.Add , View.ViewId, View.Caption
              Debug.Print View.ViewId
           Next i
                      
   End Select
   
   Set View = Nothing

End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   EventRaise "MouseDown", ButtonMenu.Key, vbLeftButton
   EventRaise "MouseUp", ButtonMenu.Key, vbLeftButton
   
   Debug.Print ButtonMenu.Key
End Sub

Private Sub UserControl_Initialize()
   Set m_Tabs = New HashTable
   UserControl.BackColor = Theme.InactiveFrameColor
End Sub

Private Sub UserControl_Show()
   Toolbar.Visible = True
End Sub

Private Sub UserControl_Terminate()
   Set m_Tabs = Nothing
End Sub

Public Sub EventRaise(ByVal strEvent As String, ParamArray Args() As Variant)

   ' Save key of active tab
   ActiveTabKey = Args(0)
   
   Select Case strEvent
      
      Case "MouseDown":
           
           RaiseEvent TabMouseDown(Args(0), Args(1))
           
      Case "MouseUp":
           
           RaiseEvent TabMouseUp(Args(0), Args(1))
           
      Case "TabDblClick":
           
           RaiseEvent TabDblClick(Args(0))
           
      Case "ImageClick":
           
           RaiseEvent TabImageClick(Args(0), Args(1))
           
      Case "Close":
           
           RaiseEvent TabClose(ActiveTabKey)
           
   End Select

   Refresh

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

Public Property Get ActiveTabKey() As String
      
   Dim t As Variant
   
   ActiveTabKey = m_ActiveTabKey
      
   If Len(ActiveTabKey) = 0 Then
      If Not IsEmpty(Tabs) Then
         For Each t In Tabs
            ActiveTabKey = t.Key
            Exit For
         Next
      End If
   End If
     
End Property
Public Property Let ActiveTabKey(ByVal Key As String)
   m_ActiveTabKey = Key
End Property

Public Property Get Active() As Boolean
   Active = m_Active
End Property
Public Property Let Active(ByVal NewActive As Boolean)
   
   If m_Active <> NewActive Then
      m_Active = NewActive
      Refresh
   End If
   
End Property

Public Property Get Theme() As ITheme
   
   If m_Theme Is Nothing Then
      Set m_Theme = New ThemeOffice2003
   End If
   
   Set Theme = m_Theme
   
End Property
Public Property Let Theme(ByVal NewTheme As ITheme)
   Set m_Theme = Parent.Theme
End Property

Private Sub UserControl_Resize()
   
   On Error Resume Next
   
   picBackground.Move 10, 10, ScaleWidth - 30, ScaleHeight - 30
   
   Dim c As Long
   Dim t As Variant
   
   t = Tabs
   
   If Not IsEmpty(t) Then
      c = UBound(t)
      If t(c).Left + t(c).Width > ScaleWidth - Toolbar.Width Then
         Toolbar.Buttons(1).Style = tbrDropdown
      Else
         Toolbar.Buttons(1).Style = tbrDefault
      End If
      
      Toolbar.Width = Toolbar.Buttons(1).Width
   End If
   
   Toolbar.Move ScaleWidth - Toolbar.Width - 10, 10, Toolbar.Width, ScaleHeight - 30
   
End Sub

Public Property Get Tabs() As Variant
   
   Tabs = m_Tabs.Values
   
End Property

Public Sub Add(ByVal Key As String, ByVal Caption As String, Optional ByVal Icon As Picture)
   
   Dim NewTab As Object
   
   If Not m_Tabs.Item(Key) Is Nothing Then
      Err.Raise 1000, , "Key is not unique!"
   End If
   
   Set NewTab = Controls.Add("absVbEclipse.vbeTab", "Tab_" & Key)
   
   With NewTab
       .Caption = Caption
       .Key = Key
       If IsEmpty(Tabs) Then
          .State = STATE_ACTIVE
       Else
          .State = STATE_INACTIVE
       End If
       .Visible = True
       .Image = Icon
       '.Refresh
       .ZOrder
   End With
   
   Toolbar.ZOrder
   
   m_Tabs.Add Key, NewTab
   
   Refresh
   
End Sub

Public Sub Remove(ByVal Key As String)
   
   Dim NewTab As Object
   
   If Not m_Tabs.Item(Key) Is Nothing Then
   
      Set NewTab = m_Tabs.Item(Key)
   
      Controls.Remove NewTab
   
      m_Tabs.Remove Key
   End If
   
   Refresh
   
End Sub

Public Sub Show(ByVal Key As String)
   
   If m_Tabs.Item(Key) Is Nothing Then
      Err.Raise -1, "", "No tab available for key '" & Key & "'!"
   End If
   
   If ActiveTabKey <> Key Then
      ActiveTabKey = Key

      Refresh
   End If
   
End Sub

Public Sub Refresh()
   
   Dim l As Long
   Dim t As Variant
   
   l = 60
   
   LockWindow UserControl.hWnd
   
  ' UserControl.AutoRedraw = False
   If Not IsEmpty(Tabs) Then
   
      For Each t In Tabs
         With t
            If StrComp(.Key, ActiveTabKey, vbBinaryCompare) = 0 Then
               If Active Then
                  .State = STATE_ACTIVE_HOT
               Else
                  .State = STATE_ACTIVE
               End If
            Else
               .State = STATE_INACTIVE
            End If
            
            .Move l, 40, .Width, ScaleHeight - 40
            .Refresh
            l = l + .Width
         End With
      Next
      
   End If
   
'   picBackground.Width = Screen.Width
'
'   Dim Gradient As Gradient
'   Set Gradient = New Gradient
'
'   With Gradient
'      If Me.Active Then
'         .Color1 = Theme.ActiveBackColor1
'         .Color2 = Theme.ActiveBackColor2
'         .Angle = Theme.ActiveBackColorGradientAngle
'      Else
'         .Color1 = Theme.InactiveBackColor1
'         .Color2 = Theme.InactiveBackColor2
'         .Angle = Theme.InactiveBackColorGradientAngle
'      End If
'
'      picBackground.AutoRedraw = True
'      If .Color1 <> .Color2 Then
'         .Draw picBackground
'      Else
'         picBackground.Cls
'         picBackground.BackColor = .Color1
'      End If
'      picBackground.AutoRedraw = False
'
'   End With

   UserControl_Resize

   UnLockWindow UserControl.hWnd
 
'   Set Gradient = Nothing
   
End Sub
