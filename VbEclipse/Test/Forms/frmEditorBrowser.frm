VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmEditorBrowser 
   BackColor       =   &H00C0E0FF&
   Caption         =   "C:\Test.txt"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   Icon            =   "frmEditorBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   960
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmEditorBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    
Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Sub Form_Load()
   WebBrowser1.Navigate "http://www.pscode.com"
End Sub

Private Sub Form_Resize()
   WebBrowser1.Move -30, -30, ScaleWidth + 60, ScaleHeight + 60
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   frmViewHistory.lvwHistory.ListItems.Add , , URL
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
   frmViewConsole.txtConsole = frmViewConsole.txtConsole & vbNewLine & Now & vbTab & Text
   frmViewConsole.txtConsole.SelStart = Len(frmViewConsole.txtConsole)
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    With WebBrowser1
        While Not ( _
            .ReadyState = READYSTATE_INTERACTIVE _
            Or .ReadyState = READYSTATE_COMPLETE _
            Or .Busy = True)
            DoEvents
        Wend
    End With
    
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

'    Timer1.Enabled = False
'    With WebBrowser1
'            If .ReadyState = READYSTATE_INTERACTIVE Then Exit Sub
'            If Not .ReadyState = READYSTATE_COMPLETE Then Exit Sub
'            If .Busy = True Then Exit Sub
'    End With

    
    ' Set ratio
'    Dim Ratio As Double
'    Ratio = frmMain.picScreenShot.Width / WebBrowser1.Width
'    frmMain.picScreenShot.Height = WebBrowser1.Height * Ratio
    
    ' Resize and paste to main form
    StretchBlt _
        frmViewPreview.picScreenShot.hdc, _
        0, _
        0, _
        frmViewPreview.picScreenShot.Width, _
        frmViewPreview.picScreenShot.Height, _
        GetDC(WebBrowser1.Parent.hwnd), _
        0, _
        0, _
        WebBrowser1.Width, _
        WebBrowser1.Height, _
        SRCCOPY
    
    ' Save screenshot
    
    'SavePicture frmMain.picScreenShot.Image, frmMain.txtSaveAs.Text
    
    'WebBrowser1.Navigate2 "about:blank"
    
    'Unload Me

End Sub


