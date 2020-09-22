VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThemeOffice2003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ITheme

Private Function ITheme_ActiveCaptionBackGradientAngle() As Long
   
   Select Case Scheme
      Case VbMetallic, VbNormalColor, VbHomeStead:
           ITheme_ActiveCaptionBackGradientAngle = 270
      Case Else:
           ITheme_ActiveCaptionBackGradientAngle = 180
   End Select

End Function

Private Function ITheme_ActiveCaptionForeColor() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_ActiveCaptionForeColor = RGB(0, 0, 0)
      Case Else:
           ITheme_ActiveCaptionForeColor = vbWhite
   End Select
   
End Function

Private Function ITheme_ActiveCaptionBackColor1() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_ActiveCaptionBackColor1 = RGB(252, 240, 204)
      Case Else:
           ITheme_ActiveCaptionBackColor1 = RGB(166, 202, 240)
   End Select
   
End Function

Private Function ITheme_ActiveCaptionBackColor2() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_ActiveCaptionBackColor2 = RGB(241, 168, 55)
      Case Else:
           ITheme_ActiveCaptionBackColor2 = RGB(10, 36, 106)
   End Select
   
End Function

Private Function ITheme_ActiveForeColor() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_ActiveForeColor = vbWhite
      Case VbMetallic:
           ITheme_ActiveForeColor = vbBlack
      Case VbHomeStead:
           ITheme_ActiveForeColor = vbBlack
      Case Else:
           ITheme_ActiveForeColor = vbWhite
   End Select
   
End Function

Private Function ITheme_ActiveFrameColor() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_ActiveFrameColor = RGB(241, 168, 55)
      Case Else:
           ITheme_ActiveFrameColor = RGB(81, 112, 167)
   End Select
   
End Function

Private Function ITheme_InactiveForeColor() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_InactiveForeColor = RGB(100, 100, 100)
      Case Else:
           ITheme_InactiveForeColor = vbBlack
   End Select
End Function

Private Function ITheme_InactiveFrameColor() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_InactiveFrameColor = RGB(8, 49, 217)
      Case VbMetallic:
           ITheme_InactiveFrameColor = RGB(119, 118, 151)
      Case VbHomeStead:
           ITheme_InactiveFrameColor = RGB(100, 144, 88)
      Case Else:
           ITheme_InactiveFrameColor = RGB(128, 128, 128)
   End Select
   
End Function

Private Function ITheme_ActiveBackColor1() As Long
   Select Case Scheme
      Case VbNormalColor:
           ITheme_ActiveBackColor1 = RGB(227, 239, 252)
      Case VbMetallic:
           ITheme_ActiveBackColor1 = RGB(155, 153, 182)
      Case VbHomeStead:
           ITheme_ActiveBackColor1 = RGB(242, 243, 232)
      Case Else:
           ITheme_ActiveBackColor1 = vbButtonFace 'ITheme_ActiveCaptionBackColor1
   End Select
End Function
Private Function ITheme_ActiveBackColor2() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_ActiveBackColor2 = RGB(135, 172, 227) 'RGB(157, 185, 235)
      Case VbMetallic:
           ITheme_ActiveBackColor2 = RGB(237, 237, 243)
      Case VbHomeStead:
           ITheme_ActiveBackColor2 = RGB(168, 182, 128)
      Case Else:
           ITheme_ActiveBackColor2 = vbButtonFace
   End Select
   
End Function

Private Function ITheme_ActiveBackColorGradientAngle() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_ActiveBackColorGradientAngle = 90
      Case Else:
           ITheme_ActiveBackColorGradientAngle = 90
   End Select
   
End Function

Private Function ITheme_InactiveBackColor1() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_InactiveBackColor1 = vbButtonFace
      Case VbMetallic:
           ITheme_InactiveBackColor1 = vbWhite
      Case VbHomeStead:
           ITheme_InactiveBackColor1 = RGB(237, 235, 211)
      Case Else:
           ITheme_InactiveBackColor1 = vbButtonFace
   End Select
   
End Function
Private Function ITheme_InactiveBackColor2() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_InactiveBackColor2 = RGB(227, 239, 252)
      Case VbMetallic:
           ITheme_InactiveBackColor2 = vbButtonFace
      Case VbHomeStead:
           ITheme_InactiveBackColor2 = RGB(203, 206, 182)
      Case Else:
           ITheme_InactiveBackColor2 = vbButtonFace
   End Select
   
End Function
Private Function ITheme_InactiveBackColorGradientAngle() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_InactiveBackColorGradientAngle = 90
      Case Else:
           ITheme_InactiveBackColorGradientAngle = 180
   End Select
   
End Function

Private Function ITheme_InactiveCaptionForeColor() As Long
   ITheme_InactiveCaptionForeColor = RGB(100, 100, 100)
End Function

Private Function ITheme_InactiveCaptionBackColor1() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_InactiveCaptionBackColor1 = RGB(197, 221, 250) 'RGB(237, 235, 211) 'vbWhite 'RGB(157, 185, 235)
      Case VbMetallic:
           ITheme_InactiveCaptionBackColor1 = RGB(237, 237, 243)
      Case VbHomeStead:
           ITheme_InactiveCaptionBackColor1 = RGB(234, 245, 201)
      Case Else:
           ITheme_InactiveCaptionBackColor1 = RGB(192, 192, 192) ' RGB(166, 202, 240)
   End Select
   
End Function

Private Function ITheme_InactiveCaptionBackColor2() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           ITheme_InactiveCaptionBackColor2 = RGB(25, 100, 230) 'RGB(8, 49, 217) 'RGB(121, 152, 223)
      Case VbMetallic:
           ITheme_InactiveCaptionBackColor2 = RGB(155, 153, 182)
      Case VbHomeStead:
           ITheme_InactiveCaptionBackColor2 = RGB(168, 182, 128)
      Case Else:
           ITheme_InactiveCaptionBackColor2 = RGB(128, 128, 128)
   End Select
   
End Function

Private Function ITheme_InactiveCaptionBackGradientAngle() As Long
   
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           ITheme_InactiveCaptionBackGradientAngle = 90
      Case Else:
           ITheme_InactiveCaptionBackGradientAngle = 180
   End Select
      
End Function
