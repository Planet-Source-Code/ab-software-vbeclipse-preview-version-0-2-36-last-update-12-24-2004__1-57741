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
   ITheme_ActiveCaptionBackGradientAngle = 270
End Function

Private Function ITheme_ActiveCaptionForeColor() As Long
   ITheme_ActiveCaptionForeColor = RGB(0, 0, 0)
End Function

Private Function ITheme_ActiveCaptionBackColor1() As Long
   ITheme_ActiveCaptionBackColor1 = RGB(252, 240, 204)
End Function

Private Function ITheme_ActiveCaptionBackColor2() As Long
   ITheme_ActiveCaptionBackColor2 = RGB(241, 168, 55)
End Function

Private Function ITheme_ActiveFrameColor() As Long
   ITheme_ActiveFrameColor = RGB(241, 168, 55) 'RGB(252, 240, 204) 'RGB(100, 144, 88)
End Function
Private Function ITheme_InactiveFrameColor() As Long
   ITheme_InactiveFrameColor = RGB(100, 144, 88)
End Function

Private Function ITheme_ActiveBackColor1() As Long
   ITheme_ActiveBackColor1 = RGB(242, 243, 232)  'RGB(184, 199, 146)
End Function
Private Function ITheme_ActiveBackColor2() As Long
   ITheme_ActiveBackColor2 = RGB(168, 182, 128)
End Function

Private Function ITheme_ActiveBackColorGradientAngle() As Long
   ITheme_ActiveBackColorGradientAngle = 90
End Function

Private Function ITheme_InactiveBackColor1() As Long
   ITheme_InactiveBackColor1 = RGB(237, 235, 211) 'RGB(242, 243, 232)
End Function
Private Function ITheme_InactiveBackColor2() As Long
   ITheme_InactiveBackColor2 = RGB(203, 206, 182) 'RGB(184, 199, 146)
End Function
Private Function ITheme_InactiveBackColorGradientAngle() As Long
   ITheme_InactiveBackColorGradientAngle = 90
End Function

Private Function ITheme_InactiveCaptionForeColor() As Long
   ITheme_InactiveCaptionForeColor = RGB(100, 100, 100)
End Function

Private Function ITheme_InactiveCaptionBackColor1() As Long
   ITheme_InactiveCaptionBackColor1 = RGB(234, 245, 201) 'RGB(237, 235, 211)
End Function

Private Function ITheme_InactiveCaptionBackColor2() As Long
   ITheme_InactiveCaptionBackColor2 = RGB(168, 182, 128) 'RGB(203, 206, 182)
End Function

Private Function ITheme_InactiveCaptionBackGradientAngle() As Long
   ITheme_InactiveCaptionBackGradientAngle = 90
End Function
