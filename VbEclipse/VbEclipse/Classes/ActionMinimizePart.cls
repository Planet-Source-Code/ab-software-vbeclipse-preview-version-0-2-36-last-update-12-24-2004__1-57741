VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionMinimizePart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAction

Private Property Get IAction_Enabled() As Boolean

End Property

Private Property Let IAction_Enabled(ByVal NewEnabled As Boolean)

End Property

Private Function IAction_Run(ParamArray Args() As Variant) As Boolean

End Function

Private Property Let IAction_Text(ByVal NewText As String)

End Property

Private Property Get IAction_Text() As String
   
End Property
