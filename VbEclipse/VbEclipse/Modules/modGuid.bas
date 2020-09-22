Attribute VB_Name = "modGuid"
Option Explicit

Private Declare Function UuidCreate Lib "rpcrt4" ( _
   lpGUID As GUID _
) As Long

Private Declare Function UuidToString Lib "rpcrt4" _
   Alias "UuidToStringA" ( _
   lpGUID As GUID, _
   lpGUIDString As Long _
) As Long

Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" ( _
   ByVal lpString As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" ( _
   lpDest As Any, _
   lpSource As Any, _
   ByVal cBytes As Long _
)

Private Declare Function RpcStringFree Lib "rpcrt4" _
   Alias "RpcStringFreeA" ( _
   lpGUIDString As Long _
) As Long

Const RPC_S_OK As Long = &H0
Const RPC_S_UUID_LOCAL_ONLY As Long = &H720
Const RPC_S_UUID_NO_ADDRESS As Long = &H6CB

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Function CreateGUID() As String

   Dim G As GUID
   Dim GuidByt As Long
   Dim l As Long
   Dim GuidStr As String
   Dim Buffer() As Byte

   If UuidCreate(G) <> RPC_S_UUID_NO_ADDRESS Then

      If UuidToString(G, GuidByt) = RPC_S_OK Then

         l = lstrlen(GuidByt)
         ReDim Buffer(l - 1) As Byte

         Call CopyMemory(Buffer(0), ByVal GuidByt, l)
         Call RpcStringFree(GuidByt)

         GuidStr = StrConv(Buffer, vbUnicode)
         CreateGUID = UCase$(GuidStr)

      End If

   End If

End Function
