VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMRecvBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strData As String

Public Sub AddData(Data As String)
    strData = strData & Data
End Sub

Public Function FullPacket() As Boolean
    Dim lngPacketLen As Long
    FullPacket = False
    If Len(strData) <> 0 And Left$(strData, 1) = Chr$(&H2A) Then
        lngPacketLen = GetWORD(Mid$(strData, 5, 2))
        If Len(strData) >= lngPacketLen Then: FullPacket = True
    End If
End Function

Public Function GetPacket() As String
    Dim lngPacketLen As Long
    If Left(strData, 1) = Chr(&H2A) Then
        lngPacketLen = GetWORD(Mid$(strData, 5, 2))
        GetPacket = Mid$(strData, 1, lngPacketLen + 6)
        strData = Mid$(strData, lngPacketLen + 7)
    End If
End Function

Public Sub ClearBuffer()
    strData = ""
End Sub
