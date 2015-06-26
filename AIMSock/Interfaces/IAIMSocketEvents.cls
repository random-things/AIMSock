VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAIMSocketEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub OnConnected(ByVal Index As Long, ByVal Server As String, ByVal Port As Long)
    '
End Sub

Public Sub OnConnecting(ByVal Index As Long, ByVal Server As String, ByVal Port As Long)
    '
End Sub

Public Sub OnDisconnected(ByVal Index As Long, ByVal Forced As Boolean)
    '
End Sub

Public Sub OnPacketReceived(ByVal Index As Long, ByVal strPacket As String)
    '
End Sub

Public Sub OnPacketSent(ByVal Index As Long)
    '
End Sub

Public Sub OnError(ByVal Index As Long, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    '
End Sub
