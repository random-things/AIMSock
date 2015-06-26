VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents sckAIMSocket As clsSocket
Private WithEvents sckAIMSocket As Winsock
Attribute sckAIMSocket.VB_VarHelpID = -1

Private AIMRecvBuffer As cAIMRecvBuffer

Private m_EventSink As IAIMSocketEvents
Private m_Index As Long
Private m_Tag As String
Private m_Purpose As String

Private SeqNum As Long

Public Event OnConnected(ByRef Index As Long, ByRef Server As String, ByRef Port As Long)
Public Event OnConnecting(ByRef Index As Long, ByRef Server As String, ByRef Port As Long)
Public Event OnDataArrival(ByRef Index As Long, ByRef strData As String)
Public Event OnDisconnected(ByRef Index As Long, ByRef ForcedDisconnect As Boolean)
Public Event OnSocketError(ByRef Index As Long, ByRef ErrorNumber As Long, ByRef ErrorDescription As String)

Public Property Let EventSink(ByVal AIMSocketEvents As IAIMSocketEvents)
    Set m_EventSink = AIMSocketEvents
End Property

Public Property Let Index(ByVal newIndex As Long)
    m_Index = newIndex
End Property

Public Property Get Index() As Long
    Index = m_Index
End Property

Public Property Let Tag(ByVal newTag As String)
    m_Tag = newTag
End Property

Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Purpose(ByVal newPurpose As String)
    m_Purpose = newPurpose
End Property

Public Property Get Purpose() As String
    Purpose = m_Purpose
End Property

Private Sub Class_Initialize()
    Set sckAIMSocket = New MSWinsockLib.Winsock
    Set AIMRecvBuffer = New cAIMRecvBuffer
End Sub

Public Sub Connect(ByVal Server As String, ByVal Port As Long)
    sckAIMSocket.Close
    Call sckAIMSocket.Connect(Server, Port)
    'RaiseEvent OnConnecting(m_Index, Server, Port)
    Call m_EventSink.OnConnecting(m_Index, Server, Port)
End Sub

Public Sub Disconnect()
    sckAIMSocket.Close
    Call m_EventSink.OnDisconnected(m_Index, True)
End Sub

Public Sub SendData(ByVal strData As String)
    If sckAIMSocket.State = sckConnected Then
        Debug.Print " -- Send -- "
        Debug.Print DebugOutput(strData)
        Mid$(strData, 3, 2) = MakeWORD(SeqNum)
        Call sckAIMSocket.SendData(strData)
        SeqNum = SeqNum + 1
        If SeqNum >= 65536 Then SeqNum = 0
    End If
End Sub

Private Sub sckAIMSocket_Close()
    Call m_EventSink.OnDisconnected(m_Index, False)
End Sub

Private Sub sckAIMSocket_Connect()
    Call m_EventSink.OnConnected(m_Index, sckAIMSocket.RemoteHost, sckAIMSocket.RemotePort)
End Sub

Private Sub sckAIMSocket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    Call sckAIMSocket.GetData(strData, vbString)
    'Debug.Print " -- " & bytesTotal & " bytes received --"
    'Debug.Print DebugOutput(strData)
    
    Call AIMRecvBuffer.AddData(strData)
    
    While AIMRecvBuffer.FullPacket = True
        'RaiseEvent OnDataArrival(m_Index, AIMRecvBuffer.GetPacket)
        Call m_EventSink.OnPacketReceived(m_Index, AIMRecvBuffer.GetPacket)
    Wend
End Sub

Private Sub sckAIMSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckAIMSocket.Close
    Call m_EventSink.OnDisconnected(m_Index, False)
    Call m_EventSink.OnError(m_Index, Number, Description)
End Sub

Private Sub sckAIMSocket_SendComplete()
    Call m_EventSink.OnPacketSent(m_Index)
End Sub
