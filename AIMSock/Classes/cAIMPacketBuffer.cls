VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMPacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Buffer As String

Public Sub InsertDWORD(ByVal Data As Long)
    Buffer = Buffer & MakeDWORD(Data)
End Sub

Public Sub InsertWORD(ByVal Data As Long)
    Buffer = Buffer & MakeWORD(Data)
End Sub

Public Sub InsertBYTE(ByVal Data As Byte)
    Buffer = Buffer & Chr$(Data)
End Sub

Public Sub InsertByteString(ByVal Data As String)
    Buffer = Buffer & Chr$(Len(Data)) & Data
End Sub

Public Sub InsertString(ByVal Data As String)
    Buffer = Buffer & MakeWORD(Len(Data)) & Data
End Sub

Public Sub InsertNonNTString(ByVal strData As String)
    Buffer = Buffer & strData
End Sub

Public Sub InsertTLV(ByVal lngType As Long, ByVal strData As String)
    Buffer = Buffer & MakeWORD(lngType) & MakeWORD(Len(strData)) & strData
End Sub

Public Sub InsertItem(ByVal Name As String, ByVal GroupID As Long, ByVal BuddyID As Long, ByVal TheType As Long, ByVal TLVs As String)
    InsertString Name
    InsertWORD GroupID
    InsertWORD BuddyID
    InsertWORD TheType
    InsertString TLVs
End Sub

Public Sub InsertSNAC(ByVal FamilyID As Families, ByVal SubTypeID As Long, ByVal Flags1 As Byte, ByVal Flags2 As Byte, ByVal requestID As Long)
    Buffer = Buffer & MakeWORD(FamilyID) & MakeWORD(SubTypeID) & Chr(Flags1) & Chr(Flags2) & MakeDWORD(requestID)
End Sub

Public Function GetBuffer(ByVal ChannelID As Channels) As String
    GetBuffer = Chr(&H2A) & Chr(ChannelID) & MakeWORD(0) & MakeWORD(Len(Buffer)) & Buffer
    Clear
End Function

Public Function Clear()
    Buffer = ""
End Function

Public Function SendPacket(ByVal ChannelID As Channels)
    'PrivateSocket.SendData Chr(&H2A) & Chr(ChannelID) & MakeWORD(SeqNum) & MakeWORD(Len(Buffer)) & Buffer
    Debug.Print "Send " & vbCrLf & DebugOutput(Chr(&H2A) & Chr(ChannelID) & MakeWORD(0) & MakeWORD(Len(Buffer)) & Buffer) & vbCrLf
    Clear
End Function

