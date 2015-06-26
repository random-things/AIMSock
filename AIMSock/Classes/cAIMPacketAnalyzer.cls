VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMPacketAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ePos As Long
Private Buffer As String

Public Property Get Channel() As Byte
    Channel = Asc(Mid$(Buffer, 2, 1))
End Property

Public Property Get ServerSequenceNumber() As Long
    ServerSequenceNumber = GetWORD(Mid(Buffer, 3, 2))
End Property

Public Property Get FamilyID() As Long
    FamilyID = GetWORD(Mid(Buffer, 7, 2))
End Property

Public Property Get SubTypeID() As Long
    SubTypeID = GetWORD(Mid(Buffer, 9, 2))
End Property

Public Property Get DataLength() As Long
    DataLength = GetWORD(Mid(Buffer, 5, 2))
End Property

Public Function ReadByte() As Byte
    ReadByte = Asc(Mid(Buffer, ePos, 1))
End Function

Public Function ExtractByte() As Byte
    ExtractByte = Asc(Mid(Buffer, ePos, 1))
    ePos = ePos + 1
End Function

Public Function ExtractBytes(ByVal NumberOfBytes As Integer) As String
    If NumberOfBytes < 0 Or NumberOfBytes > (Len(Buffer) - ePos + 1) Then NumberOfBytes = 0
    ExtractBytes = Mid(Buffer, ePos, NumberOfBytes)
    ePos = ePos + NumberOfBytes
End Function

Public Function ReadWord() As Long
    ReadWord = GetWORD(Mid(Buffer, ePos, 2))
End Function

Public Function ExtractWord() As Long
    ExtractWord = GetWORD(Mid(Buffer, ePos, 2))
    ePos = ePos + 2
End Function

Public Function ExtractDword() As Long
    ExtractDword = GetDWORD(Mid(Buffer, ePos, 4))
    ePos = ePos + 4
End Function

Public Function ExtractByteString() As String
    Dim byteStringLength As Byte
    byteStringLength = ExtractByte()
    ExtractByteString = Mid(Buffer, ePos, byteStringLength)
    ePos = ePos + byteStringLength
End Function

Public Function ExtractString() As String
    Dim lngStrLength As Long
    lngStrLength = ExtractWord
    ExtractString = Mid(Buffer, ePos, lngStrLength)
    ePos = ePos + lngStrLength
End Function

Public Function ExtractMessage() As String
    'Debug.Print DebugOutput(Right(Buffer, Len(Buffer) - ePos))
    Dim lngStrLength As Integer
    lngStrLength = ExtractWord()
    'Debug.Print lngStrLength
    If lngStrLength < 0 Then lngStrLength = 0
    ExtractMessage = Mid(Buffer, ePos + 4, lngStrLength)
    ePos = ePos + lngStrLength + 4
End Function

Public Function ExtractTLV(ByRef TheType As Long, ByRef Length As Long, ByRef Value As String) As Boolean
    TheType = ExtractWord()
    Length = ReadWord()
    Value = ExtractString()
End Function

Public Property Let Pos(Pos As Long)
    ePos = Pos
End Property
Public Property Get Pos() As Long
    Pos = ePos
End Property

Public Property Get Size() As Long
    Size = Len(Buffer)
End Property

Public Property Let PacketData(ByVal Data As String)
    Buffer = Data
    ePos = &H11
End Property
Public Property Get PacketData() As String
    PacketData = Buffer
End Property

Public Property Get Data() As String
    Data = Mid(Buffer, ePos, Len(Buffer) - ePos + 1)
End Property

Private Sub Class_Initialize()
    ePos = 1
End Sub

Private Function GetDWORD(Data As String) As Long
    Dim tmp As String
    tmp = ToHex(Data)
    Dim A As String, b As String, c As String, d As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    c = Mid(tmp, 5, 2)
    d = Mid(tmp, 7, 2)
    tmp = A & b & c & d
    GetDWORD = Val("&H" & tmp)
End Function

Private Function GetWORD(Data As String) As Long
    Dim tmp As String
    tmp = ToHex(Data)
    Dim A As String, b As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    tmp = A & b
    GetWORD = Val("&H" & tmp)
End Function

Private Function ToHex(Data As String) As String
    Dim i As Integer
    For i = 1 To Len(Data)
        ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
    Next i
End Function
