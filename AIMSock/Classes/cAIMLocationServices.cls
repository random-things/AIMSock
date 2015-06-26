VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMLocationServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum InfoType
    GeneralInfo = 1
    ShortUserInfo = 2
    AwayMessage = 3
    Capabilities = 4
End Enum

Public Function RequestLocationRights() As cAIMPacketBuffer
    Set RequestLocationRights = New cAIMPacketBuffer
    RequestLocationRights.InsertSNAC Information, LocationSIDs.CS_RequestLocationRights, 0, 0, &H20002
End Function

Public Function SetUserInformation(ByVal Profile As String, Optional ByVal AwayMessage As String) As cAIMPacketBuffer
    Set SetUserInformation = New cAIMPacketBuffer
    
    With SetUserInformation
        .InsertSNAC Information, LocationSIDs.CS_SetUserInformation, 0, 0, &H20004
        .InsertTLV &H1, "text/x-aolrtf; charset=" & Chr(34) & "us-ascii" & Chr(34)
        .InsertTLV &H2, Profile
        .InsertTLV &H3, "text/x-aolrtf; charset=" & Chr(34) & "us-ascii" & Chr(34)
        .InsertTLV &H4, AwayMessage
        .InsertTLV &H5, Chr(&H9) & Chr(&H46) & Chr(&H13) & Chr(&H46) & Chr(&H4C) & Chr(&H7F) & Chr(&H11) & Chr(&HD1) & Chr(&H82) & Chr(&H22) & Chr(&H44) & Chr(&H45) & Chr(&H53) & Chr(&H54) & Chr(&H0) & Chr(&H0)
    End With
End Function

Public Function RequestUserInformation(ByVal ScreenName As String, ByVal InformationType As InfoType) As cAIMPacketBuffer
    Set RequestUserInformation = New cAIMPacketBuffer
    
    With RequestUserInformation
        .InsertSNAC Information, LocationSIDs.CS_RequestUserInformation, 0, 0, &H20005
        .InsertWORD InformationType
        .InsertByteString ScreenName
    End With
End Function

Public Function InfoExtract(ByRef Packet As cAIMPacketAnalyzer, Optional ByRef bStatus As Boolean) As String
    If Not IsMissing(bStatus) Then bStatus = True
    
    With OutInfo
        .AwayMessage = ""
        .AwayMessageEncoding = ""
        .Capabilities = 0
        .CreatedTime = ""
        .Flags = 0
        .IdleTime = 0
        .MemberSince = ""
        .OnlineSince = ""
        .Profile = ""
        .ProfileEncoding = ""
        .ScreenName = ""
        .SessionsLength = 0
        .WarningLevel = 0
    End With

    InfoExtract = "0"
    Dim MyTLV As TLV
    
    Packet.Pos = &H11
    
    OutInfo.ScreenName = Packet.ExtractByteString
    OutInfo.WarningLevel = Packet.ExtractWord / 10
    
    Dim TLVCount As Integer
    TLVCount = Packet.ExtractWord
    
    Dim i As Integer
    'Debug.Print "========= TLVs =========="
    For i = 1 To TLVCount
        Packet.ExtractTLV MyTLV.TheType, MyTLV.Length, MyTLV.Value
        'Debug.Print "Type -> " & Hex(MyTLV.TheType)
        'Debug.Print "Length -> " & MyTLV.Length
        'Debug.Print "Value -> " & MyTLV.Value
        
        Select Case MyTLV.TheType
            Case &H1
                OutInfo.Flags = GetWORD(MyTLV.Value)
            Case &H2
                OutInfo.CreatedTime = GetDWORD(MyTLV.Value)
            Case &H3
                OutInfo.OnlineSince = DateAdd("s", GetDWORD(MyTLV.Value), "01/01/1970 12:00:00 AM")
            Case &H4
                OutInfo.IdleTime = GetWORD(MyTLV.Value)
            Case &H5
                OutInfo.MemberSince = DateAdd("s", GetDWORD(MyTLV.Value), "01/01/1970 12:00:00 AM")
            Case &HD
                InfoExtract = ScreenName
                bStatus = False
            Case &HF
                OutInfo.SessionsLength = GetDWORD(MyTLV.Value)
            Case &H10
                OutInfo.SessionsLength = GetDWORD(MyTLV.Value)
        End Select
    Next i
    
    If Packet.Pos < Packet.Size Then
        Dim TempTLV As TLV, TempTLV2 As TLV
        Call Packet.ExtractTLV(TempTLV.TheType, TempTLV.Length, TempTLV.Value)
        Call Packet.ExtractTLV(TempTLV2.TheType, TempTLV2.Length, TempTLV2.Value)
        
        Select Case TempTLV.TheType
            Case &H1
                OutInfo.ProfileEncoding = TempTLV.Value
            Case &H3
                OutInfo.AwayMessageEncoding = TempTLV.Value
        End Select
        
        Select Case TempTLV2.TheType
            Case &H2
                OutInfo.Profile = TempTLV2.Value
            Case &H4
                OutInfo.AwayMessage = TempTLV2.Value
        End Select
    End If
    
    InfoExtract = OutInfo.ScreenName
    'Debug.Print "Exiting with " & Packet.Pos
End Function

