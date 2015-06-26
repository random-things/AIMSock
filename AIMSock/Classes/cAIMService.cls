VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ClientReady() As cAIMPacketBuffer
    Set ClientReady = New cAIMPacketBuffer
    
    With ClientReady
        .InsertSNAC Service, ServiceSIDs.CS_ClientReady, 0, 0, &H10002
        .InsertWORD &H1           ' Generic Controls Service
        .InsertWORD &H3           '   version 0x03
        .InsertDWORD &H1100739    '   dll version 0x01100739
        .InsertWORD &H13          ' Server Side Information Service
        .InsertWORD &H3           '   version 0x03
        .InsertDWORD &H1100739    '   dll version 0x01100739
        .InsertWORD &H2           ' Location Service
        .InsertWORD &H1           '   version 0x01
        .InsertDWORD &H1100739    '   dll version 0x01100739
        .InsertWORD &H3           ' Buddy List Management Service
        .InsertWORD &H1           '   version 0x01
        .InsertDWORD &H1100739    '   dll version 0x01100739
        .InsertWORD &H4           ' ICBM Service
        .InsertWORD &H1           '   version 0x01
        .InsertDWORD &H1100739    '   dll version 0x01100739
    End With
    
End Function

Public Function SetClientVersions() As cAIMPacketBuffer
    Set SetClientVersions = New cAIMPacketBuffer
    
    With SetClientVersions
        .InsertSNAC Service, ServiceSIDs.CS_SetClientVersions, 0, 0, &H10003
        .InsertWORD &H1   ' Generic Controls Service
        .InsertWORD &H3   '   version 0x03
        .InsertWORD &H13  ' Server Side Information Servic
        .InsertWORD &H3   '   version 0x03
        .InsertWORD &H2   ' Location Service
        .InsertWORD &H1   '   version 0x01
        .InsertWORD &H3   ' Buddy List Management Service
        .InsertWORD &H1   '   version 0x01
        .InsertWORD &H4   ' ICBM Service
        .InsertWORD &H1   '   version 0x01
    End With
End Function

Public Function RequestRateParameters() As cAIMPacketBuffer
    Set RequestRateParameters = New cAIMPacketBuffer
    
    RequestRateParameters.InsertSNAC Service, ServiceSIDs.CS_RequestRateInfo, 0, 0, &H10006
End Function

Public Function GenerateRateResponse(Packet As cAIMPacketAnalyzer) As cAIMPacketBuffer
    Set GenerateRateResponse = New cAIMPacketBuffer
    
    Dim i As Long, nClasses As Integer, Temp As RateClass
    nClasses = Packet.ExtractWord ' Number of rate classes to extract

    GenerateRateResponse.InsertSNAC Service, ServiceSIDs.CS_AddRateParameter, 0, 0, &H10008
    
    For i = 1 To nClasses
        Temp.ClassID = Packet.ExtractWord
        Temp.WindowSize = Packet.ExtractDword
        Temp.ClearLevel = Packet.ExtractDword
        Temp.AlertLevel = Packet.ExtractDword
        Temp.LimitLevel = Packet.ExtractDword
        Temp.DisconnectLevel = Packet.ExtractDword
        Temp.CurrentLevel = Packet.ExtractDword
        Temp.MaxLevel = Packet.ExtractDword
        Temp.LastTime = Packet.ExtractDword
        Temp.CurrentState = Packet.ExtractByte
        
        With GenerateRateResponse
            .InsertWORD CLng(Temp.ClassID)
        End With
    Next i
End Function

Public Function RequestSelfInfo() As cAIMPacketBuffer
    Set RequestSelfInfo = New cAIMPacketBuffer
    
    RequestSelfInfo.InsertSNAC Service, ServiceSIDs.CS_RequestSelfInfo, 0, 0, &H1000E
End Function

Public Function SetIdleTime(ByVal IdleTime As Long) As cAIMPacketBuffer
    Set SetIdleTime = New cAIMPacketBuffer
    
    SetIdleTime.InsertSNAC Service, ServiceSIDs.CS_SetIdleTime, 0, 0, &H10011
    SetIdleTime.InsertDWORD IdleTime
End Function
