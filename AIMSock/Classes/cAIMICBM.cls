VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMICBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ICBMFlags
    MessagesOnChannel = &H1
    MissedCallsNotifyForChannel = &H2
    DefaultFlags = &H3
    TypingNotification = &H4
    AllFlags = &HB
End Enum

Public Function RequestParameters() As cAIMPacketBuffer
    Set RequestParameters = New cAIMPacketBuffer
    RequestParameters.InsertSNAC Messaging, ICBMSIDs.CS_RequestParameters, 0, 0, &H40004
End Function

Public Function SetICBMParameters(Optional ByVal Channel As Byte = 0, _
                                  Optional ByVal Flags As ICBMFlags = 3, _
                                  Optional ByVal MaxMessageSize As Long = 512, _
                                  Optional ByVal MaxSenderWarn As Long = 999, _
                                  Optional ByVal MaxRecvWarn As Long = 999, _
                                  Optional ByVal MinMessageInterval As Long = 0, _
                                  Optional ByVal Unknown As Long = 0) As cAIMPacketBuffer
                                  
    Set SetICBMParameters = New cAIMPacketBuffer
    
    With SetICBMParameters
        .InsertSNAC Messaging, ICBMSIDs.CS_SetParameters, 0, 0, &H40002
        .InsertWORD Channel
        .InsertDWORD Flags
        .InsertWORD MaxMessageSize
        .InsertWORD MaxSenderWarn
        .InsertWORD MaxRecvWarn
        .InsertWORD MinMessageInterval
        .InsertWORD Unknown
    End With
End Function

Public Function SendIM(ByVal strScreenName As String, ByVal strMessage As String) As cAIMPacketBuffer
    Set SendIM = New cAIMPacketBuffer

    With SendIM
        .InsertSNAC Messaging, ICBMSIDs.CS_Message, 0, 0, &H40006
        .InsertDWORD GetDWORD("MMCO")
        .InsertDWORD GetDWORD("OKIE")
        .InsertWORD 1
        .InsertByteString strScreenName
        .InsertWORD 2
        .InsertWORD Len(strMessage) + &HF
        .InsertWORD &H501   'Type
        .InsertWORD 3       'Length
        .InsertNonNTString Chr(1) & Chr(1) & Chr(2)
        .InsertWORD &H101   'Type
        .InsertWORD Len(strMessage) + 4 'Length
        .InsertDWORD 0
        .InsertNonNTString strMessage
    End With
End Function

Public Function Warn(ByVal strScreenName As String, Optional ByVal Anonymous As Byte = 0) As cAIMPacketBuffer
    Set Warn = New cAIMPacketBuffer
    
    Warn.InsertSNAC Messaging, ICBMSIDs.CS_EvilRequest, 0, 0, &H40008
    Warn.InsertWORD Anonymous
    Warn.InsertByteString strScreenName
End Function
