VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMAuthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function RequestLogin(ByVal ScreenName As String) As cAIMPacketBuffer
    Set RequestLogin = New cAIMPacketBuffer
    RequestLogin.InsertSNAC Authentication, AuthSIDs.CS_RequestLogin, 0, 0, &H170006
    RequestLogin.InsertTLV TLVs.ScreenName, ScreenName
    'RequestLogin.InsertTLV TLVs.SecurID, ""
End Function

Public Function GenerateLogin(ByVal ScreenName As String, ByVal Password As String, ByVal HashSeed As String) As cAIMPacketBuffer
    Dim MD5 As cMD5
    Set MD5 = New cMD5

    Set GenerateLogin = New cAIMPacketBuffer
    With GenerateLogin
        .InsertSNAC Authentication, AuthSIDs.CS_SendLoginData, 0, 0, &H170002
        .InsertTLV TLVs.ScreenName, ScreenName
        .InsertTLV TLVs.TransferResponse, MD5.AIMEncryptPw(HashSeed, Password)
        .InsertTLV TLVs.MachineInfo, ""
        .InsertTLV TLVs.VersionString, "AOL Instant Messenger, version 5.2.3292/WIN32"
        .InsertTLV TLVs.ClientID, MakeWORD(&H109)
        .InsertTLV TLVs.VersionMajor, MakeWORD(&H5)
        .InsertTLV TLVs.VersionMinor, MakeWORD(&H2)
        .InsertTLV TLVs.VersionPoint, MakeWORD(&H0)
        .InsertTLV TLVs.VersionBuild, MakeWORD(&HCDC)
        .InsertTLV TLVs.ClientDistribution, MakeDWORD(&HEE)
        .InsertTLV TLVs.Language, "en"
        .InsertTLV TLVs.Country, "us"
        .InsertTLV TLVs.ForceSSI, Chr(&H1)
    End With
    
    Set MD5 = Nothing
End Function
