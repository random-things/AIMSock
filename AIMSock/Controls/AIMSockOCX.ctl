VERSION 5.00
Begin VB.UserControl AIMSockOCX 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ScaleHeight     =   2880
   ScaleWidth      =   3900
End
Attribute VB_Name = "AIMSockOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAIMSocketEvents

Private OutBuffer As cAIMPacketBuffer

Private AIM() As cAIMSocket
Private AIMIndex As Long

Private strAuthorizationCookie As String
Private strScreenName As String
Private strPassword As String
Private strProfile As String
Private strAwayMessage As String

Private IsReady As Boolean

' Socket related events
Public Event OnConnected(ByVal Index As Long, ByVal Purpose As String, ByVal Server As String, ByVal Port As Long)
Public Event OnConnecting(ByVal Index As Long, ByVal Purpose As String, ByVal Server As String, ByVal Port As Long)
Public Event OnDisconnected(ByVal Index As Long, ByVal Purpose As String, ByVal Forced As Boolean)
Public Event OnError(ByVal Index As Long, ByVal Purpose As String, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)

' Got stuff events
Public Event GotAuthorizationCookie(ByVal AuthCookie As String)
Public Event GotBOSRights()
Public Event GotBuddy(ByVal ScreenName As String, _
                      ByVal GroupID As Long, _
                      ByVal BuddyID As Long, _
                      ByVal BuddyAlias As String, _
                      ByVal BuddyLocalMail As String, _
                      ByVal BuddySMS As String, _
                      ByVal BuddyComment As String, _
                      ByVal AlertType As Byte, _
                      ByVal AlertEvent As Byte, _
                      ByVal AlertSound As String, _
                      ByVal CTime As String)
Public Event GotBuddyListRights()
Public Event GotDenyListBuddy(ByVal ScreenName As String, ByVal GroupID As Long, ByVal BuddyID As Long)
Public Event GotEmailAddress(ByVal EmailAddress As String)
Public Event GotGroup(ByVal GroupName As String, _
                      ByVal GroupID As Long, _
                      ByVal BuddyID As Long, _
                      ByVal MemberBuddies As String)
Public Event GotICBMParameters()
Public Event GotIM(ByVal ScreenName As String, ByVal WarningLevel As Integer, ByVal Message As String)
Public Event GotLocationRights()
Public Event GotNewBOSHost(ByVal Server As String)
Public Event GotPermitListBuddy(ByVal ScreenName As String, ByVal GroupID As Long, ByVal BuddyID As Long)
Public Event GotSSIRights()
Public Event GotSSIInformation(ByVal NumberOfItems As Integer)
Public Event GotUserInformation(ByVal ScreenName As String)
Public Event GotWarned(ByVal ScreenName As String, ByVal Increase As Integer)

' Something happened events
Public Event OnAuthorization(ByVal Success As Boolean)
Public Event OnAuthorizationError(ByVal ErrorCode As Long, ByVal ErrorURL As String, ByVal ErrorInfo As String)
Public Event OnBuddySignedOff(ByVal ScreenName As String)
Public Event OnBuddySignedOn(ByVal ScreenName As String, ByVal StatusChange As Boolean)
Public Event OnGenericError(ByVal ErrorString As String, ByVal Family As String) ' Maybe handles all errors, we'll see.
Public Event OnLoggedOnAs(ByVal ScreenName As String)
Public Event OnServerSwitch()
Public Event OnTyping(ByVal Channel As Integer, ByVal ScreenName As String, ByVal NotifyType As Integer)
Public Event OnUpdatedOutInfo(ByVal ScreenName As String, _
                              ByVal WarningLevel As Long, _
                              ByVal IdleTime As Long, _
                              ByVal Flags As Long, _
                              ByVal CreatedTime As String, _
                              ByVal MemberSince As String, _
                              ByVal OnlineSince As String, _
                              ByVal SessionsLength As Long, _
                              ByVal Capabilities As Long, _
                              ByVal AwayMessage As String, _
                              ByVal AwayMessageEncoding As String, _
                              ByVal Profile As String, _
                              ByVal ProfileEncoding As String)
Public Event OnWarnReply(ByVal Increase As Integer, ByVal Total As Integer)

Public Event Sent(ByVal Message As String)

Public Property Let ScreenName(ByVal newScreenName As String)
    strScreenName = newScreenName
End Property

Public Property Get ScreenName() As String
    ScreenName = strScreenName
End Property

Public Property Let Password(ByVal newPassword As String)
    strPassword = newPassword
End Property

Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Let Profile(ByVal newProfile As String)
    strProfile = newProfile
End Property

Public Property Get Profile() As String
    Profile = strProfile
End Property

Public Property Let AwayMessage(ByVal newAwayMessage As String)
    strAwayMessage = newAwayMessage
End Property

Public Property Get AwayMessage() As String
    AwayMessage = strAwayMessage
End Property

Public Sub Connect(ByVal Server As String, ByVal Port As Long)
    AIM(0).Connect Server, Port
End Sub

Public Sub SetProfileAway()
    Dim AIMLoc As cAIMLocationServices
    Set AIMLoc = New cAIMLocationServices
    
    Set OutBuffer = AIMLoc.SetUserInformation(strProfile, strAwayMessage)
    
    AIM(0).SendData OutBuffer.GetBuffer(NormalTransit)
    
    Set AIMLoc = Nothing
End Sub

Public Sub Disconnect()
    AIM(0).Disconnect
End Sub

Public Sub SendMessage(ByVal ScreenName As String, ByVal Message As String) ', Optional ByVal RequireResponse As Boolean = False)
    ' Requiring a response means that if the recipient has a buddy icon
    ' their client (assuming it is AIM) will silently respond to the
    ' message, thus giving us a chance to warn them.
    
    Dim AIMICBM As cAIMICBM
    Set AIMICBM = New cAIMICBM
    
    Set OutBuffer = AIMICBM.SendIM(ScreenName, Message)
    'If RequireResponse = True Then OutBuffer.InsertTLV 9, ""
    AIM(0).SendData OutBuffer.GetBuffer(2)
    
    Set AIMICBM = Nothing
End Sub

Private Sub IAIMSocketEvents_OnConnected(ByVal Index As Long, ByVal Server As String, ByVal Port As Long)
    RaiseEvent OnConnected(Index, AIM(Index).Purpose, Server, Port)
End Sub

Private Sub IAIMSocketEvents_OnConnecting(ByVal Index As Long, ByVal Server As String, ByVal Port As Long)
    RaiseEvent OnConnected(Index, AIM(Index).Purpose, Server, Port)
End Sub

Private Sub IAIMSocketEvents_OnDisconnected(ByVal Index As Long, ByVal Forced As Boolean)
    RaiseEvent OnDisconnected(Index, AIM(Index).Purpose, Forced)
End Sub

Private Sub IAIMSocketEvents_OnError(ByVal Index As Long, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    RaiseEvent OnError(Index, AIM(Index).Purpose, ErrorNumber, ErrorDescription)
End Sub

Private Sub IAIMSocketEvents_OnPacketReceived(ByVal Index As Long, ByVal strPacket As String)
    Dim PA As cAIMPacketAnalyzer
    Set PA = New cAIMPacketAnalyzer
    
    PA.PacketData = strPacket
    Debug.Print DebugOutput(strPacket)
    
    Select Case PA.Channel
        Case Channels.NewConnection
            Call Handle_NewConnection(PA)
        Case Channels.NormalTransit
            Call Handle_NormalTransit(PA)
        Case Channels.ErrorChannel
        
        Case Channels.CloseConnection
        
        Case Channels.NoOperation
        
    End Select
End Sub

Private Sub Handle_NewConnection(ByRef Packet As cAIMPacketAnalyzer)
    ' Reply to the server-ping
    
    If Len(strAuthorizationCookie) = 0 Then
        Dim AIMAuth As cAIMAuthorization
        Set AIMAuth = New cAIMAuthorization
        
        ' Reply to server ping
        With OutBuffer
            .InsertDWORD 1
            AIM(0).SendData OutBuffer.GetBuffer(Channels.NewConnection)
            .Clear
        End With

        ' Send login request
        Set OutBuffer = AIMAuth.RequestLogin(strScreenName)
        AIM(0).SendData OutBuffer.GetBuffer(Channels.NormalTransit)
        Set AIMAuth = Nothing
        
    Else    ' Send the auth cookie in place of the ping
        OutBuffer.InsertDWORD 1
        OutBuffer.InsertTLV TLVs.AuthorizationCookie, strAuthorizationCookie
        AIM(0).SendData OutBuffer.GetBuffer(Channels.NewConnection)
    End If
End Sub

Private Sub Handle_NormalTransit(ByRef Packet As cAIMPacketAnalyzer)
    Select Case Packet.FamilyID
        Case Families.Authentication
            Call Handle_Authentication(Packet)
        Case Families.Service
            Call Handle_Service(Packet)
        Case Families.BasicOscarService
            Call Handle_BasicOscarService(Packet)
        Case Families.BuddyList
            Call Handle_BuddyList(Packet)
        Case Families.Information
            Call Handle_LocationServices(Packet)
        Case Families.Messaging
            Call Handle_ICBM(Packet)
        Case Families.ServerStoredInformation
            Call Handle_SSI(Packet)
        Case Else
            Debug.Print "Unhandled family packet: 0x" & MakeWORD(Packet.FamilyID)
    End Select
End Sub

Private Sub Handle_Authentication(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMAuth As cAIMAuthorization
    Set AIMAuth = New cAIMAuthorization

    Select Case Packet.SubTypeID
        Case AuthSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "Authentication")
        
        Case AuthSIDs.SC_LoginResponse
            Dim NewServer As String
            Dim ErrorCode As Long
            Dim ErrorURL As String
            Dim ExtraErrorInfo As String
            
            Dim TheType As Long, TheLength As Long, TheValue As String
            
            While Packet.Pos < Packet.Size
                Packet.ExtractTLV TheType, TheLength, TheValue
                
                Select Case TheType
                    Case TLVs.ScreenName
                        RaiseEvent OnLoggedOnAs(TheValue)
                        
                    Case TLVs.AuthorizationCookie
                        strAuthorizationCookie = TheValue
                        RaiseEvent GotAuthorizationCookie(strAuthorizationCookie)
                    
                    Case TLVs.BOSHost
                        NewServer = TheValue
                        RaiseEvent GotNewBOSHost(NewServer)
                        RaiseEvent OnServerSwitch
                        
                        AIM(0).Disconnect
                        AIM(0).Connect Split(NewServer, ":")(0), Split(NewServer, ":")(1)
                    
                    Case TLVs.ErrorCode
                        ErrorCode = GetWORD(TheValue)
                    
                    Case TLVs.ErrorURL
                        ErrorURL = TheValue
                        
                    Case TLVs.Email
                        RaiseEvent GotEmailAddress(TheValue)
                End Select
            Wend
            
            If Len(strAuthorizationCookie) > 0 And Len(NewServer) > 0 Then
                RaiseEvent OnAuthorization(True)
            End If
            
            If Len(ErrorCode) > 0 And Len(ErrorURL) > 0 Then
                RaiseEvent OnAuthorization(False)
                
                Select Case ErrorCode
                    Case &H1, &H4:
                        ExtraErrorInfo = "The screen name or password you entered is not valid"
                    Case &H5
                        ExtraErrorInfo = "The password you entered is not valid"
                    Case &H6
                        ExtraErrorInfo = "Internal client error: bad input to authorizer"
                    Case &H8
                        ExtraErrorInfo = "Your account has been deleted"
                    Case &HC, &HD, &H12, &H13, &H14, &H15, &H1A, &H1F
                        ExtraErrorInfo = "The AOL Instant Messenger service is temporarily unavailable"
                    Case &H11
                        ExtraErrorInfo = "Suspended account"
                    Case &H18, &H1D
                        ExtraErrorInfo = "You are attempting to sign on again too soon"
                    Case &H1B, &H1C
                        ExtraErrorInfo = "You are running an old client version and need to upgrade"
                    Case &H20
                        ExtraErrorInfo = "Invalid SecurID"
                    Case &H22
                        ExtraErrorInfo = "Account suspended because of your age (age < 13)"
                End Select
                
                RaiseEvent OnAuthorizationError(ErrorCode, ErrorURL, ExtraErrorInfo)
            End If
        
        Case AuthSIDs.SC_LoginRequestResponse
            Dim strHashingKey As String
            strHashingKey = Packet.ExtractString
            
            Set OutBuffer = AIMAuth.GenerateLogin(strScreenName, strPassword, strHashingKey)
            AIM(0).SendData OutBuffer.GetBuffer(2)
    End Select
    
    Set AIMAuth = Nothing
End Sub

Private Sub Handle_Service(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMService As cAIMService
    Set AIMService = New cAIMService
    
    Dim AIMSSI As cAIMSSI
    Dim AIMLocation As cAIMLocationServices
    Dim AIMBuddyList As cAIMBuddyList
    Dim AIMICBM As cAIMICBM

    Select Case Packet.SubTypeID
        Case ServiceSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "Service")
            
        Case ServiceSIDs.SC_AllowedSNACs
            ' We can pretty much discard these.  Screw them.
            ' However, we do need to send a response to this.
            
            Set OutBuffer = AIMService.SetClientVersions
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Set Client Versions")
            
        Case ServiceSIDs.SC_HostVersions
            ' Once again, we don't really care much about
            ' anything the server has to say.  Screw it.
            
            Set OutBuffer = AIMService.RequestRateParameters
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Request Rate Parameters")
            
        Case ServiceSIDs.SC_RequestedRateInfo
            ' The rate class information arrives here.  It
            ' would be cool to implement this at some point.
            
            Set OutBuffer = AIMService.GenerateRateResponse(Packet)
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Rate Response")
            
            Set AIMSSI = New cAIMSSI
            Set OutBuffer = AIMSSI.RequestSSIRights
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("SSI Rights Request")
            
            Set OutBuffer = AIMSSI.RequestSSIData
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("SSI Data Request")
            
            Set OutBuffer = AIMService.RequestSelfInfo
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Self info request")
            
            Set AIMLocation = New cAIMLocationServices
            Set OutBuffer = AIMLocation.RequestLocationRights
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Location Rights Request")
            
            Set AIMBuddyList = New cAIMBuddyList
            Set OutBuffer = AIMBuddyList.RequestBuddyListRights
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Buddy List Rights Request")
            
            Set AIMICBM = New cAIMICBM
            Set OutBuffer = AIMICBM.RequestParameters
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("ICBM Paramters Request")
        
        Case ServiceSIDs.SC_SelfUserInfo
            Set AIMLocation = New cAIMLocationServices
        
            Dim ResultString As String
            ResultString = AIMLocation.InfoExtract(Packet)
            If ResultString <> "0" Then
                RaiseEvent OnUpdatedOutInfo(OutInfo.ScreenName, _
                                            OutInfo.WarningLevel, _
                                            OutInfo.IdleTime, _
                                            OutInfo.Flags, _
                                            OutInfo.CreatedTime, _
                                            OutInfo.MemberSince, _
                                            OutInfo.OnlineSince, _
                                            OutInfo.SessionsLength, _
                                            OutInfo.Capabilities, _
                                            OutInfo.AwayMessage, _
                                            OutInfo.AwayMessageEncoding, _
                                            OutInfo.Profile, _
                                            OutInfo.ProfileEncoding)
            End If
            
        Case ServiceSIDs.SC_GotWarned
            '0000:  2A 02 3D E3 00 14 00 01 00 10 80 00 8C 32 5B B6   *=ã...€.Œ2[¶
            '0010:  00 06 00 01 00 02 00 03 00 32                     .....2......
            Packet.ExtractBytes 8   'The 0x06010203
            Dim WarnIncrease As Integer, Warner As String
            WarnIncrease = Packet.ExtractWord / 10
            
            If Packet.Pos >= Packet.Size Then
                Warner = "[Anonymous]"
            Else
                Warner = Packet.ExtractByteString
            End If
            
            RaiseEvent GotWarned(Warner, WarnIncrease)
    End Select
    
    Set AIMService = Nothing
    Set AIMSSI = Nothing
    Set AIMLocation = Nothing
    Set AIMBuddyList = Nothing
    Set AIMICBM = Nothing
End Sub

Private Sub Handle_BasicOscarService(ByRef Packet As cAIMPacketAnalyzer)
    Select Case Packet.SubTypeID
        Case BOSSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "BOS")
        Case BOSSIDs.SC_BOSRights
            ' TODO: Parse data
            RaiseEvent GotBOSRights
    End Select
End Sub

Private Sub Handle_LocationServices(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMLocation As cAIMLocationServices
    Set AIMLocation = New cAIMLocationServices

    Select Case Packet.SubTypeID
        Case LocationSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "Information/Location")
        Case LocationSIDs.SC_LocationRightsResponse
            RaiseEvent GotLocationRights
            ' TODO: Parse data
            
            ' Set profile
            Set OutBuffer = AIMLocation.SetUserInformation(strProfile, strAwayMessage)
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Initial User Information")
        
        Case LocationSIDs.SC_UserInformation
            RaiseEvent GotUserInformation(Packet.ExtractByteString)
            Dim ResultString As String
            ResultString = AIMLocation.InfoExtract(Packet)
            If ResultString <> "0" Then
                RaiseEvent OnUpdatedOutInfo(OutInfo.ScreenName, _
                                            OutInfo.WarningLevel, _
                                            OutInfo.IdleTime, _
                                            OutInfo.Flags, _
                                            OutInfo.CreatedTime, _
                                            OutInfo.MemberSince, _
                                            OutInfo.OnlineSince, _
                                            OutInfo.SessionsLength, _
                                            OutInfo.Capabilities, _
                                            OutInfo.AwayMessage, _
                                            OutInfo.AwayMessageEncoding, _
                                            OutInfo.Profile, _
                                            OutInfo.ProfileEncoding)
            End If
    End Select
    
    Set AIMLocation = Nothing
End Sub

Private Sub Handle_BuddyList(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMLocation As cAIMLocationServices
    Set AIMLocation = New cAIMLocationServices

    Select Case Packet.SubTypeID
        Case BuddyListSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "Buddy List")
    
        Case BuddyListSIDs.SC_BuddyListRightsResponse
            ' TODO: Parse data
            RaiseEvent GotBuddyListRights
            
        Case BuddyListSIDs.SC_BuddySignedOn
            Dim Name As String
            Dim ResultString As String
            Dim IsJustStatus As Boolean
            Name = Packet.ExtractByteString
            ResultString = AIMLocation.InfoExtract(Packet, IsJustStatus)
            If ResultString <> "0" Then
                RaiseEvent OnUpdatedOutInfo(OutInfo.ScreenName, _
                                            OutInfo.WarningLevel, _
                                            OutInfo.IdleTime, _
                                            OutInfo.Flags, _
                                            OutInfo.CreatedTime, _
                                            OutInfo.MemberSince, _
                                            OutInfo.OnlineSince, _
                                            OutInfo.SessionsLength, _
                                            OutInfo.Capabilities, _
                                            OutInfo.AwayMessage, _
                                            OutInfo.AwayMessageEncoding, _
                                            OutInfo.Profile, _
                                            OutInfo.ProfileEncoding)
            End If
            RaiseEvent OnBuddySignedOn(Name, IsJustStatus)
            
        Case BuddyListSIDs.SC_BuddySignedOff
            RaiseEvent OnBuddySignedOff(Packet.ExtractByteString)
    End Select
    
    Set AIMLocation = Nothing
End Sub

Private Sub Handle_ICBM(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMICBM As cAIMICBM
    Set AIMICBM = New cAIMICBM

    Select Case Packet.SubTypeID
        Case ICBMSIDs.B_Error
            Call GenericError(Packet.ExtractWord, "ICBM")
    
        Case ICBMSIDs.SC_ParameterInformation
            RaiseEvent GotICBMParameters
            
            Set OutBuffer = AIMICBM.SetICBMParameters(, AllFlags, 8000)
            AIM(0).SendData OutBuffer.GetBuffer(2)
            
            RaiseEvent Sent("Set ICBM Parameters")
            
        Case ICBMSIDs.SC_Message
        ' Message arrived.
            Dim MessageID(0 To 1) As Long
            MessageID(0) = Packet.ExtractDword
            MessageID(1) = Packet.ExtractDword
            Dim MessageChannel As Integer
            MessageChannel = Packet.ExtractWord
            Dim Username As String
            Username = Packet.ExtractByteString
            Dim WarningLevel As Integer
            WarningLevel = Packet.ExtractWord / 10
            Dim TLVCount As Integer
            TLVCount = Packet.ExtractWord
            Debug.Print "TLVCount -> " & TLVCount
            
            Dim TheType As Long, TheLength As Long, TheValue As String
            
            ' Pull out the TLVs
            While TLVCount > 0
                Packet.ExtractTLV TheType, TheLength, TheValue
            
                'Debug.Print "TheType -> " & TheType
                'Debug.Print "TheLength -> " & TheLength
                'Debug.Print "TheValue -> " & TheValue
                
                TLVCount = TLVCount - 1
            Wend
            
            Dim TheMessage As String
            Dim MessageBlock As cAIMPacketAnalyzer
            Set MessageBlock = New cAIMPacketAnalyzer
            Packet.ExtractTLV TheType, TheLength, TheValue
            MessageBlock.PacketData = TheValue
            MessageBlock.Pos = 1
            
            ' This means it's not really a message, but could be an icon, etc.
            If MessageBlock.ReadWord <> &H501 Then
                Set AIMICBM = Nothing
                Exit Sub
            End If
            
            Debug.Print "Message"
            Debug.Print DebugOutput(MessageBlock.Data)
            MessageBlock.ExtractTLV TheType, TheLength, TheValue    '0x501 TLV
            Debug.Print DebugOutput(MessageBlock.Data)
            MessageBlock.ExtractTLV TheType, TheLength, TheValue    '0x101 TLV - Holds message
            MessageBlock.PacketData = TheValue
            MessageBlock.Pos = 1
            Debug.Print DebugOutput(MessageBlock.Data)
            MessageBlock.ExtractDword 'Throw it away
            TheMessage = MessageBlock.ExtractBytes(TheLength - 4)
            Debug.Print "TheMessage -> " & TheMessage
            
            If Len(TheMessage) > 0 Then
                RaiseEvent GotIM(Username, WarningLevel, TheMessage)
            End If
            
            Set MessageBlock = Nothing
            
        Case ICBMSIDs.SC_EvilReply
            RaiseEvent OnWarnReply(Packet.ExtractWord / 10, Packet.ExtractWord / 10)
            
        Case ICBMSIDs.B_TypingNotification
            Dim NotifyType As Integer
            
            Packet.ExtractBytes 8   'Get rid of notification cookie.
            RaiseEvent OnTyping(Packet.ExtractWord, Packet.ExtractByteString, Packet.ExtractWord)
            
    End Select
    
    Set AIMICBM = Nothing
End Sub

Private Sub Handle_SSI(ByRef Packet As cAIMPacketAnalyzer)
    Dim AIMSSI As cAIMSSI
    Set AIMSSI = New cAIMSSI
    
    Select Case Packet.SubTypeID
        Case SSISIDs.B_Error
            Call GenericError(Packet.ExtractWord, "SSI")
            
        Case SSISIDs.SC_RightsResponse
            ' TODO: Parse data
            RaiseEvent GotSSIRights
            
        Case SSISIDs.SC_SSIResponse
            ' TODO: Parse data
            ' This is a big one, it's the buddy list!
            
            ' ============================
            ' Ok, here we go.
            
            Packet.ExtractByte  ' 0x00 (SSI version)
            Dim NumItems As Integer
            NumItems = Packet.ExtractWord
            
            ' Moved the event to the bottom so that I can parse the list
            ' after I'm sure all buddies have been added.
            
            Dim i As Integer
            Dim ScreenName As String
            Dim GroupID As Long
            Dim BuddyID As Long
            Dim ItemType As Long
            Dim TheType As Long, TheLength As Long, TheValue As String
            Dim BuddyAlias As String
            Dim BuddyLocalMail As String
            Dim BuddySMS As String
            Dim BuddyComment As String
            Dim AlertType As Byte, AlertEvent As Byte
            Dim AlertSound As String
            Dim CTime As String
            Dim MemberBuddies As String
            
            'Debug.Print "NumItems -> " & NumItems
            For i = 1 To NumItems
                'Debug.Print "Packet Data"
                'Debug.Print DebugOutput(Packet.Data)
                ScreenName = Packet.ExtractString
                'Debug.Print "ScreenName -> " & ScreenName
                GroupID = Packet.ExtractWord
                'Debug.Print "GroupID -> " & Hex(GroupID)
                BuddyID = Packet.ExtractWord
                'Debug.Print "BuddyID -> " & Hex(BuddyID)
                ItemType = Packet.ExtractWord
                'Debug.Print "ItemType -> " & Hex(ItemType)
                
                Dim InfoParse As cAIMPacketAnalyzer
                Set InfoParse = New cAIMPacketAnalyzer
                
                ' The word of extract string is the length of the
                ' rest of the data.  So in theory, all remaining
                ' comes in here.
                InfoParse.PacketData = Packet.ExtractString
                InfoParse.Pos = 1
                'Debug.Print "TLVs -> "
                'Debug.Print DebugOutput(InfoParse.PacketData)
                
                'Debug.Print InfoParse.Pos & "/" & InfoParse.Size - 3
                Do While InfoParse.Pos < InfoParse.Size - 3
                    InfoParse.ExtractTLV TheType, TheLength, TheValue
                    'Debug.Print "Extracted TLV Type -> " & TheType
                    
                    Select Case ItemType
                        Case SSIItem.Buddy
                            Select Case TheType
                                Case &H131
                                    BuddyAlias = TheValue
                                Case &H137
                                    BuddyLocalMail = TheValue
                                Case &H13A
                                    BuddySMS = TheValue
                                Case &H13C
                                    BuddyComment = TheValue
                                Case &H13D
                                    AlertType = Asc(Mid(TheValue, 1, 1))
                                    AlertEvent = Asc(Mid(TheValue, 2, 1))
                                Case &H13E
                                    AlertSound = TheValue
                                Case &H145
                                    CTime = TheValue
                            End Select
                            
                        Case SSIItem.Group
                            Select Case TheType
                                Case &HC8
                                    MemberBuddies = TheValue
                            End Select
                            
                    End Select
                Loop
                
                Select Case ItemType
                    Case SSIItem.Buddy
                        RaiseEvent GotBuddy(ScreenName, GroupID, BuddyID, BuddyAlias, BuddyLocalMail, BuddySMS, BuddyComment, AlertType, AlertEvent, AlertSound, CTime)
                    Case SSIItem.Group
                        RaiseEvent GotGroup(ScreenName, GroupID, BuddyID, MemberBuddies)
                    Case SSIItem.Permit
                        RaiseEvent GotPermitListBuddy(ScreenName, GroupID, BuddyID)
                    Case SSIItem.Deny
                        RaiseEvent GotDenyListBuddy(ScreenName, GroupID, BuddyID)
                End Select
            Next i
            
            RaiseEvent GotSSIInformation(NumItems)
            
            ' ============================
            
            'Debug.Print DebugOutput(Packet.PacketData)
            
            If IsReady = False Then
                ' Activate SSI
                Set OutBuffer = AIMSSI.ActivateSSI
                AIM(0).SendData OutBuffer.GetBuffer(2)
                
                RaiseEvent Sent("SSI Activate")
                
                ' NOTIFY SERVER THAT CLIENT IS READY
                ' *** THIS IS IMPORTANT BECAUSE IT IS A REQUIRED STEP AND ***
                ' *** LOCATED IN AN OTHERWISE NOT VERY NOTICABLE PLACE    ***
                Dim AIMService As cAIMService
                Set AIMService = New cAIMService
                
                Set OutBuffer = AIMService.ClientReady()
                AIM(0).SendData OutBuffer.GetBuffer(2)
                
                RaiseEvent Sent("Client Ready")
                IsReady = True
            End If
    End Select
    
    Set AIMSSI = Nothing
    Set AIMService = Nothing
End Sub

Private Function GenericError(ByVal ErrorCode As Long, ByVal Family As String) As String
    Select Case ErrorCode
        Case &H1
            GenericError = "Invalid SNAC header."
        Case &H2
            GenericError = "Server rate limit exceeded."
        Case &H3
            GenericError = "Client rate limit exceeded."
        Case &H4
            GenericError = "That user is not currently available."
        Case &H5
            GenericError = "Requested service is unavailable."
        Case &H6
            GenericError = "Requested service is not defined."
        Case &H7
            GenericError = "You sent an obsolete SNAC."
        Case &H8
            GenericError = "Request is not supported by the server."
        Case &H9
            GenericError = "Request is not supported by client."
        Case &HA
            GenericError = "Request was refused by the client."
        Case &HB
            GenericError = "Your reply is too large to send."
        Case &HC
            GenericError = "Responses lost."
        Case &HD
            GenericError = "Request denied."
        Case &HE
            GenericError = "Your message is formatted incorrectly.  (Client error)"
        Case &HF
            GenericError = "Insufficient rights."
        Case &H10
            GenericError = "You have the recipient of this message blocked."
        Case &H11
            GenericError = "Sender is too evil."
        Case &H12
            GenericError = "Recipient is too evil."
        Case &H13
            GenericError = "Recipient is temporarily unavailable."
        Case &H14
            GenericError = "No match."
        Case &H15
            GenericError = "List overflow."
        Case &H16
            GenericError = "Your request was ambiguous."
        Case &H17
            GenericError = "The server's queue is full.  Please try again later."
        Case &H18
            GenericError = "Not while on AOL."
        Case Else
            GenericError = "Unknown SNAC failure (0x" & Hex(ErrorCode) & ")"
    End Select
    
    RaiseEvent OnGenericError(GenericError, Family)
End Function


Private Sub IAIMSocketEvents_OnPacketSent(ByVal Index As Long)
    '
End Sub

Private Sub UserControl_Initialize()
    Set OutBuffer = New cAIMPacketBuffer

    AIMIndex = 0
    ReDim AIM(AIMIndex)
    Set AIM(0) = New cAIMSocket

    'AIM(0).Socket = sckAIM
    AIM(0).EventSink = Me
    AIM(0).Index = AIMIndex
    AIM(0).Purpose = "BOS"
End Sub
