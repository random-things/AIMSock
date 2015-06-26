Attribute VB_Name = "modAIMData"
Option Explicit

Public Enum Channels
    NewConnection = 1
    NormalTransit = 2
    ErrorChannel = 3
    CloseConnection = 4
    NoOperation = 5
End Enum

Public Enum Families
    Service = &H1
    Information = &H2
    BuddyList = &H3
    Messaging = &H4
    Advertisements = &H5
    Invitation = &H6
    Administrative = &H7
    Popup = &H8
    BasicOscarService = &H9
    UserLookup = &HA
    Stats = &HB
    Translate = &HC
    ChatNavigation = &HD
    Chat = &HE
    Search = &HF
    ServerStoredThemes = &H10
    ServerStoredInformation = &H13
    Authentication = &H17
    Email = &H18
End Enum

Public Enum AuthSIDs
    B_Error = &H1
    CS_SendLoginData = &H2
    SC_LoginResponse = &H3
    CS_RequestLogin = &H6
    SC_LoginRequestResponse = &H7
    SC_RequestSecurID = &HA
    CS_SendSecurID = &HB
End Enum

Public Enum ServiceSIDs
    B_Error = &H1
    CS_ClientReady = &H2
    SC_AllowedSNACs = &H3
    CS_RequestRateInfo = &H6
    SC_RequestedRateInfo = &H7
    CS_AddRateParameter = &H8
    CS_RequestSelfInfo = &HE
    SC_SelfUserInfo = &HF
    SC_GotWarned = &H10
    CS_SetIdleTime = &H11
    SC_MOTD = &H13
    CS_SetClientVersions = &H17
    SC_HostVersions = &H18
    CS_SetStatus = &H1E
End Enum

Public Enum LocationSIDs
    B_Error = &H1
    CS_RequestLocationRights = &H2
    SC_LocationRightsResponse = &H3
    CS_SetUserInformation = &H4
    CS_RequestUserInformation = &H5
    SC_UserInformation = &H6
End Enum

Public Enum BuddyListSIDs
    B_Error = &H1
    CS_RequestBuddyListRights = &H2
    SC_BuddyListRightsResponse = &H3
    SC_BuddySignedOn = &HB
    SC_BuddySignedOff = &HC
End Enum

Public Enum ICBMSIDs
    B_Error = &H1
    CS_SetParameters = &H2
    CS_RequestParameters = &H4
    SC_ParameterInformation = &H5
    CS_Message = &H6
    SC_Message = &H7
    CS_EvilRequest = &H8
    SC_EvilReply = &H9
    B_TypingNotification = &H14
End Enum

Public Enum BOSSIDs
    B_Error = &H1
    CS_RequestBOSRights = &H2
    SC_BOSRights = &H3
    CS_SetGroupPermissions = &H4
    CS_AddPermissions = &H5
    CS_DeletePermissions = &H6
    CS_AddDenyListEntries = &H7
    CS_DeleteDenyListEntries = &H8
    SC_BOSError = &H9
End Enum

Public Enum SSISIDs
    B_Error = &H1
    CS_RequestRights = &H2
    SC_RightsResponse = &H3
    CS_RequestSSI = &H5
    SC_SSIResponse = &H6
    CS_ActivateSSI = &H7
    CS_AddItem = &H8
    CS_ModifyGroup = &H9
    CS_RemoveItem = &HA
    CS_BeginMod = &H11
    CS_EndMod = &H12
End Enum

Public Enum TLVs
    ScreenName = &H1
    Password = &H2
    VersionString = &H3
    ErrorURL = &H4
    BOSHost = &H5
    AuthorizationCookie = &H6
    SNACVersion = &H7
    ErrorCode = &H8
    DisconnectReason = &H9
    ReconnectHost = &HA
    URL = &HB
    DebugData = &HC
    GroupID = &HD   ' SNACGroupUp?
    Country = &HE
    Language = &HF
    Script = &H10
    Email = &H11
    OldPassword = &H12
    RegistrationStatusPreference = &H13
    ClientDistribution = &H14
    PersonalizedText = &H15
    ClientID = &H16
    VersionMajor = &H17
    VersionMinor = &H18
    VersionPoint = &H19
    VersionBuild = &H1A
    ErrorText = &H1B
    MIMECharSet = &H1C
    MIMELanguage = &H1D
    QContext = &H1E
    DemoData = &H1F
    DemoEvaluation = &H20
    ErrorData = &H21
    IPAddress = &H22
    SurveyFlag = &H23
    Transfer = &H24
    TransferResponse = &H25     ' Also used for Hash
    Normalize = &H26
    Progress = &H27
    ServiceUUID = &H28
    ErrorInfoClassID = &H29
    ErrorInfoData = &H30
    NewAIMBetaBuild = &H40
    NewAIMBetaURL = &H41
    NewAIMBetaInfo = &H42
    NewAIMBetaString = &H43
    LatestAIMReleaseBuild = &H44
    LatestAIMReleaseURL = &H45
    LatestAIMReleaseInfo = &H46
    LatestAIMReleaseName = &H47
    NewAIMBetaSerial = &H48
    LatestAIMReleaseSerial = &H49
    ForceSSI = &H4A ' Multi-Connection Level?
    SecurID = &H4B
    MachineInfo = &H4C ' Also used for Use MD5
    MCToken = &H4D
    MCSiteID = &H4E
    MCLogin = &H4F
    CipherInfo = &H50
    MCHTTPHeader = &H51
    MCID = &H52
    ClientProxyType = &H53
    ChangePasswordURL = &H54
    ClientUIBits = &H55
    ICQEmailLogin = &H56
    NN_RegistrationRequired = &H57
    NN_RegistrationData = &H58
    ClientDebugOn = &H80
    MCErrorSubCode = &H81
End Enum

Public Enum UserClass
    Unconfirmed = &H1
    Administrator = &H2
    AOLUser = &H4
    CommercialAOL = &H8
    Free = &H10
    Away = &H20
    ICQ = &H40
    Wireless = &H80
    ActiveBuddy = &H400
End Enum

Public Enum SSIItem
    Buddy = &H0
    Group = &H1
    Permit = &H2
    Deny = &H3
End Enum

Public Type RateClass
    ClassID As Integer
    WindowSize As Long
    ClearLevel As Long
    AlertLevel As Long
    LimitLevel As Long
    DisconnectLevel As Long
    CurrentLevel As Long
    MaxLevel As Long
    LastTime As Long
    CurrentState As Byte
End Type

Public Type TLV
    TheType As Long
    Length As Long
    Value As String
End Type

Public Type structOutInfo
    ScreenName As String
    WarningLevel As Long
    IdleTime As Long
    Flags As Long
    CreatedTime As String
    MemberSince As String
    OnlineSince As String
    SessionsLength As Long
    Capabilities As Long
    AwayMessage As String
    AwayMessageEncoding As String
    Profile As String
    ProfileEncoding As String
End Type

Public OutInfo As structOutInfo

Public MyScreenName As String
