Attribute VB_Name = "modEnum"
Option Explicit

'modEnum - project StealthBot
'February 2004 - Stealth [stealth at stealthbot dot net]


'UDTS
Public Type udtChanList
    Current         As String
    Flags           As Long
    Designated      As String
    staticDesignee  As String
End Type

Public Type udtOutFilters
    ofFind      As String
    ofReplace   As String
End Type

Public Type udtBanList
    Username        As String
    UsernameActual  As String
    cOperator       As String
End Type

Public Type udtAutoRespond
    Check As String
    Reply As String
End Type

Public Type udtProductInfo
    Code As String
    ShortCode As String
    KeyCount As Integer
    FullName As String
    ChannelName As String
    BNLS_ID As Long
    LogonSystem As Long
    VersionByte As Long
End Type

Public Type udtCustomCommandData
    reqAccess   As Integer
    Query       As String * 20
    Action      As String * 500
End Type

Public Type udtUserAccess
    Username        As String
    Rank            As Integer
    Flags           As String
    Groups          As Collection
    BanMessage      As String
End Type

'Public Type udtQueue
'    Message     As String
'    Priority    As Byte
'End Type

Public Type udtMail
    To      As String * 30
    From    As String * 30
    Message As String * 225
End Type

Public Type udtProxyConnectionInfo
    ' config
    ServerType As enuServerTypes
    UseProxy As Boolean
    ProxyIP As String
    ProxyPort As Long
    Version As Byte
    Username As String
    Password As String
    
    ' status
    IsUsingProxy As Boolean
    Status As enuProxyStatus
    RemoteHost As String
    RemoteHostIP As String
    RemotePort As Long
    RemoteResolveHost As Boolean
End Type

Public Type udtServerRequest
    ResponseReceived As Boolean
    HandlerType As enuServerRequestHandlerType
    Command As clsCommandObj
    PacketID As Byte
    PacketCommand As Byte
    Cookie As Long
    Tag As Variant
End Type

Public Enum enuClanResponseValue
    clresSuccess = 0
    clresNameInUse = 1
    clresTooSoon = 2
    clresNotEnoughMembers = 3
    clresDecline = 4
    clresUnavailable = 5
    clresAccept = 6
    clresNotAuthorized = 7
    clresNotAllowed = 8
    clresIsFull = 9
    clresBadTag = 10
    clresBadName = 11
    clresUserNotFound = 12
End Enum

Public Enum enuClanRank
    clrankRecruit = 0
    clrankPeon = 1
    clrankGrunt = 2
    clrankShaman = 3
    clrankChieftain = 4
End Enum
   
Public Type FLASHWINFO
    cbSize      As Long
    hWnd        As Long
    dwFlags     As Long
    uCount      As Long
    dwTimeout   As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LVHITTESTINFO
    pt As POINTAPI
    Flags As Long
    iItem As Long
    iSubItem As Long
End Type

Public Enum inetQueueModes
    inqReset = 0
    inqAdd = 1
    inqGet = 2
End Enum

Public Enum enuProxyStatus
    psNotConnected = 0
    psRequestingMethod = 1
    psLoggingOn = 2
    psRequestingConn = 3
    psOnline = 4
End Enum

Public Enum enuWebProfileTypes
    W3XP = 1
    WAR3 = 2
End Enum

Public Enum eCharacterTypes
    Unknown& = &H0
    Amazon& = &H1
    Sorceress& = &H2
    Necromancer& = &H3
    Paladin& = &H4
    Barbarian& = &H5
    Druid& = &H6
    Assassin& = &H7
End Enum

Public Enum enuSettingsPanels
    spSplash = 0
    spConnectionConfig = 1
    spConnectionAdvanced = 2
    spInterfaceGeneral = 3
    spInterfaceChat = 4
    spInterfaceFontsColors = 5
    spGenModeration = 6
    spGenGreets = 7
    spGenIdles = 8
    spGenCommands = 9
    spGenLogging = 10
End Enum

Public Enum enuDBActions
    AddEntry = 1
    RemEntry = 2
    ModEntry = 3
End Enum

Public Enum enuServerRequestHandlerType
    reqInternal = 1
    reqUserInterface = 2
    reqScriptingCall = 3
    reqUserCommand = 4
End Enum

Public Enum enuSHA1Type
    ' standard SHA1
    shaStandard = 0
    ' standard SHA1 with 5 longs reverse endian (Warden)
    shaStandardRevEndian = 1
    ' broken SHA1 (OLS passwords & keys)
    shaBrokenROL = 2
End Enum
