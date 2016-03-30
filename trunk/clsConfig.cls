VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Written by Pyro
'  2016-3-25

Private Const CONFIG_VERSION        As Integer = 6

'Config sections
Private Const SECTION_MAIN          As String = "Main"
Private Const SECTION_POSITION      As String = "Position"
Private Const SECTION_OTHER         As String = "Other"
Private Const SECTION_OVERRIDE      As String = "Override"
Private Const SECTION_FILEPATH      As String = "FilePaths"

'New sections
Private Const SECTION_UI            As String = "UI"
Private Const SECTION_UI_POS        As String = "UI-Position"
Private Const SECTION_CLIENT        As String = "Client"
Private Const SECTION_FEATURES      As String = "Features"
Private Const SECTION_MODERATION    As String = "Moderation"
Private Const SECTION_LOGGING       As String = "Logging"
Private Const SECTION_QUEUE         As String = "Queue"
Private Const SECTION_SCRIPTING     As String = "Scripting"
Private Const SECTION_EMULATION     As String = "Emulation"
Private Const SECTION_DEBUG         As String = "Debug"

'Non-setting variables
Private m_ConfigVersion             As Long
Private m_ConfigPath                As String
Private m_ForceSave                 As Boolean
Private m_ProductKeys(8)            As String
Private m_DebugConfig               As Boolean


'[MAIN]
Private m_DisableNews As Boolean

'[CLIENT]
Private m_Username                  As String
Private m_Password                  As String
Private m_CDKey                     As String
Private m_EXPKey                    As String
Private m_UseSpawn                  As Boolean
Private m_Game                      As String
Private m_Server                    As String
Private m_HomeChannel               As String
Private m_AutoConnect               As Boolean
Private m_UseD2Realms               As Boolean
Private m_UseBNLS                   As Boolean
Private m_BNLSServer                As String
Private m_UseBNLSFinder             As Boolean
Private m_BNLSFinderSource          As String
Private m_UseProxy                  As Boolean
Private m_ProxyIP                   As String
Private m_ProxyPort                 As Integer
Private m_ProxyType                 As String

'[FEATURES]
Private m_UseBackupChannel          As Boolean
Private m_BackupChannel             As String
Private m_ReconnectDelay            As Long
Private m_BotMail                   As Boolean
Private m_ProfileAmp                As Boolean
Private m_VoidView                  As Boolean
Private m_GreetMessage              As Boolean
Private m_GreetMessageText          As String
Private m_WhisperGreet              As Boolean
Private m_IdleMessage               As Boolean
Private m_IdleMessageText           As String
Private m_IdleMessageDelay          As Integer
Private m_IdleMessageType           As String
Private m_Trigger                   As String
Private m_BotOwner                  As String
Private m_ChatFilters               As Boolean
Private m_WhisperWindows            As Boolean
Private m_WhisperCommands           As Boolean
Private m_ChatDelay                 As Long
Private m_MediaPlayer               As String
Private m_MediaPlayerPath           As String
Private m_Mp3Commands               As Boolean
Private m_AutocompletePostfix       As String
Private m_CaseSensitiveDBFlags      As Boolean
Private m_MultiLinePostfix          As String

'[MODERATION]
Private m_BanEvasion                As Boolean
Private m_Phrasebans                As Boolean
Private m_Phrasekick                As Boolean
Private m_LevelBanW3                As Integer
Private m_LevelBanD2                As Integer
Private m_LevelBanMessage           As String
Private m_PeonBan                   As Boolean
Private m_KickOnYell                As Boolean
Private m_ShitlistGroup             As String
Private m_TagbanGroup               As String
Private m_SafelistGroup             As String
Private m_RetainOldBans             As Boolean
Private m_StoreAllBans              As Boolean
Private m_ChannelProtectionMessage  As String
Private m_IdleBan                   As Boolean
Private m_IdleBanDelay              As Integer
Private m_IdleBanKick               As Boolean
Private m_IPBans                    As Boolean
Private m_UDPBan                    As Boolean
Private m_AutoSafelistLevel         As Integer
Private m_ChannelProtection         As Boolean
Private m_QuietTime                 As Boolean
Private m_QuietTimeKick             As Boolean
Private m_PingBan                   As Boolean
Private m_PingBanLevel              As Long

'[UI]
Private m_ShowSplashScreen          As Boolean
Private m_ShowWhisperBox            As Boolean
Private m_MinimizeOnStartup         As Boolean
Private m_UseUTF8                   As Boolean
Private m_UrlDetection              As Boolean
Private m_ShowOutgoingWhispers      As Boolean
Private m_HideWhispersInMain        As Boolean
Private m_TimestampMode             As Byte
Private m_ChatFont                  As String
Private m_ChatFontSize              As Integer
Private m_ChannelListFont           As String
Private m_ChannelListFontSize       As Integer
Private m_HideClanDisplay           As Boolean
Private m_HidePingDisplay           As Boolean
Private m_NamespaceConvention       As Byte
Private m_UseD2Naming               As Boolean
Private m_ShowStatsIcons            As Boolean
Private m_ShowFlagIcons             As Boolean
Private m_ShowJoinLeaves            As Boolean
Private m_FlashOnEvents             As Boolean
Private m_FlashOnCatchPhrases       As Boolean
Private m_MinimizeToTray            As Boolean
Private m_NameColoring              As Boolean
Private m_ShowOfflineFriends        As Boolean
Private m_DisablePrefixBox          As Boolean
Private m_DisableSuffixBox          As Boolean
Private m_MathAllowUI               As Boolean
Private m_D2NamingFormat            As String
Private m_SecondsToIdle             As Long
Private m_DisableRTBAutoCopy        As Boolean
Private m_HideBanMessages           As Boolean
Private m_NameAutoComplete          As Boolean

'[UI_POS]
Private m_PositionLeft              As Long
Private m_PositionTop               As Long
Private m_PositionHeight            As Long
Private m_PositionWidth             As Long
Private m_IsMaximized               As Boolean
Private m_LastSettingsPanel         As Byte

'[LOGGING]
Private m_LogDBActions              As Boolean
Private m_LogCommands               As Boolean
Private m_MaxBacklogSize            As Long
Private m_MaxLogFileSize            As Long
Private m_LoggingMode               As Byte

'[QUEUE]
Private m_QueueMaxCredits           As Long
Private m_QueueCostPerPacket        As Long
Private m_QueueCostPerByte          As Long
Private m_QueueCostPerByteOver      As Long
Private m_QueueStartingCredits      As Long
Private m_QueueThresholdBytes       As Long
Private m_QueueCreditRate           As Long

'[SCRIPTING]
Private m_DisableScripting          As Boolean
Private m_ScriptingAllowUI          As Boolean
Private m_ScriptViewer              As String

'[EMULATION]
Private m_IgnoreClanInvites         As Boolean
Private m_IgnoreCDKeyLength         As Boolean
Private m_DoNotUseDirectFriendList  As Boolean
Private m_PingSpoofing              As Byte
Private m_UseUDP                    As Boolean
Private m_CustomStatstring          As String
Private m_ForceDefaultLocaleID      As Boolean
Private m_CustomUDPString           As String
Private m_CDKeyOwnerName            As String
Private m_UseLowerCasePassword      As Boolean
Private m_IgnoreVersionCheck        As Boolean
Private m_PredefinedGateway         As String
Private m_DefaultChannelJoin        As Boolean
Private m_MaxMessageLength          As Integer
Private m_AutoCreateChannels        As String
Private m_RegisterEmailAction       As String
Private m_RegisterEmailDefault      As String
Private m_VersionBytes(8)           As Long     'XXVerByte
Private m_LogonSystems(8)           As Long     'XXLogonSystem

'[DEBUG]
Private m_DebugWarden               As Boolean



'-------------------------
'   SECTION: MAIN
'-------------------------

Public Property Get DisableNews() As Boolean
    DisableNews = m_DisableNews
End Property

Public Property Let DisableNews(ByVal sValue As Boolean)
    m_DisableNews = sValue
End Property


'-------------------------
'   SECTION: CLIENT
'-------------------------

Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let Username(ByVal sValue As String)
    m_Username = sValue
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Password(ByVal sValue As String)
    m_Password = sValue
End Property

Public Property Get CDKey() As String
    CDKey = m_CDKey
End Property

Public Property Let CDKey(ByVal sValue As String)
    m_CDKey = sValue
End Property

Public Property Get ExpKey() As String
    ExpKey = m_EXPKey
End Property

Public Property Let ExpKey(ByVal sValue As String)
    m_EXPKey = sValue
End Property

Public Property Get UseSpawn() As Boolean
    UseSpawn = m_UseSpawn
End Property

Public Property Let UseSpawn(ByVal sValue As Boolean)
    m_UseSpawn = sValue
End Property

Public Property Get Game() As String
    Game = m_Game
End Property

Public Property Let Game(ByVal sValue As String)
    m_Game = sValue
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Let Server(ByVal sValue As String)
    m_Server = sValue
End Property

Public Property Get HomeChannel() As String
    HomeChannel = m_HomeChannel
End Property

Public Property Let HomeChannel(ByVal sValue As String)
    m_HomeChannel = sValue
End Property

Public Property Get AutoConnect() As Boolean
    AutoConnect = m_AutoConnect
End Property

Public Property Let AutoConnect(ByVal sValue As Boolean)
    m_AutoConnect = sValue
End Property

Public Property Get UseD2Realms() As Boolean
    UseD2Realms = m_UseD2Realms
End Property

Public Property Let UseD2Realms(ByVal sValue As Boolean)
    m_UseD2Realms = sValue
End Property

Public Property Get UseBNLS() As Boolean
    UseBNLS = m_UseBNLS
End Property

Public Property Let UseBNLS(ByVal sValue As Boolean)
    m_UseBNLS = sValue
End Property

Public Property Get BNLSServer() As String
    BNLSServer = m_BNLSServer
End Property

Public Property Let BNLSServer(ByVal sValue As String)
    m_BNLSServer = sValue
End Property

Public Property Get BNLSFinder() As Boolean
    BNLSFinder = m_UseBNLSFinder
End Property

Public Property Let BNLSFinder(ByVal sValue As Boolean)
    m_UseBNLSFinder = sValue
End Property

Public Property Get BNLSFinderSource() As String
    BNLSFinderSource = m_BNLSFinderSource
End Property

Public Property Let BNLSFinderSource(ByVal sValue As String)
    m_BNLSFinderSource = sValue
End Property

Public Property Get UseProxy() As Boolean
    UseProxy = m_UseProxy
End Property

Public Property Let UseProxy(ByVal sValue As Boolean)
    m_UseProxy = sValue
End Property

Public Property Get ProxyIP() As String
    ProxyIP = m_ProxyIP
End Property

Public Property Let ProxyIP(ByVal sValue As String)
    m_ProxyIP = sValue
End Property

Public Property Get ProxyPort() As Integer
    ProxyPort = m_ProxyPort
End Property

Public Property Let ProxyPort(ByVal sValue As Integer)
    m_ProxyPort = sValue
End Property

Public Property Get ProxyType() As String
    ProxyType = m_ProxyType
End Property

Public Property Let ProxyType(ByVal sValue As String)
    m_ProxyType = sValue
End Property


'-------------------------
'   SECTION: FEATURES
'-------------------------

Public Property Get UseBackupChannel() As Boolean
    UseBackupChannel = m_UseBackupChannel
End Property

Public Property Let UseBackupChannel(ByVal sValue As Boolean)
    m_UseBackupChannel = sValue
End Property

Public Property Get BackupChannel() As String
    BackupChannel = m_BackupChannel
End Property

Public Property Let BackupChannel(ByVal sValue As String)
    m_BackupChannel = sValue
End Property

Public Property Get ReconnectDelay() As Long
    ReconnectDelay = m_ReconnectDelay
End Property

Public Property Let ReconnectDelay(ByVal sValue As Long)
    m_ReconnectDelay = sValue
End Property

Public Property Get BotMail() As Boolean
    BotMail = m_BotMail
End Property

Public Property Let BotMail(ByVal sValue As Boolean)
    m_BotMail = sValue
End Property

Public Property Get ProfileAmp() As Boolean
    ProfileAmp = m_ProfileAmp
End Property

Public Property Let ProfileAmp(ByVal sValue As Boolean)
    m_ProfileAmp = sValue
End Property

Public Property Get VoidView() As Boolean
    VoidView = m_VoidView
End Property

Public Property Let VoidView(ByVal sValue As Boolean)
    m_VoidView = sValue
End Property

Public Property Get GreetMessage() As Boolean
    GreetMessage = m_GreetMessage
End Property

Public Property Let GreetMessage(ByVal sValue As Boolean)
    m_GreetMessage = sValue
End Property

Public Property Get GreetMessageText() As String
    GreetMessageText = m_GreetMessageText
End Property

Public Property Let GreetMessageText(ByVal sValue As String)
    m_GreetMessageText = sValue
End Property

Public Property Get WhisperGreet() As Boolean
    WhisperGreet = m_WhisperGreet
End Property

Public Property Let WhisperGreet(ByVal sValue As Boolean)
    m_WhisperGreet = sValue
End Property

Public Property Get IdleMessage() As Boolean
    IdleMessage = m_IdleMessage
End Property

Public Property Let IdleMessage(ByVal sValue As Boolean)
    m_IdleMessage = sValue
End Property

Public Property Get IdleMessageText() As String
    IdleMessageText = m_IdleMessageText
End Property

Public Property Let IdleMessageText(ByVal sValue As String)
    m_IdleMessageText = sValue
End Property

Public Property Get IdleMessageDelay() As Integer
    IdleMessageDelay = m_IdleMessageDelay
End Property

Public Property Let IdleMessageDelay(ByVal sValue As Integer)
    m_IdleMessageDelay = sValue
End Property

Public Property Get IdleMessageType() As String
    IdleMessageType = m_IdleMessageType
End Property

Public Property Let IdleMessageType(ByVal sValue As String)
    m_IdleMessageType = sValue
End Property

Public Property Get Trigger() As String
    Trigger = GetProtectedString(m_Trigger)
End Property

Public Property Let Trigger(ByVal sValue As String)
    m_Trigger = "{" & sValue & "}"
End Property

Public Property Get BotOwner() As String
    BotOwner = m_BotOwner
End Property

Public Property Let BotOwner(ByVal sValue As String)
    m_BotOwner = sValue
End Property

Public Property Get ChatFilters() As Boolean
    ChatFilters = m_ChatFilters
End Property

Public Property Let ChatFilters(ByVal sValue As Boolean)
    m_ChatFilters = sValue
End Property

Public Property Get WhisperWindows() As Boolean
    WhisperWindows = m_WhisperWindows
End Property

Public Property Let WhisperWindows(ByVal sValue As Boolean)
    m_WhisperWindows = sValue
End Property

Public Property Get WhisperCommands() As Boolean
    WhisperCommands = m_WhisperCommands
End Property

Public Property Let WhisperCommands(ByVal sValue As Boolean)
    m_WhisperCommands = sValue
End Property

Public Property Get ChatDelay() As Long
    ChatDelay = m_ChatDelay
End Property

Public Property Let ChatDelay(ByVal sValue As Long)
    m_ChatDelay = sValue
End Property

Public Property Get MediaPlayer() As String
    MediaPlayer = m_MediaPlayer
End Property

Public Property Let MediaPlayer(ByVal sValue As String)
    m_MediaPlayer = sValue
End Property

Public Property Get MediaPlayerPath() As String
    MediaPlayerPath = m_MediaPlayerPath
End Property

Public Property Let MediaPlayerPath(ByVal sValue As String)
    m_MediaPlayerPath = sValue
End Property

Public Property Get Mp3Commands() As Boolean
    Mp3Commands = m_Mp3Commands
End Property

Public Property Let Mp3Commands(ByVal sValue As Boolean)
    m_Mp3Commands = sValue
End Property

Public Property Get AutoCompletePostfix() As String
    AutoCompletePostfix = GetProtectedString(m_AutocompletePostfix)
End Property

Public Property Let AutoCompletePostfix(ByVal sValue As String)
    m_AutocompletePostfix = "{" & sValue & "}"
End Property

Public Property Get CaseSensitiveDBFlags() As Boolean
    CaseSensitiveDBFlags = m_CaseSensitiveDBFlags
End Property

Public Property Let CaseSensitiveDBFlags(ByVal sValue As Boolean)
    m_CaseSensitiveDBFlags = sValue
End Property

Public Property Get MultiLinePostfix() As String
    MultiLinePostfix = GetProtectedString(m_MultiLinePostfix)
End Property

Public Property Let MultiLinePostfix(ByVal sValue As String)
    m_MultiLinePostfix = "{" & sValue & "}"
End Property


'-------------------------
'   SECTION: MODERATION
'-------------------------

Public Property Get BanEvasion() As Boolean
    BanEvasion = m_BanEvasion
End Property

Public Property Let BanEvasion(ByVal sValue As Boolean)
    m_BanEvasion = sValue
End Property

Public Property Get PhraseBans() As Boolean
    PhraseBans = m_Phrasebans
End Property

Public Property Let PhraseBans(ByVal sValue As Boolean)
    m_Phrasebans = sValue
End Property

Public Property Get PhraseKick() As Boolean
    PhraseKick = m_Phrasekick
End Property

Public Property Let PhraseKick(ByVal sValue As Boolean)
    m_Phrasekick = sValue
End Property

Public Property Get LevelBanW3() As Integer
    LevelBanW3 = m_LevelBanW3
End Property

Public Property Let LevelBanW3(ByVal sValue As Integer)
    m_LevelBanW3 = sValue
End Property

Public Property Get LevelBanD2() As Integer
    LevelBanD2 = m_LevelBanD2
End Property

Public Property Let LevelBanD2(ByVal sValue As Integer)
    m_LevelBanD2 = sValue
End Property

Public Property Get LevelBanMessage() As String
    LevelBanMessage = m_LevelBanMessage
End Property

Public Property Let LevelBanMessage(ByVal sValue As String)
    m_LevelBanMessage = sValue
End Property

Public Property Get PeonBan() As Boolean
    PeonBan = m_PeonBan
End Property

Public Property Let PeonBan(ByVal sValue As Boolean)
    m_PeonBan = sValue
End Property

Public Property Get KickOnYell() As Boolean
    KickOnYell = m_KickOnYell
End Property

Public Property Let KickOnYell(ByVal sValue As Boolean)
    m_KickOnYell = sValue
End Property

Public Property Get ShitlistGroup() As String
    ShitlistGroup = m_ShitlistGroup
End Property

Public Property Let ShitlistGroup(ByVal sValue As String)
    m_ShitlistGroup = sValue
End Property

Public Property Get TagbanGroup() As String
    TagbanGroup = m_TagbanGroup
End Property

Public Property Let TagbanGroup(ByVal sValue As String)
    m_TagbanGroup = sValue
End Property

Public Property Get SafelistGroup() As String
    SafelistGroup = m_SafelistGroup
End Property

Public Property Let SafelistGroup(ByVal sValue As String)
    m_SafelistGroup = sValue
End Property

Public Property Get RetainOldBans() As Boolean
    RetainOldBans = m_RetainOldBans
End Property

Public Property Let RetainOldBans(ByVal sValue As Boolean)
    m_RetainOldBans = sValue
End Property

Public Property Get StoreAllBans() As Boolean
    StoreAllBans = m_StoreAllBans
End Property

Public Property Let StoreAllBans(ByVal sValue As Boolean)
    m_StoreAllBans = sValue
End Property

Public Property Get ChannelProtectionMessage() As String
    ChannelProtectionMessage = m_ChannelProtectionMessage
End Property

Public Property Let ChannelProtectionMessage(ByVal sValue As String)
    m_ChannelProtectionMessage = sValue
End Property

Public Property Get IdleBan() As Boolean
    IdleBan = m_IdleBan
End Property

Public Property Let IdleBan(ByVal sValue As Boolean)
    m_IdleBan = sValue
End Property

Public Property Get IdleBanDelay() As Integer
    IdleBanDelay = m_IdleBanDelay
End Property

Public Property Let IdleBanDelay(ByVal sValue As Integer)
    m_IdleBanDelay = sValue
End Property

Public Property Get IdleBanKick() As Boolean
    IdleBanKick = m_IdleBanKick
End Property

Public Property Let IdleBanKick(ByVal sValue As Boolean)
    m_IdleBanKick = sValue
End Property

Public Property Get IPBans() As Boolean
    IPBans = m_IPBans
End Property

Public Property Let IPBans(ByVal sValue As Boolean)
    m_IPBans = sValue
End Property

Public Property Get UDPBan() As Boolean
    UDPBan = m_UDPBan
End Property

Public Property Let UDPBan(ByVal sValue As Boolean)
    m_UDPBan = sValue
End Property

Public Property Get AutoSafelistLevel() As Integer
    AutoSafelistLevel = m_AutoSafelistLevel
End Property

Public Property Let AutoSafelistLevel(ByVal sValue As Integer)
    m_AutoSafelistLevel = sValue
End Property

Public Property Get ChannelProtection() As Boolean
    ChannelProtection = m_ChannelProtection
End Property

Public Property Let ChannelProtection(ByVal sValue As Boolean)
    m_ChannelProtection = sValue
End Property

Public Property Get QuietTime() As Boolean
    QuietTime = m_QuietTime
End Property

Public Property Let QuietTime(ByVal sValue As Boolean)
    m_QuietTime = sValue
End Property

Public Property Get QuietTimeKick() As Boolean
    QuietTimeKick = m_QuietTimeKick
End Property

Public Property Let QuietTimeKick(ByVal sValue As Boolean)
    m_QuietTimeKick = sValue
End Property

Public Property Get PingBan() As Boolean
    PingBan = m_PingBan
End Property

Public Property Let PingBan(ByVal sValue As Boolean)
    m_PingBan = sValue
End Property

Public Property Get PingBanLevel() As Long
    PingBanLevel = m_PingBanLevel
End Property

Public Property Let PingBanLevel(ByVal sValue As Long)
    m_PingBanLevel = sValue
End Property


'-------------------------
'   SECTION: UI
'-------------------------

Public Property Get ShowSplashScreen() As Boolean
    ShowSplashScreen = m_ShowSplashScreen
End Property

Public Property Let ShowSplashScreen(ByVal sValue As Boolean)
    m_ShowSplashScreen = sValue
End Property

Public Property Get ShowWhisperBox() As Boolean
    ShowWhisperBox = m_ShowWhisperBox
End Property

Public Property Let ShowWhisperBox(ByVal sValue As Boolean)
    m_ShowWhisperBox = sValue
End Property

Public Property Get MinimizeOnStartup() As Boolean
    MinimizeOnStartup = m_MinimizeOnStartup
End Property

Public Property Let MinimizeOnStartup(ByVal sValue As Boolean)
    m_MinimizeOnStartup = sValue
End Property

Public Property Get UseUTF8() As Boolean
    UseUTF8 = m_UseUTF8
End Property

Public Property Let UseUTF8(ByVal sValue As Boolean)
    m_UseUTF8 = sValue
End Property

Public Property Get UrlDetection() As Boolean
    UrlDetection = m_UrlDetection
End Property

Public Property Let UrlDetection(ByVal sValue As Boolean)
    m_UrlDetection = sValue
End Property

Public Property Get ShowOutgoingWhispers() As Boolean
    ShowOutgoingWhispers = m_ShowOutgoingWhispers
End Property

Public Property Let ShowOutgoingWhispers(ByVal sValue As Boolean)
    m_ShowOutgoingWhispers = sValue
End Property

Public Property Get HideWhispersInMain() As Boolean
    HideWhispersInMain = m_HideWhispersInMain
End Property

Public Property Let HideWhispersInMain(ByVal sValue As Boolean)
    m_HideWhispersInMain = sValue
End Property

Public Property Get TimestampMode() As Byte
    TimestampMode = m_TimestampMode
End Property

Public Property Let TimestampMode(ByVal sValue As Byte)
    m_TimestampMode = sValue
End Property

Public Property Get ChatFont() As String
    ChatFont = m_ChatFont
End Property

Public Property Let ChatFont(ByVal sValue As String)
    m_ChatFont = sValue
End Property

Public Property Get ChatFontSize() As Integer
    ChatFontSize = m_ChatFontSize
End Property

Public Property Let ChatFontSize(ByVal sValue As Integer)
    m_ChatFontSize = sValue
End Property

Public Property Get ChannelListFont() As String
    ChannelListFont = m_ChannelListFont
End Property

Public Property Let ChannelListFont(ByVal sValue As String)
    m_ChannelListFont = sValue
End Property

Public Property Get ChannelListFontSize() As Integer
    ChannelListFontSize = m_ChannelListFontSize
End Property

Public Property Let ChannelListFontSize(ByVal sValue As Integer)
    m_ChannelListFontSize = sValue
End Property

Public Property Get HideClanDisplay() As Boolean
    HideClanDisplay = m_HideClanDisplay
End Property

Public Property Let HideClanDisplay(ByVal sValue As Boolean)
    m_HideClanDisplay = sValue
End Property

Public Property Get HidePingDisplay() As Boolean
    HidePingDisplay = m_HidePingDisplay
End Property

Public Property Let HidePingDisplay(ByVal sValue As Boolean)
    m_HidePingDisplay = sValue
End Property

Public Property Get NamespaceConvention() As Byte
    NamespaceConvention = m_NamespaceConvention
End Property

Public Property Let NamespaceConvention(ByVal sValue As Byte)
    m_NamespaceConvention = sValue
End Property

Public Property Get UseD2Naming() As Boolean
    UseD2Naming = m_UseD2Naming
End Property

Public Property Let UseD2Naming(ByVal sValue As Boolean)
    m_UseD2Naming = sValue
End Property

Public Property Get ShowStatsIcons() As Boolean
    ShowStatsIcons = m_ShowStatsIcons
End Property

Public Property Let ShowStatsIcons(ByVal sValue As Boolean)
    m_ShowStatsIcons = sValue
End Property

Public Property Get ShowFlagIcons() As Boolean
    ShowFlagIcons = m_ShowFlagIcons
End Property

Public Property Let ShowFlagIcons(ByVal sValue As Boolean)
    m_ShowFlagIcons = sValue
End Property

Public Property Get ShowJoinLeaves() As Boolean
    ShowJoinLeaves = m_ShowJoinLeaves
End Property

Public Property Let ShowJoinLeaves(ByVal sValue As Boolean)
    m_ShowJoinLeaves = sValue
End Property

Public Property Get FlashOnEvents() As Boolean
    FlashOnEvents = m_FlashOnEvents
End Property

Public Property Let FlashOnEvents(ByVal sValue As Boolean)
    m_FlashOnEvents = sValue
End Property

Public Property Get FlashOnCatchPhrases() As Boolean
    FlashOnCatchPhrases = m_FlashOnCatchPhrases
End Property

Public Property Let FlashOnCatchPhrases(ByVal sValue As Boolean)
    m_FlashOnCatchPhrases = sValue
End Property

Public Property Get MinimizeToTray() As Boolean
    MinimizeToTray = m_MinimizeToTray
End Property

Public Property Let MinimizeToTray(ByVal sValue As Boolean)
    m_MinimizeToTray = sValue
End Property

Public Property Get NameColoring() As Boolean
    NameColoring = m_NameColoring
End Property

Public Property Let NameColoring(ByVal sValue As Boolean)
    m_NameColoring = sValue
End Property

Public Property Get ShowOfflineFriends() As Boolean
    ShowOfflineFriends = m_ShowOfflineFriends
End Property

Public Property Let ShowOfflineFriends(ByVal sValue As Boolean)
    m_ShowOfflineFriends = sValue
End Property

Public Property Get DisablePrefixBox() As Boolean
    DisablePrefixBox = m_DisablePrefixBox
End Property

Public Property Let DisablePrefixBox(ByVal sValue As Boolean)
    m_DisablePrefixBox = sValue
End Property

Public Property Get DisableSuffixBox() As Boolean
    DisableSuffixBox = m_DisableSuffixBox
End Property

Public Property Let DisableSuffixBox(ByVal sValue As Boolean)
    m_DisableSuffixBox = sValue
End Property

Public Property Get MathAllowUI() As Boolean
    MathAllowUI = m_MathAllowUI
End Property

Public Property Let MathAllowUI(ByVal sValue As Boolean)
    m_MathAllowUI = sValue
End Property

Public Property Get D2NamingFormat() As String
    D2NamingFormat = m_D2NamingFormat
End Property

Public Property Let D2NamingFormat(ByVal sValue As String)
    m_D2NamingFormat = sValue
End Property

Public Property Get SecondsToIdle() As Long
    SecondsToIdle = m_SecondsToIdle
End Property

Public Property Let SecondsToIdle(ByVal sValue As Long)
    m_SecondsToIdle = sValue
End Property

Public Property Get DisableRTBAutoCopy() As Boolean
    DisableRTBAutoCopy = m_DisableRTBAutoCopy
End Property

Public Property Let DisableRTBAutoCopy(ByVal sValue As Boolean)
    m_DisableRTBAutoCopy = sValue
End Property

Public Property Get HideBanMessages() As Boolean
    HideBanMessages = m_HideBanMessages
End Property

Public Property Let HideBanMessages(ByVal sValue As Boolean)
    m_HideBanMessages = sValue
End Property

Public Property Get NameAutoComplete() As Boolean
    NameAutoComplete = m_NameAutoComplete
End Property

Public Property Let NameAutoComplete(ByVal sValue As Boolean)
    m_NameAutoComplete = sValue
End Property


'-------------------------
'   SECTION: UI
'-------------------------

Public Property Get PositionLeft() As Long
    PositionLeft = m_PositionLeft
End Property

Public Property Let PositionLeft(ByVal sValue As Long)
    m_PositionLeft = sValue
End Property

Public Property Get PositionTop() As Long
    PositionTop = m_PositionTop
End Property

Public Property Let PositionTop(ByVal sValue As Long)
    m_PositionTop = sValue
End Property

Public Property Get PositionHeight() As Long
    PositionHeight = m_PositionHeight
End Property

Public Property Let PositionHeight(ByVal sValue As Long)
    m_PositionHeight = sValue
End Property

Public Property Get PositionWidth() As Long
    PositionWidth = m_PositionWidth
End Property

Public Property Let PositionWidth(ByVal sValue As Long)
    m_PositionWidth = sValue
End Property

Public Property Get IsMaximized() As Boolean
    IsMaximized = m_IsMaximized
End Property

Public Property Let IsMaximized(ByVal sValue As Boolean)
    m_IsMaximized = sValue
End Property

Public Property Get LastSettingsPanel() As Byte
    LastSettingsPanel = m_LastSettingsPanel
End Property

Public Property Let LastSettingsPanel(ByVal sValue As Byte)
    m_LastSettingsPanel = sValue
End Property


'-------------------------
'   SECTION: LOGGING
'-------------------------

Public Property Get LogDBActions() As Boolean
    LogDBActions = m_LogDBActions
End Property

Public Property Let LogDBActions(ByVal sValue As Boolean)
    m_LogDBActions = sValue
End Property

Public Property Get LogCommands() As Boolean
    LogCommands = m_LogCommands
End Property

Public Property Let LogCommands(ByVal sValue As Boolean)
    m_LogCommands = sValue
End Property

Public Property Get MaxBacklogSize() As Long
    MaxBacklogSize = m_MaxBacklogSize
End Property

Public Property Let MaxBacklogSize(ByVal sValue As Long)
    m_MaxBacklogSize = sValue
End Property

Public Property Get MaxLogFileSize() As Long
    MaxLogFileSize = m_MaxLogFileSize
End Property

Public Property Let MaxLogFileSize(ByVal sValue As Long)
    m_MaxLogFileSize = sValue
End Property

Public Property Get LoggingMode() As Byte
    LoggingMode = m_LoggingMode
End Property

Public Property Let LoggingMode(ByVal sValue As Byte)
    m_LoggingMode = sValue
End Property


'-------------------------
'   SECTION: QUEUE
'-------------------------

Public Property Get QueueMaxCredits() As Long
    QueueMaxCredits = m_QueueMaxCredits
End Property

Public Property Let QueueMaxCredits(ByVal sValue As Long)
    m_QueueMaxCredits = sValue
End Property

Public Property Get QueueCostPerPacket() As Long
    QueueCostPerPacket = m_QueueCostPerPacket
End Property

Public Property Let QueueCostPerPacket(ByVal sValue As Long)
    m_QueueCostPerPacket = sValue
End Property

Public Property Get QueueCostPerByte() As Long
    QueueCostPerByte = m_QueueCostPerByte
End Property

Public Property Let QueueCostPerByte(ByVal sValue As Long)
    m_QueueCostPerByte = sValue
End Property

Public Property Get QueueCostPerByteOver() As Long
    QueueCostPerByteOver = m_QueueCostPerByteOver
End Property

Public Property Let QueueCostPerByteOver(ByVal sValue As Long)
    m_QueueCostPerByteOver = sValue
End Property

Public Property Get QueueStartingCredits() As Long
    QueueStartingCredits = m_QueueStartingCredits
End Property

Public Property Let QueueStartingCredits(ByVal sValue As Long)
    m_QueueStartingCredits = sValue
End Property

Public Property Get QueueThresholdBytes() As Long
    QueueThresholdBytes = m_QueueThresholdBytes
End Property

Public Property Let QueueThresholdBytes(ByVal sValue As Long)
    m_QueueThresholdBytes = sValue
End Property

Public Property Get QueueCreditRate() As Long
    QueueCreditRate = m_QueueCreditRate
End Property

Public Property Let QueueCreditRate(ByVal sValue As Long)
    m_QueueCreditRate = sValue
End Property


'-------------------------
'   SECTION: SCRIPTING
'-------------------------

Public Property Get DisableScripting() As Boolean
    DisableScripting = m_DisableScripting
End Property

Public Property Let DisableScripting(ByVal sValue As Boolean)
    m_DisableScripting = sValue
End Property

Public Property Get ScriptingAllowUI() As Boolean
    ScriptingAllowUI = m_ScriptingAllowUI
End Property

Public Property Let ScriptingAllowUI(ByVal sValue As Boolean)
    m_ScriptingAllowUI = sValue
End Property

Public Property Get ScriptViewer() As String
    ScriptViewer = m_ScriptViewer
End Property

Public Property Let ScriptViewer(ByVal sValue As String)
    m_ScriptViewer = sValue
End Property


'-------------------------
'   SECTION: EMULATION
'-------------------------

Public Property Get IgnoreClanInvites() As Boolean
    IgnoreClanInvites = m_IgnoreClanInvites
End Property

Public Property Let IgnoreClanInvites(ByVal sValue As Boolean)
    m_IgnoreClanInvites = sValue
End Property

Public Property Get IgnoreCDKeyLength() As Boolean
    IgnoreCDKeyLength = m_IgnoreCDKeyLength
End Property

Public Property Let IgnoreCDKeyLength(ByVal sValue As Boolean)
    m_IgnoreCDKeyLength = sValue
End Property

Public Property Get DoNotUseDirectFriendList() As Boolean
    DoNotUseDirectFriendList = m_DoNotUseDirectFriendList
End Property

Public Property Let DoNotUseDirectFriendList(ByVal sValue As Boolean)
    m_DoNotUseDirectFriendList = sValue
End Property

Public Property Get PingSpoofing() As Byte
    PingSpoofing = m_PingSpoofing
End Property

Public Property Let PingSpoofing(ByVal sValue As Byte)
    m_PingSpoofing = sValue
End Property

Public Property Get UseUDP() As Boolean
    UseUDP = m_UseUDP
End Property

Public Property Let UseUDP(ByVal sValue As Boolean)
    m_UseUDP = sValue
End Property

Public Property Get CustomStatstring() As String
    CustomStatstring = m_CustomStatstring
End Property

Public Property Let CustomStatstring(ByVal sValue As String)
    m_CustomStatstring = sValue
End Property

Public Property Get ForceDefaultLocaleID() As Boolean
    ForceDefaultLocaleID = m_ForceDefaultLocaleID
End Property

Public Property Let ForceDefaultLocaleID(ByVal sValue As Boolean)
    m_ForceDefaultLocaleID = sValue
End Property

Public Property Get CustomUDPString() As String
    CustomUDPString = m_CustomUDPString
End Property

Public Property Let CustomUDPString(ByVal sValue As String)
    m_CustomUDPString = sValue
End Property

Public Property Get CDKeyOwnerName() As String
    CDKeyOwnerName = m_CDKeyOwnerName
End Property

Public Property Let CDKeyOwnerName(ByVal sValue As String)
    m_CDKeyOwnerName = sValue
End Property

Public Property Get UseLowerCasePassword() As Boolean
    UseLowerCasePassword = m_UseLowerCasePassword
End Property

Public Property Let UseLowerCasePassword(ByVal sValue As Boolean)
    m_UseLowerCasePassword = sValue
End Property

Public Property Get IgnoreVersionCheck() As Boolean
    IgnoreVersionCheck = m_IgnoreVersionCheck
End Property

Public Property Let IgnoreVersionCheck(ByVal sValue As Boolean)
    m_IgnoreVersionCheck = sValue
End Property

Public Property Get PredefinedGateway() As String
    PredefinedGateway = m_PredefinedGateway
End Property

Public Property Let PredefinedGateway(ByVal sValue As String)
    m_PredefinedGateway = sValue
End Property

Public Property Get DefaultChannelJoin() As Boolean
    DefaultChannelJoin = m_DefaultChannelJoin
End Property

Public Property Let DefaultChannelJoin(ByVal sValue As Boolean)
    m_DefaultChannelJoin = sValue
End Property

Public Property Get MaxMessageLength() As Integer
    MaxMessageLength = m_MaxMessageLength
End Property

Public Property Let MaxMessageLength(ByVal sValue As Integer)
    m_MaxMessageLength = sValue
End Property

Public Property Get AutoCreateChannels() As String
    AutoCreateChannels = m_AutoCreateChannels
End Property

Public Property Let AutoCreateChannels(ByVal sValue As String)
    m_AutoCreateChannels = sValue
End Property

Public Property Get RegisterEmailAction() As String
    RegisterEmailAction = m_RegisterEmailAction
End Property

Public Property Let RegisterEmailAction(ByVal sValue As String)
    m_RegisterEmailAction = sValue
End Property

Public Property Get RegisterEmailDefault() As String
    RegisterEmailDefault = m_RegisterEmailDefault
End Property

Public Property Let RegisterEmailDefault(ByVal sValue As String)
    m_RegisterEmailDefault = sValue
End Property


'-------------------------
'   SECTION: DEBUG
'-------------------------

Public Property Get DebugWarden() As Boolean
    DebugWarden = m_DebugWarden
End Property

Public Property Let DebugWarden(ByVal sValue As Boolean)
    m_DebugWarden = sValue
End Property




' Returns true if the config file exists.
Public Property Get FileExists() As Boolean
    FileExists = Not CBool(Len(Dir(m_ConfigPath)) = 0)
End Property

' Returns the path this object was last loaded from.
Public Property Get FilePath() As String
    FilePath = m_ConfigPath
End Property

Public Property Get Version() As Long
    Version = m_ConfigVersion
End Property

Public Property Get DebugConfig() As Boolean
    DebugConfig = m_DebugConfig
End Property

Public Property Let DebugConfig(ByVal sValue As Boolean)
    m_DebugConfig = sValue
End Property

' Returns the path to the specified file if it is overriden by this configuration.
Public Function GetFilePath(ByVal sFileName As String) As String
    GetFilePath = ReadSetting(SECTION_FILEPATH, sFileName)
End Function

Public Function GetVersionByte(ByVal sProductCode As String) As Long
    GetVersionByte = m_VersionBytes(GetProductIndex(sProductCode))
End Function

Public Sub SetVersionByte(ByVal sProductCode As String, ByVal iValue As Long)
    m_VersionBytes(GetProductIndex(sProductCode)) = iValue
End Sub

Public Function GetLogonSystem(ByVal sProductCode As String) As Long
    GetLogonSystem = m_LogonSystems(GetProductIndex(sProductCode))
End Function

Public Sub SetLogonSystem(ByVal sProductCode As String, ByVal iValue As Long)
    m_LogonSystems(GetProductIndex(sProductCode)) = iValue
End Sub


' Attempts to load the specified file into the object.
Public Sub Load(ByVal sFilePath As String)
    m_ConfigPath = sFilePath
    
    ' If it doesn't exist, run the v6 load to put in default values.
    If Not FileExists Then
        Call LoadDefaults
        Exit Sub
    End If
    
    m_ConfigVersion = ReadSettingL(SECTION_MAIN, "ConfigVersion", 5)
    
    Select Case m_ConfigVersion
        Case 5: Call LoadVersion5Config
        Case 6: Call LoadVersion6Config
        Case Else:
            Call LoadDefaults
    End Select
End Sub

' Saves the config in the latest format.
Public Sub Save(Optional ByVal sFilePath As String = vbNullString)
    Dim normalPath As String
    normalPath = m_ConfigPath
    
    ' If we're saving to a different file
    If Len(sFilePath) > 0 Then
        m_ConfigPath = sFilePath
    End If
    
    ' Backup the old file if this is going to be an upgrade.
    If Version < CONFIG_VERSION Then
        Dim fso As Object
        Dim backupPath As String
        
        backupPath = m_ConfigPath & "-backup"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FileExists(m_ConfigPath) And (Not fso.FileExists(backupPath))) Then
            fso.MoveFile m_ConfigPath, backupPath
            
            Call frmChat.AddChat(RTBColors.InformationText, "Your config file is being updated. A backup of your old file will be placed in your profile directory.")
        End If
        Set fso = Nothing
        
        m_ForceSave = True
    End If
    
    Call ConformValues
    
    ' Update our version number.
    m_ConfigVersion = CONFIG_VERSION
    
    WriteSetting SECTION_MAIN, "ConfigVersion", m_ConfigVersion
    WriteSetting SECTION_MAIN, "DisableNews", m_DisableNews

    WriteSetting SECTION_CLIENT, "Username", m_Username
    WriteSetting SECTION_CLIENT, "Password", m_Password
    WriteSetting SECTION_CLIENT, "CdKey", m_CDKey
    WriteSetting SECTION_CLIENT, "ExpKey", m_EXPKey
    WriteSetting SECTION_CLIENT, "Spawn", m_UseSpawn
    WriteSetting SECTION_CLIENT, "Game", m_Game
    WriteSetting SECTION_CLIENT, "Server", m_Server
    WriteSetting SECTION_CLIENT, "HomeChannel", m_HomeChannel
    WriteSetting SECTION_CLIENT, "AutoConnect", m_AutoConnect
    WriteSetting SECTION_CLIENT, "UseRealm", m_UseD2Realms
    WriteSetting SECTION_CLIENT, "UseBNLS", m_UseBNLS
    WriteSetting SECTION_CLIENT, "BNLSServer", m_BNLSServer
    WriteSetting SECTION_CLIENT, "UseBNLSFinder", m_UseBNLSFinder
    WriteSetting SECTION_CLIENT, "BNLSFinderSource", m_BNLSFinderSource
    WriteSetting SECTION_CLIENT, "UseProxy", m_UseProxy
    WriteSetting SECTION_CLIENT, "ProxyIP", m_ProxyIP
    WriteSetting SECTION_CLIENT, "ProxyPort", m_ProxyPort
    WriteSetting SECTION_CLIENT, "ProxyType", m_ProxyType
    
    WriteSetting SECTION_FEATURES, "UseBackupChannel", m_UseBackupChannel
    WriteSetting SECTION_FEATURES, "BackupChannel", m_BackupChannel
    WriteSetting SECTION_FEATURES, "ReconnectDelay", m_ReconnectDelay
    WriteSetting SECTION_FEATURES, "BotMail", m_BotMail
    WriteSetting SECTION_FEATURES, "ProfileAmp", m_ProfileAmp
    WriteSetting SECTION_FEATURES, "VoidView", m_VoidView
    WriteSetting SECTION_FEATURES, "GreetMessage", m_GreetMessage
    WriteSetting SECTION_FEATURES, "GreetMessageText", m_GreetMessageText
    WriteSetting SECTION_FEATURES, "WhisperGreet", m_WhisperGreet
    WriteSetting SECTION_FEATURES, "IdleMessage", m_IdleMessage
    WriteSetting SECTION_FEATURES, "IdleText", m_IdleMessageText
    WriteSetting SECTION_FEATURES, "IdleDelay", m_IdleMessageDelay
    WriteSetting SECTION_FEATURES, "IdleType", m_IdleMessageType
    WriteSetting SECTION_FEATURES, "Trigger", m_Trigger
    WriteSetting SECTION_FEATURES, "BotOwner", m_BotOwner
    WriteSetting SECTION_FEATURES, "ChatFilters", m_ChatFilters
    WriteSetting SECTION_FEATURES, "WhisperWindows", m_WhisperWindows
    WriteSetting SECTION_FEATURES, "WhisperCommands", m_WhisperCommands
    WriteSetting SECTION_FEATURES, "ChatDelay", m_ChatDelay
    WriteSetting SECTION_FEATURES, "MediaPlayer", m_MediaPlayer
    WriteSetting SECTION_FEATURES, "PlayerPath", m_MediaPlayerPath
    WriteSetting SECTION_FEATURES, "AllowMP3", m_Mp3Commands
    WriteSetting SECTION_FEATURES, "NameAutoComplete", m_NameAutoComplete
    WriteSetting SECTION_FEATURES, "AutoCompletePostfix", m_AutocompletePostfix
    WriteSetting SECTION_FEATURES, "CaseSensitiveFlags", m_CaseSensitiveDBFlags
    WriteSetting SECTION_FEATURES, "MultiLinePostfix", m_MultiLinePostfix
    
    WriteSetting SECTION_MODERATION, "BanEvasion", m_BanEvasion
    WriteSetting SECTION_MODERATION, "PhraseBans", m_Phrasebans
    WriteSetting SECTION_MODERATION, "PhraseKick", m_Phrasekick
    WriteSetting SECTION_MODERATION, "LevelBanW3", m_LevelBanW3
    WriteSetting SECTION_MODERATION, "LevelBanD2", m_LevelBanD2
    WriteSetting SECTION_MODERATION, "LevelBanMessage", m_LevelBanMessage
    WriteSetting SECTION_MODERATION, "PeonBan", m_PeonBan
    WriteSetting SECTION_MODERATION, "KickOnYell", m_KickOnYell
    WriteSetting SECTION_MODERATION, "ShitlistGroup", m_ShitlistGroup
    WriteSetting SECTION_MODERATION, "TagbansGroup", m_TagbanGroup
    WriteSetting SECTION_MODERATION, "SafelistGroup", m_SafelistGroup
    WriteSetting SECTION_MODERATION, "RetainOldBans", m_RetainOldBans
    WriteSetting SECTION_MODERATION, "StoreAllBans", m_StoreAllBans
    WriteSetting SECTION_MODERATION, "ProtectMessage", m_ChannelProtectionMessage
    WriteSetting SECTION_MODERATION, "RemoveIdleUsers", m_IdleBan
    WriteSetting SECTION_MODERATION, "IdleBanDelay", m_IdleBanDelay
    WriteSetting SECTION_MODERATION, "KickIdleUsers", m_IdleBanKick
    WriteSetting SECTION_MODERATION, "IPBans", m_IPBans
    WriteSetting SECTION_MODERATION, "BanUDPPlugs", m_UDPBan
    WriteSetting SECTION_MODERATION, "AutoSafelistLevel", m_AutoSafelistLevel
    WriteSetting SECTION_MODERATION, "ChannelProtect", m_ChannelProtection
    WriteSetting SECTION_MODERATION, "QuietTime", m_QuietTime
    WriteSetting SECTION_MODERATION, "QuietTimeKick", m_QuietTimeKick
    WriteSetting SECTION_MODERATION, "PingBan", m_PingBan
    WriteSetting SECTION_MODERATION, "PingBanLevel", m_PingBanLevel
    
    WriteSetting SECTION_UI, "ShowSplashScreen", m_ShowSplashScreen
    WriteSetting SECTION_UI, "ShowWhisperWindow", m_ShowWhisperBox
    WriteSetting SECTION_UI, "MinimizeOnStartup", m_MinimizeOnStartup
    WriteSetting SECTION_UI, "UseUTF8", m_UseUTF8
    WriteSetting SECTION_UI, "DetectURLs", m_UrlDetection
    WriteSetting SECTION_UI, "ShowOutgoingWhispers", m_ShowOutgoingWhispers
    WriteSetting SECTION_UI, "HideWhispersInMain", m_HideWhispersInMain
    WriteSetting SECTION_UI, "TimestampMode", m_TimestampMode
    WriteSetting SECTION_UI, "ChatFont", m_ChatFont
    WriteSetting SECTION_UI, "ChatSize", m_ChatFontSize
    WriteSetting SECTION_UI, "ChannelFont", m_ChannelListFont
    WriteSetting SECTION_UI, "ChannelSize", m_ChannelListFontSize
    WriteSetting SECTION_UI, "HideClanDisplay", m_HideClanDisplay
    WriteSetting SECTION_UI, "HidePingDisplay", m_HidePingDisplay
    WriteSetting SECTION_UI, "NamespaceConvention", m_NamespaceConvention
    WriteSetting SECTION_UI, "UseD2Naming", m_UseD2Naming
    WriteSetting SECTION_UI, "ShowStatsIcons", m_ShowStatsIcons
    WriteSetting SECTION_UI, "ShowFlagsIcons", m_ShowFlagIcons
    WriteSetting SECTION_UI, "ShowJoinLeaves", m_ShowJoinLeaves
    WriteSetting SECTION_UI, "FlashOnEvents", m_FlashOnEvents
    WriteSetting SECTION_UI, "FlashOnCatchPhrases", m_FlashOnCatchPhrases
    WriteSetting SECTION_UI, "MinimizeToTray", m_MinimizeToTray
    WriteSetting SECTION_UI, "NameColoring", m_NameColoring
    WriteSetting SECTION_UI, "ShowOfflineFriends", m_ShowOfflineFriends
    WriteSetting SECTION_UI, "DisablePrefix", m_DisablePrefixBox
    WriteSetting SECTION_UI, "DisableSuffix", m_DisableSuffixBox
    WriteSetting SECTION_UI, "MathAllowUI", m_MathAllowUI
    WriteSetting SECTION_UI, "D2NamingFormat", m_D2NamingFormat
    WriteSetting SECTION_UI, "SecondsToIdle", m_SecondsToIdle
    WriteSetting SECTION_UI, "NoRTBAutomaticCopy", m_DisableRTBAutoCopy
    WriteSetting SECTION_UI, "HideBanMessages", m_HideBanMessages
    
    WriteSetting SECTION_UI_POS, "Left", m_PositionLeft
    WriteSetting SECTION_UI_POS, "Top", m_PositionTop
    WriteSetting SECTION_UI_POS, "Height", m_PositionHeight
    WriteSetting SECTION_UI_POS, "Width", m_PositionWidth
    WriteSetting SECTION_UI_POS, "Maximized", m_IsMaximized
    WriteSetting SECTION_UI_POS, "LastSettingsPanel", m_LastSettingsPanel
    
    WriteSetting SECTION_LOGGING, "LogDBActions", m_LogDBActions
    WriteSetting SECTION_LOGGING, "LogCommands", m_LogCommands
    WriteSetting SECTION_LOGGING, "MaxBacklogSize", m_MaxBacklogSize
    WriteSetting SECTION_LOGGING, "MaxLogFileSize", m_MaxLogFileSize
    WriteSetting SECTION_LOGGING, "LogMode", m_LoggingMode
    
    WriteSetting SECTION_QUEUE, "MaxCredits", m_QueueMaxCredits
    WriteSetting SECTION_QUEUE, "CostPerPacket", m_QueueCostPerPacket
    WriteSetting SECTION_QUEUE, "CostPerByte", m_QueueCostPerByte
    WriteSetting SECTION_QUEUE, "CostPerByteOverThreshhold", m_QueueCostPerByteOver
    WriteSetting SECTION_QUEUE, "StartingCredits", m_QueueStartingCredits
    WriteSetting SECTION_QUEUE, "ThreshholdBytes", m_QueueThresholdBytes
    WriteSetting SECTION_QUEUE, "CreditRate", m_QueueCreditRate
    
    WriteSetting SECTION_SCRIPTING, "DisableScripts", m_DisableScripting
    WriteSetting SECTION_SCRIPTING, "AllowUI", m_ScriptingAllowUI
    WriteSetting SECTION_SCRIPTING, "ScriptViewer", m_ScriptViewer
 
    WriteSetting SECTION_EMULATION, "IgnoreClanInvites", m_IgnoreClanInvites
    WriteSetting SECTION_EMULATION, "IgnoreKeyLength", m_IgnoreCDKeyLength
    WriteSetting SECTION_EMULATION, "DoNotUseDirectFList", m_DoNotUseDirectFriendList
    WriteSetting SECTION_EMULATION, "PingSpoof", m_PingSpoofing
    WriteSetting SECTION_EMULATION, "UseUDP", m_UseUDP
    WriteSetting SECTION_EMULATION, "CustomStatstring", m_CustomStatstring
    WriteSetting SECTION_EMULATION, "ForceDefaultLocaleID", m_ForceDefaultLocaleID
    WriteSetting SECTION_EMULATION, "UDPString", m_CustomUDPString
    WriteSetting SECTION_EMULATION, "KeyOwner", m_CDKeyOwnerName
    WriteSetting SECTION_EMULATION, "LowerCasePassword", m_UseLowerCasePassword
    WriteSetting SECTION_EMULATION, "IgnoreVersionCheck", m_IgnoreVersionCheck
    WriteSetting SECTION_EMULATION, "PredefinedGateway", m_PredefinedGateway
    WriteSetting SECTION_EMULATION, "ForceJoinDefaultChannel", m_DefaultChannelJoin
    WriteSetting SECTION_EMULATION, "MaxMessageLength", m_MaxMessageLength
    WriteSetting SECTION_EMULATION, "AutoCreateChannels", m_AutoCreateChannels
    WriteSetting SECTION_EMULATION, "RegisterEmailAction", m_RegisterEmailAction
    WriteSetting SECTION_EMULATION, "RegisterEmailDefault", m_RegisterEmailDefault
    
    WriteSetting SECTION_DEBUG, "Warden", m_DebugWarden
    
    Dim i As Integer
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        If m_VersionBytes(i) > 0 Then
            WriteSetting SECTION_EMULATION, m_ProductKeys(i) & "VerByte", Hex(m_VersionBytes(i))
        End If
    Next
    
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        If m_LogonSystems(i) > 0 Then
            WriteSetting SECTION_EMULATION, m_ProductKeys(i) & "LogonSystem", Hex(m_LogonSystems(i))
        End If
    Next

    ' Change back to the old path.
    m_ConfigPath = normalPath
    m_ForceSave = False
End Sub

' Sets default values
Private Sub LoadDefaults()
    '[Main]
    m_DisableNews = False

    '[Client]
    m_Username = vbNullString
    m_Password = vbNullString
    m_CDKey = vbNullString
    m_EXPKey = vbNullString
    m_UseSpawn = False
    m_Game = PRODUCT_STAR
    m_Server = "useast.battle.net"
    m_HomeChannel = vbNullString
    m_AutoConnect = False
    m_UseD2Realms = False
    m_UseBNLS = True
    m_BNLSServer = vbNullString
    m_UseBNLSFinder = True
    m_BNLSFinderSource = vbNullString
    m_UseProxy = False
    m_ProxyIP = "127.0.0.1"
    m_ProxyPort = 1000
    m_ProxyType = "SOCKS4"

    '[Features]
    m_UseBackupChannel = False
    m_BackupChannel = vbNullString
    m_ReconnectDelay = 1000
    m_BotMail = True
    m_ProfileAmp = False
    m_VoidView = False
    m_GreetMessage = False
    m_GreetMessageText = vbNullString
    m_WhisperGreet = False
    m_IdleMessage = False
    m_IdleMessageText = vbNullString
    m_IdleMessageDelay = 12
    m_IdleMessageType = "msg"
    m_Trigger = "{.}"
    m_BotOwner = vbNullString
    m_ChatFilters = True
    m_WhisperWindows = False
    m_WhisperCommands = False
    m_ChatDelay = 500
    m_MediaPlayer = "Winamp"
    m_MediaPlayerPath = vbNullString
    m_Mp3Commands = True
    m_NameAutoComplete = True
    m_AutocompletePostfix = "{, }"
    m_CaseSensitiveDBFlags = True
    m_MultiLinePostfix = "{ [more]}"

    '[Moderation]
    m_BanEvasion = True
    m_Phrasebans = False
    m_Phrasekick = False
    m_LevelBanW3 = 0
    m_LevelBanD2 = 0
    m_LevelBanMessage = vbNullString
    m_PeonBan = False
    m_KickOnYell = False
    m_ShitlistGroup = vbNullString
    m_TagbanGroup = vbNullString
    m_SafelistGroup = vbNullString
    m_RetainOldBans = True
    m_StoreAllBans = True
    m_ChannelProtectionMessage = vbNullString
    m_IdleBan = False
    m_IdleBanDelay = 300
    m_IdleBanKick = False
    m_IPBans = True
    m_UDPBan = False
    m_AutoSafelistLevel = 20
    m_ChannelProtection = False
    m_QuietTime = False
    m_QuietTimeKick = False
    m_PingBan = False
    m_PingBanLevel = 5000

    '[UI]
    m_ShowSplashScreen = True
    m_ShowWhisperBox = False
    m_MinimizeOnStartup = False
    m_UseUTF8 = True
    m_UrlDetection = True
    m_ShowOutgoingWhispers = True
    m_HideWhispersInMain = False
    m_TimestampMode = 1
    m_ChatFont = "Tahoma"
    m_ChatFontSize = 10
    m_ChannelListFont = "Tahoma"
    m_ChannelListFontSize = 10
    m_HideClanDisplay = False
    m_HidePingDisplay = False
    m_NamespaceConvention = 0
    m_UseD2Naming = False
    m_ShowStatsIcons = True
    m_ShowFlagIcons = True
    m_ShowJoinLeaves = True
    m_FlashOnEvents = False
    m_FlashOnCatchPhrases = True
    m_MinimizeToTray = True
    m_NameColoring = True
    m_ShowOfflineFriends = True
    m_DisablePrefixBox = False
    m_DisableSuffixBox = False
    m_MathAllowUI = False
    m_D2NamingFormat = vbNullString
    m_SecondsToIdle = 600
    m_DisableRTBAutoCopy = False
    m_HideBanMessages = False

    '[UI-Position]
    m_PositionLeft = -1
    m_PositionTop = -1
    m_PositionHeight = 600
    m_PositionWidth = 800
    m_IsMaximized = False
    m_LastSettingsPanel = 0

    '[Logging]
    m_LogDBActions = True
    m_LogCommands = True
    m_MaxBacklogSize = 10000
    m_MaxLogFileSize = 0
    m_LoggingMode = 2

    '[Queue]
    m_QueueMaxCredits = 600
    m_QueueCostPerPacket = 200
    m_QueueCostPerByte = 7
    m_QueueCostPerByteOver = 8
    m_QueueStartingCredits = 200
    m_QueueThresholdBytes = 200
    m_QueueCreditRate = 7

    '[Scripting]
    m_DisableScripting = False
    m_ScriptingAllowUI = True
    m_ScriptViewer = vbNullString

    '[Emulation]
    m_IgnoreClanInvites = False
    m_IgnoreCDKeyLength = False
    m_DoNotUseDirectFriendList = False
    m_PingSpoofing = 0
    m_UseUDP = False
    m_CustomStatstring = vbNullString
    m_ForceDefaultLocaleID = False
    m_CustomUDPString = vbNullString
    m_CDKeyOwnerName = vbNullString
    m_UseLowerCasePassword = True
    m_IgnoreVersionCheck = False
    m_PredefinedGateway = vbNullString
    m_DefaultChannelJoin = False
    m_MaxMessageLength = 223
    m_AutoCreateChannels = "ALWAYS"
    m_RegisterEmailAction = "PROMPT"
    m_RegisterEmailDefault = vbNullString

    '[Debug]
    m_DebugWarden = False
End Sub

' Loads the old (2.7.1) config.
Private Sub LoadVersion5Config()
    Call LoadDefaults
    
    m_ShowSplashScreen = ReadSettingB(SECTION_MAIN, "ShowSplash", m_ShowSplashScreen)
    m_ShowWhisperBox = ReadSettingB(SECTION_MAIN, "ShowWhisperWindow", m_ShowWhisperBox)
    m_AutoConnect = ReadSettingB(SECTION_MAIN, "ConnectOnStartup", m_AutoConnect)
    m_MinimizeOnStartup = ReadSettingB(SECTION_MAIN, "MinimizeOnStartup", m_MinimizeOnStartup)
    m_UseBNLSFinder = ReadSettingB(SECTION_MAIN, "UseAltBnls", m_UseBNLSFinder)
    m_IdleMessage = ReadSettingB(SECTION_MAIN, "Idles", m_IdleMessage)
    m_IdleMessageText = ReadSetting(SECTION_MAIN, "IdleMsg", m_IdleMessageText)
    m_IdleMessageDelay = ReadSettingL(SECTION_MAIN, "IdleWait", m_IdleMessageDelay)
    m_IdleMessageType = ReadSetting(SECTION_MAIN, "IdleType", m_IdleMessageType)
    m_Username = ReadSetting(SECTION_MAIN, "Username", m_Username)
    m_Password = ReadSetting(SECTION_MAIN, "Password", m_Password)
    m_CDKey = ReadSetting(SECTION_MAIN, "CdKey", m_CDKey)
    m_EXPKey = ReadSetting(SECTION_MAIN, "ExpKey", m_EXPKey)
    m_Game = ReadSetting(SECTION_MAIN, "Product", m_Game)
    m_Server = ReadSetting(SECTION_MAIN, "Server", m_Server)
    m_HomeChannel = ReadSetting(SECTION_MAIN, "HomeChan", m_HomeChannel)
    m_BotOwner = ReadSetting(SECTION_MAIN, "Owner", m_BotOwner)
    m_Trigger = ReadSetting(SECTION_MAIN, "Trigger", m_Trigger)
    m_BNLSServer = ReadSetting(SECTION_MAIN, "BnlsServer", m_BNLSServer)
    m_ShowOfflineFriends = ReadSettingB(SECTION_MAIN, "ShowOfflineFriends", m_ShowOfflineFriends)
    m_WhisperWindows = ReadSettingB(SECTION_MAIN, "UseWWs", m_WhisperWindows)
    m_WhisperCommands = ReadSettingB(SECTION_MAIN, "WhisperBack", m_WhisperCommands)
    m_UseBNLS = ReadSettingB(SECTION_MAIN, "UseBnls", m_UseBNLS)
    m_LogDBActions = ReadSettingB(SECTION_MAIN, "LogDbAction", m_LogDBActions)
    m_LogCommands = ReadSettingB(SECTION_MAIN, "LogCommands", m_LogCommands)
    m_MaxBacklogSize = ReadSettingL(SECTION_MAIN, "MaxBacklogSize", m_MaxBacklogSize)
    m_MaxLogFileSize = ReadSettingL(SECTION_MAIN, "MaxLogFileSize", m_MaxLogFileSize)
    m_DoNotUseDirectFriendList = ReadSettingB(SECTION_MAIN, "DoNotUseDirectFList", m_DoNotUseDirectFriendList)
    m_UrlDetection = ReadSettingB(SECTION_MAIN, "URLDetect", m_UrlDetection)
    m_ReconnectDelay = ReadSettingL(SECTION_MAIN, "ReconnectDelay", m_ReconnectDelay)
    m_UseBackupChannel = ReadSettingB(SECTION_MAIN, "UseBackupChan", m_UseBackupChannel)
    m_BackupChannel = ReadSetting(SECTION_MAIN, "BackupChan", m_BackupChannel)
    m_UseUTF8 = ReadSettingB(SECTION_MAIN, "UTF8", m_UseUTF8)
    m_ShowOutgoingWhispers = ReadSettingB(SECTION_MAIN, "ShowOutgoingWhispers", m_ShowOutgoingWhispers)
    m_HideWhispersInMain = ReadSettingB(SECTION_MAIN, "HideWhispersInMain", m_HideWhispersInMain)
    m_IgnoreClanInvites = ReadSettingB(SECTION_MAIN, "IgnoreClanInvitations", m_IgnoreClanInvites)
    m_PingSpoofing = ReadSettingL(SECTION_MAIN, "Spoof", m_PingSpoofing)
    m_ChannelProtection = ReadSettingB(SECTION_MAIN, "Protect", m_ChannelProtection)
    m_UseUDP = ReadSettingB(SECTION_MAIN, "UDP", m_UseUDP)
    m_QuietTime = ReadSettingB(SECTION_MAIN, "QuietTime", m_QuietTime)
    m_UseProxy = ReadSettingB(SECTION_MAIN, "UseProxy", m_UseProxy)
    m_ProxyPort = ReadSettingL(SECTION_MAIN, "ProxyPort", m_ProxyPort)
    m_ProxyType = IIf(ReadSettingB(SECTION_MAIN, "ProxyIsSocks5", False), "SOCKS5", "SOCKS4")
    m_UseD2Realms = ReadSettingB(SECTION_MAIN, "UseRealm", m_UseD2Realms)
    m_ProxyIP = ReadSetting(SECTION_MAIN, "ProxyIP", m_ProxyIP)
    m_HideBanMessages = ReadSettingB(SECTION_MAIN, "HideBanMessages", m_HideBanMessages)

    m_PositionLeft = ReadSettingL(SECTION_POSITION, "Left", m_PositionLeft)
    m_PositionTop = ReadSettingL(SECTION_POSITION, "Top", m_PositionTop)
    m_PositionHeight = ReadSettingL(SECTION_POSITION, "Height", m_PositionHeight)
    m_PositionWidth = ReadSettingL(SECTION_POSITION, "Width", m_PositionWidth)
    m_IsMaximized = ReadSettingB(SECTION_POSITION, "Maximized", m_IsMaximized)
    m_LastSettingsPanel = ReadSettingL(SECTION_POSITION, "LastSettingsPanel", m_LastSettingsPanel)

    m_ProfileAmp = ReadSettingB(SECTION_OTHER, "ProfileAmp", m_ProfileAmp)
    m_TimestampMode = ReadSettingL(SECTION_OTHER, "Timestamp", m_TimestampMode)
    m_ChatFont = ReadSetting(SECTION_OTHER, "ChatFont", m_ChatFont)
    m_ChatFontSize = ReadSettingL(SECTION_OTHER, "ChatSize", m_ChatFontSize)
    m_ChannelListFont = ReadSetting(SECTION_OTHER, "ChanFont", m_ChannelListFont)
    m_ChannelListFontSize = ReadSettingL(SECTION_OTHER, "ChanSize", m_ChannelListFontSize)
    m_ChatFilters = ReadSettingB(SECTION_OTHER, "Filters", m_ChatFilters)
    m_HideClanDisplay = ReadSettingB(SECTION_OTHER, "HideClanDisplay", m_HideClanDisplay)
    m_HidePingDisplay = ReadSettingB(SECTION_OTHER, "HidePingDisplay", m_HidePingDisplay)
    m_RetainOldBans = ReadSettingB(SECTION_OTHER, "RetainOldBans", m_RetainOldBans)
    m_StoreAllBans = ReadSettingB(SECTION_OTHER, "StoreAllBans", m_StoreAllBans)
    m_NamespaceConvention = ReadSettingL(SECTION_OTHER, "NamespaceConvention", m_NamespaceConvention)
    m_UseD2Naming = ReadSettingB(SECTION_OTHER, "UseD2Naming", m_UseD2Naming)
    m_ShowStatsIcons = ReadSettingB(SECTION_OTHER, "ShowStatsIcons", m_ShowStatsIcons)
    m_ShowFlagIcons = ReadSettingB(SECTION_OTHER, "ShowFlagsIcons", m_ShowFlagIcons)
    m_ShowJoinLeaves = ReadSettingB(SECTION_OTHER, "JoinLeaves", m_ShowJoinLeaves)
    m_BotMail = ReadSettingB(SECTION_OTHER, "Mail", m_BotMail)
    m_BanEvasion = ReadSettingB(SECTION_OTHER, "BanEvasion", m_BanEvasion)
    m_LoggingMode = ReadSettingL(SECTION_OTHER, "Logging", m_LoggingMode)
    m_Phrasebans = ReadSettingB(SECTION_OTHER, "Phrasebans", m_Phrasebans)
    m_CaseSensitiveDBFlags = ReadSettingB(SECTION_OTHER, "CaseSensitiveFlags", m_CaseSensitiveDBFlags)
    m_LevelBanW3 = ReadSettingL(SECTION_OTHER, "BanUnderLevel", m_LevelBanW3)
    m_LevelBanD2 = ReadSettingL(SECTION_OTHER, "BanD2UnderLevel", m_LevelBanD2)
    m_LevelBanMessage = ReadSetting(SECTION_OTHER, "LevelBanMsg", m_LevelBanMessage)
    m_PeonBan = ReadSettingB(SECTION_OTHER, "PeonBans", m_PeonBan)
    m_KickOnYell = ReadSettingB(SECTION_OTHER, "KickOnYell", m_KickOnYell)
    m_IdleBanDelay = ReadSettingL(SECTION_OTHER, "IdleBanDelay", m_IdleBanDelay)
    m_ShitlistGroup = ReadSetting(SECTION_OTHER, "DefaultShitlistGroup", m_ShitlistGroup)
    m_TagbanGroup = ReadSetting(SECTION_OTHER, "DefaultTagbansGroup", m_TagbanGroup)
    m_SafelistGroup = ReadSetting(SECTION_OTHER, "DefaultSafelistGroup", m_SafelistGroup)
    m_Mp3Commands = ReadSettingB(SECTION_OTHER, "AllowMP3", m_Mp3Commands)
    m_ChannelProtectionMessage = ReadSetting(SECTION_OTHER, "ProtectMsg", m_ChannelProtectionMessage)
    m_IdleBan = ReadSettingB(SECTION_OTHER, "IdleBans", m_IdleBan)
    m_IdleBanKick = ReadSettingB(SECTION_OTHER, "KickIdle", m_IdleBanKick)
    m_IPBans = ReadSettingB(SECTION_OTHER, "IPBans", m_IPBans)
    m_FlashOnEvents = ReadSettingB(SECTION_OTHER, "FlashWindow", m_FlashOnEvents)
    m_MinimizeToTray = ReadSettingB(SECTION_OTHER, "MinimizeToTray", m_MinimizeToTray)
    m_NameAutoComplete = Not ReadSettingB(SECTION_OTHER, "NoAutocomplete", Not m_NameAutoComplete)
    m_NameColoring = Not ReadSettingB(SECTION_OTHER, "NoColoring", Not m_NameColoring)
    m_VoidView = Not ReadSettingB(SECTION_OTHER, "DisableVoidView", Not m_VoidView)
    m_MediaPlayer = ReadSetting(SECTION_OTHER, "MediaPlayer", m_MediaPlayer)
    m_DisablePrefixBox = ReadSettingB(SECTION_OTHER, "DisablePrefix", m_DisablePrefixBox)
    m_DisableSuffixBox = ReadSettingB(SECTION_OTHER, "DisableSuffix", m_DisableSuffixBox)
    m_MathAllowUI = ReadSettingB(SECTION_OTHER, "MathAllowUI", m_MathAllowUI)
    m_GreetMessageText = ReadSetting(SECTION_OTHER, "GreetMSg", m_GreetMessageText)
    m_GreetMessage = ReadSettingB(SECTION_OTHER, "UseGreets", m_GreetMessage)
    m_WhisperGreet = ReadSettingB(SECTION_OTHER, "WhisperGreet", m_WhisperGreet)
    m_ChatDelay = ReadSettingL(SECTION_OTHER, "ChatDelay", m_ChatDelay)
    m_FlashOnCatchPhrases = ReadSettingB(SECTION_OTHER, "FlashOnCatchPhrases", m_FlashOnCatchPhrases)
    m_MediaPlayerPath = ReadSetting(SECTION_OTHER, "WinampPath", m_MediaPlayerPath)
    m_ScriptingAllowUI = ReadSettingB(SECTION_OTHER, "ScriptAllowUI", m_ScriptingAllowUI)
    m_UDPBan = ReadSettingB(SECTION_OTHER, "PlugBans", m_UDPBan)
    m_AutocompletePostfix = ReadSetting(SECTION_OTHER, "AutoCompletePostfix", m_AutocompletePostfix)

    m_DisableNews = ReadSettingB(SECTION_OVERRIDE, "DisableSBNews", m_DisableNews)
    m_BNLSFinderSource = ReadSetting(SECTION_OVERRIDE, "BnlsSource", m_BNLSFinderSource)
    m_MaxMessageLength = ReadSettingL(SECTION_OVERRIDE, "AddQMaxLength", m_MaxMessageLength)
    m_AutoSafelistLevel = ReadSettingL(SECTION_OVERRIDE, "AutoModerationSafelistValue", m_AutoSafelistLevel)
    m_D2NamingFormat = ReadSetting(SECTION_OVERRIDE, "D2NamingFormat", m_D2NamingFormat)
    m_SecondsToIdle = ReadSettingL(SECTION_OVERRIDE, "SecondsToIdle", m_SecondsToIdle)
    m_QueueMaxCredits = ReadSettingL(SECTION_OVERRIDE, "QueueMaxCredits", m_QueueMaxCredits)
    m_QueueCostPerPacket = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerPacket", m_QueueCostPerPacket)
    m_QueueCostPerByte = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerByte", m_QueueCostPerByte)
    m_QueueCostPerByteOver = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerByteOverThreshhold", m_QueueCostPerByteOver)
    m_QueueStartingCredits = ReadSettingL(SECTION_OVERRIDE, "QueueStartingCredits", m_QueueStartingCredits)
    m_QueueThresholdBytes = ReadSettingL(SECTION_OVERRIDE, "QueueThreshholdBytes", m_QueueThresholdBytes)
    m_QueueCreditRate = ReadSettingL(SECTION_OVERRIDE, "QueueCreditRate", m_QueueCreditRate)
    m_DisableRTBAutoCopy = ReadSettingB(SECTION_OVERRIDE, "NoRTBAutomaticCopy", m_DisableRTBAutoCopy)
    m_DisableScripting = ReadSettingB(SECTION_OVERRIDE, "DisableScripts", m_DisableScripting)
    m_IgnoreCDKeyLength = ReadSettingB(SECTION_OVERRIDE, "SetKeyIgnoreLength", m_IgnoreCDKeyLength)
    m_CustomStatstring = ReadSetting(SECTION_OVERRIDE, "SetBotStatstring", m_CustomStatstring)
    m_ForceDefaultLocaleID = ReadSettingB(SECTION_OVERRIDE, "ForceDefaultLocaleId", m_ForceDefaultLocaleID)
    m_CustomUDPString = ReadSetting(SECTION_OVERRIDE, "UdpString", m_CustomUDPString)
    m_CDKeyOwnerName = ReadSetting(SECTION_OVERRIDE, "OwnerName", m_CDKeyOwnerName)
    m_UseLowerCasePassword = ReadSettingB(SECTION_OVERRIDE, "LowerCasePassword", m_UseLowerCasePassword)
    m_IgnoreVersionCheck = ReadSettingB(SECTION_OVERRIDE, "Ignore0x51Reply", m_IgnoreVersionCheck)
    m_PredefinedGateway = ReadSetting(SECTION_OVERRIDE, "PredefinedGateway", m_PredefinedGateway)
    m_DefaultChannelJoin = ReadSettingB(SECTION_OVERRIDE, "DoDefaultChannelJoin", m_DefaultChannelJoin)
    m_ScriptViewer = ReadSetting(SECTION_OVERRIDE, "ScriptViewer", m_ScriptViewer)
    m_DebugWarden = ReadSettingB(SECTION_OVERRIDE, "WardenDebug", m_DebugWarden)
    m_UseSpawn = ReadSettingB(SECTION_OVERRIDE, "SpawnKey", m_UseSpawn)
    m_MultiLinePostfix = ReadSetting(SECTION_OVERRIDE, "AddQLinePostfix", m_MultiLinePostfix)
    m_AutoCreateChannels = ReadSetting(SECTION_OVERRIDE, "ChannelCreate", m_AutoCreateChannels)
    m_RegisterEmailAction = ReadSetting(SECTION_OVERRIDE, "RegisterEmailAction", m_RegisterEmailAction)
    m_RegisterEmailDefault = ReadSetting(SECTION_OVERRIDE, "RegisterEmailDefault", m_RegisterEmailDefault)
    
    If m_EXPKey = vbNullString Then m_EXPKey = ReadSetting(SECTION_MAIN, "LODKey", m_EXPKey)
    
    Dim i As Integer
    Dim sVal As String
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        sVal = ReadSetting(SECTION_OVERRIDE, m_ProductKeys(i) & "VerByte")
        If LenB(sVal) > 0 Then m_VersionBytes(i) = CLng(Val("&H" & sVal))
        sVal = ReadSetting(SECTION_OVERRIDE, m_ProductKeys(i) & "LogonSystem")
        If LenB(sVal) > 0 Then m_LogonSystems(i) = CLng(Val("&H" & sVal))
    Next
    
    Call ConformValues
End Sub

' Loads the version 6 config.
Private Sub LoadVersion6Config()
    Call LoadDefaults
    
    m_DisableNews = ReadSettingB(SECTION_MAIN, "DisableNews", m_DisableNews)

    m_Username = ReadSetting(SECTION_CLIENT, "Username", m_Username)
    m_Password = ReadSetting(SECTION_CLIENT, "Password", m_Password)
    m_CDKey = ReadSetting(SECTION_CLIENT, "CdKey", m_CDKey)
    m_EXPKey = ReadSetting(SECTION_CLIENT, "ExpKey", m_EXPKey)
    m_UseSpawn = ReadSettingB(SECTION_CLIENT, "Spawn", m_UseSpawn)
    m_Game = ReadSetting(SECTION_CLIENT, "Game", m_Game)
    m_Server = ReadSetting(SECTION_CLIENT, "Server", m_Server)
    m_HomeChannel = ReadSetting(SECTION_CLIENT, "HomeChannel", m_HomeChannel)
    m_AutoConnect = ReadSettingB(SECTION_CLIENT, "AutoConnect", m_AutoConnect)
    m_UseD2Realms = ReadSettingB(SECTION_CLIENT, "UseRealm", m_UseD2Realms)
    m_UseBNLS = ReadSettingB(SECTION_CLIENT, "UseBNLS", m_UseBNLS)
    m_BNLSServer = ReadSetting(SECTION_CLIENT, "BNLSServer", m_BNLSServer)
    m_UseBNLSFinder = ReadSettingB(SECTION_CLIENT, "UseBNLSFinder", m_UseBNLSFinder)
    m_BNLSFinderSource = ReadSetting(SECTION_CLIENT, "BNLSFinderSource", m_BNLSFinderSource)
    m_UseProxy = ReadSettingB(SECTION_CLIENT, "UseProxy", m_UseProxy)
    m_ProxyIP = ReadSetting(SECTION_CLIENT, "ProxyIP", m_ProxyIP)
    m_ProxyPort = ReadSettingL(SECTION_CLIENT, "ProxyPort", m_ProxyPort)
    m_ProxyType = ReadSetting(SECTION_CLIENT, "ProxyType", m_ProxyType)

    m_UseBackupChannel = ReadSettingB(SECTION_FEATURES, "UseBackupChannel", m_UseBackupChannel)
    m_BackupChannel = ReadSetting(SECTION_FEATURES, "BackupChannel", m_BackupChannel)
    m_ReconnectDelay = ReadSettingL(SECTION_FEATURES, "ReconnectDelay", m_ReconnectDelay)
    m_BotMail = ReadSettingB(SECTION_FEATURES, "BotMail", m_BotMail)
    m_ProfileAmp = ReadSettingB(SECTION_FEATURES, "ProfileAmp", m_ProfileAmp)
    m_VoidView = ReadSettingB(SECTION_FEATURES, "VoidView", m_VoidView)
    m_GreetMessage = ReadSettingB(SECTION_FEATURES, "GreetMessage", m_GreetMessage)
    m_GreetMessageText = ReadSetting(SECTION_FEATURES, "GreetMessageText", m_GreetMessageText)
    m_WhisperGreet = ReadSettingB(SECTION_FEATURES, "WhisperGreet", m_WhisperGreet)
    m_IdleMessage = ReadSettingB(SECTION_FEATURES, "IdleMessage", m_IdleMessage)
    m_IdleMessageText = ReadSetting(SECTION_FEATURES, "IdleText", m_IdleMessageText)
    m_IdleMessageDelay = ReadSettingL(SECTION_FEATURES, "IdleDelay", m_IdleMessageDelay)
    m_IdleMessageType = ReadSetting(SECTION_FEATURES, "IdleType", m_IdleMessageType)
    m_Trigger = ReadSetting(SECTION_FEATURES, "Trigger", m_Trigger)
    m_BotOwner = ReadSetting(SECTION_FEATURES, "BotOwner", m_BotOwner)
    m_ChatFilters = ReadSettingB(SECTION_FEATURES, "ChatFilters", m_ChatFilters)
    m_WhisperWindows = ReadSettingB(SECTION_FEATURES, "WhisperWindows", m_WhisperWindows)
    m_WhisperCommands = ReadSettingB(SECTION_FEATURES, "WhisperCommands", m_WhisperCommands)
    m_ChatDelay = ReadSettingL(SECTION_FEATURES, "ChatDelay", m_ChatDelay)
    m_MediaPlayer = ReadSetting(SECTION_FEATURES, "MediaPlayer", m_MediaPlayer)
    m_MediaPlayerPath = ReadSetting(SECTION_FEATURES, "PlayerPath", m_MediaPlayerPath)
    m_Mp3Commands = ReadSettingB(SECTION_FEATURES, "AllowMP3", m_Mp3Commands)
    m_NameAutoComplete = ReadSettingB(SECTION_FEATURES, "NameAutoComplete", m_NameAutoComplete)
    m_AutocompletePostfix = ReadSetting(SECTION_FEATURES, "AutoCompletePostfix", m_AutocompletePostfix)
    m_CaseSensitiveDBFlags = ReadSettingB(SECTION_FEATURES, "CaseSensitiveFlags", m_CaseSensitiveDBFlags)
    m_MultiLinePostfix = ReadSetting(SECTION_FEATURES, "MultiLinePostfix", m_MultiLinePostfix)

    m_BanEvasion = ReadSettingB(SECTION_MODERATION, "BanEvasion", m_BanEvasion)
    m_Phrasebans = ReadSettingB(SECTION_MODERATION, "PhraseBans", m_Phrasebans)
    m_Phrasekick = ReadSettingB(SECTION_MODERATION, "PhraseKick", m_Phrasekick)
    m_LevelBanW3 = ReadSettingL(SECTION_MODERATION, "LevelBanW3", m_LevelBanW3)
    m_LevelBanD2 = ReadSettingL(SECTION_MODERATION, "LevelBanD2", m_LevelBanD2)
    m_LevelBanMessage = ReadSetting(SECTION_MODERATION, "LevelBanMessage", m_LevelBanMessage)
    m_PeonBan = ReadSettingB(SECTION_MODERATION, "PeonBan", m_PeonBan)
    m_KickOnYell = ReadSettingB(SECTION_MODERATION, "KickOnYell", m_KickOnYell)
    m_ShitlistGroup = ReadSetting(SECTION_MODERATION, "ShitlistGroup", m_ShitlistGroup)
    m_TagbanGroup = ReadSetting(SECTION_MODERATION, "TagbansGroup", m_TagbanGroup)
    m_SafelistGroup = ReadSetting(SECTION_MODERATION, "SafelistGroup", m_SafelistGroup)
    m_RetainOldBans = ReadSettingB(SECTION_MODERATION, "RetainOldBans", m_RetainOldBans)
    m_StoreAllBans = ReadSettingB(SECTION_MODERATION, "StoreAllBans", m_StoreAllBans)
    m_ChannelProtectionMessage = ReadSetting(SECTION_MODERATION, "ProtectMessage", m_ChannelProtectionMessage)
    m_IdleBan = ReadSettingB(SECTION_MODERATION, "RemoveIdleUsers", m_IdleBan)
    m_IdleBanDelay = ReadSettingL(SECTION_MODERATION, "IdleBanDelay", m_IdleBanDelay)
    m_IdleBanKick = ReadSettingB(SECTION_MODERATION, "KickIdleUsers", m_IdleBanKick)
    m_IPBans = ReadSettingB(SECTION_MODERATION, "IPBans", m_IPBans)
    m_UDPBan = ReadSettingB(SECTION_MODERATION, "BanUDPPlugs", m_UDPBan)
    m_AutoSafelistLevel = ReadSettingL(SECTION_MODERATION, "AutoSafelistLevel", m_AutoSafelistLevel)
    m_ChannelProtection = ReadSettingB(SECTION_MODERATION, "ChannelProtect", m_ChannelProtection)
    m_QuietTime = ReadSettingB(SECTION_MODERATION, "QuietTime", m_QuietTime)
    m_QuietTime = ReadSettingB(SECTION_MODERATION, "QuietTimeKick", m_QuietTime)
    m_PingBan = ReadSettingB(SECTION_MODERATION, "PingBan", m_PingBan)
    m_PingBanLevel = ReadSettingL(SECTION_MODERATION, "PingBanLevel", m_PingBanLevel)
    
    m_ShowSplashScreen = ReadSettingB(SECTION_UI, "ShowSplashScreen", m_ShowSplashScreen)
    m_ShowWhisperBox = ReadSettingB(SECTION_UI, "ShowWhisperWindow", m_ShowWhisperBox)
    m_MinimizeOnStartup = ReadSettingB(SECTION_UI, "MinimizeOnStartup", m_MinimizeOnStartup)
    m_UseUTF8 = ReadSettingB(SECTION_UI, "UseUTF8", m_UseUTF8)
    m_UrlDetection = ReadSettingB(SECTION_UI, "DetectURLs", m_UrlDetection)
    m_ShowOutgoingWhispers = ReadSettingB(SECTION_UI, "ShowOutgoingWhispers", m_ShowOutgoingWhispers)
    m_HideWhispersInMain = ReadSettingB(SECTION_UI, "HideWhispersInMain", m_HideWhispersInMain)
    m_TimestampMode = ReadSettingL(SECTION_UI, "TimestampMode", m_TimestampMode)
    m_ChatFont = ReadSetting(SECTION_UI, "ChatFont", m_ChatFont)
    m_ChatFontSize = ReadSettingL(SECTION_UI, "ChatSize", m_ChatFontSize)
    m_ChannelListFont = ReadSetting(SECTION_UI, "ChannelFont", m_ChannelListFont)
    m_ChannelListFontSize = ReadSettingL(SECTION_UI, "ChannelSize", m_ChannelListFontSize)
    m_HideClanDisplay = ReadSettingB(SECTION_UI, "HideClanDisplay", m_HideClanDisplay)
    m_HidePingDisplay = ReadSettingB(SECTION_UI, "HidePingDisplay", m_HidePingDisplay)
    m_NamespaceConvention = ReadSettingL(SECTION_UI, "NamespaceConvention", m_NamespaceConvention)
    m_UseD2Naming = ReadSettingB(SECTION_UI, "UseD2Naming", m_UseD2Naming)
    m_ShowStatsIcons = ReadSettingB(SECTION_UI, "ShowStatsIcons", m_ShowStatsIcons)
    m_ShowFlagIcons = ReadSettingB(SECTION_UI, "ShowFlagsIcons", m_ShowFlagIcons)
    m_ShowJoinLeaves = ReadSettingB(SECTION_UI, "ShowJoinLeaves", m_ShowJoinLeaves)
    m_FlashOnEvents = ReadSettingB(SECTION_UI, "FlashOnEvents", m_FlashOnEvents)
    m_FlashOnCatchPhrases = ReadSettingB(SECTION_UI, "FlashOnCatchPhrases", m_FlashOnCatchPhrases)
    m_MinimizeToTray = ReadSettingB(SECTION_UI, "MinimizeToTray", m_MinimizeToTray)
    m_NameColoring = ReadSettingB(SECTION_UI, "NameColoring", m_NameColoring)
    m_ShowOfflineFriends = ReadSettingB(SECTION_UI, "ShowOfflineFriends", m_ShowOfflineFriends)
    m_DisablePrefixBox = ReadSettingB(SECTION_UI, "DisablePrefix", m_DisablePrefixBox)
    m_DisableSuffixBox = ReadSettingB(SECTION_UI, "DisableSuffix", m_DisableSuffixBox)
    m_MathAllowUI = ReadSettingB(SECTION_UI, "MathAllowUI", m_MathAllowUI)
    m_D2NamingFormat = ReadSetting(SECTION_UI, "D2NamingFormat", m_D2NamingFormat)
    m_SecondsToIdle = ReadSettingL(SECTION_UI, "SecondsToIdle", m_SecondsToIdle)
    m_DisableRTBAutoCopy = ReadSettingB(SECTION_UI, "NoRTBAutomaticCopy", m_DisableRTBAutoCopy)
    m_HideBanMessages = ReadSettingB(SECTION_UI, "HideBanMessages", m_HideBanMessages)

    m_PositionLeft = ReadSettingL(SECTION_UI_POS, "Left", m_PositionLeft)
    m_PositionTop = ReadSettingL(SECTION_UI_POS, "Top", m_PositionTop)
    m_PositionHeight = ReadSettingL(SECTION_UI_POS, "Height", m_PositionHeight)
    m_PositionWidth = ReadSettingL(SECTION_UI_POS, "Width", m_PositionWidth)
    m_IsMaximized = ReadSettingB(SECTION_UI_POS, "Maximized", m_IsMaximized)
    m_LastSettingsPanel = ReadSettingL(SECTION_UI_POS, "LastSettingsPanel", m_LastSettingsPanel)

    m_LogDBActions = ReadSettingB(SECTION_LOGGING, "LogDBActions", m_LogDBActions)
    m_LogCommands = ReadSettingB(SECTION_LOGGING, "LogCommands", m_LogCommands)
    m_MaxBacklogSize = ReadSettingL(SECTION_LOGGING, "MaxBacklogSize", m_MaxBacklogSize)
    m_MaxLogFileSize = ReadSettingL(SECTION_LOGGING, "MaxLogFileSize", m_MaxLogFileSize)
    m_LoggingMode = ReadSettingL(SECTION_LOGGING, "LogMode", m_LoggingMode)

    m_QueueMaxCredits = ReadSettingL(SECTION_QUEUE, "MaxCredits", m_QueueMaxCredits)
    m_QueueCostPerPacket = ReadSettingL(SECTION_QUEUE, "CostPerPacket", m_QueueCostPerPacket)
    m_QueueCostPerByte = ReadSettingL(SECTION_QUEUE, "CostPerByte", m_QueueCostPerByte)
    m_QueueCostPerByteOver = ReadSettingL(SECTION_QUEUE, "CostPerByteOverThreshhold", m_QueueCostPerByteOver)
    m_QueueStartingCredits = ReadSettingL(SECTION_QUEUE, "StartingCredits", m_QueueStartingCredits)
    m_QueueThresholdBytes = ReadSettingL(SECTION_QUEUE, "ThreshholdBytes", m_QueueThresholdBytes)
    m_QueueCreditRate = ReadSettingL(SECTION_QUEUE, "CreditRate", m_QueueCreditRate)

    m_DisableScripting = ReadSettingB(SECTION_SCRIPTING, "DisableScripts", m_DisableScripting)
    m_ScriptingAllowUI = ReadSettingB(SECTION_SCRIPTING, "AllowUI", m_ScriptingAllowUI)
    m_ScriptViewer = ReadSetting(SECTION_SCRIPTING, "ScriptViewer", m_ScriptViewer)

    m_IgnoreClanInvites = ReadSettingB(SECTION_EMULATION, "IgnoreClanInvites", m_IgnoreClanInvites)
    m_IgnoreCDKeyLength = ReadSettingB(SECTION_EMULATION, "IgnoreKeyLength", m_IgnoreCDKeyLength)
    m_DoNotUseDirectFriendList = ReadSettingB(SECTION_EMULATION, "DoNotUseDirectFList", m_DoNotUseDirectFriendList)
    m_PingSpoofing = ReadSettingL(SECTION_EMULATION, "PingSpoof", m_PingSpoofing)
    m_UseUDP = ReadSettingB(SECTION_EMULATION, "UseUDP", m_UseUDP)
    m_CustomStatstring = ReadSetting(SECTION_EMULATION, "CustomStatstring", m_CustomStatstring)
    m_ForceDefaultLocaleID = ReadSettingB(SECTION_EMULATION, "ForceDefaultLocaleID", m_ForceDefaultLocaleID)
    m_CustomUDPString = ReadSetting(SECTION_EMULATION, "UDPString", m_CustomUDPString)
    m_CDKeyOwnerName = ReadSetting(SECTION_EMULATION, "KeyOwner", m_CDKeyOwnerName)
    m_UseLowerCasePassword = ReadSettingB(SECTION_EMULATION, "LowerCasePassword", m_UseLowerCasePassword)
    m_IgnoreVersionCheck = ReadSettingB(SECTION_EMULATION, "IgnoreVersionCheck", m_IgnoreVersionCheck)
    m_PredefinedGateway = ReadSetting(SECTION_EMULATION, "PredefinedGateway", m_PredefinedGateway)
    m_DefaultChannelJoin = ReadSetting(SECTION_EMULATION, "ForceJoinDefaultChannel", m_DefaultChannelJoin)
    m_MaxMessageLength = ReadSettingL(SECTION_EMULATION, "MaxMessageLength", m_MaxMessageLength)
    m_AutoCreateChannels = ReadSetting(SECTION_EMULATION, "AutoCreateChannels", m_AutoCreateChannels)
    m_RegisterEmailAction = ReadSetting(SECTION_EMULATION, "RegisterEmailAction", m_RegisterEmailAction)
    m_RegisterEmailDefault = ReadSetting(SECTION_EMULATION, "RegisterEmailDefault", m_RegisterEmailDefault)

    m_DebugWarden = ReadSettingB(SECTION_DEBUG, "Warden", m_DebugWarden)

    
    Dim i As Integer
    Dim sVal As String
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        sVal = ReadSetting(SECTION_EMULATION, m_ProductKeys(i) & "VerByte")
        If LenB(sVal) > 0 Then m_VersionBytes(i) = CLng(Val("&H" & sVal))
        sVal = ReadSetting(SECTION_EMULATION, m_ProductKeys(i) & "LogonSystem")
        If LenB(sVal) > 0 Then m_LogonSystems(i) = CLng(Val("&H" & sVal))
    Next
    
    Call ConformValues
End Sub

' Ensures that values are within valid ranges and in the proper format.
Private Sub ConformValues()
    If m_TimestampMode > 4 Or m_TimestampMode < 0 Then m_TimestampMode = 0
    
    m_CDKey = UCase$(m_CDKey)
    m_EXPKey = UCase$(m_EXPKey)
    m_AutoCreateChannels = UCase$(m_AutoCreateChannels)
    m_RegisterEmailAction = UCase$(m_RegisterEmailAction)
    
    m_Game = GetProductInfo(m_Game).Code
    
    If m_AutoSafelistLevel < 1 Or m_AutoSafelistLevel > 200 Then m_AutoSafelistLevel = 20
    If m_NamespaceConvention < 0 Or m_NamespaceConvention > 3 Then m_NamespaceConvention = 0
    If m_SecondsToIdle > 1000000 Then m_SecondsToIdle = 600
    If m_IdleBanDelay > 32767 Then m_IdleBanDelay = 32767
    
    If m_MaxBacklogSize < 0 Then m_MaxBacklogSize = 10000
    If m_MaxLogFileSize < 0 Then m_MaxLogFileSize = 50000000
    
    If m_ReconnectDelay < 0 Then m_ReconnectDelay = 1000
    If m_ReconnectDelay > 60000 Then m_ReconnectDelay = 60000
    
    If m_PingSpoofing < 0 Or m_PingSpoofing > 2 Then m_PingSpoofing = 0
    If m_ProxyPort < 0 Or m_ProxyPort > 65535 Then m_ProxyPort = 0
    
    If m_QueueMaxCredits < 0 Then m_QueueMaxCredits = 600
    If m_QueueCostPerPacket < 0 Then m_QueueCostPerPacket = 200
    If m_QueueCostPerByte < 0 Then m_QueueCostPerByte = 6
    If m_QueueCostPerByteOver < 0 Then m_QueueCostPerByteOver = 7
    If m_QueueStartingCredits < 0 Then m_QueueStartingCredits = 200
    If m_QueueThresholdBytes < 0 Then m_QueueThresholdBytes = 200
    If m_QueueCreditRate < 0 Then m_QueueCreditRate = 7
    
    If m_MaxMessageLength < 1 Or m_MaxMessageLength > BNET_MSG_LENGTH Then m_MaxMessageLength = BNET_MSG_LENGTH
End Sub

Private Function GetProtectedString(ByVal sValue As String) As String
    If Len(sValue) > 2 Then
        GetProtectedString = Mid(sValue, 2, Len(sValue) - 2)
    Else
        GetProtectedString = sValue
    End If
End Function


Private Function ReadSetting(ByVal section As String, ByVal Key As String, Optional ByVal DefaultValue As String = vbNullString) As String
    Dim Buffer As String
    Dim length As Long
    
    ' If the config file doesn't exist, return the default.
    If Not FileExists() Then
        ReadSetting = DefaultValue
        Exit Function
    End If
    
    ' Create a buffer to read the value into.
    Buffer = String(255, vbNull)
    
    ' Read the value into the buffer.
    length = GetPrivateProfileString(section, Key, DefaultValue, Buffer, 255, m_ConfigPath)
    ReadSetting = Left$(Buffer, length)
    
    'MsgBox section & "->" & key & ": [" & CStr(length) & "] " & ReadSetting
End Function

' Reads a setting as a boolean
Private Function ReadSettingB(ByVal section As String, ByVal Key As String, Optional ByVal DefaultValue As Boolean = False) As Boolean
    Dim sVal As String
    
    sVal = ReadSetting(section, Key, CStr(DefaultValue))
    If sVal = "Y" Or sVal = "True" Or sVal = "1" Then
        ReadSettingB = True
    ElseIf sVal = "N" Or sVal = "False" Or sVal = "0" Then
        ReadSettingB = False
    Else
        ReadSettingB = DefaultValue
    End If
End Function

' Reads a setting as a number
Private Function ReadSettingL(ByVal section As String, ByVal Key As String, Optional ByVal DefaultValue As Long = -1) As Long
    Dim sVal As String
    
    sVal = ReadSetting(section, Key, CStr(DefaultValue))
    If IsNumeric(sVal) Then
        ReadSettingL = Val(sVal)
    Else
        ReadSettingL = DefaultValue
    End If
End Function

' Sets a setting if needed.
Private Sub WriteSetting(ByVal section As String, ByVal Key As String, ByVal Value As String)
    Dim currentVal As String
    currentVal = ReadSetting(section, Key)
    
    ' Only write if the value has changed or we are forcing a save.
    If (m_ForceSave Or (StrComp(currentVal, Value, vbBinaryCompare) <> 0)) Then
        WritePrivateProfileString section, Key, Value, m_ConfigPath
        
        If m_DebugConfig Then frmChat.AddChat RTBColors.InformationText, StringFormat("[CONFIG] Set setting: {0}.{1} -> {2}", section, Key, Value)
    End If
End Sub

Private Function GetProductIndex(ByVal Code As String) As Integer
    Dim i As Integer
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        If UCase$(Code) = m_ProductKeys(i) Then
            GetProductIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    m_DebugConfig = False
    
    ' Used for reading version byte and logon system overrides
    m_ProductKeys(0) = "W2"
    m_ProductKeys(1) = "SC"
    m_ProductKeys(2) = "D2"
    m_ProductKeys(3) = "D2X"
    m_ProductKeys(4) = "W3"
    m_ProductKeys(5) = "D1"
    m_ProductKeys(6) = "DS"
    m_ProductKeys(7) = "JS"
    m_ProductKeys(8) = "SS"
End Sub

