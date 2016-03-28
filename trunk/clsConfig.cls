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

'[Main]
Private m_ShowSplashScreen          As Boolean  'ShowSplash
Private m_ShowWhisperWindow         As Boolean
Private m_ConnectOnStartup          As Boolean
Private m_MinimizeOnStartup         As Boolean
Private m_UseBnlsFinder             As Boolean  'UseAltBnls
Private m_IdlesEnabled              As Boolean  'Idles
Private m_IdleMessage               As String   'IdleMsg
Private m_IdleDelay                 As Integer  'IdleWait
Private m_IdleType                  As String
Private m_Username                  As String
Private m_Password                  As String
Private m_CdKey                     As String
Private m_ExpKey                    As String
Private m_Product                   As String
Private m_Server                    As String
Private m_HomeChannel               As String   'HomeChan
Private m_BotOwner                  As String   'Owner
Private m_Trigger                   As String
Private m_BnlsServer                As String
Private m_ShowOfflineFriends        As Boolean
Private m_UseWhisperWindows         As Boolean  'UseWWs
Private m_WhisperResponses          As Boolean  'WhisperBack
Private m_UseBnls                   As Boolean
Private m_LogDbAction               As Boolean
Private m_LogCommands               As Boolean
Private m_MaxBacklogSize            As Long
Private m_MaxLogFileSize            As Long
Private m_DoNotUseDirectFList       As Boolean
Private m_DetectUrls                As Boolean  'URLDetect
Private m_ReconnectDelay            As Long
Private m_UseBackupChannel          As Boolean  'UseBackupChan
Private m_BackupChannel             As String   'BackupChan
Private m_UTF8                      As Boolean
Private m_ShowOutgoingWhispers      As Boolean
Private m_HideWhispersInMain        As Boolean
Private m_IgnoreClanInvitations     As Boolean
Private m_PingSpoofing              As Byte     'Spoof
Private m_ChannelProtection         As Boolean  'Protect
Private m_UseUDP                    As Boolean  'UDP
Private m_QuietTime                 As Boolean
Private m_UseProxy                  As Boolean
Private m_ProxyPort                 As Integer
Private m_ProxyIsSocks5             As Boolean
Private m_UseRealm                  As Boolean
Private m_ProxyIP                   As String
Private m_HideBanMessages           As Boolean

'[Position]
Private m_Left                      As Long
Private m_Top                       As Long
Private m_Height                    As Long
Private m_Width                     As Long
Private m_Maximized                 As Boolean
Private m_LastSettingsPanel         As Byte

'[Other]
Private m_UseProfileAmp             As Boolean  'ProfileAmp
Private m_Timestamp                 As Byte
Private m_ChatFont                  As String
Private m_ChatSize                  As Integer
Private m_ChannelFont               As String   'ChanFont
Private m_ChannelSize               As Integer  'ChanSize
Private m_UseChatFilters            As Boolean  'Filters
Private m_HideClanDisplay           As Boolean
Private m_HidePingDisplay           As Boolean
Private m_RetainOldBans             As Boolean
Private m_StoreAllBans              As Boolean
Private m_NamespaceConvention       As Byte
Private m_UseD2Naming               As Boolean
Private m_ShowStatsIcons            As Boolean
Private m_ShowFlagIcons             As Boolean  'ShowFlagsIcons
Private m_ShowJoinLeaves            As Boolean  'JoinLeaves
Private m_EnableMail                As Boolean  'Mail
Private m_EnforceBanEvasion         As Boolean  'BanEvasion
Private m_LoggingLevel              As Byte     'Logging
Private m_EnablePhrasebans          As Boolean  'Phrasebans
Private m_CaseSensitiveFlags        As Boolean
Private m_BanUnderLevel             As Integer
Private m_BanD2UnderLevel           As Integer
Private m_LevelBanMessage           As String   'LevelBanMsg
Private m_BanWc3Peons               As Boolean  'PeonBans
Private m_KickOnYell                As Boolean
Private m_IdleBanDelay              As Integer
Private m_DefaultShitlistGroup      As String
Private m_DefaultTagbanGroup        As String   'DefaultTagbansGroup
Private m_DefaultSafelistGroup      As String
Private m_AllowMp3Commands          As Boolean  'AllowMP3
Private m_ChannelProtectionMessage  As String   'ProtectMsg
Private m_EnforceIdleBans           As Boolean  'IdleBans
Private m_KickIdleUsers             As Boolean  'KickIdle
Private m_IpBans                    As Boolean  'IPBans
Private m_FlashOnEvents             As Boolean  'FlashWindow
Private m_MinimizeToTray            As Boolean  'NoTray (inverse)
Private m_DisableAutoComplete       As Boolean  'NoAutocomplete
Private m_DisableNameColoring       As Boolean  'NoColoring
Private m_DisableVoidView           As Boolean
Private m_MediaPlayer               As String
Private m_DisablePrefixBox          As Boolean  'DisablePrefix
Private m_DisableSuffixBox          As Boolean  'DisableSuffix
Private m_MathAllowUI               As Boolean
Private m_GreetMessage              As String   'GreetMSg
Private m_UseGreetMessage           As Boolean  'UseGreets
Private m_WhisperGreet              As Boolean
Private m_ChatDelay                 As Long
Private m_FlashOnCatchPhrases       As Boolean
Private m_WinampPath                As String
Private m_ScriptAllowUI             As Boolean
Private m_BanUdpPlugs               As Boolean  'PlugBans
Private m_AutoCompletePostfix       As String

'[Override]
Private m_DisableNews               As Boolean  'DisableSBNews
Private m_BnlsSource                As String
Private m_AddQMaxLength             As Integer
Private m_AutoModSafelistLevel      As Integer  'AutoModerationSafelistValue
Private m_D2NamingFormat            As String
Private m_SecondsToIdle             As Long
Private m_QueueMaxCredits           As Long
Private m_QueueCostPerPacket        As Long
Private m_QueueCostPerByte          As Long
Private m_QueueCostPerByteOver      As Long     'QueueCostPerByteOverThreshhold
Private m_QueueStartingCredits      As Long
Private m_QueueThreshholdBytes      As Long
Private m_QueueCreditRate           As Long
Private m_DisableRtbAutoCopy        As Boolean  'NoRTBAutomaticCopy
Private m_LocalIP                   As String
Private m_DisableScripts            As Boolean
Private m_SetKeyIgnoreLength        As Boolean
Private m_SetBotStatstring          As String
Private m_ForceDefaultLocaleId      As Boolean
Private m_UdpString                 As String
Private m_KeyOwnerName              As String   'OwnerName
Private m_LowerCasePassword         As Boolean
Private m_IgnoreVersionCheckResult  As Boolean  'Ignore0x51Reply
Private m_PredefinedGateway         As String
Private m_DoDefaultChannelJoin      As Boolean
Private m_ScriptViewer              As String
Private m_WardenDebug               As Boolean
Private m_UseSpawnKey               As Boolean  'SpawnKey
Private m_AddQLinePostfix           As String
Private m_CreateEmptyChannels       As String
Private m_RegisterEmailAction       As String
Private m_RegisterEmailDefault      As String

Private m_VersionBytes(8)           As Long     'XXVerByte
Private m_LogonSystems(8)           As Long     'XXLogonSystem


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


Public Property Get ShowSplashScreen() As Boolean
    ShowSplashScreen = m_ShowSplashScreen
End Property

Public Property Let ShowSplashScreen(ByVal value As Boolean)
    m_ShowSplashScreen = value
End Property

Public Property Get ShowWhisperWindow() As Boolean
    ShowWhisperWindow = m_ShowWhisperWindow
End Property

Public Property Let ShowWhisperWindow(ByVal value As Boolean)
    m_ShowWhisperWindow = value
End Property

Public Property Get ConnectOnStartup() As Boolean
    ConnectOnStartup = m_ConnectOnStartup
End Property

Public Property Let ConnectOnStartup(ByVal value As Boolean)
    m_ConnectOnStartup = value
End Property

Public Property Get MinimizeOnStartup() As Boolean
    MinimizeOnStartup = m_MinimizeOnStartup
End Property

Public Property Let MinimizeOnStartup(ByVal value As Boolean)
    m_MinimizeOnStartup = value
End Property

Public Property Get UseBnlsFinder() As Boolean
    UseBnlsFinder = m_UseBnlsFinder
End Property

Public Property Let UseBnlsFinder(ByVal value As Boolean)
    m_UseBnlsFinder = value
End Property

Public Property Get IdlesEnabled() As Boolean
    IdlesEnabled = m_IdlesEnabled
End Property

Public Property Let IdlesEnabled(ByVal value As Boolean)
    m_IdlesEnabled = value
End Property

Public Property Get IdleMessage() As String
    IdleMessage = m_IdleMessage
End Property

Public Property Let IdleMessage(ByVal value As String)
    m_IdleMessage = value
End Property

Public Property Get IdleDelay() As Integer
    IdleDelay = m_IdleDelay
End Property

Public Property Let IdleDelay(ByVal value As Integer)
    m_IdleDelay = value
End Property

Public Property Get IdleType() As String
    IdleType = m_IdleType
End Property

Public Property Let IdleType(ByVal value As String)
    m_IdleType = value
End Property

Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let Username(ByVal value As String)
    m_Username = value
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Password(ByVal value As String)
    m_Password = value
End Property

Public Property Get CdKey() As String
    CdKey = m_CdKey
End Property

Public Property Let CdKey(ByVal value As String)
    m_CdKey = value
End Property

Public Property Get ExpKey() As String
    ExpKey = m_ExpKey
End Property

Public Property Let ExpKey(ByVal value As String)
    m_ExpKey = value
End Property

Public Property Get Product() As String
    Product = m_Product
End Property

Public Property Let Product(ByVal value As String)
    m_Product = value
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Let Server(ByVal value As String)
    m_Server = value
End Property

Public Property Get HomeChannel() As String
    HomeChannel = m_HomeChannel
End Property

Public Property Let HomeChannel(ByVal value As String)
    m_HomeChannel = value
End Property

Public Property Get BotOwner() As String
    BotOwner = m_BotOwner
End Property

Public Property Let BotOwner(ByVal value As String)
    m_BotOwner = value
End Property

Public Property Get Trigger() As String
    If Len(m_Trigger) > 1 Then
        Trigger = Mid(m_Trigger, 2, Len(m_Trigger) - 2)
    Else
        Trigger = m_Trigger
    End If
End Property

Public Property Let Trigger(ByVal value As String)
    m_Trigger = "{" & value & "}"
End Property

Public Property Get BnlsServer() As String
    BnlsServer = m_BnlsServer
End Property

Public Property Let BnlsServer(ByVal value As String)
    m_BnlsServer = value
End Property

Public Property Get ShowOfflineFriends() As Boolean
    ShowOfflineFriends = m_ShowOfflineFriends
End Property

Public Property Let ShowOfflineFriends(ByVal value As Boolean)
    m_ShowOfflineFriends = value
End Property

Public Property Get UseWhisperWindows() As Boolean
    UseWhisperWindows = m_UseWhisperWindows
End Property

Public Property Let UseWhisperWindows(ByVal value As Boolean)
    m_UseWhisperWindows = value
End Property

Public Property Get WhisperResponses() As Boolean
    WhisperResponses = m_WhisperResponses
End Property

Public Property Let WhisperResponses(ByVal value As Boolean)
    m_WhisperResponses = value
End Property

Public Property Get UseBnls() As Boolean
    UseBnls = m_UseBnls
End Property

Public Property Let UseBnls(ByVal value As Boolean)
    m_UseBnls = value
End Property

Public Property Get LogDbAction() As Boolean
    LogDbAction = m_LogDbAction
End Property

Public Property Let LogDbAction(ByVal value As Boolean)
    m_LogDbAction = value
End Property

Public Property Get LogCommands() As Boolean
    LogCommands = m_LogCommands
End Property

Public Property Let LogCommands(ByVal value As Boolean)
    m_LogCommands = value
End Property

Public Property Get MaxBacklogSize() As Long
    MaxBacklogSize = m_MaxBacklogSize
End Property

Public Property Let MaxBacklogSize(ByVal value As Long)
    m_MaxBacklogSize = value
End Property

Public Property Get MaxLogFileSize() As Long
    MaxLogFileSize = m_MaxLogFileSize
End Property

Public Property Let MaxLogFileSize(ByVal value As Long)
    m_MaxLogFileSize = value
End Property

Public Property Get DoNotUseDirectFList() As Boolean
    DoNotUseDirectFList = m_DoNotUseDirectFList
End Property

Public Property Let DoNotUseDirectFList(ByVal value As Boolean)
    m_DoNotUseDirectFList = value
End Property

Public Property Get DetectUrls() As Boolean
    DetectUrls = m_DetectUrls
End Property

Public Property Let DetectUrls(ByVal value As Boolean)
    m_DetectUrls = value
End Property

Public Property Get ReconnectDelay() As Long
    ReconnectDelay = m_ReconnectDelay
End Property

Public Property Let ReconnectDelay(ByVal value As Long)
    m_ReconnectDelay = value
End Property

Public Property Get UseBackupChannel() As Boolean
    UseBackupChannel = m_UseBackupChannel
End Property

Public Property Let UseBackupChannel(ByVal value As Boolean)
    m_UseBackupChannel = value
End Property

Public Property Get BackupChannel() As String
    BackupChannel = m_BackupChannel
End Property

Public Property Let BackupChannel(ByVal value As String)
    m_BackupChannel = value
End Property

Public Property Get UTF8() As Boolean
    UTF8 = m_UTF8
End Property

Public Property Let UTF8(ByVal value As Boolean)
    m_UTF8 = value
End Property

Public Property Get ShowOutgoingWhispers() As Boolean
    ShowOutgoingWhispers = m_ShowOutgoingWhispers
End Property

Public Property Let ShowOutgoingWhispers(ByVal value As Boolean)
    m_ShowOutgoingWhispers = value
End Property

Public Property Get HideWhispersInMain() As Boolean
    HideWhispersInMain = m_HideWhispersInMain
End Property

Public Property Let HideWhispersInMain(ByVal value As Boolean)
    m_HideWhispersInMain = value
End Property

Public Property Get IgnoreClanInvitations() As Boolean
    IgnoreClanInvitations = m_IgnoreClanInvitations
End Property

Public Property Let IgnoreClanInvitations(ByVal value As Boolean)
    m_IgnoreClanInvitations = value
End Property

Public Property Get PingSpoofing() As Byte
    PingSpoofing = m_PingSpoofing
End Property

Public Property Let PingSpoofing(ByVal value As Byte)
    m_PingSpoofing = value
End Property

Public Property Get ChannelProtection() As Boolean
    ChannelProtection = m_ChannelProtection
End Property

Public Property Let ChannelProtection(ByVal value As Boolean)
    m_ChannelProtection = value
End Property

Public Property Get UseUDP() As Boolean
    UseUDP = m_UseUDP
End Property

Public Property Let UseUDP(ByVal value As Boolean)
    m_UseUDP = value
End Property

Public Property Get QuietTime() As Boolean
    QuietTime = m_QuietTime
End Property

Public Property Let QuietTime(ByVal value As Boolean)
    m_QuietTime = value
End Property

Public Property Get UseProxy() As Boolean
    UseProxy = m_UseProxy
End Property

Public Property Let UseProxy(ByVal value As Boolean)
    m_UseProxy = value
End Property

Public Property Get ProxyPort() As Integer
    ProxyPort = m_ProxyPort
End Property

Public Property Let ProxyPort(ByVal value As Integer)
    m_ProxyPort = value
End Property

Public Property Get ProxyIsSocks5() As Boolean
    ProxyIsSocks5 = m_ProxyIsSocks5
End Property

Public Property Let ProxyIsSocks5(ByVal value As Boolean)
    m_ProxyIsSocks5 = value
End Property

Public Property Get UseRealm() As Boolean
    UseRealm = m_UseRealm
End Property

Public Property Let UseRealm(ByVal value As Boolean)
    m_UseRealm = value
End Property

Public Property Get ProxyIP() As String
    ProxyIP = m_ProxyIP
End Property

Public Property Let ProxyIP(ByVal value As String)
    m_ProxyIP = value
End Property

Public Property Get Left() As Long
    Left = m_Left
End Property

Public Property Let Left(ByVal value As Long)
    m_Left = value
End Property

Public Property Get Top() As Long
    Top = m_Top
End Property

Public Property Let Top(ByVal value As Long)
    m_Top = value
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Height(ByVal value As Long)
    m_Height = value
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(ByVal value As Long)
    m_Width = value
End Property

Public Property Get Maximized() As Boolean
    Maximized = m_Maximized
End Property

Public Property Let Maximized(ByVal value As Boolean)
    m_Maximized = value
End Property

Public Property Get LastSettingsPanel() As Byte
    LastSettingsPanel = m_LastSettingsPanel
End Property

Public Property Let LastSettingsPanel(ByVal value As Byte)
    m_LastSettingsPanel = value
End Property

Public Property Get UseProfileAmp() As Boolean
    UseProfileAmp = m_UseProfileAmp
End Property

Public Property Let UseProfileAmp(ByVal value As Boolean)
    m_UseProfileAmp = value
End Property

Public Property Get Timestamp() As Byte
    Timestamp = m_Timestamp
End Property

Public Property Let Timestamp(ByVal value As Byte)
    m_Timestamp = value
End Property

Public Property Get ChatFont() As String
    ChatFont = m_ChatFont
End Property

Public Property Let ChatFont(ByVal value As String)
    m_ChatFont = value
End Property

Public Property Get ChatSize() As Integer
    ChatSize = m_ChatSize
End Property

Public Property Let ChatSize(ByVal value As Integer)
    m_ChatSize = value
End Property

Public Property Get ChannelFont() As String
    ChannelFont = m_ChannelFont
End Property

Public Property Let ChannelFont(ByVal value As String)
    m_ChannelFont = value
End Property

Public Property Get ChannelSize() As Integer
    ChannelSize = m_ChannelSize
End Property

Public Property Let ChannelSize(ByVal value As Integer)
    m_ChannelSize = value
End Property

Public Property Get UseChatFilters() As Boolean
    UseChatFilters = m_UseChatFilters
End Property

Public Property Let UseChatFilters(ByVal value As Boolean)
    m_UseChatFilters = value
End Property

Public Property Get HideClanDisplay() As Boolean
    HideClanDisplay = m_HideClanDisplay
End Property

Public Property Let HideClanDisplay(ByVal value As Boolean)
    m_HideClanDisplay = value
End Property

Public Property Get HidePingDisplay() As Boolean
    HidePingDisplay = m_HidePingDisplay
End Property

Public Property Let HidePingDisplay(ByVal value As Boolean)
    m_HidePingDisplay = value
End Property

Public Property Get RetainOldBans() As Boolean
    RetainOldBans = m_RetainOldBans
End Property

Public Property Let RetainOldBans(ByVal value As Boolean)
    m_RetainOldBans = value
End Property

Public Property Get StoreAllBans() As Boolean
    StoreAllBans = m_StoreAllBans
End Property

Public Property Let StoreAllBans(ByVal value As Boolean)
    m_StoreAllBans = value
End Property

Public Property Get NamespaceConvention() As Byte
    NamespaceConvention = m_NamespaceConvention
End Property

Public Property Let NamespaceConvention(ByVal value As Byte)
    m_NamespaceConvention = value
End Property

Public Property Get UseD2Naming() As Boolean
    UseD2Naming = m_UseD2Naming
End Property

Public Property Let UseD2Naming(ByVal value As Boolean)
    m_UseD2Naming = value
End Property

Public Property Get ShowStatsIcons() As Boolean
    ShowStatsIcons = m_ShowStatsIcons
End Property

Public Property Let ShowStatsIcons(ByVal value As Boolean)
    m_ShowStatsIcons = value
End Property

Public Property Get ShowFlagIcons() As Boolean
    ShowFlagIcons = m_ShowFlagIcons
End Property

Public Property Let ShowFlagIcons(ByVal value As Boolean)
    m_ShowFlagIcons = value
End Property

Public Property Get ShowJoinLeaves() As Boolean
    ShowJoinLeaves = m_ShowJoinLeaves
End Property

Public Property Let ShowJoinLeaves(ByVal value As Boolean)
    m_ShowJoinLeaves = value
End Property

Public Property Get EnableMail() As Boolean
    EnableMail = m_EnableMail
End Property

Public Property Let EnableMail(ByVal value As Boolean)
    m_EnableMail = value
End Property

Public Property Get EnforceBanEvasion() As Boolean
    EnforceBanEvasion = m_EnforceBanEvasion
End Property

Public Property Let EnforceBanEvasion(ByVal value As Boolean)
    m_EnforceBanEvasion = value
End Property

Public Property Get LoggingLevel() As Byte
    LoggingLevel = m_LoggingLevel
End Property

Public Property Let LoggingLevel(ByVal value As Byte)
    m_LoggingLevel = value
End Property

Public Property Get EnablePhrasebans() As Boolean
    EnablePhrasebans = m_EnablePhrasebans
End Property

Public Property Let EnablePhrasebans(ByVal value As Boolean)
    m_EnablePhrasebans = value
End Property

Public Property Get CaseSensitiveFlags() As Boolean
    CaseSensitiveFlags = m_CaseSensitiveFlags
End Property

Public Property Let CaseSensitiveFlags(ByVal value As Boolean)
    m_CaseSensitiveFlags = value
End Property

Public Property Get BanUnderLevel() As Integer
    BanUnderLevel = m_BanUnderLevel
End Property

Public Property Let BanUnderLevel(ByVal value As Integer)
    m_BanUnderLevel = value
End Property

Public Property Get BanD2UnderLevel() As Integer
    BanD2UnderLevel = m_BanD2UnderLevel
End Property

Public Property Let BanD2UnderLevel(ByVal value As Integer)
    m_BanD2UnderLevel = value
End Property

Public Property Get LevelBanMessage() As String
    LevelBanMessage = m_LevelBanMessage
End Property

Public Property Let LevelBanMessage(ByVal value As String)
    m_LevelBanMessage = value
End Property

Public Property Get BanWc3Peons() As Boolean
    BanWc3Peons = m_BanWc3Peons
End Property

Public Property Let BanWc3Peons(ByVal value As Boolean)
    m_BanWc3Peons = value
End Property

Public Property Get KickOnYell() As Boolean
    KickOnYell = m_KickOnYell
End Property

Public Property Let KickOnYell(ByVal value As Boolean)
    m_KickOnYell = value
End Property

Public Property Get IdleBanDelay() As Integer
    IdleBanDelay = m_IdleBanDelay
End Property

Public Property Let IdleBanDelay(ByVal value As Integer)
    m_IdleBanDelay = value
End Property

Public Property Get DefaultShitlistGroup() As String
    DefaultShitlistGroup = m_DefaultShitlistGroup
End Property

Public Property Let DefaultShitlistGroup(ByVal value As String)
    m_DefaultShitlistGroup = value
End Property

Public Property Get DefaultTagbanGroup() As String
    DefaultTagbanGroup = m_DefaultTagbanGroup
End Property

Public Property Let DefaultTagbanGroup(ByVal value As String)
    m_DefaultTagbanGroup = value
End Property

Public Property Get DefaultSafelistGroup() As String
    DefaultSafelistGroup = m_DefaultSafelistGroup
End Property

Public Property Let DefaultSafelistGroup(ByVal value As String)
    m_DefaultSafelistGroup = value
End Property

Public Property Get AllowMp3Commands() As Boolean
    AllowMp3Commands = m_AllowMp3Commands
End Property

Public Property Let AllowMp3Commands(ByVal value As Boolean)
    m_AllowMp3Commands = value
End Property

Public Property Get ChannelProtectionMessage() As String
    ChannelProtectionMessage = m_ChannelProtectionMessage
End Property

Public Property Let ChannelProtectionMessage(ByVal value As String)
    m_ChannelProtectionMessage = value
End Property

Public Property Get EnforceIdleBans() As Boolean
    EnforceIdleBans = m_EnforceIdleBans
End Property

Public Property Let EnforceIdleBans(ByVal value As Boolean)
    m_EnforceIdleBans = value
End Property

Public Property Get KickIdleUsers() As Boolean
    KickIdleUsers = m_KickIdleUsers
End Property

Public Property Let KickIdleUsers(ByVal value As Boolean)
    m_KickIdleUsers = value
End Property

Public Property Get IpBans() As Boolean
    IpBans = m_IpBans
End Property

Public Property Let IpBans(ByVal value As Boolean)
    m_IpBans = value
End Property

Public Property Get FlashOnEvents() As Boolean
    FlashOnEvents = m_FlashOnEvents
End Property

Public Property Let FlashOnEvents(ByVal value As Boolean)
    m_FlashOnEvents = value
End Property

Public Property Get MinimizeToTray() As Boolean
    MinimizeToTray = m_MinimizeToTray
End Property

Public Property Let MinimizeToTray(ByVal value As Boolean)
    m_MinimizeToTray = value
End Property

Public Property Get DisableAutoComplete() As Boolean
    DisableAutoComplete = m_DisableAutoComplete
End Property

Public Property Let DisableAutoComplete(ByVal value As Boolean)
    m_DisableAutoComplete = value
End Property

Public Property Get DisableNameColoring() As Boolean
    DisableNameColoring = m_DisableNameColoring
End Property

Public Property Let DisableNameColoring(ByVal value As Boolean)
    m_DisableNameColoring = value
End Property

Public Property Get DisableVoidView() As Boolean
    DisableVoidView = m_DisableVoidView
End Property

Public Property Let DisableVoidView(ByVal value As Boolean)
    m_DisableVoidView = value
End Property

Public Property Get MediaPlayer() As String
    MediaPlayer = m_MediaPlayer
End Property

Public Property Let MediaPlayer(ByVal value As String)
    m_MediaPlayer = value
End Property

Public Property Get DisablePrefixBox() As Boolean
    DisablePrefixBox = m_DisablePrefixBox
End Property

Public Property Let DisablePrefixBox(ByVal value As Boolean)
    m_DisablePrefixBox = value
End Property

Public Property Get DisableSuffixBox() As Boolean
    DisableSuffixBox = m_DisableSuffixBox
End Property

Public Property Let DisableSuffixBox(ByVal value As Boolean)
    m_DisableSuffixBox = value
End Property

Public Property Get MathAllowUI() As Boolean
    MathAllowUI = m_MathAllowUI
End Property

Public Property Let MathAllowUI(ByVal value As Boolean)
    m_MathAllowUI = value
End Property

Public Property Get GreetMessage() As String
    GreetMessage = m_GreetMessage
End Property

Public Property Let GreetMessage(ByVal value As String)
    m_GreetMessage = value
End Property

Public Property Get UseGreetMessage() As Boolean
    UseGreetMessage = m_UseGreetMessage
End Property

Public Property Let UseGreetMessage(ByVal value As Boolean)
    m_UseGreetMessage = value
End Property

Public Property Get WhisperGreet() As Boolean
    WhisperGreet = m_WhisperGreet
End Property

Public Property Let WhisperGreet(ByVal value As Boolean)
    m_WhisperGreet = value
End Property

Public Property Get ChatDelay() As Long
    ChatDelay = m_ChatDelay
End Property

Public Property Let ChatDelay(ByVal value As Long)
    m_ChatDelay = value
End Property

Public Property Get FlashOnCatchPhrases() As Boolean
    FlashOnCatchPhrases = m_FlashOnCatchPhrases
End Property

Public Property Let FlashOnCatchPhrases(ByVal value As Boolean)
    m_FlashOnCatchPhrases = value
End Property

Public Property Get WinampPath() As String
    WinampPath = m_WinampPath
End Property

Public Property Let WinampPath(ByVal value As String)
    m_WinampPath = value
End Property

Public Property Get ScriptAllowUI() As Boolean
    ScriptAllowUI = m_ScriptAllowUI
End Property

Public Property Let ScriptAllowUI(ByVal value As Boolean)
    m_ScriptAllowUI = value
End Property

Public Property Get BanUdpPlugs() As Boolean
    BanUdpPlugs = m_BanUdpPlugs
End Property

Public Property Let BanUdpPlugs(ByVal value As Boolean)
    m_BanUdpPlugs = value
End Property

Public Property Get AutoCompletePostfix() As String
    AutoCompletePostfix = m_AutoCompletePostfix
End Property

Public Property Let AutoCompletePostfix(ByVal value As String)
    m_AutoCompletePostfix = value
End Property

Public Property Get HideBanMessages() As Boolean
    HideBanMessages = m_HideBanMessages
End Property

Public Property Let HideBanMessages(ByVal value As Boolean)
    m_HideBanMessages = value
End Property

Public Property Get DisableNews() As Boolean
    DisableNews = m_DisableNews
End Property

Public Property Let DisableNews(ByVal value As Boolean)
    m_DisableNews = value
End Property

Public Property Get BnlsSource() As String
    BnlsSource = m_BnlsSource
End Property

Public Property Let BnlsSource(ByVal value As String)
    m_BnlsSource = value
End Property

Public Property Get AddQMaxLength() As Integer
    AddQMaxLength = m_AddQMaxLength
End Property

Public Property Let AddQMaxLength(ByVal value As Integer)
    m_AddQMaxLength = value
End Property

Public Property Get AutoModSafelistLevel() As Integer
    AutoModSafelistLevel = m_AutoModSafelistLevel
End Property

Public Property Let AutoModSafelistLevel(ByVal value As Integer)
    m_AutoModSafelistLevel = value
End Property

Public Property Get D2NamingFormat() As String
    D2NamingFormat = m_D2NamingFormat
End Property

Public Property Let D2NamingFormat(ByVal value As String)
    m_D2NamingFormat = value
End Property

Public Property Get SecondsToIdle() As Long
    SecondsToIdle = m_SecondsToIdle
End Property

Public Property Let SecondsToIdle(ByVal value As Long)
    m_SecondsToIdle = value
End Property

Public Property Get QueueMaxCredits() As Long
    QueueMaxCredits = m_QueueMaxCredits
End Property

Public Property Let QueueMaxCredits(ByVal value As Long)
    m_QueueMaxCredits = value
End Property

Public Property Get QueueCostPerPacket() As Long
    QueueCostPerPacket = m_QueueCostPerPacket
End Property

Public Property Let QueueCostPerPacket(ByVal value As Long)
    m_QueueCostPerPacket = value
End Property

Public Property Get QueueCostPerByte() As Long
    QueueCostPerByte = m_QueueCostPerByte
End Property

Public Property Let QueueCostPerByte(ByVal value As Long)
    m_QueueCostPerByte = value
End Property

Public Property Get QueueCostPerByteOver() As Long
    QueueCostPerByteOver = m_QueueCostPerByteOver
End Property

Public Property Let QueueCostPerByteOver(ByVal value As Long)
    m_QueueCostPerByteOver = value
End Property

Public Property Get QueueStartingCredits() As Long
    QueueStartingCredits = m_QueueStartingCredits
End Property

Public Property Let QueueStartingCredits(ByVal value As Long)
    m_QueueStartingCredits = value
End Property

Public Property Get QueueThreshholdBytes() As Long
    QueueThreshholdBytes = m_QueueThreshholdBytes
End Property

Public Property Let QueueThreshholdBytes(ByVal value As Long)
    m_QueueThreshholdBytes = value
End Property

Public Property Get QueueCreditRate() As Long
    QueueCreditRate = m_QueueCreditRate
End Property

Public Property Let QueueCreditRate(ByVal value As Long)
    m_QueueCreditRate = value
End Property

Public Property Get DisableRtbAutoCopy() As Boolean
    DisableRtbAutoCopy = m_DisableRtbAutoCopy
End Property

Public Property Let DisableRtbAutoCopy(ByVal value As Boolean)
    m_DisableRtbAutoCopy = value
End Property

Public Property Get LocalIP() As String
    LocalIP = m_LocalIP
End Property

Public Property Let LocalIP(ByVal value As String)
    m_LocalIP = value
End Property

Public Property Get DisableScripts() As Boolean
    DisableScripts = m_DisableScripts
End Property

Public Property Let DisableScripts(ByVal value As Boolean)
    m_DisableScripts = value
End Property

Public Property Get SetKeyIgnoreLength() As Boolean
    SetKeyIgnoreLength = m_SetKeyIgnoreLength
End Property

Public Property Let SetKeyIgnoreLength(ByVal value As Boolean)
    m_SetKeyIgnoreLength = value
End Property

Public Property Get SetBotStatstring() As String
    SetBotStatstring = m_SetBotStatstring
End Property

Public Property Let SetBotStatstring(ByVal value As String)
    m_SetBotStatstring = value
End Property

Public Property Get ForceDefaultLocaleId() As Boolean
    ForceDefaultLocaleId = m_ForceDefaultLocaleId
End Property

Public Property Let ForceDefaultLocaleId(ByVal value As Boolean)
    m_ForceDefaultLocaleId = value
End Property

Public Property Get UdpString() As String
    UdpString = m_UdpString
End Property

Public Property Let UdpString(ByVal value As String)
    m_UdpString = value
End Property

Public Property Get KeyOwnerName() As String
    KeyOwnerName = m_KeyOwnerName
End Property

Public Property Let KeyOwnerName(ByVal value As String)
    m_KeyOwnerName = value
End Property

Public Property Get LowerCasePassword() As Boolean
    LowerCasePassword = m_LowerCasePassword
End Property

Public Property Let LowerCasePassword(ByVal value As Boolean)
    m_LowerCasePassword = value
End Property

Public Property Get IgnoreVersionCheckResult() As Boolean
    IgnoreVersionCheckResult = m_IgnoreVersionCheckResult
End Property

Public Property Let IgnoreVersionCheckResult(ByVal value As Boolean)
    m_IgnoreVersionCheckResult = value
End Property

Public Property Get PredefinedGateway() As String
    PredefinedGateway = m_PredefinedGateway
End Property

Public Property Let PredefinedGateway(ByVal value As String)
    m_PredefinedGateway = value
End Property

Public Property Get DoDefaultChannelJoin() As Boolean
    DoDefaultChannelJoin = m_DoDefaultChannelJoin
End Property

Public Property Let DoDefaultChannelJoin(ByVal value As Boolean)
    m_DoDefaultChannelJoin = value
End Property

Public Property Get ScriptViewer() As String
    ScriptViewer = m_ScriptViewer
End Property

Public Property Let ScriptViewer(ByVal value As String)
    m_ScriptViewer = value
End Property

Public Property Get WardenDebug() As Boolean
    WardenDebug = m_WardenDebug
End Property

Public Property Let WardenDebug(ByVal value As Boolean)
    m_WardenDebug = value
End Property

Public Property Get UseSpawnKey() As Boolean
    UseSpawnKey = m_UseSpawnKey
End Property

Public Property Let UseSpawnKey(ByVal value As Boolean)
    m_UseSpawnKey = value
End Property

Public Property Get AddQLinePostfix() As String
    AddQLinePostfix = m_AddQLinePostfix
End Property

Public Property Let AddQLinePostfix(ByVal value As String)
    m_AddQLinePostfix = value
End Property

Public Property Get CreateEmptyChannels() As String
    CreateEmptyChannels = m_CreateEmptyChannels
End Property

Public Property Let CreateEmptyChannels(ByVal value As String)
    m_CreateEmptyChannels = value
End Property

Public Property Get RegisterEmailAction() As String
    RegisterEmailAction = m_RegisterEmailAction
End Property

Public Property Let RegisterEmailAction(ByVal value As String)
    m_RegisterEmailAction = value
End Property

Public Property Get RegisterEmailDefault() As String
    RegisterEmailDefault = m_RegisterEmailDefault
End Property

Public Property Let RegisterEmailDefault(ByVal value As String)
    m_RegisterEmailDefault = value
End Property


' Returns the path to the specified file if it is overriden by this configuration.
Public Function GetFilePath(ByVal fileName As String) As String
    GetFilePath = ReadSetting(SECTION_FILEPATH, fileName)
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
    
    'MsgBox "Loading config from: " & sFilePath
    m_ConfigVersion = ReadSettingL(SECTION_MAIN, "ConfigVersion", 5)
    'MsgBox "Version: " & m_ConfigVersion
    
    If m_ConfigVersion < 6 Then
        Call LoadVersion5Config
    ElseIf m_ConfigVersion = 6 Then
        Call LoadVersion6Config
    Else
        Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unsupported config version: ", CStr(m_ConfigVersion))
    End If
End Sub

' Saves the config in the latest format.
Public Sub Save(Optional ByVal sFilePath As String = vbNullString)
    Dim normalPath As String
    normalPath = m_ConfigPath
    
    ' If we're saving to a different file
    If Len(sFilePath) > 0 Then
        m_ConfigPath = sFilePath
    End If
    
    'MsgBox "Saving config to: " & m_ConfigPath
    
    ' Backup the old file if this is going to be an upgrade.
    If Version < CONFIG_VERSION Then
        Dim fso As Object
        Dim backupPath As String
        
        backupPath = m_ConfigPath & "-backup"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FileExists(m_ConfigPath) And (Not fso.FileExists(backupPath))) Then
            fso.MoveFile m_ConfigPath, backupPath
            'MsgBox "Backup created."
        End If
        Set fso = Nothing
        
        m_ForceSave = True
    End If
    
    Call ConformValues
    
    ' Update our version number.
    m_ConfigVersion = CONFIG_VERSION
    
    WriteSetting SECTION_MAIN, "ConfigVersion", CONFIG_VERSION
    WriteSetting SECTION_MAIN, "DisableNews", m_DisableNews

    WriteSetting SECTION_CLIENT, "Username", m_Username
    WriteSetting SECTION_CLIENT, "Password", m_Password
    WriteSetting SECTION_CLIENT, "CdKey", m_CdKey
    WriteSetting SECTION_CLIENT, "ExpKey", m_ExpKey
    WriteSetting SECTION_CLIENT, "Spawn", m_UseSpawnKey
    WriteSetting SECTION_CLIENT, "Game", m_Product
    WriteSetting SECTION_CLIENT, "Server", m_Server
    WriteSetting SECTION_CLIENT, "HomeChannel", m_HomeChannel
    WriteSetting SECTION_CLIENT, "AutoConnect", m_ConnectOnStartup
    WriteSetting SECTION_CLIENT, "UseRealm", m_UseRealm
    WriteSetting SECTION_CLIENT, "UseBNLS", m_UseBnls
    WriteSetting SECTION_CLIENT, "BNLSServer", m_BnlsServer
    WriteSetting SECTION_CLIENT, "UseBNLSFinder", m_UseBnlsFinder
    WriteSetting SECTION_CLIENT, "BNLSFinderSource", m_BnlsSource
    WriteSetting SECTION_CLIENT, "UseProxy", m_UseProxy
    WriteSetting SECTION_CLIENT, "ProxyIP", m_ProxyIP
    WriteSetting SECTION_CLIENT, "ProxyPort", m_ProxyPort
    WriteSetting SECTION_CLIENT, "ProxyIsSocks5", m_ProxyIsSocks5
    
    WriteSetting SECTION_FEATURES, "UseBackupChannel", m_UseBackupChannel
    WriteSetting SECTION_FEATURES, "BackupChan", m_BackupChannel
    WriteSetting SECTION_FEATURES, "ReconnectDelay", m_ReconnectDelay
    WriteSetting SECTION_FEATURES, "BotMail", m_EnableMail
    WriteSetting SECTION_FEATURES, "ProfileAmp", m_UseProfileAmp
    WriteSetting SECTION_FEATURES, "VoidView", Not m_DisableVoidView
    WriteSetting SECTION_FEATURES, "GreetMessage", m_GreetMessage
    WriteSetting SECTION_FEATURES, "SendGreetMessage", m_UseGreetMessage
    WriteSetting SECTION_FEATURES, "WhisperGreet", m_WhisperGreet
    WriteSetting SECTION_FEATURES, "AntiIdle", m_IdlesEnabled
    WriteSetting SECTION_FEATURES, "IdleMessage", m_IdleMessage
    WriteSetting SECTION_FEATURES, "IdleDelay", m_IdleDelay
    WriteSetting SECTION_FEATURES, "IdleType", m_IdleType
    WriteSetting SECTION_FEATURES, "Trigger", m_Trigger
    WriteSetting SECTION_FEATURES, "BotOwner", m_BotOwner
    WriteSetting SECTION_FEATURES, "ChatFilters", m_UseChatFilters
    WriteSetting SECTION_FEATURES, "WhisperWindows", m_UseWhisperWindows
    WriteSetting SECTION_FEATURES, "WhisperResponses", m_WhisperResponses
    WriteSetting SECTION_FEATURES, "ChatDelay", m_ChatDelay
    WriteSetting SECTION_FEATURES, "MediaPlayer", m_MediaPlayer
    WriteSetting SECTION_FEATURES, "PlayerPath", m_WinampPath
    WriteSetting SECTION_FEATURES, "AllowMP3", m_AllowMp3Commands
    WriteSetting SECTION_FEATURES, "AutoCompletePostfix", m_AutoCompletePostfix
    WriteSetting SECTION_FEATURES, "CaseSensitiveFlags", m_CaseSensitiveFlags
    WriteSetting SECTION_FEATURES, "MultiLinePostfix", m_AddQLinePostfix
    
    WriteSetting SECTION_MODERATION, "BanEvasion", m_EnforceBanEvasion
    WriteSetting SECTION_MODERATION, "PhraseBans", m_EnablePhrasebans
    WriteSetting SECTION_MODERATION, "LevelBan", m_BanUnderLevel
    WriteSetting SECTION_MODERATION, "D2LevelBan", m_BanD2UnderLevel
    WriteSetting SECTION_MODERATION, "LevelBanMessage", m_LevelBanMessage
    WriteSetting SECTION_MODERATION, "PeonBans", m_BanWc3Peons
    WriteSetting SECTION_MODERATION, "KickOnYell", m_KickOnYell
    WriteSetting SECTION_MODERATION, "ShitlistGroup", m_DefaultShitlistGroup
    WriteSetting SECTION_MODERATION, "TagbansGroup", m_DefaultTagbanGroup
    WriteSetting SECTION_MODERATION, "SafelistGroup", m_DefaultSafelistGroup
    WriteSetting SECTION_MODERATION, "RetainOldBans", m_RetainOldBans
    WriteSetting SECTION_MODERATION, "StoreAllBans", m_StoreAllBans
    WriteSetting SECTION_MODERATION, "ProtectMessage", m_ChannelProtectionMessage
    WriteSetting SECTION_MODERATION, "RemoveIdleUsers", m_EnforceIdleBans
    WriteSetting SECTION_MODERATION, "IdleBanDelay", m_IdleBanDelay
    WriteSetting SECTION_MODERATION, "KickIdleUsers", m_KickIdleUsers
    WriteSetting SECTION_MODERATION, "IPBans", m_IpBans
    WriteSetting SECTION_MODERATION, "BanUDPPlugs", m_BanUdpPlugs
    WriteSetting SECTION_MODERATION, "AutoModerationSafelistLevel", m_AutoModSafelistLevel
    WriteSetting SECTION_MODERATION, "ChannelProtect", m_ChannelProtection
    WriteSetting SECTION_MODERATION, "QuietTime", m_QuietTime
    
    WriteSetting SECTION_UI, "ShowSplashScreen", m_ShowSplashScreen
    WriteSetting SECTION_UI, "ShowWhisperWindow", m_ShowWhisperWindow
    WriteSetting SECTION_UI, "MinimizeOnStartup", m_MinimizeOnStartup
    WriteSetting SECTION_UI, "UseUTF8", m_UTF8
    WriteSetting SECTION_UI, "DetectURLs", m_DetectUrls
    WriteSetting SECTION_UI, "ShowOutgoingWhispers", m_ShowOutgoingWhispers
    WriteSetting SECTION_UI, "HideWhispersInMain", m_HideWhispersInMain
    WriteSetting SECTION_UI, "TimestampFormat", m_Timestamp
    WriteSetting SECTION_UI, "ChatFont", m_ChatFont
    WriteSetting SECTION_UI, "ChatSize", m_ChatSize
    WriteSetting SECTION_UI, "ChannelFont", m_ChannelFont
    WriteSetting SECTION_UI, "ChannelSize", m_ChannelSize
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
    WriteSetting SECTION_UI, "DisableNameColors", m_DisableNameColoring
    WriteSetting SECTION_UI, "ShowOfflineFriends", m_ShowOfflineFriends
    WriteSetting SECTION_UI, "DisablePrefix", m_DisablePrefixBox
    WriteSetting SECTION_UI, "DisableSuffix", m_DisableSuffixBox
    WriteSetting SECTION_UI, "MathAllowUI", m_MathAllowUI
    WriteSetting SECTION_UI, "D2NamingFormat", m_D2NamingFormat
    WriteSetting SECTION_UI, "SecondsToIdle", m_SecondsToIdle
    WriteSetting SECTION_UI, "NoRTBAutomaticCopy", m_DisableRtbAutoCopy
    WriteSetting SECTION_UI, "HideBanMessages", m_HideBanMessages
    
    WriteSetting SECTION_UI_POS, "Left", m_Left
    WriteSetting SECTION_UI_POS, "Top", m_Top
    WriteSetting SECTION_UI_POS, "Height", m_Height
    WriteSetting SECTION_UI_POS, "Width", m_Width
    WriteSetting SECTION_UI_POS, "Maximized", m_Maximized
    WriteSetting SECTION_UI_POS, "LastSettingsPanel", m_LastSettingsPanel
    
    WriteSetting SECTION_LOGGING, "LogDBActions", m_LogDbAction
    WriteSetting SECTION_LOGGING, "LogCommands", m_LogCommands
    WriteSetting SECTION_LOGGING, "MaxBacklogSize", m_MaxBacklogSize
    WriteSetting SECTION_LOGGING, "MaxLogFileSize", m_MaxLogFileSize
    WriteSetting SECTION_LOGGING, "LogMode", m_LoggingLevel
    
    WriteSetting SECTION_QUEUE, "MaxCredits", m_QueueMaxCredits
    WriteSetting SECTION_QUEUE, "CostPerPacket", m_QueueCostPerPacket
    WriteSetting SECTION_QUEUE, "CostPerByte", m_QueueCostPerByte
    WriteSetting SECTION_QUEUE, "CostPerByteOverThreshhold", m_QueueCostPerByteOver
    WriteSetting SECTION_QUEUE, "StartingCredits", m_QueueStartingCredits
    WriteSetting SECTION_QUEUE, "ThreshholdBytes", m_QueueThreshholdBytes
    WriteSetting SECTION_QUEUE, "CreditRate", m_QueueCreditRate
    
    WriteSetting SECTION_SCRIPTING, "DisableScripts", m_DisableScripts
    WriteSetting SECTION_SCRIPTING, "AllowUI", m_ScriptAllowUI
    WriteSetting SECTION_SCRIPTING, "ScriptViewer", m_ScriptViewer
 
    WriteSetting SECTION_EMULATION, "IgnoreClanInvitations", m_IgnoreClanInvitations
    WriteSetting SECTION_EMULATION, "IgnoreKeyLength", m_SetKeyIgnoreLength
    WriteSetting SECTION_EMULATION, "DoNotUseDirectFList", m_DoNotUseDirectFList
    WriteSetting SECTION_EMULATION, "PingSpoof", m_PingSpoofing
    WriteSetting SECTION_EMULATION, "UseUDP", m_UseUDP
    WriteSetting SECTION_EMULATION, "SetStatstring", m_SetBotStatstring
    WriteSetting SECTION_EMULATION, "ForceDefaultLocaleID", m_ForceDefaultLocaleId
    WriteSetting SECTION_EMULATION, "UDPString", m_UdpString
    WriteSetting SECTION_EMULATION, "KeyOwner", m_KeyOwnerName
    WriteSetting SECTION_EMULATION, "LowerCasePassword", m_LowerCasePassword
    WriteSetting SECTION_EMULATION, "IgnoreVersionCheck", m_IgnoreVersionCheckResult
    WriteSetting SECTION_EMULATION, "PredefinedGateway", m_PredefinedGateway
    WriteSetting SECTION_EMULATION, "ForceJoinDefaultChannel", m_DoDefaultChannelJoin
    WriteSetting SECTION_EMULATION, "MaxMessageLength", m_AddQMaxLength
    WriteSetting SECTION_EMULATION, "CreateEmptyChannels", m_CreateEmptyChannels
    WriteSetting SECTION_EMULATION, "RegisterEmailAction", m_RegisterEmailAction
    WriteSetting SECTION_EMULATION, "RegisterEmailDefault", m_RegisterEmailDefault
    
    WriteSetting SECTION_DEBUG, "WardenDebug", m_WardenDebug
    
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

' Loads the old (2.7.1) config.
Private Sub LoadVersion5Config()
    m_ShowSplashScreen = ReadSettingB(SECTION_MAIN, "ShowSplash", True)
    m_ShowWhisperWindow = ReadSettingB(SECTION_MAIN, "ShowWhisperWindow", False)
    m_ConnectOnStartup = ReadSettingB(SECTION_MAIN, "ConnectOnStartup", False)
    m_MinimizeOnStartup = ReadSettingB(SECTION_MAIN, "MinimizeOnStartup", False)
    m_UseBnlsFinder = ReadSettingB(SECTION_MAIN, "UseAltBnls", True)
    m_IdlesEnabled = ReadSettingB(SECTION_MAIN, "Idles", False)
    m_IdleMessage = ReadSetting(SECTION_MAIN, "IdleMsg", vbNullString)
    m_IdleDelay = ReadSettingL(SECTION_MAIN, "IdleWait", 60)
    m_IdleType = ReadSetting(SECTION_MAIN, "IdleType", "msg")
    m_Username = ReadSetting(SECTION_MAIN, "Username", vbNullString)
    m_Password = ReadSetting(SECTION_MAIN, "Password", vbNullString)
    m_CdKey = ReadSetting(SECTION_MAIN, "CdKey", vbNullString)
    m_ExpKey = ReadSetting(SECTION_MAIN, "ExpKey", vbNullString)
    m_Product = ReadSetting(SECTION_MAIN, "Product", "STAR")
    m_Server = ReadSetting(SECTION_MAIN, "Server", "useast.battle.net")
    m_HomeChannel = ReadSetting(SECTION_MAIN, "HomeChan", vbNullString)
    m_BotOwner = ReadSetting(SECTION_MAIN, "Owner", vbNullString)
    m_Trigger = ReadSetting(SECTION_MAIN, "Trigger", "{.}")
    m_BnlsServer = ReadSetting(SECTION_MAIN, "BnlsServer", vbNullString)
    m_ShowOfflineFriends = ReadSettingB(SECTION_MAIN, "ShowOfflineFriends", False)
    m_UseWhisperWindows = ReadSettingB(SECTION_MAIN, "UseWWs", False)
    m_WhisperResponses = ReadSettingB(SECTION_MAIN, "WhisperBack", False)
    m_UseBnls = ReadSettingB(SECTION_MAIN, "UseBnls", True)
    m_LogDbAction = ReadSettingB(SECTION_MAIN, "LogDbAction", False)
    m_LogCommands = ReadSettingB(SECTION_MAIN, "LogCommands", False)
    m_MaxBacklogSize = ReadSettingL(SECTION_MAIN, "MaxBacklogSize", 10000)
    m_MaxLogFileSize = ReadSettingL(SECTION_MAIN, "MaxLogFileSize", 50000000)
    m_DoNotUseDirectFList = ReadSettingB(SECTION_MAIN, "DoNotUseDirectFList", False)
    m_DetectUrls = ReadSettingB(SECTION_MAIN, "URLDetect", True)
    m_ReconnectDelay = ReadSettingL(SECTION_MAIN, "ReconnectDelay", 1000)
    m_UseBackupChannel = ReadSettingB(SECTION_MAIN, "UseBackupChan", False)
    m_BackupChannel = ReadSetting(SECTION_MAIN, "BackupChan", vbNullString)
    m_UTF8 = ReadSettingB(SECTION_MAIN, "UTF8", True)
    m_ShowOutgoingWhispers = ReadSettingB(SECTION_MAIN, "ShowOutgoingWhispers", False)
    m_HideWhispersInMain = ReadSettingB(SECTION_MAIN, "HideWhispersInMain", False)
    m_IgnoreClanInvitations = ReadSettingB(SECTION_MAIN, "IgnoreClanInvitations", False)
    m_PingSpoofing = ReadSettingL(SECTION_MAIN, "Spoof", 0)
    m_ChannelProtection = ReadSettingB(SECTION_MAIN, "Protect", False)
    m_UseUDP = ReadSettingB(SECTION_MAIN, "UDP", False)
    m_QuietTime = ReadSettingB(SECTION_MAIN, "QuietTime", False)
    m_UseProxy = ReadSettingB(SECTION_MAIN, "UseProxy", False)
    m_ProxyPort = ReadSettingL(SECTION_MAIN, "ProxyPort", 0)
    m_ProxyIsSocks5 = ReadSettingB(SECTION_MAIN, "ProxyIsSocks5", False)
    m_UseRealm = ReadSettingB(SECTION_MAIN, "UseRealm", False)
    m_ProxyIP = ReadSetting(SECTION_MAIN, "ProxyIP", vbNullString)
    m_HideBanMessages = ReadSettingB(SECTION_MAIN, "HideBanMessages", False)
    
    m_Left = ReadSettingL(SECTION_POSITION, "Left", -1)
    m_Top = ReadSettingL(SECTION_POSITION, "Top", -1)
    m_Height = ReadSettingL(SECTION_POSITION, "Height", -1)
    m_Width = ReadSettingL(SECTION_POSITION, "Width", -1)
    m_Maximized = ReadSettingB(SECTION_POSITION, "Maximized", False)
    m_LastSettingsPanel = ReadSettingL(SECTION_POSITION, "LastSettingsPanel", 1)
    
    m_UseProfileAmp = ReadSettingB(SECTION_OTHER, "ProfileAmp", False)
    m_Timestamp = ReadSettingL(SECTION_OTHER, "Timestamp", 0)
    m_ChatFont = ReadSetting(SECTION_OTHER, "ChatFont", "Tahoma")
    m_ChatSize = ReadSettingL(SECTION_OTHER, "ChatSize", 9)
    m_ChannelFont = ReadSetting(SECTION_OTHER, "ChanFont", "Tahoma")
    m_ChannelSize = ReadSettingL(SECTION_OTHER, "ChanSize", 9)
    m_UseChatFilters = ReadSettingB(SECTION_OTHER, "Filters", False)
    m_HideClanDisplay = ReadSettingB(SECTION_OTHER, "HideClanDisplay", False)
    m_HidePingDisplay = ReadSettingB(SECTION_OTHER, "HidePingDisplay", False)
    m_RetainOldBans = ReadSettingB(SECTION_OTHER, "RetainOldBans", False)
    m_StoreAllBans = ReadSettingB(SECTION_OTHER, "StoreAllBans", False)
    m_NamespaceConvention = ReadSettingL(SECTION_OTHER, "NamespaceConvention", 0)
    m_UseD2Naming = ReadSettingB(SECTION_OTHER, "UseD2Naming", False)
    m_ShowStatsIcons = ReadSettingB(SECTION_OTHER, "ShowStatsIcons", True)
    m_ShowFlagIcons = ReadSettingB(SECTION_OTHER, "ShowFlagsIcons", True)
    m_ShowJoinLeaves = ReadSettingB(SECTION_OTHER, "JoinLeaves", True)
    m_EnableMail = ReadSettingB(SECTION_OTHER, "Mail", True)
    m_EnforceBanEvasion = ReadSettingB(SECTION_OTHER, "BanEvasion", False)
    m_LoggingLevel = ReadSettingL(SECTION_OTHER, "Logging", 1)
    m_EnablePhrasebans = ReadSettingB(SECTION_OTHER, "Phrasebans", False)
    m_CaseSensitiveFlags = ReadSettingB(SECTION_OTHER, "CaseSensitiveFlags", False)
    m_BanUnderLevel = ReadSettingL(SECTION_OTHER, "BanUnderLevel", -1)
    m_BanD2UnderLevel = ReadSettingL(SECTION_OTHER, "BanD2UnderLevel", -1)
    m_LevelBanMessage = ReadSetting(SECTION_OTHER, "LevelBanMsg", vbNullString)
    m_BanWc3Peons = ReadSettingB(SECTION_OTHER, "PeonBans", False)
    m_KickOnYell = ReadSettingB(SECTION_OTHER, "KickOnYell", False)
    m_IdleBanDelay = ReadSettingL(SECTION_OTHER, "IdleBanDelay", 0)
    m_DefaultShitlistGroup = ReadSetting(SECTION_OTHER, "DefaultShitlistGroup", vbNullString)
    m_DefaultTagbanGroup = ReadSetting(SECTION_OTHER, "DefaultTagbansGroup", vbNullString)
    m_DefaultSafelistGroup = ReadSetting(SECTION_OTHER, "DefaultSafelistGroup", vbNullString)
    m_AllowMp3Commands = ReadSettingB(SECTION_OTHER, "AllowMP3", True)
    m_ChannelProtectionMessage = ReadSetting(SECTION_OTHER, "ProtectMsg", "Channel Protection")
    m_EnforceIdleBans = ReadSettingB(SECTION_OTHER, "IdleBans", False)
    m_KickIdleUsers = ReadSettingB(SECTION_OTHER, "KickIdle", False)
    m_IpBans = ReadSettingB(SECTION_OTHER, "IPBans", False)
    m_FlashOnEvents = ReadSettingB(SECTION_OTHER, "FlashWindow", False)
    m_MinimizeToTray = ReadSettingB(SECTION_OTHER, "MinimizeToTray", False)
    m_DisableAutoComplete = ReadSettingB(SECTION_OTHER, "NoAutocomplete", False)
    m_DisableNameColoring = ReadSettingB(SECTION_OTHER, "NoColoring", False)
    m_DisableVoidView = ReadSettingB(SECTION_OTHER, "DisableVoidView", True)
    m_MediaPlayer = ReadSetting(SECTION_OTHER, "MediaPlayer", "Winamp")
    m_DisablePrefixBox = ReadSettingB(SECTION_OTHER, "DisablePrefix", False)
    m_DisableSuffixBox = ReadSettingB(SECTION_OTHER, "DisableSuffix", False)
    m_MathAllowUI = ReadSettingB(SECTION_OTHER, "MathAllowUI", False)
    m_GreetMessage = ReadSetting(SECTION_OTHER, "GreetMSg", vbNullString)
    m_UseGreetMessage = ReadSettingB(SECTION_OTHER, "UseGreets", False)
    m_WhisperGreet = ReadSettingB(SECTION_OTHER, "WhisperGreet", False)
    m_ChatDelay = ReadSettingL(SECTION_OTHER, "ChatDelay", 500)
    m_FlashOnCatchPhrases = ReadSettingB(SECTION_OTHER, "FlashOnCatchPhrases", True)
    m_WinampPath = ReadSetting(SECTION_OTHER, "WinampPath", vbNullString)
    m_ScriptAllowUI = ReadSettingB(SECTION_OTHER, "ScriptAllowUI", True)
    m_BanUdpPlugs = ReadSettingB(SECTION_OTHER, "PlugBans", False)
    m_AutoCompletePostfix = ReadSetting(SECTION_OTHER, "AutoCompletePostfix", vbNullString)
    
    m_DisableNews = ReadSettingB(SECTION_OVERRIDE, "DisableSBNews", False)
    m_BnlsSource = ReadSetting(SECTION_OVERRIDE, "BnlsSource", vbNullString)
    m_AddQMaxLength = ReadSettingL(SECTION_OVERRIDE, "AddQMaxLength", BNET_MSG_LENGTH)
    m_AutoModSafelistLevel = ReadSettingL(SECTION_OVERRIDE, "AutoModerationSafelistValue", 20)
    m_D2NamingFormat = ReadSetting(SECTION_OVERRIDE, "D2NamingFormat", "char (*name)")
    m_SecondsToIdle = ReadSettingL(SECTION_OVERRIDE, "SecondsToIdle", 600)
    m_QueueMaxCredits = ReadSettingL(SECTION_OVERRIDE, "QueueMaxCredits", 600)
    m_QueueCostPerPacket = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerPacket", 200)
    m_QueueCostPerByte = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerByte", 7)
    m_QueueCostPerByteOver = ReadSettingL(SECTION_OVERRIDE, "QueueCostPerByteOverThreshhold", 8)
    m_QueueStartingCredits = ReadSettingL(SECTION_OVERRIDE, "QueueStartingCredits", 200)
    m_QueueThreshholdBytes = ReadSettingL(SECTION_OVERRIDE, "QueueThreshholdBytes", 200)
    m_QueueCreditRate = ReadSettingL(SECTION_OVERRIDE, "QueueCreditRate", 7)
    m_DisableRtbAutoCopy = ReadSettingB(SECTION_OVERRIDE, "NoRTBAutomaticCopy", False)
    m_LocalIP = ReadSetting(SECTION_OVERRIDE, "LocalIP", vbNullString)
    m_DisableScripts = ReadSettingB(SECTION_OVERRIDE, "DisableScripts", False)
    m_SetKeyIgnoreLength = ReadSettingB(SECTION_OVERRIDE, "SetKeyIgnoreLength", False)
    m_SetBotStatstring = ReadSetting(SECTION_OVERRIDE, "SetBotStatstring", vbNullString)
    m_ForceDefaultLocaleId = ReadSettingB(SECTION_OVERRIDE, "ForceDefaultLocaleId", False)
    m_UdpString = ReadSetting(SECTION_OVERRIDE, "UdpString", vbNullString)
    m_KeyOwnerName = ReadSetting(SECTION_OVERRIDE, "OwnerName", vbNullString)
    m_LowerCasePassword = ReadSettingB(SECTION_OVERRIDE, "LowerCasePassword", False)
    m_IgnoreVersionCheckResult = ReadSettingB(SECTION_OVERRIDE, "Ignore0x51Reply", False)
    m_PredefinedGateway = ReadSetting(SECTION_OVERRIDE, "PredefinedGateway", vbNullString)
    m_DoDefaultChannelJoin = ReadSettingB(SECTION_OVERRIDE, "DoDefaultChannelJoin", False)
    m_ScriptViewer = ReadSetting(SECTION_OVERRIDE, "ScriptViewer", vbNullString)
    m_WardenDebug = ReadSettingB(SECTION_OVERRIDE, "WardenDebug", False)
    m_UseSpawnKey = ReadSettingB(SECTION_OVERRIDE, "SpawnKey", False)
    m_AddQLinePostfix = ReadSetting(SECTION_OVERRIDE, "AddQLinePostfix", " [more]")
    m_CreateEmptyChannels = ReadSetting(SECTION_OVERRIDE, "ChannelCreate", "ALWAYS")
    m_RegisterEmailAction = ReadSetting(SECTION_OVERRIDE, "RegisterEmailAction", "PROMPT")
    m_RegisterEmailDefault = ReadSetting(SECTION_OVERRIDE, "RegisterEmailDefault", vbNullString)
    
    If m_ExpKey = vbNullString Then m_ExpKey = ReadSetting(SECTION_MAIN, "LODKey", vbNullString)
    
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
    m_DisableNews = ReadSetting(SECTION_MAIN, "DisableNews", False)
        
    m_Username = ReadSetting(SECTION_CLIENT, "Username", vbNullString)
    m_Password = ReadSetting(SECTION_CLIENT, "Password", vbNullString)
    m_CdKey = ReadSetting(SECTION_CLIENT, "CdKey", vbNullString)
    m_ExpKey = ReadSetting(SECTION_CLIENT, "ExpKey", vbNullString)
    m_UseSpawnKey = ReadSettingB(SECTION_CLIENT, "Spawn", False)
    m_Product = ReadSetting(SECTION_CLIENT, "Game", "STAR")
    m_Server = ReadSetting(SECTION_CLIENT, "Server", "useast.battle.net")
    m_HomeChannel = ReadSetting(SECTION_CLIENT, "HomeChannel", vbNullString)
    m_ConnectOnStartup = ReadSettingB(SECTION_CLIENT, "AutoConnect", False)
    m_UseRealm = ReadSettingB(SECTION_CLIENT, "UseRealm", False)
    m_UseBnls = ReadSettingB(SECTION_CLIENT, "UseBNLS", True)
    m_BnlsServer = ReadSetting(SECTION_CLIENT, "BNLSServer", vbNullString)
    m_UseBnlsFinder = ReadSettingB(SECTION_CLIENT, "UseBNLSFinder", True)
    m_BnlsSource = ReadSetting(SECTION_CLIENT, "BNLSFinderSource", vbNullString)
    m_UseProxy = ReadSettingB(SECTION_CLIENT, "UseProxy", False)
    m_ProxyIP = ReadSetting(SECTION_CLIENT, "ProxyIP", vbNullString)
    m_ProxyPort = ReadSettingL(SECTION_CLIENT, "ProxyPort", 0)
    m_ProxyIsSocks5 = ReadSettingB(SECTION_CLIENT, "ProxyIsSocks5", False)
    
    m_UseBackupChannel = ReadSettingB(SECTION_FEATURES, "UseBackupChannel", False)
    m_BackupChannel = ReadSetting(SECTION_FEATURES, "BackupChan", vbNullString)
    m_ReconnectDelay = ReadSettingL(SECTION_FEATURES, "ReconnectDelay", 1000)
    m_EnableMail = ReadSettingB(SECTION_FEATURES, "BotMail", True)
    m_UseProfileAmp = ReadSettingB(SECTION_FEATURES, "ProfileAmp", False)
    m_GreetMessage = ReadSetting(SECTION_FEATURES, "GreetMessage", vbNullString)
    m_UseGreetMessage = ReadSettingB(SECTION_FEATURES, "SendGreetMessage", False)
    m_WhisperGreet = ReadSettingB(SECTION_FEATURES, "WhisperGreet", False)
    m_IdlesEnabled = ReadSettingB(SECTION_FEATURES, "AntiIdle", False)
    m_IdleMessage = ReadSetting(SECTION_FEATURES, "IdleMessage", vbNullString)
    m_IdleDelay = ReadSettingL(SECTION_FEATURES, "IdleDelay", 60)
    m_IdleType = ReadSetting(SECTION_FEATURES, "IdleType", "msg")
    m_Trigger = ReadSetting(SECTION_FEATURES, "Trigger", "{.}")
    m_BotOwner = ReadSetting(SECTION_FEATURES, "BotOwner", vbNullString)
    m_UseChatFilters = ReadSettingB(SECTION_FEATURES, "ChatFilters", False)
    m_UseWhisperWindows = ReadSettingB(SECTION_FEATURES, "WhisperWindows", False)
    m_WhisperResponses = ReadSettingB(SECTION_FEATURES, "WhisperResponses", False)
    m_ChatDelay = ReadSettingL(SECTION_FEATURES, "ChatDelay", 500)
    m_MediaPlayer = ReadSetting(SECTION_FEATURES, "MediaPlayer", "Winamp")
    m_WinampPath = ReadSetting(SECTION_FEATURES, "PlayerPath", vbNullString)
    m_AllowMp3Commands = ReadSettingB(SECTION_FEATURES, "AllowMP3", True)
    m_AutoCompletePostfix = ReadSetting(SECTION_FEATURES, "AutoCompletePostfix", vbNullString)
    m_CaseSensitiveFlags = ReadSettingB(SECTION_FEATURES, "CaseSensitiveFlags", False)
    m_AddQLinePostfix = ReadSetting(SECTION_FEATURES, "MultiLinePostfix", " [more]")
    
    m_EnforceBanEvasion = ReadSettingB(SECTION_MODERATION, "BanEvasion", False)
    m_EnablePhrasebans = ReadSettingB(SECTION_MODERATION, "PhraseBans", False)
    m_BanUnderLevel = ReadSettingL(SECTION_MODERATION, "LevelBan", -1)
    m_BanD2UnderLevel = ReadSettingL(SECTION_MODERATION, "D2LevelBan", -1)
    m_LevelBanMessage = ReadSetting(SECTION_MODERATION, "LevelBanMessage", vbNullString)
    m_BanWc3Peons = ReadSettingB(SECTION_MODERATION, "PeonBans", False)
    m_KickOnYell = ReadSettingB(SECTION_MODERATION, "KickOnYell", False)
    m_DefaultShitlistGroup = ReadSetting(SECTION_MODERATION, "ShitlistGroup", vbNullString)
    m_DefaultTagbanGroup = ReadSetting(SECTION_MODERATION, "TagbansGroup", vbNullString)
    m_DefaultSafelistGroup = ReadSetting(SECTION_MODERATION, "SafelistGroup", vbNullString)
    m_RetainOldBans = ReadSettingB(SECTION_MODERATION, "RetainOldBans", False)
    m_StoreAllBans = ReadSettingB(SECTION_MODERATION, "StoreAllBans", False)
    m_ChannelProtectionMessage = ReadSetting(SECTION_MODERATION, "ProtectMessage", "Channel Protection")
    m_EnforceIdleBans = ReadSettingB(SECTION_MODERATION, "RemoveIdleUsers", False)
    m_IdleBanDelay = ReadSettingL(SECTION_MODERATION, "IdleBanDelay", 0)
    m_KickIdleUsers = ReadSettingB(SECTION_MODERATION, "KickIdleUsers", False)
    m_IpBans = ReadSettingB(SECTION_MODERATION, "IPBans", False)
    m_BanUdpPlugs = ReadSettingB(SECTION_MODERATION, "BanUDPPlugs", False)
    m_AutoModSafelistLevel = ReadSettingL(SECTION_MODERATION, "AutoModerationSafelistLevel", 20)
    m_ChannelProtection = ReadSettingB(SECTION_MODERATION, "ChannelProtect", False)
    m_QuietTime = ReadSettingB(SECTION_MODERATION, "QuietTime", False)
    
    m_ShowSplashScreen = ReadSettingB(SECTION_UI, "ShowSplashScreen", True)
    m_ShowWhisperWindow = ReadSettingB(SECTION_UI, "ShowWhisperWindow", False)
    m_MinimizeOnStartup = ReadSettingB(SECTION_UI, "MinimizeOnStartup", False)
    m_UTF8 = ReadSettingB(SECTION_UI, "UseUTF8", True)
    m_DetectUrls = ReadSettingB(SECTION_UI, "DetectURLs", True)
    m_ShowOutgoingWhispers = ReadSettingB(SECTION_UI, "ShowOutgoingWhispers", False)
    m_HideWhispersInMain = ReadSettingB(SECTION_UI, "HideWhispersInMain", False)
    m_Timestamp = ReadSetting(SECTION_UI, "TimestampFormat", 0)
    m_ChatFont = ReadSetting(SECTION_UI, "ChatFont", "Tahoma")
    m_ChatSize = ReadSettingL(SECTION_UI, "ChatSize", 9)
    m_ChannelFont = ReadSetting(SECTION_UI, "ChannelFont", "Tahoma")
    m_ChannelSize = ReadSettingL(SECTION_UI, "ChannelSize", 9)
    m_HideClanDisplay = ReadSettingB(SECTION_UI, "HideClanDisplay", False)
    m_HidePingDisplay = ReadSettingB(SECTION_UI, "HidePingDisplay", False)
    m_NamespaceConvention = ReadSettingL(SECTION_UI, "NamespaceConvention", 0)
    m_UseD2Naming = ReadSettingB(SECTION_UI, "UseD2Naming", False)
    m_ShowStatsIcons = ReadSettingB(SECTION_UI, "ShowStatsIcons", True)
    m_ShowFlagIcons = ReadSettingB(SECTION_UI, "ShowFlagsIcons", True)
    m_ShowJoinLeaves = ReadSettingB(SECTION_UI, "ShowJoinLeaves", True)
    m_FlashOnEvents = ReadSettingB(SECTION_UI, "FlashOnEvents", False)
    m_FlashOnCatchPhrases = ReadSettingB(SECTION_UI, "FlashOnCatchPhrases", True)
    m_MinimizeToTray = ReadSettingB(SECTION_UI, "MinimizeToTray", False)
    m_DisableNameColoring = ReadSettingB(SECTION_UI, "DisableNameColors", False)
    m_ShowOfflineFriends = ReadSettingB(SECTION_UI, "ShowOfflineFriends", False)
    m_DisablePrefixBox = ReadSettingB(SECTION_UI, "DisablePrefix", False)
    m_DisableSuffixBox = ReadSettingB(SECTION_UI, "DisableSuffix", False)
    m_MathAllowUI = ReadSettingB(SECTION_UI, "MathAllowUI", False)
    m_D2NamingFormat = ReadSetting(SECTION_UI, "D2NamingFormat", "char (*name)")
    m_SecondsToIdle = ReadSettingL(SECTION_UI, "SecondsToIdle", 600)
    m_DisableRtbAutoCopy = ReadSettingB(SECTION_UI, "NoRTBAutomaticCopy", False)
    m_HideBanMessages = ReadSettingB(SECTION_UI, "HideBanMessages", False)
    
    m_Left = ReadSettingL(SECTION_UI_POS, "Left", -1)
    m_Top = ReadSettingL(SECTION_UI_POS, "Top", -1)
    m_Height = ReadSettingL(SECTION_UI_POS, "Height", -1)
    m_Width = ReadSettingL(SECTION_UI_POS, "Width", -1)
    m_Maximized = ReadSettingB(SECTION_UI_POS, "Maximized", False)
    m_LastSettingsPanel = ReadSettingL(SECTION_UI_POS, "LastSettingsPanel", 1)
    
    m_LogDbAction = ReadSettingB(SECTION_LOGGING, "LogDBActions", False)
    m_LogCommands = ReadSettingB(SECTION_LOGGING, "LogCommands", False)
    m_MaxBacklogSize = ReadSettingL(SECTION_LOGGING, "MaxBacklogSize", 10000)
    m_MaxLogFileSize = ReadSettingL(SECTION_LOGGING, "MaxLogFileSize", 50000000)
    m_LoggingLevel = ReadSetting(SECTION_LOGGING, "LogMode", 1)
    
    m_QueueMaxCredits = ReadSettingL(SECTION_QUEUE, "MaxCredits", 600)
    m_QueueCostPerPacket = ReadSettingL(SECTION_QUEUE, "CostPerPacket", 200)
    m_QueueCostPerByte = ReadSettingL(SECTION_QUEUE, "CostPerByte", 7)
    m_QueueCostPerByteOver = ReadSettingL(SECTION_QUEUE, "CostPerByteOverThreshhold", 8)
    m_QueueStartingCredits = ReadSettingL(SECTION_QUEUE, "StartingCredits", 200)
    m_QueueThreshholdBytes = ReadSettingL(SECTION_QUEUE, "ThreshholdBytes", 200)
    m_QueueCreditRate = ReadSettingL(SECTION_QUEUE, "CreditRate", 7)
    
    m_DisableScripts = ReadSettingB(SECTION_SCRIPTING, "DisableScripts", False)
    m_ScriptAllowUI = ReadSettingB(SECTION_SCRIPTING, "AllowUI", True)
    m_ScriptViewer = ReadSetting(SECTION_SCRIPTING, "ScriptViewer", vbNullString)
    
    m_IgnoreClanInvitations = ReadSettingB(SECTION_EMULATION, "IgnoreClanInvitations", False)
    m_SetKeyIgnoreLength = ReadSettingB(SECTION_EMULATION, "IgnoreKeyLength", False)
    m_DoNotUseDirectFList = ReadSettingB(SECTION_EMULATION, "DoNotUseDirectFList", False)
    m_PingSpoofing = ReadSettingL(SECTION_EMULATION, "PingSpoof", 0)
    m_UseUDP = ReadSettingB(SECTION_EMULATION, "UseUDP", False)
    m_SetBotStatstring = ReadSetting(SECTION_EMULATION, "SetStatstring", vbNullString)
    m_ForceDefaultLocaleId = ReadSettingB(SECTION_EMULATION, "ForceDefaultLocaleID", False)
    m_UdpString = ReadSetting(SECTION_EMULATION, "UDPString", vbNullString)
    m_KeyOwnerName = ReadSetting(SECTION_EMULATION, "KeyOwner", vbNullString)
    m_LowerCasePassword = ReadSettingB(SECTION_EMULATION, "LowerCasePassword", False)
    m_IgnoreVersionCheckResult = ReadSettingB(SECTION_EMULATION, "IgnoreVersionCheck", False)
    m_PredefinedGateway = ReadSetting(SECTION_EMULATION, "PredefinedGateway", vbNullString)
    m_DoDefaultChannelJoin = ReadSettingB(SECTION_EMULATION, "ForceJoinDefaultChannel", False)
    m_AddQMaxLength = ReadSettingL(SECTION_EMULATION, "MaxMessageLength", BNET_MSG_LENGTH)
    m_CreateEmptyChannels = ReadSetting(SECTION_EMULATION, "CreateEmptyChannels", "ALWAYS")
    m_RegisterEmailAction = ReadSetting(SECTION_EMULATION, "RegisterEmailAction", "PROMPT")
    m_RegisterEmailDefault = ReadSetting(SECTION_EMULATION, "RegisterEmailDefault", vbNullString)
    
    m_WardenDebug = ReadSettingB(SECTION_DEBUG, "WardenDebug", False)
    
    
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
    If m_Timestamp > 4 Or m_Timestamp < 0 Then m_Timestamp = 0
    
    m_CdKey = UCase$(m_CdKey)
    m_ExpKey = UCase$(m_ExpKey)
    m_CreateEmptyChannels = UCase$(m_CreateEmptyChannels)
    m_RegisterEmailAction = UCase$(m_RegisterEmailAction)
    
    If m_AutoModSafelistLevel < 1 Or m_AutoModSafelistLevel > 200 Then m_AutoModSafelistLevel = 20
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
    If m_QueueThreshholdBytes < 0 Then m_QueueThreshholdBytes = 200
    If m_QueueCreditRate < 0 Then m_QueueCreditRate = 7
    
    If m_AddQMaxLength < 1 Or m_AddQMaxLength > BNET_MSG_LENGTH Then m_AddQMaxLength = BNET_MSG_LENGTH
    
    
End Sub


Private Function ReadSetting(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim buffer As String
    Dim length As Long
    
    ' If the config file doesn't exist, return nothing.
    If Not FileExists() Then
        ReadSetting = vbNullString
        Exit Function
    End If
    
    ' Create a buffer to read the value into.
    buffer = String(255, vbNull)
    
    ' Read the value into the buffer.
    length = GetPrivateProfileString(section, key, defaultValue, buffer, 255, m_ConfigPath)
    ReadSetting = Left$(buffer, length)
    
    'MsgBox section & "->" & key & ": [" & CStr(length) & "] " & ReadSetting
End Function

' Reads a setting as a boolean
Private Function ReadSettingB(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim sVal As String
    
    sVal = ReadSetting(section, key, CStr(defaultValue))
    If sVal = "Y" Or sVal = "True" Or sVal = "1" Then
        ReadSettingB = True
    ElseIf sVal = "N" Or sVal = "False" Or sVal = "0" Then
        ReadSettingB = False
    Else
        ReadSettingB = defaultValue
    End If
End Function

' Reads a setting as a number
Private Function ReadSettingL(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Long = -1) As Long
    Dim sVal As String
    
    sVal = ReadSetting(section, key, CStr(defaultValue))
    If StrictIsNumeric(sVal) Then
        ReadSettingL = Val(sVal)
    Else
        ReadSettingL = defaultValue
    End If
End Function

' Sets a setting if needed.
Private Sub WriteSetting(ByVal section As String, ByVal key As String, ByVal value As String)
    Dim currentVal As String
    currentVal = ReadSetting(section, key)
    
    ' Only write if the value has changed or we are forcing a save.
    If (m_ForceSave Or (StrComp(currentVal, value, vbBinaryCompare) <> 0)) Then
        WritePrivateProfileString section, key, value, m_ConfigPath
    End If
End Sub

Private Function GetProductIndex(ByVal code As String) As Integer
    Dim i As Integer
    For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
        If UCase$(code) = m_ProductKeys(i) Then
            GetProductIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
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
