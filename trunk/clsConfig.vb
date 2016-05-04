Option Strict Off
Option Explicit On
Friend Class clsConfig
	
	' Written by Pyro
	'  2016-3-25
	
	Private Const CONFIG_VERSION As Short = 6
	
	'Config sections
	Private Const SECTION_MAIN As String = "Main"
	Private Const SECTION_POSITION As String = "Position"
	Private Const SECTION_OTHER As String = "Other"
	Private Const SECTION_OVERRIDE As String = "Override"
	Private Const SECTION_FILEPATH As String = "FilePaths"
	
	'New sections
	Private Const SECTION_UI As String = "UI"
	Private Const SECTION_UI_POS As String = "UI-Position"
	Private Const SECTION_CLIENT As String = "Client"
	Private Const SECTION_FEATURES As String = "Features"
	Private Const SECTION_MODERATION As String = "Moderation"
	Private Const SECTION_LOGGING As String = "Logging"
	Private Const SECTION_QUEUE As String = "Queue"
	Private Const SECTION_SCRIPTING As String = "Scripting"
	Private Const SECTION_EMULATION As String = "Emulation"
	Private Const SECTION_DEBUG As String = "Debug"
	
	'Non-setting variables
	Private m_ConfigVersion As Integer
	Private m_ConfigPath As String
	Private m_ForceSave As Boolean
	Private m_ProductKeys(8) As String
	Private m_DebugConfig As Boolean
	
	
	'[MAIN]
	Private m_DisableNews As Boolean
	
	'[CLIENT]
	Private m_Username As String
	Private m_Password As String
	Private m_CDKey As String
	Private m_EXPKey As String
	Private m_UseSpawn As Boolean
	Private m_Game As String
	Private m_Server As String
	Private m_HomeChannel As String
	Private m_AutoConnect As Boolean
	Private m_UseD2Realms As Boolean
	Private m_UseBNLS As Boolean
	Private m_BNLSServer As String
	Private m_UseBNLSFinder As Boolean
	Private m_BNLSFinderSource As String
	Private m_UseProxy As Boolean
	Private m_ProxyIP As String
	Private m_ProxyPort As Short
	Private m_ProxyType As String
	
	'[FEATURES]
	Private m_UseBackupChannel As Boolean
	Private m_BackupChannel As String
	Private m_ReconnectDelay As Integer
	Private m_BotMail As Boolean
	Private m_ProfileAmp As Boolean
	Private m_VoidView As Boolean
	Private m_GreetMessage As Boolean
	Private m_GreetMessageText As String
	Private m_WhisperGreet As Boolean
	Private m_IdleMessage As Boolean
	Private m_IdleMessageText As String
	Private m_IdleMessageDelay As Short
	Private m_IdleMessageType As String
	Private m_Trigger As String
	Private m_BotOwner As String
	Private m_ChatFilters As Boolean
	Private m_WhisperWindows As Boolean
	Private m_WhisperCommands As Boolean
	Private m_ChatDelay As Integer
	Private m_MediaPlayer As String
	Private m_MediaPlayerPath As String
	Private m_Mp3Commands As Boolean
	Private m_AutocompletePostfix As String
	Private m_CaseSensitiveDBFlags As Boolean
	Private m_MultiLinePostfix As String
	Private m_FriendsListTab As Boolean
	Private m_RealmAutoChooseServer As String
	Private m_RealmAutoChooseCharacter As String
	Private m_RealmAutoChooseDelay As Integer
	
	'[MODERATION]
	Private m_BanEvasion As Boolean
	Private m_Phrasebans As Boolean
	Private m_Phrasekick As Boolean
	Private m_LevelBanW3 As Short
	Private m_LevelBanD2 As Short
	Private m_LevelBanMessage As String
	Private m_PeonBan As Boolean
	Private m_KickOnYell As Boolean
	Private m_ShitlistGroup As String
	Private m_TagbanGroup As String
	Private m_SafelistGroup As String
	Private m_RetainOldBans As Boolean
	Private m_StoreAllBans As Boolean
	Private m_ChannelProtectionMessage As String
	Private m_IdleBan As Boolean
	Private m_IdleBanDelay As Short
	Private m_IdleBanKick As Boolean
	Private m_IPBans As Boolean
	Private m_UDPBan As Boolean
	Private m_AutoSafelistLevel As Short
	Private m_ChannelProtection As Boolean
	Private m_QuietTime As Boolean
	Private m_QuietTimeKick As Boolean
	Private m_PingBan As Boolean
	Private m_PingBanLevel As Integer
	
	'[UI]
	Private m_ShowSplashScreen As Boolean
	Private m_ShowWhisperBox As Boolean
	Private m_MinimizeOnStartup As Boolean
	Private m_UseUTF8 As Boolean
	Private m_UrlDetection As Boolean
	Private m_ShowOutgoingWhispers As Boolean
	Private m_HideWhispersInMain As Boolean
	Private m_TimestampMode As Byte
	Private m_ChatFont As String
	Private m_ChatFontSize As Short
	Private m_ChannelListFont As String
	Private m_ChannelListFontSize As Short
	Private m_HideClanDisplay As Boolean
	Private m_HidePingDisplay As Boolean
	Private m_NamespaceConvention As Byte
	Private m_UseD2Naming As Boolean
	Private m_ShowStatsIcons As Boolean
	Private m_ShowFlagIcons As Boolean
	Private m_ShowJoinLeaves As Boolean
	Private m_FlashOnEvents As Boolean
	Private m_FlashOnCatchPhrases As Boolean
	Private m_MinimizeToTray As Boolean
	Private m_NameColoring As Boolean
	Private m_ShowOfflineFriends As Boolean
	Private m_DisablePrefixBox As Boolean
	Private m_DisableSuffixBox As Boolean
	Private m_MathAllowUI As Boolean
	Private m_D2NamingFormat As String
	Private m_SecondsToIdle As Integer
	Private m_DisableRTBAutoCopy As Boolean
	Private m_HideBanMessages As Boolean
	Private m_NameAutoComplete As Boolean
	Private m_RealmHideMotd As Boolean
	
	'[UI-POSITION]
	Private m_PositionLeft As Integer
	Private m_PositionTop As Integer
	Private m_PositionHeight As Integer
	Private m_PositionWidth As Integer
	Private m_IsMaximized As Boolean
	Private m_LastSettingsPanel As Byte
	Private m_EnforceBounds As Boolean
	Private m_MonitorCount As Integer
	
	'[LOGGING]
	Private m_LogDBActions As Boolean
	Private m_LogCommands As Boolean
	Private m_MaxBacklogSize As Integer
	Private m_MaxLogFileSize As Integer
	Private m_LoggingMode As Byte
	
	'[QUEUE]
	Private m_QueueMaxCredits As Integer
	Private m_QueueCostPerPacket As Integer
	Private m_QueueCostPerByte As Integer
	Private m_QueueCostPerByteOver As Integer
	Private m_QueueStartingCredits As Integer
	Private m_QueueThresholdBytes As Integer
	Private m_QueueCreditRate As Integer
	
	'[SCRIPTING]
	Private m_DisableScripting As Boolean
	Private m_ScriptingAllowUI As Boolean
	Private m_ScriptViewer As String
	
	'[EMULATION]
	Private m_IgnoreClanInvites As Boolean
	Private m_IgnoreCDKeyLength As Boolean
	Private m_PingSpoofing As Byte
	Private m_UseUDP As Boolean
	Private m_CustomStatstring As String
	Private m_ForceDefaultLocaleID As Boolean
	Private m_UDPString As String
	Private m_CDKeyOwnerName As String
	Private m_UseLowerCasePassword As Boolean
	Private m_IgnoreVersionCheck As Boolean
	Private m_PredefinedGateway As String
	Private m_DefaultChannelJoin As Boolean
	Private m_MaxMessageLength As Short
	Private m_AutoCreateChannels As String
	Private m_RegisterEmailAction As String
	Private m_RegisterEmailDefault As String
	Private m_RealmServerPassword As String
	Private m_ProtocolID As Integer
	Private m_PlatformID As String
	Private m_ProductLanguage As String
	Private m_ServerCommandList As String
	Private m_VersionBytes(8) As Integer 'XXVerByte
	Private m_LogonSystems(8) As Integer 'XXLogonSystem
	
	'[DEBUG]
	Private m_DebugWarden As Boolean
	
	
	
	'-------------------------
	'   SECTION: MAIN
	'-------------------------
	
	
	Public Property DisableNews() As Boolean
		Get
			DisableNews = m_DisableNews
		End Get
		Set(ByVal Value As Boolean)
			m_DisableNews = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: CLIENT
	'-------------------------
	
	
	Public Property Username() As String
		Get
			Username = m_Username
		End Get
		Set(ByVal Value As String)
			m_Username = Value
		End Set
	End Property
	
	
	Public Property Password() As String
		Get
			Password = m_Password
		End Get
		Set(ByVal Value As String)
			m_Password = Value
		End Set
	End Property
	
	
	Public Property CDKey() As String
		Get
			CDKey = m_CDKey
		End Get
		Set(ByVal Value As String)
			m_CDKey = Value
		End Set
	End Property
	
	
	Public Property ExpKey() As String
		Get
			ExpKey = m_EXPKey
		End Get
		Set(ByVal Value As String)
			m_EXPKey = Value
		End Set
	End Property
	
	
	Public Property UseSpawn() As Boolean
		Get
			UseSpawn = m_UseSpawn
		End Get
		Set(ByVal Value As Boolean)
			m_UseSpawn = Value
		End Set
	End Property
	
	
	Public Property Game() As String
		Get
			Game = m_Game
		End Get
		Set(ByVal Value As String)
			m_Game = Value
		End Set
	End Property
	
	
	Public Property Server() As String
		Get
			Server = m_Server
		End Get
		Set(ByVal Value As String)
			m_Server = Value
		End Set
	End Property
	
	
	Public Property HomeChannel() As String
		Get
			HomeChannel = m_HomeChannel
		End Get
		Set(ByVal Value As String)
			m_HomeChannel = Value
		End Set
	End Property
	
	
	Public Property AutoConnect() As Boolean
		Get
			AutoConnect = m_AutoConnect
		End Get
		Set(ByVal Value As Boolean)
			m_AutoConnect = Value
		End Set
	End Property
	
	
	Public Property UseD2Realms() As Boolean
		Get
			UseD2Realms = m_UseD2Realms
		End Get
		Set(ByVal Value As Boolean)
			m_UseD2Realms = Value
		End Set
	End Property
	
	
	Public Property UseBNLS() As Boolean
		Get
			UseBNLS = m_UseBNLS
		End Get
		Set(ByVal Value As Boolean)
			m_UseBNLS = Value
		End Set
	End Property
	
	
	Public Property BNLSServer() As String
		Get
			BNLSServer = m_BNLSServer
		End Get
		Set(ByVal Value As String)
			m_BNLSServer = Value
		End Set
	End Property
	
	
	Public Property BNLSFinder() As Boolean
		Get
			BNLSFinder = m_UseBNLSFinder
		End Get
		Set(ByVal Value As Boolean)
			m_UseBNLSFinder = Value
		End Set
	End Property
	
	
	Public Property BNLSFinderSource() As String
		Get
			BNLSFinderSource = m_BNLSFinderSource
		End Get
		Set(ByVal Value As String)
			m_BNLSFinderSource = Value
		End Set
	End Property
	
	
	Public Property UseProxy() As Boolean
		Get
			UseProxy = m_UseProxy
		End Get
		Set(ByVal Value As Boolean)
			m_UseProxy = Value
		End Set
	End Property
	
	
	Public Property ProxyIP() As String
		Get
			ProxyIP = m_ProxyIP
		End Get
		Set(ByVal Value As String)
			m_ProxyIP = Value
		End Set
	End Property
	
	
	Public Property ProxyPort() As Short
		Get
			ProxyPort = m_ProxyPort
		End Get
		Set(ByVal Value As Short)
			m_ProxyPort = Value
		End Set
	End Property
	
	
	Public Property ProxyType() As String
		Get
			ProxyType = m_ProxyType
		End Get
		Set(ByVal Value As String)
			m_ProxyType = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: FEATURES
	'-------------------------
	
	
	Public Property UseBackupChannel() As Boolean
		Get
			UseBackupChannel = m_UseBackupChannel
		End Get
		Set(ByVal Value As Boolean)
			m_UseBackupChannel = Value
		End Set
	End Property
	
	
	Public Property BackupChannel() As String
		Get
			BackupChannel = m_BackupChannel
		End Get
		Set(ByVal Value As String)
			m_BackupChannel = Value
		End Set
	End Property
	
	
	Public Property ReconnectDelay() As Integer
		Get
			ReconnectDelay = m_ReconnectDelay
		End Get
		Set(ByVal Value As Integer)
			m_ReconnectDelay = Value
		End Set
	End Property
	
	
	Public Property BotMail() As Boolean
		Get
			BotMail = m_BotMail
		End Get
		Set(ByVal Value As Boolean)
			m_BotMail = Value
		End Set
	End Property
	
	
	Public Property ProfileAmp() As Boolean
		Get
			ProfileAmp = m_ProfileAmp
		End Get
		Set(ByVal Value As Boolean)
			m_ProfileAmp = Value
		End Set
	End Property
	
	
	Public Property VoidView() As Boolean
		Get
			VoidView = m_VoidView
		End Get
		Set(ByVal Value As Boolean)
			m_VoidView = Value
		End Set
	End Property
	
	
	Public Property GreetMessage() As Boolean
		Get
			GreetMessage = m_GreetMessage
		End Get
		Set(ByVal Value As Boolean)
			m_GreetMessage = Value
		End Set
	End Property
	
	
	Public Property GreetMessageText() As String
		Get
			GreetMessageText = m_GreetMessageText
		End Get
		Set(ByVal Value As String)
			m_GreetMessageText = Value
		End Set
	End Property
	
	
	Public Property WhisperGreet() As Boolean
		Get
			WhisperGreet = m_WhisperGreet
		End Get
		Set(ByVal Value As Boolean)
			m_WhisperGreet = Value
		End Set
	End Property
	
	
	Public Property IdleMessage() As Boolean
		Get
			IdleMessage = m_IdleMessage
		End Get
		Set(ByVal Value As Boolean)
			m_IdleMessage = Value
		End Set
	End Property
	
	
	Public Property IdleMessageText() As String
		Get
			IdleMessageText = m_IdleMessageText
		End Get
		Set(ByVal Value As String)
			m_IdleMessageText = Value
		End Set
	End Property
	
	
	Public Property IdleMessageDelay() As Short
		Get
			IdleMessageDelay = m_IdleMessageDelay
		End Get
		Set(ByVal Value As Short)
			m_IdleMessageDelay = Value
		End Set
	End Property
	
	
	Public Property IdleMessageType() As String
		Get
			IdleMessageType = m_IdleMessageType
		End Get
		Set(ByVal Value As String)
			m_IdleMessageType = Value
		End Set
	End Property
	
	
	Public Property Trigger() As String
		Get
			Trigger = GetProtectedString(m_Trigger)
		End Get
		Set(ByVal Value As String)
			m_Trigger = "{" & Value & "}"
		End Set
	End Property
	
	
	Public Property BotOwner() As String
		Get
			BotOwner = m_BotOwner
		End Get
		Set(ByVal Value As String)
			m_BotOwner = Value
		End Set
	End Property
	
	
	Public Property ChatFilters() As Boolean
		Get
			ChatFilters = m_ChatFilters
		End Get
		Set(ByVal Value As Boolean)
			m_ChatFilters = Value
		End Set
	End Property
	
	
	Public Property WhisperWindows() As Boolean
		Get
			WhisperWindows = m_WhisperWindows
		End Get
		Set(ByVal Value As Boolean)
			m_WhisperWindows = Value
		End Set
	End Property
	
	
	Public Property WhisperCommands() As Boolean
		Get
			WhisperCommands = m_WhisperCommands
		End Get
		Set(ByVal Value As Boolean)
			m_WhisperCommands = Value
		End Set
	End Property
	
	
	Public Property ChatDelay() As Integer
		Get
			ChatDelay = m_ChatDelay
		End Get
		Set(ByVal Value As Integer)
			m_ChatDelay = Value
		End Set
	End Property
	
	
	Public Property MediaPlayer() As String
		Get
			MediaPlayer = m_MediaPlayer
		End Get
		Set(ByVal Value As String)
			m_MediaPlayer = Value
		End Set
	End Property
	
	
	Public Property MediaPlayerPath() As String
		Get
			MediaPlayerPath = m_MediaPlayerPath
		End Get
		Set(ByVal Value As String)
			m_MediaPlayerPath = Value
		End Set
	End Property
	
	
	Public Property Mp3Commands() As Boolean
		Get
			Mp3Commands = m_Mp3Commands
		End Get
		Set(ByVal Value As Boolean)
			m_Mp3Commands = Value
		End Set
	End Property
	
	
	Public Property AutoCompletePostfix() As String
		Get
			AutoCompletePostfix = GetProtectedString(m_AutocompletePostfix)
		End Get
		Set(ByVal Value As String)
			m_AutocompletePostfix = "{" & Value & "}"
		End Set
	End Property
	
	
	Public Property CaseSensitiveDBFlags() As Boolean
		Get
			CaseSensitiveDBFlags = m_CaseSensitiveDBFlags
		End Get
		Set(ByVal Value As Boolean)
			m_CaseSensitiveDBFlags = Value
		End Set
	End Property
	
	
	Public Property MultiLinePostfix() As String
		Get
			MultiLinePostfix = GetProtectedString(m_MultiLinePostfix)
		End Get
		Set(ByVal Value As String)
			m_MultiLinePostfix = "{" & Value & "}"
		End Set
	End Property
	
	
	Public Property FriendsListTab() As Boolean
		Get
			FriendsListTab = m_FriendsListTab
		End Get
		Set(ByVal Value As Boolean)
			m_FriendsListTab = Value
		End Set
	End Property
	
	
	Public Property RealmAutoChooseServer() As String
		Get
			RealmAutoChooseServer = m_RealmAutoChooseServer
		End Get
		Set(ByVal Value As String)
			m_RealmAutoChooseServer = Value
		End Set
	End Property
	
	
	Public Property RealmAutoChooseCharacter() As String
		Get
			RealmAutoChooseCharacter = m_RealmAutoChooseCharacter
		End Get
		Set(ByVal Value As String)
			m_RealmAutoChooseCharacter = Value
		End Set
	End Property
	
	
	Public Property RealmAutoChooseDelay() As Integer
		Get
			RealmAutoChooseDelay = m_RealmAutoChooseDelay
		End Get
		Set(ByVal Value As Integer)
			m_RealmAutoChooseDelay = Value
		End Set
	End Property
	
	'-------------------------
	'   SECTION: MODERATION
	'-------------------------
	
	
	Public Property BanEvasion() As Boolean
		Get
			BanEvasion = m_BanEvasion
		End Get
		Set(ByVal Value As Boolean)
			m_BanEvasion = Value
		End Set
	End Property
	
	
	Public Property PhraseBans() As Boolean
		Get
			PhraseBans = m_Phrasebans
		End Get
		Set(ByVal Value As Boolean)
			m_Phrasebans = Value
		End Set
	End Property
	
	
	Public Property PhraseKick() As Boolean
		Get
			PhraseKick = m_Phrasekick
		End Get
		Set(ByVal Value As Boolean)
			m_Phrasekick = Value
		End Set
	End Property
	
	
	Public Property LevelBanW3() As Short
		Get
			LevelBanW3 = m_LevelBanW3
		End Get
		Set(ByVal Value As Short)
			m_LevelBanW3 = Value
		End Set
	End Property
	
	
	Public Property LevelBanD2() As Short
		Get
			LevelBanD2 = m_LevelBanD2
		End Get
		Set(ByVal Value As Short)
			m_LevelBanD2 = Value
		End Set
	End Property
	
	
	Public Property LevelBanMessage() As String
		Get
			LevelBanMessage = m_LevelBanMessage
		End Get
		Set(ByVal Value As String)
			m_LevelBanMessage = Value
		End Set
	End Property
	
	
	Public Property PeonBan() As Boolean
		Get
			PeonBan = m_PeonBan
		End Get
		Set(ByVal Value As Boolean)
			m_PeonBan = Value
		End Set
	End Property
	
	
	Public Property KickOnYell() As Boolean
		Get
			KickOnYell = m_KickOnYell
		End Get
		Set(ByVal Value As Boolean)
			m_KickOnYell = Value
		End Set
	End Property
	
	
	Public Property ShitlistGroup() As String
		Get
			ShitlistGroup = m_ShitlistGroup
		End Get
		Set(ByVal Value As String)
			m_ShitlistGroup = Value
		End Set
	End Property
	
	
	Public Property TagbanGroup() As String
		Get
			TagbanGroup = m_TagbanGroup
		End Get
		Set(ByVal Value As String)
			m_TagbanGroup = Value
		End Set
	End Property
	
	
	Public Property SafelistGroup() As String
		Get
			SafelistGroup = m_SafelistGroup
		End Get
		Set(ByVal Value As String)
			m_SafelistGroup = Value
		End Set
	End Property
	
	
	Public Property RetainOldBans() As Boolean
		Get
			RetainOldBans = m_RetainOldBans
		End Get
		Set(ByVal Value As Boolean)
			m_RetainOldBans = Value
		End Set
	End Property
	
	
	Public Property StoreAllBans() As Boolean
		Get
			StoreAllBans = m_StoreAllBans
		End Get
		Set(ByVal Value As Boolean)
			m_StoreAllBans = Value
		End Set
	End Property
	
	
	Public Property ChannelProtectionMessage() As String
		Get
			ChannelProtectionMessage = m_ChannelProtectionMessage
		End Get
		Set(ByVal Value As String)
			m_ChannelProtectionMessage = Value
		End Set
	End Property
	
	
	Public Property IdleBan() As Boolean
		Get
			IdleBan = m_IdleBan
		End Get
		Set(ByVal Value As Boolean)
			m_IdleBan = Value
		End Set
	End Property
	
	
	Public Property IdleBanDelay() As Short
		Get
			IdleBanDelay = m_IdleBanDelay
		End Get
		Set(ByVal Value As Short)
			m_IdleBanDelay = Value
		End Set
	End Property
	
	
	Public Property IdleBanKick() As Boolean
		Get
			IdleBanKick = m_IdleBanKick
		End Get
		Set(ByVal Value As Boolean)
			m_IdleBanKick = Value
		End Set
	End Property
	
	
	Public Property IPBans() As Boolean
		Get
			IPBans = m_IPBans
		End Get
		Set(ByVal Value As Boolean)
			m_IPBans = Value
		End Set
	End Property
	
	
	Public Property UDPBan() As Boolean
		Get
			UDPBan = m_UDPBan
		End Get
		Set(ByVal Value As Boolean)
			m_UDPBan = Value
		End Set
	End Property
	
	
	Public Property AutoSafelistLevel() As Short
		Get
			AutoSafelistLevel = m_AutoSafelistLevel
		End Get
		Set(ByVal Value As Short)
			m_AutoSafelistLevel = Value
		End Set
	End Property
	
	
	Public Property ChannelProtection() As Boolean
		Get
			ChannelProtection = m_ChannelProtection
		End Get
		Set(ByVal Value As Boolean)
			m_ChannelProtection = Value
		End Set
	End Property
	
	
	Public Property QuietTime() As Boolean
		Get
			QuietTime = m_QuietTime
		End Get
		Set(ByVal Value As Boolean)
			m_QuietTime = Value
		End Set
	End Property
	
	
	Public Property QuietTimeKick() As Boolean
		Get
			QuietTimeKick = m_QuietTimeKick
		End Get
		Set(ByVal Value As Boolean)
			m_QuietTimeKick = Value
		End Set
	End Property
	
	
	Public Property PingBan() As Boolean
		Get
			PingBan = m_PingBan
		End Get
		Set(ByVal Value As Boolean)
			m_PingBan = Value
		End Set
	End Property
	
	
	Public Property PingBanLevel() As Integer
		Get
			PingBanLevel = m_PingBanLevel
		End Get
		Set(ByVal Value As Integer)
			m_PingBanLevel = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: UI
	'-------------------------
	
	
	Public Property ShowSplashScreen() As Boolean
		Get
			ShowSplashScreen = m_ShowSplashScreen
		End Get
		Set(ByVal Value As Boolean)
			m_ShowSplashScreen = Value
		End Set
	End Property
	
	
	Public Property ShowWhisperBox() As Boolean
		Get
			ShowWhisperBox = m_ShowWhisperBox
		End Get
		Set(ByVal Value As Boolean)
			m_ShowWhisperBox = Value
		End Set
	End Property
	
	
	Public Property MinimizeOnStartup() As Boolean
		Get
			MinimizeOnStartup = m_MinimizeOnStartup
		End Get
		Set(ByVal Value As Boolean)
			m_MinimizeOnStartup = Value
		End Set
	End Property
	
	
	Public Property UseUTF8() As Boolean
		Get
			UseUTF8 = m_UseUTF8
		End Get
		Set(ByVal Value As Boolean)
			m_UseUTF8 = Value
		End Set
	End Property
	
	
	Public Property UrlDetection() As Boolean
		Get
			UrlDetection = m_UrlDetection
		End Get
		Set(ByVal Value As Boolean)
			m_UrlDetection = Value
		End Set
	End Property
	
	
	Public Property ShowOutgoingWhispers() As Boolean
		Get
			ShowOutgoingWhispers = m_ShowOutgoingWhispers
		End Get
		Set(ByVal Value As Boolean)
			m_ShowOutgoingWhispers = Value
		End Set
	End Property
	
	
	Public Property HideWhispersInMain() As Boolean
		Get
			HideWhispersInMain = m_HideWhispersInMain
		End Get
		Set(ByVal Value As Boolean)
			m_HideWhispersInMain = Value
		End Set
	End Property
	
	
	Public Property TimestampMode() As Byte
		Get
			TimestampMode = m_TimestampMode
		End Get
		Set(ByVal Value As Byte)
			m_TimestampMode = Value
		End Set
	End Property
	
	
	Public Property ChatFont() As String
		Get
			ChatFont = m_ChatFont
		End Get
		Set(ByVal Value As String)
			m_ChatFont = Value
		End Set
	End Property
	
	
	Public Property ChatFontSize() As Short
		Get
			ChatFontSize = m_ChatFontSize
		End Get
		Set(ByVal Value As Short)
			m_ChatFontSize = Value
		End Set
	End Property
	
	
	Public Property ChannelListFont() As String
		Get
			ChannelListFont = m_ChannelListFont
		End Get
		Set(ByVal Value As String)
			m_ChannelListFont = Value
		End Set
	End Property
	
	
	Public Property ChannelListFontSize() As Short
		Get
			ChannelListFontSize = m_ChannelListFontSize
		End Get
		Set(ByVal Value As Short)
			m_ChannelListFontSize = Value
		End Set
	End Property
	
	
	Public Property HideClanDisplay() As Boolean
		Get
			HideClanDisplay = m_HideClanDisplay
		End Get
		Set(ByVal Value As Boolean)
			m_HideClanDisplay = Value
		End Set
	End Property
	
	
	Public Property HidePingDisplay() As Boolean
		Get
			HidePingDisplay = m_HidePingDisplay
		End Get
		Set(ByVal Value As Boolean)
			m_HidePingDisplay = Value
		End Set
	End Property
	
	
	Public Property NamespaceConvention() As Byte
		Get
			NamespaceConvention = m_NamespaceConvention
		End Get
		Set(ByVal Value As Byte)
			m_NamespaceConvention = Value
		End Set
	End Property
	
	
	Public Property UseD2Naming() As Boolean
		Get
			UseD2Naming = m_UseD2Naming
		End Get
		Set(ByVal Value As Boolean)
			m_UseD2Naming = Value
		End Set
	End Property
	
	
	Public Property ShowStatsIcons() As Boolean
		Get
			ShowStatsIcons = m_ShowStatsIcons
		End Get
		Set(ByVal Value As Boolean)
			m_ShowStatsIcons = Value
		End Set
	End Property
	
	
	Public Property ShowFlagIcons() As Boolean
		Get
			ShowFlagIcons = m_ShowFlagIcons
		End Get
		Set(ByVal Value As Boolean)
			m_ShowFlagIcons = Value
		End Set
	End Property
	
	
	Public Property ShowJoinLeaves() As Boolean
		Get
			ShowJoinLeaves = m_ShowJoinLeaves
		End Get
		Set(ByVal Value As Boolean)
			m_ShowJoinLeaves = Value
		End Set
	End Property
	
	
	Public Property FlashOnEvents() As Boolean
		Get
			FlashOnEvents = m_FlashOnEvents
		End Get
		Set(ByVal Value As Boolean)
			m_FlashOnEvents = Value
		End Set
	End Property
	
	
	Public Property FlashOnCatchPhrases() As Boolean
		Get
			FlashOnCatchPhrases = m_FlashOnCatchPhrases
		End Get
		Set(ByVal Value As Boolean)
			m_FlashOnCatchPhrases = Value
		End Set
	End Property
	
	
	Public Property MinimizeToTray() As Boolean
		Get
			MinimizeToTray = m_MinimizeToTray
		End Get
		Set(ByVal Value As Boolean)
			m_MinimizeToTray = Value
		End Set
	End Property
	
	
	Public Property NameColoring() As Boolean
		Get
			NameColoring = m_NameColoring
		End Get
		Set(ByVal Value As Boolean)
			m_NameColoring = Value
		End Set
	End Property
	
	
	Public Property ShowOfflineFriends() As Boolean
		Get
			ShowOfflineFriends = m_ShowOfflineFriends
		End Get
		Set(ByVal Value As Boolean)
			m_ShowOfflineFriends = Value
		End Set
	End Property
	
	
	Public Property DisablePrefixBox() As Boolean
		Get
			DisablePrefixBox = m_DisablePrefixBox
		End Get
		Set(ByVal Value As Boolean)
			m_DisablePrefixBox = Value
		End Set
	End Property
	
	
	Public Property DisableSuffixBox() As Boolean
		Get
			DisableSuffixBox = m_DisableSuffixBox
		End Get
		Set(ByVal Value As Boolean)
			m_DisableSuffixBox = Value
		End Set
	End Property
	
	
	Public Property MathAllowUI() As Boolean
		Get
			MathAllowUI = m_MathAllowUI
		End Get
		Set(ByVal Value As Boolean)
			m_MathAllowUI = Value
		End Set
	End Property
	
	
	Public Property D2NamingFormat() As String
		Get
			D2NamingFormat = m_D2NamingFormat
		End Get
		Set(ByVal Value As String)
			m_D2NamingFormat = Value
		End Set
	End Property
	
	
	Public Property SecondsToIdle() As Integer
		Get
			SecondsToIdle = m_SecondsToIdle
		End Get
		Set(ByVal Value As Integer)
			m_SecondsToIdle = Value
		End Set
	End Property
	
	
	Public Property DisableRTBAutoCopy() As Boolean
		Get
			DisableRTBAutoCopy = m_DisableRTBAutoCopy
		End Get
		Set(ByVal Value As Boolean)
			m_DisableRTBAutoCopy = Value
		End Set
	End Property
	
	
	Public Property HideBanMessages() As Boolean
		Get
			HideBanMessages = m_HideBanMessages
		End Get
		Set(ByVal Value As Boolean)
			m_HideBanMessages = Value
		End Set
	End Property
	
	
	Public Property NameAutoComplete() As Boolean
		Get
			NameAutoComplete = m_NameAutoComplete
		End Get
		Set(ByVal Value As Boolean)
			m_NameAutoComplete = Value
		End Set
	End Property
	
	
	Public Property RealmHideMotd() As Boolean
		Get
			RealmHideMotd = m_RealmHideMotd
		End Get
		Set(ByVal Value As Boolean)
			m_RealmHideMotd = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: UI-POSITION
	'-------------------------
	
	
	Public Property PositionLeft() As Integer
		Get
			PositionLeft = m_PositionLeft
		End Get
		Set(ByVal Value As Integer)
			m_PositionLeft = Value
		End Set
	End Property
	
	
	Public Property PositionTop() As Integer
		Get
			PositionTop = m_PositionTop
		End Get
		Set(ByVal Value As Integer)
			m_PositionTop = Value
		End Set
	End Property
	
	
	Public Property PositionHeight() As Integer
		Get
			PositionHeight = m_PositionHeight
		End Get
		Set(ByVal Value As Integer)
			m_PositionHeight = Value
		End Set
	End Property
	
	
	Public Property PositionWidth() As Integer
		Get
			PositionWidth = m_PositionWidth
		End Get
		Set(ByVal Value As Integer)
			m_PositionWidth = Value
		End Set
	End Property
	
	
	Public Property IsMaximized() As Boolean
		Get
			IsMaximized = m_IsMaximized
		End Get
		Set(ByVal Value As Boolean)
			m_IsMaximized = Value
		End Set
	End Property
	
	
	Public Property LastSettingsPanel() As Byte
		Get
			LastSettingsPanel = m_LastSettingsPanel
		End Get
		Set(ByVal Value As Byte)
			m_LastSettingsPanel = Value
		End Set
	End Property
	
	
	Public Property EnforceScreenBounds() As Boolean
		Get
			EnforceScreenBounds = m_EnforceBounds
		End Get
		Set(ByVal Value As Boolean)
			m_EnforceBounds = Value
		End Set
	End Property
	
	
	Public Property MonitorCount() As Integer
		Get
			MonitorCount = m_MonitorCount
		End Get
		Set(ByVal Value As Integer)
			m_MonitorCount = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: LOGGING
	'-------------------------
	
	
	Public Property LogDBActions() As Boolean
		Get
			LogDBActions = m_LogDBActions
		End Get
		Set(ByVal Value As Boolean)
			m_LogDBActions = Value
		End Set
	End Property
	
	
	Public Property LogCommands() As Boolean
		Get
			LogCommands = m_LogCommands
		End Get
		Set(ByVal Value As Boolean)
			m_LogCommands = Value
		End Set
	End Property
	
	
	Public Property MaxBacklogSize() As Integer
		Get
			MaxBacklogSize = m_MaxBacklogSize
		End Get
		Set(ByVal Value As Integer)
			m_MaxBacklogSize = Value
		End Set
	End Property
	
	
	Public Property MaxLogFileSize() As Integer
		Get
			MaxLogFileSize = m_MaxLogFileSize
		End Get
		Set(ByVal Value As Integer)
			m_MaxLogFileSize = Value
		End Set
	End Property
	
	
	Public Property LoggingMode() As Byte
		Get
			LoggingMode = m_LoggingMode
		End Get
		Set(ByVal Value As Byte)
			m_LoggingMode = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: QUEUE
	'-------------------------
	
	
	Public Property QueueMaxCredits() As Integer
		Get
			QueueMaxCredits = m_QueueMaxCredits
		End Get
		Set(ByVal Value As Integer)
			m_QueueMaxCredits = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerPacket() As Integer
		Get
			QueueCostPerPacket = m_QueueCostPerPacket
		End Get
		Set(ByVal Value As Integer)
			m_QueueCostPerPacket = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerByte() As Integer
		Get
			QueueCostPerByte = m_QueueCostPerByte
		End Get
		Set(ByVal Value As Integer)
			m_QueueCostPerByte = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerByteOver() As Integer
		Get
			QueueCostPerByteOver = m_QueueCostPerByteOver
		End Get
		Set(ByVal Value As Integer)
			m_QueueCostPerByteOver = Value
		End Set
	End Property
	
	
	Public Property QueueStartingCredits() As Integer
		Get
			QueueStartingCredits = m_QueueStartingCredits
		End Get
		Set(ByVal Value As Integer)
			m_QueueStartingCredits = Value
		End Set
	End Property
	
	
	Public Property QueueThresholdBytes() As Integer
		Get
			QueueThresholdBytes = m_QueueThresholdBytes
		End Get
		Set(ByVal Value As Integer)
			m_QueueThresholdBytes = Value
		End Set
	End Property
	
	
	Public Property QueueCreditRate() As Integer
		Get
			QueueCreditRate = m_QueueCreditRate
		End Get
		Set(ByVal Value As Integer)
			m_QueueCreditRate = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: SCRIPTING
	'-------------------------
	
	
	Public Property DisableScripting() As Boolean
		Get
			DisableScripting = m_DisableScripting
		End Get
		Set(ByVal Value As Boolean)
			m_DisableScripting = Value
		End Set
	End Property
	
	
	Public Property ScriptingAllowUI() As Boolean
		Get
			ScriptingAllowUI = m_ScriptingAllowUI
		End Get
		Set(ByVal Value As Boolean)
			m_ScriptingAllowUI = Value
		End Set
	End Property
	
	
	Public Property ScriptViewer() As String
		Get
			ScriptViewer = m_ScriptViewer
		End Get
		Set(ByVal Value As String)
			m_ScriptViewer = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: EMULATION
	'-------------------------
	
	
	Public Property IgnoreClanInvites() As Boolean
		Get
			IgnoreClanInvites = m_IgnoreClanInvites
		End Get
		Set(ByVal Value As Boolean)
			m_IgnoreClanInvites = Value
		End Set
	End Property
	
	
	Public Property IgnoreCDKeyLength() As Boolean
		Get
			IgnoreCDKeyLength = m_IgnoreCDKeyLength
		End Get
		Set(ByVal Value As Boolean)
			m_IgnoreCDKeyLength = Value
		End Set
	End Property
	
	
	Public Property PingSpoofing() As Byte
		Get
			PingSpoofing = m_PingSpoofing
		End Get
		Set(ByVal Value As Byte)
			m_PingSpoofing = Value
		End Set
	End Property
	
	
	Public Property UseUDP() As Boolean
		Get
			UseUDP = m_UseUDP
		End Get
		Set(ByVal Value As Boolean)
			m_UseUDP = Value
		End Set
	End Property
	
	
	Public Property CustomStatstring() As String
		Get
			CustomStatstring = m_CustomStatstring
		End Get
		Set(ByVal Value As String)
			m_CustomStatstring = Value
		End Set
	End Property
	
	
	Public Property ForceDefaultLocaleID() As Boolean
		Get
			ForceDefaultLocaleID = m_ForceDefaultLocaleID
		End Get
		Set(ByVal Value As Boolean)
			m_ForceDefaultLocaleID = Value
		End Set
	End Property
	
	
	Public Property UDPString() As String
		Get
			UDPString = m_UDPString
		End Get
		Set(ByVal Value As String)
			m_UDPString = Value
		End Set
	End Property
	
	
	Public Property CDKeyOwnerName() As String
		Get
			CDKeyOwnerName = m_CDKeyOwnerName
		End Get
		Set(ByVal Value As String)
			m_CDKeyOwnerName = Value
		End Set
	End Property
	
	
	Public Property UseLowerCasePassword() As Boolean
		Get
			UseLowerCasePassword = m_UseLowerCasePassword
		End Get
		Set(ByVal Value As Boolean)
			m_UseLowerCasePassword = Value
		End Set
	End Property
	
	
	Public Property IgnoreVersionCheck() As Boolean
		Get
			IgnoreVersionCheck = m_IgnoreVersionCheck
		End Get
		Set(ByVal Value As Boolean)
			m_IgnoreVersionCheck = Value
		End Set
	End Property
	
	
	Public Property PredefinedGateway() As String
		Get
			PredefinedGateway = m_PredefinedGateway
		End Get
		Set(ByVal Value As String)
			m_PredefinedGateway = Value
		End Set
	End Property
	
	
	Public Property DefaultChannelJoin() As Boolean
		Get
			DefaultChannelJoin = m_DefaultChannelJoin
		End Get
		Set(ByVal Value As Boolean)
			m_DefaultChannelJoin = Value
		End Set
	End Property
	
	
	Public Property MaxMessageLength() As Short
		Get
			MaxMessageLength = m_MaxMessageLength
		End Get
		Set(ByVal Value As Short)
			m_MaxMessageLength = Value
		End Set
	End Property
	
	
	Public Property AutoCreateChannels() As String
		Get
			AutoCreateChannels = m_AutoCreateChannels
		End Get
		Set(ByVal Value As String)
			m_AutoCreateChannels = Value
		End Set
	End Property
	
	
	Public Property RegisterEmailAction() As String
		Get
			RegisterEmailAction = m_RegisterEmailAction
		End Get
		Set(ByVal Value As String)
			m_RegisterEmailAction = Value
		End Set
	End Property
	
	
	Public Property RegisterEmailDefault() As String
		Get
			RegisterEmailDefault = m_RegisterEmailDefault
		End Get
		Set(ByVal Value As String)
			m_RegisterEmailDefault = Value
		End Set
	End Property
	
	
	Public Property RealmServerPassword() As String
		Get
			RealmServerPassword = m_RealmServerPassword
		End Get
		Set(ByVal Value As String)
			m_RealmServerPassword = Value
		End Set
	End Property
	
	
	Public Property ProtocolID() As Integer
		Get
			ProtocolID = m_ProtocolID
		End Get
		Set(ByVal Value As Integer)
			m_ProtocolID = Value
		End Set
	End Property
	
	
	Public Property PlatformID() As String
		Get
			PlatformID = m_PlatformID
		End Get
		Set(ByVal Value As String)
			m_PlatformID = Value
		End Set
	End Property
	
	
	Public Property ProductLanguage() As String
		Get
			ProductLanguage = m_ProductLanguage
		End Get
		Set(ByVal Value As String)
			m_ProductLanguage = Value
		End Set
	End Property
	
	
	Public Property ServerCommandList() As String
		Get
			ServerCommandList = m_ServerCommandList
		End Get
		Set(ByVal Value As String)
			m_ServerCommandList = Value
		End Set
	End Property
	
	
	'-------------------------
	'   SECTION: DEBUG
	'-------------------------
	
	
	Public Property DebugWarden() As Boolean
		Get
			DebugWarden = m_DebugWarden
		End Get
		Set(ByVal Value As Boolean)
			m_DebugWarden = Value
		End Set
	End Property
	
	
	
	
	' Returns true if the config file exists.
	Public ReadOnly Property FileExists() As Boolean
		Get
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileExists = Not CBool(Len(Dir(m_ConfigPath)) = 0)
		End Get
	End Property
	
	' Returns the path this object was last loaded from.
	Public ReadOnly Property FilePath() As String
		Get
			FilePath = m_ConfigPath
		End Get
	End Property
	
	Public ReadOnly Property Version() As Integer
		Get
			Version = m_ConfigVersion
		End Get
	End Property
	
	
	Public Property DebugConfig() As Boolean
		Get
			DebugConfig = m_DebugConfig
		End Get
		Set(ByVal Value As Boolean)
			m_DebugConfig = Value
		End Set
	End Property
	
	' Returns the path to the specified file if it is overriden by this configuration.
	Public Function GetFilePath(ByVal sFileName As String) As String
		GetFilePath = ReadSetting(SECTION_FILEPATH, sFileName)
	End Function
	
	Public Function GetVersionByte(ByVal sProductCode As String) As Integer
		Dim iIndex As Short
		iIndex = GetProductIndex(GetProductInfo(sProductCode).ShortCode)
		
		If (iIndex >= LBound(m_VersionBytes) And iIndex <= UBound(m_VersionBytes)) Then
			GetVersionByte = m_VersionBytes(iIndex)
		Else
			GetVersionByte = 0
		End If
	End Function
	
	Public Sub SetVersionByte(ByVal sProductCode As String, ByVal iValue As Integer)
		Dim iIndex As Short
		iIndex = GetProductIndex(GetProductInfo(sProductCode).ShortCode)
		
		If (iIndex >= LBound(m_VersionBytes) And iIndex <= UBound(m_VersionBytes)) Then
			m_VersionBytes(iIndex) = iValue
		End If
	End Sub
	
	Public Function GetLogonSystem(ByVal sProductCode As String) As Integer
		Dim iIndex As Short
		iIndex = GetProductIndex(GetProductInfo(sProductCode).ShortCode)
		
		If (iIndex >= LBound(m_LogonSystems) And iIndex <= UBound(m_LogonSystems)) Then
			GetLogonSystem = m_LogonSystems(iIndex)
		Else
			GetLogonSystem = 0
		End If
	End Function
	
	Public Sub SetLogonSystem(ByVal sProductCode As String, ByVal iValue As Integer)
		Dim iIndex As Short
		iIndex = GetProductIndex(GetProductInfo(sProductCode).ShortCode)
		
		If (iIndex >= LBound(m_LogonSystems) And iIndex <= UBound(m_LogonSystems)) Then
			m_LogonSystems(iIndex) = iValue
		End If
	End Sub
	
	
	' Attempts to load the specified file into the object.
	Public Sub Load(ByVal sFilePath As String)
		m_ConfigPath = sFilePath
		
		' If it doesn't exist, run the v6 load to put in default values.
		If Not FileExists Then
			Call LoadDefaults()
			Exit Sub
		End If
		
		m_ConfigVersion = CInt(ReadSetting(SECTION_MAIN, "ConfigVersion", CStr(0)))
		
		Select Case m_ConfigVersion
			Case 5 : Call LoadVersion5Config()
			Case 6 : Call LoadVersion6Config()
			Case Else
				Call LoadDefaults()
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
		Dim fso As Object
		Dim backupPath As String
		Dim iPostfix As Short
		If Version < CONFIG_VERSION Then
			
			backupPath = m_ConfigPath & "-backup"
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If (Len(Dir(m_ConfigPath)) > 0) Then
                iPostfix = 1

                ' If a backup already exists, find an available filename
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                Do While (Len(Dir(backupPath)) > 0)
                    iPostfix = iPostfix + 1
                    backupPath = m_ConfigPath & "-backup" & CStr(iPostfix)
                Loop

                ' Copy the file and delete the original
                Call FileCopy(m_ConfigPath, backupPath)
                Call Kill(m_ConfigPath)

                Call frmChat.AddChat(RTBColors.InformationText, "Your config file is being updated. A backup of your old file has been placed in your bot folder.")
            End If
			'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			fso = Nothing
			
			m_ForceSave = True
		End If
		
		Call ConformValues()
		
		' Update our version number.
		m_ConfigVersion = CONFIG_VERSION
		
		WriteSetting(SECTION_MAIN, "ConfigVersion", CStr(m_ConfigVersion))
		WriteSetting(SECTION_MAIN, "DisableNews", CStr(m_DisableNews))
		
		WriteSetting(SECTION_CLIENT, "Username", m_Username)
		WriteSetting(SECTION_CLIENT, "Password", m_Password)
		WriteSetting(SECTION_CLIENT, "CdKey", m_CDKey)
		WriteSetting(SECTION_CLIENT, "ExpKey", m_EXPKey)
		WriteSetting(SECTION_CLIENT, "Spawn", CStr(m_UseSpawn))
		WriteSetting(SECTION_CLIENT, "Game", m_Game)
		WriteSetting(SECTION_CLIENT, "Server", m_Server)
		WriteSetting(SECTION_CLIENT, "HomeChannel", m_HomeChannel)
		WriteSetting(SECTION_CLIENT, "AutoConnect", CStr(m_AutoConnect))
		WriteSetting(SECTION_CLIENT, "UseRealm", CStr(m_UseD2Realms))
		WriteSetting(SECTION_CLIENT, "UseBNLS", CStr(m_UseBNLS))
		WriteSetting(SECTION_CLIENT, "BNLSServer", m_BNLSServer)
		WriteSetting(SECTION_CLIENT, "UseBNLSFinder", CStr(m_UseBNLSFinder))
		WriteSetting(SECTION_CLIENT, "BNLSFinderSource", m_BNLSFinderSource)
		WriteSetting(SECTION_CLIENT, "UseProxy", CStr(m_UseProxy))
		WriteSetting(SECTION_CLIENT, "ProxyIP", m_ProxyIP)
		WriteSetting(SECTION_CLIENT, "ProxyPort", CStr(m_ProxyPort))
		WriteSetting(SECTION_CLIENT, "ProxyType", m_ProxyType)
		
		WriteSetting(SECTION_FEATURES, "UseBackupChannel", CStr(m_UseBackupChannel))
		WriteSetting(SECTION_FEATURES, "BackupChannel", m_BackupChannel)
		WriteSetting(SECTION_FEATURES, "ReconnectDelay", CStr(m_ReconnectDelay))
		WriteSetting(SECTION_FEATURES, "BotMail", CStr(m_BotMail))
		WriteSetting(SECTION_FEATURES, "ProfileAmp", CStr(m_ProfileAmp))
		WriteSetting(SECTION_FEATURES, "VoidView", CStr(m_VoidView))
		WriteSetting(SECTION_FEATURES, "GreetMessage", CStr(m_GreetMessage))
		WriteSetting(SECTION_FEATURES, "GreetMessageText", m_GreetMessageText)
		WriteSetting(SECTION_FEATURES, "WhisperGreet", CStr(m_WhisperGreet))
		WriteSetting(SECTION_FEATURES, "IdleMessage", CStr(m_IdleMessage))
		WriteSetting(SECTION_FEATURES, "IdleText", m_IdleMessageText)
		WriteSetting(SECTION_FEATURES, "IdleDelay", CStr(m_IdleMessageDelay))
		WriteSetting(SECTION_FEATURES, "IdleType", m_IdleMessageType)
		WriteSetting(SECTION_FEATURES, "Trigger", m_Trigger)
		WriteSetting(SECTION_FEATURES, "BotOwner", m_BotOwner)
		WriteSetting(SECTION_FEATURES, "ChatFilters", CStr(m_ChatFilters))
		WriteSetting(SECTION_FEATURES, "WhisperWindows", CStr(m_WhisperWindows))
		WriteSetting(SECTION_FEATURES, "WhisperCommands", CStr(m_WhisperCommands))
		WriteSetting(SECTION_FEATURES, "ChatDelay", CStr(m_ChatDelay))
		WriteSetting(SECTION_FEATURES, "MediaPlayer", m_MediaPlayer)
		WriteSetting(SECTION_FEATURES, "PlayerPath", m_MediaPlayerPath)
		WriteSetting(SECTION_FEATURES, "AllowMP3", CStr(m_Mp3Commands))
		WriteSetting(SECTION_FEATURES, "NameAutoComplete", CStr(m_NameAutoComplete))
		WriteSetting(SECTION_FEATURES, "AutoCompletePostfix", m_AutocompletePostfix)
		WriteSetting(SECTION_FEATURES, "CaseSensitiveFlags", CStr(m_CaseSensitiveDBFlags))
		WriteSetting(SECTION_FEATURES, "MultiLinePostfix", m_MultiLinePostfix)
		WriteSetting(SECTION_FEATURES, "FriendsListTab", CStr(m_FriendsListTab))
		WriteSetting(SECTION_FEATURES, "RealmAutoChooseServer", m_RealmAutoChooseServer)
		WriteSetting(SECTION_FEATURES, "RealmAutoChooseCharacter", m_RealmAutoChooseCharacter)
		WriteSetting(SECTION_FEATURES, "RealmAutoChooseDelay", CStr(m_RealmAutoChooseDelay))
		
		WriteSetting(SECTION_MODERATION, "BanEvasion", CStr(m_BanEvasion))
		WriteSetting(SECTION_MODERATION, "PhraseBans", CStr(m_Phrasebans))
		WriteSetting(SECTION_MODERATION, "PhraseKick", CStr(m_Phrasekick))
		WriteSetting(SECTION_MODERATION, "LevelBanW3", CStr(m_LevelBanW3))
		WriteSetting(SECTION_MODERATION, "LevelBanD2", CStr(m_LevelBanD2))
		WriteSetting(SECTION_MODERATION, "LevelBanMessage", m_LevelBanMessage)
		WriteSetting(SECTION_MODERATION, "PeonBan", CStr(m_PeonBan))
		WriteSetting(SECTION_MODERATION, "KickOnYell", CStr(m_KickOnYell))
		WriteSetting(SECTION_MODERATION, "ShitlistGroup", m_ShitlistGroup)
		WriteSetting(SECTION_MODERATION, "TagbansGroup", m_TagbanGroup)
		WriteSetting(SECTION_MODERATION, "SafelistGroup", m_SafelistGroup)
		WriteSetting(SECTION_MODERATION, "RetainOldBans", CStr(m_RetainOldBans))
		WriteSetting(SECTION_MODERATION, "StoreAllBans", CStr(m_StoreAllBans))
		WriteSetting(SECTION_MODERATION, "ProtectMessage", m_ChannelProtectionMessage)
		WriteSetting(SECTION_MODERATION, "RemoveIdleUsers", CStr(m_IdleBan))
		WriteSetting(SECTION_MODERATION, "IdleBanDelay", CStr(m_IdleBanDelay))
		WriteSetting(SECTION_MODERATION, "KickIdleUsers", CStr(m_IdleBanKick))
		WriteSetting(SECTION_MODERATION, "IPBans", CStr(m_IPBans))
		WriteSetting(SECTION_MODERATION, "BanUDPPlugs", CStr(m_UDPBan))
		WriteSetting(SECTION_MODERATION, "AutoSafelistLevel", CStr(m_AutoSafelistLevel))
		WriteSetting(SECTION_MODERATION, "ChannelProtect", CStr(m_ChannelProtection))
		WriteSetting(SECTION_MODERATION, "QuietTime", CStr(m_QuietTime))
		WriteSetting(SECTION_MODERATION, "QuietTimeKick", CStr(m_QuietTimeKick))
		WriteSetting(SECTION_MODERATION, "PingBan", CStr(m_PingBan))
		WriteSetting(SECTION_MODERATION, "PingBanLevel", CStr(m_PingBanLevel))
		
		WriteSetting(SECTION_UI, "ShowSplashScreen", CStr(m_ShowSplashScreen))
		WriteSetting(SECTION_UI, "ShowWhisperWindow", CStr(m_ShowWhisperBox))
		WriteSetting(SECTION_UI, "MinimizeOnStartup", CStr(m_MinimizeOnStartup))
		WriteSetting(SECTION_UI, "UseUTF8", CStr(m_UseUTF8))
		WriteSetting(SECTION_UI, "DetectURLs", CStr(m_UrlDetection))
		WriteSetting(SECTION_UI, "ShowOutgoingWhispers", CStr(m_ShowOutgoingWhispers))
		WriteSetting(SECTION_UI, "HideWhispersInMain", CStr(m_HideWhispersInMain))
		WriteSetting(SECTION_UI, "TimestampMode", CStr(m_TimestampMode))
		WriteSetting(SECTION_UI, "ChatFont", m_ChatFont)
		WriteSetting(SECTION_UI, "ChatSize", CStr(m_ChatFontSize))
		WriteSetting(SECTION_UI, "ChannelFont", m_ChannelListFont)
		WriteSetting(SECTION_UI, "ChannelSize", CStr(m_ChannelListFontSize))
		WriteSetting(SECTION_UI, "HideClanDisplay", CStr(m_HideClanDisplay))
		WriteSetting(SECTION_UI, "HidePingDisplay", CStr(m_HidePingDisplay))
		WriteSetting(SECTION_UI, "NamespaceConvention", CStr(m_NamespaceConvention))
		WriteSetting(SECTION_UI, "UseD2Naming", CStr(m_UseD2Naming))
		WriteSetting(SECTION_UI, "ShowStatsIcons", CStr(m_ShowStatsIcons))
		WriteSetting(SECTION_UI, "ShowFlagsIcons", CStr(m_ShowFlagIcons))
		WriteSetting(SECTION_UI, "ShowJoinLeaves", CStr(m_ShowJoinLeaves))
		WriteSetting(SECTION_UI, "FlashOnEvents", CStr(m_FlashOnEvents))
		WriteSetting(SECTION_UI, "FlashOnCatchPhrases", CStr(m_FlashOnCatchPhrases))
		WriteSetting(SECTION_UI, "MinimizeToTray", CStr(m_MinimizeToTray))
		WriteSetting(SECTION_UI, "NameColoring", CStr(m_NameColoring))
		WriteSetting(SECTION_UI, "ShowOfflineFriends", CStr(m_ShowOfflineFriends))
		WriteSetting(SECTION_UI, "DisablePrefix", CStr(m_DisablePrefixBox))
		WriteSetting(SECTION_UI, "DisableSuffix", CStr(m_DisableSuffixBox))
		WriteSetting(SECTION_UI, "MathAllowUI", CStr(m_MathAllowUI))
		WriteSetting(SECTION_UI, "D2NamingFormat", m_D2NamingFormat)
		WriteSetting(SECTION_UI, "SecondsToIdle", CStr(m_SecondsToIdle))
		WriteSetting(SECTION_UI, "NoRTBAutomaticCopy", CStr(m_DisableRTBAutoCopy))
		WriteSetting(SECTION_UI, "HideBanMessages", CStr(m_HideBanMessages))
		WriteSetting(SECTION_UI, "RealmHideMotd", CStr(m_RealmHideMotd))
		
		WriteSetting(SECTION_UI_POS, "Left", CStr(m_PositionLeft))
		WriteSetting(SECTION_UI_POS, "Top", CStr(m_PositionTop))
		WriteSetting(SECTION_UI_POS, "Height", CStr(m_PositionHeight))
		WriteSetting(SECTION_UI_POS, "Width", CStr(m_PositionWidth))
		WriteSetting(SECTION_UI_POS, "Maximized", CStr(m_IsMaximized))
		WriteSetting(SECTION_UI_POS, "LastSettingsPanel", CStr(m_LastSettingsPanel))
		WriteSetting(SECTION_UI_POS, "EnforceBounds", CStr(m_EnforceBounds))
		WriteSetting(SECTION_UI_POS, "MonitorCount", CStr(m_MonitorCount))
		
		WriteSetting(SECTION_LOGGING, "LogDBActions", CStr(m_LogDBActions))
		WriteSetting(SECTION_LOGGING, "LogCommands", CStr(m_LogCommands))
		WriteSetting(SECTION_LOGGING, "MaxBacklogSize", CStr(m_MaxBacklogSize))
		WriteSetting(SECTION_LOGGING, "MaxLogFileSize", CStr(m_MaxLogFileSize))
		WriteSetting(SECTION_LOGGING, "LogMode", CStr(m_LoggingMode))
		
		WriteSetting(SECTION_QUEUE, "MaxCredits", CStr(m_QueueMaxCredits))
		WriteSetting(SECTION_QUEUE, "CostPerPacket", CStr(m_QueueCostPerPacket))
		WriteSetting(SECTION_QUEUE, "CostPerByte", CStr(m_QueueCostPerByte))
		WriteSetting(SECTION_QUEUE, "CostPerByteOverThreshhold", CStr(m_QueueCostPerByteOver))
		WriteSetting(SECTION_QUEUE, "StartingCredits", CStr(m_QueueStartingCredits))
		WriteSetting(SECTION_QUEUE, "ThreshholdBytes", CStr(m_QueueThresholdBytes))
		WriteSetting(SECTION_QUEUE, "CreditRate", CStr(m_QueueCreditRate))
		
		WriteSetting(SECTION_SCRIPTING, "DisableScripts", CStr(m_DisableScripting))
		WriteSetting(SECTION_SCRIPTING, "AllowUI", CStr(m_ScriptingAllowUI))
		WriteSetting(SECTION_SCRIPTING, "ScriptViewer", m_ScriptViewer)
		
		WriteSetting(SECTION_EMULATION, "IgnoreClanInvites", CStr(m_IgnoreClanInvites))
		WriteSetting(SECTION_EMULATION, "IgnoreKeyLength", CStr(m_IgnoreCDKeyLength))
		WriteSetting(SECTION_EMULATION, "PingSpoof", CStr(m_PingSpoofing))
		WriteSetting(SECTION_EMULATION, "UseUDP", CStr(m_UseUDP))
		WriteSetting(SECTION_EMULATION, "CustomStatstring", m_CustomStatstring)
		WriteSetting(SECTION_EMULATION, "ForceDefaultLocaleID", CStr(m_ForceDefaultLocaleID))
		WriteSetting(SECTION_EMULATION, "UDPString", m_UDPString)
		WriteSetting(SECTION_EMULATION, "KeyOwner", m_CDKeyOwnerName)
		WriteSetting(SECTION_EMULATION, "LowerCasePassword", CStr(m_UseLowerCasePassword))
		WriteSetting(SECTION_EMULATION, "IgnoreVersionCheck", CStr(m_IgnoreVersionCheck))
		WriteSetting(SECTION_EMULATION, "PredefinedGateway", m_PredefinedGateway)
		WriteSetting(SECTION_EMULATION, "ForceJoinDefaultChannel", CStr(m_DefaultChannelJoin))
		WriteSetting(SECTION_EMULATION, "MaxMessageLength", CStr(m_MaxMessageLength))
		WriteSetting(SECTION_EMULATION, "AutoCreateChannels", m_AutoCreateChannels)
		WriteSetting(SECTION_EMULATION, "RegisterEmailAction", m_RegisterEmailAction)
		WriteSetting(SECTION_EMULATION, "RegisterEmailDefault", m_RegisterEmailDefault)
		WriteSetting(SECTION_EMULATION, "RealmServerPassword", m_RealmServerPassword)
		WriteSetting(SECTION_EMULATION, "ProtocolID", CStr(m_ProtocolID))
		WriteSetting(SECTION_EMULATION, "PlatformID", m_PlatformID)
		WriteSetting(SECTION_EMULATION, "ProductLanguage", m_ProductLanguage)
		WriteSetting(SECTION_EMULATION, "ServerCommandList", m_ServerCommandList)
		
		WriteSetting(SECTION_DEBUG, "Warden", CStr(m_DebugWarden))
		
		Dim i As Short
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			If m_VersionBytes(i) > 0 Then
				WriteSetting(SECTION_EMULATION, m_ProductKeys(i) & "VerByte", Hex(m_VersionBytes(i)))
			End If
		Next 
		
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			If m_LogonSystems(i) > 0 Then
				WriteSetting(SECTION_EMULATION, m_ProductKeys(i) & "LogonSystem", Hex(m_LogonSystems(i)))
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
		m_AutocompletePostfix = "{,}"
		m_CaseSensitiveDBFlags = True
		m_MultiLinePostfix = "{ [more]}"
		m_FriendsListTab = True
		m_RealmAutoChooseServer = vbNullString
		m_RealmAutoChooseCharacter = vbNullString
		m_RealmAutoChooseDelay = 30
		
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
		m_RetainOldBans = False
		m_StoreAllBans = False
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
		m_RealmHideMotd = False
		
		'[UI-Position]
		m_PositionLeft = 0
		m_PositionTop = 0
		m_PositionHeight = 600
		m_PositionWidth = 800
		m_IsMaximized = False
		m_LastSettingsPanel = 0
		m_EnforceBounds = True
		m_MonitorCount = 1
		
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
		m_PingSpoofing = 0
		m_UseUDP = False
		m_CustomStatstring = vbNullString
		m_ForceDefaultLocaleID = False
		m_UDPString = "bnet"
		m_CDKeyOwnerName = vbNullString
		m_UseLowerCasePassword = True
		m_IgnoreVersionCheck = False
		m_PredefinedGateway = vbNullString
		m_DefaultChannelJoin = False
		m_MaxMessageLength = 223
		m_AutoCreateChannels = "ALWAYS"
		m_RegisterEmailAction = "PROMPT"
		m_RegisterEmailDefault = vbNullString
		m_RealmServerPassword = "password"
		m_ProtocolID = 0
		m_PlatformID = "IX86"
		m_ProductLanguage = vbNullString
		m_ServerCommandList = "w %,whisper %,m %,msg %,f m,clan mail,c mail,clan motd,c motd,away,dnd,ban %,kick %,j,join,channel,me,emote"
		
		Dim i As Short
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			m_VersionBytes(i) = -1
			m_LogonSystems(i) = -1
		Next 
		
		'[Debug]
		m_DebugWarden = False
	End Sub
	
	' Loads the old (2.7.1) config.
	Private Sub LoadVersion5Config()
		Call LoadDefaults()
		
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
		m_FriendsListTab = Not ReadSettingB(SECTION_MAIN, "DoNotUseDirectFList", m_FriendsListTab)
		
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
		m_UDPString = ReadSetting(SECTION_OVERRIDE, "UdpString", m_UDPString)
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
		
		Dim i As Short
		Dim sVal As String
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			sVal = ReadSetting(SECTION_OVERRIDE, m_ProductKeys(i) & "VerByte")
            If Len(sVal) > 0 Then m_VersionBytes(i) = CInt(Val("&H" & sVal))
			sVal = ReadSetting(SECTION_OVERRIDE, m_ProductKeys(i) & "LogonSystem")
            If Len(sVal) > 0 Then m_LogonSystems(i) = CInt(Val("&H" & sVal))
		Next 
		
		Call ConformValues()
	End Sub
	
	' Loads the version 6 config.
	Private Sub LoadVersion6Config()
		Call LoadDefaults()
		
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
		m_FriendsListTab = ReadSettingL(SECTION_FEATURES, "FriendsListTab", m_FriendsListTab)
		m_RealmAutoChooseServer = ReadSetting(SECTION_FEATURES, "RealmAutoChooseServer", m_RealmAutoChooseServer)
		m_RealmAutoChooseCharacter = ReadSetting(SECTION_FEATURES, "RealmAutoChooseCharacter", m_RealmAutoChooseCharacter)
		m_RealmAutoChooseDelay = ReadSettingL(SECTION_FEATURES, "RealmAutoChooseDelay", m_RealmAutoChooseDelay)
		
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
		m_RealmHideMotd = ReadSettingB(SECTION_UI, "RealmHideMotd", m_RealmHideMotd)
		
		m_PositionLeft = ReadSettingL(SECTION_UI_POS, "Left", m_PositionLeft)
		m_PositionTop = ReadSettingL(SECTION_UI_POS, "Top", m_PositionTop)
		m_PositionHeight = ReadSettingL(SECTION_UI_POS, "Height", m_PositionHeight)
		m_PositionWidth = ReadSettingL(SECTION_UI_POS, "Width", m_PositionWidth)
		m_IsMaximized = ReadSettingB(SECTION_UI_POS, "Maximized", m_IsMaximized)
		m_LastSettingsPanel = ReadSettingL(SECTION_UI_POS, "LastSettingsPanel", m_LastSettingsPanel)
		m_EnforceBounds = ReadSettingB(SECTION_UI_POS, "EnforceBounds", m_EnforceBounds)
		m_MonitorCount = ReadSettingL(SECTION_UI_POS, "MonitorCount", m_MonitorCount)
		
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
		m_PingSpoofing = ReadSettingL(SECTION_EMULATION, "PingSpoof", m_PingSpoofing)
		m_UseUDP = ReadSettingB(SECTION_EMULATION, "UseUDP", m_UseUDP)
		m_CustomStatstring = ReadSetting(SECTION_EMULATION, "CustomStatstring", m_CustomStatstring)
		m_ForceDefaultLocaleID = ReadSettingB(SECTION_EMULATION, "ForceDefaultLocaleID", m_ForceDefaultLocaleID)
		m_UDPString = ReadSetting(SECTION_EMULATION, "UDPString", m_UDPString)
		m_CDKeyOwnerName = ReadSetting(SECTION_EMULATION, "KeyOwner", m_CDKeyOwnerName)
		m_UseLowerCasePassword = ReadSettingB(SECTION_EMULATION, "LowerCasePassword", m_UseLowerCasePassword)
		m_IgnoreVersionCheck = ReadSettingB(SECTION_EMULATION, "IgnoreVersionCheck", m_IgnoreVersionCheck)
		m_PredefinedGateway = ReadSetting(SECTION_EMULATION, "PredefinedGateway", m_PredefinedGateway)
		m_DefaultChannelJoin = CBool(ReadSetting(SECTION_EMULATION, "ForceJoinDefaultChannel", CStr(m_DefaultChannelJoin)))
		m_MaxMessageLength = ReadSettingL(SECTION_EMULATION, "MaxMessageLength", m_MaxMessageLength)
		m_AutoCreateChannels = ReadSetting(SECTION_EMULATION, "AutoCreateChannels", m_AutoCreateChannels)
		m_RegisterEmailAction = ReadSetting(SECTION_EMULATION, "RegisterEmailAction", m_RegisterEmailAction)
		m_RegisterEmailDefault = ReadSetting(SECTION_EMULATION, "RegisterEmailDefault", m_RegisterEmailDefault)
		m_RealmServerPassword = ReadSetting(SECTION_EMULATION, "RealmServerPassword", m_RealmServerPassword)
		m_ProtocolID = ReadSettingL(SECTION_EMULATION, "ProtocolID", m_ProtocolID)
		m_PlatformID = ReadSetting(SECTION_EMULATION, "PlatformID", m_PlatformID)
		m_ProductLanguage = ReadSetting(SECTION_EMULATION, "ProductLanguage", m_ProductLanguage)
		m_ServerCommandList = ReadSetting(SECTION_EMULATION, "ServerCommandList", m_ServerCommandList)
		
		m_DebugWarden = ReadSettingB(SECTION_DEBUG, "Warden", m_DebugWarden)
		
		
		Dim i As Short
		Dim sVal As String
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			sVal = ReadSetting(SECTION_EMULATION, m_ProductKeys(i) & "VerByte")
            If Len(sVal) > 0 Then m_VersionBytes(i) = CInt(Val("&H" & sVal))
			
			sVal = ReadSetting(SECTION_EMULATION, m_ProductKeys(i) & "LogonSystem")
            If Len(sVal) > 0 Then m_LogonSystems(i) = CInt(Val("&H" & sVal))
		Next 
		
		Call ConformValues()
	End Sub
	
	' Ensures that values are within valid ranges and in the proper format.
	Private Sub ConformValues()
		If m_TimestampMode > 4 Or m_TimestampMode < 0 Then m_TimestampMode = 0
		
		m_CDKey = UCase(m_CDKey)
		m_EXPKey = UCase(m_EXPKey)
		m_AutoCreateChannels = UCase(m_AutoCreateChannels)
		m_RegisterEmailAction = UCase(m_RegisterEmailAction)
		
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
		If Len(sValue) > 1 Then
			GetProtectedString = Mid(sValue, 2, Len(sValue) - 2)
		Else
			GetProtectedString = sValue
		End If
	End Function
	
	
	Private Function ReadSetting(ByVal section As String, ByVal Key As String, Optional ByVal DefaultValue As String = vbNullString) As String
		Dim Buffer As String
		Dim length As Integer
		
		' If the config file doesn't exist, return the default.
		If Not FileExists() Then
			ReadSetting = DefaultValue
			Exit Function
		End If
		
		' Create a buffer to read the value into.
		Buffer = New String(Chr(VariantType.Null), 255)
		
		' Read the value into the buffer.
		length = GetPrivateProfileString(section, Key, DefaultValue, Buffer, 255, m_ConfigPath)
		ReadSetting = Left(Buffer, length)
		
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
	Private Function ReadSettingL(ByVal section As String, ByVal Key As String, Optional ByVal DefaultValue As Integer = -1) As Integer
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
		If (m_ForceSave Or (StrComp(currentVal, Value, CompareMethod.Binary) <> 0)) Then
			WritePrivateProfileString(section, Key, Value, m_ConfigPath)
			
			If m_DebugConfig Then frmChat.AddChat(RTBColors.InformationText, StringFormat("[CONFIG] Set setting: {0}.{1} -> {2}", section, Key, Value))
		End If
	End Sub
	
	Private Function GetProductIndex(ByVal Code As String) As Short
		Dim i As Short
		For i = LBound(m_ProductKeys) To UBound(m_ProductKeys)
			If UCase(Code) = m_ProductKeys(i) Then
				GetProductIndex = i
				Exit Function
			End If
		Next 
		GetProductIndex = -1
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
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
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class