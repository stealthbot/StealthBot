Option Strict Off
Option Explicit On
Friend Class clsBotVars
	'---------------------------------------------------------------------------------------
	' Module    : clsBotVars
	' DateTime  : 1/22/2005 16:09
	' Author    : Andy (andy@stealthbot.net)
	' Purpose   : stores bot variables at runtime and is exposed to the scripting subsystem
	'  - Added Realm string in response to topic 29227 -andy
	'---------------------------------------------------------------------------------------
	
	Private m_sUsername As String
	Private m_sPassword As String
	Private m_sCDKey As String
	Private m_sExpKey As String
	Private m_sProduct As String
	Private m_sServer As String
	Private m_byBanUnderLevel As Byte
	Private m_byBanD2UnderLevel As Byte
	Private m_sBanUnderLevelMsg As String
	Private m_byKickOnYell As Byte
	Private m_sHashFilePath As String
	Private m_bBNLS As Boolean
	Private m_bLoadMode As Boolean
	Private m_byLogging As Byte
	Private m_bySpoof As Byte
	Private m_bDisableMP3Commands As Boolean
	Private m_bUseUDP As Boolean
	Private m_byFlashOnEvents As Byte
	Private m_bUseProxy As Boolean
	Private m_sProxyIP As String
	Private m_lProxyPort As Integer
	Private m_eProxyStatus As modEnum.enuProxyStatus
	Private m_bProxyIsSocks5 As Boolean
	Private m_lMaxBacklogSize As Integer
	Private m_lMaxLogfileSize As Integer
	Private m_bUseGreet As Boolean
	Private m_sGreetMsg As String
	Private m_bWhisperGreet As Boolean
	Private m_bUseRealm As Boolean
	Private m_bNoTray As Boolean
	Private m_bShowOfflineFriends As Boolean
	Private m_lReconnectDelay As Integer
	Private m_lAutofilterMS As Integer
	Private m_bNoAutocompletion As Boolean
	Private m_bNoColoring As Boolean
	Private m_lSecondsToIdle As Integer
	Private m_sTrigger As String
	Private m_bScriptMessageVeto As Boolean
	Private m_bIPBans As Boolean
	Private m_bPlugBan As Boolean
	Private m_bClientBans As Boolean
	Private m_byTSSetting As Byte
	Private m_sHomeChannel As String
	Private m_sLastChannel As String
	Private m_colPublicChannels As Collection
	Private m_bQuietTime As Boolean
	Private m_bUseBackupChan As Boolean
	Private m_sBackupChan As String
	Private m_byIB_On As Byte
	Private m_iIB_Wait As Short
	Private m_bIB_Kick As Boolean
	Private m_sChannelPassword As String
	Private m_byChannelPasswordDelay As Byte
	Private m_bBanEvasion As Boolean
	Private m_bLogDBActions As Boolean
	Private m_bLogCommands As Boolean
	Private m_sBotOwner As String
	Private m_bBanPeons As Boolean
	Private m_bWhisperCmds As Boolean
	Private m_bDoScroll As Boolean
	Private m_bLockChat As Boolean
	Private m_iJoinWatch As Short
	Private m_sBNLSServer As String
	Private m_bNoRTBAutomaticCopy As Boolean
	Private m_bUseGameConventions As Boolean
	Private m_bUseD2GameConventions As Boolean
	Private m_bUseW3GameConventions As Boolean
	Private m_sClan As String
	Private m_sGateway As String
	Private m_sRealm As String
	Private m_NoSupportMultiCharTrigger As Boolean
	Private m_MediaPlayer As String
	Private m_ChatDelay As Integer
	Private m_ScriptLock As Boolean
	Private m_shitlist_group As String
	Private m_safelist_group As String
	Private m_tagbans_group As String
	Private m_autocomplete_postfix As String
	Private m_show_stats_icons As Boolean
	Private m_show_flags_icons As Boolean
	Private m_retain_old_bans As Boolean
	Private m_store_all_bans As Boolean
	Private m_use_alt_bnls As Boolean
	Private m_case_sensitive_flags As Boolean
	Private m_byGatewayConventions As Byte
	Private m_bUseD2Naming As Boolean
	Private m_sD2NamingFormat As String
	' Queue stuff
	Private m_lMaxCredits As Integer
	Private m_lCostPerPacket As Integer
	Private m_lCostPerByte As Integer
	Private m_lCostPerByteOverThreshhold As Integer
	Private m_lStartingCredits As Integer
	Private m_lThreshholdBytes As Integer
	Private m_lCreditRate As Integer
	
	
	
	Public Property Username() As String
		Get
			Username = m_sUsername
		End Get
		Set(ByVal Value As String)
			m_sUsername = Value
		End Set
	End Property
	
	
	Public Property Password() As String
		Get
			Password = m_sPassword
		End Get
		Set(ByVal Value As String)
			m_sPassword = Value
		End Set
	End Property
	
	
	Public Property CDKey() As String
		Get
			CDKey = m_sCDKey
		End Get
		Set(ByVal Value As String)
			m_sCDKey = Value
		End Set
	End Property
	
	
	Public Property ExpKey() As String
		Get
			ExpKey = m_sExpKey
		End Get
		Set(ByVal Value As String)
			m_sExpKey = Value
		End Set
	End Property
	
	
	Public Property Product() As String
		Get
			Product = m_sProduct
		End Get
		Set(ByVal Value As String)
			m_sProduct = Value
		End Set
	End Property
	
	
	Public Property Server() As String
		Get
			Server = m_sServer
		End Get
		Set(ByVal Value As String)
			m_sServer = Value
		End Set
	End Property
	
	
	Public Property BanUnderLevel() As Byte
		Get
			BanUnderLevel = m_byBanUnderLevel
		End Get
		Set(ByVal Value As Byte)
			m_byBanUnderLevel = Value
		End Set
	End Property
	
	
	Public Property BanD2UnderLevel() As Byte
		Get
			BanD2UnderLevel = m_byBanD2UnderLevel
		End Get
		Set(ByVal Value As Byte)
			m_byBanD2UnderLevel = Value
		End Set
	End Property
	
	
	Public Property BanUnderLevelMsg() As String
		Get
			BanUnderLevelMsg = m_sBanUnderLevelMsg
		End Get
		Set(ByVal Value As String)
			m_sBanUnderLevelMsg = Value
		End Set
	End Property
	
	
	Public Property KickOnYell() As Byte
		Get
			KickOnYell = m_byKickOnYell
		End Get
		Set(ByVal Value As Byte)
			m_byKickOnYell = Value
		End Set
	End Property
	
	
	Public Property HashFilePath() As String
		Get
			HashFilePath = m_sHashFilePath
		End Get
		Set(ByVal Value As String)
			m_sHashFilePath = Value
		End Set
	End Property
	
	
	Public Property BNLS() As Boolean
		Get
			BNLS = m_bBNLS
		End Get
		Set(ByVal Value As Boolean)
			m_bBNLS = Value
		End Set
	End Property
	
	
	Public Property LoadMode() As Boolean
		Get
			LoadMode = m_bLoadMode
		End Get
		Set(ByVal Value As Boolean)
			m_bLoadMode = Value
		End Set
	End Property
	
	
	Public Property Logging() As Byte
		Get
			Logging = m_byLogging
		End Get
		Set(ByVal Value As Byte)
			m_byLogging = Value
		End Set
	End Property
	
	
	Public Property Spoof() As Byte
		Get
			Spoof = m_bySpoof
		End Get
		Set(ByVal Value As Byte)
			m_bySpoof = Value
		End Set
	End Property
	
	
	Public Property DisableMP3Commands() As Boolean
		Get
			DisableMP3Commands = m_bDisableMP3Commands
		End Get
		Set(ByVal Value As Boolean)
			m_bDisableMP3Commands = Value
		End Set
	End Property
	
	
	Public Property UseUDP() As Boolean
		Get
			UseUDP = m_bUseUDP
		End Get
		Set(ByVal Value As Boolean)
			m_bUseUDP = Value
		End Set
	End Property
	
	
	Public Property FlashOnEvents() As Byte
		Get
			FlashOnEvents = m_byFlashOnEvents
		End Get
		Set(ByVal Value As Byte)
			m_byFlashOnEvents = Value
		End Set
	End Property
	
	
	Public Property UseProxy() As Boolean
		Get
			UseProxy = m_bUseProxy
		End Get
		Set(ByVal Value As Boolean)
			m_bUseProxy = Value
		End Set
	End Property
	
	
	Public Property ProxyPort() As Integer
		Get
			ProxyPort = m_lProxyPort
		End Get
		Set(ByVal Value As Integer)
			m_lProxyPort = Value
		End Set
	End Property
	
	
	Public Property ProxyStatus() As modEnum.enuProxyStatus
		Get
			ProxyStatus = m_eProxyStatus
		End Get
		Set(ByVal Value As modEnum.enuProxyStatus)
			m_eProxyStatus = Value
		End Set
	End Property
	
	
	Public Property ProxyIP() As String
		Get
			ProxyIP = m_sProxyIP
		End Get
		Set(ByVal Value As String)
			m_sProxyIP = Value
		End Set
	End Property
	
	
	Public Property ProxyIsSocks5() As Boolean
		Get
			ProxyIsSocks5 = m_bProxyIsSocks5
		End Get
		Set(ByVal Value As Boolean)
			m_bProxyIsSocks5 = Value
		End Set
	End Property
	
	
	Public Property MaxBacklogSize() As Integer
		Get
			MaxBacklogSize = m_lMaxBacklogSize
		End Get
		Set(ByVal Value As Integer)
			m_lMaxBacklogSize = Value
		End Set
	End Property
	
	
	Public Property MaxLogFileSize() As Integer
		Get
			MaxLogFileSize = m_lMaxLogfileSize
		End Get
		Set(ByVal Value As Integer)
			m_lMaxLogfileSize = Value
		End Set
	End Property
	
	
	Public Property UseGreet() As Boolean
		Get
			UseGreet = m_bUseGreet
		End Get
		Set(ByVal Value As Boolean)
			m_bUseGreet = Value
		End Set
	End Property
	
	
	Public Property GreetMsg() As String
		Get
			GreetMsg = m_sGreetMsg
		End Get
		Set(ByVal Value As String)
			m_sGreetMsg = Value
		End Set
	End Property
	
	
	Public Property UseRealm() As Boolean
		Get
			UseRealm = m_bUseRealm
		End Get
		Set(ByVal Value As Boolean)
			m_bUseRealm = Value
		End Set
	End Property
	
	
	Public Property WhisperGreet() As Boolean
		Get
			WhisperGreet = m_bWhisperGreet
		End Get
		Set(ByVal Value As Boolean)
			m_bWhisperGreet = Value
		End Set
	End Property
	
	
	Public Property NoTray() As Boolean
		Get
			NoTray = m_bNoTray
		End Get
		Set(ByVal Value As Boolean)
			m_bNoTray = Value
		End Set
	End Property
	
	
	Public Property ShowOfflineFriends() As Boolean
		Get
			ShowOfflineFriends = m_bShowOfflineFriends
		End Get
		Set(ByVal Value As Boolean)
			m_bShowOfflineFriends = Value
		End Set
	End Property
	
	
	Public Property ReconnectDelay() As Integer
		Get
			ReconnectDelay = m_lReconnectDelay
		End Get
		Set(ByVal Value As Integer)
			m_lReconnectDelay = Value
		End Set
	End Property
	
	
	Public Property UsingDirectFList() As Boolean
		Get
			UsingDirectFList = Config.FriendsListTab
		End Get
		Set(ByVal Value As Boolean)
			Config.FriendsListTab = Value
		End Set
	End Property
	
	
	Public Property AutofilterMS() As Integer
		Get
			AutofilterMS = m_lAutofilterMS
		End Get
		Set(ByVal Value As Integer)
			m_lAutofilterMS = Value
		End Set
	End Property
	
	
	Public Property NoColoring() As Boolean
		Get
			NoColoring = m_bNoColoring
		End Get
		Set(ByVal Value As Boolean)
			m_bNoColoring = Value
		End Set
	End Property
	
	
	Public Property ScriptMessageVeto() As Boolean
		Get
			ScriptMessageVeto = m_bScriptMessageVeto
		End Get
		Set(ByVal Value As Boolean)
			m_bScriptMessageVeto = Value
		End Set
	End Property
	
	
	Public Property NoAutocompletion() As Boolean
		Get
			NoAutocompletion = m_bNoAutocompletion
		End Get
		Set(ByVal Value As Boolean)
			m_bNoAutocompletion = Value
		End Set
	End Property
	
	
	Public Property SecondsToIdle() As Integer
		Get
			SecondsToIdle = m_lSecondsToIdle
		End Get
		Set(ByVal Value As Integer)
			m_lSecondsToIdle = Value
		End Set
	End Property
	
	
	Public Property Trigger() As String
		Get
			If (m_NoSupportMultiCharTrigger) Then
				Trigger = Left(m_sTrigger, 1)
			Else
				Trigger = m_sTrigger
			End If
			
		End Get
		Set(ByVal Value As String)
			'Triggers longer or shorter than 1 character will break some stuff
			m_sTrigger = Value
		End Set
	End Property
	
	Public ReadOnly Property TriggerLong() As String
		Get
			TriggerLong = m_sTrigger
		End Get
	End Property
	
	
	Public Property IPBans() As Boolean
		Get
			IPBans = m_bIPBans
		End Get
		Set(ByVal Value As Boolean)
			m_bIPBans = Value
		End Set
	End Property
	
	
	Public Property PlugBan() As Boolean
		Get
			PlugBan = m_bPlugBan
		End Get
		Set(ByVal Value As Boolean)
			m_bPlugBan = Value
		End Set
	End Property
	
	
	Public Property ClientBans() As Boolean
		Get
			ClientBans = m_bClientBans
		End Get
		Set(ByVal Value As Boolean)
			m_bClientBans = Value
		End Set
	End Property
	
	
	Public Property TSSetting() As Byte
		Get
			TSSetting = m_byTSSetting
		End Get
		Set(ByVal Value As Byte)
			m_byTSSetting = Value
		End Set
	End Property
	
	
	Public Property HomeChannel() As String
		Get
			HomeChannel = m_sHomeChannel
		End Get
		Set(ByVal Value As String)
			m_sHomeChannel = Value
		End Set
	End Property
	
	
	Public Property LastChannel() As String
		Get
			LastChannel = m_sLastChannel
		End Get
		Set(ByVal Value As String)
			m_sLastChannel = Value
		End Set
	End Property
	
	
	Public Property PublicChannels() As Collection
		Get
			PublicChannels = m_colPublicChannels
		End Get
		Set(ByVal Value As Collection)
			m_colPublicChannels = Value
		End Set
	End Property
	
	
	Public Property QuietTime() As Boolean
		Get
			QuietTime = m_bQuietTime
		End Get
		Set(ByVal Value As Boolean)
			m_bQuietTime = Value
		End Set
	End Property
	
	
	Public Property UseBackupChan() As Boolean
		Get
			UseBackupChan = m_bUseBackupChan
		End Get
		Set(ByVal Value As Boolean)
			m_bUseBackupChan = Value
		End Set
	End Property
	
	
	Public Property BackupChan() As String
		Get
			BackupChan = m_sBackupChan
		End Get
		Set(ByVal Value As String)
			m_sBackupChan = Value
		End Set
	End Property
	
	
	Public Property IB_On() As Byte
		Get
			IB_On = m_byIB_On
		End Get
		Set(ByVal Value As Byte)
			m_byIB_On = Value
		End Set
	End Property
	
	
	Public Property IB_Wait() As Short
		Get
			IB_Wait = m_iIB_Wait
		End Get
		Set(ByVal Value As Short)
			m_iIB_Wait = Value
		End Set
	End Property
	
	
	Public Property IB_Kick() As Boolean
		Get
			IB_Kick = m_bIB_Kick
		End Get
		Set(ByVal Value As Boolean)
			m_bIB_Kick = Value
		End Set
	End Property
	
	
	Public Property ChannelPassword() As String
		Get
			ChannelPassword = m_sChannelPassword
		End Get
		Set(ByVal Value As String)
			m_sChannelPassword = Value
		End Set
	End Property
	
	
	Public Property ChannelPasswordDelay() As Byte
		Get
			ChannelPasswordDelay = m_byChannelPasswordDelay
		End Get
		Set(ByVal Value As Byte)
			m_byChannelPasswordDelay = Value
		End Set
	End Property
	
	
	Public Property BanEvasion() As Boolean
		Get
			BanEvasion = m_bBanEvasion
		End Get
		Set(ByVal Value As Boolean)
			m_bBanEvasion = Value
		End Set
	End Property
	
	
	Public Property LogDBActions() As Boolean
		Get
			LogDBActions = m_bLogDBActions
		End Get
		Set(ByVal Value As Boolean)
			m_bLogDBActions = Value
		End Set
	End Property
	
	
	Public Property LogCommands() As Boolean
		Get
			LogCommands = m_bLogCommands
		End Get
		Set(ByVal Value As Boolean)
			m_bLogCommands = Value
		End Set
	End Property
	
	
	Public Property BotOwner() As String
		Get
			BotOwner = m_sBotOwner
		End Get
		Set(ByVal Value As String)
			m_sBotOwner = Value
		End Set
	End Property
	
	
	Public Property BanPeons() As Boolean
		Get
			BanPeons = m_bBanPeons
		End Get
		Set(ByVal Value As Boolean)
			m_bBanPeons = Value
		End Set
	End Property
	
	
	Public Property WhisperCmds() As Boolean
		Get
			WhisperCmds = m_bWhisperCmds
		End Get
		Set(ByVal Value As Boolean)
			m_bWhisperCmds = Value
		End Set
	End Property
	
	
	Public Property DoScroll() As Boolean
		Get
			DoScroll = m_bDoScroll
		End Get
		Set(ByVal Value As Boolean)
			m_bDoScroll = Value
		End Set
	End Property
	
	
	Public Property LockChat() As Boolean
		Get
			LockChat = m_bLockChat
		End Get
		Set(ByVal Value As Boolean)
			m_bLockChat = Value
		End Set
	End Property
	
	
	Public Property JoinWatch() As Short
		Get
			JoinWatch = m_iJoinWatch
		End Get
		Set(ByVal Value As Short)
			m_iJoinWatch = Value
		End Set
	End Property
	
	
	Public Property BNLSServer() As String
		Get
			BNLSServer = m_sBNLSServer
		End Get
		Set(ByVal Value As String)
			m_sBNLSServer = Value
		End Set
	End Property
	
	
	Public Property NoRTBAutomaticCopy() As Boolean
		Get
			NoRTBAutomaticCopy = m_bNoRTBAutomaticCopy
		End Get
		Set(ByVal Value As Boolean)
			m_bNoRTBAutomaticCopy = Value
		End Set
	End Property
	
	
	Public Property GatewayConventions() As Byte
		Get
			GatewayConventions = m_byGatewayConventions
		End Get
		Set(ByVal Value As Byte)
			m_byGatewayConventions = Value
		End Set
	End Property
	
	
	Public Property UseD2Naming() As Boolean
		Get
			UseD2Naming = m_bUseD2Naming
		End Get
		Set(ByVal Value As Boolean)
			m_bUseD2Naming = Value
		End Set
	End Property
	
	
	Public Property D2NamingFormat() As String
		Get
			D2NamingFormat = m_sD2NamingFormat
		End Get
		Set(ByVal Value As String)
			m_sD2NamingFormat = Value
		End Set
	End Property
	
	
	Public Property Clan() As String
		Get
			Clan = m_sClan
		End Get
		Set(ByVal Value As String)
			If (m_ScriptLock = False) Then
				m_sClan = Value
			End If
		End Set
	End Property
	
	
	Public Property Gateway() As String
		Get
			Gateway = m_sGateway
		End Get
		Set(ByVal Value As String)
			If (m_ScriptLock = False) Then
				m_sGateway = Value
			End If
		End Set
	End Property
	
	
	Public Property Realm() As String
		Get
			Realm = m_sRealm
		End Get
		Set(ByVal Value As String)
			m_sRealm = Value
		End Set
	End Property
	
	
	Public Property NoSupportMultiCharTrigger() As Boolean
		Get
			NoSupportMultiCharTrigger = m_NoSupportMultiCharTrigger
		End Get
		Set(ByVal Value As Boolean)
			m_NoSupportMultiCharTrigger = Value
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
	
	
	Public Property ChatDelay() As Integer
		Get
			ChatDelay = m_ChatDelay
		End Get
		Set(ByVal Value As Integer)
			m_ChatDelay = Value
		End Set
	End Property
	
	
	Public Property ScriptLock() As Boolean
		Get
			ScriptLock = m_ScriptLock
		End Get
		Set(ByVal Value As Boolean)
			If (m_ScriptLock = False) Then
				m_ScriptLock = Value
			End If
			
		End Set
	End Property
	
	
	Public Property DefaultShitlistGroup() As String
		Get
			DefaultShitlistGroup = m_shitlist_group
		End Get
		Set(ByVal Value As String)
			m_shitlist_group = Value
			
		End Set
	End Property
	
	
	Public Property DefaultSafelistGroup() As String
		Get
			DefaultSafelistGroup = m_safelist_group
		End Get
		Set(ByVal Value As String)
			m_safelist_group = Value
			
		End Set
	End Property
	
	
	Public Property DefaultTagbansGroup() As String
		Get
			DefaultTagbansGroup = m_tagbans_group
		End Get
		Set(ByVal Value As String)
			m_tagbans_group = Value
			
		End Set
	End Property
	
	
	Public Property AutoCompletePostfix() As String
		Get
			AutoCompletePostfix = m_autocomplete_postfix
		End Get
		Set(ByVal Value As String)
			m_autocomplete_postfix = Value
		End Set
	End Property
	
	
	
	Public Property ShowStatsIcons() As Boolean
		Get
			ShowStatsIcons = m_show_stats_icons
		End Get
		Set(ByVal Value As Boolean)
			m_show_stats_icons = Value
		End Set
	End Property
	
	
	
	Public Property ShowFlagsIcons() As Boolean
		Get
			ShowFlagsIcons = m_show_flags_icons
		End Get
		Set(ByVal Value As Boolean)
			m_show_flags_icons = Value
		End Set
	End Property
	
	
	Public Property RetainOldBans() As Boolean
		Get
			RetainOldBans = m_retain_old_bans
		End Get
		Set(ByVal Value As Boolean)
			m_retain_old_bans = Value
		End Set
	End Property
	
	
	Public Property StoreAllBans() As Boolean
		Get
			StoreAllBans = m_store_all_bans
		End Get
		Set(ByVal Value As Boolean)
			m_store_all_bans = Value
		End Set
	End Property
	
	
	Public Property UseAltBnls() As Boolean
		Get
			UseAltBnls = m_use_alt_bnls
		End Get
		Set(ByVal Value As Boolean)
			m_use_alt_bnls = Value
		End Set
	End Property
	
	
	Public Property CaseSensitiveFlags() As Boolean
		Get
			CaseSensitiveFlags = m_case_sensitive_flags
		End Get
		Set(ByVal Value As Boolean)
			m_case_sensitive_flags = Value
		End Set
	End Property
	
	
	Public Property JoinLeaveMessages() As Boolean
		Get
			JoinLeaveMessages = Not JoinMessagesOff
		End Get
		Set(ByVal Value As Boolean)
			JoinMessagesOff = Not Value
		End Set
	End Property
	
	' ############## QUEUE VARIABLES
	
	Public Property QueueMaxCredits() As Integer
		Get
			QueueMaxCredits = m_lMaxCredits
		End Get
		Set(ByVal Value As Integer)
			m_lMaxCredits = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerPacket() As Integer
		Get
			QueueCostPerPacket = m_lCostPerPacket
		End Get
		Set(ByVal Value As Integer)
			m_lCostPerPacket = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerByte() As Integer
		Get
			QueueCostPerByte = m_lCostPerByte
		End Get
		Set(ByVal Value As Integer)
			m_lCostPerByte = Value
		End Set
	End Property
	
	
	Public Property QueueCostPerByteOverThreshhold() As Integer
		Get
			QueueCostPerByteOverThreshhold = m_lCostPerByteOverThreshhold
		End Get
		Set(ByVal Value As Integer)
			m_lCostPerByteOverThreshhold = Value
		End Set
	End Property
	
	
	Public Property QueueStartingCredits() As Integer
		Get
			QueueStartingCredits = m_lStartingCredits
		End Get
		Set(ByVal Value As Integer)
			m_lStartingCredits = Value
		End Set
	End Property
	
	
	Public Property QueueThreshholdBytes() As Integer
		Get
			QueueThreshholdBytes = m_lThreshholdBytes
		End Get
		Set(ByVal Value As Integer)
			m_lThreshholdBytes = Value
		End Set
	End Property
	
	
	Public Property QueueCreditRate() As Integer
		Get
			QueueCreditRate = m_lCreditRate
		End Get
		Set(ByVal Value As Integer)
			m_lCreditRate = Value
		End Set
	End Property
	' ############## END QUEUE VARIABLES
End Class