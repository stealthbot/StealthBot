Option Strict Off
Option Explicit On
Module modGlobals
	'modGlobals - global variables
	
	Public CVERSION As String
	
	Public Config As New clsConfig
	
	'Timer variables
	Public uTicks As Integer
	Public ReconnectTimerID As Integer
	Public ExReconnectTimerID As Integer
	Public SCReloadTimerID As Integer
	Public ClanAcceptTimerID As Integer
	Public QueueTimerID As Integer
	Public rtbWhispersVisible As Boolean
	Public cboSendHadFocus As Boolean
	Public cboSendSelStart As Integer
	Public cboSendSelLength As Integer
	Public LogPacketTraffic As Boolean
	Public cfgVersion As Integer
	
	Public VoteInitiator As udtGetAccessResponse
	Public ProductList(12) As udtProductInfo
	
	Public g_Queue As New clsQueue
	Public g_OSVersion As New clsOSVersion
	Public SharedScriptSupport As New clsScriptSupportClass
	Public BNCSBuffer As New clsBNCSRecvBuffer
	Public BNLSBuffer As New clsBNLSRecvBuffer
	Public MCPBuffer As New clsBNLSRecvBuffer
	Public GErrorHandler As clsErrorHandler
	
	Public ConfigOverride As String
	Public CommandLine As String
	Public lLauncherVersion As Integer
	
	Public rtbChatLength As Integer
	
	
	
	Public g_Connected As Boolean
	Public g_Online As Boolean
	Public AwaitingSystemKeys As Byte
	Public AwaitingSelfRemoval As Byte
	
	Public g_Quotes As New clsQuotesObj
	Public g_Channel As New clsChannelObj
	Public g_Logger As New clsLogger
	Public g_Clan As New clsClanObj
	Public g_Friends As New Collection
	Public g_BNCSQueue As New clsBNCSQueue
	
	Public UserCancelledConnect As Boolean
	
	'For closing the bot with the quit command
	Public BotIsClosing As Boolean
	
	'To determine when to reset the BNLS list for the auto-BNLS server locator
	Public BNLSFinderGotList As Boolean
	Public BNLSFinderEntries() As String
	Public BNLSFinderIndex As Short
	Public BNLSFinderLatest As String
	
	Public Const ERROR_FINDBNLSSERVER As Short = 12345
	
	
	
	'VARIABLES
	Public CurrentUsername As String
	Public ProfileRequest As Boolean
	Public Protect As Boolean
	'UPGRADE_WARNING: Lower bound of array QC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public QC(9) As String
	
	Public JoinMessagesOff As Boolean
	Public Filters As Boolean
	Public mail As Boolean
	Public ProtectMsg As String
	
	Public LastWhisper As String
	Public LastWhisperFromTime As Date
	Public LastWhisperTo As String
	Public Caching As Boolean
	Public Dii As Boolean
	Public PPL As Boolean
	Public PPLRespondTo As String
	Public MyFlags As Integer
	Public BanCount As Short
	Public g_lastQueueUser As String
	Public PassedClanMotdCheck As Boolean
	Public g_request_receipt As Boolean
	
	Public WatchUser As String
	Public BotVars As clsBotVars
	
	Public colLastSeen As Collection
	Public colProfiles As Collection
	
	'ARRAYS
	Public Phrases() As String
	Public PhraseBans As Boolean
	Public SuppressProfileOutput As Boolean
	Public SpecificProfileKey As String
	'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Catch_Renamed() As String
	Public gBans() As udtBanList
	Public gOutFilters() As udtOutFilters
	Public gFilters() As String
	Public DB() As udtDatabase
	
	Public AutoModSafelistValue As Short
	
	Public ExReconTicks As Integer
	Public ExReconMinutes As Integer
	
	Public VoteDuration As Short
	Public PBuffer As New clsDataBuffer
	Public ListToolTip As clsCTooltip
	
	
	'cboSend_KeyEvents
	Public MatchIndex As Integer
	
	
	
	Public Function isDebug() As Boolean
		If (MDebug("all")) Then
			isDebug = True
		End If
	End Function
End Module