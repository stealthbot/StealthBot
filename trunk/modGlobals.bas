Attribute VB_Name = "modGlobals"
'modGlobals - global variables
Option Explicit

Public CVERSION As String

Public Config As New clsConfig
Public Database As New clsDatabase

'Timer variables
Public ReconnectTimerID As Long
Public SCReloadTimerID As Long
Public QueueTimerID As Long

Public LogPacketTraffic As Boolean
Public cfgVersion As Long

Public VoteInitiator As udtUserAccess
Public ProductList(12) As udtProductInfo

Public g_Queue As New clsQueue
Public g_OSVersion As New clsOSVersion
Public SharedScriptSupport As New clsScriptSupportClass
Public ReceiveBuffer(0 To 2) As clsDataBuffer
Public ProxyConnInfo(0 To 2) As udtProxyConnectionInfo

Public ConfigOverride As String
Public CommandLine As String
Public lLauncherVersion As Long

Public g_Connected As Boolean
Public g_ConnectionAlive As Boolean
Public g_Online As Boolean

Public g_Quotes As New clsQuotesObj
Public g_Logger As New clsLogger
Public g_BNCSQueue As New clsBNCSQueue
Public g_Channel As clsChannelObj
Public g_Friends As Collection
Public g_Clan As clsClanObj
Public g_Color As clsColor

Public AutoReconnectActive As Boolean
Public AutoReconnectTry    As Long
Public AutoReconnectTicks  As Long
Public AutoReconnectIn     As Long

'For closing the bot with the quit command
Public BotIsClosing As Boolean

'To determine when to reset the BNLS list for the auto-BNLS server locator
Public BNLSFinderGotList   As Boolean
Public BNLSFinderEntries() As String
Public BNLSFinderIndex     As Integer
Public BNLSFinderLatest    As String

Public Const ERROR_FINDBNLSSERVER As Integer = 12345



'VARIABLES
Public CurrentUsername As String
Public Protect As Boolean
Public QC(1 To 9) As String
Public ServerRequests() As udtServerRequest

Public JoinMessagesOff As Boolean
Public Filters As Boolean
Public mail As Boolean
Public ProtectMsg As String

Public LastWhisper As String
Public LastWhisperFromTime As Date
Public LastWhisperTo As String
Public Caching As Boolean
Public Dii As Boolean
Public MyFlags As Long
Public BanCount As Integer
Public g_request_receipt As Boolean

Public WatchUser As String
Public BotVars As clsBotVars

Public colLastSeen As Collection
Public colProfiles As Collection

'ARRAYS
Public Phrases() As String
Public PhraseBans As Boolean
Public Catch() As String
Public gBans() As udtBanList
Public gOutFilters() As udtOutFilters
Public gFilters() As String
Public g_Blocklist() As String

Public AutoModSafelistValue As Integer

Public ExReconTicks As Long
Public ExReconMinutes As Long

Public VoteDuration As Integer
Public ListToolTip As clsCTooltip


'cboSend_KeyEvents
Public MatchIndex  As Long



Public Function isDebug() As Boolean
    If (MDebug("all")) Then
        isDebug = True
    End If
End Function

