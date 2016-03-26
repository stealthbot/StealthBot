Attribute VB_Name = "modGlobals"
'modGlobals - global variables
Option Explicit

Public CVERSION As String

Public Config As New clsConfig

'Timer variables
Public uTicks As Long
Public ReconnectTimerID As Long
Public ExReconnectTimerID As Long
Public SCReloadTimerID As Long
Public ClanAcceptTimerID As Long
Public QueueTimerID As Long
Public rtbWhispersVisible As Boolean
Public cboSendHadFocus As Boolean
Public cboSendSelStart As Long
Public cboSendSelLength As Long
Public RealmError As Boolean
Public LogPacketTraffic As Boolean
Public cfgVersion As Long

Public VoteInitiator As udtGetAccessResponse

Public g_Queue As New clsQueue
Public g_OSVersion As New clsOSVersion
Public SharedScriptSupport As New clsScriptSupportClass
Public BNCSBuffer As New clsBNCSRecvBuffer
Public BNLSBuffer As New clsBNLSRecvBuffer
Public GErrorHandler As clsErrorHandler

Public ConfigOverride As String
Public CommandLine As String
Public lLauncherVersion As Long

Public rtbChatLength As Long



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
Public GotBNLSList As Boolean
Public LocatingAltBNLS As Boolean

'VARIABLES
Public CurrentUsername As String
Public ProfileRequest As Boolean
Public Protect As Boolean
Public QC(1 To 9) As String
Public PublicChannels As Collection
Public SkipUICEvents As Boolean

Public JoinMessagesOff As Boolean
Public Filters As Boolean
Public mail As Boolean
Public BotLoaded As Boolean
Public ProtectMsg As String

Public LastWhisper As String
Public LastWhisperFromTime As Date
Public LastWhisperTo As String
Public Caching As Boolean
Public Dii As Boolean
Public PPL As Boolean
Public PPLRespondTo As String
Public MyFlags As Long
Public BanCount As Integer
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
Public Catch() As String
Public gBans() As udtBanList
Public gOutFilters() As udtOutFilters
Public gFilters() As String
Public DB() As udtDatabase

Public AutoModSafelistValue As Integer

Public ExReconTicks As Long
Public ExReconMinutes As Long

Public VoteDuration As Integer
Public PBuffer As New clsDataBuffer
Public ListToolTip As clsCTooltip


'cboSend_KeyEvents
Public MatchIndex  As Long



Public Function isDebug() As Boolean
    If (MDebug("all")) Then
        isDebug = True
    End If
End Function

