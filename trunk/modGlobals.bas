Attribute VB_Name = "modGlobals"
'modGlobals - global variables
Option Explicit

'Timer variables
Public uTicks As Long
Public ForcedJoinsOn As Byte
Public ReconnectTimerID As Long
Public ExReconnectTimerID As Long
Public UnsquelchTimerID As Long
Public SCReloadTimerID As Long
Public ClanAcceptTimerID As Long
Public QueueTimerID As Long
Public AutoChatFilter As Long
Public rtbWhispersVisible As Boolean
Public cboSendHadFocus As Boolean
Public cboSendSelStart As Long
Public cboSendSelLength As Long
Public RealmError As Boolean
Public PacketLogFilePath As String
Public LogPacketTraffic As Boolean
Public ScriptMenu_ParentID As Long
Public CfgVersion As Long

Public VoteInitiator As udtGetAccessResponse

Public g_Queue As New clsQueue
Public g_OSVersion As New clsOSVersion
Public SharedScriptSupport As New clsScriptSupportClass
Public BNCSBuffer As New clsBNCSRecvBuffer
Public BNLSBuffer As New clsBNLSRecvBuffer
Public GErrorHandler As clsErrorHandler

Public AttemptedFirstReconnect As Boolean

Public ConfigOverride As String
Public CommandLine As String
Public lLauncherVersion As Long

Public rtbChatLength As Long



Public g_Connected As Boolean
Public g_Online As Boolean
Public AwaitingSystemKeys As Byte
Public AwaitingSelfRemoval As Byte
'Public AttemptedNewVerbyte As Boolean

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

Public DisableMonitor As Boolean
Public JoinMessagesOff As Boolean
Public Filters As Boolean
Public mail As Boolean
Public Mimic As String
Public BotLoaded As Boolean
Public ProtectMsg As String

Public LastWhisper As String
Public LastWhisperFromTime As Date
Public LastWhisperTo As String
Public Caching As Boolean
Public UserTracker(14) As String
Public Dii As Boolean
Public PPL As Boolean
Public PPLRespondTo As String
Public MyFlags As Long
Public Unsquelching As Boolean
Public BanCount As Integer
Public g_lastQueueUser As String
Public PassedClanMotdCheck As Boolean
Public g_request_receipt As Boolean

Public LastAdd As Byte
Public WatchUser As String
Public bFlood As Boolean
Public BotVars As clsBotVars

Public colLastSeen As Collection
Public colProfiles As Collection

Public LastWhisperTime As Long

'ARRAYS
Public Phrases() As String
Public ClientBans() As String
Public PhraseBans As Boolean
Public SuppressProfileOutput As Boolean
Public SpecificProfileKey As String
Public Catch() As String
Public gChannel As udtChanList
Public gBans() As udtBanList
Public gOutFilters() As udtOutFilters
Public colSafelist As Collection
Public gFilters() As String
'Public Queue() As udtQueue
Public colQueue As Collection
Public DB() As udtDatabase
Public gFloodSafelist() As String
Public Last4Messages(0 To 3) As String

Public AutoModSafelistValue As Integer

Public ExReconTicks As Long
Public ExReconMinutes As Long

Public VoteDuration As Integer
'Public MonitorForm As frmMonitor
Public PBuffer As New clsDataBuffer
Public ListToolTip As clsCTooltip
Public AwaitingEmailReg As Byte
Public g_ThisIconCode As Integer


'cboSend_KeyEvents
Public Highlighted As Boolean
Public MatchIndex  As Long



Public Function isDebug() As Boolean
    If (MDebug("all")) Then
        isDebug = True
    End If
End Function

