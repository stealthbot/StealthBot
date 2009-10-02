Attribute VB_Name = "modGlobals"
'modGlobals - global variables and constants
Option Explicit

'CONSTANTS
Public CVERSION As String
Public Const REVISION As String = ""
Public Const VERCODE As String = "2697"
Public Const BNCSUTIL_VERSION As String = "1.3.1"
Public Const CONFIG_VERSION As String = "5"

Public Const BETA_AUTH_URL As String = _
    "http://www.stealthbot.net/board/sbauth.php?username="
Public Const BETA_AUTH_URL_CRC32 As Long = 716038006

Public Const COLOR_TEAL& = &H99CC00
Public Const COLOR_BLUE& = &HCC9900
Public Const GWL_WNDPROC = (-4)

Public Const WM_VSCROLL = &H115
Public Const SB_VERT As Long = 1
Public Const EM_GETTHUMB = &HBE
Public Const SB_HORZ As Long = 0
Public Const SB_BOTH As Long = 3
Public Const SB_THUMBPOSITION = &H4

Public Const ID_USER = &H1
Public Const ID_JOIN = &H2
Public Const ID_LEAVE = &H3
Public Const ID_WHISPER = &H4
Public Const ID_TALK = &H5
Public Const ID_BROADCAST = &H6
Public Const ID_CHANNEL = &H7
Public Const ID_USERFLAGS = &H9
Public Const ID_WHISPERSENT = &HA
Public Const ID_CHANNELFULL = &HD
Public Const ID_CHANNELDOESNOTEXIST = &HE
Public Const ID_CHANNELRESTRICTED = &HF
Public Const ID_INFO = &H12
Public Const ID_ERROR = &H13
Public Const ID_EMOTE = &H17
' Additional event constants for logging
Public Const ID_CONNECTED = &H18
Public Const ID_DISCONNECTED = &H19


Public Const USER_BLIZZREP& = &H1
Public Const USER_CHANNELOP& = &H2
Public Const USER_SPEAKER& = &H4
Public Const USER_SYSOP& = &H8
Public Const USER_NOUDP& = &H10
Public Const USER_BEEPENABLED& = &H100
Public Const USER_KBKOFFICIAL& = &H1000
Public Const USER_JAILED& = &H100000
Public Const USER_SQUELCHED& = &H20
Public Const USER_PGLPLAYER& = &H200
Public Const USER_GFPLAYER& = &H200000
Public Const USER_GUEST& = &H40
Public Const USER_PGLOFFICIAL& = &H400
Public Const USER_KBKPLAYER& = &H800

Public Const ICON_START_WAR3& = 62
Public Const ICON_START_W3XP& = 25

Public Const LVM_FIRST = &H1000&
Public Const LVM_HITTEST = LVM_FIRST + 18

Public Const CPTALK As Byte = 0
Public Const CPEMOTE As Byte = 1
Public Const CPWHISPER As Byte = 2

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

Public rtbChatLength As Long

Public Const BTRUE& = 1
Public Const BFALSE& = 0

Public Const IF_AWAITING_CHPW& = 3
Public Const IF_SUBJECT_TO_IDLEBANS = 9
Public Const IF_CHPW_AND_IDLEBANS = 12

Public Const LOAD_SAFELIST = 1
Public Const LOAD_FILTERS = 2
Public Const LOAD_PHRASES = 3
Public Const LOAD_DB = 4

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
Public QC(0 To 8) As String

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

Public HashCmd As String
Public MPQName As String
Public Checksum As Long
Public EXEInfo As String
Public SessionKey As Long
Public GTC As Long
Public AutoModSafelistValue As Integer

Public ExReconTicks As Long
Public ExReconMinutes As Long

Public VoteDuration As Integer
'Public MonitorForm As frmMonitor
Public PBuffer As New clsDataBuffer
Public ListToolTip As clsCTooltip
Public AwaitingEmailReg As Byte
Public g_ThisIconCode As Integer

'Mode1 Values
Public Const BVT_VOTE_ADD As Byte = 1
Public Const BVT_VOTE_START As Byte = 0
Public Const BVT_VOTE_END As Byte = 2
Public Const BVT_VOTE_TALLY As Byte = 3

'Mode2 Values
Public Const BVT_VOTE_ADDYES As Byte = 1
Public Const BVT_VOTE_ADDNO As Byte = 0
Public Const BVT_VOTE_BAN As Byte = 2
Public Const BVT_VOTE_STD As Byte = 3
Public Const BVT_VOTE_KICK As Byte = 4
Public Const BVT_VOTE_CANCEL As Byte = 5

'Icon constants
Public Const ICSTAR = 1
Public Const ICSEXP = 2
Public Const ICD2DV = 3
Public Const ICD2XP = 4
Public Const ICW2BN = 5
Public Const ICWAR3 = 6
Public Const ICGAVEL = 7
Public Const ICUNKNOWN = 8
Public Const ICBLIZZ = 9
Public Const ICSYSOP = 103
Public Const ICCHAT = 10
Public Const ICDIABLO = 11
Public Const ICDIABLOSW = 12
Public Const ICSQUELCH = 13
Public Const ICJSTR = 14
Public Const ICWAR3X = 15
Public Const LAG_PLUG = 16
Public Const LAG_1 = 17
Public Const LAG_2 = 18
Public Const LAG_3 = 19
Public Const LAG_4 = 20
Public Const LAG_5 = 21
Public Const LAG_6 = 22
Public Const MONITOR_ONLINE = 23
Public Const MONITOR_OFFLINE = 24
Public Const ICSCSW = 25

'World Cyber Games icons
Public Const IC_WCG_PLAYER = 89
Public Const IC_WCG_REF = 90

'UPDATED WCG Icons
Public Const IC_WCRF = 104
Public Const IC_WCPL = 105
Public Const IC_WCGO = 106
Public Const IC_WCSI = 107
Public Const IC_WCBR = 108
Public Const IC_WCPG = 109

'PGTour icons
Public Const IC_PGT_A = 92
Public Const IC_PGT_B = 95
Public Const IC_PGT_C = 98
Public Const IC_PGT_D = 101

'Friends constants
Public Const FRL_OFFLINE& = &H0
Public Const FRL_NOTINCHAT& = &H1
Public Const FRL_INCHAT& = &H2
Public Const FRL_PUBLICGAME& = &H3
Public Const FRL_PRIVATEGAME& = &H5

Public Const FRS_NONE& = &H0
Public Const FRS_MUTUAL& = &H1
Public Const FRS_DND& = &H2
Public Const FRS_AWAY& = &H4

'FlashWindow constants
Public Const FLASHW_CAPTION = 1
Public Const FLASHW_TRAY = 2
Public Const FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY
Public Const FLASHW_TIMERNOFG = 12 Or FLASHW_ALL
Public Const FLW_SIZE = 20

'Key constants
Public Const KEY_PGDN = 34
Public Const KEY_PGUP = 33
Public Const KEY_HOME = 36
Public Const KEY_ALTN = 78
Public Const KEY_END = 35
Public Const KEY_INSERT = 45
Public Const KEY_ENTER = 13
Public Const KEY_SPACE = 32
Public Const KEY_V = 86
Public Const KEY_A = 65
Public Const KEY_S = 83
Public Const KEY_D = 68
Public Const KEY_B = 66
Public Const KEY_J = 74
Public Const KEY_I = 73
Public Const KEY_U = 85
Public Const KEY_DELETE = 46

'cboSend_KeyEvents
Public Highlighted As Boolean
Public MatchIndex  As Long

'Nagle algorithm constants
Public Const IPPROTO_TCP& = 6   'socketlevel
Public Const TCP_NODELAY& = &H1 'optname
Public Const NAGLE_OPTLEN& = 4  'optlen
Public Const NAGLE_ON& = 1      'optval
Public Const NAGLE_OFF& = 0     'optval

Public Const SO_KEEPALIVE = &H8

Public Const CAP_SHOW$ = "v"
Public Const CAP_HIDE$ = "^"
Public Const SHOW_SIZE& = 255
Public Const HIDE_SIZE& = 2300
Public Const TIP_HIDE$ = "Hide the whisper window"
Public Const TIP_SHOW$ = "Show the whisper window"

Public Const LVW_BUTTON_CHANNEL = 0
Public Const LVW_BUTTON_FRIENDS = 1
Public Const LVW_BUTTON_CLAN = 2

'BNCSUTIL NLS buffer size constants
Public Const NLS_ACCOUNTCREATE_     As Long = 65
Public Const NLS_ACCOUNTLOGON_      As Long = 33
Public Const NLS_GET_S_             As Long = 32
Public Const NLS_GET_V_             As Long = 32
Public Const NLS_GET_A_             As Long = 32
Public Const NLS_GET_K_             As Long = 40
Public Const NLS_GET_M1_            As Long = 20

Public Function isDebug() As Boolean

    ' ...
    If (MDebug("all")) Then
        isDebug = True
    End If

End Function

