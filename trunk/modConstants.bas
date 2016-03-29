Attribute VB_Name = "modConstants"
Option Explicit

Public Const REVISION As String = ""
Public Const VERCODE As String = "2697"
Public Const BNCSUTIL_VERSION As String = "1.3.1"
Public Const CONFIG_VERSION As String = "5"

' URLs
Public Const BETA_AUTH_URL As String = _
    "http://www.stealthbot.net/board/sbauth.php?username="
Public Const BETA_AUTH_URL_CRC32 As Long = 716038006

' Files
Public Const FILE_COMMANDS = "Commands.xml"
Public Const FILE_USERDB = "Users.txt"
Public Const FILE_QUOTES = "Quotes.txt"
Public Const FILE_CAUGHT_PHRASES = "CaughtPhrases.htm"
Public Const FILE_FILTERS = "Filters.ini"
Public Const FILE_PHRASE_BANS = "PhraseBans.txt"
Public Const FILE_MAILDB = "Mail.dat"
Public Const FILE_SCRIPT_INI = "Scripts.ini"
Public Const FILE_CATCH_PHRASES = "CatchPhrases.txt"
Public Const FILE_QUICK_CHANNELS = "QuickChannels.txt"
Public Const FILE_COLORS = "Colors.sclf"
Public Const FILE_CREV_INI = "CheckRevision.ini"
Public Const FILE_WARDEN_INI = "Warden.ini"
Public Const FILE_BNLS_LIST = "AdditionalBNLSservers.txt"
Public Const FILE_SERVER_LIST = "Servers.txt"

' Product codes
Public Const PRODUCT_STAR = "STAR"
Public Const PRODUCT_SEXP = "SEXP"
Public Const PRODUCT_W2BN = "W2BN"
Public Const PRODUCT_D2DV = "D2DV"
Public Const PRODUCT_D2XP = "D2XP"
Public Const PRODUCT_WAR3 = "WAR3"
Public Const PRODUCT_W3XP = "W3XP"
Public Const PRODUCT_SSHR = "SSHR"
Public Const PRODUCT_JSTR = "JSTR"
Public Const PRODUCT_DSHR = "DSHR"
Public Const PRODUCT_DRTL = "DRTL"
Public Const PRODUCT_CHAT = "CHAT"


Public Const BNET_MSG_LENGTH = 233

' Colors
Public Const COLOR_TEAL& = &H99CC00
Public Const COLOR_BLUE& = &HCC9900


' Icons
Public Const ICON_START_WAR3& = 62
Public Const ICON_START_W3XP& = 25
Public Const ICON_START_D2& = 110
Public Const ICON_START_SC& = 117
Public Const ICON_START_W2& = 128
Public Const ICON_START_DRTL& = 139
Public Const ICON_START_DSHR& = 142

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


' Misc

Public Const LVM_FIRST = &H1000&
Public Const LVM_HITTEST = LVM_FIRST + 18

Public Const CPTALK As Byte = 0
Public Const CPEMOTE As Byte = 1
Public Const CPWHISPER As Byte = 2

Public Const GWL_WNDPROC = (-4)

Public Const WM_VSCROLL = &H115
Public Const SB_VERT As Long = 1
Public Const EM_GETTHUMB = &HBE
Public Const SB_HORZ As Long = 0
Public Const SB_BOTH As Long = 3
Public Const SB_THUMBPOSITION = &H4

Public Const BTRUE& = 1
Public Const BFALSE& = 0

Public Const IF_AWAITING_CHPW& = 3
Public Const IF_SUBJECT_TO_IDLEBANS = 9
Public Const IF_CHPW_AND_IDLEBANS = 12

Public Const LOAD_SAFELIST = 1
Public Const LOAD_FILTERS = 2
Public Const LOAD_PHRASES = 3
Public Const LOAD_DB = 4
