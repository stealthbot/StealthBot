Attribute VB_Name = "modConstants"
Option Explicit

Public Const REVISION As String = ""
Public Const VERCODE As String = "2697"
Public Const BNCSUTIL_VERSION As String = "1.3.1"
Public Const CONFIG_VERSION As String = "6"

' URLs
Public Const BNLS_DEFAULT_SOURCE As String = _
    "http://stealthbot.net/sb/bnls.php"
    
Public Const VERBYTE_SOURCE As String = _
    "http://www.stealthbot.net/sb/verbytes/versionbytes.txt"

' Files
Public Const FILE_COMMANDS = "Commands.xml"
Public Const FILE_USERDB = "Users.txt"
Public Const FILE_MAILDB = "Mail.dat"
Public Const FILE_CAUGHT_PHRASES = "CaughtPhrases.htm"
Public Const FILE_FILTERS = "Filters.ini"
Public Const FILE_SCRIPT_INI = "Scripts.ini"
Public Const FILE_COLORS = "Colors.ini"
Public Const FILE_CREV_INI = "CheckRevision.ini"
Public Const FILE_WARDEN_INI = "Warden.ini"
Public Const FILE_PHRASE_BANS = "PhraseBans.txt"
Public Const FILE_CATCH_PHRASES = "CatchPhrases.txt"
Public Const FILE_QUICK_CHANNELS = "QuickChannels.txt"
Public Const FILE_QUOTE_LIST = "Quotes.txt"
Public Const FILE_BNLS_LIST = "AdditionalBNLSservers.txt"
Public Const FILE_SERVER_LIST = "Servers.txt"
Public Const FILE_KEY_LIST = "Keys.txt"

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

' prod icons
Public Const ICUNKNOWN = 1
Public Const ICSTAR = 2
Public Const ICSEXP = 3
Public Const ICD2DV = 4
Public Const ICD2XP = 5
Public Const ICW2BN = 6
Public Const ICWAR3 = 7
Public Const ICWAR3X = 8
Public Const ICDIABLO = 9
Public Const ICDIABLOSW = 10
Public Const ICJSTR = 11
Public Const ICSCSW = 12
Public Const ICCHAT = 13

' stats spawn icons
Public Const IC_STAR_SPAWN = 140
Public Const IC_JSTR_SPAWN = 141
Public Const IC_W2BN_SPAWN = 153
Public Const IC_DIAB_SPAWN = 166

' stats icon sequences
Public Const ICON_START_WAR3 = 76
Public Const ICON_START_W3XP = 39
Public Const ICON_START_D2 = 122
Public Const ICON_START_SC = 129
Public Const ICON_START_W2 = 142
Public Const ICON_START_D1 = 154

' flags icons
Public Const ICGAVEL = 14
Public Const ICBLIZZ = 15
Public Const ICSYSOP = 16
Public Const ICSQUELCH = 17
Public Const ICSPEAKER = 18
Public Const ICSPECS = 19

' flags icons: GF
Public Const IC_GF_OFFICIAL = 20
Public Const IC_GF_PLAYER = 21

' Lag icons
Public Const LAG_PLUG = 22
Public Const LAG_1 = 23
Public Const LAG_2 = 24
Public Const LAG_3 = 25
Public Const LAG_4 = 26
Public Const LAG_5 = 27
Public Const LAG_6 = 28

' State icons
Public Const IC_FRIEND_START = 29
Public Const IC_FRIEND_MUTUAL_START = 34

' World Cyber Games icons
Public Const IC_WCG_PLAYER = 102
Public Const IC_WCG_REF = 103

' UPDATED WCG Icons
Public Const IC_WCRF = 116
Public Const IC_WCPL = 117
Public Const IC_WCGO = 118
Public Const IC_WCSI = 119
Public Const IC_WCBR = 120
Public Const IC_WCPG = 121

' PGTour icons
Public Const IC_PGT_A = 104
Public Const IC_PGT_B = 107
Public Const IC_PGT_C = 110
Public Const IC_PGT_D = 113

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
Public Const FRL_PRIVATEGAME& = &H4
Public Const FRL_PRIVATEGAME_MUTUAL& = &H5

Public Const FRS_NONE& = &H0
Public Const FRS_MUTUAL& = &H1
Public Const FRS_DND& = &H2
Public Const FRS_AWAY& = &H4

'Clan list constants
Public Const IC_CLAN_PEON = 1
Public Const IC_CLAN_GRUNT = 2
Public Const IC_CLAN_SHAMAN = 3
Public Const IC_CLAN_CHIEFTAIN = 4
Public Const IC_CLAN_UNKNOWN = 5
Public Const IC_CLAN_STATUS_OFFLINE = 6
Public Const IC_CLAN_STATUS_ONLINE = 7

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

Public Const BTRUE& = 1
Public Const BFALSE& = 0

Public Const IF_AWAITING_CHPW& = 3
Public Const IF_SUBJECT_TO_IDLEBANS = 9
Public Const IF_CHPW_AND_IDLEBANS = 12

Public Const LOAD_SAFELIST = 1
Public Const LOAD_FILTERS = 2
Public Const LOAD_PHRASES = 3
Public Const LOAD_DB = 4
Public Const LOAD_BLOCKLIST = 5

Public Const DB_TYPE_USER = "USER"
Public Const DB_TYPE_GROUP = "GROUP"
Public Const DB_TYPE_CLAN = "CLAN"
Public Const DB_TYPE_GAME = "GAME"

Public Const CMD_PARAM_PREFIX = "--"
