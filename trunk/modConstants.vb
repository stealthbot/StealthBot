Option Strict Off
Option Explicit On
Module modConstants
	
	Public Const REVISION As String = ""
	Public Const VERCODE As String = "2697"
	Public Const BNCSUTIL_VERSION As String = "1.3.1"
	Public Const CONFIG_VERSION As String = "6"
	
	' URLs
	Public Const BETA_AUTH_URL As String = "http://www.stealthbot.net/board/sbauth.php?username="
	Public Const BETA_AUTH_URL_CRC32 As Integer = 716038006
	
	Public Const BNLS_DEFAULT_SOURCE As String = "http://stealthbot.net/sb/bnls.php"
	
	Public Const VERBYTE_SOURCE As String = "http://www.stealthbot.net/sb/verbytes/versionbytes.txt"
	
	' Files
	Public Const FILE_COMMANDS As String = "Commands.xml"
	Public Const FILE_USERDB As String = "Users.txt"
	Public Const FILE_QUOTES As String = "Quotes.txt"
	Public Const FILE_CAUGHT_PHRASES As String = "CaughtPhrases.htm"
	Public Const FILE_FILTERS As String = "Filters.ini"
	Public Const FILE_PHRASE_BANS As String = "PhraseBans.txt"
	Public Const FILE_MAILDB As String = "Mail.dat"
	Public Const FILE_SCRIPT_INI As String = "Scripts.ini"
	Public Const FILE_CATCH_PHRASES As String = "CatchPhrases.txt"
	Public Const FILE_QUICK_CHANNELS As String = "QuickChannels.txt"
	Public Const FILE_COLORS As String = "Colors.sclf"
	Public Const FILE_CREV_INI As String = "CheckRevision.ini"
	Public Const FILE_WARDEN_INI As String = "Warden.ini"
	Public Const FILE_BNLS_LIST As String = "AdditionalBNLSservers.txt"
	Public Const FILE_SERVER_LIST As String = "Servers.txt"
	
	' Product codes
	Public Const PRODUCT_STAR As String = "STAR"
	Public Const PRODUCT_SEXP As String = "SEXP"
	Public Const PRODUCT_W2BN As String = "W2BN"
	Public Const PRODUCT_D2DV As String = "D2DV"
	Public Const PRODUCT_D2XP As String = "D2XP"
	Public Const PRODUCT_WAR3 As String = "WAR3"
	Public Const PRODUCT_W3XP As String = "W3XP"
	Public Const PRODUCT_SSHR As String = "SSHR"
	Public Const PRODUCT_JSTR As String = "JSTR"
	Public Const PRODUCT_DSHR As String = "DSHR"
	Public Const PRODUCT_DRTL As String = "DRTL"
	Public Const PRODUCT_CHAT As String = "CHAT"
	
	
	Public Const BNET_MSG_LENGTH As Short = 233
	
	' Colors
	Public Const COLOR_TEAL As Integer = &H99CC00
	Public Const COLOR_BLUE As Integer = &HCC9900
	
	' prod icons
	Public Const ICUNKNOWN As Short = 1
	Public Const ICSTAR As Short = 2
	Public Const ICSEXP As Short = 3
	Public Const ICD2DV As Short = 4
	Public Const ICD2XP As Short = 5
	Public Const ICW2BN As Short = 6
	Public Const ICWAR3 As Short = 7
	Public Const ICCHAT As Short = 8
	Public Const ICDIABLO As Short = 9
	Public Const ICDIABLOSW As Short = 10
	Public Const ICJSTR As Short = 11
	Public Const ICWAR3X As Short = 12
	Public Const ICSCSW As Short = 13
	
	' stats spawn icons
	Public Const IC_STAR_SPAWN As Short = 132
	Public Const IC_JSTR_SPAWN As Short = 133
	Public Const IC_W2BN_SPAWN As Short = 145
	Public Const IC_DIAB_SPAWN As Short = 158
	
	' stats icon sequences
	Public Const ICON_START_WAR3 As Short = 68
	Public Const ICON_START_W3XP As Short = 31
	Public Const ICON_START_D2 As Short = 114
	Public Const ICON_START_SC As Short = 121
	Public Const ICON_START_W2 As Short = 134
	Public Const ICON_START_D1 As Short = 146
	
	' flags icons
	Public Const ICGAVEL As Short = 14
	Public Const ICBLIZZ As Short = 15
	Public Const ICSYSOP As Short = 16
	Public Const ICSQUELCH As Short = 17
	Public Const ICSPEAKER As Short = 18
	Public Const ICSPECS As Short = 19
	
	' flags icons: GF
	Public Const IC_GF_OFFICIAL As Short = 20
	Public Const IC_GF_PLAYER As Short = 21
	
	' Lag icons
	Public Const LAG_PLUG As Short = 22
	Public Const LAG_1 As Short = 23
	Public Const LAG_2 As Short = 24
	Public Const LAG_3 As Short = 25
	Public Const LAG_4 As Short = 26
	Public Const LAG_5 As Short = 27
	Public Const LAG_6 As Short = 28
	
	' State icons
	Public Const MONITOR_ONLINE As Short = 29
	Public Const MONITOR_OFFLINE As Short = 30
	
	' World Cyber Games icons
	Public Const IC_WCG_PLAYER As Short = 94
	Public Const IC_WCG_REF As Short = 95
	
	' UPDATED WCG Icons
	Public Const IC_WCRF As Short = 108
	Public Const IC_WCPL As Short = 109
	Public Const IC_WCGO As Short = 110
	Public Const IC_WCSI As Short = 111
	Public Const IC_WCBR As Short = 112
	Public Const IC_WCPG As Short = 113
	
	' PGTour icons
	Public Const IC_PGT_A As Short = 96
	Public Const IC_PGT_B As Short = 99
	Public Const IC_PGT_C As Short = 102
	Public Const IC_PGT_D As Short = 105
	
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
	Public Const FRL_OFFLINE As Integer = &H0
	Public Const FRL_NOTINCHAT As Integer = &H1
	Public Const FRL_INCHAT As Integer = &H2
	Public Const FRL_PUBLICGAME As Integer = &H3
	Public Const FRL_PRIVATEGAME As Integer = &H5
	
	Public Const FRS_NONE As Integer = &H0
	Public Const FRS_MUTUAL As Integer = &H1
	Public Const FRS_DND As Integer = &H2
	Public Const FRS_AWAY As Integer = &H4
	
	'FlashWindow constants
	Public Const FLASHW_CAPTION As Short = 1
	Public Const FLASHW_TRAY As Short = 2
	Public Const FLASHW_ALL As Boolean = FLASHW_CAPTION Or FLASHW_TRAY
	Public Const FLASHW_TIMERNOFG As Boolean = 12 Or FLASHW_ALL
	Public Const FLW_SIZE As Short = 20
	
	'Key constants
	Public Const KEY_PGDN As Short = 34
	Public Const KEY_PGUP As Short = 33
	Public Const KEY_HOME As Short = 36
	Public Const KEY_ALTN As Short = 78
	Public Const KEY_END As Short = 35
	Public Const KEY_INSERT As Short = 45
	Public Const KEY_ENTER As Short = 13
	Public Const KEY_SPACE As Short = 32
	Public Const KEY_V As Short = 86
	Public Const KEY_A As Short = 65
	Public Const KEY_S As Short = 83
	Public Const KEY_D As Short = 68
	Public Const KEY_B As Short = 66
	Public Const KEY_J As Short = 74
	Public Const KEY_I As Short = 73
	Public Const KEY_U As Short = 85
	Public Const KEY_DELETE As Short = 46
	
	'Nagle algorithm constants
	Public Const IPPROTO_TCP As Integer = 6 'socketlevel
	Public Const TCP_NODELAY As Integer = &H1 'optname
	Public Const NAGLE_OPTLEN As Integer = 4 'optlen
	Public Const NAGLE_ON As Integer = 1 'optval
	Public Const NAGLE_OFF As Integer = 0 'optval
	
	Public Const SO_KEEPALIVE As Integer = &H8
	
	Public Const CAP_SHOW As String = "v"
	Public Const CAP_HIDE As String = "^"
	Public Const SHOW_SIZE As Integer = 255
	Public Const HIDE_SIZE As Integer = 2300
	Public Const TIP_HIDE As String = "Hide the whisper window"
	Public Const TIP_SHOW As String = "Show the whisper window"
	
	Public Const LVW_BUTTON_CHANNEL As Short = 0
	Public Const LVW_BUTTON_FRIENDS As Short = 1
	Public Const LVW_BUTTON_CLAN As Short = 2
	
	
	' Misc
	
	Public Const LVM_FIRST As Integer = &H1000
	Public Const LVM_HITTEST As Decimal = LVM_FIRST + 18
	
	Public Const CPTALK As Byte = 0
	Public Const CPEMOTE As Byte = 1
	Public Const CPWHISPER As Byte = 2
	
	Public Const GWL_WNDPROC As Short = (-4)
	
	Public Const WM_VSCROLL As Integer = &H115
	Public Const SB_VERT As Integer = 1
	Public Const EM_GETTHUMB As Integer = &HBE
	Public Const SB_HORZ As Integer = 0
	Public Const SB_BOTH As Integer = 3
	Public Const SB_THUMBPOSITION As Integer = &H4
	
	Public Const BTRUE As Integer = 1
	Public Const BFALSE As Integer = 0
	
	Public Const IF_AWAITING_CHPW As Integer = 3
	Public Const IF_SUBJECT_TO_IDLEBANS As Short = 9
	Public Const IF_CHPW_AND_IDLEBANS As Short = 12
	
	Public Const LOAD_SAFELIST As Short = 1
	Public Const LOAD_FILTERS As Short = 2
	Public Const LOAD_PHRASES As Short = 3
	Public Const LOAD_DB As Short = 4
End Module