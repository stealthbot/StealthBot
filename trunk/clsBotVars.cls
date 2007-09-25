VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBotVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsBotVars
' DateTime  : 1/22/2005 16:09
' Author    : Andy (andy@stealthbot.net)
' Purpose   : stores bot variables at runtime and is exposed to the scripting subsystem
'---------------------------------------------------------------------------------------
Option Explicit

Private m_sUsername                     As String
Private m_sPassword                     As String
Private m_sCDKey                        As String
Private m_sLODKey                       As String
Private m_sProduct                      As String
Private m_sServer                       As String
Private m_byBanUnderLevel               As Byte
Private m_byBanD2UnderLevel             As Byte
Private m_byKickOnYell                  As Byte
Private m_sHashFilePath                 As String
Private m_bBNLS                         As Boolean
Private m_bLoadMode                     As Boolean
Private m_byLogging                     As Byte
Private m_bySpoof                       As Byte
Private m_bDisableMP3Commands           As Boolean
Private m_bUseUDP                       As Boolean
Private m_byFlashOnEvents               As Byte
Private m_bUseProxy                     As Boolean
Private m_sProxyIP                      As String
Private m_lProxyPort                    As Long
Private m_eProxyStatus                  As enuProxyStatus
Private m_bProxyIsSocks5                As Boolean
Private m_lMaxLogfileSize               As Long
Private m_bUseGreet                     As Boolean
Private m_sGreetMsg                     As String
Private m_bWhisperGreet                 As Boolean
Private m_bUseRealm                     As Boolean
Private m_bNoTray                       As Boolean
Private m_bShowOfflineFriends           As Boolean
Private m_lReconnectDelay               As Long
Private m_bUsingDirectFList             As Boolean
Private m_lAutofilterMS                 As Long
Private m_bNoAutocompletion             As Boolean
Private m_bNoColoring                   As Boolean
Private m_lSecondsToIdle                As Long
Private m_sTrigger                      As String
Private m_bScriptMessageVeto            As Boolean
Private m_bIPBans                       As Boolean
Private m_bPlugBan                      As Boolean
Private m_bClientBans                   As Boolean
Private m_byTSSetting                   As Byte
Private m_sHomeChannel                  As String
Private m_bQuietTime                    As Boolean
Private m_bUseBackupChan                As Boolean
Private m_sBackupChan                   As String
Private m_byIB_On                       As Byte
Private m_iIB_Wait                      As Integer
Private m_bIB_Kick                      As Boolean
Private m_sChannelPassword              As String
Private m_byChannelPasswordDelay        As Byte
Private m_bBanEvasion                   As Boolean
Private m_bLogDBActions                 As Boolean
Private m_bLogCommands                  As Boolean
Private m_sBotOwner                     As String
Private m_byBanPeons                    As Byte
Private m_bWhisperCmds                  As Boolean
Private m_bDoScroll                     As Boolean
Private m_bLockChat                     As Boolean
Private m_iJoinWatch                    As Integer
Private m_sBNLSServer                   As String
Private m_bNoRTBAutomaticCopy           As Boolean


Public Property Get Username() As String

    Username = m_sUsername

End Property

Public Property Let Username(ByVal sUsername As String)

    m_sUsername = sUsername

End Property

Public Property Get Password() As String

    Password = m_sPassword

End Property

Public Property Let Password(ByVal sPassword As String)

    m_sPassword = sPassword

End Property

Public Property Get CDKey() As String

    CDKey = m_sCDKey

End Property

Public Property Let CDKey(ByVal sCDKey As String)

    m_sCDKey = sCDKey

End Property

Public Property Get LODKey() As String

    LODKey = m_sLODKey

End Property

Public Property Let LODKey(ByVal sLODKey As String)

    m_sLODKey = sLODKey

End Property

Public Property Get Product() As String

    Product = m_sProduct

End Property

Public Property Let Product(ByVal sProduct As String)

    m_sProduct = sProduct

End Property

Public Property Get Server() As String

    Server = m_sServer

End Property

Public Property Let Server(ByVal sServer As String)

    m_sServer = sServer

End Property

Public Property Get BanUnderLevel() As Byte

    BanUnderLevel = m_byBanUnderLevel

End Property

Public Property Let BanUnderLevel(ByVal byBanUnderLevel As Byte)

    m_byBanUnderLevel = byBanUnderLevel

End Property

Public Property Get BanD2UnderLevel() As Byte

    BanD2UnderLevel = m_byBanD2UnderLevel

End Property

Public Property Let BanD2UnderLevel(ByVal byBanD2UnderLevel As Byte)

    m_byBanD2UnderLevel = byBanD2UnderLevel

End Property

Public Property Get KickOnYell() As Byte

    KickOnYell = m_byKickOnYell

End Property

Public Property Let KickOnYell(ByVal byKickOnYell As Byte)

    m_byKickOnYell = byKickOnYell

End Property

Public Property Get HashFilePath() As String

    HashFilePath = m_sHashFilePath

End Property

Public Property Let HashFilePath(ByVal sHashFilePath As String)

    m_sHashFilePath = sHashFilePath

End Property

Public Property Get BNLS() As Boolean

    BNLS = m_bBNLS

End Property

Public Property Let BNLS(ByVal bBNLS As Boolean)

    m_bBNLS = bBNLS

End Property

Public Property Get LoadMode() As Boolean

    LoadMode = m_bLoadMode

End Property

Public Property Let LoadMode(ByVal bLoadMode As Boolean)

    m_bLoadMode = bLoadMode

End Property

Public Property Get Logging() As Byte

    Logging = m_byLogging

End Property

Public Property Let Logging(ByVal byLogging As Byte)

    m_byLogging = byLogging

End Property

Public Property Get Spoof() As Byte

    Spoof = m_bySpoof

End Property

Public Property Let Spoof(ByVal bySpoof As Byte)

    m_bySpoof = bySpoof

End Property

Public Property Get DisableMP3Commands() As Boolean

    DisableMP3Commands = m_bDisableMP3Commands

End Property

Public Property Let DisableMP3Commands(ByVal bDisableMP3Commands As Boolean)

    m_bDisableMP3Commands = bDisableMP3Commands

End Property

Public Property Get UseUDP() As Boolean

    UseUDP = m_bUseUDP

End Property

Public Property Let UseUDP(ByVal bUseUDP As Boolean)

    m_bUseUDP = bUseUDP

End Property

Public Property Get FlashOnEvents() As Byte

    FlashOnEvents = m_byFlashOnEvents

End Property

Public Property Let FlashOnEvents(ByVal byFlashOnEvents As Byte)

    m_byFlashOnEvents = byFlashOnEvents

End Property

Public Property Get UseProxy() As Boolean

    UseProxy = m_bUseProxy

End Property

Public Property Let UseProxy(ByVal bUseProxy As Boolean)

    m_bUseProxy = bUseProxy

End Property

Public Property Get ProxyPort() As Long

    ProxyPort = m_lProxyPort

End Property

Public Property Let ProxyPort(ByVal lProxyPort As Long)

    m_lProxyPort = lProxyPort

End Property

Public Property Get ProxyStatus() As enuProxyStatus

    ProxyStatus = m_eProxyStatus

End Property

Public Property Let ProxyStatus(ByVal eProxyStatus As enuProxyStatus)

    m_eProxyStatus = eProxyStatus

End Property

Public Property Get ProxyIP() As String

    ProxyIP = m_sProxyIP

End Property

Public Property Let ProxyIP(ByVal sProxyIP As String)

    m_sProxyIP = sProxyIP

End Property

Public Property Get ProxyIsSocks5() As Boolean

    ProxyIsSocks5 = m_bProxyIsSocks5

End Property

Public Property Let ProxyIsSocks5(ByVal bProxyIsSocks5 As Boolean)

    m_bProxyIsSocks5 = bProxyIsSocks5

End Property

Public Property Get MaxLogfileSize() As Long

    MaxLogfileSize = m_lMaxLogfileSize

End Property

Public Property Let MaxLogfileSize(ByVal lMaxLogfileSize As Long)

    m_lMaxLogfileSize = lMaxLogfileSize

End Property

Public Property Get UseGreet() As Boolean

    UseGreet = m_bUseGreet

End Property

Public Property Let UseGreet(ByVal bUseGreet As Boolean)

    m_bUseGreet = bUseGreet

End Property

Public Property Get GreetMsg() As String

    GreetMsg = m_sGreetMsg

End Property

Public Property Let GreetMsg(ByVal sGreetMsg As String)

    m_sGreetMsg = sGreetMsg

End Property

Public Property Get UseRealm() As Boolean

    UseRealm = m_bUseRealm

End Property

Public Property Let UseRealm(ByVal bUseRealm As Boolean)

    m_bUseRealm = bUseRealm

End Property

Public Property Get WhisperGreet() As Boolean

    WhisperGreet = m_bWhisperGreet

End Property

Public Property Let WhisperGreet(ByVal bWhisperGreet As Boolean)

    m_bWhisperGreet = bWhisperGreet

End Property

Public Property Get NoTray() As Boolean

    NoTray = m_bNoTray

End Property

Public Property Let NoTray(ByVal bNoTray As Boolean)

    m_bNoTray = bNoTray

End Property

Public Property Get ShowOfflineFriends() As Boolean

    ShowOfflineFriends = m_bShowOfflineFriends

End Property

Public Property Let ShowOfflineFriends(ByVal bShowOfflineFriends As Boolean)

    m_bShowOfflineFriends = bShowOfflineFriends

End Property

Public Property Get ReconnectDelay() As Long

    ReconnectDelay = m_lReconnectDelay

End Property

Public Property Let ReconnectDelay(ByVal lReconnectDelay As Long)

    m_lReconnectDelay = lReconnectDelay

End Property

Public Property Get UsingDirectFList() As Boolean

    UsingDirectFList = m_bUsingDirectFList

End Property

Public Property Let UsingDirectFList(ByVal bUsingDirectFList As Boolean)

    m_bUsingDirectFList = bUsingDirectFList

End Property

Public Property Get AutofilterMS() As Long

    AutofilterMS = m_lAutofilterMS

End Property

Public Property Let AutofilterMS(ByVal lAutofilterMS As Long)

    m_lAutofilterMS = lAutofilterMS

End Property

Public Property Get NoColoring() As Boolean

    NoColoring = m_bNoColoring

End Property

Public Property Let NoColoring(ByVal bNoColoring As Boolean)

    m_bNoColoring = bNoColoring

End Property

Public Property Get ScriptMessageVeto() As Boolean

    ScriptMessageVeto = m_bScriptMessageVeto

End Property

Public Property Let ScriptMessageVeto(ByVal bScriptMessageVeto As Boolean)

    m_bScriptMessageVeto = bScriptMessageVeto

End Property

Public Property Get NoAutocompletion() As Boolean

    NoAutocompletion = m_bNoAutocompletion

End Property

Public Property Let NoAutocompletion(ByVal bNoAutocompletion As Boolean)

    m_bNoAutocompletion = bNoAutocompletion

End Property

Public Property Get SecondsToIdle() As Long

    SecondsToIdle = m_lSecondsToIdle

End Property

Public Property Let SecondsToIdle(ByVal lSecondsToIdle As Long)

    m_lSecondsToIdle = lSecondsToIdle

End Property

Public Property Get Trigger() As String

    Trigger = m_sTrigger

End Property

Public Property Let Trigger(ByVal sTrigger As String)

    'Triggers longer or shorter than 1 character will break some stuff
    If Len(sTrigger) = 1 Then m_sTrigger = sTrigger

End Property

Public Property Get IPBans() As Boolean

    IPBans = m_bIPBans

End Property

Public Property Let IPBans(ByVal bIPBans As Boolean)

    m_bIPBans = bIPBans

End Property

Public Property Get PlugBan() As Boolean

    PlugBan = m_bPlugBan

End Property

Public Property Let PlugBan(ByVal bPlugBan As Boolean)

    m_bPlugBan = bPlugBan

End Property

Public Property Get ClientBans() As Boolean

    ClientBans = m_bClientBans

End Property

Public Property Let ClientBans(ByVal bClientBans As Boolean)

    m_bClientBans = bClientBans

End Property

Public Property Get TSSetting() As Byte

    TSSetting = m_byTSSetting

End Property

Public Property Let TSSetting(ByVal byTSSetting As Byte)

    m_byTSSetting = byTSSetting

End Property

Public Property Get HomeChannel() As String

    HomeChannel = m_sHomeChannel

End Property

Public Property Let HomeChannel(ByVal sHomeChannel As String)

    m_sHomeChannel = sHomeChannel

End Property

Public Property Get QuietTime() As Boolean

    QuietTime = m_bQuietTime

End Property

Public Property Let QuietTime(ByVal bQuietTime As Boolean)

    m_bQuietTime = bQuietTime

End Property

Public Property Get UseBackupChan() As Boolean

    UseBackupChan = m_bUseBackupChan

End Property

Public Property Let UseBackupChan(ByVal bUseBackupChan As Boolean)

    m_bUseBackupChan = bUseBackupChan

End Property

Public Property Get BackupChan() As String

    BackupChan = m_sBackupChan

End Property

Public Property Let BackupChan(ByVal sBackupChan As String)

    m_sBackupChan = sBackupChan

End Property

Public Property Get IB_On() As Byte

    IB_On = m_byIB_On

End Property

Public Property Let IB_On(ByVal byIB_On As Byte)

    m_byIB_On = byIB_On

End Property

Public Property Get IB_Wait() As Integer

    IB_Wait = m_iIB_Wait

End Property

Public Property Let IB_Wait(ByVal iIB_Wait As Integer)

    m_iIB_Wait = iIB_Wait

End Property

Public Property Get IB_Kick() As Boolean

    IB_Kick = m_bIB_Kick

End Property

Public Property Let IB_Kick(ByVal bIB_Kick As Boolean)

    m_bIB_Kick = bIB_Kick

End Property

Public Property Get ChannelPassword() As String

    ChannelPassword = m_sChannelPassword

End Property

Public Property Let ChannelPassword(ByVal sChannelPassword As String)

    m_sChannelPassword = sChannelPassword

End Property

Public Property Get ChannelPasswordDelay() As Byte

    ChannelPasswordDelay = m_byChannelPasswordDelay

End Property

Public Property Let ChannelPasswordDelay(ByVal byChannelPasswordDelay As Byte)

    m_byChannelPasswordDelay = byChannelPasswordDelay

End Property

Public Property Get BanEvasion() As Boolean

    BanEvasion = m_bBanEvasion

End Property

Public Property Let BanEvasion(ByVal bBanEvasion As Boolean)

    m_bBanEvasion = bBanEvasion

End Property

Public Property Get LogDBActions() As Boolean

    LogDBActions = m_bLogDBActions

End Property

Public Property Let LogDBActions(ByVal bLogDBActions As Boolean)

    m_bLogDBActions = bLogDBActions

End Property

Public Property Get LogCommands() As Boolean

    LogCommands = m_bLogCommands

End Property

Public Property Let LogCommands(ByVal bLogCommands As Boolean)

    m_bLogCommands = bLogCommands

End Property

Public Property Get BotOwner() As String

    BotOwner = m_sBotOwner

End Property

Public Property Let BotOwner(ByVal sBotOwner As String)

    m_sBotOwner = sBotOwner

End Property

Public Property Get BanPeons() As Byte

    BanPeons = m_byBanPeons

End Property

Public Property Let BanPeons(ByVal byBanPeons As Byte)

    m_byBanPeons = byBanPeons

End Property

Public Property Get WhisperCmds() As Boolean

    WhisperCmds = m_bWhisperCmds

End Property

Public Property Let WhisperCmds(ByVal bWhisperCmds As Boolean)

    m_bWhisperCmds = bWhisperCmds

End Property

Public Property Get DoScroll() As Boolean

    DoScroll = m_bDoScroll

End Property

Public Property Let DoScroll(ByVal bDoScroll As Boolean)

    m_bDoScroll = bDoScroll

End Property

Public Property Get LockChat() As Boolean

    LockChat = m_bLockChat

End Property

Public Property Let LockChat(ByVal bLockChat As Boolean)

    m_bLockChat = bLockChat

End Property

Public Property Get JoinWatch() As Integer

    JoinWatch = m_iJoinWatch

End Property

Public Property Let JoinWatch(ByVal iJoinWatch As Integer)

    m_iJoinWatch = iJoinWatch

End Property

Public Property Get BNLSServer() As String

    BNLSServer = m_sBNLSServer

End Property

Public Property Let BNLSServer(ByVal sBNLSServer As String)

    m_sBNLSServer = sBNLSServer

End Property

Public Property Get NoRTBAutomaticCopy() As Boolean

    NoRTBAutomaticCopy = m_bNoRTBAutomaticCopy

End Property

Public Property Let NoRTBAutomaticCopy(ByVal bNoRTBAutomaticCopy As Boolean)

    m_bNoRTBAutomaticCopy = bNoRTBAutomaticCopy

End Property
