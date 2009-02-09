VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsConfiguration.cls
' ...

Option Explicit

Private Type PROFILE_GLOBAL
    ' user interface
    Left              As Integer
    Top               As Integer
    Height            As Integer
    Maximized         As Boolean
    LastSettingsPanel As Integer
    MaxBacklogSize    As Integer
    URLDetect         As Boolean
    FlashWindowOnMsg  As Boolean
    SystemTrayIcon    As Boolean
    ChatFont          As String
    ChatSize          As Integer
    ChanFont          As String
    ChanSize          As Integer
    DisablePrefix     As Boolean
    DisableSuffix     As Boolean
    
    ' miscellaneous
    ReconnectDelay       As Integer
    ShowSplash           As Boolean
    MaxLogFileSize       As Integer
    LogCommands          As Boolean
    LogDBActions         As Boolean
    UTF8                 As Boolean
    NoColoring           As Boolean
    NoAutocomplete       As Boolean
    AllowMP3             As Boolean
    UseGameConventions   As Boolean
    UseD2GameConventions As Boolean
    UseW3GameConventions As Boolean
    W2VerByte            As Integer
    SCVerByte            As Integer
    D2VerByte            As Integer
    W3VerByte            As Integer
    MediaPlayer          As String
End Type

Private Type Profile
    ' connection
    Username       As String
    Password       As String
    Server         As String
    ServerPort     As Integer
    UseProxy       As Integer
    ProxySrv       As String
    ProxyPort      As Integer
    game           As String
    CDKey()        As String
    CDKeyCount     As Integer
    CDKeyOwner     As String
    HomeChannel    As String
    SPOOF_UDP      As Integer
    SPOOF_PING     As Integer
    RealmConnect   As Integer
    RealmCharacter As String
    LocalIP        As String
    UseRLoginSrv   As Integer
    
    ' anti-idle
    UseIdle        As Integer
    IdleType       As Integer
    IdleInterval   As Integer
    IdleMessage    As String
    
    ' welcome greet
    UseGreets      As Integer
    WhisperGreet   As Integer
    GreetMessage   As String
    
    ' channel moderation
    Lockdown       As Integer
    PeonBan        As Integer
    IdleBan        As Integer
    PlugBan        As Integer
    KickOnYell     As Integer
    W3BanLevel     As Integer
    D2BanLevel     As Integer
    IPBan          As Integer
    BanEvasion     As Integer
    UsePhraseBans  As Integer
    PhraseBans()   As String
End Type

Private m_profile As String ' ...

Public Property Get Profile() As String

    Profile = m_profile

End Property

Public Property Let Profile(ByVal str As String)

    m_profile = str

End Property

Public Sub WriteLine(ByVal Key As String, ByVal Value As String)

    If (StrComp(Profile, "global", vbTextCompare) = 0) Then
    
    Else
    
    End If

End Sub

Public Function ReadLine(ByVal Key As String) As String

    If (StrComp(Profile, "global", vbTextCompare) = 0) Then
    
    Else
    
    End If

End Function

