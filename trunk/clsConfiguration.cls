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

Private Type PROFILE
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

Public Sub WriteLine(ByVal Section As String, ByVal Key As String, ByVal Value As String)

End Sub

Public Function ReadLine(ByVal Section As String, ByVal Key As String) As String

End Function

