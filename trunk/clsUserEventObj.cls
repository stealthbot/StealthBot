VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserEventObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsUserEventObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_event_id    As Long
Private m_gtc         As Long
Private m_ping        As Long
Private m_flags       As Long
Private m_message     As String
Private m_stats       As clsUserStats
Private m_clan        As String
Private m_game_id     As String
Private m_icon_code   As String
Private m_stat_string As String

Private Sub Class_Initialize()

    ' ...
    EventTick = GetTickCount()
    
    ' ...
    Set m_stats = New clsUserStats

End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_stats = Nothing

End Sub

Public Property Get EventID() As Long

    ' ...
    EventID = m_event_id

End Property

Public Property Let EventID(ByVal lng As Long)

    ' ...
    m_event_id = lng

End Property

Public Property Let EventTick(ByVal lng As Long)

    ' ...
    m_gtc = lng

End Property


Public Property Get EventTick() As Long

    ' ...
    EventTick = m_gtc

End Property

Public Property Get Ping() As Long

    ' ...
    Ping = m_ping

End Property

Public Property Let Ping(ByVal lng As Long)

    ' ...
    m_ping = lng

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(ByVal lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Get Message() As String

    ' ...
    Message = m_message

End Property

Public Property Let Message(ByVal str As String)

    ' ...
    m_message = str

End Property

Public Property Get GameID() As String

    ' ...
    GameID = m_game_id

End Property

Public Property Let GameID(ByVal str As String)

    ' ...
    m_game_id = str

End Property

Public Property Get Clan() As String

    ' ...
    Clan = m_clan

End Property

Public Property Let Clan(ByVal str As String)

    ' ...
    m_clan = str

End Property

Public Property Get Statstring() As String

    ' ...
    Statstring = m_stat_string

End Property

Public Property Let Statstring(ByVal str As String)

    ' ...
    m_stat_string = str
    
    ' ...
    m_stats.Statstring = str

End Property

Public Property Get IconCode() As String

    ' ...
    IconCode = m_icon_code

End Property

Public Property Let IconCode(ByVal str As String)

    ' ...
    m_icon_code = str

End Property

Public Function Clone() As Object

    ' ...
    Set Clone = New clsUserEventObj
    
    ' ...
    Clone.EventID = EventID
    Clone.EventTick = EventTick
    Clone.Flags = Flags
    Clone.GameID = GameID
    Clone.Clan = Clan
    Clone.IconCode = IconCode
    Clone.Message = Message
    Clone.Ping = Ping
    Clone.Statstring = Statstring

End Function

