VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

' battle.net user flag constants
Private Const USER_BLIZZREP    As Long = &H1
Private Const USER_CHANNELOP   As Long = &H2
Private Const USER_SPEAKER     As Long = &H4
Private Const USER_SYSOP       As Long = &H8
Private Const USER_NOUDP       As Long = &H10
Private Const USER_BEEPENABLED As Long = &H100
Private Const USER_KBKOFFICIAL As Long = &H1000
Private Const USER_JAILED      As Long = &H100000
Private Const USER_SQUELCHED   As Long = &H20
Private Const USER_PGLPLAYER   As Long = &H200
Private Const USER_GFPLAYER    As Long = &H200000
Private Const USER_GUEST       As Long = &H40
Private Const USER_PGLOFFICIAL As Long = &H400
Private Const USER_KBKPLAYER   As Long = &H800

Private m_flags            As Long
Private m_ping             As Long
Private m_actual_name      As String
Private m_display_name     As String
Private m_clan             As String
Private m_clan_rank        As Integer
Private m_join_date        As Date
Private m_last_speak_date  As Date
Private m_stats_string     As String
Private m_game             As String
Private m_total_bans       As Long
Private m_total_kicks      As Long
Private m_operator_date    As Date

Public Property Get Name() As String

    ' ...
    Name = m_actual_name

End Property

Public Property Let Name(str As String)

    ' ...
    m_actual_name = str

End Property

Public Property Get Game() As String

    ' ...
    Game = m_game

End Property

Public Property Let Game(str As String)

    ' ...
    m_game = str

End Property

Public Property Get Clan() As String

    ' ...
    Clan = m_clan

End Property

Public Property Let Clan(str As String)

    ' ...
    m_clan = str

End Property

Public Property Get DisplayName() As String

    ' ...
    DisplayName = m_display_name

End Property

Public Property Let DisplayName(str As String)

    ' ...
    m_display_name = str

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Get IsOperator() As Boolean

    ' ...
    IsOperator = ((m_flags And USER_CHANNELOP&) = USER_CHANNELOP&)

End Property

Public Property Get IsSquelched() As Boolean

    ' ...
    IsSquelched = ((m_flags And USER_SQUELCHED&) = USER_SQUELCHED&)

End Property

Public Property Get Ping() As Long

    ' ...
    Ping = m_ping

End Property

Public Property Let Ping(lng As Long)

    ' ...
    m_ping = lng

End Property


Public Property Let JoinDate(jdate As Date)

    ' ...
    m_join_date = jdate

End Property

Public Property Get JoinDate() As Date

    ' ...
    JoinDate = m_join_date

End Property
