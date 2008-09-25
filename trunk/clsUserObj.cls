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
Private m_character_name   As String
Private m_clan             As String
Private m_clan_rank        As Integer
Private m_join_date        As Date
Private m_last_speak_date  As Date
Private m_stats_string     As String
Private m_game             As String
Private m_total_bans       As Long
Private m_total_kicks      As Long
Private m_operator_date    As Date
Private m_queue            As Collection
Private m_passed_chan_auth As Boolean
Private m_stats            As clsUserStats

Private Sub Class_Initialize()
    
    ' ...
    Set m_queue = New Collection
    Set m_stats = New clsUserStats
    
    ' ...
    LastTalkTime = UtcNow
    JoinTime = UtcNow
        
End Sub

Public Property Get Name() As String

    ' ...
    Name = modEvents.CleanUsername(m_actual_name)

End Property

Public Property Let Name(str As String)

    ' ...
    m_actual_name = str

End Property

Public Property Get CharacterName() As String

    ' ...
    Name = m_character_name

End Property

Public Property Let CharacterName(str As String)

    ' ...
    m_character_name = str

End Property

Public Property Get game() As String

    ' ...
    game = m_game

End Property

Public Property Let game(str As String)

    ' ...
    m_game = str

End Property

Public Property Get IsUsingDII() As Boolean

    ' ...
    IsUsingDII = ((game = "D2DV") Or (game = "D2XP"))

End Property

Public Property Get IsUsingWarIII() As Boolean

    ' ...
    IsUsingWarIII = ((game = "WAR3") Or (game = "W3XP"))

End Property

Public Property Get Statstring() As String

    ' ...
    Statstring = m_stats_string

End Property

Public Property Let Statstring(str As String)

    ' ...
    m_stats_string = str

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
    DisplayName = convertUsername(m_actual_name)

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Get PassedChannelAuth() As Boolean

    ' ...
    PassedChannelAuth = m_passed_chan_auth

End Property

Public Property Let PassedChannelAuth(bln As Boolean)

    ' ...
    m_passed_chan_auth = bln

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

Public Property Let LastTalkTime(TDate As Date)

    ' ...
    m_last_speak_date = TDate

End Property

Public Property Get LastTalkTime() As Date

    ' ...
    LastTalkTime = m_last_speak_date

End Property


Public Property Let JoinTime(JDate As Date)

    ' ...
    m_join_date = JDate

End Property

Public Property Get JoinTime() As Date

    ' ...
    JoinTime = m_join_date

End Property

Public Function TimeSinceTalk() As Double

    ' ...
    On Error GoTo ERROR_HANDLER

    ' ...
    TimeSinceTalk = DateDiff("s", LastTalkTime, UtcNow)
    
    ' ...
    Exit Function

ERROR_HANDLER:

    ' ...
    Exit Function

End Function

Public Function TimeInChannel() As Double

    ' ...
    On Error GoTo ERROR_HANDLER
    
    ' ...
    TimeInChannel = DateDiff("s", JoinTime, UtcNow)
    
    ' ...
    Exit Function

ERROR_HANDLER:

    ' ...
    Exit Function
    
End Function

Public Sub ClearQueue()

    ' ...
    Set m_queue = New Collection

End Sub

Public Property Get Queue() As Collection

    ' ...
    Set Queue = m_queue
    
End Property

Public Property Get Stats() As clsUserStats

    ' ...
    Set Stats = m_stats
    
End Property

' ...
'Added error handling - FrOzeN
Public Sub DisplayQueue()
    On Error GoTo ERROR_HANDLER

    Dim CurrentEvent As clsUserEventObj
    Dim j            As Integer
    
    ' ...
    For j = 1 To Queue.Count
    
        ' ...
        Set CurrentEvent = Queue(j)
    
        ' ...
        Select Case (CurrentEvent.EventID)
            ' ...
            Case ID_USER
                Call Event_UserInChannel(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, _
                    CurrentEvent.GameID, CurrentEvent.Clan, CurrentEvent.Statstring, CurrentEvent.IconCode, j)
                    
            ' ...
            Case ID_JOIN
                Call Event_UserJoins(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, _
                    CurrentEvent.GameID, CurrentEvent.Clan, CurrentEvent.Statstring, CurrentEvent.IconCode, j)
            
            ' ...
            Case ID_TALK
                Call Event_UserTalk(Name, CurrentEvent.Flags, CurrentEvent.Message, CurrentEvent.Ping, j)
            
            ' ...
            Case ID_EMOTE
                Call Event_UserEmote(Name, CurrentEvent.Flags, CurrentEvent.Message, j)
            
            ' ...
            Case ID_USERFLAGS
                Call Event_FlagsUpdate(Name, CurrentEvent.Statstring, CurrentEvent.Flags, CurrentEvent.Ping, _
                    CurrentEvent.GameID, j)
        End Select
    Next j
    
    ' ...
    ClearQueue

    Exit Sub

ERROR_HANDLER:
    Exit Sub
End Sub

