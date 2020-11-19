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

Option Explicit


Private m_Flags             As Long             ' users BNET flags
Private m_Ping              As Long
Private m_ActualName        As String           ' unmodified name as provided by the server
Private m_CharacterName     As String           ' D2 character name
Private m_Clan              As String           ' W3 clan tag
Private m_JoinDate          As Date             ' date-time joined the channel
Private m_LastSpeakDate     As Date             ' time last seen talking
Private m_Statstring        As String
Private m_Game              As String           ' product code (WAR3, D2DV, STAR, etc)
Private m_Queue             As Collection       ' queue'd chat events
Private m_PassedChannelAuth As Boolean
Private m_StatsObj          As clsUserStats
Private m_PendingBan        As Boolean
Private m_Phantom           As Boolean
Private m_Weight            As Long

Private m_CleanName         As String
Private m_DisplayName       As String           ' name displayed according to conventions
Private m_FullName          As String           ' full name including gateway

Private Sub Class_Initialize()
    Set m_Queue = New Collection
    Set m_StatsObj = New clsUserStats
    
    LastTalkTime = UtcNow
    JoinTime = UtcNow
End Sub

Private Sub Class_Terminate()
    ClearQueue

    Set m_Queue = Nothing
    Set m_StatsObj = Nothing
End Sub


Public Property Get Name() As String
    Name = m_CleanName
End Property

Public Property Let Name(ByVal sValue As String)
    ' Only update if value has actually changed.
    If StrComp(m_ActualName, sValue, vbBinaryCompare) = 0 Then
        Exit Property
    End If
    
    m_ActualName = sValue
    
    ' Update derived names
    m_CleanName = modEvents.CleanUsername(m_ActualName)
    m_DisplayName = ConvertUsername(m_ActualName, , Me)
    m_FullName = ConvertUsername(m_ActualName, 3)       ' 3 = show all gateways
End Property

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Property Get FullName() As String
    FullName = m_FullName
End Property


Public Property Get Flags() As Long
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal iValue As Long)
    m_Flags = iValue
End Property


Public Property Get Ping() As Long
    Ping = m_Ping
End Property

Public Property Let Ping(ByVal iValue As Long)
    m_Ping = iValue
End Property


Public Property Get Statstring() As String
    Statstring = m_Statstring
End Property

Public Property Let Statstring(ByVal sValue As String)
    m_Statstring = sValue
    
    m_StatsObj.Statstring = m_Statstring
End Property


Public Property Get Stats() As clsUserStats
    Set Stats = m_StatsObj
End Property


Public Property Get Game() As String
    Game = m_StatsObj.Game
End Property

Public Property Get IsUsingDII() As Boolean
    IsUsingDII = ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP))
End Property

Public Property Get IsUsingWarIII() As Boolean
    IsUsingWarIII = ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP))
End Property


Public Property Get Clan() As String
    Clan = m_StatsObj.Clan
End Property



Public Property Get CharacterName() As String
    CharacterName = m_CharacterName
End Property

Public Property Let CharacterName(ByVal sValue As String)
    m_CharacterName = sValue
End Property



Public Property Get PassedChannelAuth() As Boolean
    PassedChannelAuth = m_PassedChannelAuth
End Property

Public Property Let PassedChannelAuth(ByVal bln As Boolean)
    m_PassedChannelAuth = bln
End Property


Public Property Get PendingBan() As Boolean
    PendingBan = m_PendingBan
End Property

Public Property Let PendingBan(ByVal bValue As Boolean)
    m_PendingBan = bValue
End Property


Public Property Get IsPhantom() As Boolean
    IsPhantom = m_Phantom
End Property

Public Property Let IsPhantom(ByVal bValue As Boolean)
    m_Phantom = bValue
End Property


Public Property Get IsBnetAdmin() As Boolean
    IsBnetAdmin = ((m_Flags And USER_SYSOP) = USER_SYSOP)
End Property

Public Property Get IsBlizzRep() As Boolean
    IsBlizzRep = ((m_Flags And USER_BLIZZREP) = USER_BLIZZREP)
End Property

Public Property Get IsSpeaker() As Boolean
    IsSpeaker = ((m_Flags And USER_SPEAKER) = USER_SPEAKER)
End Property

Public Property Get IsSquelched() As Boolean
    IsSquelched = ((m_Flags And USER_SQUELCHED) = USER_SQUELCHED)
End Property

Public Property Get IsOperator() As Boolean
    IsOperator = frmChat.IsPriorityUser(m_Flags)
End Property


Public Property Let LastTalkTime(ByVal dtValue As Date)
    m_LastSpeakDate = dtValue
End Property

Public Property Get LastTalkTime() As Date
    LastTalkTime = m_LastSpeakDate
End Property


Public Property Let JoinTime(ByVal dtValue As Date)
    m_JoinDate = dtValue
End Property

Public Property Get JoinTime() As Date
    JoinTime = m_JoinDate
End Property

Public Property Let UserlistWeight(ByVal lng As Long)
    m_Weight = lng
End Property

Public Property Get UserlistWeight() As Long
    UserlistWeight = m_Weight
End Property


Public Function TimeSinceTalk() As Double
    On Error GoTo ERROR_HANDLER

    TimeSinceTalk = DateDiff("s", LastTalkTime, UtcNow)
    
    Exit Function

ERROR_HANDLER:
    Exit Function
End Function

Public Function TimeInChannel() As Double
    On Error GoTo ERROR_HANDLER
    
    TimeInChannel = DateDiff("s", JoinTime, UtcNow)
    
    Exit Function

ERROR_HANDLER:
    Exit Function
End Function


Public Property Get Queue() As Collection
    Set Queue = m_Queue
End Property

Public Sub DisplayQueue()
    On Error GoTo ERROR_HANDLER

    Dim CurrentEvent As clsUserEventObj
    Dim j            As Integer
    
    If (Queue Is Nothing) Then
        Exit Sub
    End If
    
    For j = 1 To Queue.Count
        If (j > Queue.Count) Then
            Exit For
        End If
    
        Set CurrentEvent = Queue(j)
    
        Select Case (CurrentEvent.EventID)
            Case ID_USER
                Call Event_UserInChannel(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
                    
            Case ID_JOIN
                Call Event_UserJoins(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
            
            Case ID_TALK
                Call Event_UserTalk(Name, CurrentEvent.Flags, CurrentEvent.Message, CurrentEvent.Ping, j)
            
            Case ID_EMOTE
                Call Event_UserEmote(Name, CurrentEvent.Flags, CurrentEvent.Message, j)
            
            Case ID_USERFLAGS
                Call Event_FlagsUpdate(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
        End Select
    Next j
    
    ClearQueue

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in clsUserObj::DisplayQueue()."

    Exit Sub
End Sub

Public Sub ClearQueue()
    Dim i As Integer
    
    For i = Queue.Count To 1 Step -1
        Queue.Remove i
    Next i
End Sub


Public Function Clone() As clsUserObj
    Dim i As Integer

    Set Clone = New clsUserObj

    Clone.Name = Name
    Clone.Ping = Ping
    Clone.Flags = Flags
    Clone.CharacterName = CharacterName
    Clone.JoinTime = JoinTime
    Clone.LastTalkTime = LastTalkTime
    Clone.PassedChannelAuth = PassedChannelAuth
    Clone.PendingBan = PendingBan
    Clone.IsPhantom = IsPhantom
    Clone.UserlistWeight = UserlistWeight
    Clone.Statstring = Statstring
    
    For i = 1 To Queue.Count
        Clone.Queue.Add Queue(i).Clone()
    Next i
End Function
