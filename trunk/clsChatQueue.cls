VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChatQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BATTLE.NET EVENT CONSTANTS
Private Const ID_USER                As Long = &H1
Private Const ID_JOIN                As Long = &H2
Private Const ID_LEAVE               As Long = &H3
Private Const ID_WHISPER             As Long = &H4
Private Const ID_TALK                As Long = &H5
Private Const ID_BROADCAST           As Long = &H6
Private Const ID_CHANNEL             As Long = &H7
Private Const ID_USERFLAGS           As Long = &H9
Private Const ID_WHISPERSENT         As Long = &HA
Private Const ID_CHANNELFULL         As Long = &HD
Private Const ID_CHANNELDOESNOTEXIST As Long = &HE
Private Const ID_CHANNELRESTRICTED   As Long = &HF
Private Const ID_INFO                As Long = &H12
Private Const ID_ERROR               As Long = &H13
Private Const ID_EMOTE               As Long = &H17

Private Type udtChatPacket
    EventID    As Long
    Flags      As Long
    Ping       As Long
    Product    As String
    Clan       As String
    Statstring As String
    Message    As String
    w3icon     As String
End Type

Private m_arrChatPackets() As udtChatPacket

Private m_Username         As String
Private m_Count            As Long
Private m_GTC              As Long

Private Sub Class_Initialize()
    ReDim m_arrChatPackets(0)
End Sub

Public Sub StoreJoin(ByVal bnFlags As Long, ByVal Latency As Long, ByVal Game As String, _
    ByVal Clan As String, ByVal Statstring As String, ByVal w3icon As String)

    With m_arrChatPackets(m_Count)
        .EventID = ID_JOIN
        .Flags = bnFlags
        .Ping = Latency
        .Product = Game
        .Clan = Clan
        .Statstring = Statstring
        .w3icon = w3icon
    End With

    m_Count = (m_Count + 1)
    
    ReDim Preserve m_arrChatPackets(m_Count)
End Sub

Public Sub StoreStatusUpdate(ByVal bnFlags As Long, ByVal Latency As Long, ByVal Game As String, _
    ByVal Clan As String, ByVal Statstring As String, ByVal w3icon As String)

    With m_arrChatPackets(m_Count)
        .EventID = ID_USERFLAGS
        .Flags = bnFlags
        .Ping = Latency
        .Product = Game
        .Clan = Clan
        .Statstring = Statstring
        .w3icon = w3icon
    End With

    m_Count = (m_Count + 1)
    
    ReDim Preserve m_arrChatPackets(m_Count)
End Sub

Public Sub StoreTalk(ByVal bnFlags As Long, ByVal Latency As Long, ByVal Message As String)

    With m_arrChatPackets(m_Count)
        .EventID = ID_TALK
        .Flags = bnFlags
        .Ping = Latency
        .Message = Message
    End With

    m_Count = (m_Count + 1)
    
    ReDim Preserve m_arrChatPackets(m_Count)
End Sub

Public Sub StoreEmote(ByVal bnFlags As Long, ByVal Latency As Long, ByVal Message As String)

    With m_arrChatPackets(m_Count)
        .EventID = ID_EMOTE
        .Flags = bnFlags
        .Ping = Latency
        .Message = Message
    End With

    m_Count = (m_Count + 1)
    
    ReDim Preserve m_arrChatPackets(m_Count)
End Sub

Public Sub Show()
    Dim i As Integer ' ...
    
    For i = 0 To (m_Count - 1)
        ' ...
        Dim ChatPacket As udtChatPacket
        
        ChatPacket = m_arrChatPackets(i)
        
        With ChatPacket
            Select Case (.EventID)
                Case ID_JOIN:
                    Call modChatQueue.Event_QueuedJoin(m_Username, _
                            .Flags, .Ping, .Product, .Clan, .Statstring, _
                                .w3icon)
                                
                Case ID_USERFLAGS:
                    Call modChatQueue.Event_QueuedStatusUpdate
                    
                Case ID_TALK:
                    Call modChatQueue.Event_QueuedTalk(m_Username, _
                            .Flags, .Ping, .Message)
                            
                Case ID_EMOTE:
                    Call modChatQueue.Event_QueuedEmote(m_Username, _
                            .Flags, .Ping, .Message)
            End Select
        End With
    Next i
    
    Call Clear
End Sub

Private Sub Clear()
    ReDim m_arrChatPackets(0)
End Sub

' CLASS PROPERTIES

Public Property Get Time() As Long
    Time = m_GTC
End Property

Public Property Let Time(ByVal GTC As Long)
    m_GTC = GTC
End Property

Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let Username(ByVal Username As String)
    m_Username = Username
End Property
