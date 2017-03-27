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

Option Explicit

Private m_EventID     As Long
Private m_GTC         As Currency
Private m_Ping        As Long
Private m_Flags       As Long
Private m_Message     As String
Private m_Clan        As String
Private m_GameID      As String
Private m_IconCode    As String
Private m_Statstring  As String
Private m_Displayed   As Boolean

Private Sub Class_Initialize()
    EventTick = GetTickCountMS()
End Sub

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let EventID(ByVal lng As Long)
    m_EventID = lng
End Property

Public Property Let EventTick(ByVal lng As Currency)
    m_GTC = lng
End Property


Public Property Get EventTick() As Currency
    EventTick = m_GTC
End Property

Public Property Get Ping() As Long
    Ping = m_Ping
End Property

Public Property Let Ping(ByVal lng As Long)
    m_Ping = lng
End Property

Public Property Get Flags() As Long
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal lng As Long)
    m_Flags = lng
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

Public Property Let Message(ByVal str As String)
    m_Message = str
End Property

Public Property Get GameID() As String
    GameID = m_GameID
End Property

Public Property Let GameID(ByVal str As String)
    m_GameID = str
End Property

Public Property Get Clan() As String
    Clan = m_Clan
End Property

Public Property Let Clan(ByVal str As String)
    m_Clan = str
End Property

Public Property Get Statstring() As String
    Statstring = m_Statstring
End Property

Public Property Let Statstring(ByVal str As String)
    m_Statstring = str
End Property

Public Property Get IconCode() As String
    IconCode = m_IconCode
End Property

Public Property Let IconCode(ByVal str As String)
    m_IconCode = str
End Property

Public Property Get Displayed() As Boolean
    ' returns whether this event has been displayed in the RTB, useful for combining events
    ' in messages if chatdelay > 0, such as ops acquired and stats updates
    Displayed = m_Displayed
End Property

Public Property Let Displayed(ByVal bln As Boolean)
    ' sets whether this event has been displayed in the RTB
    m_Displayed = bln
End Property

Public Function Clone() As Object
    Set Clone = New clsUserEventObj
    
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
