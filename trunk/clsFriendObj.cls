VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFriendObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : g_FriendObj
' DateTime  : 3/15/2004 18:58
' Author    : Stealth
' Purpose   : stores friendlist data
'---------------------------------------------------------------------------------------

Option Explicit

Private m_Username    As String
Private m_DisplayName As String
Private m_Status      As Byte
Private m_LocationID  As Byte
Private m_Location    As String
Private m_Game        As String

Public Property Get Location() As String
    Location = m_Location
End Property

Public Property Let Location(ByVal sChannel As String)
    m_Location = sChannel
End Property

Public Property Get Game() As String
    Game = m_Game
End Property

Public Property Let Game(ByVal sProduct As String)
    m_Game = KillNull(sProduct)
End Property

Public Property Get LocationID() As Byte
    LocationID = m_LocationID
End Property

Public Property Let LocationID(ByVal byLocation As Byte)
    m_LocationID = byLocation
End Property

Public Property Get Status() As Byte
    Status = m_Status
End Property

Public Property Let Status(ByVal byStatus As Byte)
    m_Status = byStatus
End Property

Public Property Get IsMutual() As Boolean
    IsMutual = (Status And 1)
End Property

Public Property Get IsAway() As Boolean
    IsAway = (Status And 4)
End Property

Public Property Get DontDisturb() As Boolean
    DontDisturb = (Status And 2)
End Property

Public Property Get IsOnline() As Boolean
    IsOnline = (LocationID <> 0)
End Property

Public Property Get IsInChannel() As Boolean
    IsInChannel = (LocationID = 2)
End Property

Public Property Get IsInGame() As Boolean
    IsInGame = (LocationID > 2)
End Property

Public Property Get Name() As String
    Name = m_Username
End Property

Public Property Let Name(ByVal sUsername As String)
    m_Username = KillNull(sUsername)
    m_DisplayName = ConvertUsername(m_Username)
End Property

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Function Clone() As clsFriendObj
    Set Clone = New clsFriendObj
    
    Clone.Name = Name
    Clone.Game = Game
    Clone.Status = Status
    Clone.LocationID = LocationID
    Clone.Location = Location
End Function
