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

Private m_username    As String
Private m_status      As Byte
Private m_location_id As Byte
Private m_location    As String
Private m_game        As String

Public Property Get Location() As String

    Location = m_location

End Property

Public Property Let Location(ByVal sChannel As String)

    m_location = sChannel
    
End Property

Public Property Get game() As String

    game = m_game

End Property

Public Property Let game(ByVal sProduct As String)

    m_game = KillNull(sProduct)

End Property

Public Property Get LocationID() As Byte

    LocationID = m_location_id

End Property

Public Property Let LocationID(ByVal byLocation As Byte)

    m_location_id = byLocation

End Property

Public Property Get Status() As Byte

    Status = m_status

End Property

Public Property Let Status(ByVal byStatus As Byte)

    m_status = byStatus

End Property

Public Property Get IsMutual() As Boolean

    IsMutual = (Status >= 1)

End Property

Public Property Get IsAway() As Boolean

    IsAway = (Status = 4)

End Property

Public Property Get DontDisturb() As Boolean

    DontDisturb = (Status = 2)

End Property

Public Property Get IsOnline() As Boolean

    IsOnline = (LocationID > 0)

End Property

Public Property Get IsInChannel() As Boolean

    IsInChannel = (LocationID = 2)

End Property

Public Property Get IsInGame() As Boolean

    IsInGame = (LocationID > 2)

End Property

Public Property Get Name() As String

    Name = m_username

End Property

Public Property Let Name(ByVal sUsername As String)

    m_username = KillNull(sUsername)

End Property

Public Property Get DisplayName() As String

    DisplayName = convertUsername(m_username)

End Property

Public Function Clone() As Object

    ' ...
    Set Clone = New clsFriendObj
    
    ' ...
    Clone.Name = Name
    Clone.game = game
    Clone.Status = Status
    Clone.LocationID = LocationID
    Clone.Location = Location
    
End Function
