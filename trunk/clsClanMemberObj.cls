VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanMemberObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsClanMemberObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name      As String
Private m_rank      As Integer
Private m_join_date As Date
Private m_status    As Integer
Private m_location  As String

' ...
Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

' ...
Public Property Let Name(str As String)

    ' ...
    m_name = str

End Property

' ...
Public Property Get DisplayName() As String

    ' ...
    DisplayName = convertUsername(Name)

End Property

' ...
Public Property Get Rank() As Integer

    ' ...
    Rank = m_rank

End Property

' ...
Public Property Let Rank(i As Integer)

    ' ...
    m_rank = i

End Property

' ...
Public Property Get JoinTime() As Date

    ' ...
    JoinTime = m_join_date

End Property

' ...
Public Property Let JoinTime(dJoined As Date)

    ' ...
    m_join_date = dJoined

End Property

' ...
Public Property Get Status() As Integer

    ' ...
    Status = m_status

End Property

' ...
Public Property Let Status(i As Integer)

    ' ...
    m_status = i

End Property

' ...
Public Property Get IsOnline() As Boolean

    ' ...
    IsOnline = (Status > 0)

End Property

' ...
Public Property Get Location() As String

    ' ...
    Location = m_location

End Property

' ...
Public Property Let Location(ByVal str As String)

    ' ...
    m_location = str

End Property

Public Sub MakeChieftain()

    ' ...
    Call MakeMemberChieftain(m_name)

End Sub

Public Sub Promote(Optional ByVal Rank As Integer = 0)

    ' ...
    Call PromoteMember(m_name, IIf(Rank, Rank, m_rank + 1))

End Sub

Public Sub Demote(Optional ByVal Rank As Integer = 0)

    ' ...
    Call DemoteMember(m_name, IIf(Rank, Rank, m_rank - 1))

End Sub

Public Sub KickOut()

    ' ...
    Call RemoveMember(m_name)

End Sub

