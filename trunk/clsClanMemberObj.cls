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

Option Explicit

Private m_Name     As String
Private m_Rank     As enuClanRank
Private m_JoinDate As Date
Private m_Status   As Byte
Private m_Location As String

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal str As String)
    m_Name = str
End Property

Public Property Get DisplayName() As String
    DisplayName = ConvertUsername(Name)
End Property

Public Property Get Rank() As enuClanRank
    Rank = m_Rank
End Property

Public Property Let Rank(ByVal i As enuClanRank)
    m_Rank = i
End Property

Public Property Get RankName() As String
    RankName = frmChat.ClanHandler.GetRank(CByte(m_Rank))
End Property

Public Property Get JoinTime() As Date
    JoinTime = m_JoinDate
End Property

Public Property Let JoinTime(ByVal dJoined As Date)
    m_JoinDate = dJoined
End Property

Public Property Get Status() As Byte
    Status = m_Status
End Property

Public Property Let Status(ByVal i As Byte)
    m_Status = i
End Property

Public Property Get IsOnline() As Boolean
    IsOnline = (Status > 0)
End Property

Public Property Get Location() As String
    Location = m_Location
End Property

Public Property Let Location(ByVal str As String)
    m_Location = str
End Property

Public Property Get IsSelf()
    IsSelf = (StrComp(m_Name, BotVars.Username, vbTextCompare) = 0)
End Property

Public Sub MakeChieftain()
    Call frmChat.ClanHandler.MakeMemberChieftain(m_Name, reqScriptingCall)
End Sub

Public Sub Promote(Optional ByVal Rank As Integer = -1)
    If ((Rank > -1) And (Rank <= m_Rank)) Then
        Exit Sub
    End If
    Call frmChat.ClanHandler.PromoteMember(m_Name, IIf(Rank > -1, Rank, m_Rank + 1), reqScriptingCall)
End Sub

Public Sub Demote(Optional ByVal Rank As Integer = -1)
    If ((Rank > -1) And (Rank >= m_Rank)) Then
        Exit Sub
    End If
    Call frmChat.ClanHandler.DemoteMember(m_Name, IIf(Rank > -1, Rank, m_Rank - 1), reqScriptingCall)
End Sub

Public Sub KickOut()
    Call frmChat.ClanHandler.RemoveMember(m_Name, IsSelf, reqScriptingCall)
End Sub

Public Function Clone() As Object
    Set Clone = New clsClanMemberObj
    
    Clone.Name = Name
    Clone.Location = Location
    Clone.Rank = Rank
    Clone.Status = Status
    Clone.JoinTime = JoinTime
End Function
