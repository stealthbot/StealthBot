VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsMyClanObj.cls
' Copyright (C) 2008 Eric Evans

Option Explicit

Private m_PI       As Boolean
Private m_PICookie As Long
Private m_PMOTD    As Boolean
Private m_InClan   As Boolean
Private m_MyRank   As Byte
Private m_Tag      As String
Private m_Name     As String
Private m_MOTD     As String
Private m_Members  As Collection

Private Sub Class_Initialize()
    Set m_Members = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Members = Nothing
End Sub

Public Property Get Name() As String
    Name = m_Tag
End Property

Public Property Let Name(str As String)
    m_Tag = str
End Property

' alias for Name
Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Tag(str As String)
    m_Tag = str
End Property

Public Property Get FullName() As String
    FullName = m_Name
End Property

Public Property Let FullName(str As String)
    m_Name = str
End Property

Public Property Get MOTD() As String
    MOTD = m_MOTD
End Property

Public Property Let MOTD(str As String)
    m_MOTD = str
End Property

Public Property Get InClan() As Boolean
    InClan = m_InClan
End Property

Public Property Let InClan(bool As Boolean)
    m_InClan = bool
End Property

Public Property Get PendingInvitation() As Boolean
    PendingInvitation = m_PI
End Property

Public Property Let PendingInvitation(bool As Boolean)
    m_PI = bool
End Property

Public Property Get PendingInvitationCookie() As Long
    PendingInvitationCookie = m_PICookie
End Property

Public Property Let PendingInvitationCookie(lng As Long)
    m_PICookie = lng
End Property

Public Property Get PendingClanMOTD() As Boolean
    PendingClanMOTD = m_PMOTD
End Property

Public Property Let PendingClanMOTD(bool As Boolean)
    m_PMOTD = bool
End Property

Public Property Get Self() As clsClanMemberObj
    Dim i As Integer

    Set Self = New clsClanMemberObj

    For i = 1 To Members.Count
        'Changed 10/08/09 - Hdx - Now uses BotVars.Username instead of CurrentUsername so that if you login as a multiple (#2 #2) it will still work.
        If (StrComp(Members(i).Name, BotVars.Username, vbTextCompare) = 0) Then
            Set Self = Members(i)
            Exit For
        End If
    Next i
End Property

Public Property Get Members() As Collection
    Set Members = m_Members
End Property

Public Property Get Chieftain() As clsClanMemberObj
    Dim i As Integer
    Set Chieftain = New clsClanMemberObj
    
    For i = 1 To Members.Count
        If (Members(i).Rank >= 4) Then
            Set Chieftain = Members(i)
        
            Exit For
        End If
    Next i
End Property

Public Property Get Shamans() As Collection
    Dim i As Integer
    Set Shamans = New Collection
    
    For i = 1 To Members.Count
        If (Members(i).Rank = 3) Then
            Shamans.Add Members(i)
        End If
    Next i
End Property

Public Property Get Grunts() As Collection
    Dim i As Integer
    Set Grunts = New Collection
    
    For i = 1 To Members.Count
        If (Members(i).Rank = 2) Then
            Grunts.Add Members(i)
        End If
    Next i
End Property

Public Property Get Peons() As Collection
    Dim i As Integer
    Set Peons = New Collection
    
    For i = 1 To Members.Count
        If ((Members(i).Rank = 0) Or (Members(i).Rank = 1)) Then
            Peons.Add Members(i)
        End If
    Next i
End Property

' alias for GetUser()
Public Function GetMember(ByVal Username As String) As Object
    Set GetMember = GetUser(Username)
End Function

Public Function GetUser(ByVal Username As String) As Object
    Dim Index As Integer
    
    If (StrictIsNumeric(Username)) Then
        Index = CInt(Val(Username))
    Else
        Index = GetUserIndex(Username)
    End If
    
    If ((Index >= 1) And (Index <= Members.Count)) Then
        Set GetUser = Members(Index)
    Else
        Set GetUser = New clsClanMemberObj
    End If
End Function

' alias for GetUserIndex()
Public Function GetMemberIndex(ByVal Username As String) As Integer
    GetMemberIndex = GetUserIndex(Username)
End Function

Public Function GetUserIndex(ByVal Username As String) As Integer
    Dim i As Integer
    
    For i = 1 To Members.Count
        If (StrComp(Members(i).DisplayName, Username, vbTextCompare) = 0) Then
            GetUserIndex = i
        
            Exit Function
        End If
    Next i
    
    GetUserIndex = 0
End Function

' alias for GetUserEx()
Public Function GetMemberEx(ByVal Username As String) As Object
    Set GetMemberEx = GetUserEx(Username)
End Function

Public Function GetUserEx(ByVal Username As String) As Object
    Dim Index As Integer
    
    If (StrictIsNumeric(Username)) Then
        Index = CInt(Val(Username))
    Else
        Index = GetUserIndexEx(Username)
    End If
    
    If ((Index >= 1) And (Index <= Members.Count)) Then
        Set GetUserEx = Members(Index)
    Else
        Set GetUserEx = New clsClanMemberObj
    End If
End Function

' alias for GetUserIndexEx
Public Function GetMemberIndexEx(ByVal Username As String) As Integer
    GetMemberIndexEx = GetUserIndexEx(Username)
End Function

Public Function GetUserIndexEx(ByVal Username As String) As Integer
    Dim i As Integer
    
    For i = 1 To Members.Count
        If (StrComp(Members(i).Name, Username, vbTextCompare) = 0) Then
            GetUserIndexEx = i
        
            Exit Function
        End If
    Next i
    
    GetUserIndexEx = 0
End Function

Public Sub Clear()
    Set m_Members = New Collection
End Sub

Public Sub Disband()
    Call frmChat.ClanHandler.DisbandClan(reqScriptingCall)
End Sub

Public Sub Invite(ByVal Username As String)
    Call frmChat.ClanHandler.InviteToClan(Username, reqScriptingCall)
End Sub

Public Sub FindCandidates(ByVal ClanTag As String)
    Call frmChat.ClanHandler.CreateFindCandidates(ClanTag, reqScriptingCall)
End Sub

Public Sub Create(ByVal ClanTag As String, ByVal ClanName As String, ByRef Users() As Variant)
    Dim Candidates() As String
    Dim i As Integer

    If LBound(Users) < UBound(Users) Then
        ReDim Candidates(0 To UBound(Users) - LBound(Users))
        For i = LBound(Users) To UBound(Users)
            Candidates(i) = CStr(Users(i + LBound(Users)))
        Next i
        Call frmChat.ClanHandler.CreateInviteMultiple(ClanTag, ClanName, Candidates(), reqScriptingCall)
    End If
End Sub

Public Sub RequestMemberList()
    Call frmChat.ClanHandler.RequestClanList(reqScriptingCall)
End Sub

Public Sub RequestMOTD()
    Call frmChat.ClanHandler.RequestClanMOTD(reqScriptingCall)
End Sub

Public Sub RequestMemberInfo(ByVal Username As String, ByVal ClanTag As String)
    Call frmChat.ClanHandler.RequestClanMemberInfo(Username, ClanTag, reqScriptingCall)
End Sub

Public Sub SetMOTD(ByVal MOTD As String)
    Call frmChat.ClanHandler.SetClanMOTD(MOTD, reqScriptingCall)
End Sub

Public Function Clone() As Object
    Dim i As Integer
    Set Clone = New clsClanObj
    
    Clone.Name = Name
    Clone.MOTD = MOTD
    
    For i = 1 To Members.Count
        Clone.Members.Add Members(i).Clone()
    Next i
End Function
