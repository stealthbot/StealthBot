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
' ...

Option Explicit

Private m_name    As String
Private m_motd    As String
Private m_members As Collection

Private Sub Class_Initialize()

    ' ...
    Set m_members = New Collection

End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_members = Nothing

End Sub

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
Public Property Get MOTD() As String

    ' ...
    MOTD = m_motd

End Property

' ...
Public Property Let MOTD(str As String)

    ' ...
    m_motd = str

End Property

' ...
Public Property Get Self() As clsClanMemberObj

    Dim i As Integer ' ...

    ' ...
    Set Self = New clsClanMemberObj
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (StrComp(Members(i).Name, CurrentUsername, vbTextCompare) = 0) Then
            ' ...
            Set Self = Members(i)
        
            ' ...
            Exit For
        End If
    Next i

End Property

' ...
Public Property Get Members() As Collection

    ' ...
    Set Members = m_members

End Property

' ...
Public Property Get Chieftain() As clsClanMemberObj

    Dim i As Integer ' ...

    ' ...
    Set Chieftain = New clsClanMemberObj
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (Members(i).Rank >= 4) Then
            ' ...
            Set Chieftain = Members(i)
        
            ' ...
            Exit For
        End If
    Next i

End Property

' ...
Public Property Get Shamans() As Collection

    Dim i As Integer ' ...

    ' ...
    Set Shamans = New Collection
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (Members(i).Rank = 3) Then
            Shamans.Add Members(i)
        End If
    Next i

End Property

' ...
Public Property Get Grunts() As Collection

    Dim i As Integer ' ...

    ' ...
    Set Grunts = New Collection
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (Members(i).Rank = 2) Then
            Grunts.Add Members(i)
        End If
    Next i

End Property

' ...
Public Property Get Peons() As Collection

    Dim i As Integer ' ...

    ' ...
    Set Peons = New Collection
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If ((Members(i).Rank >= 0) And (Members(i).Rank <= 0)) Then
            Peons.Add Members(i)
        End If
    Next i

End Property

Public Function GetMember(ByVal Username As String) As Object

    Set GetMember = GetUser(Username)

End Function

Public Function GetUser(ByVal Username As String) As Object

    Dim Index As Integer ' ...
    
    ' ...
    If (StrictIsNumeric(Username)) Then
        Index = CInt(Val(Username))
    Else
        Index = GetUserIndex(Username)
    End If
    
    ' ...
    If ((Index >= 1) And (Index <= Members.Count)) Then
        Set GetUser = Members(Index)
    Else
        Set GetUser = New clsClanMemberObj
    End If

End Function

Public Function GetMemberIndex(ByVal Username As String) As Integer

    GetMemberIndex = GetUserIndex(Username)

End Function

Public Function GetUserIndex(ByVal Username As String) As Integer

    Dim i As Integer ' ...
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (StrComp(Members(i).DisplayName, Username, vbTextCompare) = 0) Then
            ' ...
            GetUserIndex = i
        
            ' ...
            Exit Function
        End If
    Next i
    
    ' ...
    GetUserIndex = 0

End Function

Public Function GetMemberEx(ByVal Username As String) As Object

    Set GetMemberEx = GetUserEx(Username)

End Function

Public Function GetUserEx(ByVal Username As String) As Object

    Dim Index As Integer ' ...
    
    ' ...
    If (StrictIsNumeric(Username)) Then
        Index = CInt(Val(Username))
    Else
        Index = GetUserIndexEx(Username)
    End If
    
    ' ...
    If ((Index >= 1) And (Index <= Members.Count)) Then
        Set GetUserEx = Members(Index)
    Else
        Set GetUserEx = New clsClanMemberObj
    End If

End Function

Public Function GetMemberIndexEx(ByVal Username As String) As Integer
    
    GetMemberIndexEx = GetUserIndexEx(Username)
    
End Function

Public Function GetUserIndexEx(ByVal Username As String) As Integer

    Dim i As Integer ' ...
    
    ' ...
    For i = 1 To Members.Count
        ' ...
        If (StrComp(Members(i).Name, Username, vbTextCompare) = 0) Then
            ' ...
            GetUserIndexEx = i
        
            ' ...
            Exit Function
        End If
    Next i
    
    ' ...
    GetUserIndexEx = 0

End Function

Public Sub Clear()
    
    ' ...
    Set m_members = New Collection

End Sub

Public Sub Disband()

    Call DisbandClan

End Sub

Public Sub SetMOTD(ByVal MOTD As String)

    Call modWar3Clan.SetClanMOTD(MOTD)

End Sub

Public Function Clone() As Object

    Dim i As Integer ' ...

    ' ...
    Set Clone = New clsClanObj
    
    ' ...
    Clone.Name = Name
    Clone.MOTD = MOTD
    
    ' ...
    For i = 1 To Members.Count
        Clone.Members.Add Members(i).Clone()
    Next i
    
End Function
