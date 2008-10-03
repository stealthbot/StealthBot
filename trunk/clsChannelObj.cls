VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannelObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private Const CHAN_PUBLIC     As Long = &H1
Private Const CHAN_MODERATED  As Long = &H2
Private Const CHAN_RESTRICTED As Long = &H4
Private Const CHAN_SILENT     As Long = &H8
Private Const CHAN_SYSTEM     As Long = &H10
Private Const CHAN_PRODUCT    As Long = &H20
Private Const CHAN_GLOBAL     As Long = &H1000

Private m_name             As String
Private m_flags            As Long
Private m_designated_heir  As String
Private m_num_joins        As Long
Private m_num_bans         As Long
Private m_num_kicks        As Long
Private m_join_date        As Date
Private m_users            As Collection
Private m_banlist          As Collection

Public Property Get SType() As String
    
    Dim tmp As String ' ...
    
    ' ...
    If ((Flags And CHAN_PUBLIC) = CHAN_PUBLIC) Then
        tmp = tmp & "public, "
    End If
    
    ' ...
    If ((Flags And CHAN_MODERATED) = CHAN_MODERATED) Then
        tmp = tmp & "moderated, "
    End If
    
    ' ...
    If ((Flags And CHAN_RESTRICTED) = CHAN_RESTRICTED) Then
        tmp = tmp & "restricted, "
    End If
    
    ' ...
    If ((Flags And CHAN_SILENT) = CHAN_SILENT) Then
        tmp = tmp & "silent, "
    End If
    
    ' ...
    If ((Flags And CHAN_SYSTEM) = CHAN_SYSTEM) Then
        tmp = tmp & "system, "
    End If
    
    ' ...
    If ((Flags And CHAN_PRODUCT) = CHAN_PRODUCT) Then
        tmp = tmp & "product-specific, "
    End If
    
    ' ...
    If ((Flags And CHAN_GLOBAL) = CHAN_GLOBAL) Then
        tmp = tmp & "global, "
    End If
    
    ' ...
    If (Flags = &H0) Then
        tmp = "private, "
    End If
    
    ' ...
    tmp = Mid$(tmp, 1, Len(tmp) - 2)
    
    ' ...
    SType = tmp
    
End Property


Private Sub Class_Initialize()
    
    ' ...
    Set m_users = New Collection
    Set m_banlist = New Collection
    
End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_users = Nothing
    Set m_banlist = Nothing

End Sub

Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

Public Property Let Name(str As String)

    ' ...
    m_name = str

End Property

Public Property Get OperatorHeir() As String

    ' ...
    OperatorHeir = m_designated_heir

End Property

Public Property Let OperatorHeir(str As String)

    ' ...
    m_designated_heir = str

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Let JoinTime(JDate As Date)

    ' ...
    m_join_date = JDate

End Property

Public Property Get JoinTime() As Date

    ' ...
    JoinTime = m_join_date

End Property

Public Property Get IsSilent() As Boolean

    ' ...
    IsSilent = ((m_flags And CHAN_SILENT) = CHAN_SILENT)

End Property

Public Property Get JoinCount() As Long

    ' ...

End Property

Public Property Get BanCount() As Long

    ' ...
    BanCount = m_num_bans

End Property

Public Property Let BanCount(lng As Long)

    ' ...
    m_num_bans = lng
    
End Property

Public Property Get KickCount() As Long

    ' ...

End Property

Public Property Get Users() As Collection

    ' ...
    Set Users = m_users

End Property

Public Sub ClearUsers()

    ' ...
    Set m_users = New Collection

End Sub

Public Sub ClearBanlist()

    ' ...
    Set m_banlist = New Collection

End Sub

Public Property Get Self() As clsUserObj

    Dim I As Integer ' ...

    ' ...
    Set Self = New clsUserObj
    
    ' ...
    For I = 1 To Users.Count
        ' ...
        If (StrComp(Users(I).Name, CurrentUsername, vbTextCompare) = 0) Then
            ' ...
            Set Self = Users(I)
            
            ' ...
            Exit Property
        End If
    Next I

End Property

Public Function GetUserIndexEx(ByVal AccountName As String, Optional SearchLimit As Integer = 0) As Integer

    Dim I As Integer ' ...
    
    ' ...
    AccountName = modEvents.CleanUsername(AccountName)

    ' ...
    For I = 1 To Users.Count
        ' ...
        If (SearchLimit > 0) Then
            ' ...
            If (I >= SearchLimit) Then
                Exit For
            End If
        End If
    
        ' ...
        If (StrComp(Users(I).Name, AccountName, vbTextCompare) = 0) Then
            ' ...
            GetUserIndexEx = I
            
            ' ...
            Exit Function
        End If
    Next I
    
    ' ...
    GetUserIndexEx = 0

End Function

Public Function GetUserIndex(ByVal Username As String, Optional SearchLimit As Integer = 0) As Integer

    Dim I As Integer ' ...

    ' ...
    For I = 1 To m_users.Count
        ' ...
        If (SearchLimit > 0) Then
            ' ...
            If (I >= SearchLimit) Then
                Exit For
            End If
        End If
    
        ' ...
        If (StrComp(m_users(I).DisplayName, Username, vbTextCompare) = 0) Then
            ' ...
            GetUserIndex = I
            
            ' ...
            Exit Function
        End If
    Next I
    
    ' ...
    GetUserIndex = 0

End Function

Public Property Get Banlist() As Collection
    
    ' ...
    Set Banlist = m_banlist

End Property

Public Function IsOnBanList(ByVal Username As String, Optional Operator As String = vbNullString) As Integer

    Dim I   As Integer ' ...
    Dim bln As Boolean ' ...
    
    ' ...
    Username = CleanUsername(Username)
    
    ' ...
    For I = 1 To Banlist.Count
        ' ...
        If (StrComp(Banlist(I).Name, Username, vbTextCompare) = 0) Then
            ' ...
            If (Operator <> vbNullString) Then
                ' ...
                If (StrComp(Banlist(I).Operator, Operator, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
        
            ' ...
            If (bln) Then
                IsOnBanList = I
                
                Exit Function
            End If
        End If
    Next I

    IsOnBanList = 0
End Function

Public Sub RemoveBansFromOperator(ByVal Username As String)

    Dim I    As Integer ' ...
    Dim bln  As Boolean ' ...
    Dim Pos  As Integer ' ...
    Dim Name As String  ' ...
    
    ' ...
    Username = CleanUsername(Username)

    ' ...
    Do
        ' ...
        bln = False
    
        ' ...
        For I = 1 To Banlist.Count
            ' ...
            If (StrComp(Banlist(I).Operator, Username, vbTextCompare) = 0) Then
                ' ...
                Name = Banlist(I).Name
            
                ' ...
                If (BotVars.RetainOldBans = False) Then
                    ' ...
                    Banlist.Remove I
                    
                    ' ...
                    Pos = IsOnBanList(Name)
                    
                    ' ...
                    If (Pos > 0) Then
                        Banlist(Pos).IsDuplicateBan = False
                    End If
                    
                    ' ...
                    bln = True
                    
                    ' ...
                    Exit For
                Else
                    ' ...
                    Banlist(Pos).IsActive = False
                End If
            End If
        Next I
    Loop While (bln = True)

End Sub

' ...
Public Function CheckUsers() As Integer

    ' ...
    If (Self.IsOperator) Then
        Dim CurrentUser As clsUserObj
        Dim DBEntry     As udtGetAccessResponse
        Dim I           As Integer ' ...
        Dim Message     As String  ' ...
        Dim doCheck     As Boolean ' ...
        Dim HighRank    As Integer ' ...
        Dim HighIndex   As Integer ' ...
        Dim Count       As Integer ' ...
        
        ' ...
        doCheck = True
        
        ' ...
        If (OperatorHeir = vbNullString) Then
            ' ...
            For I = 1 To Users.Count
                ' ...
                Set CurrentUser = Users(I)
            
                ' ...
                If (CurrentUser.IsOperator = False) Then
                    ' ...
                    DBEntry = _
                        GetCumulativeAccess(CurrentUser.DisplayName, "USER")
                
                    ' ...
                    If (InStr(1, DBEntry.Flags, "D", vbBinaryCompare) <> 0) Then
                        ' ...
                        If (DBEntry.Access > HighRank) Then
                            ' ...
                            HighRank = DBEntry.Access
                            
                            ' ...
                            HighIndex = I
                        End If
                    End If
                End If
            Next I
            
            ' ...
            If (HighIndex > 0) Then
                ' ...
                frmChat.AddQ "/designate " & Users(HighIndex).DisplayName
            End If
        End If
    
        ' ...
        For I = Users.Count To 1 Step -1
            ' ...
            Set CurrentUser = Users(I)
        
            ' ...
            If (CurrentUser.IsOperator = False) Then
                ' ...
                Message = GetShitlist(CurrentUser.DisplayName)
                       
                ' ...
                If (Message <> vbNullString) Then
                    ' ...
                    frmChat.AddQ "/ban " & Message
                    
                    ' ...
                    doCheck = False
                Else
                    ' ...
                    If (GetSafelist(CurrentUser.DisplayName) = False) Then
                        ' ...
                        If (Protect) Then
                            ' ...
                            Ban CurrentUser.DisplayName & Space$(1) & ProtectMsg, (AutoModSafelistValue - 1)
                            
                            ' ...
                            doCheck = False
                        Else
                            Dim j As Integer ' ...
                
                            ' ...
                            If ((doCheck) And (BotVars.BanEvasion)) Then
                                ' ...
                                If (IsOnBanList(CurrentUser.Name)) Then
                                    ' ...
                                    Ban CurrentUser.DisplayName & " Ban Evasion", (AutoModSafelistValue - 1)
                                    
                                    ' ...
                                    doCheck = False
                                End If
                            End If
                            
                            ' ...
                            If ((doCheck) And (BotVars.IPBans)) Then
                                ' ...
                                If (CurrentUser.IsSquelched) Then
                                    ' ...
                                    Ban CurrentUser.DisplayName & " IPBanned.", (AutoModSafelistValue - 1)
                                    
                                    ' ...
                                    doCheck = False
                                End If
                            End If
                            
                            ' ...
                            If ((doCheck) And (BotVars.PlugBan)) Then
                                ' ...
                                If ((CurrentUser.Flags And USER_NOUDP) = USER_NOUDP) Then
                                    ' ...
                                    Ban CurrentUser.DisplayName & " PlugBan", (AutoModSafelistValue - 1)
                                    
                                    ' ...
                                    doCheck = False
                                End If
                            End If
                            
                            ' ...
                            If (doCheck) Then
                                ' ...
                                If (CurrentUser.IsUsingDII) Then
                                    ' ...
                                    If (BotVars.BanD2UnderLevel) Then
                                        ' ...
                                        If (CurrentUser.Stats.Level < BotVars.BanD2UnderLevel) Then
                                            ' ...
                                            Ban CurrentUser.DisplayName & Space$(1) & "You are under the required " & _
                                                "level for entry.", (AutoModSafelistValue - 1)
                                            
                                            ' ...
                                            doCheck = False
                                        End If
                                    End If
                                ElseIf (CurrentUser.IsUsingWarIII) Then
                                    ' ...
                                    If (BotVars.BanPeons) Then
                                        ' ...
                                        If (CurrentUser.Stats.Tier = 1) Then
                                            ' ...
                                            Ban CurrentUser.DisplayName & Space$(1) & "PeonBan", _
                                                (AutoModSafelistValue - 1)
                                        
                                            ' ...
                                            doCheck = False
                                        End If
                                    End If
                                End If
                            End If
                            
                            ' ...
                            'If (doCheck) Then
                            '    ' ...
                            '    If ((BotVars.ChannelPasswordDelay) And (Len(BotVars.ChannelPassword) > 0)) Then
                            '        ' ...
                            '        If (CurrentUser.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
                            '            ' ...
                            '            Ban CurrentUser.DisplayName & " Password time is up", _
                            '                (AutoModSafelistValue - 1)
                            '
                            '            ' ...
                            '            doCheck = False
                            '        End If
                            '    End If
                            'End If
                            '
                            ' ...
                            'If (doCheck) Then
                            '    ' ...
                            '    If ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0)) Then
                            '        ' ...
                            '        If (CurrentUser.TimeSinceTalk() > BotVars.IB_Wait) Then
                            '            ' ...
                            '            Ban CurrentUser.DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", _
                            '                (AutoModSafelistValue - 1), IIf(BotVars.IB_Kick, 1, 0)
                            '
                            '            ' ...
                            '            doCheck = False
                            '        End If
                            '    End If
                            'End If
                        End If
                    End If
                End If
            End If
            
            ' ...
            If (doCheck = False) Then
                Count = (Count + 1)
            End If
            
            ' ...
            Message = vbNullString
            
            ' ...
            doCheck = True
        Next I
    End If
    
    ' ...
    CheckUsers = Count
    
End Function

Public Function CheckQueue(ByVal Username As String) As Boolean

    Dim CurrentUser As clsUserObj

    ' ...
    For Each CurrentUser In Users
        ' ...
        If (StrComp(CurrentUser.DisplayName, Username, vbTextCompare) = 0) Then
            ' ...
            If (CurrentUser.Queue.Count) Then
                CheckQueue = True
            End If
            
            ' ...
            Exit For
        End If
    Next

End Function
