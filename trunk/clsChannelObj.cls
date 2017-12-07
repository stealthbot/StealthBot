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

Option Explicit

Private Const CHAN_PUBLIC     As Long = &H1
Private Const CHAN_MODERATED  As Long = &H2
Private Const CHAN_RESTRICTED As Long = &H4
Private Const CHAN_SILENT     As Long = &H8
Private Const CHAN_SYSTEM     As Long = &H10
Private Const CHAN_PRODUCT    As Long = &H20
Private Const CHAN_GLOBAL     As Long = &H1000

Private m_Name           As String
Private m_Flags          As Long
Private m_DesignatedHeir As String
Private m_NumJoins       As Long
Private m_NumBans        As Long
Private m_NumKicks       As Long
Private m_JoinDate       As Date
Private m_Users          As Collection
Private m_Banlist        As Collection

Public Property Get sType() As String
    
    Dim tmp As String
    
    If ((Flags And CHAN_RESTRICTED) = CHAN_RESTRICTED) Then
        tmp = tmp & "restricted, "
    End If
    
    If ((Flags And CHAN_GLOBAL) = CHAN_GLOBAL) Then
        tmp = tmp & "global, "
    End If
    
    If ((Flags And CHAN_PUBLIC) = CHAN_PUBLIC) Then
        tmp = tmp & "public, "
    End If
    
    If ((Flags And CHAN_MODERATED) = CHAN_MODERATED) Then
        tmp = tmp & "moderated, "
    End If
    
    If ((Flags And CHAN_PRODUCT) = CHAN_PRODUCT) Then
        tmp = tmp & "product-specific, "
    End If
    
    If ((Flags And CHAN_SYSTEM) = CHAN_SYSTEM) Then
        tmp = tmp & "system, "
    End If
    
    If ((Flags And CHAN_SILENT) = CHAN_SILENT) Then
        tmp = tmp & "silent, "
    End If
    
    If (Flags = &H0) Then
        tmp = "private, "
    End If
    
    tmp = Mid$(tmp, 1, Len(tmp) - 2)
    
    sType = tmp
    
End Property


Private Sub Class_Initialize()
    
    Set m_Users = New Collection
    Set m_Banlist = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_Users = Nothing
    Set m_Banlist = Nothing
    
End Sub

Public Property Get Name() As String

    Name = m_Name

End Property

Public Property Let Name(str As String)

    m_Name = str

End Property

Public Property Get OperatorHeir() As String

    OperatorHeir = m_DesignatedHeir

End Property

Public Property Let OperatorHeir(str As String)

    m_DesignatedHeir = str

End Property

Public Property Get Flags() As Long

    Flags = m_Flags

End Property

Public Property Let Flags(ByVal lng As Long)

    m_Flags = lng

End Property

Public Property Let JoinTime(ByVal JDate As Date)

    m_JoinDate = JDate

End Property

Public Property Get JoinTime() As Date

    JoinTime = m_JoinDate

End Property

Public Property Get IsSilent() As Boolean

    IsSilent = ((m_Flags And CHAN_SILENT) = CHAN_SILENT)

End Property

Public Property Get JoinCount() As Long

    JoinCount = m_NumJoins

End Property

Public Property Let JoinCount(ByVal lng As Long)

    m_NumJoins = lng

End Property

Public Property Get BanCount() As Long

    BanCount = m_NumBans

End Property

Public Property Let BanCount(ByVal lng As Long)

    m_NumBans = lng
    
End Property

Public Property Get KickCount() As Long

    KickCount = m_NumKicks

End Property

Public Property Let KickCount(ByVal lng As Long)

    m_NumKicks = lng

End Property

Public Property Get Users() As Collection

    Set Users = m_Users

End Property

Public Property Get PriorityUsers() As Collection

    Dim i As Integer
    Set PriorityUsers = New Collection
    
    For i = Users.Count To 1 Step -1
        If (Users(i).IsOperator) Then
            PriorityUsers.Add Users(i)
        End If
    Next i
    For i = 1 To Users.Count
        If (Not Users(i).IsOperator) Then
            PriorityUsers.Add Users(i)
        End If
    Next i

End Property

Public Property Get Operators() As Collection

    Dim i As Integer
    Set Operators = New Collection
    
    For i = 1 To Users.Count
        If (Users(i).IsOperator) Then
            Operators.Add Users(i)
        End If
    Next i

End Property


Public Sub ClearUsers()

    Set m_Users = Nothing

    Set m_Users = New Collection

End Sub

Public Sub ClearBanlist()

    Set m_Banlist = New Collection

End Sub

Public Property Get Self() As clsUserObj

    Dim i As Integer

    Set Self = New clsUserObj
    
    For i = Users.Count To 1 Step -1
        If (StrComp(Users(i).Name, CleanUsername(CurrentUsername), vbTextCompare) = 0) And (Users(i).IsPhantom = False) Then
            Set Self = Users(i)
            Exit Property
        End If
    Next i

End Property

Public Function GetUserEx(ByVal AccountName As String, Optional ByVal SearchLimit As Integer = 0) As Object

    Dim Index As Integer
    
    If (StrictIsNumeric(AccountName)) Then
        Index = CInt(Val(AccountName))
    Else
        Index = GetUserIndexEx(AccountName, SearchLimit)
    End If
    
    If ((Index >= 1) And (Index <= Users.Count)) Then
        Set GetUserEx = Users(Index)
    Else
        Set GetUserEx = New clsUserObj
    End If

End Function

Public Function GetUserIndexEx(ByVal AccountName As String, Optional ByVal SearchLimit As Integer = 0) As Integer

    Dim i As Integer
    
    AccountName = modEvents.CleanUsername(AccountName)

    For i = Users.Count To 1 Step -1
        If (SearchLimit > 0) Then
            If (i >= SearchLimit) Then
                Exit For
            End If
        End If
    
        If (StrComp(Users(i).Name, AccountName, vbTextCompare) = 0) And (Users(i).IsPhantom = False) Then
            GetUserIndexEx = i
            Exit Function
        End If
    Next i
    
    GetUserIndexEx = 0

End Function

Public Function GetUser(ByVal Username As String, Optional ByVal SearchLimit As Integer = 0) As Object

    Dim Index As Integer
    
    If (StrictIsNumeric(Username) And (Len(Username) <= 4)) Then
        Index = CInt(Val(Username))
        If (Index > Users.Count) Then
            Index = GetUserIndex(Username, SearchLimit)
        End If
    Else
        Index = GetUserIndex(Username, SearchLimit)
    End If
    
    If ((Index >= 1) And (Index <= Users.Count)) Then
        Set GetUser = Users(Index)
    Else
        Set GetUser = New clsUserObj
    End If

End Function

Public Function GetUserIndex(ByVal Username As String, Optional ByVal SearchLimit As Integer = 0) As Integer

    Dim i As Integer

    For i = m_Users.Count To 1 Step -1
        If (SearchLimit > 0) Then
            If (i >= SearchLimit) Then
                Exit For
            End If
        End If
    
        If (StrComp(m_Users(i).DisplayName, Username, vbTextCompare) = 0) And (Users(i).IsPhantom = False) Then
            GetUserIndex = i
            Exit Function
        End If
    Next i
    
    GetUserIndex = 0

End Function

Public Function GetUserIndexByPriority(ByVal Username As String, Optional ByVal SearchLimit As Integer = 0, Optional ByVal SkipDelayed As Boolean = False) As Integer

    Dim i     As Integer
    Dim pos   As Integer
    Dim PUser As clsUserObj

    pos = 0
    For i = 1 To PriorityUsers.Count
        Set PUser = PriorityUsers(i)
        If SkipDelayed Imp (PUser.Queue.Count = 0) Then
            pos = pos + 1
            If (SearchLimit > 0) Then
                If (pos >= SearchLimit) Then
                    Exit For
                End If
            End If

            If (StrComp(PUser.Name, Username, vbTextCompare) = 0) And (PUser.IsPhantom = False) Then
                GetUserIndexByPriority = pos
                Exit Function
            End If
        End If
    Next i
    
    GetUserIndexByPriority = 0

End Function

Public Property Get Banlist() As Collection
    
    Set Banlist = m_Banlist

End Property

Public Function IsOnBanList(ByVal Username As String, Optional ByVal Operator As String = vbNullString) As Integer

    Dim i   As Integer
    Dim bln As Boolean
    
    Username = CleanUsername(Username)
    
    For i = 1 To Banlist.Count
        If (StrComp(Banlist(i).Name, Username, vbTextCompare) = 0) Then
            If (Operator <> vbNullString) Then
                If (StrComp(Banlist(i).Operator, Operator, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
        
            If (bln) Then
                IsOnBanList = i
                Exit Function
            End If
        End If
    Next i

    IsOnBanList = 0
    
End Function

Public Function IsOnRecentBanList(ByVal Username As String, Optional ByVal Operator As String = vbNullString) As Integer

    Dim i   As Integer
    Dim bln As Boolean
    
    Username = CleanUsername(Username)
    
    For i = Banlist.Count To (Banlist.Count - 5) Step -1
        If (i <= 0) Then
            Exit For
        End If

        If (StrComp(Banlist(i).Name, Username, vbTextCompare) = 0) Then
            If (Operator <> vbNullString) Then
                If (StrComp(Banlist(i).Operator, Operator, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
        
            If (bln) Then
                'frmChat.AddChat vbYellow, DateDiff("s", Banlist(I).DateOfBan, UtcNow)
            
                If (DateDiff("s", Banlist(i).DateOfBan, UtcNow) <= 3) Then
                    IsOnRecentBanList = i
                
                    Exit Function
                End If
            End If
        End If
    Next i

    IsOnRecentBanList = 0
    
End Function

Public Sub RemoveBansFromOperator(ByVal Username As String)

    Dim i    As Integer
    Dim bln  As Boolean
    Dim pos  As Integer
    Dim Name As String
    
    Username = CleanUsername(Username)

    Do
        bln = False
    
        For i = 1 To Banlist.Count
            If (StrComp(Banlist(i).Operator, Username, vbTextCompare) = 0) Then
                Name = Banlist(i).Name
            
                If (BotVars.RetainOldBans = False) Then
                    Banlist.Remove i
                    
                    pos = IsOnBanList(Name)
                    
                    If (pos > 0) Then
                        Banlist(pos).IsDuplicateBan = False
                    End If
                    
                    bln = True
                    
                    Exit For
                Else
                    Banlist(i).IsActive = False
                End If
            End If
        Next i
    Loop While (bln = True)

End Sub

Public Function CheckUser(Username As String, Optional ByRef CurrentUser As clsUserObj = Nothing) As Integer

    Dim doCheck As Boolean

    doCheck = True

    If (Self.IsOperator) Then
        Dim DBEntry  As udtUserAccess
        Dim i        As Integer
        Dim Message  As String
        
        If (CurrentUser Is Nothing) Then
            Set CurrentUser = Users(GetUserIndex(Username))
        End If
        
        DBEntry = Database.GetUserAccess(CurrentUser.DisplayName)
        
        If (DBEntry.Rank < AutoModSafelistValue) Then
            ' Not high enough rank to be safe from auto-mod
            
            If (InStr(1, DBEntry.Flags, "S", vbBinaryCompare) = 0) Then
                ' Not on the safelist
                
                If (InStr(1, DBEntry.Flags, "B", vbBinaryCompare) <> 0) Then
                    ' On the shitlist
                    
                    Message = GetShitlist(CurrentUser.DisplayName)
                    
                    frmChat.AddQ "/ban " & Message
                    
                    doCheck = False
                Else
                    
                    ' Is channel protection enabled?
                    If (Protect) Then
                        Ban CurrentUser.DisplayName & Space$(1) & ProtectMsg, (AutoModSafelistValue - 1)
                        
                        doCheck = False
                    Else
                        Dim j As Integer
                        
                        ' Have they previously been banned?
                        If ((doCheck) And (BotVars.BanEvasion)) Then
                            If (IsOnBanList(CurrentUser.Name)) Then
                                Ban CurrentUser.DisplayName & " Ban Evasion", (AutoModSafelistValue - 1)
                                
                                doCheck = False
                            End If
                        End If
                        
                        ' Are they IP banned?
                        If ((doCheck) And (BotVars.IPBans)) Then
                            If (CurrentUser.IsSquelched) Then
                                Ban CurrentUser.DisplayName & " IP Banned", (AutoModSafelistValue - 1)
                                
                                doCheck = False
                            End If
                        End If
                        
                        ' Is the user's ping acceptable?
                        If ((doCheck) And (Config.PingBan)) Then
                            If (Config.PingBanLevel < 1) Then
                                If (CurrentUser.Ping = Config.PingBanLevel) Then
                                    Ban CurrentUser.DisplayName & " Ping Ban", (AutoModSafelistValue - 1)
                                    
                                    doCheck = False
                                End If
                            Else
                                If (CurrentUser.Ping > Config.PingBanLevel) Then
                                    Ban CurrentUser.DisplayName & " Ping Ban", (AutoModSafelistValue - 1)
                                    
                                    doCheck = False
                                End If
                            End If
                        End If
                        
                        ' Do they have a UDP plug?
                        If ((doCheck) And (BotVars.PlugBan)) Then
                            If ((CurrentUser.Flags And USER_NOUDP) = USER_NOUDP) Then
                                Ban CurrentUser.DisplayName & " PlugBan", (AutoModSafelistValue - 1)
                                
                                doCheck = False
                            End If
                        End If
                        
                        ' Level bans
                        If (doCheck) Then
                            If (CurrentUser.IsUsingDII) Then
                                If (BotVars.BanD2UnderLevel) Then
                                    If (CurrentUser.Stats.Level < BotVars.BanD2UnderLevel) Then
                                        Message = BotVars.BanUnderLevelMsg
                                        
                                        If LenB(Message) = 0 Then
                                            Message = "You are below the required level for entry."
                                        End If
                                        
                                        If InStr(1, Message, "%cl", vbTextCompare) > 0 Then
                                            Message = Replace$(Message, "%cl", CStr(CurrentUser.Stats.Level))
                                        End If
                                        
                                        If InStr(1, Message, "%rl", vbTextCompare) > 0 Then
                                            Message = Replace$(Message, "%rl", CStr(BotVars.BanD2UnderLevel))
                                        End If
                                        
                                        Ban CurrentUser.DisplayName & Space$(1) & Message, _
                                            (AutoModSafelistValue - 1)
                                        
                                        doCheck = False
                                    End If
                                End If
                            ElseIf (CurrentUser.IsUsingWarIII) Then
                                If (BotVars.BanPeons) Then
                                    If (StrComp(CurrentUser.Stats.IconName, "peon", vbTextCompare) = 0) Then
                                        Ban CurrentUser.DisplayName & Space$(1) & "PeonBan", _
                                            (AutoModSafelistValue - 1)
                                        
                                        doCheck = False
                                    End If
                                ElseIf (BotVars.BanUnderLevel) Then
                                    If (CurrentUser.Stats.Level < BotVars.BanUnderLevel) Then
                                        Message = BotVars.BanUnderLevelMsg
                                        
                                        If LenB(Message) = 0 Then
                                            Message = "You are below the required level for entry."
                                        End If
                                        
                                        If InStr(1, Message, "%cl", vbTextCompare) > 0 Then
                                            Message = Replace$(Message, "%cl", CStr(CurrentUser.Stats.Level))
                                        End If
                                        
                                        If InStr(1, Message, "%rl", vbTextCompare) > 0 Then
                                            Message = Replace$(Message, "%rl", CStr(BotVars.BanUnderLevel))
                                        End If
                                        
                                        Ban CurrentUser.DisplayName & Space$(1) & Message, _
                                            (AutoModSafelistValue - 1)
                                        
                                        doCheck = False
                                    End If
                                End If
                            End If
                        End If
                        
                        'If (doCheck) Then
                        '    If ((BotVars.ChannelPasswordDelay) And (Len(BotVars.ChannelPassword) > 0)) Then
                        '        If (CurrentUser.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
                        '            Ban CurrentUser.DisplayName & " Password time is up", _
                        '                (AutoModSafelistValue - 1)
                        '
                        '            doCheck = False
                        '        End If
                        '    End If
                        'End If
                        '
                        'If (doCheck) Then
                        '    If ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0)) Then
                        '        If (CurrentUser.TimeSinceTalk() > BotVars.IB_Wait) Then
                        '            Ban CurrentUser.DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", _
                        '                (AutoModSafelistValue - 1), IIf(BotVars.IB_Kick, 1, 0)
                        '
                        '            doCheck = False
                        '        End If
                        '    End If
                        'End If
                    End If
                End If
            End If
        End If
    End If
    
    If (doCheck = False) Then
        CurrentUser.PendingBan = True
    End If

End Function

Public Function CheckUsers() As Integer

    If (Self.IsOperator) Then
        Dim CurrentUser As clsUserObj
        Dim DBEntry     As udtUserAccess
        Dim i           As Integer
        Dim Message     As String
        Dim doCheck     As Boolean
        Dim HighRank    As Integer
        Dim HighIndex   As Integer
        Dim Count       As Integer
        
        doCheck = True
        
        If (OperatorHeir = vbNullString) Then
            For i = Users.Count To 1 Step -1
                Set CurrentUser = Users(i)
            
                If (CurrentUser.IsOperator = False) And (CurrentUser.IsPhantom = False) Then
                    DBEntry = _
                        Database.GetUserAccess(CurrentUser.DisplayName)
                
                    If (InStr(1, DBEntry.Flags, "D", vbBinaryCompare) <> 0) Then
                        If (DBEntry.Rank > HighRank) Then
                            HighRank = DBEntry.Rank
                            
                            HighIndex = i
                        End If
                    End If
                End If
            Next i
            
            If (HighIndex > 0) Then
                frmChat.AddQ "/designate " & Users(HighIndex).DisplayName
            End If
        End If
    
        For i = Users.Count To 1 Step -1
            Set CurrentUser = Users(i)
        
            If (CurrentUser.IsOperator = False) And (CurrentUser.IsPhantom = False) Then
                If (CheckUser(CurrentUser.DisplayName, CurrentUser)) Then
                    Count = (Count + 1)
                End If
            End If
        Next i
    End If
    
    CheckUsers = Count
    
End Function

Public Function CheckQueue(ByVal Username As String) As Boolean

    Dim i           As Integer
    Dim CurrentUser As clsUserObj

    For i = Users.Count To 1 Step -1
        Set CurrentUser = Users(i)
        If (StrComp(CurrentUser.DisplayName, Username, vbTextCompare) = 0) And (CurrentUser.IsPhantom = False) Then
            If (CurrentUser.Queue.Count) Then
                CheckQueue = True
            End If
            Exit For
        End If
    Next i

End Function

Public Function Clone() As Object

    Dim i As Integer

    Set Clone = New clsChannelObj
    
    Clone.Name = Name
    Clone.Flags = Flags
    Clone.JoinTime = JoinTime
    Clone.BanCount = BanCount
    Clone.KickCount = KickCount
    Clone.JoinCount = JoinCount
    Clone.OperatorHeir = OperatorHeir
    
    For i = 1 To Users.Count
        Clone.Users.Add Users(i).Clone()
    Next i

    For i = 1 To Banlist.Count
        Clone.Banlist.Add Banlist(i).Clone()
    Next i

End Function
