Attribute VB_Name = "modEvents"
'StealthBot Project - modEvents.bas
' Andy T (andy@stealthbot.net) March 2005

Option Explicit

Public Sub Event_FlagsUpdate(ByVal Username As String, ByVal Message As String, ByVal Flags As Long, _
    ByVal Ping As Long, ByVal Product As String)
    
    On Error GoTo ERROR_HANDLER
    
    Dim clsChatQueue As clsChatQueue
    Dim UserObj      As clsChannelUserObj
    
    Dim UserIndex    As Integer  ' ...
    Dim i            As Integer  ' ...
    Dim prevflags    As Long     ' ...
    Dim Clan         As String
    Dim parsed       As String
    
    ' if our username is for some reason null, we don't
    ' want to continue, possibly causing further errors
    If (LenB(Username) < 1) Then
        Exit Sub
    End If
    
    ' ...
    Call ParseStatstring(Message, parsed, Clan)
    
    ' ...
    UserIndex = _
        g_Channel.GetUserIndexByName(CleanDiablo2Username(Username))
    
    ' ...
    If (UserIndex > 0) Then
        ' ...
        Set UserObj = g_Channel.Users(UserIndex)
    Else
        Set UserObj = New clsChannelUserObj
            
        ' ...
        If (g_Channel.IsSilent = False) Then
            frmChat.AddChat vbRed, "Error: (1) There was a flags update received for a user that we do " & _
                    "not have a record for."
                    
            Exit Sub
        Else
            With frmChat.tmrSilentChannel(0)
                .Enabled = False
                .Enabled = True
            End With
        End If
    End If
    
    ' ...
    With UserObj
        .Name = Username
        .DisplayName = convertUsername(Username)
        .Flags = Flags
        .Ping = Ping
        .game = Product
        .Clan = Clan
    End With
    
    ' ...
    If (g_Channel.IsSilent) Then
        g_Channel.Users.Add UserObj
    End If

    ' convert username to appropriate
    ' display format
    Username = UserObj.DisplayName
    
    ' check for user in channel
    i = UsernameToIndex(Username)
    
    ' the user is already in the
    ' internal channel listings,
    ' right?
    If (i > 0) Then
        With colUsersInChannel.Item(i)
            ' create a copy of previous flags for determining
            ' if a user's flags have just been changed
            prevflags = .Flags
            
            ' update user flags
            .Flags = Flags
        End With
    Else
        If (g_Channel.IsSilent) Then
            Dim UserToAdd As clsUserInfo
            
            ' ...
            Set UserToAdd = New clsUserInfo
        
            ' ...
            With UserToAdd
                .Flags = Flags
                .Username = Username
                .Ping = Ping
                .Product = Product
                .Safelisted = GetSafelist(Username)
                .Statstring = Message
                .JoinTime = GetTickCount
                .Clan = Clan
                .IsSelf = (StrComp(Username, CurrentUsername, _
                    vbBinaryCompare) = 0)
                .InternalFlags = 0
            End With
            
            ' ...
            Call colUsersInChannel.Add(UserToAdd)
        Else
            frmChat.AddChat vbRed, "Error: (2) There was a flags update received for a user that we do " & _
                    "not have a record for."
                    
            Exit Sub
        End If
    End If
    
    ' are we receiving a flag update for ourselves?
    If (StrComp(Username, CurrentUsername, vbBinaryCompare) = 0) Then
        ' assign my current flags to the
        ' relevant internal variable
        MyFlags = Flags
        
        ' assign my current flags to the
        ' relevant scripting variable
        SharedScriptSupport.BotFlags = MyFlags
        
        ' if we're on ops, check for the presence of a user that we
        ' should designate as an heir
        If (g_Channel.Self.IsOperator) Then
            If (gChannel.Designated = vbNullString) Then
                ' loop through list of users
                For i = 1 To colUsersInChannel.Count
                    With GetCumulativeAccess(colUsersInChannel.Item(i).Username)
                        ' check for auto-designation flag
                        If (InStr(1, .Flags, "D", vbBinaryCompare) > 0) Then
                            ' is the user already an op?
                            If ((colUsersInChannel(i).Flags And USER_CHANNELOP&) <> _
                                 USER_CHANNELOP&) Then
                                
                                ' designate user
                                frmChat.AddQ "/designate " & _
                                    colUsersInChannel(i).Username
    
                                ' store designee name for future reference
                                gChannel.staticDesignee = _
                                    colUsersInChannel(i).Username
                                
                                ' we can only designate a
                                ' single person
                                Exit For
                            End If
                        End If
                    End With
                Next i
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' handle the display of user event
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (BotVars.ChatDelay) Then
        For i = 1 To colChatQueue.Count
            ' ...
            Set clsChatQueue = colChatQueue(i)
            
            If (StrComp(Username, clsChatQueue.Username, _
                vbBinaryCompare) = 0) Then
            
                Exit For
            End If
        Next i
    End If
    
    If ((BotVars.ChatDelay = 0) Or _
        ((colChatQueue.Count = 0) Or (i >= (colChatQueue.Count + 1)))) Then
        
        Call Event_QueuedStatusUpdate(Username, Flags, prevflags, Ping, Product, _
            Clan, Message, vbNullString)
    Else
        Set clsChatQueue = colChatQueue(i)
        
        Call clsChatQueue.StoreStatusUpdate(Flags, prevflags, Ping, Product, _
            Clan, Message, vbNullString)
    End If
    
    ' ...
    If (g_Channel.Self.IsOperator) Then
        ' we don't want anyone here that isn't
        ' supposed to be here.
        Call checkUsers
    End If
    
    ' destroy instance of chat queue
    Set clsChatQueue = Nothing
    
    Exit Sub
    
ERROR_HANDLER:
    MsgBox "Error: " & Err.description & " in Event_FlagsUpdate()"

    Exit Sub
End Sub

Public Sub Event_JoinedChannel(ByVal ChannelName As String, ByVal Flags As Long)
    Dim mailCount As Integer ' ...
    Dim ToANSI    As String  ' ...
    
    ' ...
    Set g_Channel = New clsChannelObj
    
    ' ...
    If (frmChat.mnuUTF8.Checked) Then
        ' ...
        ToANSI = UTF8Decode(ChannelName)
        
        ' ...
        If (Len(ToANSI) > 0) Then
            ChannelName = ToANSI
        End If
    End If
    
    ' ...
    With g_Channel
        .Name = ChannelName
        .Flags = Flags
        .JoinDate = Now
    End With

    ' clear chat queue when
    ' joining new channel
    Call ClearChatQueue
    
    ' we want to reset our filter
    ' values when we join a new channel
    BotVars.JoinWatch = 0
    
    ' if our channel is for some reason null, we don't
    ' want to continue, possibly causing further errors
    If (Len(ChannelName) < 1) Then
        Exit Sub
    End If
    
    With gChannel
        .Current = ChannelName
        .Flags = Flags
    End With
    
    SharedScriptSupport.MyChannel = ChannelName
    
    If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
        PassedClanMotdCheck = False
    End If

    ' if we've just left another channel, call event script
    ' function indicating that we've done so.
    If (LenB(g_Channel.Name)) Then
        On Error Resume Next
        
        frmChat.SControl.Run "Event_ChannelLeave"
    End If

    BanCount = 0
    
    Call frmChat.ClearChannel
    
    frmChat.AddChat RTBColors.JoinedChannelText, "-- Joined channel: ", _
        RTBColors.JoinedChannelName, ChannelName, RTBColors.JoinedChannelText, " --"
    
    SetTitle CurrentUsername & ", online in channel " & _
        g_Channel.Name
    
    ' have we just joined the void?
    If (g_Channel.IsSilent) Then
        ' lets inform user of potential lag issues while in this channel
        frmChat.AddChat RTBColors.InformationText, "If you experience a lot of lag while within " & _
                "this channel, try selecting 'Disable Silent Channel View' from the Window menu."
        
        ' if we've joined the void, lets try to grab the list of
        ' users within the channel by attempting to force a user
        ' update message using Battle.net's unignore command.
        If (frmChat.mnuDisableVoidView.Checked = False) Then
            ' ...
            frmChat.tmrSilentChannel(1).Enabled = True
        
            ' ...
            frmChat.AddQ "/unignore " & CurrentUsername
        End If
    Else
        ' ...
        frmChat.tmrSilentChannel(1).Enabled = False
    End If

    ' lets update our configuration file with the
    ' current channel name so that we join the channel
    ' again automatically if we disconnect or close the bot.
    Call WriteINI("Other", "LastChannel", ChannelName)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' check for mail
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    mailCount = GetMailCount(CurrentUsername)
        
    If (mailCount) Then
        frmChat.AddChat RTBColors.ConsoleText, "You have " & _
            mailCount & " new message" & IIf(mailCount = 1, "", "s") & _
                ". Type /getmail to retrieve."
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_ChannelJoin", ChannelName, Flags
End Sub

Public Sub Event_KeyReturn(ByVal KeyName As String, ByVal KeyValue As String)
    On Error Resume Next
    
    Dim s() As String
    Dim u   As String
    Dim i   As Integer

    ' Some of the oldest code in this project lives right here
    If SuppressProfileOutput Then
        
        ' // We're receiving profile information from a scripter request
        ' // No need to do anything at all with it except set Suppress = False after
        ' // the description comes in, and of course hadn it over to the scripters
        frmChat.SControl.Run "Event_KeyReturn", KeyName, KeyValue
        
        If KeyName = "Profile\Description" Then
            SuppressProfileOutput = False
        End If
    
    ElseIf ProfileRequest = True Then
    
        If KeyName = "Profile\Age" Then
            frmWriteProfile.txtAge.text = KeyValue
        ElseIf KeyName = "Profile\Location" Then
            frmWriteProfile.txtLoc.text = KeyValue
        ElseIf KeyName = "Profile\Description" Then
            frmWriteProfile.txtDescr.text = KeyValue
        ElseIf KeyName = "Profile\Sex" Then
            frmWriteProfile.txtSex.text = KeyValue
        End If
        
        frmWriteProfile.SetFocus
        
        frmChat.SControl.Run "Event_KeyReturn", KeyName, KeyValue
        
    ' Public Profile Listing
    ElseIf PPL = True Then
        
        If LenB(PPLRespondTo) > 0 Then
            u = "/w " & IIf(Dii, "*", "") & PPLRespondTo & " "
        Else
            u = ""
        End If
        
        If KeyName = "Profile\Location" Then
Repeat2:
            i = InStr(1, KeyValue, Chr(13))
            
            If Len(KeyValue) > 90 Then
                If i <> 0 Then
                    frmChat.AddQ u & "[Location] " & Left$(KeyValue, Len(KeyValue) - i)
                    KeyValue = Right(KeyValue, Len(KeyValue) - i)
                    
                    GoTo Repeat2
                Else
                    frmChat.AddQ u & "[Location] " & KeyValue
                End If
            Else
                If i <> 0 Then
                    frmChat.AddQ u & "[Location] " & Left$(KeyValue, Len(KeyValue) - i)
                    KeyValue = Right(KeyValue, Len(KeyValue) - i)
                    GoTo Repeat2
                Else
                    frmChat.AddQ u & "[Location] " & KeyValue
                End If
            End If
            
        ElseIf KeyName = "Profile\Description" Then
        
            Dim X() As String
            
            X() = Split(KeyValue, Chr(13))
            ReDim s(0)
            
            For i = LBound(X) To UBound(X)
                s(0) = X(i)
                
                If Len(s(0)) > 200 Then s(0) = Left$(s(0), 200)
                
                If i = LBound(X) Then
                    frmChat.AddQ u & "[Descr] " & s(0)
                Else
                    frmChat.AddQ u & "[Descr] " & Right(s(0), Len(s(0)) - 1)
                End If
            Next i
            
            PPL = False
            
            If LenB(PPLRespondTo) > 0 Then
                PPLRespondTo = ""
            End If
            
        ElseIf KeyName = "Profile\Sex" Then
Repeat4:
            If Len(KeyValue) > 90 Then
                frmChat.AddQ u & "[Sex] " & Left$(KeyValue, 80) & " [more]"
                KeyValue = Right(KeyValue, Len(KeyValue) - 80)
                GoTo Repeat4
            Else
                frmChat.AddQ u & "[Sex] " & KeyValue
            End If
            
        Else
            
        End If
        
    ElseIf Left$(KeyName, 7) = "System\" Then

        'frmchat.addchat RTBColors.ConsoleText, KeyName & ": " & KeyValue
        
        If InStr(1, KeyValue, " ", vbTextCompare) > 0 Then '// If it's a FILETIME
        
            Dim FT As FILETIME
            Dim sT As SYSTEMTIME
            
            FT.dwHighDateTime = CLng(Left$(KeyValue, InStr(1, KeyValue, " ", vbTextCompare)))
            
            On Error Resume Next
            
            KeyValue = Mid$(KillNull(KeyValue), InStr(1, KeyValue, " ", vbTextCompare) + 1)
            'keyvalue = Left$(keyvalue, Len(keyvalue) - 1)
            
            FT.dwLowDateTime = KeyValue 'CLng(KeyValue & "0")
            
            FileTimeToSystemTime FT, sT
            
            With sT
                frmChat.AddChat RTBColors.ServerInfoText, Right$(KeyName, Len(KeyName) - 7) & ": " & _
                        SystemTimeToString(sT) & " (Battle.net time)"
            End With
            
        Else    '// it's a SECONDS type
            If StrictIsNumeric(KeyValue) Then
                'On Error Resume Next
                frmChat.AddChat RTBColors.ServerInfoText, "Time Logged: " & ConvertTime(KeyValue, 1)
            End If
        End If
        
    Else
    
        frmProfile.Show
        
        If KeyName = "Profile\Age" Then
            frmProfile.txtAge.text = KeyValue
        ElseIf KeyName = "Profile\Location" Then
            frmProfile.txtLoc.text = KeyValue
        ElseIf KeyName = "Profile\Description" Then
            frmProfile.rtbProfile.text = vbNullString
            frmProfile.AddText vbWhite, KeyValue
        ElseIf KeyName = "Profile\Sex" Then
            frmProfile.txtSex.text = KeyValue
        End If
        
        frmProfile.SetFocus
        frmChat.SControl.Run "Event_KeyReturn", KeyName, KeyValue
        
    End If
End Sub

Public Sub Event_LoggedOnAs(Username As String, Product As String)
    LastWhisper = vbNullString

    'If InStr(1, Username, "*", vbBinaryCompare) <> 0 Then
    '    Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbBinaryCompare))
    'End If
    
    Call g_Queue.Clear
    
    g_Online = True
    
    DestroyNLSObject
    
    AttemptedFirstReconnect = False
    
    Call SetNagelStatus(frmChat.sckBNet.SocketHandle, True)
    
    Call EnableSO_KEEPALIVE(frmChat.sckBNet.SocketHandle)
    
    If (BotVars.UsingDirectFList) Then
        Call frmChat.FriendListHandler.RequestFriendsList(PBuffer)
    End If
    
    CurrentUsername = KillNull(Username)
    
    If (StrComp(Left$(CurrentUsername, 2), "w#", vbTextCompare) = 0) Then
        CurrentUsername = Mid(CurrentUsername, 3)
    End If

    SharedScriptSupport.myUsername = CurrentUsername
    
    With frmChat
        .InitListviewTabs
    
        .AddChat RTBColors.InformationText, "[BNET] Logged on as ", _
                RTBColors.SuccessText, Username, RTBColors.InformationText, "."
            
        .UpTimer.Interval = 1000
        
        .Timer.Interval = 30000
    
        'If (Not (DisableMonitor)) Then
        '    .AddChat RTBColors.SuccessText, "User monitor initialized."
        '
        '    InitMonitor
        'End If
    End With
    
    If (frmChat.sckBNLS.State <> 0) Then
        frmChat.sckBNLS.Close
    End If
    
    'INetQueue inqReset
    
    'FullJoin BotVars.HomeChannel

    QueueLoad = (QueueLoad + 2)
    
    Call frmChat.UpdateTrayTooltip
    
    If (ExReconnectTimerID > 0) Then
        Call KillTimer(0, ExReconnectTimerID)
        
        ExReconnectTimerID = 0
    End If
    
    RequestSystemKeys
    
    With PBuffer
        .InsertNTString "/whoami"
        .SendPacket &HE
    End With
    
    Call FullJoin(BotVars.HomeChannel)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error Resume Next
    
    frmChat.SControl.Run "Event_LoggedOn", Username, Product
End Sub

' updated 8-10-05 for new logging system
Public Sub Event_LogonEvent(ByVal Message As Byte, Optional ByVal ExtraInfo As String)
    Dim lColor       As Long
    Dim sMessage     As String
    Dim UseExtraInfo As Boolean

    Select Case (Message)
        Case 0:
            lColor = RTBColors.ErrorMessageText
            
            sMessage = "Login error - account does not exist."
            
        Case 1:
            lColor = RTBColors.ErrorMessageText
            
            sMessage = "Login error - invalid password."
            
        Case 2:
            lColor = RTBColors.SuccessText
            
            sMessage = "Login successful."
            
        Case 3:
            lColor = RTBColors.InformationText
            
            sMessage = "Attempting to create account..."
            
        Case 4:
            lColor = RTBColors.SuccessText
            
            sMessage = "Account created successfully."
            
        Case 5:
            sMessage = ExtraInfo
            
            lColor = RTBColors.ErrorMessageText
    End Select
    
    frmChat.AddChat lColor, "[BNET] " & sMessage
End Sub

Public Sub Event_RealmConnected()
    frmChat.AddChat RTBColors.SuccessText, "Realm: Connected! Please wait, " & _
        "logging in to the Diablo II realm may take a moment."
End Sub

Public Sub Event_RealmConnecting()
    frmChat.AddChat RTBColors.InformationText, "Realm: Connecting..."
End Sub

Public Sub Event_RealmError(ErrorNumber As Integer, description As String)
    frmChat.AddChat RTBColors.ErrorMessageText, "Realm: Error " & _
        ErrorNumber & ": " & description
End Sub

Public Sub Event_ServerError(ByVal Message As String)
    frmChat.AddChat RTBColors.ErrorMessageText, Message
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_ServerError", Message
End Sub

Public Sub Event_ServerInfo(ByVal Username As String, ByVal Message As String)
    Const MSG_BANNED      As String = " was banned by "
    Const MSG_UNBANNED    As String = " was unbanned by "
    Const MSG_SQUELCHED   As String = " has been squelched."
    Const MSG_UNSQUELCHED As String = " has been unsquelched."
    Const MSG_KICKEDOUT   As String = " kicked you out of the channel!"
    Const MSG_FRIENDS     As String = "Your friends are:"
    
    Dim i      As Integer
    Dim Temp   As String
    Dim bHide  As Boolean
    Dim ToANSI As String
    
    ' ...
    If (frmChat.mnuUTF8.Checked) Then
        ' ...
        ToANSI = UTF8Decode(Message)
        
        ' ...
        If (Len(ToANSI) > 0) Then
            Message = ToANSI
        End If
    End If
    
    ' ...
    If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
        ' ...
        If (PassedClanMotdCheck = False) Then
            ' ...
            If (Message <> vbNullString) Then
                Call frmChat.AddChat(RTBColors.ServerInfoText, Message)
            End If

            ' ...
            PassedClanMotdCheck = True
            
            ' ...
            Exit Sub
        End If
    End If
    
    ' ...
    If (g_request_receipt) Then ' for .cs and .cb commands
        ' ...
        Caching = True
    
        ' ...
        Cache Message, 1
        
        ' ...
        With frmChat.quLower
            .Enabled = False
            .Enabled = True
        End With
    End If
    
    ' what is our current gateway name?
    If (BotVars.Gateway = vbNullString) Then
        ' ...
        If (InStr(1, Message, "You are ", vbTextCompare) > 0) And (InStr(1, Message, ", using ", _
                vbTextCompare) > 0) Then
                
            ' ...
            If ((InStr(1, Message, "channel", vbTextCompare) = 0) And _
                    (InStr(1, Message, "game", vbTextCompare) = 0)) Then
                    
                ' ...
                i = InStrRev(Message, Space$(1))
                
                ' ...
                BotVars.Gateway = Mid$(Message, i + 1)
    
                ' we want our username to accurately reflect
                ' our new discovery of the realm name
                CurrentUsername = convertUsername(CurrentUsername)
    
                Exit Sub
            End If
        End If
    End If

    If (InStr(1, Message, Space$(1), vbBinaryCompare) <> 0) Then
        If (InStr(1, Message, "are still marked", vbTextCompare) <> 0) Then
            Exit Sub
        End If
        
        If ((InStr(1, Message, " from your friends list.", vbBinaryCompare) > 0) Or _
            (InStr(1, Message, " to your friends list.", vbBinaryCompare) > 0)) Then
            
            frmChat.lvFriendList.ListItems.Clear
            
            Call frmChat.FriendListHandler.RequestFriendsList(PBuffer)
            
            frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString
            
            unsquelching = True
        End If
        
        'Ban Evasion and banned-user tracking
        Temp = Split(Message, " ")(1)
        
        ' added 1/21/06 thanks to
        ' http://www.stealthbot.net/forum/index.php?showtopic=24582
        
        If (Len(Temp) > 0) Then
            Dim Banning    As Boolean
            Dim Unbanning  As Boolean
            Dim user       As String  ' ...
            Dim cOperator  As String  ' ...
            Dim msgPos     As Integer ' ...
            Dim Pos        As Integer ' ...
            Dim tmp        As String
            Dim banpos     As Integer ' ...
            Dim j          As Integer
            
            If (InStr(1, Message, MSG_BANNED, vbTextCompare) > 0) Then
                ' ...
                user = Left$(Message, _
                                (InStr(1, Message, MSG_BANNED, vbBinaryCompare) - 1))
                
                ' ...
                If (Len(user) > 0) Then
                    ' ...
                    g_Channel.TotalBanCount = _
                                (g_Channel.TotalBanCount + 1)
                
                    ' ...
                    Pos = g_Channel.GetUserIndexByName(CleanDiablo2Username(Username))
                    
                    ' ...
                    If (Pos > 0) Then
                        Dim BanlistObj As clsBannedUserObj
                        
                        ' ...
                        If (g_Channel.Users(Pos).IsOnBanList(user) = False) Then
                            ' ...
                            Set BanlistObj = New clsBannedUserObj
                            
                            ' ...
                            With BanlistObj
                                .Name = user
                                .DateOfBan = Now
                            End With
                            
                            ' ...
                            Call g_Channel.Users(Pos).Banlist.Add(BanlistObj)
                        End If
                    End If
                End If
                
                ' ...
                Call RemoveBanFromQueue(user)
            ElseIf (InStr(1, Message, MSG_UNBANNED, vbTextCompare) > 0) Then
                ' ...
                user = Left$(Message, _
                                (InStr(1, Message, MSG_UNBANNED, vbBinaryCompare) - 1))
                                
                ' ...
                If (Len(user) > 0) Then
                    ' ...
                    For i = 1 To g_Channel.Users.Count
                        ' ...
                        If (g_Channel.Users(i).IsOperator) Then
                            ' ...
                            banpos = g_Channel.Users(i).IsOnBanList(user)
                            
                            ' ...
                            If (banpos > 0) Then
                                Call g_Channel.Users(i).Banlist.Remove(banpos)
                            End If
                        End If
                    Next i
                End If
            End If
    
            '// backup channel
            If (InStr(Len(Temp), Message, "kicked you out", vbTextCompare) > 0) Then
                If ((StrComp(g_Channel.Name, "Op [vL]", vbTextCompare) <> 0) And _
                    (StrComp(g_Channel.Name, "Op Fatal-Error", vbTextCompare) <> 0)) Then
                        
                    If (BotVars.UseBackupChan) Then
                        If (Len(BotVars.BackupChan) > 1) Then
                            frmChat.AddQ "/join " & BotVars.BackupChan
                        End If
                    Else
                        frmChat.AddQ "/join " & g_Channel.Name
                    End If
                End If
            End If
            
            ' ...
            If (InStr(Len(Temp), Message, " has been unsquelched", vbTextCompare) > 0) Then
                unsquelching = True
            End If
        End If
        
        ' ...
        If (InStr(1, Message, "designated heir", vbTextCompare) <> 0) Then
            gChannel.Designated = Left$(Message, Len(Message) - 29)
        End If
        
        
        Temp = "Your friends are:"
        
        If (StrComp(Left$(Message, Len(Temp)), Temp) = 0) Then
            If (Not (BotVars.ShowOfflineFriends)) Then
                Message = Message & _
                    "  ÿci(StealthBot is hiding your offline friends)"
            End If
        End If
    
    End If ' message contains a space
    
    If (StrComp(Right$(Message, 9), ", offline", vbTextCompare) = 0) Then
        If (BotVars.ShowOfflineFriends) Then
            frmChat.AddChat RTBColors.ServerInfoText, Message
        End If
    Else
        'If (Not (bHide)) Then
            frmChat.AddChat RTBColors.ServerInfoText, Message
        'End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_ServerInfo", Message
End Sub

Public Sub Event_SomethingUnknown(ByVal UnknownString As String)
    frmChat.AddChat RTBColors.ErrorMessageText, "Something unknown has happened... " & _
        "Did Battle.Net change something? The Unknown Event is as follows:"
    frmChat.AddChat RTBColors.ErrorMessageText, "[" & UnknownString & "]"
    frmChat.AddChat RTBColors.ErrorMessageText, "Please report this event to Stealth as soon " & _
        "as possible, copy/paste this entire message."
End Sub

Public Sub Event_UserEmote(ByVal Username As String, ByVal Flags As Long, ByVal Message As String)
    Dim i      As Integer ' ...
    Dim ToANSI As String ' ...

    ' ...
    Username = convertUsername(Username)
    
    ' ...
    If (frmChat.mnuUTF8.Checked) Then
        ' ...
        ToANSI = UTF8Decode(Message)
        
        ' ...
        If (Len(ToANSI) > 0) Then
            Message = ToANSI
        End If
    End If
    
    i = UsernameToIndex(Username)
    
    If (i > 0) Then
        colUsersInChannel.Item(i).Acts
    End If
    
    If (Catch(0) <> vbNullString) Then
        Call CheckPhrase(Username, Message, CPEMOTE)
    End If
    
    If ((Phrasebans) Or (BotVars.QuietTime)) Then
        If ((g_Channel.Self.IsOperator) And _
            ((Flags And USER_CHANNELOP&) <> USER_CHANNELOP&)) Then
            
            If (BotVars.QuietTime) Then
                If (Not (GetSafelist(Username))) Then
                    frmChat.AddQ "/ban " & Username & " Quiet-time is enabled."
                End If
            End If
            
            If (Phrasebans) Then
                For i = LBound(Phrases) To UBound(Phrases)
                    If ((LCase$(Phrases(i)) <> vbNullString) And _
                        (LCase$(Phrases(i)) <> Space$(1))) Then
                    
                        If (InStr(1, Message, Phrases(i), vbTextCompare) <> 0) Then
                            If (Not (GetSafelist(Username))) Then
                                frmChat.AddQ "/ban " & Username & " Banned phrase: " & _
                                    Phrases(i), 1
                            End If
                            
                            GoTo theEnd
                        End If
                    End If
                Next i
            End If
        End If
    End If
    
    If (Len(Message) >= 100) Then
        BotVars.JoinWatch = (BotVars.JoinWatch + 5)
    End If
    
    If (BotVars.JoinWatch >= 20) Then
        If (Filters = False) Then
            frmChat.AddChat RTBColors.TalkBotUsername, _
                "Chat filters have been activated due to excessive rejoins and/or " & _
                    "spam; deactivate them by pressing CTRL + F."
    
            Call WriteINI("Other", "Filters", "Y")
                    
            Filters = True
            
            AutoChatFilter = GetTickCount()
        End If
            
        BotVars.JoinWatch = 0
        
        If (AutoChatFilter) Then
            AutoChatFilter = GetTickCount()
        End If
    End If

theEnd:
        ' ...
        If (BotVars.ChatDelay) Then
            For i = 1 To colChatQueue.Count
                ' ...
                Dim clsChatQueue As clsChatQueue
            
                ' ...
                Set clsChatQueue = colChatQueue(i)
                
                ' ...
                If (StrComp(Username, clsChatQueue.Username, _
                    vbBinaryCompare) = 0) Then
                
                    Exit For
                End If
            Next i
        End If
        
        ' ...
        If ((BotVars.ChatDelay = 0) Or _
            ((colChatQueue.Count = 0) Or (i >= (colChatQueue.Count + 1)))) Then
            
            ' ...
            Call Event_QueuedEmote(Username, Flags, 0, Message)
        Else
            ' ...
            Call clsChatQueue.StoreEmote(Flags, 0, Message)
            
            ' ...
            Set clsChatQueue = Nothing
        End If
    
End Sub

'Ping, Product, sClan, InitStatstring, W3Icon
Public Sub Event_UserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, Optional ByVal w3icon As String)

    On Error GoTo ERROR_HANDLER

    Dim clsChatQueue As clsChatQueue ' ...
    Dim UserObj      As clsChannelUserObj
    
    Dim UserIndex    As Integer ' ...
    Dim i            As Integer ' ...
    Dim strCompare   As String  ' ...
    Dim Level        As Byte    ' ...
    Dim StatUpdate   As Boolean ' ...
    Dim index        As Long    ' ...
    
    If (Len(Username) < 1) Then
        Exit Sub
    End If
    
    ' Error correction code added April 2005
    ' to fix a mysterious ghosting bug
    If (Ping >= 200000000) Then
        Exit Sub
    End If
    
    ' ...
    UserIndex = g_Channel.GetUserIndexByName(Username)
    
    ' ...
    If (UserIndex > 0) Then
        ' ...
        Set UserObj = g_Channel.Users(UserIndex)
    Else
        ' ...
        Set UserObj = New clsChannelUserObj
    End If
    
    ' ...
    With UserObj
        ' ...
        .Name = CleanDiablo2Username(Username)
        .DisplayName = convertUsername(Username)
        .Flags = Flags
        .Ping = Ping
        .JoinDate = g_Channel.JoinDate
    End With
    
    ' ...
    If (UserIndex = 0) Then
        g_Channel.Users.Add UserObj
    End If
    
    ' ...
    Username = UserObj.DisplayName
    
    ' are we receiving my user information?
    If (StrComp(Username, CurrentUsername, vbBinaryCompare) = 0) Then

        ' we don't want to have an out-of-date
        ' flag value for ourselves
        MyFlags = Flags
        
        ' we don't want to have an out-of-date
        ' flag value for ourselves in scripts,
        ' either
        SharedScriptSupport.BotFlags = MyFlags
    End If

    ' ...
    StatUpdate = (checkChannel(Username))

    ' ...
    If (StatUpdate = False) Then
        ' ...
        If (BotVars.ChatDelay) Then
            ' ...
            For i = 1 To colChatQueue.Count
                ' ...
                Set clsChatQueue = colChatQueue(i)
                
                ' ...
                If (StrComp(Username, clsChatQueue.Username, _
                    vbBinaryCompare) = 0) Then
                
                    Exit For
                End If
            Next i
            
            If (i < (colChatQueue.Count + 1)) Then
                StatUpdate = True
            End If
        End If
    Else
        ' if we found the user in the channel then we can assume that
        ' we won't find him again in the incoming chat queue
        i = (colChatQueue.Count + 1)
    End If
    
    ' ...
    If (StatUpdate) Then
        ' ...
        UserIndex = UsernameToIndex(Username)
    
        ' ...
        If (UserIndex) Then
            With colUsersInChannel(UserIndex)
                .Username = Username
                .Flags = Flags
                .Ping = Ping
                .Clan = sClan
                .Statstring = OriginalStatstring
            End With
        End If
    
        ' ...
        If ((BotVars.ChatDelay = 0) Or _
            ((colChatQueue.Count = 0) Or (i >= (colChatQueue.Count + 1)))) Then

            Call Event_QueuedUserInChannel(Username, Flags, Ping, Product, sClan, _
                OriginalStatstring, w3icon)
        Else
            Call clsChatQueue.StoreUserInChannel(Flags, Ping, Product, sClan, _
                OriginalStatstring, w3icon)
        End If
        
        ' ...
        Set clsChatQueue = Nothing
    Else
        Dim UserToAdd As clsUserInfo ' ...
        
        ' create new instance of class
        Set UserToAdd = New clsUserInfo
    
        ' dunno what this does...
        If ((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
            If (StrComp(Username, CurrentUsername, _
                vbTextCompare) <> 0) Then
                
                gChannel.Designated = Username
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add user to collection
        ' *necessary*
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        With UserToAdd
            .Username = Username
            .Flags = Flags
            .Ping = Ping
            .Product = Product
            .Safelisted = GetSafelist(Username)
            .Statstring = OriginalStatstring
            .JoinTime = GetTickCount
            .Clan = sClan
            .IsSelf = (StrComp(Username, CurrentUsername, _
                vbBinaryCompare) = 0)
        
            ' if the user isn't safelisted, lets make sure he's abides by the
            ' channel rules and, if required, inputs the correct channel password.
            If (Not (.Safelisted)) Then
                'BotVars.Gateway  the user an operator?  ... or his "he" actually "me"?
                If (((Flags And USER_CHANNELOP&) <> USER_CHANNELOP&) And _
                     (StrComp(Username, CurrentUsername, vbBinaryCompare) <> 0)) Then
                    
                    ' do we have a channel password set?
                    If ((Len(BotVars.ChannelPassword) > 0) And _
                        (BotVars.ChannelPasswordDelay > 0)) Then
                        
                        .InternalFlags = (.InternalFlags + IF_AWAITING_CHPW)
                    End If
                    
                    ' do we have idle bans on?
                    If (BotVars.IB_On = 1) Then
                        .InternalFlags = (.InternalFlags + _
                            IF_SUBJECT_TO_IDLEBANS)
                    End If
                End If
            End If
        End With
        
        ' lets add our new friend to the collection
        Call colUsersInChannel.Add(UserToAdd)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add to the channel list
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        If (InStr(1, Message, "in clan ", vbTextCompare) > 0) Then
            strCompare = Mid$(Message, InStr(1, Message, "in clan ", vbTextCompare) + 8)
            strCompare = Left$(strCompare, Len(strCompare) - 1)
        
            Call AddName(Username, Product, Flags, Ping, strCompare)
        Else
            Call AddName(Username, Product, Flags, Ping)
        End If
        
        Call DoLastSeen(Username)
        
        frmChat.lblCurrentChannel.Caption = _
                frmChat.GetChannelString()
        
        ' destroy class
        Set UserToAdd = Nothing
    End If
    
    If (MDebug("statstrings")) Then
        frmChat.AddChat vbMagenta, "Username: " & Username & ", Statstring: " & _
            OriginalStatstring
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_UserInChannel", Username, Flags, Message, Ping, _
        Product, StatUpdate
        
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in Event_UserInChannel().")
    
    Exit Sub
End Sub

Public Sub Event_UserJoins(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    On Error GoTo ERROR_HANDLER
    
    Dim UserObj     As clsChannelUserObj
    Dim UserToAdd   As clsUserInfo
    
    Dim toCheck     As String
    Dim strCompare  As String
    Dim i           As Long
    Dim Temp        As Byte
    Dim Level       As Byte
    Dim l           As Long
    Dim Banned      As Boolean
    Dim f           As Integer
    Dim UserIndex   As Integer ' ...
    Dim BanningUser As Boolean ' ...
    
    If (Len(Username) < 1) Then
        Exit Sub
    End If
    
    ' ...
    UserIndex = g_Channel.GetUserIndexByName(Username)
    
    ' ...
    If (UserIndex = 0) Then
        ' ...
        Set UserObj = New clsChannelUserObj
        
        ' ...
        With UserObj
            .Name = CleanDiablo2Username(Username)
            .DisplayName = convertUsername(Username)
            .Flags = Flags
            .Ping = Ping
            .game = Product
            .JoinDate = Now
        End With
        
        ' ...
        g_Channel.Users.Add UserObj
    Else
        frmChat.AddChat vbRed, "Error: We have received a join event for a user that we had thought was " & _
                "already present within the channel."
        
        Exit Sub
    End If
    
    ' ...
    Username = UserObj.DisplayName
     
    Set UserToAdd = New clsUserInfo
     
    Banned = True
    
    f = FreeFile
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Add user to collection
    ' *necessary*
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With UserToAdd
        .Flags = Flags
        .Username = Username
        .Ping = Ping
        .Product = Product
        .Safelisted = GetSafelist(Username)
        .Statstring = OriginalStatstring
        .JoinTime = GetTickCount
        .Clan = sClan
        .IsSelf = (StrComp(Username, CurrentUsername, _
            vbBinaryCompare) = 0)
        .InternalFlags = 0
        
        If (Not (.Safelisted)) Then
            If (((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) And _
                 (StrComp(Username, CurrentUsername, vbTextCompare) <> 0)) Then
                
                If (BotVars.IB_On = 1) Then
                    .InternalFlags = (.InternalFlags + _
                        IF_SUBJECT_TO_IDLEBANS)
                End If
            End If
            
            If ((Len(BotVars.ChannelPassword) > 0) And _
                (BotVars.ChannelPasswordDelay > 0)) Then
                
                If (g_Channel.Self.IsOperator) Then
                    If ((Len(BotVars.ChannelPassword) > 0) And _
                        (BotVars.ChannelPasswordDelay > 0)) Then
                        
                        .InternalFlags = (.InternalFlags + _
                            IF_AWAITING_CHPW)
                        
                        frmChat.AddQ "/w " & Username & " You have " & _
                            BotVars.ChannelPasswordDelay & " seconds to whisper a valid " & _
                                "password or you will be banned."
                    End If
                End If
            End If
        End If
    End With
    
    ' ...
    Call colUsersInChannel.Add(UserToAdd)
    
                
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' What level are they?
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ((Product = "WAR3") Or _
            (Product = "W3XP")) Then
        
        i = InStr(1, Message, "Level: ", vbTextCompare)
    ElseIf Product = "D2DV" Or Product = "D2XP" Then
        i = InStr(1, Message, " level ", vbTextCompare)
    End If
    
    If (i > 0) Then
        strCompare = Mid(Message, i + 7, 2)
        
        If (Right$(strCompare, 1) = ",") Then
            strCompare = Left$(strCompare, 1)
        End If
        
        Level = CByte(Val(strCompare))
    Else
        Level = 0
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Flash window
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (frmChat.mnuFlash.Checked) Then
        Call FlashWindow
    End If
    
    If (BotVars.ChatDelay) Then
        Dim clsChatQueue As clsChatQueue
    
        Set clsChatQueue = New clsChatQueue
    
        With clsChatQueue
            .Username = Username
            .Time = GetTickCount()
        End With
        
        Call clsChatQueue.StoreJoin(Flags, Ping, Product, sClan, _
            OriginalStatstring, w3icon)
        
        Call colChatQueue.Add(clsChatQueue)
    Else
        Call Event_QueuedJoin(Username, Flags, Ping, Product, sClan, _
            OriginalStatstring, w3icon)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' AUTOMATIC MODERATION FEATURES
    '  These are all dependent on OPS (determined here)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (g_Channel.Self.IsOperator) Then
        ' There's no sense trying to perform moderatory actions on
        ' a moderator
        If ((Flags And USER_CHANNELOP&) <> USER_CHANNELOP&) Then
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Designate them?
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (InStr(1, GetCumulativeAccess(Username).Flags, "D", _
                vbBinaryCompare) > 0) Then
                
                If (gChannel.Designated = vbNullString) Then
                    If Mid$(LCase$(g_Channel.Name), 1, 3) = "op " Then
                        If (StrComp(Mid$(g_Channel.Name, 4), StripRealm(Username), _
                            vbTextCompare)) <> 0 Then
                            
                            frmChat.AddQ "/designate " & Username
    
                            gChannel.staticDesignee = Username
                        End If
                    Else
                        frmChat.AddQ "/designate " & Username
    
                        gChannel.staticDesignee = Username
                    End If
                End If
            End If
    
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' AUTOMATIC MODERATION FEATURES
            '  These are all dependent on OPS (above if control) and the user's
            '  SAFELISTED STATUS (determined here)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (Not (UserToAdd.Safelisted)) Then
    
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Warcraft III players: various checks
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If ((Product = "WAR3") Or (Product = "W3XP")) Then
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' Should they be banned for being low?
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (BotVars.BanUnderLevel > 0) Then
                        If (Level < BotVars.BanUnderLevel) Then
                            strCompare = ReadCFG("Other", "LevelbanMsg")
                            
                            If (strCompare <> vbNullString) Then
                                Ban Username & Space(1) & strCompare, _
                                    (AutoModSafelistValue - 1)
                            Else
                                Ban Username & Space(1) & _
                                    "You are under the required level for entry.", _
                                        (AutoModSafelistValue - 1)
                            End If
                            
                            GoTo theEnd
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' How about peon-banned?
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (BotVars.BanPeons = 1) Then
                        If (InStr(1, Message, "peon icon", vbTextCompare) > 0) Then
                            If (Len(ReadCFG("Main", "PeonBanMsg")) > 0) Then
                                Ban Username & Space(1) & ReadCFG("Main", "PeonBanMsg"), _
                                    (AutoModSafelistValue - 1)
                            Else
                                Ban Username & " PeonBan", (AutoModSafelistValue - 1)
                            End If
                            
                            GoTo theEnd
                        End If
                    End If
                End If ' (end Warcraft III checks)
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Diablo II players: are they levelbanned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If ((Product = "D2DV") Or (Product = "D2XP")) Then
                    If (BotVars.BanD2UnderLevel > 0) Then
                        If (Level < BotVars.BanD2UnderLevel) Then
                            strCompare = ReadCFG("Other", "LevelbanMsg")
                            
                            If (strCompare <> vbNullString) Then
                                Ban Username & Space(1) & strCompare, _
                                    (AutoModSafelistValue - 1)
                            Else
                                Ban Username & Space(1) & _
                                    "You are under the required level for entry.", _
                                        (AutoModSafelistValue - 1)
                            End If
                            GoTo theEnd
                        End If
                    End If
                End If
NoLevel:
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they plugbanned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If (BotVars.PlugBan) Then
                    If ((Flags And USER_NOUDP) = USER_NOUDP) Then
                        
                        Ban Username & " PlugBan", _
                            (AutoModSafelistValue - 1)
                        
                        GoTo theEnd
                    End If
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they evading a ban?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If (InStr(1, Username, "#", vbBinaryCompare) <> 0) Then
                    toCheck = LCase(Left$(Username, InStr(1, _
                        Username, "#", vbTextCompare) - 1))
                Else
                    toCheck = LCase$(Username)
                End If
                
                toCheck = StripRealm(toCheck)
                
                If (BotVars.BanEvasion) Then
                    For i = 0 To UBound(gBans)
                        If (StrComp(toCheck, gBans(i).Username, vbTextCompare) = 0) Then
                            Ban Username & " Ban Evasion", _
                                (AutoModSafelistValue - 1)
                            
                            GoTo theEnd
                        End If
                    Next i
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they shitlisted?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                toCheck = GetShitlist(Username)
                
                If (Len(toCheck) > 1) Then
                    Call frmChat.AddQ("/ban " & toCheck)
                    
                    GoTo theEnd
                End If
SLSkip:
                'Close #f
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Is the channel in lockdown?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If (Protect) Then
                    Ban Username & Space(1) & ProtectMsg, (AutoModSafelistValue - 1)
                    
                    GoTo theEnd
                End If
checkIPBan:
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they IP-banned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If (BotVars.IPBans) Then
                    If ((Flags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
                        Ban Username & " IPBanned.", (AutoModSafelistValue - 1)
                        
                        GoTo theEnd
                    End If
                End If
            End If ' (user not safelisted)
        End If ' (bot has ops)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If we've gotten this far, and haven't been banned yet, they're
    '   eligible for a greet message!
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Banned = False
    
    If (BotVars.UseGreet) Then
        If (LenB(BotVars.GreetMsg) > 0) Then
            If (StrComp(g_Channel.Name, "Clan SBs", vbTextCompare) <> 0) Then
                
                If (QueueLoad = 0) Then
                    QueueLoad = (QueueLoad + 1)
                End If
                
                If (BotVars.WhisperGreet) Then
                    frmChat.AddQ "/w " & Username & Space$(1) & _
                        DoReplacements(BotVars.GreetMsg, Username, Ping)
                Else
                    frmChat.AddQ DoReplacements(BotVars.GreetMsg, Username, Ping)
                End If
            End If
        End If
    End If
    
    
theEnd:
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Is channel flooding or mass-joining happening?
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    BotVars.JoinWatch = (BotVars.JoinWatch + 2)
    
    If (BotVars.JoinWatch >= 20) Then
        'If (Not (JoinMessagesOff)) Then
        '    If (ForcedJoinsOn = 0) Then
        '        frmChat.AddChat RTBColors.TalkBotUsername, _
        '            "Rejoin flooding and/or massloading detected!"
        '
        '        frmChat.AddChat RTBColors.TalkBotUsername, _
        '            "Join/Leave Messages have been disabled due to rejoin flooding. " & _
        '                "Reactivate them by pressing CTRL + J."
        '
        '        JoinMessagesOff = True
        '
        '        ForcedJoinsOn = 2
        '    End If
        'End If
        
        If (Filters = False) Then
            frmChat.AddChat RTBColors.TalkBotUsername, _
                "Chat filters have been activated due to excessive rejoins and/or " & _
                    "spam; deactivate them by pressing CTRL + F."
    
            Call WriteINI("Other", "Filters", "Y")
                    
            Filters = True
            
            AutoChatFilter = GetTickCount()
        End If
        
        BotVars.JoinWatch = 0
        
        If (AutoChatFilter) Then
            AutoChatFilter = GetTickCount()
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do they have mail?
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Mail) Then
        l = GetMailCount(Username)
        
        If (l > 0) Then
            frmChat.AddQ "/w " & Username & " You have " & l & _
                " new message" & IIf(l = 1, "", "s") & ". Type !inbox to retrieve."
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Print their statstring, if desired
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (MDebug("statstrings")) Then
        frmChat.AddChat RTBColors.ErrorMessageText, OriginalStatstring
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Add them to the LastSeen list
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call DoLastSeen(Username)
    
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in Event_UserJoins().")
    
    Exit Sub
End Sub

Public Sub Event_UserLeaves(ByVal Username As String, ByVal Flags As Long)
    On Error GoTo ERROR_HANDLER

    Dim UserObj   As clsChannelUserObj
    
    Dim UserIndex As Integer
    Dim i         As Integer
    Dim ii        As Integer
    Dim Holder()  As Variant
    Dim Pos       As Integer
    Dim bln       As Boolean
    
    ' ...
    UserIndex = _
        g_Channel.GetUserIndexByName(CleanDiablo2Username(Username))
    
    ' ...
    If (UserIndex > 0) Then
        ' ...
        g_Channel.Users.Remove UserIndex
    Else
        frmChat.AddChat vbRed, "Error: We have received a leave event for a user that we didn't know " & _
                "was in the channel."
    
        Exit Sub
    End If
    
    ' ...
    Username = convertUsername(Username)
    
    ' ...
    i = UsernameToIndex(Username)
    
    ' ...
    If (i) Then
        Call colUsersInChannel.Remove(i)
    End If
    
    Do Until (bln = True)
        ' ...
        For i = 0 To UBound(gBans)
            If (StrComp(gBans(i).cOperator, reverseUsername(Username), vbTextCompare) = 0) Then
                Call UnbanBannedUser(gBans(i).UsernameActual, reverseUsername(Username))
                
                Exit For
            End If
        Next i
        
        If (i >= UBound(gBans) + 1) Then
            bln = True
        End If
    Loop
    
    ' ...
    If (StrComp(Username, gChannel.Designated, vbTextCompare) = 0) Then
        ' ...
        gChannel.Designated = vbNullString
        
        ' ...
        For i = 1 To colUsersInChannel.Count
            With GetAccess(colUsersInChannel.Item(i).Username)
                ' ...
                If (InStr(1, .Flags, "D", vbBinaryCompare) > 0) Then
                    ' ...
                    If ((colUsersInChannel.Item(i).Flags And USER_CHANNELOP&) = _
                         USER_CHANNELOP&) Then
                        
                        ' ...
                        frmChat.AddQ "/designate " & _
                            colUsersInChannel.Item(i).Username

                        ' ...
                        gChannel.staticDesignee = colUsersInChannel.Item(i).Username
                        
                        ' ...
                        Exit For
                    End If
                End If
            End With
        Next i
    End If
    
    ' ...
    Call RemoveBanFromQueue(Username)
    
    ' ...
    If ((JoinMessagesOff = False) And (bFlood = False)) Then
        If (BotVars.ChatDelay) Then
            ' ...
            Dim clsChatQueue As clsChatQueue
        
            ' ...
            For i = 1 To colChatQueue.Count
                ' ...
                Set clsChatQueue = colChatQueue(i)
                
                ' ...
                If (StrComp(Username, clsChatQueue.Username, _
                    vbBinaryCompare) = 0) Then
                
                    Exit For
                End If
            Next i
        End If
        
        If ((BotVars.ChatDelay = 0) Or _
            ((colChatQueue.Count = 0) Or (i >= (colChatQueue.Count + 1)))) Then
            
            frmChat.AddChat RTBColors.JoinText, "-- ", RTBColors.JoinUsername, Username, _
                RTBColors.JoinText, " has left the channel."
        Else
            Call colChatQueue.Remove(i)
        End If
    End If
    
    ' ...
    UserIndex = checkChannel(Username)
    
    ' ...
    If (UserIndex) Then
        ' ...
        If (frmChat.mnuFlash.Checked) Then
            Call FlashWindow
        End If
    
        ' ...
        With frmChat.lvChannel
            .ListItems.Remove UserIndex

            .Refresh
        End With
        
        ' ...
        frmChat.lblCurrentChannel.Caption = _
            frmChat.GetChannelString()
    End If
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_UserLeaves", Username, Flags
    
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in Event_UserLeaves().")
    
    Exit Sub
End Sub

Public Sub Event_UserTalk(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, _
    ByVal Ping As Long)
    
    Dim strSend     As String
    Dim s           As String
    Dim u           As String
    Dim strCompare  As String
    Dim i           As Integer
    Dim ColIndex    As Integer
    Dim b           As Boolean
    Dim ToANSI      As String
    Dim BanningUser As Boolean
    
    ' ...
    Username = convertUsername(Username)
    
    ' ...
    If (frmChat.mnuUTF8.Checked) Then
        ' ...
        ToANSI = UTF8Decode(Message)
        
        ' ...
        If (Len(ToANSI) > 0) Then
            Message = ToANSI
        End If
    End If

    i = UsernameToIndex(Username)
    
    If (i > 0) Then
        colUsersInChannel.Item(i).Acts
    End If
        
    If (VoteDuration > 0) Then
        If (InStr(1, LCase(Message), "yes", vbTextCompare) > 0) Then
            Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
        ElseIf (InStr(1, LCase(Message), "no", vbTextCompare) > 0) Then
            Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
        End If
    End If
            
    If (Len(Message) >= 100) Then
        BotVars.JoinWatch = (BotVars.JoinWatch + 5)
    End If
    
    If (BotVars.JoinWatch >= 20) Then
        If (Filters = False) Then
            frmChat.AddChat RTBColors.TalkBotUsername, _
                "Chat filters have been activated due to excessive rejoins and/or " & _
                    "spam; deactivate them by pressing CTRL + F."
    
            Call WriteINI("Other", "Filters", "Y")

            Filters = True
            
            AutoChatFilter = GetTickCount()
        End If
            
        BotVars.JoinWatch = 0
        
        If (AutoChatFilter) Then
            AutoChatFilter = GetTickCount()
        End If
    End If
    
    b = False
    
    If (frmChat.mnuFlash.Checked) Then
        Call FlashWindow
    End If
    
    If (BotVars.ChatDelay) Then
        For i = 1 To colChatQueue.Count
            ' ...
            Dim clsChatQueue As clsChatQueue
        
            ' ...
            Set clsChatQueue = colChatQueue(i)
            
            If (StrComp(Username, clsChatQueue.Username, vbBinaryCompare) = 0) Then
                Exit For
            End If
        Next i
    End If
    
    If ((BotVars.ChatDelay = 0) Or _
        ((colChatQueue.Count = 0) Or (i >= (colChatQueue.Count + 1)))) Then
       
        Call Event_QueuedTalk(Username, Flags, Ping, Message)
    Else
        Set clsChatQueue = colChatQueue(i)
        
        Call clsChatQueue.StoreTalk(Flags, Ping, Message)
    End If
        
    If (g_Channel.Self.IsOperator) Then
        If (GetSafelist(Username) = False) Then
            If (Phrasebans) Then
                For i = LBound(Phrases) To UBound(Phrases)
                    If ((Phrases(i) <> vbNullString) And _
                            (Phrases(i) <> Space$(1))) Then
                        
                        If ((InStr(1, Message, Phrases(i), vbTextCompare)) <> 0) Then
                            Ban Username & " Banned phrase: " & Phrases(i), _
                                    (AutoModSafelistValue - 1)
                            
                            BanningUser = True
                            
                            Exit For
                        End If
                    End If
                Next i
            End If
            
            If (BanningUser = False) Then
                If (BotVars.QuietTime) Then
                    Ban Username & " Quiet-time is enabled.", (AutoModSafelistValue - 1)
                ElseIf (BotVars.KickOnYell = 1) Then
                    If (Len(Message) > 5) Then
                        If (PercentActualUppercase(Message) > 90) Then
                            Ban Username & " Yelling", (AutoModSafelistValue - 1), 1
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If (Mail) Then
        If (StrComp(Left$(Message, 6), "!inbox", vbTextCompare) = 0) Then
            Dim Msg As udtMail ' ...
            
            If (GetMailCount(Username) > 0) Then
                Call GetMailMessage(Username, Msg)
                
                If (Len(RTrim(Msg.To)) > 0) Then
                    frmChat.AddQ "/w " & Username & " Message from " & RTrim$(Msg.From) & _
                            ": " & RTrim$(Msg.Message)
                End If
            End If
        End If
    End If
    
    Call ProcessCommand(Username, Message, False, False)
End Sub

Public Sub Event_VersionCheck(Message As Long, ExtraInfo As String)
    Dim l As Long

    Select Case (Message)
        Case 0:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Client version accepted!"
        
        Case 1:
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Version check failed! " & _
                "The version byte for this attempt was 0x" & Hex(GetVerByte(BotVars.Product)) & "."

            If (BotVars.BNLS) Then
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] BNLS has not been updated yet, " & _
                    "or you experienced an error. Try connecting again."
            Else
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Please ensure you " & _
                    "have updated your hash files using more current ones from the directory " & _
                        "of the game you're connecting with."
                
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] In addition, you can try " & _
                    "choosing ""Update version bytes from StealthBot.net"" from the Bot menu."
                
                Message = 0
            End If
        
        Case 2:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is invalid!"
        
        Case 3:
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Version check failed! " & _
                "BNLS has not been updated yet.. Try reconnecting in an hour or two."
        
        Case 4:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is for another game."
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit " & _
                "http://www.blizzard.com/support/?id=awr0639p ."
        
        Case 5:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is banned. " & _
                "For more information, visit http://www.blizzard.com/support/?id=asc0638p ."
        
        Case 6:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is currently in " & _
                "use under the owner name: " & ExtraInfo & "."
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit " & _
                "http://www.blizzard.com/support/?id=asc0729p ."
        
        Case 7:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is invalid."
        
        Case 8:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is currently " & _
                "in use under the owner name: " & ExtraInfo & "."
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit " & _
                "http://www.blizzard.com/support/?id=asc0729p ."
        
        Case 9:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is banned. " & _
                "For more information, visit http://www.blizzard.com/support/?id=asc0638p ."
        
        Case 10:
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is for the wrong game."
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit " & _
                "http://www.blizzard.com/support/?id=awr0639p ."
        
        Case Else
            frmChat.AddChat RTBColors.ErrorMessageText, "Unhandled 0x51 response! Value: " & Message
    End Select
    
    If (Message > 0) Then
        Call frmChat.DoDisconnect
    End If
End Sub

Public Sub Event_WhisperFromUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String)
    Dim s       As String
    Dim lCarats As Long
    Dim WWIndex As Integer
    Dim ToANSI  As String
    
    Username = convertUsername(Username)
    
    ' ...
    ToANSI = UTF8Decode(Message)
    
    ' ...
    If (Len(ToANSI) > 0) Then
        Message = ToANSI
    End If

    If (frmChat.mnuUTF8.Checked) Then
        Message = ToANSI
        
        If (Message = vbNullString) Then
            Exit Sub
        End If
    End If
    
    'If ((GetTickCount() - LastWhisperTime) > _
    '    BotVars.AutofilterMS) Then

    If (0 = 0) Then
        If (Not (CheckBlock(Username))) Then
            LastWhisper = Username

            LastWhisperFromTime = Now
        End If
        
        If (Catch(0) <> vbNullString) Then
            Call CheckPhrase(Username, Message, CPWHISPER)
        End If
        
        If (frmChat.mnuFlash.Checked) Then
            FlashWindow
        End If
        
        If (StrComp(Message, BotVars.ChannelPassword, vbTextCompare) = 0) Then
            lCarats = UsernameToIndex(Username)
            
            If (lCarats > 0) Then
                With colUsersInChannel.Item(lCarats)
                    If (.InternalFlags >= IF_AWAITING_CHPW) Then
                        .InternalFlags = .InternalFlags - IF_AWAITING_CHPW
                    End If
                End With
                
                frmChat.AddQ "/w " & Username & " Password accepted."
            End If
        End If
        
        If (VoteDuration > 0) Then
            If (InStr(1, Message, "yes", vbTextCompare) > 0) Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
            ElseIf (InStr(Message, "no", vbTextCompare) > 0) Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
            End If
        End If
                
        lCarats = RTBColors.WhisperCarats
        
        If (Flags And &H1) Then
            lCarats = COLOR_BLUE
        End If
        
        '####### Mail check
        If (Mail) Then
            If (StrComp(Left$(Message, 6), "!inbox", vbTextCompare) = 0) Then
                Dim Msg As udtMail
                
                If (GetMailCount(Username) > 0) Then
                    Call GetMailMessage(Username, Msg)
                    
                    If (Len(RTrim(Msg.To)) > 0) Then
                        frmChat.AddQ "/w " & Username & " Message from " & _
                            RTrim$(Msg.From) & ": " & RTrim$(Msg.Message)
                    End If
                End If
            End If
        End If
        '#######
        
        If ((Not (CheckMsg(Message, Username, -5))) And _
            (Not (CheckBlock(Username)))) Then
        
            If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
                frmChat.AddChat lCarats, "<From ", RTBColors.WhisperUsernames, _
                    Username, lCarats, "> ", RTBColors.WhisperText, Message
            End If
            
            frmChat.AddWhisper lCarats, "<From ", RTBColors.WhisperUsernames, _
                Username, lCarats, "> ", RTBColors.WhisperText, Message
                
            frmChat.rtbWhispers.Visible = rtbWhispersVisible
                           
            If ((frmChat.mnuToggleWWUse.Checked) And _
                (frmChat.WindowState <> vbMinimized)) Then
                
                If (Not (IrrelevantWhisper(Message, Username))) Then
                    WWIndex = AddWhisperWindow(Username)
                    
                    With colWhisperWindows.Item(WWIndex)
                        If (.Shown = False) Then
                            'window was previously hidden
                            
                            ShowWW WWIndex
                        End If
                        
                        .Caption = "Whisper Window: " & Username
                        
                        .AddWhisper RTBColors.WhisperUsernames, "> " & Username, lCarats, _
                            ": ", RTBColors.WhisperText, Message
                    End With
                End If
            End If
        
            Call ProcessCommand(Username, Message, False, True)
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        ' ...
        g_lastQueueUser = Username
        
        ' ...
        frmChat.SControl.Run "Event_WhisperFromUser", Username, Flags, Message
    End If
    
    LastWhisperTime = GetTickCount
End Sub

' Flags and ping are deliberately not used at this time
Public Sub Event_WhisperToUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    Dim WWIndex As Integer
    Dim ToANSI  As String
    
    ' ...
    ToANSI = UTF8Decode(Message)
    
    ' ...
    If (Len(ToANSI) > 0) Then
        Message = ToANSI
    End If
    
    ' ...
    If (StrComp(Username, "your friends", vbTextCompare) <> 0) Then
        Username = convertUsername(Username)
        
        LastWhisperTo = Username
    Else
        LastWhisperTo = "%f%"
    End If

    If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
        frmChat.AddChat RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, _
            IIf(Dii, Mid$(Username, InStr(Username, "*") + 1), Username), _
                RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If
    
    If ((frmChat.mnuHideWhispersInrtbChat.Checked) Or _
        (frmChat.mnuToggleShowOutgoing.Checked)) Then
        
        frmChat.AddWhisper RTBColors.WhisperCarats, "<To ", _
            RTBColors.WhisperUsernames, IIf(Dii, Mid$(Username, InStr(Username, "*") + 1), _
                Username), RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If

    If (frmChat.mnuToggleWWUse.Checked) Then
        If ((InStr(1, Message, "ß~ß") = 0) And _
            (StrComp(Username, "your friends") <> 0)) Then
            
            WWIndex = AddWhisperWindow(Username)
            
            If (frmChat.WindowState <> vbMinimized) Then
                Call ShowWW(WWIndex)
            End If
            
            colWhisperWindows.Item(WWIndex).Caption = "Whisper Window: " & Username
            colWhisperWindows.Item(WWIndex).AddWhisper RTBColors.TalkBotUsername, "> " & _
                CurrentUsername, RTBColors.WhisperCarats, ": ", RTBColors.WhisperText, Message
        End If
    End If
    
    If (Not (rtbWhispersVisible)) Then
        If (frmChat.rtbWhispers.Visible = True) Then
            frmChat.rtbWhispers.Visible = False
        End If
    End If
End Sub

Public Function Event_AccountCreateResponse(ByVal result As Long) As Boolean
    Dim Success As Boolean
    Dim sOut    As String
    
    Success = (result = 0)
    
    Select Case (result)
        Case 1, 6: sOut = "Your desired account name does not contain enough alphanumeric characters."
        Case 2:    sOut = "Your desired account name contains invalid characters."
        Case 3:    sOut = "Your desired account name contains a banned word."
        Case 4:    sOut = "Your desired account name already exists."
        Case Else: sOut = "Unknown response to 0x3D. Result code: " & result
    End Select
    
    If (Success) Then
        frmChat.AddChat RTBColors.SuccessText, _
            "[BNET] Account created successfully!"
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, _
            "There was an error in trying to create a new account."
        frmChat.AddChat RTBColors.ErrorMessageText, sOut
    End If
    
    Event_AccountCreateResponse = Success
End Function

Public Function Event_RealmStatusError(ByVal Status As Long)
    Select Case (Status)
        Case &H80000001:
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] The Diablo II Realm is currently " & _
                "unavailable. Please try again later."
        Case &H80000002:
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Diablo II Realm logon has failed. " & _
                "Please try again later."
        Case Else:
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Login to the Diablo II Realm " & _
                "has failed for an unknown reason (0x" & ZeroOffset(Status, 8) & "). Please try again later."
    End Select
    
    RealmError = True
End Function

'11/22/07 - Hdx - Pass the channel listing (0x0B) directly off to scriptors for there needs. (What other use is there?)
Public Sub Event_ChannelList(sChannels() As String)
    If (MDebug("all")) Then
        Dim X As Integer
        
        frmChat.AddChat RTBColors.InformationText, "Received Channel List: "
        
        For X = 0 To UBound(sChannels)
            frmChat.AddChat RTBColors.InformationText, vbTab & _
                sChannels(X)
        Next X
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    frmChat.SControl.Run "Event_ChannelList", sChannels
End Sub

Private Function CleanDiablo2Username(ByVal Username As String) As String
    
    Dim tmp As String  ' ...
    Dim Pos As Integer ' ...
    
    ' ...
    tmp = Username
    
    ' ...
    Pos = InStr(1, tmp, "*", vbBinaryCompare)

    ' ...
    If (Pos > 0) Then
        tmp = Mid$(Username, Pos + 1)
        
        ' ...
        If (Right$(tmp, 1) = ")") Then
            tmp = Left$(tmp, Len(tmp) - 1)
        End If
    End If
    
    ' ...
    CleanDiablo2Username = tmp
    
End Function

Private Function GetDiablo2CharacterName(ByVal Username As String) As String

    Dim tmp As String  ' ...
    Dim Pos As Integer ' ...
    
    ' ...
    Pos = InStr(1, Username, "*", vbBinaryCompare)

    ' ...
    If (Pos > 0) Then
        tmp = Mid$(Username, 1, Pos - 1)
    End If
    
    ' ...
    GetDiablo2CharacterName = tmp

End Function
