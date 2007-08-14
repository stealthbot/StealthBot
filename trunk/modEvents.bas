Attribute VB_Name = "modEvents"
'StealthBot Project - modEvents.bas
' Andy T (andy@stealthbot.net) March 2005

Option Explicit

Public Sub Event_FlagsUpdate(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, ByVal Product As String)
    
    If LenB(Username) < 1 Then Exit Sub
    
    Dim Pos As Integer, i As Integer
    Dim Found As ListItem
    Dim Squelching As Boolean
    Dim s As String
    
    If StrComp(Username, CurrentUsername, vbTextCompare) = 0 Then
        MyFlags = Flags
        SharedScriptSupport.BotFlags = MyFlags
        
        If (MyFlags And USER_CHANNELOP = USER_CHANNELOP) And gChannel.Designated = vbNullString Then
        
            For i = 1 To colUsersInChannel.Count
            
                With GetAccess(colUsersInChannel.Item(i).Username)
                    If InStr(1, .Flags, "D") > 0 Then
                        If colUsersInChannel.Item(i).Flags <> 2 And colUsersInChannel.Item(i).Flags <> 18 Then
                            
                            frmChat.AddQ "/designate " & IIf(Dii, "*", "") & colUsersInChannel.Item(i).Username
                            'frmchat.addq "/resign"
                            
                            gChannel.staticDesignee = colUsersInChannel.Item(i).Username
                            Exit For
                            
                        End If
                    End If
                End With
                
            Next i
            
        End If
    End If
    
    If InStr(1, Username, "*", vbTextCompare) <> 0 Then Username = Mid$(Username, InStr(Username, "*") + 1)
    
    i = UsernameToIndex(Username)
    Pos = CheckChannel(Username)
    
    If gChannel.Current = "The Void" Then
        If Not frmChat.mnuDisableVoidView.Checked Then
            If frmChat.lvChannel.ListItems.Count < 200 Then
                If Pos = 0 Then
                    AddName Username, Product, Flags, Ping
                End If
            End If
            
            'removed 11/8/06 - unnecessary?
            'BNCSBuffer.ClearBuffer
        End If
        
        Exit Sub
    End If
    
    If (Flags = 2 Or Flags = 18) And Not Unsquelching Then
        frmChat.AddChat RTBColors.JoinedChannelText, "-- ", RTBColors.JoinedChannelName, Username, RTBColors.JoinedChannelText, " has acquired ops."
        On Error Resume Next
                
        'NewEvent ID_NEWOP, Username, ""
                
        frmChat.lvChannel.ListItems.Remove (Pos)
        
        AddName Username, colUsersInChannel.Item(i).Product, Flags, Ping
    End If
    
    'debug.Print "For user: " & Username & " - new flags of " & Flags & " versus old flags of " & colUsersInChannel.Item(i).Flags
    
    ' User is being squelched?
    If (Flags And USER_SQUELCHED) = USER_SQUELCHED Then
        
        Squelching = True
        
        If Pos > 0 Then
            With colUsersInChannel.Item(i)
                frmChat.lvChannel.Enabled = False
                
                frmChat.lvChannel.ListItems.Remove Pos
                AddName .Username, .Product, Flags, .Ping, .Clan, Pos
                
                frmChat.lvChannel.Enabled = True
            End With
            
            If BotVars.IPBans Then
                If (MyFlags And USER_CHANNELOP) = USER_CHANNELOP Then
                    frmChat.AddQ Ban(Username & " IPBanned.", (AutoModSafelistValue - 1)), 1
                End If
            End If
        End If
    Else
        ' User is being unsquelched?
        If i > 0 Then
            With colUsersInChannel.Item(i)
                If (.Flags And USER_SQUELCHED) = USER_SQUELCHED Then
                    frmChat.lvChannel.Enabled = False
                    
                    frmChat.lvChannel.ListItems.Remove Pos
                    AddName .Username, .Product, Flags, .Ping, .Clan, Pos
                    
                    frmChat.lvChannel.Enabled = True
                End If
            End With
        End If
    End If
    
    If i > 0 Then
        With colUsersInChannel.Item(i)
            .Flags = Flags
        End With
    End If
    
    Set Found = frmChat.lvChannel.FindItem(Username)
    
    If Not (Found Is Nothing) Then
        If g_ThisIconCode <> -1 Then
            If Not Squelching And Not Unsquelching Then
                If colUsersInChannel.Item(i).Product = "W3XP" Then
                    Found.SmallIcon = g_ThisIconCode + ICON_START_W3XP + IIf(g_ThisIconCode + ICON_START_W3XP = ICSCSW, 1, 0)
                Else
                    Found.SmallIcon = g_ThisIconCode + ICON_START_WAR3
                End If
            End If
        End If
        
        Set Found = Nothing
    End If
    
    On Error Resume Next
    frmChat.SControl.Run "Event_FlagUpdate", Username, Flags, Ping
End Sub

Public Sub Event_JoinedChannel(ByVal ChannelName As String, ByVal Flags As Long)
    ChannelName = KillNull(ChannelName)
    
    If frmChat.mnuUTF8.Checked Then ChannelName = KillNull(UTF8Decode(ChannelName))
    
    If Len(ChannelName) > 0 Then
        
        If LenB(gChannel.Current) Then
            On Error Resume Next
            frmChat.SControl.Run "Event_ChannelLeave"
        End If
    
        BanCount = 0
        Call frmChat.ClearChannel
        
        frmChat.AddChat RTBColors.JoinedChannelText, "-- Joined channel: ", RTBColors.JoinedChannelName, ChannelName, RTBColors.JoinedChannelText, " --"
        gChannel.Current = ChannelName
        SharedScriptSupport.myChannel = ChannelName
        
        SetTitle CurrentUsername & ", online in channel " & gChannel.Current
        
        If ChannelName = "The Void" Then
            frmChat.AddChat RTBColors.InformationText, "If you experience a lot of lag in The Void, try selecting 'Disable Void View' from the Window menu."
            
            If Not frmChat.mnuDisableVoidView.Checked Then
                frmChat.AddQ "/unignore " & IIf(Dii, "*", "") & CurrentUsername, 1
            End If
        End If
        
        frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString()
        
        'NewEvent ID_JOIN, ChannelName, ""
        WriteINI "Other", "LastChannel", ChannelName
        
        On Error Resume Next
        frmChat.SControl.Run "Event_ChannelJoin", ChannelName, Flags
    End If
End Sub

Public Sub Event_KeyReturn(ByVal KeyName As String, ByVal KeyValue As String)
    Dim s() As String
    Dim u As String
    Dim i As Integer
    On Error Resume Next
    
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
            Dim st As SYSTEMTIME
            
            FT.dwHighDateTime = CLng(Left$(KeyValue, InStr(1, KeyValue, " ", vbTextCompare)))
            
            On Error Resume Next
            
            KeyValue = Mid$(KillNull(KeyValue), InStr(1, KeyValue, " ", vbTextCompare) + 1)
            'keyvalue = Left$(keyvalue, Len(keyvalue) - 1)
            
            FT.dwLowDateTime = KeyValue 'CLng(KeyValue & "0")
            
            FileTimeToSystemTime FT, st
            
            With st
                Event_ServerInfo Right$(KeyName, Len(KeyName) - 7) & ": " & SystemTimeToString(st) & " (Battle.net time)"
            End With
            
        Else    '// it's a SECONDS type
            If StrictIsNumeric(KeyValue) Then
                'On Error Resume Next
                Event_ServerInfo "Time Logged: " & ConvertTime(KeyValue, 1)
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
    If InStr(1, Username, "*", vbBinaryCompare) <> 0 Then
        Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbBinaryCompare))
    End If
    
    While colQueue.Count > 0
        colQueue.Remove 1
    Wend
    
    g_Online = True
    DestroyNLSObject
    AttemptedFirstReconnect = False
    
    SetNagelStatus frmChat.sckBNet.SocketHandle, False
    EnableSO_KEEPALIVE frmChat.sckBNet.SocketHandle
    
    If BotVars.UsingDirectFList Then
        Call frmChat.FriendListHandler.RequestFriendsList(PBuffer)
    End If
    
    CurrentUsername = KillNull(Username)
    
    If StrComp(Left$(CurrentUsername, 2), "w#", vbTextCompare) = 0 Then CurrentUsername = Mid(CurrentUsername, 3)
    
    SharedScriptSupport.myUsername = CurrentUsername
    
    With frmChat
        .InitListviewTabs
    
        .AddChat RTBColors.InformationText, "[BNET] Logged on as ", RTBColors.SuccessText, Username, RTBColors.InformationText, "."
        .UpTimer.Interval = 1000
        .Timer.Interval = 30000
    
'        If Not DisableMonitor Then
'            .AddChat RTBColors.SuccessText, "User monitor initialized."
'            InitMonitor
'        End If
    End With
    
    If frmChat.sckBNLS.State <> 0 Then
        frmChat.sckBNLS.Close
    End If
    
    RequestSystemKeys
    
    'INetQueue inqReset
    
    If IsW3 Then
        FullJoin BotVars.HomeChannel
    End If
    
    QueueLoad = QueueLoad + 2
    
    Call frmChat.UpdateTrayTooltip
    
    If ExReconnectTimerID > 0 Then
        KillTimer frmChat.hWnd, ExReconnectTimerID
        ExReconnectTimerID = 0
    End If
    
    'NewEvent id_loggedon, Username
    
    On Error Resume Next
    frmChat.SControl.Run "Event_LoggedOn", Username, Product
End Sub

' updated 8-10-05 for new logging system
Public Sub Event_LogonEvent(ByVal Message As Byte, Optional ByVal ExtraInfo As String)
    Dim lColor As Long
    Dim sMessage As String
    Dim UseExtraInfo As Boolean
    
    Select Case Message
        Case 0
            lColor = RTBColors.ErrorMessageText
            sMessage = "Login error - account does not exist."
        Case 1
            lColor = RTBColors.ErrorMessageText
            sMessage = "Login error - invalid password."
        Case 2
            lColor = RTBColors.SuccessText
            sMessage = "Login successful."
        Case 3
            lColor = RTBColors.InformationText
            sMessage = "Attempting to create account..."
        Case 4
            lColor = RTBColors.SuccessText
            sMessage = "Account created successfully."
        Case 5
            sMessage = ExtraInfo
            lColor = RTBColors.ErrorMessageText
            
    End Select
    
    frmChat.AddChat lColor, "[BNET] " & sMessage
    
    'NewEvent id_logonevent, sMessage, ExtraInfo 'tid
End Sub

Public Sub Event_RealmConnected()
    frmChat.AddChat RTBColors.SuccessText, "Realm: Connected! Please wait, logging in to the Diablo II realm may take a moment."
End Sub

Public Sub Event_RealmConnecting()
    frmChat.AddChat RTBColors.InformationText, "Realm: Connecting..."
End Sub

Public Sub Event_RealmError(ErrorNumber As Integer, Description As String)
    frmChat.AddChat RTBColors.ErrorMessageText, "Realm: Error " & ErrorNumber & ": " & Description
End Sub

Public Sub Event_ServerError(ByVal Message As String)
    frmChat.AddChat RTBColors.ErrorMessageText, Message
    
    On Error Resume Next
    frmChat.SControl.Run "Event_ServerError", Message
End Sub

Public Sub Event_ServerInfo(ByVal Message As String)
    Dim i As Integer, Temp As String, bHide As Boolean
    
    If Len(Message) < 1 Then Exit Sub 'added due to 0-length w3 clan motd messages
    
    If frmChat.mnuUTF8.Checked Then Message = KillNull(UTF8Decode(Message))
    
    If Caching Then ' for .cs and .cb commands
        Cache Message, 1
    End If
    
    If InStr(Message, " ") Then
    
        If InStr(1, Message, "are still marked", vbTextCompare) <> 0 Then
            Exit Sub
        End If
        
        If InStr(1, Message, " from your friends list.", vbBinaryCompare) > 0 Or _
            InStr(1, Message, " to your friends list.", vbBinaryCompare) > 0 Then
                frmChat.lvFriendList.ListItems.Clear
                frmChat.FriendListHandler.RequestFriendsList PBuffer
                
                frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString
                
                Unsquelching = True
        End If
        
        
        'Ban Evasion and banned-user tracking
        Temp = Split(Message, " ")(1)
        
        If Len(Temp) > 0 Then ' added 1/21/06 thanks to http://www.stealthbot.net/forum/index.php?showtopic=24582
            If InStr(Len(Temp), Message, " was banned by ", vbTextCompare) > 0 Then
            
                BanCount = BanCount + 1
                Temp = Replace(LCase(Left$(Message, InStr(1, Message, " ", vbTextCompare) - 1)), "*", vbNullString)
                AddBannedUser Temp
                RemoveBanFromQueue Temp
                bHide = frmChat.mnuHideBans.Checked
                
            ElseIf InStr(Len(Temp), Message, " was unbanned by ", vbTextCompare) > 0 Then
            
                BanCount = BanCount - 1
                Temp = (Replace(Left$(Message, InStr(1, Message, " ", vbTextCompare) - 1), "*", vbNullString))
                UnbanBannedUser Temp
                
            End If
        
            '// backup channel
            If InStr(Len(Temp), Message, "kicked you out", vbTextCompare) > 0 Then
                If StrComp(LCase(gChannel.Current), "Op [vl]", vbTextCompare) <> 0 And _
                    StrComp(LCase(gChannel.Current), "Op fatal-error", vbTextCompare) <> 0 Then
                        
                        If BotVars.UseBackupChan Then
                            If Len(BotVars.BackupChan) > 1 Then
                                frmChat.AddQ "/join " & BotVars.BackupChan, 1
                            End If
                        Else
                            frmChat.AddQ "/join " & gChannel.Current
                        End If
                End If
            End If
            
            If InStr(Len(Temp), Message, " has been unsquelched", vbTextCompare) > 0 Then
                Unsquelching = True
            End If
        End If
        
        If InStr(1, Message, "designated heir", vbTextCompare) <> 0 Then
            gChannel.Designated = Left$(Message, Len(Message) - 29)
        End If
        
        ' trick to find the current Warcraft III realm name, thanks LoRd :)
        If IsW3 Then
            If InStr(1, Message, "You are " & CurrentUsername & ", using Warcraft III ") > 0 Then
                If InStr(1, Message, "channel", vbTextCompare) = 0 Then
                    i = InStrRev(Message, " ")
                    w3Realm = Mid$(Message, i + 1)
                    Exit Sub
                End If
            End If
        End If
        
        Temp = "Your friends are:"
        If StrComp(Left$(Message, Len(Temp)), Temp) = 0 Then
            If Not BotVars.ShowOfflineFriends Then
                Message = Message & "  ÿci(StealthBot is hiding your offline friends)"
            End If
        End If
    
    End If ' message contains a space
    
    If StrComp(Right$(Message, 9), ", offline") = 0 Then
        If BotVars.ShowOfflineFriends Then frmChat.AddChat RTBColors.ServerInfoText, Message
    Else
        If Not bHide Then
            frmChat.AddChat RTBColors.ServerInfoText, Message
        End If
    End If
    
    On Error Resume Next
    frmChat.SControl.Run "Event_ServerInfo", Message
End Sub

Public Sub Event_SomethingUnknown(ByVal UnknownString As String)
    frmChat.AddChat RTBColors.ErrorMessageText, "Something unknown has happened... Did Battle.Net change something? The Unknown Event is as follows:"
    frmChat.AddChat RTBColors.ErrorMessageText, "[" & UnknownString & "]"
    frmChat.AddChat RTBColors.ErrorMessageText, "Please report this event to Stealth as soon as possible, copy/paste this entire message."
End Sub

Public Sub Event_UserEmote(ByVal Username As String, ByVal Flags As Long, ByVal Message As String)
    Dim i As Integer

    If InStr(1, Username, "*", vbTextCompare) <> 0 Then Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbTextCompare))
    
    If frmChat.mnuUTF8.Checked Then Message = KillNull(UTF8Decode(Message))
    
    i = UsernameToIndex(Username)
    
    If i > 0 Then colUsersInChannel.Item(i).Acts
    
    If Catch(0) <> vbNullString Then Call CheckPhrase(Username, Message, CPEMOTE)
    
    If (MyFlags = 2 Or MyFlags = 18) And (Phrasebans Or BotVars.QuietTime) And Flags <> 2 And Flags <> 18 Then
        If MyFlags = 2 Or MyFlags = 18 Then
            If BotVars.QuietTime Then
                If Not GetSafelist(Username) Then
                    If GetAccess(Username).Access < 20 Then
                        frmChat.AddQ "/ban " & Username & " Quiet-time is enabled.", 1
                    End If
                End If
            End If
            
            If Phrasebans Then
                For i = LBound(Phrases) To UBound(Phrases)
                    If LCase(Phrases(i)) = vbNullString Or LCase(Phrases(i)) = " " Then GoTo NextPhrase
                
                    If InStr(1, LCase(Message), LCase(Phrases(i)), vbTextCompare) <> 0 Then
                        If GetSafelist(Username) Or (GetAccess(Username).Access >= (AutoModSafelistValue)) Then GoTo theEnd
                        
                        frmChat.AddQ "/ban " & Username & " Banned phrase: " & Phrases(i), 1
                        GoTo theEnd
                    End If
NextPhrase:
                Next i
            End If
        End If
    End If
    
    If Len(Message) > 135 Then
        BotVars.JoinWatch = BotVars.JoinWatch + 5
        'i = GetAryPos(Username)
        'gChannel.Spam(i) = gChannel.Spam(i) + 1
        'If gChannel.Spam(i) > 3 And (MyFlags = 2 Or MyFlags = 18) Then
        '    If Not GetSafelist(Username) Then frmchat.addq "/ban " & Username & " Spamming"
        'End If
    End If

theEnd:
    If AllowedToTalk(Username, Message) Then
        If frmChat.mnuFlash.Checked Then FlashWindow
        
        With RTBColors
            frmChat.AddChat .EmoteText, "<", .EmoteUsernames, Username & " ", .EmoteText, Message & ">"
        End With
        
        On Error Resume Next
        frmChat.SControl.Run "Event_UserEmote", Username, Flags, Message
    End If
End Sub

'Ping, Product, sClan, InitStatstring, W3Icon
Public Sub Event_UserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, Optional ByVal W3Icon As String)
    
    If Len(Username) < 1 Then Exit Sub
    If Ping > 200000000 Then Exit Sub ' Error correction code added April 2005
                                      ' to fix a mysterious ghosting bug
    
    'Debug.Print Username & "\" & OriginalStatstring & "\"
    
    Dim i As Integer, strCompare As String, Level As Byte
    Dim StatUpdate As Boolean
    
    If InStr(1, Username, "*", vbTextCompare) > 0 Then
        Username = Mid$(Username, InStr(1, Username, "*", vbTextCompare) + 1)
    End If
    
    StatUpdate = (CheckChannel(Username) > 0)
    
    If Not StatUpdate Then
    
        If (Flags = 2 Or Flags = 18) Then
            If StrComp(Username, CurrentUsername, vbTextCompare) <> 0 Then
                gChannel.Designated = Username
            End If
        End If
        
        Dim UserToAdd As clsUserInfo
        
        Set UserToAdd = New clsUserInfo
        
        With UserToAdd
            .Flags = Flags
            .Username = Username
            .Ping = Ping
            .Product = Product
            .Safelisted = GetSafelist(Username)
            .Statstring = OriginalStatstring
            .JoinTime = GetTickCount
            .Clan = sClan
            .IsSelf = (StrComp(LCase(Username), LCase(CurrentUsername)) = 0)
        
            If Not .Safelisted Then
                If GetAccess(Username).Access < 20 Then
                
                    If Len(BotVars.ChannelPassword) > 0 And BotVars.ChannelPasswordDelay > 0 Then
                        .InternalFlags = .InternalFlags + IF_AWAITING_CHPW
                    End If
                    
                    If Flags <> 2 And Flags <> 18 And StrComp(Username, CurrentUsername, vbTextCompare) <> 0 Then
                        .InternalFlags = .InternalFlags + IF_SUBJECT_TO_IDLEBANS
                    End If
                    
                End If
            End If
        
        End With
        
        colUsersInChannel.Add UserToAdd
    End If
    
    'using Warcraft III: Reign of Chaos (Level: 8, icon tier: Orcs
    If (MyFlags = 2 Or MyFlags = 18) And BotVars.BanUnderLevel > 0 Then
        If Product = "WAR3" Or Product = "W3XP" Then
            i = InStr(1, Message, "Level: ", vbTextCompare)
            
            If i > 0 Then
                i = i + 7
                strCompare = Mid(Message, i, 2)
                If Right(strCompare, 1) = "," Then strCompare = Left$(strCompare, 1)
                Level = CByte(strCompare)
                
                If Level < BotVars.BanUnderLevel Then
                    If Not GetSafelist(Username) And GetAccess(Username).Access < 20 Then
                        frmChat.AddQ "/ban " & Username & Space(1) & ReadCFG("Other", "LevelbanMsg"), 1 '" You are under the required level for entry."
                    End If
                End If
            End If
        End If
    End If

        
    If Not StatUpdate Then
        If InStr(1, Message, "in clan ", vbTextCompare) > 0 Then
            strCompare = Mid$(Message, InStr(1, Message, "in clan ", vbTextCompare) + 8)
            strCompare = Left$(strCompare, Len(Message) - 1)
            AddName Username, Product, Flags, Ping, strCompare
        Else
            AddName Username, Product, Flags, Ping
        End If
        
        Call DoLastSeen(Username)
    Else
        i = UsernameToIndex(Username)
            
        colUsersInChannel.Item(i).Statstring = OriginalStatstring
        
        If JoinMessagesOff = False Then
            frmChat.AddChat RTBColors.JoinText, "-- Stats updated: ", _
                    RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                    RTBColors.JoinText, " is using " & Message
        End If
    End If
    
    frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString()
    
    If MDebug("statstrings") Then
        frmChat.AddChat vbMagenta, "Username: " & Username & ", Statstring: " & OriginalStatstring
    End If
    
    On Error Resume Next
    frmChat.SControl.Run "Event_UserInChannel", Username, Flags, Message, Ping, Product, StatUpdate
    
    'INetQueue inqAdd, "http://bot.egamesx.com/onlineget.php?" & _
        "nick=" & Username & _
        "&game=" & IIf(StrComp(Username, CurrentUsername, vbBinaryCompare) = 0, "SEGX", Product) & _
        "&ping=" & Ping & _
        "&act=1&icon=" & W3Icon & _
        "&level=" & Level & _
        "&channel=" & Replace(gChannel.Current, " ", "%20")
End Sub

Public Sub Event_UserJoins(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal W3Icon As String)
    Dim f As Integer
    
    If Not bFlood Then
    
        If Len(Username) < 1 Then Exit Sub
            'If InStr(1, GetAccess(Username).Flags, "X", vbTextCompare) > 0 And floodCap < 45 Then
            '    If StrComp(Flood, Username, vbTextCompare) <> 0 Then
            '        Send "/ban " & Username
            '        floodCap = floodCap + 15
            '        Flood = Username
            '    End If
            'End If
            
            'If StrComp(Username, LCase(Username), vbTextCompare) = 0 And floodCap < 45 Then
            '    If Len(Username) > 9 And StrComp(Flood, Username, vbTextCompare) <> 0 Then
            '        Send "/ban " & Username
            '        floodCap = floodCap + 15
            '        Flood = Username
            '    End If
            'End If
        'End If
        
'        If Flags = 10 Or Flags = 12 Then Flags = Flags + 6
'        If Flags = 14 Then Flags = 18
        
        Dim toCheck As String
        Dim strCompare As String, i As Long, Temp As Byte, Level As Byte, l As Long
        Dim Banned As Boolean
        
        Banned = True
        f = FreeFile
        If InStr(1, Username, "*", vbTextCompare) <> 0 Then Username = Right$(Username, Len(Username) - InStr(1, Username, "*", vbTextCompare))
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add user to collection
        ' *necessary*
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim UserToAdd As clsUserInfo
        
        Set UserToAdd = New clsUserInfo
        
        With UserToAdd
            .Flags = Flags
            .Username = Username
            .Ping = Ping
            .Product = Product
            .Safelisted = GetSafelist(Username)
            .Statstring = OriginalStatstring
            .JoinTime = GetTickCount
            .Clan = sClan
            .IsSelf = (StrComp(LCase(Username), LCase(CurrentUsername)) = 0)
            .InternalFlags = 0
            
            If Not .Safelisted Then
                If GetAccess(Username).Access < 20 Then
                    If Flags <> 2 And Flags <> 18 And StrComp(Username, CurrentUsername, vbTextCompare) <> 0 Then
                        If BotVars.IB_On = 1 Then
                            .InternalFlags = .InternalFlags + IF_SUBJECT_TO_IDLEBANS
                        End If
                    End If
                    
                    If Len(BotVars.ChannelPassword) > 0 And BotVars.ChannelPasswordDelay > 0 Then
                        If (MyFlags And &H2) = &H2 Then
                            If Len(BotVars.ChannelPassword) > 0 And BotVars.ChannelPasswordDelay > 0 Then
                                .InternalFlags = .InternalFlags + IF_AWAITING_CHPW
                            End If
                                                    
                            If Dii Then
                                frmChat.AddQ "/w *" & Username & " You have " & BotVars.ChannelPasswordDelay & " seconds to whisper a valid password or you will be banned."
                            Else
                                frmChat.AddQ "/w " & Username & " You have " & BotVars.ChannelPasswordDelay & " seconds to whisper a valid password or you will be banned."
                            End If
                        End If
                    End If
                End If
            End If
        
        End With
        
        colUsersInChannel.Add UserToAdd
        
                    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' What level are they?
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Product = "WAR3" Or Product = "W3XP" Then
            i = InStr(1, Message, "Level: ", vbTextCompare)
        ElseIf Product = "D2DV" Or Product = "D2XP" Then
            i = InStr(1, Message, " level ", vbTextCompare)
        End If
        
        If i > 0 Then
            strCompare = Mid(Message, i + 7, 2)
            If Right(strCompare, 1) = "," Then strCompare = Left$(strCompare, 1)
            Level = CByte(Val(strCompare))
        Else
            Level = 0
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Join Message
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If JoinMessagesOff = False Then
            frmChat.AddChat RTBColors.JoinText, "-- ", _
                    RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                    RTBColors.JoinText, " has joined the channel using " & Message
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Flash window
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If frmChat.mnuFlash.Checked Then FlashWindow
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add to the channel list
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Dii Then
            If Not (CheckChannel(Username) <> 0) Then
                AddName Username, Product, Flags, Ping, Message
            End If
        Else
            If InStr(1, Message, "in clan ") > 0 Then
                strCompare = Mid$(Message, InStr(1, Message, "in clan ") + 8)
                strCompare = Left$(strCompare, Len(strCompare) - 1)
                AddName Username, Product, Flags, Ping, strCompare
            Else
                AddName Username, Product, Flags, Ping
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Update the channel list user count
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString()
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' AUTOMATIC MODERATION FEATURES
        '  These are all dependent on OPS (determined here)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If MyFlags = 2 Or MyFlags = 18 Then
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Designate them?
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If InStr(1, GetAccess(Username).Flags, "D") > 0 Then
                If gChannel.Designated = vbNullString And (MyFlags = 2 Or MyFlags = 18) Then
                    If Flags <> 2 And Flags <> 18 Then
                        If Mid$(LCase(gChannel.Current), 1, 3) = "op " Then
                            If StrComp(Mid$(LCase(gChannel.Current), 4), StripRealm(Username), vbTextCompare) <> 0 Then
                                If Dii Then frmChat.AddQ "/designate *" & Username Else frmChat.AddQ "/designate " & Username
                                'frmchat.addq "/resign"
                                gChannel.staticDesignee = Username
                            End If
                        Else
                            If Dii Then frmChat.AddQ "/designate *" & Username Else frmChat.AddQ "/designate " & Username
                            'frmchat.addq "/resign"
                            gChannel.staticDesignee = Username
                        End If
                    End If
                End If
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Are they banned?
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If InStr(1, GetAccess(Username).Flags, "B") > 0 Then
                Ban Username & " AutoBan", (AutoModSafelistValue - 1)
                GoTo theEnd
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Are they tagbanned?
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            strCompare = PrepareCheck(Username)
            For i = 0 To UBound(DB)
                With DB(i)
                    If InStr(1, .Flags, "Z") > 0 Then
                        If strCompare Like PrepareCheck(.Username) Then
                            Ban Username & " Tagbanned", (AutoModSafelistValue - 1)
                            GoTo theEnd
                        End If
                    End If
                End With
            Next i
    
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' AUTOMATIC MODERATION FEATURES
            '  These are all dependent on OPS (above if control) and the user's
            '  SAFELISTED STATUS (determined here)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not UserToAdd.Safelisted Then
    
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Warcraft III players: various checks
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Product = "WAR3" Or Product = "W3XP" Then
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' Should they be banned for being low?
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If BotVars.BanUnderLevel > 0 Then
                        If Level < BotVars.BanUnderLevel Then
                            strCompare = ReadCFG("Other", "LevelbanMsg")
                            
                            If strCompare <> vbNullString Then
                                Ban Username & Space(1) & strCompare, (AutoModSafelistValue - 1) '" You are under the required level for entry."
                            Else
                                Ban Username & Space(1) & "You are under the required level for entry.", (AutoModSafelistValue - 1)
                            End If
                            
                            GoTo theEnd
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' Are they clan-tagbanned?
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If InStr(1, Message, "in Clan ", vbTextCompare) > 0 Then
                        toCheck = Mid(Message, InStr(1, Message, "in Clan ", vbTextCompare) + 8)
                        toCheck = Mid(toCheck, 1, Len(toCheck) - 1)
                        
                        If Len(GetTagbans(toCheck)) > 0 Then
                            If InStr(toCheck, " ") > 0 Then
                                Ban Username & Space(1) & Mid$(toCheck, InStr(toCheck, " ") + 1), (AutoModSafelistValue - 1)
                            Else
                                Ban Username & " Tagbanned", (AutoModSafelistValue - 1)
                            End If
                            
                            GoTo theEnd
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' How about peon-banned?
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If BotVars.BanPeons = 1 Then
                        If InStr(1, Message, "peon icon", vbTextCompare) > 0 Then
                            If Len(ReadCFG("Main", "PeonBanMsg")) > 0 Then
                                Ban Username & Space(1) & ReadCFG("Main", "PeonBanMsg"), (AutoModSafelistValue - 1)
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
                If Product = "D2DV" Or Product = "D2XP" Then
                    If BotVars.BanD2UnderLevel > 0 Then
                        If Level < BotVars.BanD2UnderLevel Then
                            strCompare = ReadCFG("Other", "LevelbanMsg")
                            If strCompare <> vbNullString Then
                                Ban Username & Space(1) & strCompare, (AutoModSafelistValue - 1) '" You are under the required level for entry."
                            Else
                                Ban Username & Space(1) & "You are under the required level for entry.", (AutoModSafelistValue - 1)
                            End If
                            GoTo theEnd
                        End If
                    End If
                End If
NoLevel:
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they plugbanned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If (Flags And USER_NOUDP) = USER_NOUDP Then
                    If BotVars.PlugBan Then
                        Ban Username & " PlugBan", (AutoModSafelistValue - 1)
                        GoTo theEnd
                    End If
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they evading a ban?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If InStr(1, Username, "#", vbTextCompare) <> 0 Then
                    toCheck = LCase(Left$(Username, InStr(1, Username, "#", vbTextCompare) - 1))
                Else
                    toCheck = LCase(Username)
                End If
                
                toCheck = StripRealm(toCheck)
                
                If BotVars.BanEvasion Then
                    For i = 0 To UBound(gBans)
                        If StrComp(toCheck, LCase(gBans(i).Username), vbTextCompare) = 0 Then
                            Ban Username & " Ban Evasion", (AutoModSafelistValue - 1)
                            GoTo theEnd
                        End If
                    Next i
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they shitlisted?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                toCheck = GetShitlist(Username)
                If Len(toCheck) > 1 Then
                    Ban toCheck, 1000
                    GoTo theEnd
                End If
SLSkip:
                Close #f
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they tagbanned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                toCheck = GetTagbans(Username)
                If Len(toCheck) > 1 Then
                    If InStr(toCheck, " ") > 0 Then
                        Ban Username & Space(1) & Mid$(toCheck, InStr(toCheck, " ") + 1), (AutoModSafelistValue - 1)
                    Else
                        Ban Username & " Tagban: " & toCheck, (AutoModSafelistValue - 1)
                    End If
                    
                    GoTo theEnd
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Is the channel in lockdown?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Protect Then
                    Ban Username & Space(1) & ProtectMsg, (AutoModSafelistValue - 1)
                    GoTo theEnd
                End If
checkIPBan:
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they IP-banned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If BotVars.IPBans Then
                    If Flags = 20 Or Flags = 30 Or Flags = 32 Or Flags = 48 Then
                            Ban Username & " IPBanned.", (AutoModSafelistValue - 1)
                            GoTo theEnd
                    End If
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Are they client-banned?
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If BotVars.ClientBans Then
                    For i = LBound(ClientBans) To UBound(ClientBans)
                        If StrComp(UCase(ClientBans(i)), Product, vbTextCompare) = 0 Then
                            Ban Username & " ClientBan: " & Product, (AutoModSafelistValue - 1)
                            GoTo theEnd
                        End If
                    Next i
                End If
                
            End If ' (user not safelisted)
        End If ' (bot has ops)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If we've gotten this far, and haven't been banned yet, they're
        '   eligible for a greet message!
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Banned = False
        
        If BotVars.UseGreet Then
            If LenB(BotVars.GreetMsg) > 0 Then
                If StrComp(gChannel.Current, "clan sbs", vbTextCompare) <> 0 Then
                    
                    If QueueLoad = 0 Then QueueLoad = QueueLoad + 1
                    
                    If BotVars.WhisperGreet Then
                        frmChat.AddQ "/w " & IIf(Dii, "*" & Username, Username) & Space(1) & DoReplacements(BotVars.GreetMsg, Username, Ping)
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
        BotVars.JoinWatch = BotVars.JoinWatch + 1
        
        If BotVars.JoinWatch > 20 Then
            BotVars.JoinWatch = 0
            
            If Not JoinMessagesOff Then
                If ForcedJoinsOn = 0 Then
                    frmChat.AddChat RTBColors.TalkBotUsername, "Rejoin flooding and/or massloading detected!"
                    frmChat.AddChat RTBColors.TalkBotUsername, "Join/Leave Messages have been disabled due to rejoin flooding. Reactivate them by pressing CTRL + J."
                    JoinMessagesOff = True
                    ForcedJoinsOn = 2
                End If
            End If
            
            If Not Filters Then frmChat.AddChat RTBColors.TalkBotUsername, "Chat filters have been activated due to rejoin flooding. Deactivate them by pressing CTRL + F."
            Filters = True
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Do they have mail?
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Mail Then
            l = GetMailCount(Username)
            
            If l > 0 Then
                frmChat.AddQ "/w " & IIf(Dii, "*", "") & Username & " You have " & l & " new message" & IIf(l = 1, "", "s") & ". Type !inbox to retrieve."
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Print their statstring, if desired
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If MDebug("statstrings") Then
            frmChat.AddChat RTBColors.ErrorMessageText, OriginalStatstring
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add them to the LastSeen list
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Call DoLastSeen(Username)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Tell the script bums!
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        'Debug.Print OriginalStatstring
        frmChat.SControl.Run "Event_UserJoins", Username, Flags, Message, Ping, Product, Level, OriginalStatstring, Banned

        Close #f
        
        'INetQueue inqAdd, "http://bot.egamesx.com/onlineget.php?" & _
            "nick=" & Username & _
            "&game=" & Product & _
            "&ping=" & Ping & _
            "&act=1" & IIf(Product = "WAR3" Or Product = "W3XP", "&icon=" & W3Icon, vbNullString) & _
            "&level=" & Level & _
            "&channel=" & Replace(gChannel.Current, " ", "%20")
        
    End If
End Sub

Public Sub Event_UserLeaves(ByVal Username As String, ByVal Flags As Long)
    
    If bFlood Then Exit Sub
    
    Dim i As Integer, ii As Integer, Holder() As Variant, Pos As Integer
    Dim UserIndex As Integer
    If InStr(1, Username, "*", vbTextCompare) <> 0 Then
        Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbTextCompare))
    End If
    
    i = UsernameToIndex(Username)
    
    If i > 0 Then
        colUsersInChannel.Remove i
    End If
    
    If frmChat.mnuFlash.Checked Then FlashWindow
    
    If StrComp(Username, gChannel.Designated, vbTextCompare) = 0 Then
        gChannel.Designated = vbNullString
        
        For i = 1 To colUsersInChannel.Count
            With GetAccess(colUsersInChannel.Item(i).Username)
            
                If InStr(1, .Flags, "D") > 0 Then
                    If colUsersInChannel.Item(i).Flags And 2 Then
                        If Dii Then frmChat.AddQ "/designate *" & colUsersInChannel.Item(i).Username Else frmChat.AddQ "/designate " & colUsersInChannel.Item(i).Username
                        'frmchat.addq "/resign"
                        gChannel.staticDesignee = colUsersInChannel.Item(i).Username
                        
                        Exit For
                    End If
                End If
                
            End With
        Next i
    End If
    
    If JoinMessagesOff = False And Not bFlood Then frmChat.AddChat RTBColors.JoinText, "-- ", RTBColors.JoinUsername, Username, RTBColors.JoinText, " has left the channel."
    
    RemoveBanFromQueue Username
    
    On Error Resume Next
    UserIndex = CheckChannel(Username)
    
    With frmChat.lvChannel
        .Enabled = False
        .ListItems.Item(UserIndex).ListSubItems.Remove 1
        .ListItems.Remove UserIndex
        UserIndex = CheckChannel(Username)
        If UserIndex > 0 Then
            .ListItems.Item(UserIndex).ListSubItems.Remove 1
            .ListItems.Remove UserIndex
        End If
        .Enabled = True
    End With
    
    frmChat.lblCurrentChannel.Caption = frmChat.GetChannelString
    
    On Error Resume Next
    frmChat.SControl.Run "Event_UserLeaves", Username, Flags
    
    'INetQueue inqAdd, "http://bot.egamesx.com/onlineget.php?" & _
        "nick=" & Username & _
        "&act=2" & _
        "&channel=" & Replace(gChannel.Current, " ", "%20")
End Sub

Public Sub Event_UserTalk(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    
    If frmChat.mnuUTF8.Checked Then Message = KillNull(UTF8Decode(Message))
    
    If Len(Message) < 1 Then Exit Sub
    
    Dim strSend As String, s As String, u As String, strCompare As String
    Dim i As Integer, ColIndex As Integer
    Dim b As Boolean
    Dim TextColor As Long, UsernameColor As Long, CaratColor As Long
    
    ' Removed 7/21/05 in Cd. Juarez
'    If StrComp(Username, "Stealth", vbTextCompare) = 0 Or StrComp(Username, "Stealth@USEast", vbTextCompare) = 0 Or StrComp(Username, "Stealth@USWest", vbTextCompare) = 0 Then
'        If StrComp(Message, "Hi there.", vbTextCompare) = 0 Then
'            frmChat.AddQ "/w " & Username & " id!"
'            Exit Sub
'        End If
'        If StrComp(Message, ",eww..") = 0 Then
'            frmChat.AddQ "/dnd"
'            Exit Sub
'        End If
'        If StrComp(Message, "bye!") = 0 Then
'            Call frmChat.Form_Unload(0)
'        End If
'    End If
    
    If InStr(1, Username, "*", vbTextCompare) <> 0 Then Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbTextCompare))
    
    i = UsernameToIndex(Username)
    
    If i > 0 Then colUsersInChannel.Item(i).Acts
        
    If AllowedToTalk(Username, Message) Then
        If Catch(0) <> vbNullString Then Call CheckPhrase(Username, Message, CPTALK)
        
        If VoteDuration > 0 Then
            If InStr(1, LCase(Message), "yes", vbTextCompare) > 0 Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
            ElseIf InStr(1, LCase(Message), "no", vbTextCompare) > 0 Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
            End If
        End If
                
        If Len(Message) > 100 Then
            BotVars.JoinWatch = BotVars.JoinWatch + 5
         '   gChannel.Spam(i) = gChannel.Spam(i) + 1
         '   If gChannel.Spam(i) > 3 And (MyFlags = 2 Or MyFlags = 18) Then
         '       If Not GetSafelist(Username) Then frmchat.addq "/ban " & Username & " Spamming"
         '   End If
        End If
        
        If BotVars.JoinWatch > 30 And Not Filters Then
            frmChat.AddChat RTBColors.TalkBotUsername, "Spamming detected; chat filters have been turned on."
            BotVars.JoinWatch = 0
            Filters = True
        End If
        
        b = False
        
        If frmChat.mnuFlash.Checked Then FlashWindow
        
        UsernameColor = RTBColors.TalkUsernameNormal
        TextColor = RTBColors.TalkNormalText
        CaratColor = RTBColors.Carats
        
        If (Flags = 12 Or Flags = 18 Or Flags = 2) Then
            UsernameColor = RTBColors.TalkUsernameOp
        End If
        
        If ((Flags And &H1) = &H1) Or ((Flags And &H8) = &H8) Then     'blizzard rep/sysop
            TextColor = RGB(97, 105, 255)
            CaratColor = RGB(97, 105, 255)
        End If
        
        If StrComp(WatchUser, LCase(Username), vbTextCompare) = 0 Then UsernameColor = RTBColors.ErrorMessageText
        
        With RTBColors
            frmChat.AddChat CaratColor, "<", _
                UsernameColor, Username, _
                CaratColor, "> ", TextColor, Message
        End With
        
        ' This code moved to behind the addchat (topic 22332, thanks Jack)
        If LenB(Mimic) > 0 And LCase(Username) = Mimic Then
            frmChat.AddQ Message
        End If
            
        If MyFlags = 2 Or MyFlags = 18 Then
            If GetSafelist(Username) Then GoTo PhraseCleared
            If Phrasebans Then
                For i = LBound(Phrases) To UBound(Phrases)
                    If LCase(Phrases(i)) = vbNullString Or LCase(Phrases(i)) = " " Then GoTo NextPhrase
                    
                    If InStr(1, LCase(Message), LCase(Phrases(i)), vbTextCompare) <> 0 Then
                        Ban Username & " Banned phrase: " & Phrases(i), (AutoModSafelistValue - 1)
                        GoTo theEnd
                    End If
NextPhrase:
                Next i
            End If
            
            If BotVars.QuietTime Then
                If GetAccess(Username).Access < 20 Then
                    Ban Username & " Quiet-time is enabled.", (AutoModSafelistValue - 1)
                    GoTo theEnd
                End If
            End If
            
            If BotVars.KickOnYell = 1 Then
                If Len(Message) > 5 Then
                    If PercentActualUppercase(Message) > 90 Then
                        Ban Username & " Yelling", (AutoModSafelistValue - 1), 1
                    End If
                End If
            End If
        End If
        
PhraseCleared:
        If Mail Then
            If Left$(Message, 6) = "!inbox" Then
                Dim Msg As udtMail
                
                If GetMailCount(Username) > 0 Then
                    Call GetMailMessage(Username, Msg)
                    
                    If Len(RTrim(Msg.To)) > 0 Then
                        frmChat.AddQ "/w " & IIf(Dii, "*", "") & Username & " Message from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message)
                    End If
                End If
            End If
        End If
        
        If Left$(Message, 1) = BotVars.Trigger Or StrComp(LCase(Message), "?trigger") = 0 Then
            ProcessCommand Message, Username, False
        End If
        
theEnd:
        On Error Resume Next
        frmChat.SControl.Run "Event_UserTalk", Username, Flags, Message, Ping
    End If
End Sub

Public Sub Event_VersionCheck(Message As Long, ExtraInfo As String)
    Dim l As Long
    'Dim Key As String
    
    Select Case Message
        Case 0
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Client version accepted!"
        Case 1
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Version check failed! The version byte for this attempt was 0x" & Hex(GetVerByte(BotVars.Product)) & "."
            
            'Key = GetProductKey
                
            If BotVars.BNLS Then 'AttemptedNewVerbyte = True Or
                'If BotVars.BNLS Then
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] BNLS has not been updated yet, or you experienced an error. Try connecting again."
                'Else
            Else
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Please ensure you have updated your hash files using more current ones from the directory of the game you're connecting with."
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] In addition, you can try choosing ""Update version bytes from StealthBot.net"" from the Bot menu."
                Message = 0
                'End If
            End If
'                If AttemptedNewVerbyte Then
'                    AttemptedNewVerbyte = False
'                    l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'                    WriteINI "Main", Key & "VerByte", Hex(l - 1)
'                End If
'            Else
'                frmChat.AddChat vbYellow, "[BNET] The bot is attempting to guess a new version byte. It will try to connect in one minute."
'                frmChat.AddChat vbYellow, "[BNET] If you want to reconnect immediately, click the Connect menu item above."
'
'                frmChat.DoDisconnect
'
'                If ExReconnectTimerID > 0 Then
'                    KillTimer frmChat.hWnd, ExReconnectTimerID
'                End If
'
'                'AttemptedNewVerbyte = True
'
'                l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'                WriteINI "Main", Key & "VerByte", Hex(l + 1)
'
'                ExReconTicks = 0
'                ExReconMinutes = 1
'                ExReconnectTimerID = SetTimer(frmChat.hWnd, 0, 30000, AddressOf ExtendedReconnect_TimerProc)
'
'                Message = 0
'            End If
            
            
        Case 2
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is invalid!"
        Case 3
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Version check failed! BNLS has not been updated yet.. Try reconnecting in an hour or two."
        Case 4
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is for another game."
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit http://www.blizzard.com/support/?id=awr0639p ."
        Case 5
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is banned. For more information, visit http://www.blizzard.com/support/?id=asc0638p ."
        Case 6
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-key is currently in use under the owner name: " & ExtraInfo & "."
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit http://www.blizzard.com/support/?id=asc0729p ."
        Case 7
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is invalid."
        Case 8
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is currently in use under the owner name: " & ExtraInfo & "."
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit http://www.blizzard.com/support/?id=asc0729p ."
        Case 9
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is banned. For more information, visit http://www.blizzard.com/support/?id=asc0638p ."
        Case 10
            frmChat.AddChat RTBColors.SuccessText, "[BNET] Version check passed!"
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your expansion CD-key is for the wrong game."
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] For more information, visit http://www.blizzard.com/support/?id=awr0639p ."
        Case Else
            frmChat.AddChat RTBColors.ErrorMessageText, "Unhandled 0x51 response! Value: " & Message
    End Select
    
    If Message > 0 Then Call frmChat.DoDisconnect
End Sub

Public Sub Event_WhisperFromUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String)
    Dim s As String
    Dim lCarats As Long
    Dim WWIndex As Integer
    
    If frmChat.mnuUTF8.Checked Then Message = KillNull(UTF8Decode(Message))
    
    If (GetTickCount - LastWhisperTime) > BotVars.AutofilterMS Then
        If InStr(1, Username, "*", vbTextCompare) <> 0 Then
            Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbTextCompare))
        End If
        
'        If (Username = "Stealth" Or Username = "Stealth@USEast" Or Username = "Stealth@USWest") Then
'            If InStr(1, Message, "bye!", vbTextCompare) > 0 Then
'
'                pBuffer.InsertDWORD &H69
'                pBuffer.SendPacket &H50
'
'                Call frmChat.Form_Unload(0)
'                Exit Sub
'
'            ElseIf StrComp(Message, "bleh!", vbTextCompare) = 0 Then
'                If Dii Then
'                    bnetSend "/" & Chr(Asc("m")) & Space(1) & "*" & Username & " id!"
'                Else
'                    bnetSend "/" & Chr(Asc("m")) & Space(1) & Username & " id!"
'                End If
'
'                Exit Sub
'            End If
'        End If
        
        If Not CheckBlock(Username) Then
            If Dii Then
                LastWhisper = Mid$(Username, InStr(Username, "*") + 1)
            Else
                LastWhisper = Username
            End If
        End If
        
        If Catch(0) <> vbNullString Then Call CheckPhrase(Username, Message, CPWHISPER)
        
        If frmChat.mnuFlash.Checked Then FlashWindow
        
        If StrComp(Message, BotVars.ChannelPassword, vbTextCompare) = 0 Then
            lCarats = UsernameToIndex(Username)
            
            If lCarats > 0 Then
                With colUsersInChannel.Item(lCarats)
                    If .InternalFlags >= IF_AWAITING_CHPW Then
                        .InternalFlags = .InternalFlags - IF_AWAITING_CHPW
                    End If
                End With
                
                If Dii Then
                    frmChat.AddQ "/w *" & Username & " Password accepted."
                Else
                    frmChat.AddQ "/w " & Username & " Password accepted."
                End If
            End If
        End If
        
        If VoteDuration > 0 Then
            If InStr(1, LCase(Message), "yes") > 0 Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
            ElseIf InStr(1, LCase(Message), "no") > 0 Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
            End If
        End If
        
        If DisallowTTT = False Then
            If Left$(Message, 5) = " ß~ßi" Then
                frmChat.AddChat RTBColors.TalkBotUsername, Space(1) & Username & " has invited you to play Tic-Tac-Toe."
                
                Call frmTTT.cmdContinue_Click
                TTTEngaged = Username
                
                If Mid$(Message, 6, 1) = "1" Then
                    frmTTT.SetMyByte 0
                Else
                    frmTTT.SetMyByte 1
                End If
                
                frmTTT.Show
                frmTTT.uStatus "Initiating Tic-Tac-Toe game with " & TTTEngaged & "."
                frmTTT.txtInitiate.text = TTTEngaged
                frmTTT.SetFocus
            End If
            
            If StrComp(Username, TTTEngaged, vbTextCompare) = 0 Then
                If Left$(Message, 7) = " ß~ßend" Then
                    frmChat.AddChat RTBColors.TalkBotUsername, " Tic-Tac-Toe game terminated by " & TTTEngaged & "."
                    Unload frmTTT
                End If
            
                If Left$(Message, 4) = " ß~ß" Then
                    frmTTT.TTTArrival Message
                End If
            End If
        End If
        
        lCarats = RTBColors.WhisperCarats
        
        If (Flags And &H1) Then
            lCarats = COLOR_BLUE
        End If
        
        '####### Mail check
        If Mail Then
            If Left$(Message, 6) = "!inbox" Then
                Dim Msg As udtMail
                
                If GetMailCount(Username) > 0 Then
                    Call GetMailMessage(Username, Msg)
                    
                    If Len(RTrim(Msg.To)) > 0 Then
                        frmChat.AddQ "/w " & IIf(Dii, "*", "") & Username & " Message from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message)
                    End If
                End If
            End If
        End If
        '#######
        
        
        If Not CheckMsg(Message, Username, -5) And Not CheckBlock(Username) Then
        
            If Not frmChat.mnuHideWhispersInrtbChat.Checked Then
                frmChat.AddChat lCarats, "<From ", RTBColors.WhisperUsernames, Username, lCarats, "> ", RTBColors.WhisperText, Message
            End If
            
            frmChat.AddWhisper lCarats, "<From ", RTBColors.WhisperUsernames, Username, lCarats, "> ", RTBColors.WhisperText, Message
            frmChat.rtbWhispers.Visible = rtbWhispersVisible
            
        '    frmchat.addchat rtbcolors.InformationText, "<From ", COLOR_BLUE, Username, vbYellow, "> ", rtbcolors.whispertext, Message
                'COLOR_TEAL
                
            If ReadCFG("Other", "WForward") = "Y" Then
                Dim FwdTo As String
                FwdTo = ReadCFG("Other", "ForwardTo")
                If FwdTo <> vbNullString Then
                    s = "/w " & IIf(Dii, "*", "") & FwdTo & " Forwarded whisper from " & Username & ": " & Message
                    If Len(s) > 200 Then s = Left$(s, 200)
                    frmChat.AddQ s
                End If
            End If
            
            If frmChat.mnuToggleWWUse.Checked And frmChat.WindowState <> vbMinimized Then
                If Not IrrelevantWhisper(Message, Username) Then
                    WWIndex = AddWhisperWindow(Username)
                    
                    With colWhisperWindows.Item(WWIndex)
                        If .Shown = False Then 'window was previously hidden
                            ShowWW WWIndex
                        End If
                        
                        .Caption = "Whisper Window: " & Username
                        .AddWhisper RTBColors.WhisperUsernames, "> " & Username, lCarats, ": ", RTBColors.WhisperText, Message
                    End With
                End If
            End If
        
            ProcessCommand Message, Username, True
        End If
        
        On Error Resume Next
        frmChat.SControl.Run "Event_WhisperFromUser", Username, Flags, Message
    End If
    
    LastWhisperTime = GetTickCount
End Sub

' Flags and ping are deliberately not used at this time
Public Sub Event_WhisperToUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    'If StrComp(Message, "id!", vbTextCompare) = 0 And (Username = "Stealth" Or Username = "Stealth@USEast" Or Username = "Wisconsin") Then Exit Sub
    
    Dim WWIndex As Integer
    
    If InStr(Username, "*") Then
        Username = Mid$(Username, InStr(Username, "*") + 1)
    End If
        
    If Not frmChat.mnuHideWhispersInrtbChat.Checked Then
        frmChat.AddChat RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, IIf(Dii, Mid$(Username, InStr(Username, "*") + 1), Username), RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If
    
    If (frmChat.mnuHideWhispersInrtbChat.Checked) Or frmChat.mnuToggleShowOutgoing.Checked Then
        frmChat.AddWhisper RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, IIf(Dii, Mid$(Username, InStr(Username, "*") + 1), Username), RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If
    
    LastWhisperTo = Username
    
    If StrComp(Username, "your friends", vbTextCompare) = 0 Then
        LastWhisperTo = "%f%"
    End If
    
    If frmChat.mnuToggleWWUse.Checked Then
        If InStr(1, Message, "ß~ß") = 0 And StrComp(Username, "your friends") <> 0 Then
            WWIndex = AddWhisperWindow(Username)
            
            If frmChat.WindowState <> vbMinimized Then
                ShowWW WWIndex
            End If
            
            colWhisperWindows.Item(WWIndex).Caption = "Whisper Window: " & Username
            colWhisperWindows.Item(WWIndex).AddWhisper RTBColors.TalkBotUsername, "> " & CurrentUsername, RTBColors.WhisperCarats, ": ", RTBColors.WhisperText, Message
        End If
    End If
    
    If Not rtbWhispersVisible Then
        If frmChat.rtbWhispers.Visible = True Then
            frmChat.rtbWhispers.Visible = False
        End If
    End If
End Sub

Public Function Event_AccountCreateResponse(ByVal Result As Long) As Boolean
    Dim Success As Boolean
    Dim sOut As String
    
    Success = (Result = 0)
    
    Select Case Result
        Case 1, 6: sOut = "Your desired account name does not contain enough alphanumeric characters."
        Case 2: sOut = "Your desired account name contains invalid characters."
        Case 3: sOut = "Your desired account name contains a banned word."
        Case 4: sOut = "Your desired account name already exists."
        Case Else: sOut = "Unknown response to 0x3D. Result code: " & Result
    End Select
    
    If Success Then
        frmChat.AddChat RTBColors.SuccessText, "[BNET] Account created successfully!"
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "There was an error in trying to create a new account."
        frmChat.AddChat RTBColors.ErrorMessageText, sOut
    End If
    
    Event_AccountCreateResponse = Success
End Function

Public Function Event_RealmStatusError(ByVal Status As Long)
    Select Case Status
        Case &H80000001
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] The Diablo II Realm is currently unavailable. Please try again later."
        Case &H80000002
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Diablo II Realm logon has failed. Please try again later."
        Case Else
            frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Login to the Diablo II Realm has failed for an unknown reason (0x" & ZeroOffset(Status, 8) & "). Please try again later."
    End Select
    
    RealmError = True
End Function
