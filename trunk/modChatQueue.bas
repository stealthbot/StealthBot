Attribute VB_Name = "modChatQueue"
' ...

Option Explicit

' ...
Public colChatQueue  As Collection

' ...
Private m_TimerID    As Long
Private m_QueueCount As Long
Private m_QueueGTC   As Long

' ...
Public Sub ChatQueue_Initialize()
    ' ...
    If (BotVars.ChatDelay) Then
        ' ...
        Set colChatQueue = New Collection
    
        ' ...
        m_TimerID = SetTimer(0, m_TimerID, BotVars.ChatDelay, _
            AddressOf ChatQueueTimerProc)
    End If
End Sub

' ...
Public Sub ChatQueue_Terminate()
    ' ...
    If (m_TimerID) Then
        ' ...
        m_TimerID = KillTimer(0, m_TimerID)
        
        ' ...
        Set colChatQueue = Nothing
    End If
End Sub

' ...
Public Function ChatQueueTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
    ByVal dwTimer As Long)
    
    Dim i       As Integer ' ...
    Dim doLoop  As Boolean ' ...
    Dim found   As Boolean ' ...
    Dim blnShow As Boolean ' ...
    
    ' ...
    m_QueueCount = colChatQueue.Count
    
    ' ...
    doLoop = True
    
    ' ...
    Do While (doLoop)
        ' ...
        For i = 1 To colChatQueue.Count
            ' ...
            Dim clsChatQueue As clsChatQueue
            
            ' ...
            Set clsChatQueue = colChatQueue(i)

            With clsChatQueue
                ' ...
                If (GetTickCount() - .Time() >= 300) Then
                    blnShow = True
                End If

                ' ...
                If (blnShow) Then
                    ' ...
                    Call .Show
                    
                    ' ...
                    Call colChatQueue.Remove(i)
                End If
            End With
            
            ' ...
            blnShow = False
        Next
        
        ' ...
        If (found) Then
            doLoop = True
        Else
            doLoop = False
        End If
        
        ' ...
        found = False
    Loop
End Function

' CHATQUEUE EVENTS

' ...
Public Sub Event_QueuedJoin(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
    Dim game   As String ' ...
    Dim pStats As String ' ...
    Dim Clan   As String ' ...
    
    game = ParseStatstring(OriginalStatstring, pStats, Clan)

    If (JoinMessagesOff = False) Then
        Call frmChat.AddChat(RTBColors.JoinText, "-- ", _
            RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                RTBColors.JoinText, " has joined the channel using " & pStats)
    End If
    
    If (Clan <> vbNullString) Then
        Call AddName(Username, Product, Flags, Ping, Clan)
    Else
        Call AddName(Username, Product, Flags, Ping)
    End If
    
    frmChat.lblCurrentChannel.Caption = _
        frmChat.GetChannelString
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    'frmChat.SControl.Run "Event_UserJoins", Username, flags, Message, Ping, _
    '    Level, OriginalStatstring, Banned
    
    frmChat.SControl.Run "Event_UserJoins", Username, Flags, OriginalStatstring, Ping, _
        Product, 0, OriginalStatstring, False
End Sub

' ...
Public Sub Event_QueuedUserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
    Dim found  As ListItem ' ...
    
    Dim i      As Integer  ' ...
    Dim game   As String   ' ...
    Dim pStats As String   ' ...
    Dim Clan   As String   ' ...
    Dim Pos    As Integer  ' ...

    i = UsernameToIndex(Username)
    
    Pos = checkChannel(Username)

    game = ParseStatstring(OriginalStatstring, pStats, Clan)
    
    If (JoinMessagesOff = False) Then
        Call frmChat.AddChat(RTBColors.JoinText, "-- Stats updated: ", _
            RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                RTBColors.JoinText, " is using " & pStats)
    End If
    
    If (Flags = &H0) Then
        If (Pos) Then
            Set found = frmChat.lvChannel.ListItems(Pos)
            
            If (g_ThisIconCode <> -1) Then
                If (colUsersInChannel.Item(i).Product = "WAR3") Then
                    If (g_ThisIconCode = ICON_START_WAR3) Then
                        found.SmallIcon = (g_ThisIconCode + ICON_START_WAR3)
                    End If
                ElseIf (colUsersInChannel.Item(i).Product = "W3XP") Then
                    If (g_ThisIconCode = ICON_START_W3XP) Then
                        found.SmallIcon = (g_ThisIconCode + ICON_START_W3XP + _
                            IIf(g_ThisIconCode + ICON_START_W3XP = ICSCSW, 1, 0))
                    End If
                End If
            End If
        
            Set found = Nothing
        End If
    End If
End Sub

' ...
Public Sub Event_QueuedStatusUpdate(ByVal Username As String, ByVal Flags As Long, ByVal prevflags As Long, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
        ByVal w3icon As String)
    
    Dim found      As ListItem ' ...
    
    Dim i          As Integer  ' ...
    Dim Pos        As Integer  ' ...
    Dim squelching As Boolean  ' ...
    
    i = UsernameToIndex(Username)
    
    Pos = checkChannel(Username)
    
    If (((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) And _
        ((prevflags And USER_CHANNELOP&) <> USER_CHANNELOP&)) Then

        Call frmChat.AddChat(RTBColors.JoinedChannelText, "-- ", _
            RTBColors.JoinedChannelName, Username, RTBColors.JoinedChannelText, _
                " has acquired ops.")
                
        Call frmChat.lvChannel.ListItems.Remove(Pos)
    
        Call AddName(Username, colUsersInChannel.Item(i).Product, Flags, Ping)
    End If
    
    If (StrComp(gChannel.Current, "The Void", vbBinaryCompare) = 0) Then
        If (Not (frmChat.mnuDisableVoidView.Checked)) Then
            If (frmChat.lvChannel.ListItems.Count < 200) Then
                If (Not (Pos)) Then
                    frmChat.lvChannel.Enabled = False
                
                    Call AddName(Username, Product, Flags, Ping)
                    
                    frmChat.lvChannel.Enabled = True
                End If
            End If
        End If
    
        Exit Sub
    End If

    ' is user being squelched?
    If ((Flags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
        squelching = True
    
        If (Pos) Then
            With colUsersInChannel(i)
                frmChat.lvChannel.Enabled = False
    
                Call frmChat.lvChannel.ListItems.Remove(Pos)
    
                Call AddName(.Username, .Product, Flags, Ping, .Clan, Pos)
    
                frmChat.lvChannel.Enabled = True
            End With
        End If
    Else
        ' is user being unsquelched?
        If (Pos) Then
            If ((prevflags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
                With colUsersInChannel(i)
                    frmChat.lvChannel.Enabled = False
    
                    Call frmChat.lvChannel.ListItems.Remove(Pos)
    
                    Call AddName(.Username, .Product, Flags, Ping, .Clan, Pos)
                    
                    frmChat.lvChannel.Enabled = True
                End With
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next

    frmChat.SControl.Run "Event_FlagUpdate", Username, Flags, Ping
End Sub

' ...
Public Sub Event_QueuedTalk(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    Dim UsernameColor As Long ' ...
    Dim TextColor     As Long ' ...
    Dim CaratColor    As Long ' ...
    
    If (StrComp(WatchUser, Username, vbTextCompare) = 0) Then
        UsernameColor = RTBColors.ErrorMessageText
    ElseIf ((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        UsernameColor = RTBColors.TalkUsernameOp
    Else
        UsernameColor = RTBColors.TalkUsernameNormal
    End If
    
    If (((Flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or _
        ((Flags And USER_SYSOP&) = USER_SYSOP&)) Then
       
        TextColor = RGB(97, 105, 255)
        CaratColor = RGB(97, 105, 255)
    Else
        TextColor = RTBColors.TalkNormalText
        CaratColor = RTBColors.Carats
    End If
    
    Call frmChat.AddChat(CaratColor, "<", UsernameColor, Username, _
        CaratColor, "> ", TextColor, Message)
        
    ' scripts
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
            vbBinaryCompare) = 0) Then
            
            Message = BotVars.TriggerLong & _
                Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
            vbBinaryCompare) = 0) Then
            
            Message = BotVars.Trigger & _
                Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If

    frmChat.SControl.Run "Event_UserTalk", Username, Flags, Message, Ping
End Sub

' ...
Public Sub Event_QueuedEmote(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    frmChat.AddChat RTBColors.EmoteText, "<", RTBColors.EmoteUsernames, _
        Username & Space(1), RTBColors.EmoteText, Message & ">"
        
    If (frmChat.mnuFlash.Checked) Then
        Call FlashWindow
    End If
    
    ' scripts
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
            vbBinaryCompare) = 0) Then
            
            Message = BotVars.TriggerLong & _
                Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
            vbBinaryCompare) = 0) Then
            
            Message = BotVars.Trigger & _
                Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If
    
    frmChat.SControl.Run "Event_UserEmote", Username, Flags, Message
End Sub

' ...
Public Sub Event_DroppedJoin(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
End Sub

' ...
Public Sub Event_DroppedUserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
End Sub

' ...
Public Sub Event_DroppedStatusUpdate(ByVal Username As String, ByVal Flags As Long, ByVal prevflags As Long, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
        ByVal w3icon As String)
    
End Sub

' ...
Public Sub Event_DroppedTalk(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
End Sub

' ...
Public Sub Event_DroppedEmote(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
End Sub

' ...
Public Sub ClearChatQueue()
    Set colChatQueue = New Collection
End Sub
