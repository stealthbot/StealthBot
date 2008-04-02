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
    If (BotVars.ChatDelay > 0) Then
        ' ...
        Set colChatQueue = New Collection
    
        ' ...
        m_TimerID = SetTimer(0, m_TimerID, _
            IIf(BotVars.ChatDelay <= 500, BotVars.ChatDelay, 500), AddressOf ChatQueueTimerProc)
    End If
End Sub

' ...
Public Sub ChatQueue_Terminate()
    ' ...
    If (m_TimerID > 0) Then
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
    Dim blnShow As Boolean ' ...

    ' ...
    Do
        ' ...
        blnShow = False
    
        ' ...
        For i = 1 To colChatQueue.Count
            ' ...
            Dim clsChatQueue As clsChatQueue
            
            ' ...
            Set clsChatQueue = colChatQueue(i)

            ' ...
            If (GetTickCount() - clsChatQueue.Time() >= BotVars.ChatDelay) Then
                blnShow = True
            End If

            ' ...
            If (blnShow) Then
                ' ...
                clsChatQueue.Show
                
                ' ...
                colChatQueue.Remove i
                
                ' ...
                Exit For
            End If
        Next i
    Loop While (blnShow)
End Function

' CHATQUEUE EVENTS

' ...
Public Sub Event_QueuedJoin(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
    On Error GoTo ERROR_HANDLER
    
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
    
    ' ...
    frmChat.SControl.Run "Event_UserJoins", Username, Flags, pStats, Ping, _
        Product, 0, OriginalStatstring, False
        
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in Event_QueuedUserJoin().")
    
    Exit Sub
End Sub

' ...
Public Sub Event_QueuedUserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, ByVal w3icon As String)
    
    Dim found  As ListItem ' ...
    
    Dim i      As Integer  ' ...
    Dim Stats  As String   ' ...
    Dim Clan   As String   ' ...
    Dim Pos    As Integer  ' ...

    ' ...
    ParseStatstring OriginalStatstring, Stats, Clan
    
    ' ...
    If (JoinMessagesOff = False) Then
        ' ...
        Call frmChat.AddChat(RTBColors.JoinText, "-- Stats updated: ", _
                RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                        RTBColors.JoinText, " is using " & Stats)
    End If
    
    ' ...
    If (BotVars.ShowStatsIcons) Then
        ' ...
        Pos = checkChannel(Username)
        
        ' ...
        If (Pos) Then
            ' ...
            Set found = frmChat.lvChannel.ListItems(Pos)
            
            ' ...
            i = UsernameToIndex(Username)
            
            ' ...
            If (g_ThisIconCode <> -1) Then
                ' ...
                If (colUsersInChannel.Item(i).Product = "WAR3") Then
                    ' ...
                    If (found.SmallIcon = ICWAR3) Then
                        ' ...
                        found.SmallIcon = (g_ThisIconCode + ICON_START_WAR3)
                    End If
                ElseIf (colUsersInChannel.Item(i).Product = "W3XP") Then
                    ' ...
                    If (found.SmallIcon = ICWAR3X) Then
                        ' ...
                        found.SmallIcon = ((g_ThisIconCode + ICON_START_W3XP) + IIf(g_ThisIconCode + _
                                ICON_START_W3XP = ICSCSW, 1, 0))
                    End If
                End If
            End If
            
            ' ...
            Set found = Nothing
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in Event_QueuedUserInChannel().")
    
    Exit Sub
End Sub

' ...
Public Sub Event_QueuedStatusUpdate(ByVal Username As String, ByVal Flags As Long, ByVal prevflags As Long, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
        ByVal w3icon As String)
    
    On Error GoTo ERROR_HANDLER
    
    Dim Pos      As Integer  ' ...
    Dim doUpdate As Boolean  ' ...
    
    ' are we in the void?
    If (g_Channel.IsSilent) Then
        ' ...
        If (frmChat.mnuDisableVoidView.Checked = False) Then
            'Call AddName(Username, Product, Flags, Ping)
        End If
    
        Exit Sub
    Else
        ' ...
        Pos = checkChannel(Username)
    
        ' is user being designated?
        If (((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) And _
                ((prevflags And USER_CHANNELOP&) <> USER_CHANNELOP&)) Then
    
            ' ...
            frmChat.lvChannel.ListItems.Remove Pos
            
            ' ...
            Call AddName(Username, Product, Flags, Ping, sClan)
        
            ' ...
            Call frmChat.AddChat(RTBColors.JoinedChannelText, "-- ", _
                    RTBColors.JoinedChannelName, Username, RTBColors.JoinedChannelText, _
                            " has acquired ops.")
                            
            ' ...
            Exit Sub
        End If
        
        ' is user being squelched?
        If (((Flags And USER_SQUELCHED&) = USER_SQUELCHED&) And _
                ((prevflags And USER_SQUELCHED&) <> USER_SQUELCHED&)) Then
        
            ' ...
            doUpdate = True
        Else
            ' is user being unsquelched?
            If ((prevflags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
                ' ...
                doUpdate = True
            End If
        End If
    
        ' ...
        If (doUpdate = True) Then
            ' ...
            If (Pos) Then
                ' ...
                frmChat.lvChannel.ListItems.Remove Pos
                
                ' ...
                Call AddName(Username, Product, Flags, Ping, sClan, Pos)
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    ' ...
    frmChat.SControl.Run "Event_FlagUpdate", Username, Flags, Ping
    
    Exit Sub
    
ERROR_HANDLER:
    MsgBox "Error: " & Err.description & " in Event_QueuedStatusUpdate()."
    
    Exit Sub
End Sub

' ...
Public Sub Event_QueuedTalk(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    Dim UsernameColor As Long ' ...
    Dim TextColor     As Long ' ...
    Dim CaratColor    As Long ' ...
    
    ' ...
    If (AllowedToTalk(Username, Message)) Then
        If (Catch(0) <> vbNullString) Then
            Call CheckPhrase(Username, Message, CPTALK)
        End If
    
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
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    ' ...
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        ' ...
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
                vbBinaryCompare) = 0) Then
            
            ' ...
            Message = BotVars.Trigger & Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If

    ' ...
    frmChat.SControl.Run "Event_UserTalk", Username, Flags, Message, Ping
End Sub

' ...
Public Sub Event_QueuedEmote(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    ' ...
    If (AllowedToTalk(Username, Message)) Then
        ' ...
        frmChat.AddChat RTBColors.EmoteText, "<", RTBColors.EmoteUsernames, _
            Username & Space(1), RTBColors.EmoteText, Message & ">"
        
        ' ...
        If (frmChat.mnuFlash.Checked) Then
            Call FlashWindow
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    ' ...
    If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
        If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
            vbBinaryCompare) = 0) Then
            
            Message = BotVars.Trigger & Mid$(Message, Len(BotVars.TriggerLong) + 1)
        End If
    End If
    
    ' ...
    g_lastQueueUser = Username
    
    ' ...
    frmChat.SControl.Run "Event_UserEmote", Username, Flags, Message
End Sub

' ...
Public Sub ClearChatQueue()
    Set colChatQueue = New Collection
End Sub
