Attribute VB_Name = "modChatQueue"
' ...

Option Explicit

' ...
Public colChatQueue As Collection

' ...
Private m_TimerID   As Long

' ...
Public Sub ChatQueue_Initialize()
    Set colChatQueue = New Collection

    m_TimerID = SetTimer(frmChat.hWnd, 0, 1000, AddressOf ChatQueueTimerProc)
End Sub

' ...
Public Sub ChatQueue_Terminate()
    Set colChatQueue = Nothing

    m_TimerID = KillTimer(frmChat.hWnd, m_TimerID)
End Sub

' ...
Public Function ChatQueueTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal idEvent As Long, ByVal dwTimer As Long)
    
    Dim i    As Integer ' ...
    
    ' ...
    For i = 1 To colChatQueue.Count
        ' ...
        Dim clsChatQueue As clsChatQueue
        
        ' ...
        Set clsChatQueue = colChatQueue(i)
    
        With clsChatQueue
            ' ...
            If (GetTickCount() - .Time() >= 3000) Then
                ' ...
                Call .Show
                
                ' ...
                Call colChatQueue.Remove(i)
            End If
        End With
    Next
End Function

' CHATQUEUE EVENTS

' ...
Public Sub Event_QueuedJoin(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
    ByVal w3icon As String)
    
    Dim Game   As String ' ...
    Dim pStats As String ' ...
    Dim Clan   As String ' ...
    
    Game = ParseStatstring(OriginalStatstring, pStats, Clan)

    If (Not (JoinMessagesOff)) Then
        frmChat.AddChat RTBColors.JoinText, "-- ", _
            RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                RTBColors.JoinText, " has joined the channel using " & pStats
    End If
    
    If (Dii) Then
        If (Not (checkChannel(Username) <> 0)) Then
            Call AddName(Username, Product, Flags, Ping, Clan)
        End If
    Else
        If (Len(Clan)) Then
            Call AddName(Username, Product, Flags, Ping, Clan)
        Else
            Call AddName(Username, Product, Flags, Ping)
        End If
    End If
    
    frmChat.lblCurrentChannel.Caption = _
        frmChat.GetChannelString()
End Sub

' ...
Public Sub Event_QueuedUserInChannel(ByVal Username As String, ByVal Flags As Long, _
    ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
    ByVal w3icon As String)
    
    Dim i      As Integer ' ...
    Dim Game   As String  ' ...
    Dim pStats As String  ' ...
    Dim Clan   As String  ' ...

    i = UsernameToIndex(Username)
    
    colUsersInChannel.Item(i).Statstring = _
        OriginalStatstring
        
    Game = ParseStatstring(OriginalStatstring, pStats, Clan)
    
    If (JoinMessagesOff = False) Then
        frmChat.AddChat RTBColors.JoinText, "-- Stats updated: ", _
            RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                RTBColors.JoinText, " is using " & pStats
    End If
End Sub

' ...
Public Sub Event_QueuedStatusUpdate(ByVal Username As String, ByVal Flags As Long, _
    ByVal prevflags As Long, ByVal Ping As Long, ByVal Product As String, ByVal sClan As String, _
    ByVal OriginalStatstring As String, ByVal w3icon As String)
    
    Dim found      As ListItem ' ...
    
    Dim i          As Integer  ' ...
    Dim Pos        As Integer  ' ...
    Dim squelching As Boolean  ' ...
    
    i = UsernameToIndex(Username)
    
    Pos = checkChannel(Username)
    
    If (((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) And _
        ((prevflags And USER_CHANNELOP&) <> USER_CHANNELOP&)) Then

        frmChat.AddChat RTBColors.JoinedChannelText, "-- ", _
            RTBColors.JoinedChannelName, Username, RTBColors.JoinedChannelText, _
                " has acquired ops."
                
        Call frmChat.lvChannel.ListItems.Remove(Pos)
    
        Call AddName(Username, colUsersInChannel.Item(i).Product, Flags, Ping)
    End If
    
    If (StrComp(gChannel.Current, "The Void", vbBinaryCompare) = 0) Then
        If (Not (frmChat.mnuDisableVoidView.Checked)) Then
            If (frmChat.lvChannel.ListItems.Count < 200) Then
                If (Pos) Then
                    Call AddName(Username, Product, Flags, Ping)
                End If
            End If
        End If
    
        Exit Sub
    End If

    ' User is being squelched?
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
        ' User is being unsquelched?
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

    If (Pos) Then
        Set found = frmChat.lvChannel.ListItems(Pos)
        
        If (g_ThisIconCode <> -1) Then
            If ((Not (squelching)) And _
                (Not (Unsquelching))) Then
                
                If (colUsersInChannel.Item(i).Product = "W3XP") Then
                    found.SmallIcon = (g_ThisIconCode + ICON_START_W3XP + _
                        IIf(g_ThisIconCode + ICON_START_W3XP = ICSCSW, 1, 0))
                Else
                    found.SmallIcon = (g_ThisIconCode + ICON_START_WAR3)
                End If
            End If
        End If
    
        Set found = Nothing
    End If
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
    
    frmChat.AddChat CaratColor, "<", UsernameColor, Username, _
        CaratColor, "> ", TextColor, Message
End Sub

' ...
Public Sub Event_QueuedEmote(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    Call frmChat.AddChat(vbRed, "Event_QueuedEmote() has been fired.")
End Sub
