Attribute VB_Name = "modChatQueue"
' ...

Option Explicit

' ...
Public colChatQueue  As Collection

' ...
Private m_TimerID    As Long
Private m_QueueCount As Long

' ...
Public Sub ChatQueue_Initialize()
    Set colChatQueue = New Collection

    m_TimerID = SetTimer(0, m_TimerID, 1000, AddressOf ChatQueueTimerProc)
End Sub

' ...
Public Sub ChatQueue_Terminate()
    m_TimerID = KillTimer(0, m_TimerID)
    
    Set colChatQueue = Nothing
End Sub

' ...
Public Function ChatQueueTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal idEvent As Long, ByVal dwTimer As Long)
    
    Dim i      As Integer ' ...
    Dim doLoop As Boolean ' ...
    Dim found  As Boolean ' ...
        
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
                If (GetTickCount() - .Time() >= 3000) Then
                    ' ...
                    Call .Show
                    
                    ' ...
                    Call colChatQueue.Remove(i)
                    
                    ' ...
                    found = True
                    
                    ' ...
                    Exit For
                End If
            End With
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
    ByVal Product As String, ByVal sClan As String, ByVal OriginalStatstring As String, _
    ByVal w3icon As String)
    
    Dim game   As String ' ...
    Dim pStats As String ' ...
    Dim Clan   As String ' ...
    
    game = ParseStatstring(OriginalStatstring, pStats, Clan)

    If (Not (JoinMessagesOff)) Then
        If (m_QueueCount <= 3) Then
            Call frmChat.AddChat(RTBColors.JoinText, "-- ", _
                RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                    RTBColors.JoinText, " has joined the channel using " & pStats)
        Else
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "-- ", _
                RTBColors.ErrorMessageText, Username & " [" & Ping & "ms]", _
                    RTBColors.ErrorMessageText, " has joined the channel using " & pStats)
        End If
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
    Dim game   As String  ' ...
    Dim pStats As String  ' ...
    Dim Clan   As String  ' ...

    i = UsernameToIndex(Username)
    
    colUsersInChannel.Item(i).Statstring = _
        OriginalStatstring
        
    game = ParseStatstring(OriginalStatstring, pStats, Clan)
    
    If (JoinMessagesOff = False) Then
        Call frmChat.AddChat(RTBColors.JoinText, "-- Stats updated: ", _
            RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
                RTBColors.JoinText, " is using " & pStats)
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

    If (Pos) Then
        Set found = frmChat.lvChannel.ListItems(Pos)
        
        If (g_ThisIconCode <> -1) Then
            If ((Not (squelching)) And _
                (Not (unsquelching))) Then
                
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
    
    Call frmChat.AddChat(CaratColor, "<", UsernameColor, Username, _
        CaratColor, "> ", TextColor, Message)
End Sub

' ...
Public Sub Event_QueuedEmote(ByVal Username As String, ByVal Flags As Long, ByVal Ping As Long, _
    ByVal Message As String)
    
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Event_QueuedEmote() has been fired.")
End Sub

' ...
Public Sub ClearChatQueue()
    Dim i As Integer ' ...
    
    ' clear chat queue
    'For i = 1 To colChatQueue.Count
    '    Call colChatQueue(1).Remove
    'Next i
    
    If (colChatQueue.Count) Then
        Call colChatQueue(1).Remove
    End If
End Sub
