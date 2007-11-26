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

    frmChat.AddChat RTBColors.JoinText, "-- ", _
        RTBColors.JoinUsername, Username & " [" & Ping & "ms]", _
            RTBColors.JoinText, " has joined the channel using " & pStats
End Sub

' ...
Public Sub Event_QueuedStatusUpdate()
    Call frmChat.AddChat(vbRed, "Event_QueuedJoin() has been fired.")
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
