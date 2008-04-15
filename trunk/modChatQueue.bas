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
    
    Dim CurrentUser  As clsUserObj
    Dim CurrentEvent As clsUserEventObj
    Dim i            As Integer ' ...
    Dim j            As Integer ' ...
    Dim blnShow      As Boolean ' ...
    
    ' ...
    If (g_Channel.IsSilent) Then
        Exit Function
    End If
    
    ' ...
    For i = 1 To g_Channel.Users.Count
        ' ...
        Set CurrentUser = g_Channel.Users(i)
    
        ' ...
        If (CurrentUser.Queue.Count > 0) Then
            ' ...
            If ((GetTickCount() - CurrentUser.Queue(1).EventTick) >= BotVars.ChatDelay) Then
                ' ...
                DisplayEvents CurrentUser
            End If
        End If
    Next i
    
End Function

' ...
Public Function DisplayEvents(ByRef CurrentUser As clsUserObj)

    Dim CurrentEvent As clsUserEventObj
    Dim j            As Integer
    
    ' ...
    For j = 1 To CurrentUser.Queue.Count
        ' ...
        Set CurrentEvent = CurrentUser.Queue(j)
    
        ' ...
        Select Case (CurrentEvent.EventID)
            ' ...
            Case ID_USER
                Call Event_UserInChannel(CurrentUser.Name, CurrentEvent.Flags, CurrentEvent.Statstring, _
                    CurrentEvent.Ping, CurrentEvent.GameID, CurrentEvent.Clan, CurrentEvent.Statstring, _
                        CurrentEvent.IconCode, j)
                    
            ' ...
            Case ID_JOIN
                Call Event_UserJoins(CurrentUser.Name, CurrentEvent.Flags, CurrentEvent.Statstring, _
                    CurrentEvent.Ping, CurrentEvent.GameID, CurrentEvent.Clan, CurrentEvent.Statstring, _
                        CurrentEvent.IconCode, j)
            
            ' ...
            Case ID_TALK
                Call Event_UserTalk(CurrentUser.Name, CurrentEvent.Flags, CurrentEvent.Message, _
                    CurrentEvent.Ping, j)
            
            ' ...
            Case ID_EMOTE
                Call Event_UserEmote(CurrentUser.Name, CurrentEvent.Flags, CurrentEvent.Message, j)
            
            ' ...
            Case ID_USERFLAGS
                Call Event_FlagsUpdate(CurrentUser.Name, CurrentEvent.Statstring, CurrentEvent.Flags, _
                    CurrentEvent.Ping, CurrentEvent.GameID, j)
        End Select
    Next j
    
    ' ...
    CurrentUser.ClearQueue

End Function
