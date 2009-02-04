Attribute VB_Name = "modChatQueue"
' ...

Option Explicit

' ...
'Public colChatQueue  As Collection

' ...
Private m_TimerID    As Long
'Private m_QueueCount As Long
'Private m_QueueGTC   As Long

' ...
Public Sub ChatQueue_Initialize()

    ' ...
    If (BotVars.ChatDelay > 0) Then
        ' ...
        'Set colChatQueue = New Collection
    
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
    End If
    
End Sub

' ...
Public Function ChatQueueTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
    ByVal dwTimer As Long)
    
    Dim CurrentUser  As clsUserObj
    Dim i            As Integer ' ...
    Dim j            As Integer ' ...
    Dim lastTimer    As Long    ' ...
    
    ' ...
    If (g_Channel Is Nothing) Then
        Exit Function
    End If

    ' ...
    If (g_Channel.IsSilent) Then
        Exit Function
    End If
    
    ' ...
    'If (GetTickCount() - lastTimer < BotVars.ChatDelay) Then
    '    Exit Function
    'End If
    
    ' ...
    For i = 1 To g_Channel.Users.Count
        ' ...
        If (i > g_Channel.Users.Count) Then
            Exit For
        End If
    
        ' ...
        Set CurrentUser = g_Channel.Users(i)
    
        ' ...
        If (CurrentUser.Queue.Count > 0) Then
            ' ...
            If ((GetTickCount() - CurrentUser.Queue(1).EventTick) >= BotVars.ChatDelay) Then
                ' ...
                CurrentUser.DisplayQueue
            End If
        End If
    Next i
    
    ' ...
    'lastTimer = GetTickCount()
    
    ' ...
    Set CurrentUser = Nothing
    
End Function
