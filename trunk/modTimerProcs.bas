Attribute VB_Name = "modTimerProcs"
Option Explicit

Public Sub Reconnect_TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    If (AutoReconnectActive) Then
        If (AutoReconnectTicks >= AutoReconnectIn) Then
            Call KillTimer(0, ReconnectTimerID)
            Call frmChat.Connect
        Else
            AutoReconnectTicks = AutoReconnectTicks + 1
        End If
    Else
        Call KillTimer(0, ReconnectTimerID)
    End If
End Sub

Public Sub ScriptReload_TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    On Error Resume Next
    Static hasRunAlready As Boolean
    
    KillTimer frmChat.hWnd, SCReloadTimerID
    SCReloadTimerID = 0
    
    If hasRunAlready Then
        hasRunAlready = False
    Else
        If idEvent > 0 Then
            frmChat.SControl.Timeout = idEvent
        Else
            Call frmChat.mnuReloadScripts_Click
        End If
        
        hasRunAlready = True
    End If
End Sub


' Timer procedure for the queue
Public Sub QueueTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Dim NewDelay As Long
    Dim ExtraDelay As Long

    Dim Message  As String
    Dim Tag      As String
    Dim pri      As Integer
    Dim ID       As Double
    
    Call KillTimer(0&, QueueTimerID)
    QueueTimerID = 0
    
    On Error GoTo ERROR_HANDLER
    
    If ((g_Queue.Count) And (g_Online)) Then
        With g_Queue.Peek
            Message = .Message
            Tag = .Tag
            pri = .PRIORITY
            ID = .ID
        End With
        
        If (StrComp(Message, "%%%%%blankqueuemessage%%%%%", vbBinaryCompare) = 0) Then
            '// This is a dummy queue message faking a 70-character queue entry
            ' just pop and move on
            Call g_Queue.Pop
        Else
            Call bnetSend(Message, Tag, ID)
            Call g_Queue.Pop
            
            ' OK, we've sent our message, now figure out the delay to the next one
            If (g_Queue.Count) Then
                With g_Queue.Peek()
                    NewDelay = g_BNCSQueue.GetDelay(.Message)
                    
                    If .PRIORITY = PRIORITY.CHANNEL_MODERATION_MESSAGE Then
                        ExtraDelay = g_BNCSQueue.BanDelay()
                    End If
                End With
                
                QueueTimerID = SetTimer(0&, 0&, NewDelay + ExtraDelay, _
                    AddressOf QueueTimerProc)
            End If
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in QueueTimer_Timer()."

    Exit Sub
End Sub

