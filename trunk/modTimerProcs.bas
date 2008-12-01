Attribute VB_Name = "modTimerProcs"
Option Explicit

Public Sub Reconnect_TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    If (Not (UserCancelledConnect)) Then
        Call frmChat.Connect
    End If
    
    UserCancelledConnect = False
    
    Call KillTimer(0, ReconnectTimerID)
    
    ReconnectTimerID = 0
End Sub

Public Sub ExtendedReconnect_TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal idEvent As Long, ByVal dwTimer As Long)
    
    '// Ticks = number of 30-second intervals passed since timer was initiated
    ExReconTicks = (ExReconTicks + 1)
    
    If ((ExReconTicks >= ExReconMinutes) And (UserCancelledConnect = False)) Then
        Call KillTimer(0, ExReconnectTimerID)
        
        UserCancelledConnect = False
        
        Call frmChat.Connect
        
        ExReconnectTimerID = 0
        ExReconTicks = 0
        ExReconMinutes = 0
    End If
    
    If ((ExReconTicks >= ExReconMinutes) Or (UserCancelledConnect)) Then
        Call KillTimer(0, ExReconnectTimerID)
        
        UserCancelledConnect = False
        
        ExReconnectTimerID = 0
        ExReconTicks = 0
        ExReconMinutes = 0
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
            Call frmChat.mnuReloadScript_Click
        End If
        
        hasRunAlready = True
    End If
End Sub

'/* Timer proc for invite accept form - deny accept/decline of the invitation for 3 sec */
Public Function ClanInviteTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Call KillTimer(frmClanInvite.hWnd, ClanAcceptTimerID)
    ClanAcceptTimerID = 0
    
    frmClanInvite.cmdAccept.Enabled = True
    frmClanInvite.cmdDecline.Enabled = True
    
    ClanAcceptTimerID = SetTimer(frmClanInvite.hWnd, 0, 28000, AddressOf ClanInviteTimerProc2)
End Function

'/* Timer proc 2 for invite accept form - autodeclines after 30 seconds */
Public Function ClanInviteTimerProc2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Call KillTimer(frmClanInvite.hWnd, ClanAcceptTimerID)
    Call frmClanInvite.cmdDecline_Click
End Function
