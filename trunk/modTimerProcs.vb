Option Strict Off
Option Explicit On
Module modTimerProcs
	
	Public Sub Reconnect_TimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer)
		If (Not (UserCancelledConnect)) Then
			Call frmChat.Connect()
		End If
		
		UserCancelledConnect = False
		
		Call KillTimer(0, ReconnectTimerID)
		
		ReconnectTimerID = 0
	End Sub
	
	Public Sub ExtendedReconnect_TimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer)
		
		'// Ticks = number of 30-second intervals passed since timer was initiated
		ExReconTicks = (ExReconTicks + 1)
		
		If ((ExReconTicks >= ExReconMinutes) And (UserCancelledConnect = False)) Then
			Call KillTimer(0, ExReconnectTimerID)
			
			UserCancelledConnect = False
			
			Call frmChat.Connect()
			
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
	
	Public Sub ScriptReload_TimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer)
		On Error Resume Next
		Static hasRunAlready As Boolean
		
		KillTimer(frmChat.Handle.ToInt32, SCReloadTimerID)
		SCReloadTimerID = 0
		
		If hasRunAlready Then
			hasRunAlready = False
		Else
			If idEvent > 0 Then
				frmChat.SControl.Timeout = idEvent
			Else
				Call frmChat.mnuReloadScripts_Click(Nothing, New System.EventArgs())
			End If
			
			hasRunAlready = True
		End If
	End Sub
	
	'/* Timer proc for invite accept form - deny accept/decline of the invitation for 3 sec */
	Public Function ClanInviteTimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer) As Object
		Call KillTimer(frmClanInvite.Handle.ToInt32, ClanAcceptTimerID)
		ClanAcceptTimerID = 0
		
		frmClanInvite.cmdAccept.Enabled = True
		frmClanInvite.cmdDecline.Enabled = True
		
		'UPGRADE_WARNING: Add a delegate for AddressOf ClanInviteTimerProc2 Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		ClanAcceptTimerID = SetTimer(frmClanInvite.Handle.ToInt32, 0, 28000, AddressOf ClanInviteTimerProc2)
	End Function
	
	'/* Timer proc 2 for invite accept form - autodeclines after 30 seconds */
	Public Function ClanInviteTimerProc2(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer) As Object
		Call KillTimer(frmClanInvite.Handle.ToInt32, ClanAcceptTimerID)
		Call frmClanInvite.cmdDecline_Click(Nothing, New System.EventArgs())
	End Function
	
	
	' Timer procedure for the queue
	Public Sub QueueTimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTimer As Integer)
		Dim NewDelay As Integer
		Dim ExtraDelay As Integer
		
		Dim Message As String
		Dim Tag As String
		Dim pri As Short
		Dim ID As Double
		
		Call KillTimer(0, QueueTimerID)
		QueueTimerID = 0
		
		On Error GoTo ERROR_HANDLER
		
		If ((g_Queue.Count) And (g_Online)) Then
			With g_Queue.Peek
				Message = .Message
				Tag = .Tag
				pri = .PRIORITY
				ID = .ID
			End With
			
			If (StrComp(Message, "%%%%%blankqueuemessage%%%%%", CompareMethod.Binary) = 0) Then
				'// This is a dummy queue message faking a 70-character queue entry
				' just pop and move on
				Call g_Queue.Pop()
			Else
				Call bnetSend(Message, Tag, ID)
				Call g_Queue.Pop()
				
				' OK, we've sent our message, now figure out the delay to the next one
				If (g_Queue.Count) Then
					With g_Queue.Peek()
						NewDelay = g_BNCSQueue.GetDelay(.Message)
						
						If .PRIORITY = modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE Then
							ExtraDelay = g_BNCSQueue.BanDelay()
						End If
					End With
					
					'UPGRADE_WARNING: Add a delegate for AddressOf QueueTimerProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
					QueueTimerID = SetTimer(0, 0, NewDelay + ExtraDelay, AddressOf QueueTimerProc)
				End If
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in QueueTimer_Timer().")
		
		Exit Sub
	End Sub
End Module