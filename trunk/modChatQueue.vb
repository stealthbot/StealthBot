Option Strict Off
Option Explicit On
Module modChatQueue
	
	
	Public Sub ChatQueue_Initialize()
		If (BotVars.ChatDelay > 0) Then
			With frmChat.ChatQueueTimer
				'UPGRADE_WARNING: Timer property ChatQueueTimer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
				.Interval = IIf(BotVars.ChatDelay <= 500, BotVars.ChatDelay, 500)
				
				.Enabled = True
			End With
		End If
		
	End Sub
	
	Public Function ChatQueueTimerProc() As Object
		
		Dim CurrentUser As clsUserObj
		Dim I As Short
		Dim j As Short
		Dim lastTimer As Integer
		
		If (g_Channel Is Nothing) Then
			Exit Function
		End If
		
		If (g_Channel.IsSilent) Then
			Exit Function
		End If
		
		For I = 1 To g_Channel.Users.Count()
			If (I > g_Channel.Users.Count()) Then
				Exit For
			End If
			
			CurrentUser = g_Channel.Users.Item(I)
			
			If (CurrentUser.Queue.Count() > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Queue(1).EventTick. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ((GetTickCount() - CurrentUser.Queue.Item(1).EventTick) >= BotVars.ChatDelay) Then
					CurrentUser.DisplayQueue()
				End If
			End If
		Next I
		
		'UPGRADE_NOTE: Object CurrentUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CurrentUser = Nothing
	End Function
End Module