Option Strict Off
Option Explicit On
Module modCommandsChat
	'This module will contain all the command code for commands that affect the chat environment
	'Things like Say, DND, Squelch, etc...
	Private m_AwayMessage As String
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAway(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(m_AwayMessage) > 0) Then
			Call frmChat.AddQ("/away", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
			If (Not Command_Renamed.IsLocal) Then
				If (m_AwayMessage = " - ") Then
					Call frmChat.AddQ("/me is back.", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				Else
					Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				End If
			End If
			m_AwayMessage = vbNullString
		Else
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(Command_Renamed.Argument("Message")) > 0) Then
				m_AwayMessage = Command_Renamed.Argument("Message")
				Call frmChat.AddQ("/away " & m_AwayMessage, (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				
				If (Command_Renamed.IsLocal) Then Call frmChat.AddQ("/me is away (" & m_AwayMessage & ")")
			Else
				m_AwayMessage = " - "
				Call frmChat.AddQ("/away", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				If (Not Command_Renamed.IsLocal) Then Call frmChat.AddQ("/me is away.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBack(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(m_AwayMessage) > 0) Then
			Call frmChat.AddQ("/away", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
			
			If (Not Command_Renamed.IsLocal) Then
				If (m_AwayMessage = " - ") Then
					Call frmChat.AddQ("/me is back.", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				Else
					Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				End If
				m_AwayMessage = vbNullString
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBlock(ByRef Command_Renamed As clsCommandObj)
		Dim FiltersPath As String
		Dim Total As Short
		Dim TotalString As String
		
		If (Command_Renamed.IsValid) Then
			FiltersPath = GetFilePath(FILE_FILTERS)
			If (CheckBlock(Command_Renamed.Argument("Username"))) Then 'Should prevent adding filters that are the same as other filters
				Command_Renamed.Respond("That username is already in the block list, or is under a wildcard block.")
			Else
				TotalString = ReadINI("BlockList", "Total", FiltersPath)
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (LenB(TotalString) = 0) Then TotalString = "0"
				
				If (StrictIsNumeric(TotalString)) Then
					Total = Int(CDbl(TotalString))
					WriteINI("BlockList", "Filter" & (Total + 1), Command_Renamed.Argument("Username"), FiltersPath)
					WriteINI("BlockList", "Total", CStr(Total + 1), FiltersPath)
					Command_Renamed.Respond(StringFormat("Added {0}{1}{0} to the username block list.", Chr(34), Command_Renamed.Argument("Username")))
				Else
					Command_Renamed.Respond("Your filters file has been edited manually and is no longer valid. Please delete it.")
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnConnect(ByRef Command_Renamed As clsCommandObj)
		frmChat.DoConnect()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnCQ(ByRef Command_Renamed As clsCommandObj)
		g_Queue.Clear()
		Command_Renamed.Respond("Queue cleared.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDisconnect(ByRef Command_Renamed As clsCommandObj)
		frmChat.DoDisconnect()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnExpand(ByRef Command_Renamed As clsCommandObj)
		Dim sMessage As String
		Dim tmpSend As String
		Dim i As Short
		If (Command_Renamed.IsValid) Then
			sMessage = Command_Renamed.Argument("Message")
			
			If (Len(tmpSend) > 223) Then tmpSend = Left(tmpSend, 223)
			For i = 1 To Len(sMessage)
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tmpSend = StringFormat("{0}{1}{2}", tmpSend, Mid(sMessage, i, 1), IIf(i = Len(sMessage), vbNullString, Space(1)))
			Next i
			
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Not Command_Renamed.Restriction("RAW_USAGE")) Then tmpSend = StringFormat("{0} Says: {1}", Command_Renamed.Username, tmpSend)
			
			If (Len(tmpSend) > 223) Then
				tmpSend = Left(tmpSend, 223)
			End If
			
			Command_Renamed.PublicOutput = True
			Command_Renamed.Respond(tmpSend)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFAdd(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call frmChat.AddQ("/f a " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
			Command_Renamed.Respond(StringFormat("Added user {0}{1}{0} to this account's friends list.", Chr(34), Command_Renamed.Argument("Username")))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFilter(ByRef Command_Renamed As clsCommandObj)
		Dim FiltersPath As String
		Dim Total As Short
		Dim TotalString As String
		
		If (Command_Renamed.IsValid) Then
			FiltersPath = GetFilePath(FILE_FILTERS)
			If (CheckMsg(Command_Renamed.Argument("Filter"))) Then 'Should prevent adding filters that are the same as other filters
				Command_Renamed.Respond("That filter is already in the list, or is under a wildcard.")
			Else
				TotalString = ReadINI("TextFilters", "Total", FiltersPath)
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (LenB(TotalString) = 0) Then TotalString = "0"
				
				If (StrictIsNumeric(TotalString)) Then
					Total = Int(CDbl(TotalString))
					WriteINI("TextFilters", "Filter" & (Total + 1), Command_Renamed.Argument("Filter"), FiltersPath)
					WriteINI("TextFilters", "Total", CStr(Total + 1), FiltersPath)
					Command_Renamed.Respond(StringFormat("Added {0}{1}{0} to the message filter list.", Chr(34), Command_Renamed.Argument("Filter")))
					Call frmChat.LoadArray(LOAD_FILTERS, gFilters)
				Else
					Command_Renamed.Respond("Your filters file has been edited manually and is no longer valid. Please delete it.")
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnForceJoin(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call FullJoin(Command_Renamed.Argument("Channel"))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFRem(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call frmChat.AddQ("/f r " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
			Command_Renamed.Respond(StringFormat("Removed user {0}{1}{0} from this account's friends list.", Chr(34), Command_Renamed.Argument("Username")))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnHome(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(Config.HomeChannel) = 0) Then
			' do product home join instead
			Call DoChannelJoinProductHome()
		Else
			' go home
			Call FullJoin((Config.HomeChannel), 2)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnReturn(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(BotVars.LastChannel) > 0) Then
			' go to last channel
			Call FullJoin((BotVars.LastChannel), 2)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIgnore(ByRef Command_Renamed As clsCommandObj)
		Dim dbTarget As udtGetAccessResponse
		Dim dbCaller As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
			If (Command_Renamed.IsLocal) Then
				Command_Renamed.Respond(StringFormat("Ignoring messages from {0}{1}{0}.", Chr(34), Command_Renamed.Argument("Username")))
				frmChat.AddQ("/ignore " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object dbTarget. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dbTarget = GetCumulativeAccess(Command_Renamed.Argument("Username"))
				'UPGRADE_WARNING: Couldn't resolve default property of object dbCaller. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dbCaller = GetCumulativeAccess(Command_Renamed.Username)
				
				If (dbTarget.Rank > dbCaller.Rank) Then
					Command_Renamed.Respond("That user has a higher rank then you.")
				ElseIf (dbTarget.Rank = dbCaller.Rank) Then 
					Command_Renamed.Respond("That user has the same rank as you.")
				Else
					Command_Renamed.Respond(StringFormat("Ignoring messages from {0}{1}{0}.", Chr(34), Command_Renamed.Argument("Username")))
					frmChat.AddQ("/ignore " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIgPriv(ByRef Command_Renamed As clsCommandObj)
		Call frmChat.AddQ("/o igpriv", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
		Command_Renamed.Respond("Ignoring messages from non-friends in private channels.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnJoin(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call frmChat.AddQ("/join " & Command_Renamed.Argument("Channel"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnBlock(ByRef Command_Renamed As clsCommandObj)
		
		Dim i As Short
		Dim Total As Short
		Dim TotalString As String
		Dim FiltersPath As String
		
		If (Command_Renamed.IsValid) Then
			FiltersPath = GetFilePath(FILE_FILTERS)
			
			TotalString = ReadINI("BlockList", "Total", FiltersPath)
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(TotalString) = 0) Then TotalString = "0"
			
			If (StrictIsNumeric(TotalString)) Then
				Total = Int(CDbl(TotalString))
				For i = 0 To Total
					If (StrComp(Command_Renamed.Argument("Username"), ReadINI("BlockList", "Filter" & (i + 1), FiltersPath), CompareMethod.Text) = 0) Then
						Exit For
					ElseIf (i = Total) Then 
						Command_Renamed.Respond(StringFormat("{0}{1}{0} is not blocked.", Chr(34), Command_Renamed.Argument("Username")))
						Exit Sub
					End If
				Next i
				
				If (i < Total) Then
					For i = i To (Total - 1)
						WriteINI("BlockList", "Filter" & (i + 1), ReadINI("BlockList", "Filter" & (i + 2), FiltersPath), FiltersPath)
					Next i
					WriteINI("BlockList", "Total", CStr(Total - 1), FiltersPath)
					Command_Renamed.Respond(StringFormat("Removed {0}{1}{0} from the blocked users list.", Chr(34), Command_Renamed.Argument("Username")))
				End If
			Else
				Command_Renamed.Respond("Your filters file has been edited manually and is no longer valid. Please delete it.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnFilter(ByRef Command_Renamed As clsCommandObj)
		Dim i As Short
		Dim Total As Short
		Dim TotalString As String
		Dim FiltersPath As String
		
		If (Command_Renamed.IsValid) Then
			FiltersPath = GetFilePath(FILE_FILTERS)
			
			TotalString = ReadINI("TextFilters", "Total", FiltersPath)
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(TotalString) = 0) Then TotalString = "0"
			
			If (StrictIsNumeric(TotalString)) Then
				Total = Int(CDbl(TotalString))
				For i = 0 To Total
					If (StrComp(Command_Renamed.Argument("Filter"), ReadINI("TextFilters", "Filter" & (i + 1), FiltersPath), CompareMethod.Text) = 0) Then
						Exit For
					ElseIf (i = Total) Then 
						Command_Renamed.Respond(StringFormat("{0}{1}{0} is not filtered.", Chr(34), Command_Renamed.Argument("Filter")))
						Exit Sub
					End If
				Next i
				
				If (i < Total) Then
					For i = i To (Total - 1)
						WriteINI("TextFilters", "Filter" & (i + 1), ReadINI("TextFilters", "Filter" & (i + 2), FiltersPath), FiltersPath)
					Next i
					WriteINI("TextFilters", "Total", CStr(Total - 1), FiltersPath)
					Command_Renamed.Respond(StringFormat("Removed {0}{1}{0} from the message filter list.", Chr(34), Command_Renamed.Argument("Filter")))
					Call frmChat.LoadArray(LOAD_FILTERS, gFilters)
				End If
			Else
				Command_Renamed.Respond("Your filters file has been edited manually and is no longer valid. Please delete it.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnIgnore(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Command_Renamed.Respond(StringFormat("Receiving messages from {0}{1}{0}.", Chr(34), Command_Renamed.Argument("Username")))
			frmChat.AddQ("/unignore " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnIgPriv(ByRef Command_Renamed As clsCommandObj)
		Call frmChat.AddQ("/o unigpriv", (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
		Command_Renamed.Respond("Allowing messages from non-friends in private channels.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnQuickRejoin(ByRef Command_Renamed As clsCommandObj)
		Call RejoinChannel((g_Channel.Name))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnReconnect(ByRef Command_Renamed As clsCommandObj)
		Dim LastChannel As String
		If (g_Online) Then
			
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If LenB(g_Channel.Name) = 0 And LenB(BotVars.LastChannel) > 0 Then
				' already outside chat environment
				LastChannel = BotVars.LastChannel
			Else
				' in chat room
				LastChannel = g_Channel.Name
			End If
			
			Call frmChat.DoDisconnect()
			
			'frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Reconnecting by command, please wait..."
			
			Pause(1)
			
			'frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
			
			Call frmChat.DoConnect()
			
			' reinstate last channel
			BotVars.LastChannel = LastChannel
			PrepareHomeChannelMenu()
			PrepareQuickChannelMenu()
		Else
			frmChat.AddChat(RTBColors.SuccessText, "Connection initialized.")
			
			Call frmChat.DoConnect()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnReJoin(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call frmChat.AddQ(StringFormat("/join {0} Rejoin", GetCurrentUsername), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
		Call frmChat.AddQ("/join " & g_Channel.Name, (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSay(ByRef Command_Renamed As clsCommandObj)
		Dim tmpSend As String
		
		If (Command_Renamed.IsValid) Then
			If (Not Command_Renamed.Restriction("RAW_USAGE")) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tmpSend = StringFormat("{0} Says: {1}", Command_Renamed.Username, Command_Renamed.Argument("Message"))
			Else
				tmpSend = Command_Renamed.Argument("Message")
			End If
			
			If (Len(tmpSend) > 223) Then tmpSend = Left(tmpSend, 223)
			
			Command_Renamed.PublicOutput = True
			Command_Renamed.Respond(tmpSend)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSCQ(ByRef Command_Renamed As clsCommandObj)
		g_Queue.Clear()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShout(ByRef Command_Renamed As clsCommandObj)
		Dim tmpSend As String
		
		If (Command_Renamed.IsValid) Then
			
			If (Not Command_Renamed.Restriction("RAW_USAGE")) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tmpSend = StringFormat("{0} Shouts: {1}", Command_Renamed.Username, UCase(Command_Renamed.Argument("Message")))
			Else
				tmpSend = UCase(Command_Renamed.Argument("Message"))
			End If
			
			If (Len(tmpSend) > 223) Then tmpSend = Left(tmpSend, 223)
			
			Command_Renamed.PublicOutput = True
			Command_Renamed.Respond(tmpSend)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWatch(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			WatchUser = Command_Renamed.Argument("Username")
			Command_Renamed.Respond(StringFormat("Now watching {0}{1}{0}.", Chr(34), Command_Renamed.Argument("Username")))
		Else
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(WatchUser) > 0) Then
				Command_Renamed.Respond(StringFormat("Stopped watching {0}{1}{0}.", Chr(34), WatchUser))
				WatchUser = vbNullString
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWatchOff(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(WatchUser) > 0) Then
			Command_Renamed.Respond(StringFormat("Stopped watching {0}{1}{0}.", Chr(34), WatchUser))
			WatchUser = vbNullString
		Else
			Command_Renamed.Respond("Watch is disabled.")
		End If
	End Sub
End Module