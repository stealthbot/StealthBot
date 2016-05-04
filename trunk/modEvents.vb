Option Strict Off
Option Explicit On
Module modEvents
	'StealthBot Project - modEvents.bas
	' Andy T (andy@stealthbot.net) March 2005
	
	Private Const OBJECT_NAME As String = "modEvents"
	
	Private Const MSG_FILTER_MAX_EVENTS As Integer = 100 ' maximum number of storable events
	Private Const MSG_FILTER_DELAY_INT As Integer = 500 ' interval for event count measuring
	Private Const MSG_FILTER_MSG_COUNT As Integer = 3 ' message count maximums
	
	Private Structure MSGFILTER
		Dim UserObj As Object
		Dim EventObj As Object
		Dim EventTime As Date
	End Structure
	
	Private m_arrMsgEvents() As MSGFILTER
	Private m_eventCount As Short
	
	Public Sub Event_FlagsUpdate(ByVal Username As String, ByVal Flags As Integer, ByVal Message As String, ByVal Ping As Integer, Optional ByRef QueuedEventID As Short = 0)
		
		On Error GoTo ERROR_HANDLER
		
		Dim UserObj As clsUserObj
		Dim PreviousUserObj As clsUserObj
		Dim UserEvent As clsUserEventObj
		
		Dim UserIndex As Short
		Dim i As Short
		Dim PreviousFlags As Integer
		Dim Clan As String
		Dim parsed As String
		Dim pos As Short
		Dim doUpdate As Boolean
		Dim Displayed As Boolean ' stores whether this event has been displayed by another event in the RTB
		
		' if our username is for some reason null, we don't
		' want to continue, possibly causing further errors
        If (Len(Username) < 1) Then
            Exit Sub
        End If
		
		
		UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		If (UserIndex > 0) Then
			UserObj = g_Channel.Users.Item(UserIndex)
			
			If (QueuedEventID = 0) Then
				If (UserObj.Queue.Count() > 0) Then
					UserEvent = New clsUserEventObj
					
					With UserEvent
						.EventID = ID_USERFLAGS
						.Flags = Flags
						.Ping = Ping
						.GameID = UserObj.Game
					End With
					
					UserObj.Queue.Add(UserEvent)
				Else
					PreviousFlags = UserObj.Flags
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object UserObj.Queue().Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PreviousFlags = UserObj.Queue.Item(QueuedEventID - 1).Flags
			End If
			
			Clan = UserObj.Clan
		Else
			If (g_Channel.IsSilent = False) Then
				frmChat.AddChat(RTBColors.ErrorMessageText, "Warning: There was a flags update received for a user that we do " & "not have a record for.  This may be indicative of a server split or other technical difficulty.")
				
				Exit Sub
			Else
				If (g_Channel.Users.Count() >= 200) Then
					Exit Sub
				End If
				
				UserObj = New clsUserObj
				
				With UserObj
					.Name = Username
					.Statstring = Message
				End With
			End If
		End If
		
		With UserObj
			.Flags = Flags
			.Ping = Ping
		End With
		
		If (g_Channel.IsSilent) Then
			g_Channel.Users.Add(UserObj)
		End If
		
		' convert username to appropriate
		' display format
		Username = UserObj.DisplayName
		
		' are we receiving a flag update for ourselves?
		If (StrComp(Username, GetCurrentUsername, CompareMethod.Binary) = 0) Then
			' assign my current flags to the
			' relevant internal variable
			MyFlags = Flags
			
			' assign my current flags to the
			' relevant scripting variable
			SharedScriptSupport.BotFlags = MyFlags
		End If
		
		' we aren't in a silent channel, are we?
		Dim NewFlags As Integer
		Dim LostFlags As Integer
		Dim FDescN As String
		Dim FDescO As String
		If (g_Channel.IsSilent) Then
			AddName(Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan)
		Else
			If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
				If (Flags <> PreviousFlags) Then
					If (g_Channel.Self.IsOperator) Then
						If ((Username = GetCurrentUsername) And ((PreviousFlags And USER_CHANNELOP) <> USER_CHANNELOP)) Then
							
							g_Channel.CheckUsers()
						Else
							g_Channel.CheckUser(Username)
						End If
					End If
					
					pos = checkChannel(Username)
					
					If (pos) Then
						
						frmChat.lvChannel.Items.RemoveAt(pos)
						
						' voodoo magic: only show flags that are new
						NewFlags = Not (VB6.Imp(Flags, PreviousFlags))
						LostFlags = Not (VB6.Imp(PreviousFlags, Flags))
						
						If (NewFlags And USER_CHANNELOP) = USER_CHANNELOP Or (NewFlags And USER_BLIZZREP) = USER_BLIZZREP Or (NewFlags And USER_SYSOP) = USER_SYSOP Then
							pos = 1
						End If
						
						AddName(Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan, pos)
						
						' default to display this event
						Displayed = False
						
						' check whether it has been
						If QueuedEventID > 0 And UserObj.Queue.Count() >= QueuedEventID Then
							UserEvent = UserObj.Queue.Item(QueuedEventID)
							Displayed = UserEvent.Displayed
						End If
						
						' display if it has not
						If Not Displayed Then
							FDescN = FlagDescription(NewFlags, False)
							FDescO = FlagDescription(LostFlags, False)
							
                            If Len(FDescN) > 0 Then
                                frmChat.AddChat(RTBColors.JoinUsername, "-- ", RTBColors.JoinedChannelName, Username, RTBColors.JoinText, " is now a " & FDescN & ".")
                            End If
							
                            If Len(FDescO) > 0 Then
                                frmChat.AddChat(RTBColors.JoinUsername, "-- ", RTBColors.JoinedChannelName, Username, RTBColors.JoinText, " is no longer a " & FDescO & ".")
                            End If
						End If
					End If
				End If
			End If
		End If
		
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			RunInAll("Event_FlagUpdate", Username, Flags, Ping)
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_FlagsUpdate()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_JoinedChannel(ByVal ChannelName As String, ByVal Flags As Integer)
		On Error GoTo ERROR_HANDLER
		
		Dim mailCount As Short
		Dim ToANSI As String
		Dim LastChannel As String
		Dim sChannel As String
		
		' if our channel is for some reason null, we don't
		' want to continue, possibly causing further errors
        If (Len(ChannelName) = 0) Then
            Exit Sub
        End If
		
		LastChannel = g_Channel.Name
		
		Call frmChat.ClearChannel()
		
		If (frmChat.mnuUTF8.Checked) Then
			ToANSI = UTF8Decode(ChannelName)
			
			If (Len(ToANSI) > 0) Then
				ChannelName = ToANSI
			End If
		End If
		
		' we want to reset our filter
		' Values() when we join a new channel
		'BotVars.JoinWatch = 0
		
		'frmChat.tmrSilentChannel(0).Enabled = False
		
		'If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
		'    PassedClanMotdCheck = False
		'End If
		
		' show home channel in menu
		BotVars.LastChannel = LastChannel
		
		' if we've just left another channel, call event script
		' function indicating that we've done so.
        If (Len(LastChannel) > 0) Then
            On Error Resume Next

            RunInAll("Event_ChannelLeave")

            On Error GoTo ERROR_HANDLER
        End If
		
		With g_Channel
			.Name = ChannelName
			.Flags = Flags
			.JoinTime = UtcNow
		End With
		
		PrepareHomeChannelMenu()
		PrepareQuickChannelMenu()
		
		SharedScriptSupport.MyChannel = ChannelName
		
		If (Len(g_Clan.Name) > 0) Then
			If (StrComp(g_Channel.Name, "Clan " & g_Clan.Name, CompareMethod.Text) = 0) Then
				RequestClanMOTD(1)
			End If
		End If
		
		frmChat.AddChat(RTBColors.JoinedChannelText, "-- Joined channel: ", RTBColors.JoinedChannelName, ChannelName, RTBColors.JoinedChannelText, " --")
		
		SetTitle(GetCurrentUsername & ", online in channel " & g_Channel.Name)
		
		frmChat.UpdateTrayTooltip()
		
		frmChat.ListviewTabs.SelectedIndex = LVW_BUTTON_CHANNEL
		Call frmChat.ListviewTabs_SelectedIndexChanged(Nothing, New System.EventArgs())
		
		' have we just joined the void?
		If (g_Channel.IsSilent) Then
			' if we've joined the void, lets try to grab the list of
			' users within the channel by attempting to force a user
			' update message using Battle.net's unignore command.
			If (frmChat.mnuDisableVoidView.Checked = False) Then
				' lets inform user of potential lag issues while in this channel
				frmChat.AddChat(RTBColors.InformationText, "If you experience a lot of lag while within " & "this channel, try selecting 'Disable Silent Channel View' from the Window menu.")
				
				frmChat.tmrSilentChannel(1).Enabled = True
				
				frmChat.AddQ("/unignore " & GetCurrentUsername)
			End If
		Else
			frmChat.tmrSilentChannel(1).Enabled = False
		End If
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' check for mail
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		mailCount = GetMailCount(GetCurrentUsername)
		
		If (mailCount) Then
			frmChat.AddChat(RTBColors.ConsoleText, "You have " & mailCount & " new message" & IIf(mailCount = 1, "", "s") & ". Type /inbox to retrieve.")
		End If
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' call event script function
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		On Error Resume Next
		
		RunInAll("Event_ChannelJoin", ChannelName, Flags)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_JoinedChannel()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_KeyReturn(ByVal KeyName As String, ByVal KeyValue As String)
		On Error Resume Next
		
		Dim ft As FILETIME
		Dim st As SYSTEMTIME
		Dim s() As String
		Dim U As String
		Dim i As Short
		
		Static KeysReceived As Short
		
		'MsgBox PPL
		
		' Some of the oldest code in this project lives right here
		Dim x() As String
		If SuppressProfileOutput Then
			
			' // We're receiving profile information from a scripter request
			' // No need to do anything at all with it except set Suppress = False after
			' // the description comes in, and of course hadn it over to the scripters
			RunInAll("Event_KeyReturn", KeyName, KeyValue)
			
			' clean up variables once profile keys are received:
			Select Case KeyName
				Case "Profile\Age", "Profile\Sex", "Profile\Location", "Profile\Description"
					If (StrComp(SpecificProfileKey, KeyName, CompareMethod.Binary) = 0) Then
						' SSC.RequestProfileKey() called for one of the profile keys
						SuppressProfileOutput = False
						SpecificProfileKey = vbNullString
					Else
						' normal: wait for 4 keys then reset suppress
						KeysReceived = KeysReceived + 1
						If (KeysReceived >= 4) Then
							SuppressProfileOutput = False
							KeysReceived = 0
						End If
					End If
					
				Case Else
					' SSC.RequestProfileKey() called for a different key
					SuppressProfileOutput = False
					SpecificProfileKey = vbNullString
					
			End Select
			
		ElseIf ProfileRequest = True Then 
			
			'MsgBox "!!"
			' writing profile
			
			frmProfile.SetKey(KeyName, KeyValue)
			
			RunInAll("Event_KeyReturn", KeyName, KeyValue)
			
			' wait for 4 keys then reset request var
			KeysReceived = KeysReceived + 1
			If KeysReceived >= 4 Then
				ProfileRequest = False
				KeysReceived = 0
			End If
			
			' Public Profile Listing
		ElseIf PPL = True Then 
			
			'MsgBox PPLRespondTo
			
            If Len(PPLRespondTo) > 0 Then
                U = "/w " & PPLRespondTo & " "
            Else
                U = vbNullString
            End If
			
			If KeyName = "Profile\Location" Then
Repeat2: 
				i = InStr(1, KeyValue, Chr(13))
				
				If Len(KeyValue) > 90 Then
					If i <> 0 Then
						frmChat.AddQ(U & "[Location] " & Left(KeyValue, Len(KeyValue) - i))
						KeyValue = Right(KeyValue, Len(KeyValue) - i)
						
						GoTo Repeat2
					Else
						frmChat.AddQ(U & "[Location] " & KeyValue)
					End If
				Else
					If i <> 0 Then
						frmChat.AddQ(U & "[Location] " & Left(KeyValue, Len(KeyValue) - i))
						KeyValue = Right(KeyValue, Len(KeyValue) - i)
						GoTo Repeat2
					Else
						frmChat.AddQ(U & "[Location] " & KeyValue)
					End If
				End If
				
			ElseIf KeyName = "Profile\Description" Then 
				
				
				x = Split(KeyValue, Chr(13))
				ReDim s(0)
				
				For i = LBound(x) To UBound(x)
					s(0) = x(i)
					
					If Len(s(0)) > 200 Then s(0) = Left(s(0), 200)
					
					If i = LBound(x) Then
						frmChat.AddQ(U & "[Descr] " & s(0))
					Else
						frmChat.AddQ(U & "[Descr] " & Right(s(0), Len(s(0)) - 1))
					End If
				Next i
				
				PPL = False
				
                If Len(PPLRespondTo) > 0 Then
                    PPLRespondTo = vbNullString
                End If
				
			ElseIf KeyName = "Profile\Sex" Then 
Repeat4: 
				If Len(KeyValue) > 90 Then
					frmChat.AddQ(U & "[Sex] " & Left(KeyValue, 80) & " [more]")
					KeyValue = Right(KeyValue, Len(KeyValue) - 80)
					GoTo Repeat4
				Else
					frmChat.AddQ(U & "[Sex] " & KeyValue)
				End If
				
			ElseIf Left(KeyName, 7) = "System\" Then 
				
				If InStr(1, KeyValue, " ", CompareMethod.Text) > 0 Then '// If it's a FILETIME
					
					'Dim FT As FILETIME
					'Dim sT As SYSTEMTIME
					
					ft.dwHighDateTime = CInt(Left(KeyValue, InStr(1, KeyValue, " ", CompareMethod.Text)))
					
					'On Error Resume Next
					
					KeyValue = Mid(KillNull(KeyValue), InStr(1, KeyValue, " ", CompareMethod.Text) + 1)
					'keyvalue = Left$(keyvalue, Len(keyvalue) - 1)
					
					ft.dwLowDateTime = CInt(KeyValue) 'CLng(KeyValue & "0")
					
					FileTimeToSystemTime(ft, st)
					
					With st
						frmChat.AddQ(U & Right(KeyName, Len(KeyName) - 7) & ": " & SystemTimeToString(st) & " (Battle.net time)")
					End With
					
				Else '// it's a SECONDS type
					If StrictIsNumeric(KeyValue) Then
						'On Error Resume Next
						frmChat.AddQ(U & "Time Logged: " & ConvertTime(CDbl(KeyValue), 1))
					End If
				End If
				
			End If
			
		ElseIf Left(KeyName, 7) = "System\" Then 
			
			'frmchat.addchat RTBColors.ConsoleText, KeyName & ": " & KeyValue
			
			If InStr(1, KeyValue, " ", CompareMethod.Text) > 0 Then '// If it's a FILETIME
				
				'Dim FT As FILETIME
				'Dim sT As SYSTEMTIME
				
				ft.dwHighDateTime = CInt(Left(KeyValue, InStr(1, KeyValue, " ", CompareMethod.Text)))
				
				'On Error Resume Next
				
				KeyValue = Mid(KillNull(KeyValue), InStr(1, KeyValue, " ", CompareMethod.Text) + 1)
				'keyvalue = Left$(keyvalue, Len(keyvalue) - 1)
				
				ft.dwLowDateTime = CInt(KeyValue) 'CLng(KeyValue & "0")
				
				FileTimeToSystemTime(ft, st)
				
				With st
					frmChat.AddChat(RTBColors.ServerInfoText, Right(KeyName, Len(KeyName) - 7) & ": " & SystemTimeToString(st) & " (Battle.net time)")
				End With
				
			Else '// it's a SECONDS type
				If StrictIsNumeric(KeyValue) Then
					'On Error Resume Next
					frmChat.AddChat(RTBColors.ServerInfoText, "Time Logged: " & ConvertTime(CDbl(KeyValue), 1))
				End If
			End If
			
		Else
			
			' viewing profile (or &H26 sent by script and &H26 isn't veto'd in script)
			
			frmProfile.SetKey(KeyName, KeyValue)
			
			RunInAll("Event_KeyReturn", KeyName, KeyValue)
			
		End If
	End Sub
	
	Public Sub Event_LeftChatEnvironment()
		On Error GoTo ERROR_HANDLER
		
		BotVars.LastChannel = g_Channel.Name
		PrepareHomeChannelMenu()
		PrepareQuickChannelMenu()
		
		frmChat.ClearChannel()
		
		SetTitle(GetCurrentUsername & ", online on " & BotVars.Gateway)
		
		frmChat.lblCurrentChannel.Text = frmChat.GetChannelString()
		
		frmChat.AddChat(RTBColors.JoinedChannelText, "-- Left channel --")
		
		On Error Resume Next
		
		RunInAll("Event_ChannelLeave")
		
		On Error GoTo ERROR_HANDLER
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_LeftChatEnvironment()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_LoggedOnAs(ByRef Username As String, ByRef Statstring As String, ByRef AccountName As String)
		On Error GoTo ERROR_HANDLER
		Dim sChannel As String
		Dim ShowW3 As Boolean
		Dim ShowD2 As Boolean
		Dim Stats As New clsUserStats
		
		LastWhisper = vbNullString
		
		'If InStr(1, Username, "*", vbBinaryCompare) <> 0 Then
		'    Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbBinaryCompare))
		'End If
		
		Call g_Queue.Clear()
		
		g_Online = True
		
		' in case this wasn't set before
		ds.EnteredChatFirstTime = True
		
		Stats.Statstring = Statstring
		
		CurrentUsername = KillNull(Username)
		
		'RequestSystemKeys
		
		Call SetNagelStatus(frmChat.sckBNet.SocketHandle, True)
		
		Call EnableSO_KEEPALIVE(frmChat.sckBNet.SocketHandle)
		
		If (StrComp(Left(CurrentUsername, 2), "w#", CompareMethod.Text) = 0) Then
			CurrentUsername = Mid(CurrentUsername, 3)
		End If
		
		' if D2 and on a char, we need to tell the whole world this so that Self is known later on
		If (StrComp(Stats.Game, PRODUCT_D2DV, CompareMethod.Binary) = 0) Or (StrComp(Stats.Game, PRODUCT_D2XP, CompareMethod.Binary) = 0) Then
            If (Len(Stats.CharacterName) > 0) Then
                CurrentUsername = Stats.CharacterName & "*" & CurrentUsername
            End If
		End If
		
		' show home channel in menu
		PrepareHomeChannelMenu()
		PrepareQuickChannelMenu()
		
		' setup Bot menu game-specific features
		ShowW3 = (StrComp(Stats.Game, PRODUCT_WAR3, CompareMethod.Binary) = 0) Or (StrComp(Stats.Game, PRODUCT_W3XP, CompareMethod.Binary) = 0)
		ShowD2 = (StrComp(Stats.Game, PRODUCT_D2DV, CompareMethod.Binary) = 0) Or (StrComp(Stats.Game, PRODUCT_D2XP, CompareMethod.Binary) = 0)
		frmChat.mnuSepZ.Visible = (ShowW3 Or ShowD2)
		frmChat.mnuIgnoreInvites.Visible = ShowW3
		frmChat.mnuRealmSwitch.Visible = ShowD2
		
		SharedScriptSupport.myUsername = GetCurrentUsername
		
		With frmChat
			.InitListviewTabs()
			
			.AddChat(RTBColors.InformationText, "[BNCS] Logged on as ", RTBColors.SuccessText, Username, RTBColors.InformationText, StringFormat(" using {0}.", Stats.ToString_Renamed))
			
			.tmrAccountLock.Enabled = False
			
			.UpTimer.Interval = 1000
			
			'.Timer.Interval = 30000
			.tmrIdleTimer.Interval = 1000
			
			'.tmrClanUpdate.Enabled = True
		End With
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNLS.CtlState <> 0) Then
			frmChat.sckBNLS.Close()
		End If
		
		If (ExReconnectTimerID > 0) Then
			Call KillTimer(0, ExReconnectTimerID)
			
			ExReconnectTimerID = 0
		End If
		
		If Config.FriendsListTab Then
			Call frmChat.FriendListHandler.RequestFriendsList(PBuffer)
		End If
		
		'UPGRADE_NOTE: Object Stats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Stats = Nothing
		
		RequestSystemKeys()
        If (Len(BotVars.Gateway) > 0) Then
            ' PvPGN: we already have our gateway, we're logged on
            SetTitle(GetCurrentUsername() & ", online in channel " & g_Channel.Name)

            Call InsertDummyQueueEntry()

            On Error Resume Next

            RunInAll("Event_LoggedOn", CurrentUsername, BotVars.Product)
        End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_LoggedOnAs()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' updated 8-10-05 for new logging system
	Public Sub Event_LogonEvent(ByVal Message As Byte, Optional ByVal ExtraInfo As String = "")
		On Error GoTo ERROR_HANDLER
		Dim lColor As Integer
		Dim sMessage As String
		'Dim UseExtraInfo As Boolean
		
		Select Case (Message)
			Case 0
				lColor = RTBColors.ErrorMessageText
				
				sMessage = "Login error - account does not exist."
				
			Case 1
				lColor = RTBColors.ErrorMessageText
				
				sMessage = "Login error - invalid password."
				
			Case 2
				lColor = RTBColors.SuccessText
				
				sMessage = "Logon successful."
				
				frmChat.tmrAccountLock.Enabled = False
				
			Case 3
				lColor = RTBColors.InformationText
				
				sMessage = "Attempting to create account..."
				
			Case 4
				lColor = RTBColors.SuccessText
				
				sMessage = "Account created successfully."
				
			Case 5
				sMessage = ExtraInfo
				
				lColor = RTBColors.ErrorMessageText
		End Select
		
		frmChat.AddChat(lColor, "[BNCS] " & sMessage)
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_LogonEvent()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_ServerError(ByVal Message As String)
		On Error GoTo ERROR_HANDLER
		frmChat.AddChat(RTBColors.ErrorMessageText, Message)
		
		RunInAll("Event_ServerError", Message)
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_ServerError()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_ChannelJoinError(ByVal EventID As Short, ByVal ChannelName As String)
		On Error GoTo ERROR_HANDLER
		Dim ChannelJoinError As String
		Dim ChannelJoinButtons As MsgBoxStyle
		Dim ChannelJoinResult As MsgBoxResult
		Dim Message As String
		Dim ChannelCreateOption As String
		
		'frmChat.AddChat RTBColors.ErrorMessageText, Message
		
        If (Len(BotVars.Gateway) = 0) Then
            ' continue gateway discovery
            SEND_SID_CHATCOMMAND("/whoami")
        Else
            ChannelCreateOption = Config.AutoCreateChannels

            Select Case ChannelCreateOption
                Case "ALERT"
                    Select Case EventID
                        Case ID_CHANNELDOESNOTEXIST
                            ChannelJoinError = "Channel does not exist." & vbNewLine & "Do you want to create it?"
                            ChannelJoinButtons = MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton1
                        Case ID_CHANNELFULL
                            ChannelJoinError = "Channel is full."
                            ChannelJoinButtons = MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation Or MsgBoxStyle.DefaultButton1
                        Case ID_CHANNELRESTRICTED
                            ChannelJoinError = "Channel is restricted."
                            ChannelJoinButtons = MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation Or MsgBoxStyle.DefaultButton1
                    End Select

                    ChannelJoinResult = MsgBox("Failed to join " & ChannelName & ":" & vbNewLine & ChannelJoinError, ChannelJoinButtons, "StealthBot")

                    If ChannelJoinResult = MsgBoxResult.Yes Then
                        Call FullJoin(ChannelName, 2)
                    End If

                Case Else
                    ' "ALWAYS" - handle it as error to bot
                    ' "NEVER" - failed to join or create
                    Select Case EventID
                        Case ID_CHANNELDOESNOTEXIST
                            Message = "[BNCS] Channel does not exist."
                        Case ID_CHANNELFULL
                            Message = "[BNCS] Channel is full."
                        Case ID_CHANNELRESTRICTED
                            Message = "[BNCS] Channel is restricted."
                    End Select

                    frmChat.AddChat(RTBColors.ErrorMessageText, Message)

            End Select

            'should we expose?
            'RunInAll "Event_ChannelJoinError", EventID, ChannelName
        End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_ChannelJoinError()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_ServerInfo(ByVal Username As String, ByVal Message As String)
		On Error GoTo ERROR_HANDLER
		
		Const MSG_BANNED As String = " was banned by "
		Const MSG_UNBANNED As String = " was unbanned by "
		Const MSG_SQUELCHED As String = " has been squelched."
		Const MSG_UNSQUELCHED As String = " has been unsquelched."
		Const MSG_KICKEDOUT As String = " kicked you out of the channel!"
		Const MSG_FRIENDS As String = "Your friends are:"
		
		Dim i As Short
		Dim temp As String
		Dim bHide As Boolean
		Dim ToANSI As String
		
		If (Message = vbNullString) Then
			Exit Sub
		End If
		
		Username = ConvertUsername(Username)
		
		If (frmChat.mnuUTF8.Checked) Then
			ToANSI = UTF8Decode(Message)
			
			If (Len(ToANSI) > 0) Then
				Message = ToANSI
			End If
		End If
		
		If (StrComp(g_Channel.Name, "Clan " & Clan.Name, CompareMethod.Text) = 0) Then
			If (PassedClanMotdCheck = False) Then
				Call frmChat.AddChat(RTBColors.ServerInfoText, Message)
				
				Exit Sub
			End If
		End If
		
		If (g_request_receipt) Then ' for .cs and .cb commands
			Caching = True
			
			
			' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
			'Call CacheChannelList(Message, 1)
			Call CacheChannelList(modCommandsOps.CacheChanneListEnum.enAdd, Message)
			
			'With frmChat.cacheTimer
			'    .Enabled = False
			'    .Enabled = True
			'End With
		End If
		
		' what is our current gateway name?
		If (BotVars.Gateway = vbNullString) Then
			If (InStr(1, Message, "You are ", CompareMethod.Text) > 0) And (InStr(1, Message, ", using ", CompareMethod.Text) > 0) Then
				
				If ((InStr(1, Message, " in channel ", CompareMethod.Text) = 0) And (InStr(1, Message, " in game ", CompareMethod.Text) = 0) And (InStr(1, Message, " a private ", CompareMethod.Text) = 0)) Then
					
					i = InStrRev(Message, Space(1))
					
					BotVars.Gateway = Mid(Message, i + 1)
					
					SetTitle(GetCurrentUsername & ", online on " & BotVars.Gateway)
					
					Call DoChannelJoinHome()
					
					Call InsertDummyQueueEntry()
					
					RunInAll("Event_LoggedOn", CurrentUsername, BotVars.Product)
					
					Exit Sub
				End If
			End If
		End If
		
		Dim Banning As Boolean
		Dim Unbanning As Boolean
		Dim User As String
		Dim cOperator As String
		Dim msgPos As Short
		Dim pos As Short
		Dim tmp As String
		Dim banpos As Short
		Dim j As Short
		Dim Reason As String
		Dim BanlistObj As clsBanlistUserObj
		If (InStr(1, Message, Space(1), CompareMethod.Binary) <> 0) Then
			If (InStr(1, Message, "are still marked", CompareMethod.Text) <> 0) Then
				Exit Sub
			End If
			
			If ((InStr(1, Message, " from your friends list.", CompareMethod.Binary) > 0) Or (InStr(1, Message, " to your friends list.", CompareMethod.Binary) > 0) Or (InStr(1, Message, " in your friends list.", CompareMethod.Binary) > 0) Or (InStr(1, Message, " of your friends list.", CompareMethod.Binary) > 0)) Then
				
				If Config.FriendsListTab Then
					If Not frmChat.FriendListHandler.SupportsFriendPackets(Config.Game) Then
						Call frmChat.FriendListHandler.RequestFriendsList(PBuffer)
					End If
				End If
			End If
			
			'Ban Evasion and banned-user tracking
			temp = Split(Message, " ")(1)
			
			' added 1/21/06 thanks to
			' http://www.stealthbot.net/forum/index.php?showtopic=24582
			
			If (Len(temp) > 0) Then
				
				If (InStr(1, Message, MSG_BANNED, CompareMethod.Text) > 0) Then
					User = Left(Message, InStr(1, Message, MSG_BANNED, CompareMethod.Binary) - 1)
					
					Reason = Mid(Message, InStr(1, Message, MSG_BANNED, CompareMethod.Binary) + Len(MSG_BANNED) + 1) ' trim out username and banned message
					If (InStr(1, Reason, " (", CompareMethod.Binary)) Then 'Did they give a message?
						Reason = Mid(Reason, InStr(1, Reason, " (") + 2) 'trim out the banning name (Note, when banned by a rep using Len(Username) won't work as its banned "By a Blizzard Representative")
						Reason = Left(Reason, Len(Reason) - 2) 'Trim off the trailing ")."
					Else
						Reason = vbNullString
					End If
					
					If (Len(User) > 0) Then
						pos = g_Channel.GetUserIndex(Username)
						
						If (pos > 0) Then
							
							banpos = g_Channel.IsOnBanList(User, Username)
							
							If (banpos > 0) Then
								g_Channel.Banlist.Remove(banpos)
							Else
								g_Channel.BanCount = (g_Channel.BanCount + 1)
							End If
							
							If ((BotVars.StoreAllBans) Or (StrComp(Username, GetCurrentUsername, CompareMethod.Binary) = 0)) Then
								
								BanlistObj = New clsBanlistUserObj
								
								With BanlistObj
									.Name = User
									.Operator_Renamed = Username
									.DateOfBan = UtcNow
									.IsDuplicateBan = (g_Channel.IsOnBanList(User) > 0)
									.Reason = Reason
								End With
								
								If (BanlistObj.IsDuplicateBan) Then
									With g_Channel.Banlist.Item(g_Channel.IsOnBanList(User))
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().IsDuplicateBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										.IsDuplicateBan = False
									End With
								End If
								
								g_Channel.Banlist.Add(BanlistObj)
							End If
						End If
						
						Call RemoveBanFromQueue(User)
					End If
					
					If (frmChat.mnuHideBans.Checked) Then
						bHide = True
					End If
				ElseIf (InStr(1, Message, MSG_UNBANNED, CompareMethod.Text) > 0) Then 
					User = Left(Message, InStr(1, Message, MSG_UNBANNED, CompareMethod.Binary) - 1)
					
					If (Len(User) > 0) Then
						g_Channel.BanCount = (g_Channel.BanCount - 1)
						
						Do 
							banpos = g_Channel.IsOnBanList(User)
							
							If (banpos > 0) Then
								g_Channel.Banlist.Remove(banpos)
							End If
						Loop While (banpos <> 0)
					End If
				End If
				
				'// backup channel
				If (InStr(1, Message, "kicked you out", CompareMethod.Text) > 0) Then
					If (BotVars.UseBackupChan) Then
						If (Len(BotVars.BackupChan) > 0) Then
							frmChat.AddQ("/join " & BotVars.BackupChan)
						End If
					Else
						frmChat.AddQ("/join " & g_Channel.Name)
					End If
				End If
				
				If (InStr(1, Message, " has been unsquelched", CompareMethod.Text) > 0) Then
					If ((g_Channel.IsSilent) And (frmChat.mnuDisableVoidView.Checked = False)) Then
						frmChat.lvChannel.Items.Clear()
					End If
				End If
			End If
			
			If (InStr(1, Message, "designated heir", CompareMethod.Text) <> 0) Then
				g_Channel.OperatorHeir = Left(Message, Len(Message) - 29)
			End If
			
			
			temp = "Your friends are:"
			
			If (StrComp(Left(Message, Len(temp)), temp) = 0) Then
				If (Not (BotVars.ShowOfflineFriends)) Then
					Message = Message & "  ÿci(StealthBot is hiding your offline friends)"
				End If
			End If
			
		End If ' message contains a space
		
		If (StrComp(Right(Message, 9), ", offline", CompareMethod.Text) = 0) Then
			If (BotVars.ShowOfflineFriends) Then
				frmChat.AddChat(RTBColors.ServerInfoText, Message)
			End If
		Else
			If (Not (bHide)) Then
				frmChat.AddChat(RTBColors.ServerInfoText, Message)
			End If
		End If
		
		RunInAll("Event_ServerInfo", Message)
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_ServerInfo()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_UserEmote(ByVal Username As String, ByVal Flags As Integer, ByVal Message As String, Optional ByRef QueuedEventID As Short = 0)
		
		On Error GoTo ERROR_HANDLER
		
		Dim UserEvent As clsUserEventObj
		Dim UserObj As clsUserObj
		
		Dim i As Short
		Dim ToANSI As String
		Dim pos As Short
		Dim PassedQueue As Boolean
		
		pos = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		If (pos > 0) Then
			UserObj = g_Channel.Users.Item(pos)
			
			If (QueuedEventID = 0) Then
				UserObj.LastTalkTime = UtcNow
				
				If (UserObj.Queue.Count() > 0) Then
					UserEvent = New clsUserEventObj
					
					With UserEvent
						.EventID = ID_EMOTE
						.Flags = Flags
						.Message = Message
					End With
					
					UserObj.Queue.Add(UserEvent)
				End If
			End If
		Else
			' create new user object for invisible representatives...
			UserObj = New clsUserObj
			
			' store user name
			UserObj.Name = Username
		End If
		
		' convert user name
		Username = UserObj.DisplayName
		
		If (frmChat.mnuUTF8.Checked) Then
			ToANSI = UTF8Decode(Message)
			
			If (Len(ToANSI) > 0) Then
				Message = ToANSI
			End If
		End If
		
		If (QueuedEventID = 0) Then
			If (g_Channel.Self.IsOperator) Then
				If (GetSafelist(Username) = False) Then
					CheckMessage(Username, Message)
				End If
			End If
		End If
		
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
			If (AllowedToTalk(Username, Message)) Then
				'If (GetVeto = False) Then
				frmChat.AddChat(RTBColors.EmoteText, "<", RTBColors.EmoteUsernames, Username & Space(1), RTBColors.EmoteText, Message & ">")
				'End If
				
				If (Catch_Renamed(0) <> vbNullString) Then
					CheckPhrase(Username, Message, CPEMOTE)
				End If
				
				If (frmChat.mnuFlash.Checked) Then
					FlashWindow()
				End If
			End If
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
				If (StrComp(Left(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, CompareMethod.Binary) = 0) Then
					
					Message = BotVars.Trigger & Mid(Message, Len(BotVars.TriggerLong) + 1)
				End If
			End If
			
			RunInAll("Event_UserEmote", Username, Flags, Message)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_UserEmote()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_UserInChannel(ByVal Username As String, ByVal Flags As Integer, ByVal Statstring As String, ByVal Ping As Integer, Optional ByRef QueuedEventID As Short = 0)
		On Error GoTo ERROR_HANDLER
		
		Dim UserEvent As clsUserEventObj
		Dim UserObj As clsUserObj
		Dim found As System.Windows.Forms.ListViewItem
		
		Dim UserIndex As Short
		Dim i As Short
		Dim strCompare As String
		Dim Level As Byte
		Dim StatUpdate As Boolean
		Dim Index As Integer
		Dim Stats As String
		Dim Clan As String
		Dim pos As Short
		Dim showUpdate As Boolean
		Dim Displayed As Boolean ' whether this event has been displayed in the RTB (if combined with another)
		Dim AcqOps As Boolean
		Dim NewIcon As Integer ' temp store new icon
		
        If (Len(Username) < 1) Then
            Exit Sub
        End If
		
		UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		If (UserIndex > 0) Then
			
			UserObj = g_Channel.Users.Item(UserIndex)
			
			If (QueuedEventID = 0) Then
				If (UserObj.Queue.Count() > 0) Then
					If (UserObj.Stats.Statstring = vbNullString) Then
						showUpdate = True
					End If
					
					UserEvent = New clsUserEventObj
					
					With UserEvent
						.EventID = ID_USER
						.Flags = Flags
						.Ping = Ping
						.GameID = UserObj.Game
						.Clan = UserObj.Clan
						.Statstring = Statstring
					End With
					
					UserObj.Queue.Add(UserEvent)
				End If
			End If
			
			StatUpdate = True
		Else
			UserObj = New clsUserObj
		End If
		
		With UserObj
			.Name = Username
			.Flags = Flags
			.Ping = Ping
			.JoinTime = g_Channel.JoinTime
			.Statstring = Statstring
		End With
		
		If (UserIndex = 0) Then
			g_Channel.Users.Add(UserObj)
		End If
		
		Username = UserObj.DisplayName
		
		'ParseStatstring OriginalStatstring, Stats, Clan
		Dim UserColor As Integer
		Dim FDesc As String
		If (StatUpdate = False) Then
			'frmChat.AddChat vbRed, UserObj.Stats.IconCode
			
			AddName(Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan)
			
			frmChat.lblCurrentChannel.Text = frmChat.GetChannelString()
			
			frmChat.ListviewTabs_SelectedIndexChanged(Nothing, New System.EventArgs())
			
			DoLastSeen(Username)
		Else
			If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
				If (JoinMessagesOff = False) Then
					' default to display this event
					Displayed = False
					
					' check whether it has been
					If QueuedEventID > 0 And UserObj.Queue.Count() >= QueuedEventID Then
						UserEvent = UserObj.Queue.Item(QueuedEventID)
						Displayed = UserEvent.Displayed
					End If
					
					' display if it has not already been
					If Not Displayed Then
						FDesc = FlagDescription(Flags, False)
						
                        If Len(FDesc) > 0 Then
                            FDesc = " as a " & FDesc
                        End If
						
						' display message
						If (Flags And USER_BLIZZREP) Then
							UserColor = RGB(97, 105, 255)
						ElseIf (Flags And USER_SYSOP) Then 
							UserColor = RGB(97, 105, 255)
						ElseIf (Flags And USER_CHANNELOP) Then 
							UserColor = RTBColors.TalkUsernameOp
						Else
							UserColor = RTBColors.JoinUsername
						End If
						
						frmChat.AddChat(RTBColors.JoinText, "-- Stats updated: ", UserColor, Username, RTBColors.JoinUsername, " [" & Ping & "ms]", RTBColors.JoinText, " is using " & UserObj.Stats.ToString_Renamed, RTBColors.JoinUsername, FDesc, RTBColors.JoinText, ".")
					End If
				End If
				
				pos = checkChannel(Username)
				
				If (pos > 0) Then
					
					'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					found = frmChat.lvChannel.Items.Item(pos)
					
					' if the update occured to a D2 user ...
					If ((StrComp(UserObj.Game, PRODUCT_D2DV) = 0) Or (StrComp(UserObj.Game, PRODUCT_D2XP) = 0)) Then
						' the username could have changed!
						If (StrComp(UserObj.DisplayName, found.Text, CompareMethod.Binary) <> 0) Then
							' it did, so update user name text in channel list
							found.Text = UserObj.DisplayName
							
							' now check if this is Self
							If (StrComp(UserObj.Name, CleanUsername(CurrentUsername), CompareMethod.Binary) = 0) Then
								' it is! we have to do some magic to tell SB we have a new name
								CurrentUsername = UserObj.Stats.CharacterName & "*" & CleanUsername(CurrentUsername)
								
								' tell scripting
								SharedScriptSupport.myUsername = GetCurrentUsername
								
								' set form title
								SetTitle(GetCurrentUsername & ", online in channel " & g_Channel.Name)
								
								' tell tray icon
								Call frmChat.UpdateTrayTooltip()
							End If
						End If
					End If
					
					' if we are showing stats icons ...
					If (BotVars.ShowStatsIcons) Then 'and the icon code is valid
						If (UserObj.Stats.IconCode <> -1) Then
							' if the icon in the list is not the icon found by stats, update
							NewIcon = GetSmallIcon(UserObj.Game, UserObj.Flags, UserObj.Stats.IconCode)
							'UPGRADE_ISSUE: MSComctlLib.ListItem property found.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							If (found.SmallIcon <> NewIcon) Then
								'UPGRADE_ISSUE: MSComctlLib.ListItem property found.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								found.SmallIcon = NewIcon
							End If
						End If
					End If
					
					If (found.SubItems.Count > 0) Then
						'UPGRADE_WARNING: Lower bound of collection found.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						found.SubItems.Item(1).Text = UserObj.Clan
					End If
					
					'UPGRADE_NOTE: Object found may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					found = Nothing
				End If
			End If
		End If
		
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			RunInAll("Event_UserInChannel", Username, Flags, UserObj.Stats.ToString_Renamed, Ping, UserObj.Game, StatUpdate)
		End If
		
		If (MDebug("statstrings")) Then
			frmChat.AddChat(RTBColors.InformationText, "Username: " & Username & ", Statstring: " & Statstring)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_UserInChannel()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_UserJoins(ByVal Username As String, ByVal Flags As Integer, ByVal Statstring As String, ByVal Ping As Integer, Optional ByRef QueuedEventID As Short = 0)
		
		On Error GoTo ERROR_HANDLER
		
		Dim UserObj As clsUserObj
		Dim UserEvent As clsUserEventObj
		
		Dim toCheck As String
		Dim strCompare As String
		Dim i As Integer
		Dim temp As Byte
		Dim Level As Byte
		Dim L As Integer
		Dim Banned As Boolean
		Dim f As Short
		Dim UserIndex As Short
		Dim BanningUser As Boolean
		Dim pStats As String
		Dim IsBanned As Boolean
		Dim AcqFlags As Integer
		Dim ToDisplay As Boolean
		
		If (Len(Username) < 1) Then
			Exit Sub
		End If
		
		UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		If (QueuedEventID > 0) Then
			If (UserIndex = 0) Then
				frmChat.AddChat(RTBColors.ErrorMessageText, "Error: We have received a queued join event for a user that we " & "couldn't find in the channel.")
				
				Exit Sub
			End If
			
			UserObj = g_Channel.Users.Item(UserIndex)
		Else
			If (UserIndex = 0) Then
				UserObj = New clsUserObj
				
				With UserObj
					.Name = Username
					.Flags = Flags
					.Ping = Ping
					.JoinTime = UtcNow
					.Statstring = Statstring
				End With
				
				If (BotVars.ChatDelay > 0) Then
					UserEvent = New clsUserEventObj
					
					With UserEvent
						.EventID = ID_JOIN
						.Flags = Flags
						.Ping = Ping
						.GameID = UserObj.Game
						.Statstring = Statstring
						.Clan = UserObj.Clan
						.IconCode = UserObj.Stats.Icon
					End With
					
					UserObj.Queue.Add(UserEvent)
				End If
				
				g_Channel.Users.Add(UserObj)
			Else
				frmChat.AddChat(RTBColors.ErrorMessageText, "Warning: We have received a join event for a user that we had thought was " & "already present within the channel.  This may be indicative of a server split or other technical difficulty.")
				
				Exit Sub
			End If
		End If
		
		Username = UserObj.DisplayName
		
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID = 0)) Then
			If (g_Channel.Self.IsOperator) Then
				g_Channel.CheckUser(Username, UserObj)
			End If
		End If
		
		Dim UserColor As Integer
		Dim FDesc As String
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' GUI
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			' if we have join/leaves on
			If (JoinMessagesOff = False) Then
				
				' does this event have events delayed after it?
				If QueuedEventID > 0 And UserObj.Queue.Count() > 0 Then
					
					' loop through the events occuring after this one
					For i = QueuedEventID To UserObj.Queue.Count()
						
						' get the event
						UserEvent = UserObj.Queue.Item(i)
						
						' default to not combine with userjoins
						ToDisplay = False
						
						Select Case UserEvent.EventID
							
							' user flags update
							Case ID_USERFLAGS
								' will combine with userjoins
								ToDisplay = True
								
								AcqFlags = UserEvent.Flags
								
								' user stats update / user in channel
							Case ID_USER
								' will combine with userjoins
								ToDisplay = True
								
								' is stats different / provided?
                                If Len(UserEvent.Statstring) > 0 Then
                                    If StrComp(UserEvent.Statstring, UserObj.Statstring) Then

                                        ' store stats update stats in object used in userjoins message generation
                                        UserObj.Statstring = UserEvent.Statstring
                                    End If
                                End If
								
						End Select
						
						' if we're going to combine this event with userjoins ...
						If ToDisplay Then 'then set .displayed on the queue'd event so it is not displayed separately
							UserEvent.Displayed = True
							
							' also update in collection
							UserObj.Queue.Remove(i)
							UserObj.Queue.Add(UserEvent,  ,  , i - 1)
						End If
						
					Next i
					
				End If
				
				If (Not CheckBlock(Username)) Then
					FDesc = FlagDescription(AcqFlags Or Flags, False)
					
                    If Len(FDesc) > 0 Then
                        FDesc = " as a " & FDesc
                    End If
					
					' display message
					If (AcqFlags And USER_BLIZZREP) Or (Flags And USER_BLIZZREP) Then
						UserColor = RGB(97, 105, 255)
					ElseIf (AcqFlags And USER_SYSOP) Or (Flags And USER_SYSOP) Then 
						UserColor = RGB(97, 105, 255)
					ElseIf (AcqFlags And USER_CHANNELOP) Or (Flags And USER_CHANNELOP) Then 
						UserColor = RTBColors.TalkUsernameOp
					Else
						UserColor = RTBColors.JoinUsername
					End If
					
					frmChat.AddChat(RTBColors.JoinText, "-- ", UserColor, Username, RTBColors.JoinUsername, " [" & Ping & "ms]", RTBColors.JoinText, " has joined the channel using " & UserObj.Stats.ToString_Renamed, RTBColors.JoinUsername, FDesc, RTBColors.JoinText, ".")
				End If
			End If
			
			' add to user list
			AddName(Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan)
			
			' update caption
			frmChat.lblCurrentChannel.Text = frmChat.GetChannelString
			
			' focus on channel tab
			frmChat.ListviewTabs_SelectedIndexChanged(Nothing, New System.EventArgs())
			
			' flash window
			If (frmChat.mnuFlash.Checked) Then
				FlashWindow()
			End If
			
			' update last seen info
			Call DoLastSeen(Username)
			
			' check is banned
			IsBanned = (UserObj.PendingBan)
			
			'frmChat.AddChat vbRed, IsBanned
			
			' if not banned...
			If (IsBanned = False) Then
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				' Greet message
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				
				If (BotVars.UseGreet) Then
                    If (Len(BotVars.GreetMsg)) Then
                        If (BotVars.WhisperGreet) Then
                            frmChat.AddQ("/w " & Username & Space(1) & DoReplacements(BotVars.GreetMsg, Username, Ping))
                        Else
                            frmChat.AddQ(DoReplacements(BotVars.GreetMsg, Username, Ping))
                        End If
                    End If
				End If
				
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				' Botmail
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				
				If (mail) Then
					L = GetMailCount(Username)
					
					If (L > 0) Then
						frmChat.AddQ("/w " & Username & " You have " & L & " new message" & IIf(L = 1, "", "s") & ". Type !inbox to retrieve.")
					End If
				End If
			End If
			
			' print their statstring, if desired
			If (MDebug("statstrings")) Then
				frmChat.AddChat(RTBColors.ErrorMessageText, Statstring)
			End If
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			'frmChat.AddChat vbRed, frmChat.SControl.Error.Number
			
			RunInAll("Event_UserJoins", Username, Flags, UserObj.Stats.ToString_Renamed, Ping, UserObj.Game, UserObj.Stats.Level, Statstring, IsBanned)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_UserJoins()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_UserLeaves(ByVal Username As String, ByVal Flags As Integer)
		On Error GoTo ERROR_HANDLER
		
		Dim UserObj As clsUserObj
		
		Dim UserIndex As Short
		Dim i As Short
		Dim ii As Short
		Dim Holder() As Object
		Dim pos As Short
		Dim bln As Boolean
		
		UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		Dim UserColor As Integer
		If (UserIndex > 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(UserIndex).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (g_Channel.Users.Item(UserIndex).IsOperator) Then
				g_Channel.RemoveBansFromOperator(Username)
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(UserIndex).Queue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (g_Channel.Users.Item(UserIndex).Queue.Count = 0) Then
				If ((Not JoinMessagesOff) And (Not CheckBlock(Username))) Then
					'If (GetVeto = False) Then
					
					' display message
					If (Flags And USER_BLIZZREP) Then
						UserColor = RGB(97, 105, 255)
					ElseIf (Flags And USER_SYSOP) Then 
						UserColor = RGB(97, 105, 255)
					ElseIf (Flags And USER_CHANNELOP) Then 
						UserColor = RTBColors.TalkUsernameOp
					Else
						UserColor = RTBColors.JoinUsername
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.AddChat(RTBColors.JoinText, "-- ", UserColor, g_Channel.Users.Item(UserIndex).DisplayName, RTBColors.JoinText, " has left the channel.")
					'End If
				End If
			End If
			
			g_Channel.Users.Remove(UserIndex)
		Else
			frmChat.AddChat(RTBColors.ErrorMessageText, "Warning: We have received a leave event for a user that we didn't know " & "was in the channel.  This may be indicative of a server split or other technical difficulty.")
			
			Exit Sub
		End If
		
		If (StrComp(Username, g_Channel.OperatorHeir, CompareMethod.Text) = 0) Then
			g_Channel.OperatorHeir = vbNullString
			
			Call g_Channel.CheckUsers()
		End If
		
		Username = ConvertUsername(Username)
		
		RemoveBanFromQueue(Username)
		
		pos = checkChannel(Username)
		
		If (pos > 0) Then
			If (frmChat.mnuFlash.Checked) Then
				FlashWindow()
			End If
			
			With frmChat.lvChannel
				.Items.RemoveAt(pos)
				
				.Refresh()
			End With
			
			frmChat.lblCurrentChannel.Text = frmChat.GetChannelString()
			
			frmChat.ListviewTabs_SelectedIndexChanged(Nothing, New System.EventArgs())
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			RunInAll("Event_UserLeaves", CleanUsername(Username), Flags)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_UserLeaves()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_UserTalk(ByVal Username As String, ByVal Flags As Integer, ByVal Message As String, ByVal Ping As Integer, Optional ByRef QueuedEventID As Short = 0)
		
		On Error GoTo ERROR_HANDLER
		
		Dim UserObj As clsUserObj
		Dim UserEvent As clsUserEventObj
		
		Dim strSend As String
		Dim s As String
		Dim U As String
		Dim strCompare As String
		Dim i As Short
		Dim ColIndex As Short
		Dim B As Boolean
		Dim ToANSI As String
		Dim BanningUser As Boolean
		Dim UsernameColor As Integer
		Dim TextColor As Integer
		Dim CaratColor As Integer
		Dim pos As Short
		Dim blnCheck As Boolean
		
		pos = g_Channel.GetUserIndexEx(CleanUsername(Username))
		
		If (pos > 0) Then
			UserObj = g_Channel.Users.Item(pos)
			
			UserObj.LastTalkTime = UtcNow
			
			If (QueuedEventID = 0) Then
				If (UserObj.Queue.Count() > 0) Then
					UserEvent = New clsUserEventObj
					
					With UserEvent
						.EventID = ID_TALK
						.Flags = Flags
						.Ping = Ping
						.Message = Message
					End With
					
					UserObj.Queue.Add(UserEvent)
				End If
			End If
		Else
			' create new user object for invisible representatives...
			UserObj = New clsUserObj
			
			' store user name
			UserObj.Name = Username
		End If
		
		' convert user name
		Username = UserObj.DisplayName
		
		If (frmChat.mnuUTF8.Checked) Then
			ToANSI = UTF8Decode(Message)
			
			If (Len(ToANSI) > 0) Then
				Message = ToANSI
			End If
		End If
		
		If (QueuedEventID = 0) Then
			If (g_Channel.Self.IsOperator) Then
				If (GetSafelist(Username) = False) Then
					CheckMessage(Username, Message)
				End If
			End If
		End If
		
		If ((UserObj.Queue.Count() = 0) Or (QueuedEventID > 0)) Then
			If (Message <> vbNullString) Then
				If (AllowedToTalk(Username, Message)) Then
					' are we watching the user?
					'If (StrComp(WatchUser, Username, vbTextCompare) = 0) Then
					If (PrepareCheck(Username) Like PrepareCheck(WatchUser)) Then
						UsernameColor = RTBColors.ErrorMessageText
						
						' is user an operator?
					ElseIf ((Flags And USER_CHANNELOP) = USER_CHANNELOP) Then 
						UsernameColor = RTBColors.TalkUsernameOp
					Else
						UsernameColor = RTBColors.TalkUsernameNormal
					End If
					
					If (((Flags And USER_BLIZZREP) = USER_BLIZZREP) Or ((Flags And USER_SYSOP) = USER_SYSOP)) Then
						
						TextColor = RGB(97, 105, 255)
						
						CaratColor = RGB(97, 105, 255)
					Else
						TextColor = RTBColors.TalkNormalText
						
						CaratColor = RTBColors.Carats
					End If
					
					'If (GetVeto = False) Then
					frmChat.AddChat(CaratColor, "<", UsernameColor, Username, CaratColor, "> ", TextColor, Message)
					'End If
					
					If (Catch_Renamed(0) <> vbNullString) Then
						CheckPhrase(Username, Message, CPTALK)
					End If
					
					If (frmChat.mnuFlash.Checked) Then
						FlashWindow()
					End If
				End If
			End If
			
			If (VoteDuration > 0) Then
				If (InStr(1, LCase(Message), "yes", CompareMethod.Text) > 0) Then
					Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
				ElseIf (InStr(1, LCase(Message), "no", CompareMethod.Text) > 0) Then 
					Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
				End If
			End If
			
			Call ProcessCommand(Username, Message, False, False)
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' call event script function
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			On Error Resume Next
			
			If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
				If (StrComp(Left(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, CompareMethod.Binary) = 0) Then
					
					Message = BotVars.Trigger & Mid(Message, Len(BotVars.TriggerLong) + 1)
				End If
			End If
			
			RunInAll("Event_UserTalk", Username, Flags, Message, Ping)
		End If
		
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_UserTalk()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Private Function CheckMessage(ByRef Username As String, ByRef Message As String) As Boolean
		On Error GoTo ERROR_HANDLER
		Dim BanningUser As Boolean
		Dim i As Short
		
		If (PhraseBans) Then
			For i = LBound(Phrases) To UBound(Phrases)
				If ((Phrases(i) <> vbNullString) And (Phrases(i) <> Space(1))) Then
					If ((InStr(1, Message, Phrases(i), CompareMethod.Text)) <> 0) Then
						Ban(Username & " Banned phrase: " & Phrases(i), AutoModSafelistValue - 1, System.Math.Abs(CInt(Config.PhraseKick)))
						
						BanningUser = True
						
						Exit For
					End If
				End If
			Next i
		End If
		
		If (BanningUser = False) Then
			If (BotVars.QuietTime) Then
				Ban(Username & " Quiet-time", AutoModSafelistValue - 1, System.Math.Abs(CInt(Config.QuietTimeKick)))
			Else
				If (BotVars.KickOnYell = 1) Then
					If (Len(Message) > 5) Then
						If (PercentActualUppercase(Message) > 90) Then
							Ban(Username & " Yelling", AutoModSafelistValue - 1, 1)
						End If
					End If
				End If
			End If
			
			If ((BotVars.QuietTime) Or (BotVars.KickOnYell = 1)) Then
				BanningUser = True
			End If
		End If
		
		CheckMessage = BanningUser
		
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.CheckMessage()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Sub Event_VersionCheck(ByRef Message As Integer, ByRef ExtraInfo As String)
		On Error GoTo ERROR_HANDLER
		Select Case (Message)
			Case 0
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Client version accepted!")
				
				' if using server finder
				If ((BotVars.BNLS) And (BotVars.UseAltBnls)) Then
					' save BNLS server so future instances of the bot won't need to get the list, connection succeeded
					If Config.BNLSServer <> BotVars.BNLSServer Then
						Config.BNLSServer = BotVars.BNLSServer
						Call Config.Save()
					End If
				End If
				
			Case 1
                frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Version check failed! " & "The version byte for this attempt was 0x" & Hex(GetVerByte((BotVars.Product))) & "." & IIf(Len(ExtraInfo) = 0, vbNullString, " Extra Information: " & ExtraInfo))
				
				If (BotVars.BNLS) Then
					If (frmChat.HandleBnlsError("[BNCS] BNLS has not been updated yet, " & "or you experienced an error. Try connecting again.")) Then
						' if we are using the finder, then don't close all connections
						Message = 0
					End If
				Else
					frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Please ensure you " & "have updated your hash files using more current ones from the directory " & "of the game you're connecting with.")
					
					frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] In addition, you can try " & "choosing ""Update version bytes from StealthBot.net"" from the Bot menu.")
				End If
				
			Case 2
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your CD-key is invalid!")
				
			Case 3
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Version check failed! " & "BNLS has not been updated yet.. Try reconnecting in an hour or two.")
				
			Case 4
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your CD-key is for another game.")
				
			Case 5
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your CD-key is banned. " & "For more information, visit http://us.blizzard.com/support/article.xml?locale=en_US&articleId=20637 .")
				
			Case 6
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your CD-key is currently in " & "use under the owner name: " & ExtraInfo & ".")
				
			Case 7
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is invalid.")
				
			Case 8
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is currently " & "in use under the owner name: " & ExtraInfo & ".")
				
			Case 9
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is banned. " & "For more information, visit http://us.blizzard.com/support/article.xml?locale=en_US&articleId=20637 .")
				
			Case 10
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Version check passed!")
				
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is for another game.")
				
			Case Else
				frmChat.AddChat(RTBColors.ErrorMessageText, "Unhandled 0x51 response! Value: " & Message)
		End Select
		
		If (Message > 0) Then
			Call frmChat.DoDisconnect()
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_VersionCheck()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub Event_WhisperFromUser(ByVal Username As String, ByVal Flags As Integer, ByVal Message As String, ByVal Ping As Integer)
		On Error GoTo ERROR_HANDLER
		'Dim s       As String
		Dim lCarats As Integer
		Dim WWIndex As Short
		Dim ToANSI As String
		
		Username = ConvertUsername(Username)
		
		ToANSI = UTF8Decode(Message)
		
		If (Len(ToANSI) > 0) Then
			Message = ToANSI
		End If
		
		If (frmChat.mnuUTF8.Checked) Then
			Message = ToANSI
			
			If (Message = vbNullString) Then
				Exit Sub
			End If
		End If
		
		If (Catch_Renamed(0) <> vbNullString) Then
			Call CheckPhrase(Username, Message, CPWHISPER)
		End If
		
		If (frmChat.mnuFlash.Checked) Then
			FlashWindow()
		End If
		
		If (StrComp(Message, BotVars.ChannelPassword, CompareMethod.Text) = 0) Then
			lCarats = g_Channel.GetUserIndex(Username)
			
			If (lCarats > 0) Then
				With g_Channel.Users.Item(lCarats)
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().PassedChannelAuth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.PassedChannelAuth = True
				End With
				
				frmChat.AddQ("/w " & Username & " Password accepted.")
			End If
		End If
		
		If (VoteDuration > 0) Then
			If (InStr(1, Message, "yes", CompareMethod.Text) > 0) Then
				Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
			ElseIf (InStr(1, Message, "no", CompareMethod.Text) > 0) Then 
				Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
			End If
		End If
		
		lCarats = RTBColors.WhisperCarats
		
		If (Flags And &H1) Then
			lCarats = COLOR_BLUE
		End If
		
		'####### Mail check
		Dim Msg As udtMail
		If (mail) Then
			If (StrComp(Left(Message, 6), "!inbox", CompareMethod.Text) = 0) Then
				
				If (GetMailCount(Username) > 0) Then
					Call GetMailMessage(Username, Msg)
					
					If (Len(RTrim(Msg.To_Renamed)) > 0) Then
						frmChat.AddQ("/w " & Username & " Message from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message))
					End If
				End If
			End If
		End If
		'#######
		
		If ((Not (CheckMsg(Message, Username, -5))) And (Not (CheckBlock(Username)))) Then
			
			If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
				frmChat.AddChat(lCarats, "<From ", RTBColors.WhisperUsernames, Username, lCarats, "> ", RTBColors.WhisperText, Message)
			End If
			
			frmChat.AddWhisper(lCarats, "<From ", RTBColors.WhisperUsernames, Username, lCarats, "> ", RTBColors.WhisperText, Message)
			
			frmChat.rtbWhispers.Visible = rtbWhispersVisible
			
			If (frmChat.mnuToggleWWUse.Checked) Then
				'If ((frmChat.mnuToggleWWUse.Checked) And _
				''(frmChat.WindowState <> vbMinimized)) Then
				
				If (Not (IrrelevantWhisper(Message, Username))) Then
					WWIndex = AddWhisperWindow(Username)
					
					With colWhisperWindows.Item(WWIndex)
						'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item(WWIndex).Shown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (.Shown = False) Then
							'window was previously hidden
							
							ShowWW(WWIndex)
						End If
						
						'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Caption = "Whisper Window: " & Username
						
						'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().AddWhisper. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.AddWhisper(RTBColors.WhisperUsernames, "> " & Username, lCarats, ": ", RTBColors.WhisperText, Message)
					End With
				End If
			End If
			
			Call ProcessCommand(Username, Message, False, True)
		End If
		
		If (Not (CheckBlock(Username))) Then
			LastWhisper = Username
			LastWhisperFromTime = Now
		End If
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' call event script function
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		If BotIsClosing Then Exit Sub
		
		On Error Resume Next
		
		g_lastQueueUser = Username
		
		RunInAll("Event_WhisperFromUser", Username, Flags, Message, Ping)
		'End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_WhisperFromUser()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' Flags and ping are deliberately not used at this time
	Public Sub Event_WhisperToUser(ByVal Username As String, ByVal Flags As Integer, ByVal Message As String, ByVal Ping As Integer)
		On Error GoTo ERROR_HANDLER
		Dim WWIndex As Short
		Dim ToANSI As String
		
		ToANSI = UTF8Decode(Message)
		
		If (Len(ToANSI) > 0) Then
			Message = ToANSI
		End If
		
		'frmChat.AddChat vbRed, Username
		
		If (StrComp(Username, "your friends", CompareMethod.Text) <> 0) Then
			Username = ConvertUsername(Username)
			
			LastWhisperTo = Username
		Else
			LastWhisperTo = "%f%"
		End If
		
		If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
			frmChat.AddChat(RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, Username, RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message)
		End If
		
		If ((frmChat.mnuHideWhispersInrtbChat.Checked) Or (frmChat.mnuToggleShowOutgoing.Checked)) Then
			
			frmChat.AddWhisper(RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, Username, RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message)
		End If
		
		If (frmChat.mnuToggleWWUse.Checked) Then
			If ((InStr(1, Message, "ß~ß") = 0) And (StrComp(Username, "your friends") <> 0)) Then
				
				WWIndex = AddWhisperWindow(Username)
				
				If (frmChat.WindowState <> System.Windows.Forms.FormWindowState.Minimized) Then
					Call ShowWW(WWIndex)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				colWhisperWindows.Item(WWIndex).Caption = "Whisper Window: " & Username
				'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().AddWhisper. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				colWhisperWindows.Item(WWIndex).AddWhisper(RTBColors.TalkBotUsername, "> " & GetCurrentUsername, RTBColors.WhisperCarats, ": ", RTBColors.WhisperText, Message)
			End If
		End If
		
		If (Not (rtbWhispersVisible)) Then
			If (frmChat.rtbWhispers.Visible = True) Then
				frmChat.rtbWhispers.Visible = False
			End If
		End If
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_WhisperToUser()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	
	'11/22/07 - Hdx - Pass the channel listing (0x0B) directly off to scriptors for there needs. (What other use is there?)
	Public Sub Event_ChannelList(ByRef sChannels() As String)
		On Error GoTo ERROR_HANDLER
		Dim x As Short
		Dim sChannel As String
		
		If (MDebug("all")) Then
			frmChat.AddChat(RTBColors.InformationText, "Received Channel List: ")
		End If
		
		' save public channels
		BotVars.PublicChannels = New Collection
		
		For x = 0 To UBound(sChannels)
			sChannel = sChannels(x)
			
            If Len(sChannel) > 0 Then
                BotVars.PublicChannels.Add(sChannel)
            End If
		Next x
		
		PreparePublicChannelMenu()
		
		RunInAll("Event_ChannelList", New Object(){ConvertStringArray(sChannels)})
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_ChannelList()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'10/01/09 - Hdx - This is for SID_MESSAGEBOX, for now it'll raise it's own event, and Event_ServerError
	Public Function Event_MessageBox(ByRef lStyle As Integer, ByRef sText As String, ByRef sCaption As String) As Object
		On Error GoTo ERROR_HANDLER
		Call Event_ServerError(sText)
		
		RunInAll("Event_MessageBox", lStyle, sText, sCaption)
		
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Event_MessageBox()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Function CleanUsername(ByVal Username As String, Optional ByVal PrependNamingStar As Boolean = False) As String
		On Error GoTo ERROR_HANDLER
		Dim tmp As String
		Dim pos As Short
		
		tmp = Username
		
		If (tmp <> vbNullString) Then
			pos = InStr(1, tmp, "*", CompareMethod.Binary)
			
			If (pos > 0) Then
				If (Right(tmp, 1) = ")") Then
					' fixed so that usernames actually ending in
					' ")" don't get trimmed (ultimately messing up
					' such bots with ops). ~Ribose/2009-11-15
					If pos > 3 Then
						' blah (*blah)
						'     ^^^
						If Mid(tmp, pos - 2, 3) = " (*" Then
							tmp = Left(tmp, Len(tmp) - 1)
						End If
					End If
				End If
				
				tmp = Mid(tmp, pos + 1)
			End If
		End If
		
		If (Dii And PrependNamingStar And BotVars.UseD2Naming = False) Then
			tmp = "*" & tmp
		End If
		
		CleanUsername = tmp
		
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.CleanUsername()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	'Private Function GetDiablo2CharacterName(ByVal Username As String) As String
	'
	'    Dim tmp As String
	'    Dim Pos As Integer
	'
	'    Pos = InStr(1, Username, "*", vbBinaryCompare)
	'
	'    If (Pos > 0) Then
	'        tmp = Mid$(Username, 1, Pos - 1)
	'    End If
	'
	'    GetDiablo2CharacterName = tmp
	'
	'End Function
End Module