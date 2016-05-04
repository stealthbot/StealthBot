Option Strict Off
Option Explicit On
Module modCommandsAdmin
	'This module will hold all the commands that relate to Andmistering the bot, Changing settings
	'Editing the database, etc..
	
	'This is a stub function for now, it still calls the old uber complicated OnAddOld function, but hey :/
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAdd(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim response() As String
		Dim i As Short
		ReDim Preserve response(0)
		
        If ((Not Command_Renamed.IsValid) Or Len(Trim(Command_Renamed.Argument("username"))) = 0) Then
            Command_Renamed.Respond("You must specify a user to add.")
            Exit Sub
        End If
		
		' special case: d2 naming conventions
		Dim Username As Object
		Dim User As clsUserObj
		Dim IsChar As Boolean
		Dim Acct As String
		If (BotVars.UseD2Naming) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Username = Command_Renamed.Argument("Username")
			If (Len(Username) > 1) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Left(Username, 1) = "*") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (InStr(2, Username, "*") = 0 And InStr(2, Username, "?") = 0) Then
						' format: *user
						' assume user means: user
						Command_Renamed.Args = Mid(Command_Renamed.Args, 2)
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (InStr(Username, "*") = 0 And InStr(Username, "?") = 0) Then
					' format: charname
					For i = 1 To g_Channel.Users.Count()
						User = g_Channel.Users.Item(i)
						'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (StrComp(User.CharacterName, Username, CompareMethod.Text) = 0) Then
							' the user provided is a character in the channel
							IsChar = True
							Acct = User.DisplayName
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (StrComp(User.DisplayName, Username, CompareMethod.Text) = 0) Then
							' if the user provided is ALSO an accountname, assume account name
							IsChar = False
							Exit For
						End If
					Next i
					If (IsChar) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Username. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Command_Renamed.Args = Replace(Command_Renamed.Args, Username, Acct, 1, 1, CompareMethod.Binary)
					End If
				End If
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dbAccess = GetCumulativeAccess(Command_Renamed.Username)
		If (Command_Renamed.IsLocal) Then
			dbAccess.Rank = 201
			dbAccess.Flags = "A"
		End If
		
		Call OnAddOld(Command_Renamed.Username, dbAccess, Command_Renamed.Args, Command_Renamed.IsLocal, response)
		
		For i = LBound(response) To UBound(response)
			Command_Renamed.Respond(response(i))
		Next i
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnClear(ByRef Command_Renamed As clsCommandObj)
		frmChat.ClearChatScreen()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDisable(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Module_Renamed As MSScriptControl.Module
		Dim Name As String
		Dim i As Short
		Dim mnu As clsMenuObj
		
		If modScripting.GetScriptSystemDisabled() Then
			Command_Renamed.Respond("Error: Scripts are globally disabled via the override.")
			Exit Sub
		End If
		
		If (Command_Renamed.IsValid) Then
			Module_Renamed = modScripting.GetModuleByName(Command_Renamed.Argument("Script"))
			If (Module_Renamed Is Nothing) Then
				Command_Renamed.Respond("Error: Could not find specified script.")
			Else
				Name = modScripting.GetScriptName((Module_Renamed.Name))
				If (StrComp(SharedScriptSupport.GetSettingsEntry("Enabled", Name), "False", CompareMethod.Text) = 0) Then
					Command_Renamed.Respond(Name & " is already disabled.")
				Else
					RunInSingle(Module_Renamed, "Event_Close")
					SharedScriptSupport.WriteSettingsEntry("Enabled", "False",  , Name)
					modScripting.DestroyObjs(Module_Renamed)
					Command_Renamed.Respond(Name & " has been disabled.")
					For i = 1 To DynamicMenus.Count()
						mnu = DynamicMenus(i)
						If (StrComp(mnu.Name, Chr(0) & Name & Chr(0) & "ENABLE|DISABLE", CompareMethod.Text) = 0) Then
							mnu.Checked = False
							Exit For
						End If
					Next i
					'UPGRADE_NOTE: Object mnu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					mnu = Nothing
				End If
			End If
		Else
			Command_Renamed.Respond("Error: You must specify a script.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDump(ByRef Command_Renamed As clsCommandObj)
		Call DumpPacketCache()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnEnable(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Module_Renamed As MSScriptControl.Module
		Dim Name As String
		Dim i As Short
		Dim mnu As clsMenuObj
		
		If modScripting.GetScriptSystemDisabled() Then
			Command_Renamed.Respond("Error: Scripts are globally disabled via the override.")
			Exit Sub
		End If
		
		If (Command_Renamed.IsValid) Then
			Module_Renamed = modScripting.GetModuleByName(Command_Renamed.Argument("Script"))
			If (Module_Renamed Is Nothing) Then
				Command_Renamed.Respond("Error: Could not find specified script.")
			Else
				Name = modScripting.GetScriptName((Module_Renamed.Name))
				If (StrComp(SharedScriptSupport.GetSettingsEntry("Enabled", Name), "True", CompareMethod.Text) = 0) Then
					Command_Renamed.Respond(Name & " is already enabled.")
				Else
					SharedScriptSupport.WriteSettingsEntry("Enabled", "True",  , Name)
					modScripting.InitScript(Module_Renamed)
					Command_Renamed.Respond(Name & " has been enabled.")
					For i = 1 To DynamicMenus.Count()
						mnu = DynamicMenus(i)
						If (StrComp(mnu.Name, Chr(0) & Name & Chr(0) & "ENABLE|DISABLE", CompareMethod.Text) = 0) Then
							mnu.Checked = True
							Exit For
						End If
					Next i
					'UPGRADE_NOTE: Object mnu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					mnu = Nothing
				End If
			End If
		Else
			Command_Renamed.Respond("Error: You must specify a script.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLockText(ByRef Command_Renamed As clsCommandObj)
		Call frmChat.mnuLock_Click(Nothing, New System.EventArgs())
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnQuit(ByRef Command_Renamed As clsCommandObj)
		BotIsClosing = True
		frmChat.Close()
		'UPGRADE_NOTE: Object frmChat may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmChat = Nothing
	End Sub
	
	'This is a stub function for now, it still calls the old uber complicated OnRemOld function, but hey :/
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnRem(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim response() As String
		Dim i As Short
		ReDim Preserve response(0)
		
		If (Not Command_Renamed.IsValid) Then
			Command_Renamed.Respond("You must specify a user to remove.")
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dbAccess = GetCumulativeAccess(Command_Renamed.Username)
		If (Command_Renamed.IsLocal) Then
			dbAccess.Rank = 201
			dbAccess.Flags = "A"
		End If
		
		Call OnRemOld(Command_Renamed.Username, dbAccess, Command_Renamed.Args, Command_Renamed.IsLocal, response)
		
		For i = LBound(response) To UBound(response)
			Command_Renamed.Respond(response(i))
		Next i
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetBnlsServer(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Config.BNLSServer = Command_Renamed.Argument("Server")
			Call Config.Save()
			
			BotVars.BNLSServer = Config.BNLSServer
			Command_Renamed.Respond(StringFormat("New BNLS server set to {0}{1}{0}.", Chr(34), BotVars.BNLSServer))
		Else
			Command_Renamed.Respond("You must specify a server.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetCommandLine(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call SetCommandLine(Command_Renamed.Argument("CommandLine"))
            If (Len(CommandLine) > 0) Then
                Command_Renamed.Respond(StringFormat("Command Line set to: {0}", CommandLine))
            Else
                Command_Renamed.Respond("Command Line cleared.")
            End If
		Else
			Command_Renamed.Respond("You must specify a new command line.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetExpKey(ByRef Command_Renamed As clsCommandObj)
		Dim strKey As String
		If (Command_Renamed.IsValid) Then
			strKey = Replace(Command_Renamed.Argument("Key"), "-", vbNullString)
			strKey = Replace(strKey, Space(1), vbNullString)
			
			If Config.IgnoreCDKeyLength Then
				If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
					strKey = vbNullString
				End If
			End If
            If (Len(strKey) > 0) Then
                Config.ExpKey = strKey
                Call Config.Save()

                BotVars.ExpKey = Config.ExpKey
                Command_Renamed.Respond("New expansion cdkey set.")
            Else
                Command_Renamed.Respond("The cdkey you specified was invalid.")
            End If
		Else
			Command_Renamed.Respond("You must specify a cdkey.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetHome(ByRef Command_Renamed As clsCommandObj)
		Dim Channel As String
		
		Channel = Command_Renamed.Argument("Channel")
		Config.HomeChannel = Channel
		Call Config.Save()
		
		BotVars.HomeChannel = Config.HomeChannel
		PrepareHomeChannelMenu()
        If Len(Channel) = 0 Then
            Command_Renamed.Respond("Home channel set to server default.")
        Else
            Command_Renamed.Respond(StringFormat("New home channel set to {0}{1}{0}.", Chr(34), Config.HomeChannel))
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetKey(ByRef Command_Renamed As clsCommandObj)
		Dim strKey As String
		If (Command_Renamed.IsValid) Then
			strKey = Replace(Command_Renamed.Argument("Key"), "-", vbNullString)
			strKey = Replace(strKey, Space(1), vbNullString)
			
			If Config.IgnoreCDKeyLength Then
				If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
					strKey = vbNullString
				End If
			End If
            If (Len(strKey) > 0) Then
                Config.CDKey = strKey
                Call Config.Save()

                BotVars.CDKey = Config.CDKey
                Command_Renamed.Respond("New CD key set.")
            Else
                Command_Renamed.Respond("The CD key you specified was invalid.")
            End If
		Else
			Command_Renamed.Respond("You must specify a CD key.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetName(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Config.Username = Command_Renamed.Argument("Username")
			Call Config.Save()
			
			BotVars.Username = Config.Username
			Command_Renamed.Respond(StringFormat("New username set to {0}{1}{0}.", Chr(34), BotVars.Username))
		Else
			Command_Renamed.Respond("You must specify a username.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetPass(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Config.Password = Command_Renamed.Argument("Password")
			Call Config.Save()
			
			BotVars.Password = Config.Password
			Command_Renamed.Respond("New password set.")
		Else
			Command_Renamed.Respond("You must specify a password.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetPMsg(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			ProtectMsg = Command_Renamed.Argument("Message")
			
			Config.ChannelProtectionMessage = ProtectMsg
			Call Config.Save()
			
			Command_Renamed.Respond("Channel protection message set.")
		Else
			Command_Renamed.Respond("You must specify a message.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetServer(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Config.Server = Command_Renamed.Argument("Server")
			Call Config.Save()
			
			BotVars.Server = Config.Server
			Command_Renamed.Respond(StringFormat("New server set to {0}{1}{0}.", Chr(34), BotVars.Server))
		Else
			Command_Renamed.Respond("You must specify a server.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetTrigger(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			
			Config.Trigger = Command_Renamed.Argument("Trigger")
			Call Config.Save()
			
			BotVars.Trigger = Config.Trigger
			Command_Renamed.Respond(StringFormat("The new trigger is {0}{1}{0}.", Chr(34), BotVars.Trigger))
		Else
			Command_Renamed.Respond("You must specify a trigger.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWhisperCmds(ByRef Command_Renamed As clsCommandObj)
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on"
				BotVars.WhisperCmds = True
				
				Command_Renamed.Respond("Command responses will now be whispered back.")
				
			Case "off"
				BotVars.WhisperCmds = False
				
				Command_Renamed.Respond("Command responses will now be displayed publicly.")
				
			Case Else
				Command_Renamed.Respond(StringFormat("Command responses are currently {0}.", IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly")))
		End Select
		
		If Config.WhisperCommands <> BotVars.WhisperCmds Then
			Config.WhisperCommands = BotVars.WhisperCmds
			Call Config.Save()
		End If
	End Sub
End Module