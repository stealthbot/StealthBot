Option Strict Off
Option Explicit On
Module modCommandsOps
	'This module will hold all commands that have to deal with holding operator over a channel
	
	Private Const ERROR_NOT_OPS As String = "Error: This command requires channel operator status."
	Public Enum CacheChanneListEnum
		enRetrieve = 0
		enAdd = 1
		enReset = 255
	End Enum
	
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAddPhrase(ByRef Command_Renamed As clsCommandObj)
		Dim sPhrase As String
		Dim i As Short
		Dim iFile As Short
		
		' grab free file handle
		iFile = FreeFile
		If (Command_Renamed.IsValid) Then
			sPhrase = Command_Renamed.Argument("Phrase")
			
			For i = LBound(Phrases) To UBound(Phrases)
				If (StrComp(sPhrase, Phrases(i), CompareMethod.Text) = 0) Then
					Exit For
				End If
			Next i
			
			If (i > UBound(Phrases)) Then
				'Thats a lot of crap.. It check if the last item in Phrases is not just whitespace
                If (Len(Trim(Phrases(UBound(Phrases)))) > 0) Then
                    ReDim Preserve Phrases(UBound(Phrases) + 1)
                End If
				
				Phrases(UBound(Phrases)) = sPhrase
				
				FileOpen(iFile, GetFilePath(FILE_PHRASE_BANS), OpenMode.Output)
				For i = LBound(Phrases) To UBound(Phrases)
                    If (Len(Trim(Phrases(i))) > 0) Then
                        PrintLine(iFile, Phrases(i))
                    End If
				Next i
				FileClose(iFile)
				
				Command_Renamed.Respond(StringFormat("Phraseban {0}{1}{0} added.", Chr(34), sPhrase))
			Else
				Command_Renamed.Respond("Error: That phrase is already banned.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBan(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
			If (g_Channel.Self.IsOperator) Then
				If (InStr(1, Command_Renamed.Argument("Username"), "*", CompareMethod.Binary) = 0) Then
					If (Command_Renamed.IsLocal) Then
						frmChat.AddQ("/ban " & Command_Renamed.Args)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dbAccess = GetCumulativeAccess(Command_Renamed.Username)
						
						Command_Renamed.Respond(Ban(Command_Renamed.Args, dbAccess.Rank))
					End If
				Else
					Command_Renamed.Respond(WildCardBan(Command_Renamed.Argument("Username"), Command_Renamed.Argument("Message"), 1))
				End If
			Else
				Command_Renamed.Respond(ERROR_NOT_OPS)
			End If
		Else
			Command_Renamed.Respond("Error: You must specify a username to ban.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnCAdd(ByRef Command_Renamed As clsCommandObj)
		Dim sArgs As String
		If (Command_Renamed.IsValid) Then
            If (Len(Command_Renamed.Argument("Message")) > 0) Then
                sArgs = "--banmsg " & Command_Renamed.Argument("Message")
            Else
                sArgs = "--banmsg Client Ban"
            End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Args = StringFormat("{0} +B --type GAME {1}", Command_Renamed.Argument("Game"), sArgs)
			Call OnAdd(Command_Renamed)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnCDel(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Command_Renamed.Args = Command_Renamed.Argument("Game") & " -B --type GAME"
			Call OnAdd(Command_Renamed)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnChPw(ByRef Command_Renamed As clsCommandObj)
		Dim delay As Short
		If (Command_Renamed.IsValid) Then
			Select Case (LCase(Command_Renamed.Argument("SubCommand")))
				Case "on", "true", "enable", "enabled"
                    If (Len(Command_Renamed.Argument("Value")) > 0) Then

                        BotVars.ChannelPassword = Command_Renamed.Argument("Value")
                        If (BotVars.ChannelPasswordDelay <= 0) Then BotVars.ChannelPasswordDelay = 30

                        Command_Renamed.Respond(StringFormat("Channel password protection enabled, delay set to {0}.", BotVars.ChannelPasswordDelay))
                    Else
                        Command_Renamed.Respond("Error: You must supply a password.")
                    End If
					
				Case "off", "false", "disable", "disabled"
					BotVars.ChannelPassword = vbNullString
					BotVars.ChannelPasswordDelay = 0
					Command_Renamed.Respond("Channel password protection disabled.")
					
				Case "delay"
					If (StrictIsNumeric(Command_Renamed.Argument("Value"))) Then
						delay = Val(Command_Renamed.Argument("Value"))
						If ((delay < 256) And (delay > 0)) Then
							BotVars.ChannelPasswordDelay = CByte(delay)
							Command_Renamed.Respond(StringFormat("Channel password delay set to {0}.", delay))
						Else
							Command_Renamed.Respond("Error: Invalid channel delay.")
						End If
					Else
						Command_Renamed.Respond("Error: Invalid channel delay.")
					End If
					
				Case Else
					
                    If ((Len(BotVars.ChannelPassword) = 0) Or (BotVars.ChannelPasswordDelay = 0)) Then
                        Command_Renamed.Respond("Channel password protection is currently disabled.")
                    Else
                        Command_Renamed.Respond(StringFormat("Channel password protection is currently enabled. Password [{0}], Delay [{1}].", BotVars.ChannelPassword, BotVars.ChannelPasswordDelay))
                    End If
			End Select
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnClearBanList(ByRef Command_Renamed As clsCommandObj)
		g_Channel.ClearBanlist()
		Command_Renamed.Respond("Banned user list cleared.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnD2LevelBan(ByRef Command_Renamed As clsCommandObj)
		Dim Level As Short
        If (Len(Command_Renamed.Argument("Level")) > 0) Then
            Level = CShort(Command_Renamed.Argument("Level"))
            If (Level > 0) Then
                If (Level < 256) Then
                    Command_Renamed.Respond(StringFormat("Banning Diablo II users under level {0}.", Level))
                    BotVars.BanD2UnderLevel = CByte(Level)
                Else
                    Command_Renamed.Respond("Error: Invalid level specified.")
                End If
            Else
                BotVars.BanD2UnderLevel = 0
                Command_Renamed.Respond("Diablo II level bans disabled.")
            End If

            If Config.LevelBanD2 <> BotVars.BanD2UnderLevel Then
                Config.LevelBanD2 = BotVars.BanD2UnderLevel
                Call Config.Save()
            End If
        Else
            If (BotVars.BanD2UnderLevel = 0) Then
                Command_Renamed.Respond("Currently not banning Diablo II users by level.")
            Else
                Command_Renamed.Respond(StringFormat("Currently banning Diablo II users under level {0}.", BotVars.BanD2UnderLevel))
            End If
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDelPhrase(ByRef Command_Renamed As clsCommandObj)
		Dim iFile As Short
		Dim sPhrase As String
		Dim bFound As Boolean
		Dim i As Short
		
		If (Command_Renamed.IsValid) Then
			sPhrase = Command_Renamed.Argument("Phrase")
			
			iFile = FreeFile
			
			FileOpen(iFile, GetFilePath(FILE_PHRASE_BANS), OpenMode.Output)
			For i = LBound(Phrases) To UBound(Phrases)
				If (Not StrComp(Phrases(i), sPhrase, CompareMethod.Text) = 0) Then
					PrintLine(iFile, Phrases(i))
				Else
					bFound = True
				End If
			Next i
			FileClose(iFile)
			
			ReDim Phrases(0)
			Call frmChat.LoadArray(LOAD_PHRASES, Phrases)
			
			If (bFound) Then
				Command_Renamed.Respond(StringFormat("Phrase {0}{1}{0} deleted.", Chr(34), sPhrase))
			Else
				Command_Renamed.Respond("Error: That phrase is not banned.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDes(ByRef Command_Renamed As clsCommandObj)
		If (g_Channel.Self.IsOperator) Then
            If (Len(Command_Renamed.Argument("Username")) > 0) Then
                Call frmChat.AddQ("/designate " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
                Command_Renamed.Respond(StringFormat("I have designated {0}.", Command_Renamed.Argument("Username")))
            Else
                If (Len(g_Channel.OperatorHeir) > 0) Then
                    Command_Renamed.Respond(StringFormat("I have designated {0}.", g_Channel.OperatorHeir))
                Else
                    Command_Renamed.Respond("No user has been designated.")
                End If
            End If
		Else
			Command_Renamed.Respond(ERROR_NOT_OPS)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnExile(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call OnShitAdd(Command_Renamed)
			Call OnIPBan(Command_Renamed)
		Else
			Command_Renamed.Respond("Error: You must specify a username.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnGiveUp(ByRef Command_Renamed As clsCommandObj)
		' This command will allow a user to designate a specified user using
		' Battle.net's "designate" command, and will then make the bot resign
		' its status as a channel moderator.  This command is useful if you are
		' lazy and you just wish to designate someone as quickly as possible.
		
		Dim i As Short
		Dim opsCount As Short
		Dim sUsername As String
		Dim colUsers As New Collection
		
		If (Command_Renamed.IsValid) Then
			sUsername = Command_Renamed.Argument("Username")
			If (g_Channel.GetUserIndex(sUsername) > 0) Then
				If (g_Channel.Self.IsOperator) Then
					opsCount = GetOpsCount
					
					If (StrComp(g_Channel.Name, "Clan " & Clan.Name, CompareMethod.Text) = 0) Then
						If (g_Clan.Self.Rank >= 4) Then
							'Lets get a count of Shamans that are in the channel
							For i = 1 To g_Clan.Shamans.Count()
								'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Shamans().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (g_Channel.GetUserIndexEx(g_Clan.Shamans.Item(i).Name) > 0) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Shamans().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									colUsers.Add(g_Clan.Shamans.Item(i).Name)
								End If
							Next i
							
							If (opsCount > (colUsers.Count() + 1)) Then 'colUser.Count is present shamans, +1 for the bot being ops
								Command_Renamed.Respond("Error: There is currently a channel moderator present that cannot be removed from his or her position.")
								Exit Sub
							End If
							
							'Lets demote the shamans
							For i = 1 To colUsers.Count()
								'UPGRADE_WARNING: Couldn't resolve default property of object colUsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Demote. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								g_Clan.Members.Item(g_Clan.GetMemberIndexEx(colUsers.Item(i))).Demote()
							Next i
							
							opsCount = GetOpsCount
						End If
					End If
					
					If (StrComp(Left(g_Channel.Name, 3), "Op ", CompareMethod.Text) = 0) Then
						If (opsCount >= 2) Then
							Command_Renamed.Respond("Error: There is currently a channel moderator present that cannot be removed from his or her position.")
							Exit Sub
						End If
					ElseIf (StrComp(Left(g_Channel.Name, 5), "Clan ", CompareMethod.Text) = 0) Then 
						If ((g_Clan.Self.Rank < 4) Or (Not StrComp(g_Channel.Name, "Clan " & Clan.Name, CompareMethod.Text) = 0)) Then
							If (opsCount >= 2) Then
								Command_Renamed.Respond("Error: There is currently a channel moderator present that cannot be removed from his or her position.")
								Exit Sub
							End If
						End If
					End If
					
					Call bnetSend("/designate " & ReverseConvertUsernameGateway(sUsername))
					Call Pause(2)
					Call bnetSend("/resign")
					
					For i = 1 To colUsers.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object colUsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Promote. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						g_Clan.Members.Item(g_Clan.GetUserIndexEx(colUsers.Item(i))).Promote()
					Next i
				Else
					Command_Renamed.Respond(ERROR_NOT_OPS)
				End If
			Else
				Command_Renamed.Respond("Error: The specified user is not present within the channel.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIdleBans(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Select Case (LCase(Command_Renamed.Argument("SubCommand")))
				Case "on", "true", "enable", "enabled"
					BotVars.IB_On = BTRUE
					
					If (StrictIsNumeric(Command_Renamed.Argument("Value"))) Then
						BotVars.IB_Wait = Val(Command_Renamed.Argument("Value"))
						Command_Renamed.Respond(StringFormat("IdleBans activated with a delay of {0}.", BotVars.IB_Wait))
					Else
						BotVars.IB_Wait = 400
						Command_Renamed.Respond("IdleBans activated using the default delay of 400.")
					End If
					
					Config.IdleBan = True
					Config.IdleBanDelay = BotVars.IB_Wait
					
				Case "off", "false", "disable", "disabled"
					BotVars.IB_On = BFALSE
					Config.IdleBan = False
					Command_Renamed.Respond("IdleBans deactivated.")
					
				Case "kick"
					Select Case LCase(Command_Renamed.Argument("Value"))
						Case "on", "true", "enable", "enabled"
							BotVars.IB_Kick = True
						Case "off", "false", "disable", "disabled"
							BotVars.IB_Kick = False
						Case "toggle"
							BotVars.IB_Kick = Not BotVars.IB_Kick
						Case Else
							Command_Renamed.Respond(StringFormat("IdleBan action is set to {0}.", IIf(Config.IdleBanKick, "kick", "ban")))
							Exit Sub
					End Select
					
					If Config.IdleBanKick <> BotVars.IB_Kick Then
						Config.IdleBanKick = BotVars.IB_Kick
						Call Config.Save()
						
						Command_Renamed.Respond(StringFormat("Idle users will now be {0}.", IIf(Config.IdleBanKick, "kicked", "banned")))
						Exit Sub
					End If
					
				Case "delay"
					If (StrictIsNumeric(Command_Renamed.Argument("Value"))) Then
						BotVars.IB_Wait = CShort(Command_Renamed.Argument("Value"))
						Config.IdleBanDelay = BotVars.IB_Wait
						Command_Renamed.Respond(StringFormat("IdleBan delay set to {0}.", BotVars.IB_Wait))
					Else
						Command_Renamed.Respond("Error: IdleBan delays require a numeric value.")
					End If
					
				Case Else
					If (BotVars.IB_On = BTRUE) Then
						Command_Renamed.Respond(StringFormat("Idle {0} is currently enabled with a delay of {1} seconds.", IIf(BotVars.IB_Kick, "kicking", "banning"), BotVars.IB_Wait))
					Else
						Command_Renamed.Respond("IdleBan is currently disabled.")
					End If
					Exit Sub
			End Select
			
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIPBan(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim dbTarget As udtGetAccessResponse
		Dim sTarget As String
		
		If (Command_Renamed.IsValid) Then
			
			If (Not g_Channel.Self.IsOperator) Then
				Command_Renamed.Respond("The bot does not currently have ops.")
				Exit Sub
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Username)
			If (Command_Renamed.IsLocal) Then
				dbAccess.Rank = 201
				dbAccess.Flags = "A"
			End If
			
			sTarget = StripInvalidNameChars(Command_Renamed.Argument("Username"))
			
            If (Len(sTarget) > 0) Then
                If (InStr(1, sTarget, "@") > 0) Then sTarget = StripRealm(sTarget)

                If (dbAccess.Rank < 101) Then
                    If (GetSafelist(sTarget) Or GetSafelist(Command_Renamed.Argument("Username"))) Then
                        Command_Renamed.Respond("Error: That user is safelisted.")
                        Exit Sub
                    End If
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object dbTarget. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dbTarget = GetCumulativeAccess(Command_Renamed.Argument("Username"))

                If ((dbTarget.Rank >= dbAccess.Rank) Or ((InStr(1, dbTarget.Flags, "A", CompareMethod.Text) > 0) And (dbAccess.Rank < 101))) Then
                    Command_Renamed.Respond("Error: You do not have enought access to do that.")
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call frmChat.AddQ(StringFormat("/ban {0} {1}", Command_Renamed.Argument("Username"), Command_Renamed.Argument("Message")), , (Command_Renamed.Username))
                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call frmChat.AddQ(StringFormat("/squelch {0}", Command_Renamed.Argument("Username")), , (Command_Renamed.Username))
                    Command_Renamed.Respond(StringFormat("User {0}{1}{0} IPBanned.", Chr(34), Command_Renamed.Argument("Username")))
                End If
            End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnKick(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
			If (g_Channel.Self.IsOperator) Then
				
				If (InStr(1, Command_Renamed.Argument("Username"), "*", CompareMethod.Text) > 0) Then
					Command_Renamed.Respond(WildCardBan(Command_Renamed.Argument("Username"), Command_Renamed.Argument("Message"), 0))
				Else
					If (Command_Renamed.IsLocal) Then
						frmChat.AddQ("/kick " & Command_Renamed.Args)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dbAccess = GetCumulativeAccess(Command_Renamed.Username)
						Command_Renamed.Respond(Ban(Command_Renamed.Args, dbAccess.Rank, 1))
					End If
				End If
			Else
				Command_Renamed.Respond(ERROR_NOT_OPS)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIPBans(ByRef Command_Renamed As clsCommandObj)
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				BotVars.IPBans = True
				Command_Renamed.Respond("IP banning activated.")
				
				g_Channel.CheckUsers()
				
			Case "off", "false", "disable", "disabled"
				BotVars.IPBans = False
				Command_Renamed.Respond("IP banning deactivated.")
				
			Case Else
				Command_Renamed.Respond(StringFormat("IP banning is currently {0}activated.", IIf(BotVars.IPBans, vbNullString, "de")))
		End Select
		
		If Config.IPBans <> BotVars.IPBans Then
			Config.IPBans = BotVars.IPBans
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnKickOnYell(ByRef Command_Renamed As clsCommandObj)
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				BotVars.KickOnYell = 1
				Command_Renamed.Respond("Kick-on-yell enabled.")
				
			Case "off", "false", "disable", "disabled"
				BotVars.KickOnYell = 0
				Command_Renamed.Respond("Kick-on-yell disabled.")
				
			Case Else
				Command_Renamed.Respond(StringFormat("Kick-on-yell is currently {0}.", IIf(BotVars.KickOnYell = 1, "enabled", "disabled")))
		End Select
		
		If Config.KickOnYell <> BotVars.KickOnYell Then
			Config.KickOnYell = BotVars.KickOnYell
			Call Config.Save()
		End If
	End Sub
	
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLevelBan(ByRef Command_Renamed As clsCommandObj)
		Dim Level As Short
        If (Len(Command_Renamed.Argument("Level")) > 0) Then
            Level = CShort(Command_Renamed.Argument("Level"))
            If (Level > 0) Then
                If (Level < 256) Then
                    Command_Renamed.Respond(StringFormat("Banning Warcraft III users under level {0}.", Level))
                    BotVars.BanUnderLevel = CByte(Level)
                Else
                    Command_Renamed.Respond("Error: Invalid level specified.")
                End If
            Else
                BotVars.BanUnderLevel = 0
                Command_Renamed.Respond("Levelbans disabled.")
            End If

            If Config.LevelBanW3 <> BotVars.BanUnderLevel Then
                Config.LevelBanW3 = BotVars.BanUnderLevel
                Call Config.Save()
            End If
        Else
            If (BotVars.BanUnderLevel = 0) Then
                Command_Renamed.Respond("Currently not banning Warcraft III users by level.")
            Else
                Command_Renamed.Respond(StringFormat("Currently banning Warcraft III users under level {0}.", BotVars.BanUnderLevel))
            End If
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPeonBan(ByRef Command_Renamed As clsCommandObj)
		' This command will enable, disable, or check the status of, WarCraft III peon
		' banning.  The "Peon" class is defined by Battle.net, and is currently the lowest
		' ranking WarCraft III user classification, in which, users have less than twenty-five
		' wins on record for any given race.
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				BotVars.BanPeons = True
				Command_Renamed.Respond("Peon banning activated.")
				
			Case "off", "false", "disable", "disabled"
				BotVars.BanPeons = False
				Command_Renamed.Respond("Peon banning deactivated.")
				
			Case Else
				Command_Renamed.Respond(StringFormat("The bot is currently {0}banning peons.", IIf(BotVars.BanPeons, vbNullString, "not ")))
		End Select
		
		If Config.PeonBan <> BotVars.BanPeons Then
			Config.PeonBan = BotVars.BanPeons
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPhraseBans(ByRef Command_Renamed As clsCommandObj)
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				PhraseBans = True
				Command_Renamed.Respond("Phrasebans have been enabled.")
				
			Case "off", "false", "disable", "disabled"
				PhraseBans = False
				Command_Renamed.Respond("Phrasebans have been disabled.")
				
			Case "kick"
				Select Case LCase(Command_Renamed.Argument("Value"))
					Case "on"
						Config.PhraseKick = True
					Case "off"
						Config.PhraseKick = False
					Case "toggle"
						Config.PhraseKick = Not Config.PhraseKick
					Case Else
						Command_Renamed.Respond(StringFormat("Phraseban punishment is set to {0}.", IIf(Config.PhraseKick, "kick", "ban")))
						Exit Sub
				End Select
				
				Command_Renamed.Respond(StringFormat("The punishment for saying a banned phrase is set to: {0}", IIf(Config.PhraseKick, "kick", "ban")))
				
				Call Config.Save()
				Exit Sub
			Case Else
				Command_Renamed.Respond(StringFormat("Phrasebans are currently {0}.", IIf(PhraseBans, "enabled", "disabled")))
		End Select
		
		If (Config.PhraseBans <> PhraseBans) Then
			Config.PhraseBans = PhraseBans
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPingBan(ByRef Command_Renamed As clsCommandObj)
		Dim sValue As String
		sValue = Command_Renamed.Argument("value")
		
		If Command_Renamed.IsValid Then
			Select Case LCase(sValue)
				Case "on", "true", "enable", "enabled"
					Config.PingBan = True
				Case "off", "false", "disable", "disabled"
					Config.PingBan = False
				Case "toggle"
					Config.PingBan = Not Config.PingBan
				Case Else
					If IsNumeric(sValue) Then
						Config.PingBanLevel = CInt(sValue)
						Call Config.Save()
						
						Command_Renamed.Respond("PingBan level set to: " & Config.PingBanLevel)
						If Config.PingBan Then Call g_Channel.CheckUsers()
						Exit Sub
					Else
						If Config.PingBan Then
							Command_Renamed.Respond(StringFormat("PingBan is enabled. Users with a ping {0} {1} will be banned.", IIf(Config.PingBanLevel <= 0, "equal to", "greater than"), Config.PingBanLevel))
						Else
							Command_Renamed.Respond("PingBan is currently disabled.")
						End If
						Exit Sub
					End If
			End Select
			Call Config.Save()
			
			Command_Renamed.Respond(StringFormat("PingBan has been {0}.", IIf(Config.PingBan, "enabled", "disabled")))
			
			If Config.PingBan Then Call g_Channel.CheckUsers()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPlugBan(ByRef Command_Renamed As clsCommandObj)
		' This command will enable, disable, or check the status of, UDP plug bans.
		' UDP plugs were traditionally used, in place of lag bars, to signifiy
		' that a user was incapable of hosting (or possibly even joining) a game.
		' However, as bot development became more popular, the emulation of such
		' a connectivity issue became fairly common, and the UDP plug began to
		' represent that a user was using a bot.  This feature allows for the
		' banning of both, potential bots, and users unlikely to be capable of
		' creating and/or joining games based on the UDP protocol.
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				If (BotVars.PlugBan) Then
					Command_Renamed.Respond("PlugBan is already activated.")
				Else
					BotVars.PlugBan = True
					Command_Renamed.Respond("PlugBan activated.")
					Call g_Channel.CheckUsers()
				End If
				
			Case "off", "false", "disable", "disabled"
				If (Not BotVars.PlugBan) Then
					Command_Renamed.Respond("PlugBan is already deactivated.")
				Else
					BotVars.PlugBan = False
					Command_Renamed.Respond("PlugBan deactivated.")
				End If
				
			Case Else
				Command_Renamed.Respond(StringFormat("The bot is currently {0}banning people with the UDP plug.", IIf(BotVars.PlugBan, vbNullString, "not ")))
		End Select
		
		If Config.UDPBan <> BotVars.PlugBan Then
			Config.UDPBan = BotVars.PlugBan
			Call Config.Save()
		End If
		
		If Config.UDPBan Then
			Call g_Channel.CheckUsers()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPOff(ByRef Command_Renamed As clsCommandObj)
		PhraseBans = False
		If Config.PhraseBans Then
			Config.PhraseBans = False
			Call Config.Save()
		End If
		Command_Renamed.Respond("Phrasebans deactivated.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPOn(ByRef Command_Renamed As clsCommandObj)
		PhraseBans = True
		If Not Config.PhraseBans Then
			Config.PhraseBans = True
			Call Config.Save()
		End If
		Command_Renamed.Respond("Phrasebans activated.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnProtect(ByRef Command_Renamed As clsCommandObj)
		Select Case (LCase(Command_Renamed.Argument("SubCommand")))
			Case "on", "true", "enable", "enabled"
				If (g_Channel.Self.IsOperator) Then
					Protect = True
					
					Call WildCardBan("*", ProtectMsg, 1)
					
                    If (Len(Command_Renamed.Username) > 0) Then
                        Command_Renamed.Respond(StringFormat("Lockdown activated by {0}.", Command_Renamed.Username))
                    Else
                        Command_Renamed.Respond("Lockdown activated.")
                    End If
				Else
					Command_Renamed.Respond(ERROR_NOT_OPS)
				End If
				
			Case "off", "false", "disable", "disabled"
				If (Protect) Then
					Protect = False
					
					Command_Renamed.Respond("Lockdown deactivated.")
				Else
					Command_Renamed.Respond("Lockdown was not enabled.")
				End If
				
			Case Else
				Command_Renamed.Respond(StringFormat("Lockdown is currently {0}active.", IIf(Protect, vbNullString, "not ")))
		End Select
		
		If Protect <> Config.ChannelProtection Then
			Config.ChannelProtection = Protect
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPStatus(ByRef Command_Renamed As clsCommandObj)
		Command_Renamed.Respond(StringFormat("Phrasebans are currently {0}.", IIf(PhraseBans, "enabled", "disabled")))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnQuietTime(ByRef Command_Renamed As clsCommandObj)
		' This command will enable, disable or check the status of, quiet time.
		' Quiet time is a feature that will ban non-safelisted users from the
		' channel when they speak publicly within the channel.  This is useful
		' when a channel wishes to have a discussion while allowing public
		' attendance, but disallowing public participation.
		Select Case LCase(Command_Renamed.Argument("SubCommand"))
			Case "on", "true", "enable", "enabled"
				BotVars.QuietTime = True
				Command_Renamed.Respond("Quiet-time enabled.")
				
			Case "off", "false", "disable", "disabled"
				BotVars.QuietTime = False
				Command_Renamed.Respond("Quiet-time disabled.")
				
			Case "kick"
				Select Case LCase(Command_Renamed.Argument("Value"))
					Case "on", "true", "enable", "enabled"
						Config.QuietTimeKick = True
					Case "off", "false", "disable", "disabled"
						Config.QuietTimeKick = False
					Case "toggle"
						Config.QuietTimeKick = Not Config.QuietTimeKick
					Case Else
						Command_Renamed.Respond(StringFormat("The QuietTime action is set to {0}", IIf(Config.QuietTimeKick, "kick", "ban")))
						Exit Sub
				End Select
				
				Command_Renamed.Respond(StringFormat("QuietTime action set to {0}.", IIf(Config.QuietTimeKick, "kick", "ban")))
				Call Config.Save()
			Case Else
				Command_Renamed.Respond(StringFormat("QuietTime is currently {0}.", IIf(BotVars.QuietTime, "enabled", "disabled")))
		End Select
		
		If Config.QuietTime <> BotVars.QuietTime Then
			Config.QuietTime = BotVars.QuietTime
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnResign(ByRef Command_Renamed As clsCommandObj)
		If (Not g_Channel.Self.IsOperator) Then Exit Sub
		Call frmChat.AddQ("/resign", (modQueueObj.PRIORITY.SPECIAL_MESSAGE), (Command_Renamed.Username))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSafeAdd(ByRef Command_Renamed As clsCommandObj)
		Dim sArgs As String
		Dim dbAccess As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
            If (Len(BotVars.DefaultSafelistGroup) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dbAccess = GetAccess(BotVars.DefaultSafelistGroup, "GROUP")

                If (Len(dbAccess.Username) > 0) Then sArgs = "--group " & BotVars.DefaultSafelistGroup
            End If
			
            If (Len(sArgs) = 0) Then sArgs = "+S"
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sArgs = StringFormat("{0} {1} --type USER", Command_Renamed.Argument("Username"), sArgs)
			
			Command_Renamed.Args = sArgs
			Call OnAdd(Command_Renamed)
		Else
			Command_Renamed.Respond("Error: You must specify a username.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSafeDel(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Command_Renamed.Args = Command_Renamed.Argument("Username") & " -S --type USER"
			Call OnAdd(Command_Renamed)
		Else
			Command_Renamed.Respond("Error: You must supply a username.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShitAdd(ByRef Command_Renamed As clsCommandObj)
		Dim sArgs As String
		
		Dim dbAccess As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
            If (Len(BotVars.DefaultShitlistGroup) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dbAccess = GetAccess(BotVars.DefaultShitlistGroup, "GROUP")
                If (Len(dbAccess.Username) > 0) Then
                    sArgs = "--group " & BotVars.DefaultShitlistGroup
                End If
            End If
			
            If (Len(sArgs) = 0) Then sArgs = "+B"
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sArgs = StringFormat("{0} {1} --type USER", Command_Renamed.Argument("Username"), sArgs)
			
            If (Len(Command_Renamed.Argument("Message")) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sArgs = StringFormat("{0} --banmsg {1}", sArgs, Command_Renamed.Argument("Message"))
            End If
			
			Command_Renamed.Args = sArgs
			Call OnAdd(Command_Renamed)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShitDel(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Command_Renamed.Args = Command_Renamed.Argument("Username") & " -B --type USER"
			Call OnAdd(Command_Renamed)
		Else
			Command_Renamed.Respond("Error: You must specify a username.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSweepBan(ByRef Command_Renamed As clsCommandObj)
		' This command will grab the listing of users in the specified channel
		' using Battle.net's "who" command, and will then begin banning each
		' user from the current channel using Battle.net's "ban" command.
		If (Command_Renamed.IsValid) Then
			If (g_Channel.Self.IsOperator) Then
				' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
				'Call CacheChannelList(vbNullString, 255, "ban ")
				Call CacheChannelList(CacheChanneListEnum.enReset, "ban ")
				Call frmChat.AddQ("/who " & Command_Renamed.Argument("Channel"), (modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE), (Command_Renamed.Username), "request_receipt")
			Else
				Command_Renamed.Respond(ERROR_NOT_OPS)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSweepIgnore(ByRef Command_Renamed As clsCommandObj)
		' This command will grab the listing of users in the specified channel
		' using Battle.net's "who" command, and will then begin ignoring each
		' user using Battle.net's "squelch" command.  This command is often used
		' instead of "sweepban" to temporarily ban all users on a given ip address
		' without actually immediately banning them from the channel.  This is
		' useful if a user wishes to stay below Battle.net's limitations on bans,
		' and still prevent a number of users from joining the channel for a
		' temporary amount of time.
		If (Command_Renamed.IsValid) Then
			' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to be removed to script
			'Call CacheChannelList(vbNullString, 255, "squelch ")
			Call CacheChannelList(CacheChanneListEnum.enReset, "squelch ")
			Call frmChat.AddQ("/who " & Command_Renamed.Argument("Channel"), (modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE), (Command_Renamed.Username), "request_receipt")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTagAdd(ByRef Command_Renamed As clsCommandObj)
		Dim sArgs As String
		Dim dbAccess As udtGetAccessResponse
		If (Command_Renamed.IsValid) Then
            If (Len(BotVars.DefaultTagbansGroup) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dbAccess = GetAccess(BotVars.DefaultTagbansGroup, "GROUP")

                If (Len(dbAccess.Username) > 0) Then
                    sArgs = "--group " & BotVars.DefaultTagbansGroup
                End If
            End If
			
            If (Len(sArgs) = 0) Then sArgs = "+B"
			
            If (Len(Command_Renamed.Argument("Message")) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sArgs = StringFormat("{0} --banmsg {1}", sArgs, Command_Renamed.Argument("Message"))
            End If
			
			If (InStr(Command_Renamed.Argument("Tag"), "*") = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sArgs = StringFormat("{0} {1} --type CLAN", Command_Renamed.Argument("Tag"), sArgs)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sArgs = StringFormat("{0} {1} --type USER", Command_Renamed.Argument("Tag"), sArgs)
			End If
			
			Command_Renamed.Args = sArgs
			Call OnAdd(Command_Renamed)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTagDel(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim ResponseTag() As String
		Dim ResponseClan() As String
		Dim i As Short
		If (Command_Renamed.IsValid) Then
			' this code is OLD and needs to be redone when OnAddOld() is redone.
			' but nobody's going to read this anyway -Ribose
			ReDim Preserve ResponseTag(0)
			ReDim Preserve ResponseClan(0)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Username)
			If (Command_Renamed.IsLocal) Then
				dbAccess.Rank = 201
				dbAccess.Flags = "A"
			End If
			
			If (InStr(Command_Renamed.Argument("Tag"), "*") <> 0) Then
				Command_Renamed.Args = Command_Renamed.Argument("Tag") & " -B --Type CLAN"
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Args = StringFormat("{0} -B --Type CLAN", Command_Renamed.Argument("Tag"))
			End If
			
			Call OnAddOld(Command_Renamed.Username, dbAccess, Command_Renamed.Args, Command_Renamed.IsLocal, ResponseClan)
			
			If UBound(ResponseClan) = 0 Then
				If (StrComp(Left(ResponseClan(0), 7), "Error: ") = 0) Then
					If (InStr(Command_Renamed.Argument("Tag"), "*") <> 0) Then
						Command_Renamed.Args = Command_Renamed.Argument("Tag") & " -B --Type USER"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Command_Renamed.Args = StringFormat("*{0}* -B --Type USER", Command_Renamed.Argument("Tag"))
					End If
					
					Call OnAddOld(Command_Renamed.Username, dbAccess, Command_Renamed.Args, Command_Renamed.IsLocal, ResponseTag)
					
					For i = LBound(ResponseTag) To UBound(ResponseTag)
						Command_Renamed.Respond(ResponseTag(i))
					Next i
				Else
					Command_Renamed.Respond(ResponseClan(i))
				End If
			Else
				' something went wrong? print all
				For i = LBound(ResponseClan) To UBound(ResponseClan)
					Command_Renamed.Respond(ResponseClan(i))
				Next i
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnBan(ByRef Command_Renamed As clsCommandObj)
		Dim sTargetUser As String
		
		If (Command_Renamed.IsValid) Then
			If (g_Channel.Self.IsOperator) Then
				sTargetUser = Command_Renamed.Argument("Username")
				
				' If no user was specified, unban the last banned user.
                If (Len(sTargetUser) = 0) Then
                    If (g_Channel.Banlist.Count() > 0) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        sTargetUser = g_Channel.Banlist.Item(g_Channel.Banlist.Count()).Name
                    End If
                End If
				
				If (InStr(1, sTargetUser, "*", CompareMethod.Binary) <> 0) Then
					Call WildCardBan(sTargetUser, vbNullString, 2)
				Else
					Call frmChat.AddQ("/unban " & sTargetUser, (modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE), (Command_Renamed.Username))
				End If
			Else
				Command_Renamed.Respond(ERROR_NOT_OPS)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnExile(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Call OnShitDel(Command_Renamed)
			Call OnUnIPBan(Command_Renamed)
		Else
			Command_Renamed.Respond("Error: You must specify a user to unban.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUnIPBan(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			If (g_Channel.Self.IsOperator) Then
				Call frmChat.AddQ("/unsquelch " & Command_Renamed.Argument("Username"),  , (Command_Renamed.Username))
				Call frmChat.AddQ("/unban " & Command_Renamed.Argument("Username"),  , (Command_Renamed.Username))
				Command_Renamed.Respond(StringFormat("User {0}{1}{0} has been Un-IPBanned.", Chr(34), Command_Renamed.Argument("Username")))
			Else
				Command_Renamed.Respond(ERROR_NOT_OPS)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnVoteBan(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			If (VoteDuration = -1) Then
				If (g_Channel.Self.IsOperator) Then
					Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, Command_Renamed.Argument("Username"))
					VoteDuration = 30
					If (Command_Renamed.IsLocal) Then
						With VoteInitiator
							.Rank = 201
							.Flags = "A"
							.Username = "(Console)"
						End With
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object VoteInitiator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						VoteInitiator = GetCumulativeAccess(Command_Renamed.Username)
					End If
					
					Command_Renamed.Respond(StringFormat("30-second VoteBan vote started. Type YES to ban {0}, NO to acquit him/her.", Command_Renamed.Argument("Username")))
				Else
					Command_Renamed.Respond(ERROR_NOT_OPS)
				End If
			Else
				Command_Renamed.Respond("A vote is currently in progress.")
			End If
		Else
			Command_Renamed.Respond("You must specify a user to kick.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnVoteKick(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			If (VoteDuration = -1) Then
				If (g_Channel.Self.IsOperator) Then
					Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, Command_Renamed.Argument("Username"))
					VoteDuration = 30
					If (Command_Renamed.IsLocal) Then
						With VoteInitiator
							.Rank = 201
							.Flags = "A"
							.Username = "(Console)"
						End With
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object VoteInitiator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						VoteInitiator = GetCumulativeAccess(Command_Renamed.Username)
					End If
					
					Command_Renamed.Respond(StringFormat("30-second VoteKick vote started. Type YES to kick {0}, NO to acquit him/her.", Command_Renamed.Argument("Username")))
				Else
					Command_Renamed.Respond(ERROR_NOT_OPS)
				End If
			Else
				Command_Renamed.Respond("A vote is currently in progress.")
			End If
		Else
			Command_Renamed.Respond("You must specify a user to kick.")
		End If
	End Sub
	
	Private Function GetOpsCount(Optional ByRef strIgnore As String = vbNullString) As Short
		Dim i As Short
		For i = 1 To g_Channel.Users.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Not StrComp(g_Channel.Users.Item(i).DisplayName, strIgnore, CompareMethod.Binary) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(i).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (g_Channel.Users.Item(i).IsOperator) Then GetOpsCount = GetOpsCount + 1
			End If
		Next i
	End Function
	
	
	'This is called in ChannelBan/Squelch to reset, and cache the users in a channel
	'Then again in frmChat.tmrCache, to get a list of users and ban/squelch them
	'The CS/CB commands AddQ "/Who Channel" and have a "request_receipt" message
	'Then in Event_ServerInfo() it checks if it has a "request_receipt" if it does it call this function
	'This is vary fucking ugly, we NEED to move this to a script whenever possible.
	'Public Function CacheChannelList(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
	Public Function CacheChannelList(ByVal eMode As CacheChanneListEnum, ByRef Data As String) As String
		Static colData As New Collection
		Static sToken As String
		Static bChannelListFollows As Boolean 'renamed this variable for clarity
		
		Dim sTemp As String
		
		'If(Mode = 255) Then
		If (eMode = CacheChanneListEnum.enReset) Then
			Do While colData.Count() > 0
				colData.Remove(1)
			Loop 
			sToken = Data
			bChannelListFollows = False
		End If
		
		If (Not InStr(1, LCase(Data), "users in channel ", CompareMethod.Text) = 0) Then
			' we weren't expecting a channel list, but we are now
			bChannelListFollows = True
		Else
			If (bChannelListFollows = True) Then ' if we're expecting a channel list, process it
				Select Case (eMode)
					Case CacheChanneListEnum.enRetrieve ' RETRIEVE
						' Merge all the cache array items into one comma-delimited string
						Do While colData.Count() > 0
							'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sTemp = StringFormat("{0}{1}, ", sTemp, colData.Item(1))
							colData.Remove(1)
						Loop 
						Data = sToken
						CacheChannelList = sTemp
					Case CacheChanneListEnum.enAdd ' ADD
						colData.Add(Data)
				End Select
			End If
		End If
	End Function
	
	'This must be public becase The old OnAdd command uses it -.-
	Public Function WildCardBan(ByVal sMatch As String, ByVal sBanMsg As String, ByVal Banning As Byte) As String
		'Values for Banning byte:
		'0 = Kick
		'1 = Ban
		'2 = Unban
		
		Dim i As Short
		Dim iSafe As Short
		Dim sCommand As String
		Dim sName As String
		
		If (g_Channel.Self.IsOperator) Then
			If (g_Channel.Users.Count() < 1) Then Exit Function
			
            If (Len(sBanMsg) = 0) Then sBanMsg = sMatch
			
			sMatch = PrepareCheck(sMatch)
			
			Select Case (Banning)
				Case 1 : sCommand = "/ban"
				Case 2 : sCommand = "/unban"
				Case Else : sCommand = "/kick"
			End Select
			
			
			If (Not Banning = 2) Then
				' Kicking or Banning
				For i = 1 To g_Channel.Users.Count()
					With g_Channel.Users.Item(i)
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (Not StrComp(.DisplayName, GetCurrentUsername, CompareMethod.Binary) = 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sName = PrepareCheck(.DisplayName)
							
							If (sName Like sMatch) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (GetSafelist(.DisplayName) = False) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(i).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (Not .IsOperator) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										Call frmChat.AddQ(StringFormat("{0} {1} {2}", sCommand, .DisplayName, sBanMsg))
									End If
								Else
									iSafe = (iSafe + 1)
								End If
							End If
						End If
					End With
				Next i
				
				If (iSafe > 0) Then
					If (Not StrComp(sBanMsg, ProtectMsg, CompareMethod.Text) = 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						WildCardBan = StringFormat("Encountered {0} safelisted user{1}.", iSafe, IIf(iSafe > 1, "s", vbNullString))
					End If
				End If
				
			Else
				For i = 1 To g_Channel.Banlist.Count()
					With g_Channel.Banlist.Item(i)
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist(i).IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If ((.IsActive) And (Len(.DisplayName) > 0)) Then
                            If (sMatch = "*") Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Call frmChat.AddQ(StringFormat("{0} {1}", sCommand, .DisplayName))
                            Else
                                'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                sName = PrepareCheck(.DisplayName)
                                If (sName Like sMatch) Then
                                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    Call frmChat.AddQ(StringFormat("{0} {1}", sCommand, .DisplayName))
                                End If
                            End If
                        End If
					End With
				Next i
			End If
		End If
	End Function
End Module