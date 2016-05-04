Option Strict Off
Option Explicit On
Module modCommandsInfo
	'This module will hold all of the 'Info' Commands
	'Commands that return information, but have really no functionality
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAbout(ByRef Command_Renamed As clsCommandObj)
		Command_Renamed.Respond(".: " & CVERSION & " :.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAccountInfo(ByRef Command_Renamed As clsCommandObj)
		If ((Not Command_Renamed.IsLocal) Or Command_Renamed.PublicOutput) Then
			PPL = True
			If ((BotVars.WhisperCmds Or Command_Renamed.WasWhispered) And (Not Command_Renamed.IsLocal)) Then
				PPLRespondTo = Command_Renamed.Username
			End If
		End If
		RequestSystemKeys()
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBanCount(ByRef Command_Renamed As clsCommandObj)
		If (g_Channel.BanCount = 0) Then
			Command_Renamed.Respond("No users have been banned since I joined this channel.")
		Else
			Command_Renamed.Respond(StringFormat("Since I joined this channel, {0} user{1} have been banned.", g_Channel.BanCount, IIf(g_Channel.BanCount > 1, "s", vbNullString)))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBanListCount(ByRef Command_Renamed As clsCommandObj)
		If (g_Channel.Banlist.Count() = 0) Then
			Command_Renamed.Respond("There are no users on the internal ban list.")
		Else
			Command_Renamed.Respond(StringFormat("There {0} currently {1} user{2} on the internal ban list.", IIf(g_Channel.Banlist.Count() > 1, "are", "is"), g_Channel.Banlist.Count(), IIf(g_Channel.Banlist.Count() > 1, "s", vbNullString)))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBanned(ByRef Command_Renamed As clsCommandObj)
		Dim sResult As String
		Dim i As Short
		Dim j As Short
		Dim BanCount As Short
		
		If (g_Channel.Banlist.Count() = 0) Then
			Command_Renamed.Respond("There are presently no users on the bot's internal banlist.")
		Else
			sResult = "User(s) banned: "
			For i = 1 To g_Channel.Banlist.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist(i).IsDuplicateBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Not g_Channel.Banlist.Item(i).IsDuplicateBan) Then
					For j = 1 To g_Channel.Banlist.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist(i).DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (StrComp(g_Channel.Banlist.Item(j).DisplayName, g_Channel.Banlist.Item(i).DisplayName, CompareMethod.Text) = 0) Then
							BanCount = (BanCount + 1)
						End If
					Next j
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Banlist().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sResult = StringFormat("{0}, {1}", sResult, g_Channel.Banlist.Item(i).DisplayName)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (BanCount > 1) Then sResult = StringFormat("{0} ({1})", sResult, BanCount)
					
					If ((Len(sResult) > 90) And (Not i = g_Channel.Banlist.Count())) Then
						Command_Renamed.Respond(Replace(sResult, " , ", Space(1)))
						sResult = "Users(s) banned: "
					End If
				End If
				BanCount = 0
			Next i
            If (Len(sResult) > Len("Users(s) banned: , ")) Then 'We don't want to send an empty line
                Command_Renamed.Respond(Replace(sResult, " , ", Space(1)))
            End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnClientBans(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim strResponse As Object
		
		If (Command_Renamed.IsValid) Then
			Call SearchDatabase(bufResponse,  ,  ,  , "GAME",  ,  , "B")
			
			For	Each strResponse In bufResponse
				'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Respond(CStr(strResponse))
			Next strResponse
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDetail(ByRef Command_Renamed As clsCommandObj)
		Dim sRetAdd, sRetMod As String
		Dim i As Short
		If (Command_Renamed.IsValid) Then
			
			For i = 0 To UBound(DB)
				With DB(i)
					If (StrComp(Command_Renamed.Argument("Username"), .Username, CompareMethod.Text) = 0) Then
                        If ((Not .AddedBy = "%") And (Len(.AddedBy) > 0)) Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            sRetAdd = StringFormat("{0} was added by {1} on {2}.", .Username, .AddedBy, .AddedOn)
                        End If
						
                        If ((Not .ModifiedBy = "%") And (Len(.ModifiedBy) > 0)) Then
                            If ((Not .AddedOn = .ModifiedOn) Or (Not StrComp(.AddedBy, .ModifiedBy, CompareMethod.Text) = 0)) Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                sRetMod = StringFormat(" The entry was last modified by {0} on {1}.", .ModifiedBy, .ModifiedOn)
                            Else
                                sRetMod = " The entry has not been modified since it was added."
                            End If
                        End If
						
                        If ((Len(sRetAdd) > 0) Or (Len(sRetMod) > 0)) Then
                            If (Len(sRetAdd) > 0) Then
                                Command_Renamed.Respond(sRetAdd & sRetMod)
                            Else
                                Command_Renamed.Respond(Trim(sRetMod))
                            End If
                        Else
                            Command_Renamed.Respond("No detailed information is available for that user.")
                        End If
						
						Exit Sub
					End If
				End With
			Next i
			
			Command_Renamed.Respond("That user was not found in the database.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFind(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim bufResponse() As String
		Dim strResponse As Object
		
		If (Not Command_Renamed.IsValid) Then Exit Sub
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(GetFilePath(FILE_USERDB))) = 0) Then
            Command_Renamed.Respond("No userlist available. Place a users.txt file in the bot's root directory.")
            Exit Sub
        End If
		
		ReDim Preserve bufResponse(0)
		
		Dim LowerRank As Short
		Dim UpperRank As Short
		If (StrictIsNumeric(Command_Renamed.Argument("Username/Rank"))) Then
            If (Len(Command_Renamed.Argument("UpperRank")) = 0) Then
                Call SearchDatabase(bufResponse, , , , , Val(Command_Renamed.Argument("Username/Rank")))
            Else

                LowerRank = Val(Command_Renamed.Argument("Username/Rank"))
                UpperRank = CShort(Command_Renamed.Argument("UpperRank"))

                If (UpperRank = LowerRank) Then
                    Call SearchDatabase(bufResponse, , , , , LowerRank)
                ElseIf (UpperRank > LowerRank) Then
                    Call SearchDatabase(bufResponse, , , , , LowerRank, UpperRank)
                Else
                    Call SearchDatabase(bufResponse, , , , , UpperRank, LowerRank)
                End If
            End If
		Else
			Call SearchDatabase(bufResponse,  , PrepareCheck(Command_Renamed.Argument("Username/Rank")))
		End If
		
		For	Each strResponse In bufResponse
			'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(CStr(strResponse))
		Next strResponse
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFindAttr(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim strResponse As Object
		
		If (Command_Renamed.IsValid) Then
			Call SearchDatabase(bufResponse,  ,  ,  ,  ,  ,  , Command_Renamed.Argument("Attributes"))
			For	Each strResponse In bufResponse
				'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Respond(CStr(strResponse))
			Next strResponse
		Else
			Command_Renamed.Respond("You must specify flag(s) to search for.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFindGrp(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim strResponse As Object
		
		If (Command_Renamed.IsValid) Then
			Call SearchDatabase(bufResponse,  ,  , Command_Renamed.Argument("Group"))
			For	Each strResponse In bufResponse
				'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Respond(CStr(strResponse))
			Next strResponse
		Else
			Command_Renamed.Respond("You must specify a group to find.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnHelp(ByRef Command_Renamed As clsCommandObj)
		Dim strCommand As String
		Dim strScript As String
		Dim docs As clsCommandDocObj
		
		strCommand = IIf(Command_Renamed.IsValid, Command_Renamed.Argument("Command"), "help")
        strScript = IIf(Len(Command_Renamed.Argument("ScriptOwner")) > 0, Command_Renamed.Argument("ScriptOwner"), Chr(0))
		
		docs = OpenCommand(strCommand, strScript)
        If (Len(docs.Name) = 0) Then
            Command_Renamed.Respond("Sorry, but no related documentation could be found.")
        Else
            If (docs.aliases.Count() > 1) Then
                Command_Renamed.Respond(StringFormat("[{0} (Aliases: {4})]: {1} (Syntax: {2}). {3}", docs.Name, docs.description, docs.SyntaxString((Command_Renamed.IsLocal)), docs.RequirementsStringShort, docs.AliasString))
            ElseIf (docs.aliases.Count() = 1) Then
                Command_Renamed.Respond(StringFormat("[{0} (Alias: {4})]: {1} (Syntax: {2}). {3}", docs.Name, docs.description, docs.SyntaxString((Command_Renamed.IsLocal)), docs.RequirementsStringShort, docs.AliasString))
            Else
                Command_Renamed.Respond(StringFormat("[{0}]: {1} (Syntax: {2}). {3}", docs.Name, docs.description, docs.SyntaxString((Command_Renamed.IsLocal)), docs.RequirementsStringShort))
            End If
        End If
		'UPGRADE_NOTE: Object docs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		docs = Nothing
		
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnHelpAttr(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		
		Dim tmpbuf As String
		
		If (Command_Renamed.IsValid) Then
			tmpbuf = GetAllCommandsFor((Command_Renamed.docs),  , Command_Renamed.Argument("Flags"))
            If (Len(tmpbuf) > 0) Then
                Command_Renamed.Respond("Commands available to specified flag(s): " & tmpbuf)
            Else
                Command_Renamed.Respond("No commands are available to the given flag(s).")
            End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnHelpAttr().")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnHelpRank(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		
		Dim tmpbuf As String
		
		If (Command_Renamed.IsValid) Then
			If (CShort(Command_Renamed.Argument("Rank")) > -1) Then
				tmpbuf = GetAllCommandsFor((Command_Renamed.docs), CShort(Command_Renamed.Argument("Rank")))
                If (Len(tmpbuf) > 0) Then
                    Command_Renamed.Respond("Commands available to specified rank: " & tmpbuf)
                Else
                    Command_Renamed.Respond("No commands are available to the given rank.")
                End If
			Else
				Command_Renamed.Respond("The specified rank must be greater or equal to zero.")
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnHelpRank().")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnInfo(ByRef Command_Renamed As clsCommandObj)
		Dim UserIndex As Short
		
		If (Command_Renamed.IsValid) Then
			UserIndex = g_Channel.GetUserIndex(Command_Renamed.Argument("Username"))
			
			If (UserIndex > 0) Then
				With g_Channel.Users.Item(UserIndex)
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(UserIndex).Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(UserIndex).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Command_Renamed.Respond(StringFormat("User {0} is logged on using {1} with {2}a ping time of {3}ms.", .DisplayName, ProductCodeToFullName(.Game), IIf(.IsOperator, "ops, and ", vbNullString), .Ping))
					
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().TimeInChannel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Command_Renamed.Respond(StringFormat("He/she has been present in the channel for {0}.", ConvertTime(.TimeInChannel(), 1)))
				End With
			Else
				Command_Renamed.Respond("No such user is present.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnInitPerf(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		
		Dim ModName As String
		Dim Name As String
		Dim i As Short
		Dim strRet As String
		Dim Script As MSScriptControl.Module
		
		If modScripting.GetScriptSystemDisabled() Then
			Command_Renamed.Respond("Error: Scripts are globally disabled via the override.")
			Exit Sub
		End If
		
        If (Len(Command_Renamed.Argument("Script")) > 0) Then
            Script = modScripting.GetModuleByName(Command_Renamed.Argument("Script"))
            If (Script Is Nothing) Then
                Command_Renamed.Respond("Could not find the script specified.")
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object Script.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", CompareMethod.Text) = 1) Then
                    Command_Renamed.Respond(StringFormat("The Script {0} loaded in {1}ms.", GetScriptName((Script.Name)), GetScriptDictionary(Script)("InitPerf")))
                Else
                    Command_Renamed.Respond("That script is currently disabled.")
                End If
            End If
        Else
            If (frmChat.SControl.Modules.Count > 1) Then
                If (Command_Renamed.IsLocal And Not Command_Renamed.PublicOutput) Then
                    Command_Renamed.Respond("Script initialization performance:")
                Else
                    strRet = "Script initialization performance:"
                End If
                For i = 2 To frmChat.SControl.Modules.Count
                    Script = frmChat.SControl.Modules(i)

                    'UPGRADE_WARNING: Couldn't resolve default property of object Script.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", CompareMethod.Text) = 0) Then
                        If (Command_Renamed.IsLocal And Not Command_Renamed.PublicOutput) Then
                            Command_Renamed.Respond(StringFormat(" '{0}' loaded in {1}ms.", GetScriptName((Script.Name)), GetScriptDictionary(Script)("InitPerf")))
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strRet = StringFormat("{0} '{1}' {2}ms{3}", strRet, GetScriptName((Script.Name)), GetScriptDictionary(Script)("InitPerf"), IIf(i = frmChat.SControl.Modules.Count, vbNullString, ","))
                        End If
                    End If
                Next i

                If (Len(strRet)) Then Command_Renamed.Respond(strRet)
            Else
                Command_Renamed.Respond("There are no scripts currently loaded.")
            End If
        End If
		
		Exit Sub
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnInitPerf().")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLastSeen(ByRef Command_Renamed As clsCommandObj)
		Dim retVal As String
		Dim i As Short
		
		If (colLastSeen.Count() = 0) Then
			retVal = "I have not seen anyone yet."
		Else
			retVal = "Last 15 users seen: "
			For i = 1 To colLastSeen.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retVal = StringFormat("{0}{1}{2}", retVal, colLastSeen.Item(i), IIf(i = colLastSeen.Count(), vbNullString, ", "))
				If (i = 15) Then Exit For
			Next i
		End If
		Command_Renamed.Respond(retVal)
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLastWhisper(ByRef Command_Renamed As clsCommandObj)
        If (Len(LastWhisper) > 0) Then
            Command_Renamed.Respond(StringFormat("The last whisper to this bot was from {0} at {1} on {2}.", LastWhisper, FormatDateTime(LastWhisperFromTime, DateFormat.LongTime), FormatDateTime(LastWhisperFromTime, DateFormat.LongDate)))
        Else
            Command_Renamed.Respond("The bot has not been whispered since it logged on.")
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLocalIp(ByRef Command_Renamed As clsCommandObj)
		Command_Renamed.Respond(StringFormat("{0} local IPv4 IP address is: {1}", IIf(Command_Renamed.IsLocal, "Your", "My"), frmChat.sckBNet.LocalIP))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnOwner(ByRef Command_Renamed As clsCommandObj)
        If (Len(BotVars.BotOwner)) Then
            Command_Renamed.Respond("This bot's owner is " & BotVars.BotOwner & ".")
        Else
            Command_Renamed.Respond("There is no owner currently set.")
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPhrases(ByRef Command_Renamed As clsCommandObj)
		Dim sSubCommand As String
		Dim iCount As Short
		
		If (Command_Renamed.IsValid) Then
			sSubCommand = Command_Renamed.Argument("subcommand")
			
            If (Len(sSubCommand) = 0) Or (LCase(sSubCommand) = "list") Then
                If ListRespond(Command_Renamed, Phrases, "Banned phrases{%p}: ", IIf(Len(sSubCommand) = 0, 2, -1)) Then
                    Exit Sub
                End If
            End If
			
			iCount = GetListCount(Phrases)
			If iCount = 0 Then
				Command_Renamed.Respond("There are no banned phrases.")
			Else
				Command_Renamed.Respond(StringFormat("There are {0} banned phrases. Use the '{1} list' command to show them.", iCount, Command_Renamed.Name))
			End If
		Else
			Command_Renamed.Respond(StringFormat("Invalid command. Correct usage: {0} [list/count]", Command_Renamed.Name))
		End If
	End Sub
	
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPing(ByRef Command_Renamed As clsCommandObj)
		Dim Latency As Integer
		If (Command_Renamed.IsValid) Then
			Latency = GetPing(Command_Renamed.Argument("Username"))
			If (Latency >= -1) Then
				Command_Renamed.Respond(StringFormat("{0}'s ping at login was {1}ms.", Command_Renamed.Argument("Username"), Latency))
			Else
				Command_Renamed.Respond(StringFormat("I can not see {0} in the channel.", Command_Renamed.Argument("Username")))
			End If
		Else
			Command_Renamed.Respond("Please specify a user to ping.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPingMe(ByRef Command_Renamed As clsCommandObj)
		Dim Latency As Integer
		If (Command_Renamed.IsLocal) Then
			If (g_Online) Then
				Command_Renamed.Respond(StringFormat("Your ping at login was {0}ms.", GetPing(GetCurrentUsername)))
			Else
				Command_Renamed.Respond("Error: You are not logged on.")
			End If
		Else
			Latency = GetPing(Command_Renamed.Username)
			If (Latency >= -1) Then
				Command_Renamed.Respond(StringFormat("Your ping at login was {0}ms.", Latency))
			Else
				Command_Renamed.Respond("I can not see you in the channel.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnProfile(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			If ((Not Command_Renamed.IsLocal) Or (Command_Renamed.PublicOutput)) Then
				PPL = True
				If ((BotVars.WhisperCmds Or Command_Renamed.WasWhispered) And (Not Command_Renamed.IsLocal)) Then
					PPLRespondTo = Command_Renamed.Username
				Else
					PPLRespondTo = vbNullString
				End If
				Call RequestProfile(Command_Renamed.Argument("Username"))
			Else
				frmProfile.PrepareForProfile(Command_Renamed.Argument("Username"), False)
				Call RequestProfile(Command_Renamed.Argument("Username"))
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSafeCheck(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			If (GetSafelist(Command_Renamed.Argument("Username"))) Then
				Command_Renamed.Respond(StringFormat("{0} is on the bot's safelist.", Command_Renamed.Argument("Username")))
			Else
				Command_Renamed.Respond(StringFormat("{0} is not on the bot's safelist.", Command_Renamed.Argument("Username")))
			End If
		End If
	End Sub
	
	' handle safelist command
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSafeList(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim i As Integer
		
		Call SearchDatabase(bufResponse,  ,  ,  ,  ,  ,  , "S")
		
		For i = 0 To UBound(bufResponse)
			Command_Renamed.Respond(bufResponse(i))
		Next i
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnScriptDetail(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		
		Dim Script As MSScriptControl.Module
		
		If modScripting.GetScriptSystemDisabled() Then
			Command_Renamed.Respond("Error: Scripts are globally disabled via the override.")
			Exit Sub
		End If
		
		Dim ScriptInfo As Scripting.Dictionary
		Dim Version As String
		Dim VerTotal As Double
		Dim Author As String
		Dim Description As String
		If (Command_Renamed.IsValid) Then
			Script = modScripting.GetModuleByName(Command_Renamed.Argument("Script"))
			If (Script Is Nothing) Then
				Command_Renamed.Respond("Error: Could not find specified script.")
			Else
				
				ScriptInfo = GetScriptDictionary(Script)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object ScriptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Version = StringFormat("{0}.{1}{2}", Val(ScriptInfo("Major")), Val(ScriptInfo("Minor")), IIf(Val(ScriptInfo("Revision")) > 0, " Revision " & Val(ScriptInfo("Revision")), vbNullString))
				
				'UPGRADE_WARNING: Couldn't resolve default property of object ScriptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VerTotal = Val(ScriptInfo("Major")) + Val(ScriptInfo("Minor")) + Val(ScriptInfo("Revision"))
				
				'UPGRADE_WARNING: Couldn't resolve default property of object ScriptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Author = ScriptInfo("Author")
				'UPGRADE_WARNING: Couldn't resolve default property of object ScriptInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Description = ScriptInfo("Description")
				
                If ((Len(Author) = 0) And (VerTotal = 0) And (Len(Description) = 0)) Then
                    Command_Renamed.Respond(StringFormat("There is no additional information for the '{0}' script.", GetScriptName((Script.Name))))
                Else
                    Command_Renamed.Respond(StringFormat("{0}{1}{2}{3}", GetScriptName((Script.Name)), IIf(VerTotal > 0, " v" & Version, vbNullString), IIf(Len(Author) > 0, " by " & Author, vbNullString), IIf(Len(Description) > 0, ": " & Description, ".")))
                End If
			End If
		End If
		Exit Sub
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnScriptDetail().")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnScripts(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		
		Dim retVal As String
		Dim i As Short
		Dim Enabled As Boolean
		Dim Name As String
		Dim Count As Short
		
		If modScripting.GetScriptSystemDisabled() Then
			Command_Renamed.Respond("Error: Scripts are globally disabled via the override.")
			Exit Sub
		End If
		
		If (frmChat.SControl.Modules.Count > 1) Then
			For i = 2 To frmChat.SControl.Modules.Count
				Name = modScripting.GetScriptName(CStr(i))
				'UPGRADE_WARNING: Couldn't resolve default property of object GetModuleByName().CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Enabled = Not (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), "False", CompareMethod.Text) = 0)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retVal = StringFormat("{0}{1}{2}{3}{4}", retVal, IIf(Enabled, vbNullString, "("), Name, IIf(Enabled, vbNullString, ")"), IIf(i = frmChat.SControl.Modules.Count, vbNullString, ", "))
				
				Count = (Count + 1)
			Next i
			
			Command_Renamed.Respond(StringFormat("Loaded Scripts ({0}): {1}", Count, retVal))
		Else
			Command_Renamed.Respond("There are no scripts currently loaded.")
		End If
		
		Exit Sub
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnScripts().")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnServer(ByRef Command_Renamed As clsCommandObj)
		Dim RemoteHost As String
		Dim RemoteHostIP As String
		
		RemoteHost = frmChat.sckBNet.RemoteHost
		RemoteHostIP = frmChat.sckBNet.RemoteHostIP
		
		If (StrComp(RemoteHost, RemoteHostIP, CompareMethod.Binary) = 0) Then
			Command_Renamed.Respond("I am currently connected to " & RemoteHostIP & ".")
		Else
			Command_Renamed.Respond("I am currently connected to " & RemoteHost & " (" & RemoteHostIP & ").")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShitCheck(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		Dim compare As CompareMethod
		If (Command_Renamed.IsValid) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Argument("Username"))
			compare = IIf(BotVars.CaseSensitiveFlags, CompareMethod.Binary, CompareMethod.Text)
			If (Not InStr(1, dbAccess.Flags, "B", compare) = 0) Then
				If (Not InStr(1, dbAccess.Flags, "S", compare) = 0) Then
					Command_Renamed.Respond(Command_Renamed.Argument("Username") & "{0} is on the bot's shitlist; also on the safelist and will not be banned.")
				Else
					Command_Renamed.Respond(Command_Renamed.Argument("Username") & " is on the bot's shitlist.")
				End If
			Else
				Command_Renamed.Respond(Command_Renamed.Argument("Username") & " is not on the bot's shitlist.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShitList(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim i As Short
		
		Call SearchDatabase(bufResponse,  , "!*[*]*",  ,  ,  ,  , "B")
		
		For i = 0 To UBound(bufResponse)
			Command_Renamed.Respond(bufResponse(i))
		Next i
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTagBans(ByRef Command_Renamed As clsCommandObj)
		Dim bufResponse() As String
		Dim i As Short
		
		Call SearchDatabase(bufResponse,  , "*[*]*",  ,  ,  ,  , "B")
		
		For i = 0 To UBound(bufResponse)
			Command_Renamed.Respond(bufResponse(i))
		Next i
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTime(ByRef Command_Renamed As clsCommandObj)
		Command_Renamed.Respond(StringFormat("The current time on this computer is {0} on {1} ({2}).", TimeOfDay, VB6.Format(Today, "MM-dd-yyyy"), GetTimeZoneName()))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTrigger(ByRef Command_Renamed As clsCommandObj)
        If (Len(BotVars.TriggerLong) = 1) Then
            Command_Renamed.Respond(StringFormat("The bot's current trigger is {0} {1} {0} (Alt +0{2})", Chr(34), BotVars.TriggerLong, Asc(BotVars.TriggerLong)))
        Else
            Command_Renamed.Respond(StringFormat("The bot's current trigger is {0} {1} {0} (Length: {2})", Chr(34), BotVars.TriggerLong, Len(BotVars.TriggerLong)))
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUptime(ByRef Command_Renamed As clsCommandObj)
		Command_Renamed.Respond(StringFormat("System uptime {0}, connection uptime {1}.", ConvertTime(GetUptimeMS), ConvertTime(uTicks)))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWhere(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsLocal) Then
			Call frmChat.AddQ("/where " & Command_Renamed.Args, (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), "(console)")
		End If
		
		Command_Renamed.Respond(StringFormat("I am currently in channel {0} ({1} users present)", g_Channel.Name, g_Channel.Users.Count()))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWhoAmI(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		
		If (Command_Renamed.IsLocal) Then
			Command_Renamed.Respond("You are the bot console.")
			
			If (g_Online) Then
				Call frmChat.AddQ("/whoami", (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Username)
			If (dbAccess.Rank = 1000) Then
				Command_Renamed.Respond(StringFormat("You are the bot owner, {0}.", Command_Renamed.Username))
			Else
				If (dbAccess.Rank > 0) Then
                    If (Len(dbAccess.Flags) > 0) Then
                        Command_Renamed.Respond(StringFormat("{0} holds rank {1} and flags {2}.", Command_Renamed.Username, dbAccess.Rank, dbAccess.Flags))
                    Else
                        Command_Renamed.Respond(StringFormat("{0} holds rank {1}.", Command_Renamed.Username, dbAccess.Rank))
                    End If
				Else
                    If (Len(dbAccess.Flags) > 0) Then
                        Command_Renamed.Respond(StringFormat("{0} has flags {1}.", Command_Renamed.Username, dbAccess.Flags))
                    Else
                        Command_Renamed.Respond(StringFormat("{0} has no rank or flags.", Command_Renamed.Username))
                    End If
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnWhoIs(ByRef Command_Renamed As clsCommandObj)
		Dim dbAccess As udtGetAccessResponse
		
		If (Command_Renamed.IsValid) Then
			If (Command_Renamed.IsLocal) Then
				Call frmChat.AddQ("/whois " & Command_Renamed.Argument("Username"), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Argument("Username"))
			
            If (Len(dbAccess.Username) > 0) Then
                If (dbAccess.Rank > 0) Then
                    If (Len(dbAccess.Flags) > 0) Then
                        Command_Renamed.Respond(dbAccess.Username & " holds rank " & dbAccess.Rank & " and flags " & dbAccess.Flags & ".")
                    Else
                        Command_Renamed.Respond(dbAccess.Username & " holds rank " & dbAccess.Rank & ".")
                    End If
                Else
                    If (Len(dbAccess.Flags) > 0) Then
                        Command_Renamed.Respond(dbAccess.Username & " has flags " & dbAccess.Flags & ".")
                    End If
                End If
            Else
                Command_Renamed.Respond("There was no such user found.")
            End If
		End If
	End Sub
	
	Private Function GetAllCommandsFor(ByRef commandDoc As clsCommandDocObj, Optional ByRef Rank As Short = -1, Optional ByRef Flags As String = vbNullString) As String
		On Error GoTo ERROR_HANDLER
		
		Dim tmpbuf As String
		Dim i As Short
		Dim xmlDoc As MSXML2.DOMDocument60
		Dim commands As MSXML2.IXMLDOMNodeList
		Dim xpath As String
		Dim lastCommand As String
		Dim thisCommand As String
		Dim xFunction As String
		Dim AZ As String
		Dim Flag As String
		AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		
		'If (LenB(Dir$(GetFilePath(FILE_COMMANDS))) = 0) Then
		'    Command.Respond "Error: The XML database could not be found in the working directory."
		'    Exit Function
		'End If
		
        If (Len(Flags) > 0) Then
            If (BotVars.CaseSensitiveFlags) Then
                xFunction = "text()='{0}'"
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xFunction = StringFormat("translate(text(), '{0}', '{1}')='{2}'", UCase(AZ), LCase(AZ), "{0}")
            End If

            For i = 1 To Len(Flags)
                Flag = IIf(Mid(Flags, i, 1) = "\", "\\", Mid(Flags, i, 1))
                If (Not BotVars.CaseSensitiveFlags) Then Flag = LCase(Flag)
                If (Flag = "'") Then Flag = "&apos;"

                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xpath = StringFormat("{0}{1}{2}", xpath, StringFormat(xFunction, Flag), IIf(i = Len(Flags), vbNullString, " or "))
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xpath = StringFormat("./command/access/flags/flag[{0}]", xpath)
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xpath = StringFormat("./command/access/rank[number() <= {0}]", Rank)
        End If
		
		xmlDoc = commandDoc.XMLDocument
		
		commands = xmlDoc.documentElement.selectNodes(xpath)
		
		If (commands.length > 0) Then
			For i = 0 To commands.length - 1
                If (Len(Flags) > 0) Then
                    thisCommand = commands(i).parentNode.parentNode.parentNode.attributes.getNamedItem("name").text
                Else
                    thisCommand = commands(i).parentNode.parentNode.attributes.getNamedItem("name").text
                End If
				
				If (StrComp(thisCommand, lastCommand, CompareMethod.Text) <> 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tmpbuf = StringFormat("{0}{1}, ", tmpbuf, thisCommand)
				End If
				
				lastCommand = thisCommand
			Next i
			
			tmpbuf = Left(tmpbuf, Len(tmpbuf) - 2)
		End If
		GetAllCommandsFor = tmpbuf
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.GetAllCommandsFor().")
	End Function
	
	Public Function GetPing(ByVal Username As String) As Integer
		Dim i As Short
		
		i = g_Channel.GetUserIndex(Username)
		
		If i > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetPing = g_Channel.Users.Item(i).Ping
		Else
			GetPing = -3
		End If
	End Function
	
	Private Sub SearchDatabase(ByRef arrReturn() As String, Optional ByRef Username As String = vbNullString, Optional ByVal match As String = vbNullString, Optional ByRef Group As String = vbNullString, Optional ByRef dbType As String = vbNullString, Optional ByRef lowerBound As Short = -1, Optional ByRef upperBound As Short = -1, Optional ByRef Flags As String = vbNullString)
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		Dim found As Short
		Dim tmpbuf As String
		ReDim arrReturn(0)
		
        Dim dbAccess As udtGetAccessResponse
		Dim res As Boolean
		Dim blnChecked As Boolean
		Dim j As Short
        If (Len(Username) > 0) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dbAccess = GetAccess(Username, dbType)

            If (Not (dbAccess.Type = "%") And (Not StrComp(dbAccess.Type, "USER", CompareMethod.Text) = 0)) Then
                dbAccess.Username = dbAccess.Username & " (" & LCase(dbAccess.Type) & ")"
            End If

            If (dbAccess.Rank > 0) Then
                tmpbuf = "Found user " & dbAccess.Username & ", who holds rank " & dbAccess.Rank & IIf(Len(dbAccess.Flags) > 0, " and flags " & dbAccess.Flags, vbNullString) & "."
            ElseIf (Len(dbAccess.Flags) > 0) Then
                tmpbuf = "Found user " & dbAccess.Username & ", with flags " & dbAccess.Flags & "."
            Else
                tmpbuf = "No such user(s) found."
            End If
        Else
            For i = LBound(DB) To UBound(DB)

                If (Len(DB(i).Username) > 0) Then
                    If (Len(match) > 0) Then
                        If (Left(match, 1) = "!") Then
                            res = (Not (LCase(PrepareCheck(DB(i).Username)) Like (LCase(Mid(match, 2)))))
                        Else
                            res = (LCase(PrepareCheck(DB(i).Username)) Like (LCase(match)))
                        End If
                        blnChecked = True
                    End If

                    If (Len(Group) > 0) Then
                        If (StrComp(DB(i).Groups, Group, CompareMethod.Text) = 0) Then
                            res = IIf(blnChecked, res, True)
                        Else
                            res = False
                        End If
                        blnChecked = True
                    End If

                    If (Len(dbType) > 0) Then
                        If (StrComp(DB(i).Type, dbType, CompareMethod.Text) = 0) Then
                            res = IIf(blnChecked, res, True)
                        Else
                            res = False
                        End If
                        blnChecked = True
                    End If

                    If ((lowerBound >= 0) And (upperBound >= 0)) Then
                        If ((DB(i).Rank >= lowerBound) And (DB(i).Rank <= upperBound)) Then
                            res = IIf(blnChecked, res, True)
                        Else
                            res = False
                        End If
                        blnChecked = True
                    ElseIf (lowerBound >= 0) Then
                        If (DB(i).Rank = lowerBound) Then
                            res = IIf(blnChecked, res, True)
                        Else
                            res = False
                        End If
                        blnChecked = True
                    End If

                    If (Len(Flags) > 0) Then

                        For j = 1 To Len(Flags)
                            If (InStr(1, DB(i).Flags, Mid(Flags, j, 1), CompareMethod.Binary) = 0) Then
                                Exit For
                            End If
                        Next j

                        If (j = (Len(Flags) + 1)) Then
                            res = IIf(blnChecked, res, True)
                        Else
                            res = False
                        End If
                        blnChecked = True
                    End If

                    If (res = True) Then
                        tmpbuf = tmpbuf & DBUserToString(DB(i).Username, DB(i).Type)
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        tmpbuf = StringFormat("{0}{1}{2}, ", tmpbuf, IIf(DB(i).Rank > 0, "\" & DB(i).Rank, vbNullString), IIf(Len(DB(i).Flags) > 0, "\" & DB(i).Flags, vbNullString))
                        found = (found + 1)
                    End If
                End If

                res = False
                blnChecked = False
            Next i

            If (found = 0) Then
                arrReturn(0) = "No such user(s) found."
            Else
                Call SplitByLen(Mid(tmpbuf, 1, Len(tmpbuf) - Len(", ")), 180, arrReturn, "User(s) found: ", , ", ")
            End If
        End If
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandCode.SearchDatabase().")
	End Sub
	
	' Gets the number of non-null items in a list. Optionally passes the list formatted as a string.
	Public Function GetListCount(ByRef aList() As String, Optional ByRef sList As String = "") As Short
		Dim i As Short
		Dim iCount As Short
		sList = vbNullString
		
		iCount = 0
		For i = LBound(aList) To UBound(aList)
            If (Len(Trim(aList(i))) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sList = StringFormat("{0}{1}, ", sList, aList(i))
                iCount = iCount + 1
            End If
		Next i
		sList = Left(sList, Len(sList) - 2)
		GetListCount = iCount
	End Function
	
	' Outputs the specified list in response to the given command.
	'   If more than iFailOnSize messages will be output (and the value is > -1), or there are no items, the function will fail.
	'   If the prefix contains {%p} that will be replaced with the formatted position of the output
	'     (x/y, where x is the number of the current message and y is the total number of messages)
	Public Function ListRespond(ByRef oCommand As clsCommandObj, ByRef aList() As String, Optional ByVal sPrefix As String = vbNullString, Optional ByVal iFailOnSize As Short = -1) As Boolean
		Dim sBuffer As String
		Dim aBuffer() As String
		Dim iCount As Short
		Dim i As Short
		
		ListRespond = False
		
		iCount = GetListCount(aList, sBuffer)
		If (iCount > 0) Then
			If (oCommand.IsLocal) Then
				oCommand.Respond(StringFormat("{0}{1}", Replace(sPrefix, "{%p}", vbNullString), sBuffer))
				ListRespond = True
			Else
				Call SplitByLen(sBuffer, 200, aBuffer, sPrefix)
				If iFailOnSize = -1 Or (UBound(aBuffer) + 1) < iFailOnSize Then
					For i = LBound(aBuffer) To UBound(aBuffer)
						If UBound(aBuffer) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat( ({0}/{1}), i + 1, UBound(aBuffer) + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							oCommand.Respond(Replace(aBuffer(i), "{%p}", StringFormat(" ({0}/{1})", i + 1, UBound(aBuffer) + 1)))
						Else
							oCommand.Respond(Replace(aBuffer(i), "{%p}", vbNullString))
						End If
					Next i
					ListRespond = True
				End If
			End If
		End If
	End Function
End Module