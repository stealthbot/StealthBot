Option Strict Off
Option Explicit On
Module modCommandsMisc
	'This modules holds all other command code that i couldnt think of which catergory it fell info :P
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAddQuote(ByRef Command_Renamed As clsCommandObj)
		If (g_Quotes Is Nothing) Then
			g_Quotes = New clsQuotesObj
		End If
		If (Command_Renamed.IsValid) Then
			g_Quotes.Add(Command_Renamed.Argument("Quote"))
			Command_Renamed.Respond("Quote added!")
		Else
			Command_Renamed.Respond("You must provide a quote to add.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnBMail(ByRef Command_Renamed As clsCommandObj)
		Dim temp As udtMail
		Dim strArray() As String
		
		If (Command_Renamed.IsValid) Then
			If (Command_Renamed.IsLocal) Then
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (LenB(Command_Renamed.Username) = 0) Then Command_Renamed.Username = BotVars.Username
			End If
			With temp
				.To_Renamed = Command_Renamed.Argument("Recipient")
				.From = Command_Renamed.Username
				.Message = Command_Renamed.Argument("Message")
			End With
			Command_Renamed.Respond(StringFormat("Added mail for {0}.", Trim(temp.To_Renamed)))
			Call AddMail(temp)
		Else
			Command_Renamed.Respond("Error: You must supply a recipient and a message.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnCancel(ByRef Command_Renamed As clsCommandObj)
		If (VoteDuration > 0) Then
			Command_Renamed.Respond(Voting(BVT_VOTE_END, BVT_VOTE_CANCEL))
		Else
			Command_Renamed.Respond("No vote in progress.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnCheckMail(ByRef Command_Renamed As clsCommandObj)
		Dim Count As Short
		
		Count = GetMailCount(IIf(Command_Renamed.IsLocal, GetCurrentUsername, Command_Renamed.Username))
		If (Count > 0) Then
			Command_Renamed.Respond(StringFormat("You have {0} new message{1}. Type {2}inbox to retrieve {3}.", Count, IIf(Count > 1, "s", vbNullString), IIf(Command_Renamed.IsLocal, "/", BotVars.Trigger), IIf(Count > 1, "them", "it")))
		Else
			Command_Renamed.Respond("You have no mail.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnExec(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		Dim ErrType As String
		
		If (Command_Renamed.IsValid) Then frmChat.SControl.ExecuteStatement(Command_Renamed.Argument("Code"))
		
		Exit Sub
		
ERROR_HANDLER: 
		
		With frmChat.SControl
			ErrType = "runtime"
			
			If InStr(1, .Error.source, "compilation", CompareMethod.Binary) > 0 Then ErrType = "parsing"
			
			Command_Renamed.Respond(StringFormat("Execution {0} error #{1}: {2}", ErrType, .Error.Number, .Error.description))
			
			.Error.Clear()
		End With
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFlip(ByRef Command_Renamed As clsCommandObj)
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Command_Renamed.Respond(IIf(((Rnd() * 1000) Mod 2) = 0, "Heads.", "Tails."))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnGreet(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Select Case LCase(Command_Renamed.Argument("SubCommand"))
				Case "on"
					BotVars.UseGreet = True
					Command_Renamed.Respond("Greet messages enabled.")
					
				Case "off"
					BotVars.UseGreet = False
					Command_Renamed.Respond("Greet messages disabled.")
					
				Case "whisper"
					Select Case (LCase(Command_Renamed.Argument("Value")))
						Case "on"
							BotVars.WhisperGreet = True
						Case "off"
							BotVars.WhisperGreet = False
						Case "toggle", vbNullString
							BotVars.WhisperGreet = Not BotVars.WhisperGreet
					End Select
					
					If BotVars.WhisperGreet Then
						Command_Renamed.Respond("Greet messages will be whispered.")
					Else
						Command_Renamed.Respond("Gree messages will not be whispered.")
					End If
					
				Case "status", vbNullString
					If (BotVars.UseGreet) Then
						If (BotVars.WhisperGreet) Then
							Command_Renamed.Respond("Greet messages are currently enabled, and whispered.")
						Else
							Command_Renamed.Respond("Greet messages are currently enabled, and public.")
						End If
					Else
						Command_Renamed.Respond("Greet messages are currently disabled.")
					End If
					
				Case "set"
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					If (LenB(Command_Renamed.Argument("Value")) > 0) Then
						BotVars.GreetMsg = Command_Renamed.Argument("Value")
						Command_Renamed.Respond("Greet message set.") '1732 - 10/15/2009 - Hdx - Greet Set command will respond now.
					Else
						Command_Renamed.Respond("You must supply a greet message.")
					End If
			End Select
			
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If ((LenB(Command_Renamed.Argument("SubCommand")) > 0) And (LCase(Command_Renamed.Argument("SubCommand")) <> "status")) Then
				Config.GreetMessage = BotVars.UseGreet
				Config.WhisperGreet = BotVars.WhisperGreet
				Config.GreetMessageText = BotVars.GreetMsg
				Call Config.Save()
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIdle(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(Command_Renamed.Argument("Enable")) > 0) Then
				Select Case LCase(Command_Renamed.Argument("Enable"))
					Case "on", "true"
						Config.IdleMessage = True
					Case "off", "false"
						Config.IdleMessage = False
					Case "toggle"
						Config.IdleMessage = Not Config.IdleMessage
				End Select
				
				If Config.IdleMessage Then
					Command_Renamed.Respond("Idle messages enabled.")
				Else
					Command_Renamed.Respond("Idle messages disabled.")
				End If
				
				Call Config.Save()
			Else
				Command_Renamed.Respond("Idle messages are currently " & IIf(Config.IdleMessage, "enabled", "disabled") & ".")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIdleTime(ByRef Command_Renamed As clsCommandObj)
		Dim delay As Short
		If (Command_Renamed.IsValid) Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(Command_Renamed.Argument("Delay")) > 0) Then
				delay = Val(Command_Renamed.Argument("Delay"))
				Config.IdleMessageDelay = (delay * 2)
				Call Config.Save()
			End If
			
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Command_Renamed.Respond(StringFormat("The idle message wait time {0} {1} minute{2}.", IIf(LenB(Command_Renamed.Argument("Delay")) > 0, "has been set to", "is"), delay, IIf(delay > 1, "s", vbNullString)))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnIdleType(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Select Case (LCase(Command_Renamed.Argument("Type")))
				Case "msg", "message"
					Config.IdleMessageType = "msg"
					Command_Renamed.Respond("Idle message type set to: message")
					
				Case "quote", "quotes"
					Config.IdleMessageType = "quote"
					Command_Renamed.Respond("Idle message type set to: quote")
					
				Case "uptime"
					Config.IdleMessageType = "uptime"
					Command_Renamed.Respond("Idle message type set to: uptime")
					
				Case "mp3"
					Config.IdleMessageType = "mp3"
					Command_Renamed.Respond("Idle message type set to: MP3")
					
				Case vbNullString
					Command_Renamed.Respond("The idle message type is currently: " & Config.IdleMessageType)
				Case Else
					Command_Renamed.Respond("Unknown idle message type. Type must be: msg, quote, uptime, or mp3")
			End Select
			
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnInbox(ByRef Command_Renamed As clsCommandObj)
		Dim Msg As udtMail
		Dim mcount As Short
		Dim Index As Short
		Dim dbAccess As udtGetAccessResponse
		
		If (Command_Renamed.IsLocal) Then
			Command_Renamed.Username = IIf(g_Online, GetCurrentUsername, BotVars.Username)
			dbAccess.Rank = 201
			dbAccess.Flags = "A"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dbAccess = GetCumulativeAccess(Command_Renamed.Username)
		End If
		
		If (GetMailCount(Command_Renamed.Username) > 0) Then
			Do 
				GetMailMessage(Command_Renamed.Username, Msg)
				If (Len(RTrim(Msg.To_Renamed)) > 0) Then
					Command_Renamed.Respond(StringFormat("Message from: {0}: {1}", Trim(Msg.From), Trim(Msg.Message)))
				End If
			Loop While (GetMailCount(Command_Renamed.Username) > 0)
		Else
			If (dbAccess.Rank > 0) Then
				Command_Renamed.Respond("You do not currently have any messages in your inbox.")
			End If
		End If
		
		If (Not Command_Renamed.IsLocal) Then
			Command_Renamed.WasWhispered = True
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnMath(ByRef Command_Renamed As clsCommandObj)
		' This command will execute a specified mathematical statement using the
		' restricted script control, SCRestricted, on frmChat.  The execution
		' of any code through direct user-interaction can become quite error-prone
		' and, as such, this command requires its own error handler.  The input
		' of this command must also be properly sanitized to ensure that no
		' harmful statements are inadvertently allowed to launch on the user's
		' machine.
		
		On Error GoTo ERROR_HANDLER
		
		Dim sStatement As String
		Dim sResult As String
		
		If (Command_Renamed.IsValid) Then
			sStatement = Command_Renamed.Argument("Expression")
			
			If (InStr(1, sStatement, "CreateObject", CompareMethod.Text)) Then
				Command_Renamed.Respond("Evaluation error, CreateObject is restricted.")
			Else
				With frmChat.SCRestricted
					.AllowUI = False
					.UseSafeSubset = True
				End With
				sStatement = Replace(sStatement, Chr(34), vbNullString)
				'UPGRADE_WARNING: Couldn't resolve default property of object frmChat.SCRestricted.Eval(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sResult = frmChat.SCRestricted.Eval(sStatement)
				
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (LenB(sResult) > 0) Then
					Command_Renamed.Respond(StringFormat("The statement {0}{1}{0} evaluates to: {2}.", Chr(34), sStatement, sResult))
				Else
					Command_Renamed.Respond("Evaluation error.")
				End If
			End If
		End If
		Exit Sub
		
ERROR_HANDLER: 
		Command_Renamed.Respond("Evaluation error.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnMMail(ByRef Command_Renamed As clsCommandObj)
		Dim temp As udtMail
		Dim Rank As Integer
		Dim Flags As String
		Dim i As Short
		Dim x As Short
		Dim dbAccess As udtGetAccessResponse
		
		If (Command_Renamed.IsValid) Then
			If (Command_Renamed.IsLocal) Then
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (LenB(Command_Renamed.Username) = 0) Then Command_Renamed.Username = BotVars.Username
			End If
			
			temp.From = Command_Renamed.Username
			temp.Message = Command_Renamed.Argument("Message")
			
			If (StrictIsNumeric(Command_Renamed.Argument("Criteria"))) Then
				Rank = Val(Command_Renamed.Argument("Criteria"))
				
				For i = 0 To UBound(DB)
					If (StrComp(DB(i).Type, "USER", CompareMethod.Text) = 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dbAccess = GetCumulativeAccess(DB(i).Username)
						If (dbAccess.Rank = Rank) Then
							temp.To_Renamed = DB(i).Username
							Call AddMail(temp)
						End If
					End If
				Next i
				Command_Renamed.Respond(StringFormat("Mass mailing to users with rank {0} complete.", Rank))
			Else
				Flags = Command_Renamed.Argument("Criteria")
				For i = 0 To UBound(DB)
					If (StrComp(DB(i).Type, "USER", CompareMethod.Text) = 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						dbAccess = GetCumulativeAccess(DB(i).Username)
						For x = 1 To Len(Flags)
							If (InStr(1, dbAccess.Flags, Mid(Flags, x, 1), IIf(BotVars.CaseSensitiveFlags, CompareMethod.Binary, CompareMethod.Text)) > 0) Then
								temp.To_Renamed = DB(i).Username
								Call AddMail(temp)
								Exit For
							End If
						Next x
					End If
				Next i
				Command_Renamed.Respond(StringFormat("Mass mailing to users with any of the flags {0} complete.", Flags))
			End If
		Else
			Command_Renamed.Respond(StringFormat("Format: {0}mmail <flag(s)> <message> OR {0}mmail <access> <message>", IIf(Command_Renamed.IsLocal, "/", BotVars.Trigger)))
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnQuote(ByRef Command_Renamed As clsCommandObj)
		Dim tmpQuote As String
		tmpQuote = "Quote: " & g_Quotes.GetRandomQuote
		
		If (Len(tmpQuote) < 8) Then
			Command_Renamed.Respond("Error reading your quotes, or no quote file exists.")
		Else
			Command_Renamed.Respond(tmpQuote)
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnReadFile(ByRef Command_Renamed As clsCommandObj)
		On Error GoTo ERROR_HANDLER
		Dim sFilePath As String
		Dim iFile As Integer
		Dim sLine As String
		Dim iLineNumber As Short
		
		If (Command_Renamed.IsValid) Then
			sFilePath = Command_Renamed.Argument("File")
			If ((Not Mid(sFilePath, 2, 2) = ":\") And (Not Mid(sFilePath, 2, 2) = ":/")) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sFilePath = StringFormat("{0}\{1}", CurDir(), sFilePath)
			End If
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(Dir(sFilePath)) > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Respond(StringFormat("Contents of file {0}:", Replace(sFilePath, StringFormat("{0}\", CurDir()), vbNullString,  , 1, CompareMethod.Text)))
				
				iFile = FreeFile
				FileOpen(iFile, sFilePath, OpenMode.Input)
				Do While (Not EOF(iFile))
					sLine = LineInput(iFile)
					'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					If (LenB(sLine) > 0) Then
						iLineNumber = iLineNumber + 1
						Command_Renamed.Respond(StringFormat("Line {0}: {1}", iLineNumber, sLine))
					End If
				Loop 
				FileClose(iFile)
				
				Command_Renamed.Respond("End of File.")
			Else
				Command_Renamed.Respond("Error: The specified file could not be found.")
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		Command_Renamed.ClearResponse()
		Command_Renamed.Respond("There was an error reading the specified file.")
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnRoll(ByRef Command_Renamed As clsCommandObj)
		Dim maxValue As Integer
		Dim Number As Integer
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(Command_Renamed.Argument("Value")) > 0) Then
			maxValue = System.Math.Abs(Val(Command_Renamed.Argument("Value")))
		Else
			maxValue = 100
		End If
		
		Randomize()
		Number = CInt(Rnd() * maxValue)
		
		Command_Renamed.Respond(StringFormat("Random number (0-{0}): {1}", maxValue, Number))
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetIdle(ByRef Command_Renamed As clsCommandObj)
		If (Command_Renamed.IsValid) Then
			Config.IdleMessageText = Command_Renamed.Argument("Message")
			Call Config.Save()
			
			Command_Renamed.Respond("Idle message set.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnTally(ByRef Command_Renamed As clsCommandObj)
		If (VoteDuration > 0) Then
			Command_Renamed.Respond(Voting(BVT_VOTE_TALLY))
		Else
			Command_Renamed.Respond("No vote is currently in progress.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnVote(ByRef Command_Renamed As clsCommandObj)
		Dim Duration As Short
		If (Command_Renamed.IsValid) Then
			If (VoteDuration = -1) Then
				Duration = Val(Command_Renamed.Argument("Duration"))
				If ((Duration > 0) And (Duration < 32000)) Then
					VoteDuration = Duration
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
					Command_Renamed.Respond(Voting(BVT_VOTE_START, BVT_VOTE_STD, vbNullString))
				Else
					Command_Renamed.Respond("Vote durations must be between 1 and 32000")
				End If
			Else
				Command_Renamed.Respond("A vote is currently in progress.")
			End If
		Else
			Command_Renamed.Respond("Please enter a number of seconds for your vote to last.")
		End If
	End Sub
End Module