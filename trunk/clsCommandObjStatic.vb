Option Strict Off
Option Explicit On
Friend Class clsCommandObjStatic
	
	'// This function returns a clsCommandObj that is populated with instance of the command
	'// object.
	Public Function IsCommand(ByVal strText As String, ByVal strUsername As String, ByVal IsLocal As Boolean, ByVal WasWhispered As Boolean, Optional ByVal strScriptOwner As String = vbNullString) As Collection
		
		On Error GoTo ERROR_HANDLER
		
		Const CMD_DELIMITER As String = "; "
		
		Dim Message As String '// the raw message
		Dim messageLen As Short '// the length of the raw message
		Dim cropLen As Short '// the length of the trigger
		Dim hasTrigger As Boolean '// if true, a trigger has been found
		Dim botUsername As String '// stores the bot's username returned from modOtherCode.GetCurrentUsername
		Dim botRawUsername As String '// stores the bot's username retrieved from modGlobals.CurrentUsername
		
		
		'// used for creating command instaces
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As clsCommandObj
		Dim commandIndex As Short
		Dim commandString As String
		Dim commandStrings As Collection
		Dim PublicOutput As Boolean '// if true, then the output should be sent to the channel
		
		
		IsCommand = New Collection
		
		'// make sure actual text was passed for the message, otherwise we return an empty collection
		If (strText = vbNullString) Then
			'// Not a command
			Exit Function
		End If
		
		'// get the bot's username into variables. botUsername will be the bots username without any
		'// domain or product info (like * for D2 and @USEast/@Azeroth etc). botRawUsername will contain
		'// this information. The commands should trigger from both versions.
		botUsername = modGlobals.CurrentUsername
		botRawUsername = modOtherCode.GetCurrentUsername()
		
		If (WasWhispered And (StrComp(botUsername, strUsername, CompareMethod.Text) = 0)) Then Exit Function
		
		hasTrigger = False
		PublicOutput = False
		Message = strText
		messageLen = Len(Message)
		
		'// If this command was entered via the bot we need to check for slashes.
		'//
		'// 0 slashes - No further processing
		'// 1 slashes - publicOutput = False
		'// 2 slashes - publicOutput = True
		'// 3 slashes - No further processing
		If (IsLocal = True) Then
			
			'// commands entered through the bot should use a /
			If (Not Left(Message, 1) = "/") Then
				'// Not a command
				Exit Function
			End If
			
			'// make sure we do no further processing if the message is nothing but 1 slash
			If (Left(Message, 1) = "/") And messageLen = 1 Then
				'// Not a command
				Exit Function
			End If
			
			'// make sure we do no further processing if the message is nothing but 2 slashes
			If (Left(Message, 2) = "//") And messageLen = 2 Then
				'// Not a command
				Exit Function
			End If
			
			'// make sure we do no further processing if the message begins with ///
			If (Left(Message, 3) = "///") Then
				'// Not a command
				Exit Function
			End If
			
			'// at this point, if message begins with // than public output should be true
			If (Left(Message, 2) = "//") Then
				PublicOutput = True
				cropLen = 3
				hasTrigger = True
			Else
				cropLen = 2
				hasTrigger = True
			End If
			
		End If '// (IsLocal = True)
		
		
		'// if this command was not entered via the bot, then we need to check for the bot trigger
		'// as well as for the 2 global triggers, ops and all.
		If (IsLocal = False) Then
			
			'// check for bot trigger
			'// EXAMPLE COMMAND STRING
			'// .add SammyHagar 200
			If (Left(Message, Len(BotVars.TriggerLong)) = BotVars.TriggerLong) Then
				cropLen = Len(BotVars.TriggerLong) + 1
				hasTrigger = True
			End If
			
			'// check for "all: " or "all, ". These special triggers work for all bots
			'// EXAMPLE COMMAND STRING
			'// all: add SammyHagar 200
			If (hasTrigger = False) And (messageLen > 5) Then
				If (StrComp(Left(Message, 3), "all", CompareMethod.Text) = 0) And (Mid(Message, 4, 2) = ": " Or Mid(Message, 4, 2) = ", ") Then
					cropLen = 6
					hasTrigger = True
				End If
			End If
			
			'// check for "ops: " or "ops, ". These special triggers work for all bots that are operators
			'// EXAMPLE COMMAND STRING
			'// ops: add SammyHagar 200
			If (hasTrigger = False) And (messageLen > 5) Then
				If (StrComp(Left(Message, 3), "ops", CompareMethod.Text) = 0) And (Mid(Message, 4, 2) = ": " Or Mid(Message, 4, 2) = ", ") Then
					If (g_Channel.Self.IsOperator) Then
						cropLen = 6
						hasTrigger = True
					End If
				End If
			End If
			
			'// check for bots name as a trigger.
			'// EXAMPLE COMMAND STRING
			'// FiftyToo: add SammyHagar 200
			If (hasTrigger = False) And (messageLen > Len(botUsername) + 2) Then
				If StrComp(Left(Message, Len(botUsername)), botUsername, CompareMethod.Text) = 0 And (Mid(Message, Len(botUsername) + 1, 2) = ": " Or Mid(Message, Len(botUsername) + 1, 2) = ", ") Then
					
					cropLen = Len(botUsername) + 3
					hasTrigger = True
					
				End If
			End If
			
			'// check for bots name as a trigger, with respect to product and realm
			'// EXAMPLE COMMAND STRING
			'// *FiftyToo: add SammyHagar 200
			If (hasTrigger = False) And (messageLen > Len(botRawUsername) + 2) Then
				If StrComp(Left(Message, Len(botRawUsername)), botRawUsername, CompareMethod.Text) = 0 And (Mid(Message, Len(botRawUsername) + 1, 2) = ": " Or Mid(Message, Len(botRawUsername) + 1, 2) = ", ") Then
					
					cropLen = Len(botRawUsername) + 3
					hasTrigger = True
					
				End If
			End If
			
			'// check for a pattern that matches the bot username
			'// EXAMPLE COMMAND STRING (matches fiftytoo followed by any 3 numbers)
			'// FiftyToo###: add SammyHagar 200
			If (hasTrigger = False) And InStr(1, Message, ": ") > 0 Then
				If (UsernameRegex(botUsername, Left(Message, InStr(1, Message, ": ") - 1)) Or UsernameRegex(botRawUsername, Left(Message, InStr(1, Message, ": ") - 1))) Then
					
					cropLen = InStr(1, Message, ": ") + 2
					hasTrigger = True
				End If
			End If
			If (hasTrigger = False) And InStr(1, Message, ", ") > 0 Then
				If (UsernameRegex(botUsername, Left(Message, InStr(1, Message, ", ") - 1)) Or UsernameRegex(botRawUsername, Left(Message, InStr(1, Message, ", ") - 1))) Then
					
					cropLen = InStr(1, Message, ", ") + 2
					hasTrigger = True
				End If
			End If
			
			'// check for ?trigger and !inbox
			If (StrComp(Message, "?trigger", CompareMethod.Text) = 0) Or (StrComp(Message, "!inbox", CompareMethod.Text) = 0) Then
				
				cropLen = 2
				hasTrigger = True
			End If
			
			
			'// if we have not found a trigger, lets get out of here
			If (hasTrigger = False) Then
				'// Not a command
				Exit Function
			End If
			
		End If '// (IsLocal = False)
		
		
		'// get a collection of commands based on the split logic
		commandStrings = SplitCompleteCommandString(Mid(Message, cropLen))
		
		'// if this command string has multiple commands, lets parse them out and process
		'// them individually.
		For commandIndex = 1 To commandStrings.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object commandStrings(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			commandString = commandStrings.Item(commandIndex)
			'// lets try to parse this command and add it to the collection
			Command_Renamed = CreateCommandInstance(commandString, strUsername, strScriptOwner)
			If Not (Command_Renamed Is Nothing) Then
				'// we only want to add the command if it is enabled
				If Command_Renamed.docs.IsEnabled = True Then
					Command_Renamed.PublicOutput = PublicOutput
					Command_Renamed.IsLocal = IsLocal
					Command_Renamed.WasWhispered = WasWhispered
					IsCommand.Add(Command_Renamed)
				End If
			End If
		Next 
		
		'// all done here
		Exit Function
		
ERROR_HANDLER: 
		If (Err.Number = 93) Then
			Err.Clear()
			Exit Function
		End If
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.IsCommand().")
	End Function
	
	
	'// This function will split a message into separate commands. This message should NOT
	'// have any triggers. This will return a collection of command strings that can be used
	'// to create an instance of a command.
	'//
	'// TODO:
	'// Fix logic to allow "; " inside a quoted argument.
	Private Function SplitCompleteCommandString(ByVal completeCommandString As String) As Collection
		
		Dim i As Short
		Dim commandString As String
		Dim commandStrings() As String
		
		SplitCompleteCommandString = New Collection
		
		'// use "; " as a delimiter for commands. Allows for /; to escape a command split
		completeCommandString = Replace(completeCommandString, "\;", Chr(0))
		commandStrings = Split(completeCommandString, "; ")
		For i = LBound(commandStrings) To UBound(commandStrings)
			'// make sure these some actual text for this command, otherwise skip it
			If Len(commandStrings(i)) > 0 Then
				SplitCompleteCommandString.Add(Replace(commandStrings(i), Chr(0), ";"))
			End If
		Next i
		
	End Function
	
	'// this function takes the raw args string (everything after the command) and returns a
	'// collection of strings. Each string is a argument that is parsed using the new
	'// argument snytax.
	'// EXAMPLE
	'// mycommand "this is a \"single\" arg" and here are 5 more
	Private Function SplitArguments(ByVal strArgString As String) As Collection
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short '// counter
		Dim L As Object
		Dim r As String '// temp vars to store the left and right characters
		Dim tmp() As String '// array of words
		Dim Word As String '// stores the word
		Dim multiword As String '// stores the text of a multi-word argument
		Dim insideArg As Boolean '// used to check if a word begins a multi-word argument
		
		SplitArguments = New Collection
		
		'// take out any extra spaces
		strArgString = Trim(strArgString)
		
		If Len(strArgString) = 0 Then
			'// no arguments
			Exit Function
		End If
		
		'// if there is no space then we can just strip the quotes (if present), add
		'// it to the collection, and then return
		If InStr(1, strArgString, " ") < 1 Then
			
			Word = StripQuotes(Replace(strArgString, "\""", Chr(0)))
			If InStr(1, Word, """") > 0 Then
				'// this is bad... words cannot contain unescaped "
				'Err.Raise -1, 0&, ""Words cannot contain unescaped """. Args =: " & strArgString
				'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				SplitArguments = Nothing
				SplitArguments = New Collection
				Exit Function
			End If
			
			Word = Replace(Word, Chr(0), """")
			
			SplitArguments.Add(Word)
			Exit Function
		End If
		
		'// default some variables
		insideArg = False
		multiword = ""
		
		'// loop through each element and group the arguments
		tmp = Split(strArgString)
		For i = LBound(tmp) To UBound(tmp)
			Word = tmp(i)
			'// allow for escaping quotes
			Word = Replace(Word, "\""", Chr(0))
			
			'// if the length is 2 or more then then l and r should be the first and last character
			If Len(Word) > 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				L = Left(Word, 1)
				r = Right(Word, 1)
				'// if the length is 1, then we need to set either l or r to "" depending on insideArg
			ElseIf Len(Word) > 0 Then 
				If insideArg = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					L = Left(Word, 1)
					r = ""
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					L = ""
					r = Right(Word, 1)
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				L = ""
				r = ""
			End If
			
			'// check if this word BEGINS with a " and we ARE NOT inside an arg
			'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (L = """") And (r = """") Then
				'// this should be a single argument, if we are inside a word we have a problem
				If insideArg = True Then
					'// this is bad... words cannot contain unescaped "
					'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
					'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					SplitArguments = Nothing
					SplitArguments = New Collection
					Exit Function
				End If
				
				'// ok this is a single word arg, but we still need to fail if it contains a "
				If InStr(1, Word, """") > 0 Then
					'// this is bad... words cannot contain unescaped "
					'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
					'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					SplitArguments = Nothing
					SplitArguments = New Collection
					Exit Function
				End If
				
				Word = Replace(Word, Chr(0), """")
				
				'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (L = """") And (insideArg = False) Then 
				
				'// we are, lets start the multiword and set our bit
				multiword = Word & " "
				insideArg = True
				
				'// check if this word ENDS with a " and we ARE inside an arg
			ElseIf (r = """") And (insideArg = True) Then 
				
				'// we are, lets end the multiword, add it to the collect, and reset our vars
				multiword = multiword & Word
				SplitArguments.Add(Replace(StripQuotes(multiword), Chr(0), """"))
				insideArg = False
				multiword = ""
				
				'// check if we are inside a word, if so then we append it to multi word and be done with it
			ElseIf (insideArg = True) Then 
				
				multiword = multiword & Word & " "
				
				'// we are not inside a word, then this must be a separate argument so we need to add it
				'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (r <> """") And (L <> """") And (insideArg = False) Then 
				
				
				'// make sure this word does not have any " inside it.
				If InStr(1, Word, """") > 0 Then
					'// this is bad... words cannot contain unescaped "
					'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
					'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					SplitArguments = Nothing
					SplitArguments = New Collection
					Exit Function
				End If
				
				'// if there is no text and not inside a word, then we should ignore it
				If Len(Word) > 0 Then
					SplitArguments.Add(Replace(StripQuotes(Word), Chr(0), """"))
					insideArg = False
					multiword = ""
				End If
				'// this should never happen with valid argument syntax
			Else
				'Err.Raise -1, 0&, "Cannot determine type of word. Args =: " & strArgString
				'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				SplitArguments = Nothing
				SplitArguments = New Collection
				Exit Function
				
			End If
			
		Next i
		
		'// final test
		If insideArg = True Then
			'// this is bad... we ended inside an argument
			'Err.Raise -1, 0&, "Ended with an open arg string. Args =: " & strArgString
			'UPGRADE_NOTE: Object SplitArguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			SplitArguments = Nothing
			SplitArguments = New Collection
			Exit Function
		End If
		
		'// all good :)
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.SplitArguments().")
		
		Exit Function
		
	End Function
	
	'// this function takes a string and will return a clsCommandObj object. If strOwnerName is missing or vbNullstring
	'// then this function will check for an internal command. All triggers should be removed from strText
	'// and Len(strText) > 0. This function does NOT consider multiple commands contained inside
	'// strText. All MULTICOMMAND PARSING SHOULD TAKE PLACE PRIOR TO CALLING THIS METHOD. Since this
	'// function does not have triggers
	'//
	'// EXAMPLE:
	'// Set cmd = CreateCommandInstance("add FiftyToo 50", "someUser")
	'// If cmd.IsValidCommand Then
	'//     frmChat.AddChat vbGreen, cmd.Name
	'// End If
	Private Function CreateCommandInstance(ByRef commandString As String, ByVal strUsername As String, Optional ByVal strScriptOwner As String = vbNullString) As clsCommandObj
		
		On Error GoTo ERROR_HANDLER
		
		Dim doc As New clsCommandDocObj
		Dim cmd As clsCommandObj
		
		Dim commandName As String
		Dim commandArgs As String
		Dim tmp() As String
		
		'// separate the command's name and args from the command string
		tmp = Split(commandString, " ", 2)
		commandName = tmp(0)
		If UBound(tmp) = 1 Then
			commandArgs = tmp(1)
		End If
		
		If (Not doc.OpenCommand(commandName, strScriptOwner)) Then
			Exit Function
		End If
		
		'// ok this is actually a command, lets create the object
		cmd = New clsCommandObj
		With cmd
			.Name = doc.Name
			.Args = commandArgs
			'.docs = Me
			.Arguments = SplitArguments(commandArgs)
			.Username = strUsername
		End With
		
		'// all good in the hood :)
		CreateCommandInstance = cmd
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.CreateCommandObject().")
		
		Exit Function
		
	End Function
	
	
	
	'// this function will remove the first and last double quote from a string, but only
	'// if both are present
	Private Function StripQuotes(ByVal strText As String) As String
		
		
		Dim retVal As String
		Dim leftStripped As Boolean
		Dim rightStripped As Boolean
		
		leftStripped = False
		rightStripped = False
		
		retVal = strText
		
		If Left(retVal, 1) = """" Then
			retVal = Mid(retVal, 2)
			leftStripped = True
		End If
		
		If Right(retVal, 1) = """" Then
			retVal = Mid(retVal, 1, Len(retVal) - 1)
			rightStripped = True
		End If
		
		'// if these values are the same, then we can return retval, otherwise we should return
		'// whatever was passed into the function
		If leftStripped = rightStripped Then
			StripQuotes = retVal
		Else
			StripQuotes = strText
		End If
		
	End Function
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function CleanXPathVar(ByVal str_Renamed As String) As String
		str_Renamed = Replace(str_Renamed, "\", "\\")
		str_Renamed = Replace(str_Renamed, "'", "&apos;")
		CleanXPathVar = str_Renamed
	End Function
	
	'Function to check if command names are valid
	Public Function IsValidCommandName(ByRef sName As String) As Boolean
		Dim x As Short
		Dim sValid As String
		
		sValid = "abcdefghijklmnopqrstuvwxyz0123456789_"
		IsValidCommandName = False
		
		For x = 1 To Len(sName)
			If (InStr(1, sValid, Mid(sName, x, 1), CompareMethod.Text) = 0) Then Exit Function
		Next x
		
		IsValidCommandName = True
	End Function
End Class