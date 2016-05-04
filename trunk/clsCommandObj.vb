Option Strict Off
Option Explicit On
Friend Class clsCommandObj
	' clsCommandObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	'// This object is a representation of a command instance. A reference to this object
	'// is returned to a script module by using the abstract IsCommand() method.
	
	
	Private m_valid As Boolean
	Private m_command_docs As clsCommandDocObj
	Private m_name As String
	Private m_Username As String
	Private m_args As String
	Private m_arguments As Collection
	Private m_local As Boolean
	Private m_publicOutput As Boolean
	Private m_xmlarguments As Scripting.Dictionary
	Private m_hasaccess As Boolean
	Private m_splithasrun As Boolean
	Private m_waswhispered As Boolean
	Private m_response As Collection
	Private m_restrictions As Scripting.Dictionary
	
	Public Property Username() As String
		Get
			Username = m_Username
		End Get
		Set(ByVal Value As String)
			m_Username = Value
		End Set
	End Property
	Public Property Name() As String
		Get
			Name = m_name
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	Public Property Args() As String
		Get
			Args = m_args
		End Get
		Set(ByVal Value As String)
			m_args = Value
		End Set
	End Property
	Public Property Arguments() As Collection
		Get
			Arguments = m_arguments
		End Get
		Set(ByVal Value As Collection)
			m_arguments = Value
		End Set
	End Property
	Public Property IsLocal() As Boolean
		Get
			IsLocal = m_local
		End Get
		Set(ByVal Value As Boolean)
			m_local = Value
		End Set
	End Property
	Public Property PublicOutput() As Boolean
		Get
			PublicOutput = m_publicOutput
		End Get
		Set(ByVal Value As Boolean)
			m_publicOutput = Value
		End Set
	End Property
	Public Property IsValid() As Boolean
		Get
			If (Not m_splithasrun) Then SplitArguments()
			IsValid = m_valid
		End Get
		Set(ByVal Value As Boolean)
			m_valid = Value
		End Set
	End Property
	Public ReadOnly Property HasAccess() As Boolean
		Get
			If (Not m_splithasrun) Then SplitArguments()
			HasAccess = m_hasaccess
		End Get
	End Property
	Public Property WasWhispered() As Boolean
		Get
			WasWhispered = m_waswhispered
		End Get
		Set(ByVal Value As Boolean)
			m_waswhispered = Value
		End Set
	End Property
	
	Public ReadOnly Property source() As Short
		Get
			If m_local = True Then
				source = 4
			Else
				If m_waswhispered = True Then
					source = 3
				Else
					source = 1
				End If
			End If
		End Get
	End Property
	
	
	Public Property docs() As clsCommandDocObj
		Get
			If m_command_docs Is Nothing Then
				'// this command is nothing, lets create it
				m_command_docs = New clsCommandDocObj
				Call m_command_docs.OpenCommand(m_name, Chr(0))
				docs = m_command_docs
			Else
				'// this command already has a value, lets make sure its still valid
				If StrComp(m_command_docs.Name, m_name, Scripting.CompareMethod.TextCompare) = 0 Then
					'// all good, lets return it
					docs = m_command_docs
				Else
					'// ugh, this doc object is for a different command, we need to
					'// destroy it and start all over again
					'UPGRADE_NOTE: Object m_command_docs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					m_command_docs = Nothing
					docs = Me.docs
				End If
			End If
		End Get
		Set(ByVal Value As clsCommandDocObj)
			m_command_docs = Value
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Class_Initialize_Renamed()
		'// initialize values
		m_valid = False
		m_command_docs = New clsCommandDocObj
		m_name = vbNullString
		m_Username = vbNullString
		m_args = vbNullString
		m_arguments = New Collection
		m_local = False
		m_publicOutput = False
		m_xmlarguments = New Scripting.Dictionary
		m_xmlarguments.CompareMode = Scripting.CompareMethod.TextCompare
		m_hasaccess = False
		m_splithasrun = False
		m_waswhispered = False
		m_response = New Collection
		m_restrictions = New Scripting.Dictionary
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Class_Terminate_Renamed()
		'// clean up
		'UPGRADE_NOTE: Object m_command_docs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_command_docs = Nothing
		'UPGRADE_NOTE: Object m_arguments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_arguments = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'Strips (removes) and returns a Numeric argument from the passed string
	Private Function StripNumeric(ByRef sString As String) As String
		Dim sTemp As String
		sTemp = StripWord(sString)
		If (StrictIsNumeric(sTemp, True)) Then
			StripNumeric = sTemp
        ElseIf Len(sTemp) > 0 Then
            sString = sTemp & Space(1) & sString
		End If
	End Function
	
	'Strips (removes) and returns a Word argument from the passed string
	Private Function StripWord(ByRef sString As String) As String
		Dim i As Short
		i = InStr(sString, Space(1))
		If (i > 0) Then
			StripWord = Left(sString, i - 1)
			sString = Mid(sString, i + 1)
		ElseIf Len(sString) > 0 Then 
			StripWord = sString
			sString = vbNullString
		End If
	End Function
	
	'Strips (removes) and returns a String argument from the passed string
	'EXAs:
	'This is a String -> This is a String
	'This is a "String" -> This is a "String"
	'"This is a String" -> This is a String
	'"This is a \"String\"" -> This is a "String"
	'"This is a \\\"String\"" -> This is a \"String"
	'"This is a \String" -> This is a \String
	Private Function StripString(ByRef sString As String, ByRef IsLastArgument As Boolean) As String
		Dim i As Short
		If (IsLastArgument) Then
			StripString = sString
			sString = vbNullString
			Exit Function
		End If
		If (Left(sString, 1) = Chr(34)) Then
			sString = Replace(sString, "\\", Chr(1))
			sString = Replace(sString, "\" & Chr(34), Chr(2))
			i = InStr(2, sString & " ", Chr(34) & " ")
			If (i > 2) Then
				sString = Replace(sString, Chr(1), "\")
				sString = Replace(sString, Chr(2), Chr(34))
				StripString = Left(Mid(sString, 2), i - 2)
				sString = Mid(sString, i + 2)
			Else
				sString = Replace(sString, Chr(1), "\")
				sString = Replace(sString, Chr(2), Chr(34))
				StripString = sString
				sString = vbNullString
			End If
		Else
			StripString = sString
			sString = vbNullString
		End If
	End Function
	
	'Will Split up the Arguments for this instance of the command, Based on the XML specs of this command.
	'EXA: .Mail Username This is a message!
	'Creates a Dictionary as such:
	'  Dict("Username") = "Username"
	'  Dict("Message")  = "This is a message!"
	'This also checks the user's access to use this command in the specific restriction context.
	Private Sub SplitArguments()
		On Error GoTo ERROR_HANDLER
		Dim sArgs As String
		Dim i As Short
		Dim Param As clsCommandParamsObj
		Dim Restriction As clsCommandRestrictionObj
		Dim sTemp As String
		Dim sError As String
		Dim dbAccess As udtGetAccessResponse
		'UPGRADE_WARNING: Couldn't resolve default property of object dbAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dbAccess = GetCumulativeAccess(Me.Username)
		sArgs = Me.Args
		
		
		If (IsLocal) Then
			dbAccess.Rank = 201
			dbAccess.Flags = "A"
		End If
		
		If (dbAccess.Rank >= Me.docs.RequiredRank And Me.docs.RequiredRank > -1) Then
			m_hasaccess = True
		End If
		
		If (CheckForAnyFlags(Me.docs.RequiredFlags, dbAccess.Flags)) Then
			m_hasaccess = True
		End If
		
		IsValid = True
		
		For	Each Param In Me.docs.Parameters
			
			Select Case LCase(Param.datatype)
				Case "word" : sTemp = StripWord(sArgs)
				Case "numeric" : sTemp = StripNumeric(sArgs)
				Case "number" : sTemp = StripNumeric(sArgs)
				Case "string"
					'UPGRADE_WARNING: Couldn't resolve default property of object Me.docs.Parameters.Item().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sTemp = StripString(sArgs, Param.Name = Me.docs.Parameters.Item(Me.docs.Parameters.Count()).Name)
			End Select
			
            If (Len(Param.MatchMessage)) Then
                If (Not CheckMatch((Param.MatchMessage), sTemp, (Param.MatchCaseSensitive))) Then
                    If (Len(Param.MatchError) > 0) And (m_hasaccess) Then

                        sError = Replace(Param.MatchError, "%Value", sTemp)
                        sError = Replace(sError, "%Rank", CStr(dbAccess.Rank))
                        sError = Replace(sError, "%Flags", dbAccess.Flags)

                        Respond(sError)

                        m_splithasrun = True
                        IsValid = False
                        Exit Sub
                    Else
                        If (LCase(Param.datatype) = "string") Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            sArgs = StringFormat("{0}{1}{0} {2}", Chr(34), sTemp, sArgs)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            sArgs = StringFormat("{0} {1}", sTemp, sArgs)
                        End If
                    End If
                End If
            End If
			
            If (Len(sTemp) > 0) Then
                For Each Restriction In Param.Restrictions 'Loop Through the Restrictions
                    'UPGRADE_WARNING: Couldn't resolve default property of object m_restrictions(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    m_restrictions(Restriction.Name) = True
                    If (CheckMatch((Restriction.MatchMessage), sTemp, (Restriction.MatchCaseSensitive))) Then 'If they match (regex)
                        'If Rank = -1 It means it's missing, and it MUST have Flags. Or if Rank > User's Access
                        If (Restriction.RequiredRank = -1 Or Restriction.RequiredRank > dbAccess.Rank) Then
                            If (Not CheckForAnyFlags((Restriction.RequiredFlags), dbAccess.Flags)) Then
                                If (Len(Restriction.MatchError)) And (m_hasaccess) Then
                                    sError = Replace(Restriction.MatchError, "%Value", sTemp)
                                    sError = Replace(sError, "%Rank", CStr(dbAccess.Rank))
                                    sError = Replace(sError, "%Flags", dbAccess.Flags)

                                    Respond(sError)
                                End If
                                If (Restriction.Fatal) Then m_hasaccess = False
                                'UPGRADE_WARNING: Couldn't resolve default property of object m_restrictions(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                m_restrictions(Restriction.Name) = False
                            End If
                        End If
                    End If
                Next Restriction
            End If
			
            If (Len(sTemp) = 0 And Not Param.IsOptional) Then
                IsValid = False
            End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_xmlarguments(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_xmlarguments(Param.Name) = sTemp
		Next Param
		
		If (IsLocal) Then m_hasaccess = True
		m_splithasrun = True
		
		Exit Sub
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.SplitArguments().")
	End Sub
	
	Public Function Argument(ByRef sName As String) As String
		If (Not m_splithasrun) Then SplitArguments()
		If (m_xmlarguments.Exists(sName)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_xmlarguments.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Argument = m_xmlarguments.Item(sName)
		Else
			Argument = vbNullString
		End If
	End Function
	
	Public Function Restriction(ByRef sName As String) As Boolean
		If (Not m_splithasrun) Then SplitArguments()
		If (m_restrictions.Exists(sName)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_restrictions.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Restriction = m_restrictions.Item(sName)
		Else
			Restriction = False
		End If
	End Function
	
	Private Function CheckForAnyFlags(ByRef sNeeded As String, ByRef sHave As String) As Boolean
		On Error GoTo ERROR_HANDLER
		Dim i As Short
		CheckForAnyFlags = False
		
        If (Len(sHave) = 0) Then Exit Function
		
		For i = 1 To Len(sNeeded)
			If (InStr(1, sHave, Mid(sNeeded, i, 1), CompareMethod.Text) > 0) Then
				CheckForAnyFlags = True
				Exit Function
			End If
		Next 
		
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.CheckForAnyFlags().")
	End Function
	
	'Adds a line to the response queue
	Public Sub Respond(ByRef strResponse As Object)
        'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Len(strResponse) > 0) Then m_response.Add(CStr(strResponse))
	End Sub
	
	'Cleares the response queue
	Public Sub ClearResponse()
		Do While m_response.Count()
			m_response.Remove(1)
		Loop 
	End Sub
	
	'Gets the Response Queue
	Public Function GetResponse() As Collection
		GetResponse = m_response
	End Function
	
	'This will respond in the proper style based on how the command was used.
	'This will messup emote responses /me if it is not public output or whisper commands is turned on.
	'If your response MUST be a specific style, Use AddQ or DSP
	Public Function SendResponse() As Object
		On Error GoTo ERROR_HANDLER
		Dim i As Short
		If (IsLocal) Then
			If (PublicOutput) Then
				For i = 1 To m_response.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object m_response.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.AddQ(m_response.Item(i), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
				Next i
			Else
				For i = 1 To m_response.Count()
					frmChat.AddChat(RTBColors.ConsoleText, m_response.Item(i))
				Next i
			End If
		Else
			If ((BotVars.WhisperCmds Or WasWhispered) And (PublicOutput = False)) Then
				For i = 1 To m_response.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object m_response.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.AddQ("/w " & Username & Space(1) & m_response.Item(i), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				Next i
			Else
				For i = 1 To m_response.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object m_response.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.AddQ(m_response.Item(i), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE))
				Next i
			End If
		End If
		ClearResponse()
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.Respond().")
	End Function
	
	Private Function CheckMatch(ByRef sExpression As String, ByRef sData As String, Optional ByRef CaseSensitive As Boolean = True) As Boolean
		On Error GoTo ERROR_HANDLER
		Dim mRegExp As New VBScript_RegExp_55.RegExp
		mRegExp.Global = True
		mRegExp.Pattern = sExpression
		mRegExp.IgnoreCase = (Not CaseSensitive)
		CheckMatch = mRegExp.Test(sData)
		
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.CheckMatch()")
	End Function
End Class