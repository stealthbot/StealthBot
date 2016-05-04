Option Strict Off
Option Explicit On
Module modOtherCode
	Private Const OBJECT_NAME As String = "modOtherCode"
	Private Declare Function GetEnvironmentVariable Lib "kernel32"  Alias "GetEnvironmentVariableA"(ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Private Declare Function SetEnvironmentVariable Lib "kernel32"  Alias "SetEnvironmentVariableA"(ByVal lpName As String, ByVal lpValue As String) As Integer
	Public Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal sBuffer As String, ByRef lSize As Integer) As Integer
	Public Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	Private Const MAX_COMPUTERNAME_LENGTH As Integer = 31
	Private Const MAX_USERNAME_LENGTH As Integer = 256
	
	Public Structure COMMAND_DATA
		Dim Name As String
		Dim params As String
		Dim local As Boolean
		Dim PublicOutput As Boolean
	End Structure
	
	Public Function GetComputerLanName() As String
		Dim buff As String
		Dim length As Integer
		buff = New String(Chr(0), MAX_COMPUTERNAME_LENGTH + 1)
		length = Len(buff)
		If (GetComputerName(buff, length)) Then
			GetComputerLanName = Left(buff, length)
		Else
			GetComputerLanName = vbNullString
		End If
	End Function
	
	Public Function GetComputerUsername() As String
		Dim buff As String
		Dim length As Integer
		buff = New String(Chr(0), MAX_USERNAME_LENGTH + 1)
		length = Len(buff)
		If (GetUserName(buff, length)) Then
			GetComputerUsername = KillNull(buff)
		Else
			GetComputerUsername = vbNullString
		End If
	End Function
	
	Public Function aton(ByRef sIPAddress As String) As Integer
		Dim sIP() As String
		Dim sValue As String
		sIP = Split(sIPAddress, ".")
		If (Not UBound(sIP) = 3) Then Exit Function
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sValue = StringFormat("{0}{1}{2}{3}", Chr(CInt(sIP(0))), Chr(CInt(sIP(1))), Chr(CInt(sIP(2))), Chr(CInt(sIP(3))))
		CopyMemory(aton, sValue, 4)
		'aton = Val(sIP(0)) + _
		''      (Val(sIP(1)) * &H100) + _
		''      (Val(sIP(2)) * &H10000) + _
		''      (Val(sIP(3)) * &H1000000)
	End Function
	
	'Read/WriteIni code thanks to ickis
	Public Sub WriteINI(ByVal wiSection As String, ByVal wiKey As String, ByVal wiValue As String, Optional ByVal wiFile As String = "x")
		If (StrComp(wiFile, "x", CompareMethod.Binary) = 0) Then
			wiFile = GetConfigFilePath()
		Else
			If (InStr(1, wiFile, "\", CompareMethod.Binary) = 0) Then
				wiFile = GetFilePath(wiFile)
			End If
		End If
		
		WritePrivateProfileString(wiSection, wiKey, wiValue, wiFile)
	End Sub
	
	Public Function ReadCfg(ByVal riSection As String, ByVal riKey As String) As String
		Dim sRiBuffer As String
		Dim sRiValue As String
		Dim sRiLong As String
		Dim riFile As String
		
		riFile = GetConfigFilePath()
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir(riFile) <> vbNullString) Then
			sRiBuffer = New String(Chr(VariantType.Null), 255)
			
			sRiLong = CStr(GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, riFile))
			
			If (Left(sRiBuffer, 1) <> Chr(1)) Then
				sRiValue = Left(sRiBuffer, CInt(sRiLong))
				
				ReadCfg = sRiValue
			End If
		Else
			ReadCfg = ""
		End If
	End Function
	
	Public Function ReadINI(ByVal riSection As String, ByVal riKey As String, ByVal riFile As String) As String
		Dim sRiBuffer As String
		Dim sRiValue As String
		Dim sRiLong As String
		
		If (InStr(1, riFile, "\", CompareMethod.Binary) = 0) Then
			riFile = GetFilePath(riFile)
		End If
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir(riFile) <> vbNullString) Then
			sRiBuffer = New String(Chr(VariantType.Null), 255)
			
			sRiLong = CStr(GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, riFile))
			
			If (Left(sRiBuffer, 1) <> Chr(1)) Then
				sRiValue = Left(sRiBuffer, CInt(sRiLong))
				ReadINI = sRiValue
			End If
		Else
			ReadINI = ""
		End If
	End Function
	
	'// http://www.vbforums.com/showpost.php?p=2909245&postcount=3
	Public Sub BubbleSort1(ByRef pvarArray As Object)
		Dim i As Integer
		Dim iMin As Integer
		Dim iMax As Integer
		Dim varSwap As Object
		Dim blnSwapped As Boolean
		
		iMin = LBound(pvarArray)
		iMax = UBound(pvarArray) - 1
		
		Do 
			blnSwapped = False
			For i = iMin To iMax
				'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If pvarArray(i) > pvarArray(i + 1) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varSwap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varSwap = pvarArray(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pvarArray(i) = pvarArray(i + 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object varSwap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pvarArray(i + 1) = varSwap
					blnSwapped = True
				End If
			Next 
			iMax = iMax - 1
		Loop Until Not blnSwapped
		
	End Sub
	
	
	Public Function GetTimeStamp(Optional ByRef DateTime As Date = #12:00:00 AM#) As String
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, DateTime, CDate("00:00:00 12/30/1899")) = 0) Then
			DateTime = Now
		End If
		
		Select Case (BotVars.TSSetting)
			Case 0
				GetTimeStamp = " [" & VB6.Format(DateTime, "HH:MM:SS AM/PM") & "] "
				
			Case 1
				GetTimeStamp = " [" & VB6.Format(DateTime, "HH:MM:SS") & "] "
				
			Case 2
				GetTimeStamp = " [" & VB6.Format(DateTime, "HH:MM:SS") & "." & Right("000" & GetCurrentMS, 3) & "] "
				
			Case 3
				GetTimeStamp = vbNullString
				
			Case Else
				GetTimeStamp = " [" & VB6.Format(DateTime, "HH:MM:SS AM/PM") & "] "
		End Select
	End Function
	
	'// Converts a millisecond or second time value to humanspeak.. modified to support BNet's Time
	'// Logged report.
	'// Updated 2/11/05 to support timeGetSystemTime() unsigned longs, which in VB are doubles after conversion
	Public Function ConvertTime(ByVal dblMS As Double, Optional ByRef seconds As Byte = 0) As String
		Dim dblSeconds As Double
		Dim dblDays As Double
		Dim dblHours As Double
		Dim dblMins As Double
		Dim strSeconds As String
		Dim strDays As String
		
		If (seconds = 0) Then
			dblSeconds = (dblMS / 1000)
		Else
			dblSeconds = dblMS
		End If
		
		dblDays = Int(dblSeconds / 86400)
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dblSeconds = dblSeconds Mod 86400
		dblHours = Int(dblSeconds / 3600)
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dblSeconds = dblSeconds Mod 3600
		dblMins = Int(dblSeconds / 60)
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dblSeconds = (dblSeconds Mod 60)
		
		If (dblSeconds <> 1) Then
			strSeconds = "s"
		End If
		
		If (dblDays <> 1) Then
			strDays = "s"
		End If
		
		ConvertTime = dblDays & " day" & strDays & ", " & dblHours & " hours, " & dblMins & " minutes and " & dblSeconds & " second" & strSeconds
	End Function
	
	Public Function GetVerByte(ByRef Product As String, Optional ByVal UseHardcode As Short = 0) As Integer
		Dim Key As String
		
		Key = GetProductKey(Product)
		
		If ((Config.GetVersionByte(Key) = -1) Or (UseHardcode = 1)) Then
			
			GetVerByte = GetProductInfo(Product).VersionByte
		Else
			GetVerByte = Config.GetVersionByte(Key)
		End If
		
	End Function
	
	Public Function GetGamePath(ByVal Client As String) As String
		' [override] XXHashes= functionality replaced by checkrevision.ini -> [CRev_XX] Path=
		' removed ~Ribose/2010-08-12
		Dim CRevINIPath As String
		Dim Key As String
		Dim Path As String
		Dim sep1 As String
		Dim sep2 As String
		
		' Moved CheckRevision.ini to profile directory instead of install directory. -Pyro, 2016-03-21
		'CRevINIPath = GetFilePath(FILE_CREV_INI, StringFormat("{0}\", App.Path))
		CRevINIPath = GetFilePath(FILE_CREV_INI)
		
		Key = GetProductKey(Client)
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Path = ReadINI(StringFormat("CRev_{0}", Key), "Path", CRevINIPath)
		sep1 = vbNullString
		sep2 = vbNullString
		
		If (InStr(1, Path, ":\") = 0) Then
			If (Left(Path, 1) <> "\") Then sep1 = "\"
			If (Right(Path, 1) <> "\") Then sep2 = "\"
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Path = StringFormat("{0}{1}{2}{3}", My.Application.Info.DirectoryPath, sep1, Path, sep2)
		End If
		
		GetGamePath = Path
	End Function
	
	Function MKL(ByRef Value As Integer) As String
		Dim Result As New VB6.FixedLengthString(4)
		
		Call CopyMemory(Result.Value, Value, 4)
		
		MKL = Result.Value
	End Function
	
	Function MKI(ByRef Value As Short) As String
		Dim Result As New VB6.FixedLengthString(2)
		
		Call CopyMemory(Result.Value, Value, 2)
		
		MKI = Result.Value
	End Function
	
	Public Function CheckPath(ByVal sPath As String) As Integer
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(sPath)) = 0) Then
            frmChat.AddChat(RTBColors.ErrorMessageText, "[HASHES] " & Mid(sPath, InStrRev(sPath, "\") + 1) & " is missing.")

            CheckPath = 1
        End If
	End Function
	
	Public Function Ban(ByVal Inpt As String, ByRef SpeakerAccess As Short, Optional ByRef Kick As Short = 0) As String
		On Error GoTo ERROR_HANDLER
		
		Static LastBan As String
		
		Dim Username As String
		Dim CleanedUsername As String
		Dim i As Short
		Dim pos As Short
		
        If (Len(Inpt) > 0) Then
            If (Kick > 2) Then
                LastBan = vbNullString

                Exit Function
            End If

            If (g_Channel.Self.IsOperator) Then
                If (InStr(1, Inpt, Space(1), CompareMethod.Binary) <> 0) Then
                    Username = LCase(Left(Inpt, InStr(1, Inpt, Space(1), CompareMethod.Binary) - 1))
                Else
                    Username = LCase(Inpt)
                End If

                If (Len(Username) > 0) Then
                    LastBan = LCase(Username)

                    'CleanedUsername = StripRealm(CleanedUsername)
                    CleanedUsername = StripInvalidNameChars(Username)

                    If (SpeakerAccess < 200) Then
                        If ((GetSafelist(CleanedUsername)) Or (GetSafelist(Username))) Then
                            Ban = "Error: That user is safelisted."

                            Exit Function
                        End If
                    End If

                    If (GetCumulativeAccess(Username).Rank >= SpeakerAccess) Then
                        Ban = "Error: You do not have sufficient access to do that."

                        Exit Function
                    End If

                    pos = g_Channel.GetUserIndex(Username)

                    If (pos > 0) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(pos).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If (g_Channel.Users.Item(pos).IsOperator) Then
                            Ban = "Error: You cannot ban a channel operator."

                            Exit Function
                        End If
                    End If

                    If (Kick = 0) Then
                        Call frmChat.AddQ("/ban " & Inpt)
                    Else
                        Call frmChat.AddQ("/kick " & Inpt)
                    End If
                End If
            Else
                Ban = "The bot does not have ops."
            End If
        End If
		
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Ban().")
		
		Exit Function
	End Function
	
	' This function created in response to http://www.stealthbot.net/forum/index.php?showtopic=20550
	Public Function StripInvalidNameChars(ByVal Username As String) As String
		Dim Allowed(14) As Short
		Dim i As Short
		Dim j As Short
		Dim thisChar As Short
		Dim NewUsername As String
		Dim ThisCharOK As Boolean
		
        If (Len(Username) > 0) Then
            NewUsername = Username

            Allowed(0) = Asc("`")
            Allowed(1) = Asc("[")
            Allowed(2) = Asc("]")
            Allowed(3) = Asc("{")
            Allowed(4) = Asc("}")
            Allowed(5) = Asc("_")
            Allowed(6) = Asc("-")
            Allowed(7) = Asc("@")
            Allowed(8) = Asc("^")
            Allowed(9) = Asc(".")
            Allowed(10) = Asc("+")
            Allowed(11) = Asc("=")
            Allowed(12) = Asc("~")
            Allowed(13) = Asc("|")
            Allowed(14) = Asc("*")

            For i = 1 To Len(Username)
                thisChar = Asc(Mid(Username, i, 1))

                ThisCharOK = False

                If (Not (IsAlpha(thisChar))) Then
                    If (Not (IsNumber(thisChar))) Then
                        For j = 0 To UBound(Allowed)
                            If (thisChar = Allowed(j)) Then
                                ThisCharOK = True
                            End If
                        Next j

                        If (Not (ThisCharOK)) Then
                            NewUsername = Replace(NewUsername, Chr(thisChar), vbNullString)
                        End If
                    End If
                End If
            Next i

            StripInvalidNameChars = NewUsername
        End If
	End Function
	
	'// Utility Function for joining strings
	'// EXAMPLE
	'// StringFormat("This is an {1} of its {0}.", Array("use", "example")) '// OUTPUT: This is an example of its use.
	'// 08/29/2008 JSM - Created
	Public Function StringFormatA(ByRef source As String, ByRef params() As Object) As String
		
		On Error GoTo ERROR_HANDLER
		
		Dim retVal As String
		Dim i As Short
		retVal = source
		For i = LBound(params) To UBound(params)
			'UPGRADE_WARNING: Couldn't resolve default property of object params(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			retVal = Replace(retVal, "{" & i & "}", CStr(params(i)))
		Next 
		StringFormatA = retVal
		
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in StringFormatA().")
		
		StringFormatA = vbNullString
		
	End Function
	
	
	'// Utility Function for joining strings
	'// EXAMPLE
	'// StringFormat("This is an {1} of its {0}.", "use", "example") '// OUTPUT: This is an example of its use.
	'// 08/14/2009 JSM - Created
	'UPGRADE_WARNING: ParamArray params was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function StringFormat(ByRef source As String, ParamArray ByVal params() As Object) As Object
		
		On Error GoTo ERROR_HANDLER
		
		Dim retVal As String
		Dim i As Short
		retVal = source
		For i = LBound(params) To UBound(params)
			'UPGRADE_WARNING: Couldn't resolve default property of object params(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			retVal = Replace(retVal, "{" & i & "}", CStr(params(i)))
		Next 
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StringFormat = retVal
		
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in StringFormat().")
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StringFormat = vbNullString
		
	End Function
	
	Public Function StripAccountNumber(ByVal Username As String) As String
		Dim numpos As Short
		Dim atpos As Short
		
		numpos = InStr(1, Username, "#", CompareMethod.Binary)
		If numpos > 0 Then
			atpos = InStr(numpos, Username, "@", CompareMethod.Binary)
			If atpos > 0 Then
				StripAccountNumber = Left(Username, numpos - 1) & Mid(Username, atpos)
			Else
				StripAccountNumber = Left(Username, numpos - 1)
			End If
		Else
			StripAccountNumber = Username
		End If
	End Function
	
	Public Function StripRealm(ByVal Username As String) As String
		If (InStr(1, Username, "@", CompareMethod.Binary) > 0) Then
			Username = Replace(Username, "@USWest", vbNullString, 1)
			Username = Replace(Username, "@USEast", vbNullString, 1)
			Username = Replace(Username, "@Asia", vbNullString, 1)
			Username = Replace(Username, "@Euruope", vbNullString, 1)
			Username = Replace(Username, "@Beta", vbNullString, 1)
			
			Username = Replace(Username, "@Lordaeron", vbNullString, 1)
			Username = Replace(Username, "@Azeroth", vbNullString, 1)
			Username = Replace(Username, "@Kalimdor", vbNullString, 1)
			Username = Replace(Username, "@Northrend", vbNullString, 1)
			Username = Replace(Username, "@Westfall", vbNullString, 1)
			
			Username = Replace(Username, "@Blizzard", vbNullString, 1)
		End If
		
		StripRealm = Username
	End Function
	
	Public Sub bnetSend(ByVal Message As String, Optional ByVal Tag As String = vbNullString, Optional ByVal ID As Double = 0)
		
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNet.CtlState = 7) Then
			With PBuffer
				If (frmChat.mnuUTF8.Checked) Then
					.InsertNTString(Message, modPacketBuffer.STRINGENCODING.UTF8)
				Else
					.InsertNTString(Message)
				End If
				
				.SendPacket(SID_CHATCOMMAND)
			End With
			
			If (Tag = "request_receipt") Then
				g_request_receipt = True
				
				With PBuffer
					.SendPacket(SID_FRIENDSLIST)
				End With
			End If
		End If
		
		If (Left(Message, 1) <> "/") Then
			If (g_Channel.IsSilent) Then
				frmChat.AddChat(RTBColors.Carats, "<", RTBColors.TalkBotUsername, GetCurrentUsername, RTBColors.Carats, "> ", RTBColors.WhisperText, Message)
			Else
				frmChat.AddChat(RTBColors.Carats, "<", RTBColors.TalkBotUsername, GetCurrentUsername, RTBColors.Carats, "> ", RTBColors.TalkNormalText, Message)
			End If
		End If
		
		On Error Resume Next
		
		RunInAll("Event_MessageSent", ID, Message, Tag)
		
		Exit Sub
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in bnetSend().")
		
		Exit Sub
		
	End Sub
	
	Public Function Voting(ByVal Mode1 As Byte, Optional ByRef Mode2 As Byte = 0, Optional ByRef Username As String = "") As String
		On Error GoTo ERROR_HANDLER
		Static Voted() As String
		Static VotesYes As Short
		Static VotesNo As Short
		Static VoteMode As Byte
		Static Target As String
		
		Dim i As Short
		
		Select Case (Mode1)
			Case BVT_VOTE_ADD
				For i = LBound(Voted) To UBound(Voted)
					If (StrComp(Voted(i), LCase(Username), CompareMethod.Text) = 0) Then
						Exit Function
					End If
				Next i
				
				Select Case (Mode2)
					Case BVT_VOTE_ADDYES
						VotesYes = VotesYes + 1
					Case BVT_VOTE_ADDNO
						VotesNo = VotesNo + 1
				End Select
				
				Voted(UBound(Voted)) = LCase(Username)
				
				ReDim Preserve Voted(UBound(Voted) + 1)
				
			Case BVT_VOTE_START
				VotesYes = 0
				VotesNo = 0
				
				ReDim Voted(0)
				
				VoteMode = Mode2
				Target = Username
				Voting = "Vote started. Type YES or NO to vote. Your vote will " & "only be counted once."
				
			Case BVT_VOTE_END
				If (Mode2 = BVT_VOTE_CANCEL) Then
					Voting = "Vote cancelled. Final results: [" & VotesYes & "] YES, " & "[" & VotesNo & "] NO. "
				Else
					Select Case (VoteMode)
						Case BVT_VOTE_STD
							Voting = "Final results: [" & VotesYes & "] YES, [" & VotesNo & "] NO. "
							
							If VotesYes > VotesNo Then
								Voting = Voting & "YES wins, with " & VB6.Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
							ElseIf (VotesYes < VotesNo) Then 
								Voting = Voting & "NO wins, with " & VB6.Format(VotesNo / (VotesYes + VotesNo), "percent") & " of the vote."
							Else
								Voting = Voting & "The vote was a draw."
							End If
							
						Case BVT_VOTE_BAN
							If (VotesYes > VotesNo) Then
								Voting = Ban(Target & " Banned by vote", VoteInitiator.Rank)
							Else
								Voting = "Ban vote failed."
							End If
							
						Case BVT_VOTE_KICK
							If (VotesYes > VotesNo) Then
								Voting = Ban(Target & " Kicked by vote", VoteInitiator.Rank, 1)
							Else
								Voting = "Kick vote failed."
							End If
					End Select
				End If
				
				VoteDuration = -1
				VotesYes = 0
				VotesNo = 0
				VoteMode = 0
				Target = vbNullString
				
				ReDim Voted(0)
				
			Case BVT_VOTE_TALLY
				Voting = "Current results: [" & VotesYes & "] YES, [" & VotesNo & "] NO; " & VoteDuration & " seconds remain."
				
				If (VotesYes > VotesNo) Then
					Voting = Voting & " YES leads, with " & VB6.Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
				ElseIf (VotesYes < VotesNo) Then 
					Voting = Voting & " NO leads, with " & VB6.Format(VotesNo / (VotesYes + VotesNo), "percent") & " of the vote."
				Else
					Voting = Voting & " The vote is a draw."
				End If
		End Select
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Voting()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Function GetAccess(ByVal Username As String, Optional ByRef dbType As String = vbNullString) As udtGetAccessResponse
		
		Dim i As Short
		Dim bln As Boolean
		
		'If (Left$(Username, 1) = "*") Then
		'    Username = Mid$(Username, 2)
		'End If
		
		For i = LBound(DB) To UBound(DB)
			If (StrComp(DB(i).Username, Username, CompareMethod.Text) = 0) Then
				If (Len(dbType)) Then
					If (StrComp(DB(i).Type, dbType, CompareMethod.Binary) = 0) Then
						bln = True
					End If
				Else
					bln = True
				End If
				
				If (bln = True) Then
					With GetAccess
						.Username = DB(i).Username
						.Rank = DB(i).Rank
						.Flags = DB(i).Flags
						.AddedBy = DB(i).AddedBy
						.AddedOn = DB(i).AddedOn
						.ModifiedBy = DB(i).ModifiedBy
						.ModifiedOn = DB(i).ModifiedOn
						.Type = DB(i).Type
						.Groups = DB(i).Groups
						.BanMessage = DB(i).BanMessage
					End With
					
					Exit Function
				End If
			End If
			
			bln = False
		Next i
		
		GetAccess.Rank = -1
	End Function
	
	Public Function dbLastModified() As Date
		
		Dim temp As Date
		Dim i As Short
		
		temp = CDate("00:00:00 12/30/1899")
		
		For i = LBound(DB) To UBound(DB)
			If (DB(i).Username = vbNullString) Then
				Exit For
			End If
			
			'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
			If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, temp, DB(i).ModifiedOn) > 0) Then
				temp = DB(i).ModifiedOn
			End If
		Next i
		
		dbLastModified = temp
		
	End Function
	
	Public Function GetCumulativeAccess(ByVal Username As String, Optional ByRef dbType As String = vbNullString) As udtGetAccessResponse
		
		On Error GoTo ERROR_HANDLER
		
		Static dynGroups() As udtDatabase
		Static dModified As Date
		
		Dim gAcc As udtGetAccessResponse
		
		Dim f As Scripting.File
		Dim fso As Scripting.FileSystemObject
		Dim i As Short
		Dim K As Short
		Dim j As Short
		Dim found As Boolean
		Dim dbIndex As Short
		Dim dbCount As Short
		Dim Splt() As String
		Dim bln As Boolean
		Dim modified As FILETIME
		Dim creation As FILETIME
		Dim Access As FILETIME
		Dim nModified As Date
		
		' default index to negative one to
		' indicate that no matching users have
		' been found
		dbIndex = -1
		
		fso = New Scripting.FileSystemObject
		
		' If for some reason the users file doesn't exist, create it. ' 10/13/08 ~Pyro
		If (Not (fso.FileExists("./users.txt"))) Then
			Call fso.CreateTextFile("./users.txt", False)
		End If
		
		'Set f = fso.GetFile("./users.txt")
		
		'nModified = f.DateLastModified
		'nModified = dbLastModified
		
		'frmChat.AddChat vbRed, DateDiff("s", dModified, dbLastModified)
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, dModified, dbLastModified) > 0) Then
			ReDim dynGroups(0)
			
			With dynGroups(0)
				.Username = vbNullString
			End With
			
			For i = LBound(DB) To UBound(DB)
				If ((InStr(1, DB(i).Username, "*", CompareMethod.Binary) <> 0) Or (InStr(1, DB(i).Username, "?", CompareMethod.Binary) <> 0) Or (DB(i).Type = "GAME") Or (DB(i).Type = "CLAN")) Then
					
					If (dynGroups(0).Username <> vbNullString) Then
						ReDim Preserve dynGroups(UBound(dynGroups) + 1)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object dynGroups(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dynGroups(UBound(dynGroups)) = DB(i)
				End If
			Next i
			
			dModified = nModified
		End If
		
		'Set fso = Nothing
		
		Dim doCheck As Boolean
		Dim tmp As udtDatabase
		If (DB(LBound(DB)).Username <> vbNullString) Then
			For i = LBound(DB) To UBound(DB)
				If (StrComp(Username, DB(i).Username, CompareMethod.Text) = 0) Then
					If ((dbType = vbNullString) Or (dbType <> vbNullString) And (StrComp(dbType, DB(i).Type, CompareMethod.Text) = 0)) Then
						
						With GetCumulativeAccess
							.Username = DB(i).Username & IIf(((DB(i).Type <> "%") And (StrComp(DB(i).Type, "USER", CompareMethod.Text) <> 0)), " (" & LCase(DB(i).Type) & ")", vbNullString)
							.Rank = DB(i).Rank
							.Flags = DB(i).Flags
							.AddedBy = DB(i).AddedBy
							.AddedOn = DB(i).AddedOn
							.ModifiedBy = DB(i).ModifiedBy
							.ModifiedOn = DB(i).ModifiedOn
							.Type = IIf(((DB(i).Type <> "%") And (DB(i).Type <> vbNullString)), DB(i).Type, "USER")
							.Groups = DB(i).Groups
							.BanMessage = DB(i).BanMessage
						End With
						
						If ((Len(DB(i).Groups) > 0) And (DB(i).Groups <> "%")) Then
							If (InStr(1, DB(i).Groups, ",", CompareMethod.Binary) <> 0) Then
								Splt = Split(DB(i).Groups, ",")
							Else
								ReDim Preserve Splt(0)
								
								Splt(0) = DB(i).Groups
							End If
							
							For j = 0 To UBound(Splt)
								'UPGRADE_WARNING: Couldn't resolve default property of object gAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								gAcc = GetCumulativeGroupAccess(Splt(j))
								
								If (GetCumulativeAccess.Rank < gAcc.Rank) Then
									GetCumulativeAccess.Rank = gAcc.Rank
									
									bln = True
								End If
								
								For K = 1 To Len(gAcc.Flags)
									If (InStr(1, GetCumulativeAccess.Flags, Mid(gAcc.Flags, K, 1), CompareMethod.Binary) = 0) Then
										
										GetCumulativeAccess.Flags = GetCumulativeAccess.Flags & Mid(gAcc.Flags, K, 1)
										
										bln = True
									End If
								Next K
								
								If ((GetCumulativeAccess.BanMessage = vbNullString) Or (GetCumulativeAccess.BanMessage = "%")) Then
									
									GetCumulativeAccess.BanMessage = gAcc.BanMessage
									
									bln = True
								End If
								
								If (bln) Then
									If (dbCount = 0) Then
										GetCumulativeAccess.Username = GetCumulativeAccess.Username & IIf((i + 1), Space(1), vbNullString) & "["
									End If
									
									GetCumulativeAccess.Username = GetCumulativeAccess.Username & gAcc.Username & IIf(((gAcc.Type <> "%") And (StrComp(gAcc.Type, "USER", CompareMethod.Text) <> 0)), " (" & LCase(gAcc.Type) & ")", vbNullString) & ", "
									
									dbCount = (dbCount + 1)
								End If
								
								bln = False
							Next j
						End If
						
						dbIndex = i
						
						Exit For
					End If
				End If
			Next i
			
			If (InStr(1, GetCumulativeAccess.Flags, "I", CompareMethod.Binary) = 0) Then
				If ((InStr(1, Username, "*", CompareMethod.Binary) = 0) And (InStr(1, Username, "?", CompareMethod.Binary) = 0) And (GetCumulativeAccess.Type <> "GAME") And (GetCumulativeAccess.Type <> "CLAN") And (GetCumulativeAccess.Type <> "GROUP")) Then
					
					For i = LBound(dynGroups) To UBound(dynGroups)
						
						If (i <> dbIndex) Then
							' default type to user
							dynGroups(i).Type = IIf(((dynGroups(i).Type <> "%") And (dynGroups(i).Type <> vbNullString)), dynGroups(i).Type, "USER")
							
							If (StrComp(dynGroups(i).Type, "USER", CompareMethod.Text) = 0) Then
								If ((LCase(PrepareCheck(Username))) Like (LCase(PrepareCheck(dynGroups(i).Username)))) Then
									
									doCheck = True
								End If
							ElseIf (StrComp(dynGroups(i).Type, "GAME", CompareMethod.Text) = 0) Then 
								For j = 1 To g_Channel.Users.Count()
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (StrComp(Username, g_Channel.Users.Item(j).DisplayName, CompareMethod.Text) = 0) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(j).Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (StrComp(dynGroups(i).Username, g_Channel.Users.Item(j).Game, CompareMethod.Text) = 0) Then
											doCheck = True
										End If
										
										Exit For
									End If
								Next j
							ElseIf (StrComp(dynGroups(i).Type, "CLAN", CompareMethod.Text) = 0) Then 
								For j = 1 To g_Channel.Users.Count()
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (StrComp(Username, g_Channel.Users.Item(j).DisplayName, CompareMethod.Text) = 0) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(j).Clan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (StrComp(dynGroups(i).Username, g_Channel.Users.Item(j).Clan, CompareMethod.Text) = 0) Then
											doCheck = True
										End If
										
										Exit For
									End If
								Next j
							End If
							
							If (doCheck = True) Then
								
								'UPGRADE_WARNING: Couldn't resolve default property of object tmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								tmp = dynGroups(i)
								
								If ((Len(tmp.Groups) > 0) And (tmp.Groups <> "%")) Then
									If (InStr(1, tmp.Groups, ",", CompareMethod.Binary) <> 0) Then
										Splt = Split(tmp.Groups, ",")
									Else
										ReDim Preserve Splt(0)
										
										Splt(0) = tmp.Groups
									End If
									
									For j = 0 To UBound(Splt)
										'UPGRADE_WARNING: Couldn't resolve default property of object gAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										gAcc = GetCumulativeGroupAccess(Splt(j))
										
										If (tmp.Rank < gAcc.Rank) Then
											tmp.Rank = gAcc.Rank
										End If
										
										For K = 1 To Len(gAcc.Flags)
											If (InStr(1, tmp.Flags, Mid(gAcc.Flags, K, 1), CompareMethod.Binary) = 0) Then
												
												tmp.Flags = tmp.Flags & Mid(gAcc.Flags, K, 1)
											End If
										Next K
										
										If ((tmp.BanMessage = vbNullString) Or (tmp.BanMessage = "%")) Then
											
											tmp.BanMessage = gAcc.BanMessage
										End If
									Next j
								End If
								
								If (GetCumulativeAccess.Rank < tmp.Rank) Then
									GetCumulativeAccess.Rank = tmp.Rank
									
									bln = True
								End If
								
								For j = 1 To Len(tmp.Flags)
									If (InStr(1, GetCumulativeAccess.Flags, Mid(tmp.Flags, j, 1), CompareMethod.Binary) = 0) Then
										
										GetCumulativeAccess.Flags = GetCumulativeAccess.Flags & Mid(tmp.Flags, j, 1)
										
										bln = True
									End If
								Next j
								
								If ((GetCumulativeAccess.BanMessage = vbNullString) Or (GetCumulativeAccess.BanMessage = "%")) Then
									
									GetCumulativeAccess.BanMessage = tmp.BanMessage
									
									bln = True
								End If
								
								If (bln) Then
									If (dbCount = 0) Then
										GetCumulativeAccess.Username = GetCumulativeAccess.Username & IIf((dbIndex + 1), Space(1), vbNullString) & "["
									End If
									
									GetCumulativeAccess.Username = GetCumulativeAccess.Username & tmp.Username & IIf(((tmp.Type <> "%") And (StrComp(tmp.Type, "USER", CompareMethod.Text) <> 0)), " (" & LCase(tmp.Type) & ")", vbNullString) & ", "
									
									dbCount = (dbCount + 1)
								End If
							End If
						End If
						
						bln = False
						doCheck = False
					Next i
				End If
			End If
			
			If (dbCount = 0) Then
				If (dbIndex = -1) Then
					With GetCumulativeAccess
						.Username = vbNullString
						.Rank = 0
						.Flags = vbNullString
					End With
				End If
			Else
				GetCumulativeAccess.Username = Left(GetCumulativeAccess.Username, Len(GetCumulativeAccess.Username) - 2) & "]"
			End If
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		'Ignores error 28: "Out of stack memory"
		If Err.Number <> 28 Then
			Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in " & "GetCumulativeAccess().")
		End If
		
		Exit Function
	End Function
	
	Private Function GetCumulativeGroupAccess(ByVal Group As String) As udtGetAccessResponse
		Dim gAcc As udtGetAccessResponse
		Dim Splt() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gAcc = GetAccess(Group, "GROUP")
		
		Dim recAcc As udtGetAccessResponse
		Dim i As Short
		Dim j As Short
		If ((Len(gAcc.Groups) > 0) And (gAcc.Groups <> "%")) Then
			
			If (InStr(1, gAcc.Groups, ",", CompareMethod.Binary) <> 0) Then
				
				Splt = Split(gAcc.Groups, ",")
				
				For i = 0 To UBound(Splt)
					'UPGRADE_WARNING: Couldn't resolve default property of object recAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					recAcc = GetCumulativeGroupAccess(Splt(i))
					
					If (gAcc.Rank < recAcc.Rank) Then
						gAcc.Rank = recAcc.Rank
					End If
					
					For j = 1 To Len(recAcc.Flags)
						If (InStr(1, gAcc.Flags, Mid(recAcc.Flags, j, 1), CompareMethod.Binary) = 0) Then
							
							gAcc.Flags = gAcc.Flags & Mid(recAcc.Flags, j, 1)
						End If
					Next j
					
					If ((gAcc.BanMessage = vbNullString) Or (gAcc.BanMessage = "%")) Then
						
						gAcc.BanMessage = recAcc.BanMessage
					End If
				Next i
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object recAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				recAcc = GetCumulativeGroupAccess(gAcc.Groups)
				
				If (gAcc.Rank < recAcc.Rank) Then
					gAcc.Rank = recAcc.Rank
				End If
				
				For j = 1 To Len(recAcc.Flags)
					If (InStr(1, gAcc.Flags, Mid(recAcc.Flags, j, 1), CompareMethod.Binary) = 0) Then
						
						gAcc.Flags = gAcc.Flags & Mid(recAcc.Flags, j, 1)
					End If
				Next j
				
				If ((gAcc.BanMessage = vbNullString) Or (gAcc.BanMessage = "%")) Then
					
					gAcc.BanMessage = recAcc.BanMessage
				End If
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCumulativeGroupAccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCumulativeGroupAccess = gAcc
	End Function
	
	Public Function CheckGroup(ByVal Group As String, ByVal Check As String) As Boolean
		Dim gAcc As udtGetAccessResponse
		Dim Splt() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gAcc = GetAccess(Group, "GROUP")
		
		Dim recAcc As Boolean
		Dim i As Short
		Dim j As Short
		If ((Len(gAcc.Groups) > 0) And (gAcc.Groups <> "%")) Then
			
			If (InStr(1, gAcc.Groups, ",", CompareMethod.Binary) <> 0) Then
				
				Splt = Split(gAcc.Groups, ",")
				
				For i = 0 To UBound(Splt)
					If (StrComp(Splt(i), Check, CompareMethod.Text) = 0) Then
						CheckGroup = True
						
						Exit Function
					Else
						recAcc = CheckGroup(Splt(i), Check)
						
						If (recAcc) Then
							CheckGroup = True
							
							Exit Function
						End If
					End If
				Next i
			Else
				If (StrComp(gAcc.Groups, Check, CompareMethod.Text) = 0) Then
					CheckGroup = True
					
					Exit Function
				Else
					recAcc = CheckGroup(gAcc.Groups, Check)
					
					If (recAcc) Then
						CheckGroup = True
						
						Exit Function
					End If
				End If
			End If
		End If
		
		CheckGroup = False
	End Function
	
	Public Sub RequestSystemKeys()
		AwaitingSystemKeys = 1
		
		With PBuffer
			.InsertDWord(&H1)
			.InsertDWord(&H4)
			.InsertDWord(GetTickCount())
			.InsertNTString((BotVars.Username))
			
			.InsertNTString("System\Account Created")
			.InsertNTString("System\Last Logon")
			.InsertNTString("System\Last Logoff")
			.InsertNTString("System\Time Logged")
			
			.SendPacket(SID_READUSERDATA)
		End With
	End Sub
	
	'// parses a system time and returns in the format:
	'//     mm/dd/yy, hh:mm:ss
	'//
	Public Function SystemTimeToString(ByRef st As SYSTEMTIME) As String
		Dim buf As String
		
		With st
			buf = buf & .wMonth & "/"
			buf = buf & .wDay & "/"
			buf = buf & .wYear & ", "
			buf = buf & IIf(.wHour > 9, .wHour, "0" & .wHour) & ":"
			buf = buf & IIf(.wMinute > 9, .wMinute, "0" & .wMinute) & ":"
			buf = buf & IIf(.wSecond > 9, .wSecond, "0" & .wSecond)
		End With
		
		SystemTimeToString = buf
	End Function
	
	Public Function GetCurrentMS() As String
		Dim st As SYSTEMTIME
		GetLocalTime(st)
		
		GetCurrentMS = Right("000" & st.wMilliseconds, 3)
	End Function
	
	Public Function ZeroOffset(ByVal lInpt As Integer, ByVal lDigits As Integer) As String
		Dim sOut As String
		
		sOut = Hex(lInpt)
		ZeroOffset = Right(New String("0", lDigits) & sOut, lDigits)
	End Function
	
	Public Function ZeroOffsetEx(ByVal lInpt As Integer, ByVal lDigits As Integer) As String
		ZeroOffsetEx = Right(New String("0", lDigits) & lInpt, lDigits)
	End Function
	
	Public Function GetSmallIcon(ByVal sProduct As String, ByVal Flags As Integer, ByRef IconCode As Short) As Integer
		Dim i As Integer
		
		If (BotVars.ShowFlagsIcons = False) Then
			i = IconCode ' disable any of the below flags-based icons
		ElseIf (Flags And USER_BLIZZREP) = USER_BLIZZREP Then  'Flags = 1: blizzard rep
			i = ICBLIZZ
		ElseIf (Flags And USER_SYSOP) = USER_SYSOP Then  'Flags = 8: battle.net sysop
			i = ICSYSOP
		ElseIf (Flags And USER_CHANNELOP) = USER_CHANNELOP Then  'op
			i = ICGAVEL
		ElseIf (Flags And USER_GUEST) = USER_GUEST Then  'guest
			i = ICSPECS
		ElseIf (Flags And USER_SPEAKER) = USER_SPEAKER Then  'speaker
			i = ICSPEAKER
		ElseIf (Flags And USER_GFPLAYER) = USER_GFPLAYER Then  'GF player
			i = IC_GF_PLAYER
		ElseIf (Flags And USER_GFOFFICIAL) = USER_GFOFFICIAL Then  'GF official
			i = IC_GF_OFFICIAL
		ElseIf (Flags And USER_SQUELCHED) = USER_SQUELCHED Then  'squelched
			i = ICSQUELCH
		Else
			i = IconCode
			'Else
			'    Select Case (UCase$(sProduct))
			'        Case Is =  PRODUCT_STAR: I = ICSTAR
			'        Case Is = PRODUCT_SEXP: I = ICSEXP
			'        Case Is = PRODUCT_D2DV: I = ICD2DV
			'        Case Is = PRODUCT_D2XP: I = ICD2XP
			'        Case Is = PRODUCT_W2BN: I = ICW2BN
			'        Case Is = PRODUCT_CHAT: I = ICCHAT
			'        Case Is = PRODUCT_DRTL: I = ICDIABLO
			'        Case Is = PRODUCT_DSHR: I = ICDIABLOSW
			'        Case Is = PRODUCT_JSTR: I = ICJSTR
			'        Case Is = PRODUCT_SSHR: I = ICSCSW
			'        Case Is = PRODUCT_WAR3: I = ICWAR3
			'        Case Is = PRODUCT_W3XP: I = ICWAR3X
			'
			'        '*** Special icons for WCG added 6/24/07 ***
			'        Case Is = "WCRF": I = IC_WCRF
			'        Case Is = "WCPL": I = IC_WCPL
			'        Case Is = "WCGO": I = IC_WCGO
			'        Case Is = "WCSI": I = IC_WCSI
			'        Case Is = "WCBR": I = IC_WCBR
			'        Case Is = "WCPG": I = IC_WCPG
			'
			'        '*** Special icons for PGTour ***
			'        Case Is = "__A+": I = IC_PGT_A + 1
			'        Case Is = "___A": I = IC_PGT_A
			'        Case Is = "__A-": I = IC_PGT_A - 1
			'        Case Is = "__B+": I = IC_PGT_B + 1
			'        Case Is = "___B": I = IC_PGT_B
			'        Case Is = "__B-": I = IC_PGT_B - 1
			'        Case Is = "__C+": I = IC_PGT_C + 1
			'        Case Is = "___C": I = IC_PGT_C
			'        Case Is = "__C-": I = IC_PGT_C - 1
			'        Case Is = "__D+": I = IC_PGT_D + 1
			'        Case Is = "___D": I = IC_PGT_D
			'        Case Is = "__D-": I = IC_PGT_D - 1
			'
			'        Case Else: I = ICUNKNOWN
			'    End Select
		End If
		
		GetSmallIcon = i
	End Function
	
	Public Sub AddName(ByVal Username As String, ByVal AccountName As String, ByVal Product As String, ByVal Flags As Integer, ByVal Ping As Integer, ByRef IconCode As Short, Optional ByRef Clan As String = "", Optional ByRef ForcePosition As Short = 0)
		Dim i As Short
		Dim LagIcon As Short
		Dim isPriority As Short
		Dim IsSelf As Boolean
		
		If (StrComp(Username, GetCurrentUsername, CompareMethod.Text) = 0) Then
			MyFlags = Flags
			
			SharedScriptSupport.BotFlags = MyFlags
			
			IsSelf = True
		End If
		
		'If (checkChannel(Username) > 0) Then
		'    Exit Sub
		'End If
		
		Select Case (Ping)
			Case 0
				LagIcon = 0
			Case 1 To 199
				LagIcon = LAG_1
			Case 200 To 299
				LagIcon = LAG_2
			Case 300 To 399
				LagIcon = LAG_3
			Case 400 To 499
				LagIcon = LAG_4
			Case 500 To 599
				LagIcon = LAG_5
			Case Is >= 600 Or -1
				LagIcon = LAG_6
			Case Else
				LagIcon = ICUNKNOWN
		End Select
		
		If ((Flags And USER_NOUDP) = USER_NOUDP) Then
			LagIcon = LAG_PLUG
		End If
		
		isPriority = (frmChat.lvChannel.Items.Count + 1)
		
		i = GetSmallIcon(Product, Flags, IconCode)
		
		'Special Cases
		'If i = ICSQUELCH Then
		'    'Debug.Print "Returned a SQUELCH icon"
		'    If ForcePosition > 0 Then isPriority = ForcePosition
		'
		If (((Flags And USER_BLIZZREP) = USER_BLIZZREP) Or ((Flags And USER_CHANNELOP) = USER_CHANNELOP) Or ((Flags And USER_SYSOP) = USER_SYSOP)) Then
			
			If (ForcePosition = 0) Then
				isPriority = 1
			Else
				isPriority = ForcePosition
			End If
			
		Else
			If (ForcePosition > 0) Then
				isPriority = ForcePosition
			End If
		End If
		
		If (i > frmChat.imlIcons.Images.Count) Then
			i = frmChat.imlIcons.Images.Count
		End If
		
		With frmChat.lvChannel
			.Enabled = False
			
			'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
			.Items.Insert(isPriority, Username, i)
			
			' store account name here so popup menus work
			'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Items.Item(isPriority).Tag = AccountName
			
			'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If (VB6.PixelsToTwipsX(.Columns.Item(2).Width) > 0) Then
				'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Items.Item(isPriority).SubItems.Add(Clan)
			End If
			
			'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If (VB6.PixelsToTwipsX(.Columns.Item(3).Width) > 0) Then
				'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListSubItems method lvChannel.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				.Items.Item(isPriority).SubItems.Add( ,  ,  , LagIcon)
			End If
			
			If (BotVars.NoColoring = False) Then
				'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Items.Item(isPriority).ForeColor = System.Drawing.ColorTranslator.FromOle(GetNameColor(Flags, 0, IsSelf))
			End If
			
			.Enabled = True
			
			'.Refresh
		End With
		
		frmChat.lblCurrentChannel.Text = frmChat.GetChannelString()
	End Sub
	
	
	Public Function CheckBlock(ByVal Username As String) As Boolean
		Dim s As String
		Dim i As Short
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(GetFilePath(FILE_FILTERS))) > 0) Then
            s = ReadINI("BlockList", "Total", GetFilePath(FILE_FILTERS))

            If (StrictIsNumeric(s)) Then
                i = CShort(s)
            Else
                Exit Function
            End If

            Username = PrepareCheck(Username)

            For i = 0 To i
                s = ReadINI("BlockList", "Filter" & i, GetFilePath(FILE_FILTERS))

                If (Username Like PrepareCheck(s)) Then
                    CheckBlock = True

                    Exit Function
                End If
            Next i
        End If
	End Function
	
	Public Function CheckMsg(ByVal Msg As String, Optional ByVal Username As String = "", Optional ByVal Ping As Integer = 0) As Boolean
		
		Dim i As Short
		
		Msg = PrepareCheck(Msg)
		
		For i = 0 To UBound(gFilters)
			If (Len(gFilters(i)) > 0) Then
				If (InStr(1, gFilters(i), "%", CompareMethod.Binary) > 0) Then
					Msg = PrepareCheck(DoReplacements(gFilters(i), Username, Ping))
				End If
				
				If (Msg Like "*" & gFilters(i) & "*") Then
					CheckMsg = True
					
					Exit Function
				End If
			End If
		Next i
	End Function
	
	Public Sub UpdateProfile()
		Dim s As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = MediaPlayer.TrackName
		
		If (s = vbNullString) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SetProfile("", ":[ ProfileAmp ]:" & vbCrLf & MediaPlayer.Name & " is not currently playing " & vbCrLf & "Last updated " & TimeOfDay & ", " & VB6.Format(Today, "d-MM-yyyy") & vbCrLf & CVERSION & " - http://www.stealthbot.net")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SetProfile("", ":[ ProfileAmp ]:" & vbCrLf & MediaPlayer.Name & " is currently playing: " & vbCrLf & s & vbCrLf & "Last updated " & TimeOfDay & ", " & VB6.Format(Today, "d-MM-yyyy") & vbCrLf & CVERSION & " - http://www.stealthbot.net")
		End If
	End Sub
	
	Public Function FlashWindow() As Boolean
		Dim pfwi As FLASHWINFO
		
		'Me.WindowState = vbMinimized
		
		With pfwi
			.cbSize = 20
			.dwFlags = FLASHW_ALL Or 12
			.dwTimeout = 0
			.hWnd = frmChat.Handle.ToInt32
			.uCount = 0
		End With
		
		FlashWindow = FlashWindowEx(pfwi)
	End Function
	
	Public Sub ReadyINet()
		'UPGRADE_ISSUE: VBControlExtender property INet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmChat.INet.Cancel()
	End Sub
	
	Public Function HTMLToRGBColor(ByVal s As String) As Integer
		HTMLToRGBColor = RGB(Val("&H" & Mid(s, 1, 2)), Val("&H" & Mid(s, 3, 2)), Val("&H" & Mid(s, 5, 2)))
	End Function
	
	Public Function StrictIsNumeric(ByVal sCheck As String, Optional ByRef AllowNegatives As Boolean = False) As Boolean
		Dim i As Integer
		
		StrictIsNumeric = True
		
		If (Len(sCheck) > 0) Then
			For i = 1 To Len(sCheck)
				If ((Asc(Mid(sCheck, i, 1)) = 45)) Then
					If ((Not AllowNegatives) Or (i > 1)) Then
						StrictIsNumeric = False
						Exit Function
					End If
				ElseIf (Not ((Asc(Mid(sCheck, i, 1)) >= 48) And (Asc(Mid(sCheck, i, 1)) <= 57))) Then 
					
					StrictIsNumeric = False
					
					Exit Function
				End If
			Next i
		Else
			StrictIsNumeric = False
		End If
	End Function
	
	Public Sub GetCountryData(ByRef CountryAbbrev As String, ByRef CountryName As String, ByRef sCountryCode As String)
		Const LOCALE_USER_DEFAULT As Integer = &H400
		Const LOCALE_ICOUNTRY As Integer = &H5 'Country Code
		Const LOCALE_SABBREVCTRYNAME As Integer = &H7 'abbreviated country name
		Const LOCALE_SENGCOUNTRY As Integer = &H1002 'English name of country
		
		Dim sBuf As String
		
		sBuf = New String(Chr(0), 256)
		Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, sBuf, Len(sBuf))
		CountryAbbrev = KillNull(sBuf)
		
		sBuf = New String(Chr(0), 256)
		Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, sBuf, Len(sBuf))
		CountryName = KillNull(sBuf)
		
		sBuf = New String(Chr(0), 256)
		Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICOUNTRY, sBuf, Len(sBuf))
		sCountryCode = KillNull(sBuf)
	End Sub
	
	Public Sub SetNagelStatus(ByVal lSocketHandle As Integer, ByVal bEnabled As Boolean)
		If (lSocketHandle > 0) Then
			If (bEnabled) Then
				Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_OFF, NAGLE_OPTLEN)
			Else
				Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_ON, NAGLE_OPTLEN)
			End If
		End If
	End Sub
	
	Public Sub EnableSO_KEEPALIVE(ByVal lSocketHandle As Integer)
		Call SetSockOpt(lSocketHandle, IPPROTO_TCP, SO_KEEPALIVE, True, 4) 'thanks Eric
	End Sub
	
	Public Function ProductCodeToFullName(ByVal pCode As String) As String
		ProductCodeToFullName = GetProductInfo(pCode).FullName
	End Function
	
	' Assumes that sIn has Length >=1
	Public Function PercentActualUppercase(ByVal sIn As String) As Double
		Dim UppercaseChars As Short
		Dim i As Short
		
		sIn = Replace(sIn, Space(1), vbNullString)
		
		If (Len(sIn) > 0) Then
			For i = 1 To Len(sIn)
				If (IsAlpha(Asc(Mid(sIn, i, 1)))) Then
					If (IsUppercase(Asc(Mid(sIn, i, 1)))) Then
						UppercaseChars = (UppercaseChars + 1)
					End If
				End If
			Next i
			
			PercentActualUppercase = CDbl(100 * (UppercaseChars / Len(sIn)))
		End If
	End Function
	
	Public Function MyUCase(ByVal sIn As String) As String
		Dim i As Short
		Dim CurrentByte As Byte
		
        If (Len(sIn) > 0) Then
            For i = 1 To Len(sIn)
                CurrentByte = Asc(Mid(sIn, i, 1))

                If (IsAlpha(CurrentByte)) Then
                    If (Not (IsUppercase(CurrentByte))) Then
                        Mid(sIn, i, 1) = Chr(CurrentByte - 32)
                    End If
                End If
            Next i
        End If
		
		MyUCase = sIn
	End Function
	
	Public Function IsAlpha(ByVal bCharValue As Byte) As Boolean
		IsAlpha = ((bCharValue >= 65 And bCharValue <= 90) Or (bCharValue >= 97 And bCharValue <= 122))
	End Function
	
	Public Function IsNumber(ByVal bCharValue As Byte) As Boolean
		IsNumber = ((bCharValue >= 48 And bCharValue <= 57))
	End Function
	
	Public Function IsUppercase(ByVal bCharValue As Byte) As Boolean
		IsUppercase = (bCharValue >= 65 And bCharValue <= 90)
	End Function
	
	Public Function VBHexToHTMLHex(ByVal sIn As String) As String
		sIn = Left((sIn) & "000000", 6)
		
		VBHexToHTMLHex = Mid(sIn, 5, 2) & Mid(sIn, 3, 2) & Mid(sIn, 1, 2)
	End Function
	
	'//10-15-2009 - Hdx - Updated url to new address
	Public Sub GetW3LadderProfile(ByVal sPlayer As String, ByVal eType As modEnum.enuWebProfileTypes)
		Const W3LadderURLFormat As String = "http://{0}.battle.net/war3/ladder/{1}-player-profile.aspx?Gateway={2}&PlayerName={3}"
		Dim W3LadderURL As String
		Dim W3WebProfileType As String
		Dim W3Realm As String
		Dim W3Domain As String
		W3Domain = "classic"
		
        If (Len(sPlayer) > 0) Then
            W3WebProfileType = IIf(eType = modEnum.enuWebProfileTypes.W3XP, PRODUCT_W3XP, PRODUCT_WAR3)
            W3Realm = GetW3Realm(sPlayer)
            If W3Realm = "Kalimdor" Then W3Domain = "asialadders"
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            W3LadderURL = StringFormat(W3LadderURLFormat, W3Domain, LCase(W3WebProfileType), W3Realm, NameWithoutRealm(sPlayer, 1))

            ShellOpenURL(W3LadderURL, sPlayer & "'s " & UCase(W3WebProfileType) & " ladder profile")
        End If
	End Sub
	
	Public Sub DoLastSeen(ByVal Username As String)
		Dim i As Short
		Dim found As Boolean
		
		If (colLastSeen.Count() > 0) Then
			For i = 1 To colLastSeen.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object colLastSeen.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (StrComp(colLastSeen.Item(i), Username, CompareMethod.Text) = 0) Then
					
					found = True
					
					Exit For
				End If
			Next i
		End If
		
		If (Not (found)) Then
			colLastSeen.Add(Username)
			
			If (colLastSeen.Count() > 15) Then
				Call colLastSeen.Remove(1)
			End If
		End If
	End Sub
	
	Public Sub SetTitle(ByVal sTitle As String)
		frmChat.Text = "[" & sTitle & "]" & " - " & CVERSION
	End Sub
	
	Public Function NameWithoutRealm(ByVal Username As String, Optional ByVal Strict As Byte = 0) As String
		If ((IsW3) And (Strict = 0)) Then
			NameWithoutRealm = Username
		Else
			If (InStr(1, Username, "@", CompareMethod.Binary) > 0) Then
				NameWithoutRealm = Left(Username, InStr(1, Username, "@") - 1)
			Else
				NameWithoutRealm = Username
			End If
		End If
	End Function
	
	Public Function GetCurrentUsername() As String
		
		GetCurrentUsername = ConvertUsername(CurrentUsername)
		
	End Function
	
	Public Function GetW3Realm(Optional ByVal Username As String = "") As String
        If (Len(Username) = 0) Then
            GetW3Realm = BotVars.Gateway
        Else
            If (InStr(1, Username, "@", CompareMethod.Binary) > 0) Then
                GetW3Realm = Mid(Username, InStr(1, Username, "@", CompareMethod.Binary) + 1)
            Else
                GetW3Realm = BotVars.Gateway
            End If
        End If
	End Function
	
	Public Function GetConfigFilePath() As String
		Static FilePath As String
		
        If (Len(FilePath) = 0) Then
            If ((Len(ConfigOverride) > 0)) Then
                FilePath = ConfigOverride
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FilePath = StringFormat("{0}Config.ini", GetProfilePath())
            End If
        End If
		
		If (InStr(1, FilePath, "\", CompareMethod.Binary) = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FilePath = StringFormat("{0}\{1}", CurDir(), FilePath)
		End If
		
		GetConfigFilePath = FilePath
	End Function
	
	Public Function GetFilePath(ByVal FileName As String, Optional ByRef DefaultPath As String = vbNullString) As String
		Dim s As String
		
		If (InStr(FileName, "\") = 0) Then
            If (Len(DefaultPath) = 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GetFilePath = StringFormat("{0}{1}", GetProfilePath(), FileName)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GetFilePath = StringFormat("{0}{1}", DefaultPath, FileName)
            End If
			
			s = Config.GetFilePath(FileName)
			
            If (Len(s) > 0) Then
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If (Len(Dir(s))) Then
                    GetFilePath = s
                End If
            End If
		Else
			GetFilePath = FileName
		End If
	End Function
	
	Public Function GetFolderPath(ByVal sFolderName As String) As String
		On Error GoTo ERROR_HANDLER
		Dim sPath As String
		sPath = GetFilePath(sFolderName)
		If (Not Right(sPath, 1) = "\") Then sPath = sPath & "\"
		GetFolderPath = sPath
		
		Exit Function
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.{3}()", Err.Number, Err.Description, "modOtherCode", "GetFolderPath"))
	End Function
	
	'Public Function OKToDoAutocompletion(ByRef sText As String, ByVal KeyAscii As Integer) As Boolean
	'    If BotVars.NoAutocompletion Then
	'        OKToDoAutocompletion = False
	'    Else
	'        If (InStr(sText, " ") = 0) And KeyAscii <> 32 Then       ' one word only
	'            OKToDoAutocompletion = True
	'        Else                                                            ' > 1 words
	'            If StrComp(Left$(sText, 1), "/") = 0 And KeyAscii <> 32 Then ' left character is a /
	'
	'                If InStr(InStr(sText, " ") + 1, sText, " ") = 0 Then    ' only two words
	'                    If StrComp(Left$(sText, 3), "/m ") = 0 Or _
	''                        StrComp(Left$(sText, 3), "/w ") = 0 Then
	'
	'                            OKToDoAutocompletion = True
	'                    Else
	'                        OKToDoAutocompletion = False
	'                    End If
	'                Else                                                    ' more than two words
	'                    OKToDoAutocompletion = False
	'                End If
	'            Else
	'                OKToDoAutocompletion = False
	'            End If
	'        End If
	'    End If
	'End Function
	
	' ProfileIndex param should only be used when changing profiles as a SET
	'   - colProfiles MUST be instantiated in order to call with a ProfileIndex > 0!
	Public Function GetProfilePath(Optional ByVal ProfileIndex As Short = 0) As String
		Static LastPath As String
		
		Dim s As String
		
		'    If ProfileIndex > 0 Then
		'        If ProfileIndex <= colProfiles.Count Then
		'            s = colProfiles.Item(ProfileIndex)
		'            'check for path\
		'            If Right$(s, 1) <> "\" Then
		'                s = s & "\"
		'            End If
		'
		'            GetProfilePath = s
		'        Else
		'            If LenB(LastPath) > 0 Then
		'                GetProfilePath = LastPath
		'            Else
		'                GetProfilePath = App.Path & "\"
		'            End If
		'        End If
		'    Else
        If Len(LastPath) > 0 Then
            GetProfilePath = LastPath
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetProfilePath = StringFormat("{0}\", CurDir())
        End If
		'    End If
		
		LastPath = GetProfilePath
	End Function
	
	Public Sub OpenReadme()
		ShellOpenURL("http://www.stealthbot.net/wiki/Main_Page", "the StealthBot Wiki")
	End Sub
	
	Sub ShellOpenURL(ByVal FullURL As String, Optional ByVal Description As String = vbNullString, Optional ByVal DisplayMessage As Boolean = True, Optional ByVal Verb As String = "open")
		ShellExecute(frmChat.Handle.ToInt32, Verb, FullURL, vbNullString, vbNullString, AppWinStyle.NormalFocus)
		
		If DisplayMessage Then
            If Len(Description) > 0 Then Description = Description & " at "
			frmChat.AddChat(RTBColors.ConsoleText, "Opening " & Description & "[ " & FullURL & " ]...")
		End If
	End Sub
	
	'Checks the queue for duplicate bans
	Public Sub RemoveBanFromQueue(ByVal sUser As String)
		Dim tmp As String
		
		tmp = "/ban " & sUser
		
		g_Queue.RemoveLines(tmp & "*")
		
		Dim strGateway As String
		If ((StrReverse(BotVars.Product) = PRODUCT_WAR3) Or (StrReverse(BotVars.Product) = PRODUCT_W3XP)) Then
			
			
			Select Case (BotVars.Gateway)
				Case "Lordaeron" : strGateway = "@USWest"
				Case "Azeroth" : strGateway = "@USEast"
				Case "Kalimdor" : strGateway = "@Asia"
				Case "Northrend" : strGateway = "@Europe"
			End Select
			
			If (InStr(1, tmp, strGateway, CompareMethod.Text) = 0) Then
				g_Queue.RemoveLines(tmp & strGateway & "*")
			End If
		End If
		
		'frmChat.AddChat vbRed, tmp & "*" & " : " & tmp & strGateway & "*"
	End Sub
	
	Public Function AllowedToTalk(ByVal sUser As String, ByVal Msg As String) As Boolean
		Dim i As Short
		
		' default to true
		AllowedToTalk = True
		
		If (Filters) Then
			If ((CheckBlock(sUser)) Or (CheckMsg(Msg, sUser, -5))) Then
				AllowedToTalk = False
			End If
		End If
	End Function
	
	
	' Used by the Individual Whisper Window system to determine whether a message should be
	'  forwarded to an IWW
	Public Function IrrelevantWhisper(ByVal sIn As String, ByVal sUser As String) As Boolean
		IrrelevantWhisper = False
		
		If InStr(sIn, Chr(223) & Chr(126) & Chr(223)) Then
			IrrelevantWhisper = True
			Exit Function
		End If
		
		sUser = NameWithoutRealm(sUser, 1)
		
		'Debug.Print "strComp(" & Left$(sIn, 12 + Len(sUser)) & ", Your friend " & sUser & ")"
		
		If StrComp(Left(sIn, 12 + Len(sUser)), "Your friend " & sUser) = 0 Then
			IrrelevantWhisper = True
			Exit Function
		End If
	End Function
	
	Public Sub UpdateSafelistedStatus(ByVal sUser As String, ByVal bStatus As Boolean)
		Dim i As Short
		
		'i = UsernameToIndex(sUser)
		
		If i > 0 Then
			'colUsersInChannel.Item(i).Safelisted = bStatus
		End If
	End Sub
	
	Public Sub AddBanlistUser(ByVal sUser As String, ByVal cOperator As String)
		Const MAX_BAN_COUNT As Short = 80
		
		Dim i As Short
		Dim bCount As Short
		
		' check for duplicate entry in banlist
		For i = 0 To UBound(gBans)
			If (StrComp(gBans(i).Username, StripRealm(sUser), CompareMethod.Text) = 0) Then
				Exit Sub
			End If
		Next i
		
		' count bans for channel operator
		For i = 0 To UBound(gBans)
			If (StrComp(gBans(i).cOperator, cOperator, CompareMethod.Text) = 0) Then
				bCount = (bCount + 1)
			End If
		Next i
		
		' if ban count for operator greater than operator
		' max, begin removing oldest bans.
		If (bCount >= MAX_BAN_COUNT) Then
			For i = 1 To (MAX_BAN_COUNT - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object gBans(i - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gBans(i - 1) = gBans(i)
			Next i
			
			With gBans(MAX_BAN_COUNT - 1)
				.Username = StripRealm(sUser)
				.UsernameActual = sUser
				.cOperator = cOperator
			End With
		Else
			With gBans(UBound(gBans))
				.Username = StripRealm(sUser)
				.UsernameActual = sUser
				.cOperator = cOperator
			End With
			
			ReDim Preserve gBans(UBound(gBans) + 1)
		End If
	End Sub
	
	' collapse array on top of the removed user
	Public Sub UnbanBanlistUser(ByVal sUser As String, ByVal cOperator As String)
		Dim i As Short
		Dim c As Short
		Dim NumRemoved As Short
		Dim iterations As Integer
		Dim uBnd As Short
		
		sUser = StripRealm(sUser)
		
		uBnd = UBound(gBans)
		
		While (i <= (uBnd - NumRemoved))
			If (StrComp(sUser, gBans(i).Username, CompareMethod.Text) = 0) Then
				If (i <> UBound(gBans)) Then
					For c = i To UBound(gBans)
						'UPGRADE_WARNING: Couldn't resolve default property of object gBans(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gBans(i) = gBans(i + 1)
					Next c
				End If
				
				' UBound(gBans) - 1 when UBound(gBans) = 0
				' causes an RTE.  Thanks PyroManiac and
				' Phix (2008-10-1).
				If (UBound(gBans)) Then
					ReDim Preserve gBans(UBound(gBans) - 1)
				Else
					ReDim gBans(0)
				End If
				
				NumRemoved = (NumRemoved + 1)
			Else
				i = (i + 1)
			End If
			
			iterations = (iterations + 1)
			
			If (iterations > 9000) Then
				If (MDebug("debug")) Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "Warning: Loop size limit exceeded " & "in UnbanBanlistUser()!")
					frmChat.AddChat(RTBColors.ErrorMessageText, "The banned-user list has been reset.. " & "hope it works!")
				End If
				
				ReDim gBans(0)
				
				Exit Sub
			End If
		End While
	End Sub
	
	Public Function isbanned(ByVal sUser As String) As Boolean
		Dim i As Short
		
		If (InStr(1, sUser, "#", CompareMethod.Binary)) Then
			sUser = Left(sUser, InStr(1, sUser, "#", CompareMethod.Binary) - 1)
			
			Debug.Print(sUser)
		End If
		
		For i = 0 To UBound(gBans)
			If (StrComp(sUser, gBans(i).UsernameActual, CompareMethod.Text) = 0) Then
				
				isbanned = True
				
				Exit Function
			End If
		Next i
	End Function
	
	Public Function IsValidIPAddress(ByVal sIn As String) As Boolean
		Dim s() As String
		Dim i As Short
		
		IsValidIPAddress = True
		
		If (InStr(1, sIn, ".", CompareMethod.Binary)) Then
			s = Split(sIn, ".")
			
			If (UBound(s) = 3) Then
				For i = 0 To 3
					If (Not (StrictIsNumeric(s(i)))) Then
						IsValidIPAddress = False
					End If
				Next i
			Else
				IsValidIPAddress = False
			End If
		Else
			IsValidIPAddress = False
		End If
	End Function
	
	Public Function GetNameColor(ByVal Flags As Integer, ByVal IdleTime As Integer, ByVal IsSelf As Boolean) As Integer
		'/* Self */
		If (IsSelf) Then
			'Debug.Print "Assigned color IsSelf"
			GetNameColor = FormColors.ChannelListSelf
			
			Exit Function
		End If
		
		'/* Squelched */
		If ((Flags And USER_SQUELCHED) = USER_SQUELCHED) Then
			'Debug.Print "Assigned color SQUELCH"
			GetNameColor = FormColors.ChannelListSquelched
			
			Exit Function
		End If
		
		'/* Blizzard */
		If (((Flags And USER_BLIZZREP) = USER_BLIZZREP) Or ((Flags And USER_SYSOP) = USER_SYSOP)) Then
			
			GetNameColor = COLOR_BLUE
			
			Exit Function
		End If
		
		'/* Operator */
		If ((Flags And USER_CHANNELOP) = USER_CHANNELOP) Then
			'Debug.Print "Assigned color OP"
			GetNameColor = FormColors.ChannelListOps
			Exit Function
		End If
		
		'/* Idle */
		If (IdleTime > BotVars.SecondsToIdle) Then
			'Debug.Print "Assigned color IDLE"
			GetNameColor = FormColors.ChannelListIdle
			Exit Function
		End If
		
		'/* Default */
		'Debug.Print "Assigned color NORMAL"
		GetNameColor = FormColors.ChannelListText
	End Function
	
	Public Function FlagDescription(ByVal Flags As Integer, ByVal ShowAll As Boolean) As String
		Dim sOut As String
		Dim sSep As String
		
		sOut = vbNullString
		sSep = vbNullString
		
		If (Flags And USER_SQUELCHED) = USER_SQUELCHED And ShowAll Then
			sOut = sOut & sSep & "squelched"
			sSep = ", "
		End If
		
		If (Flags And USER_CHANNELOP) = USER_CHANNELOP Then
			sOut = sOut & sSep & "channel operator"
			sSep = ", "
		End If
		
		If (Flags And USER_BLIZZREP) = USER_BLIZZREP Then
			sOut = sOut & sSep & "Blizzard representative"
			sSep = ", "
		End If
		
		If (Flags And USER_SYSOP) = USER_SYSOP Then
			sOut = sOut & sSep & "Battle.net system operator"
			sSep = ", "
		End If
		
		If (Flags And USER_NOUDP) = USER_NOUDP And ShowAll Then
			sOut = sOut & sSep & "UDP plug"
			sSep = ", "
		End If
		
		If (Flags And USER_BEEPENABLED) = USER_BEEPENABLED And ShowAll Then
			sOut = sOut & sSep & "beep enabled"
			sSep = ", "
		End If
		
		If (Flags And USER_GUEST) = USER_GUEST Then
			sOut = sOut & sSep & "guest"
			sSep = ", "
		End If
		
		If (Flags And USER_SPEAKER) = USER_SPEAKER Then
			sOut = sOut & sSep & "speaker"
			sSep = ", "
		End If
		
		If (Flags And USER_GFOFFICIAL) = USER_GFOFFICIAL Then
			sOut = sOut & sSep & "GF official"
			sSep = ", "
		End If
		
		If (Flags And USER_GFPLAYER) = USER_GFPLAYER Then
			sOut = sOut & sSep & "GF player"
			sSep = ", "
		End If
		
        If (Len(sOut) = 0) And ShowAll Then
            If (Flags = &H0) Then
                sOut = "normal"
            Else
                sOut = "unknown"
            End If
        End If
		
		FlagDescription = sOut
		
		If ShowAll Then
			FlagDescription = FlagDescription & " [0x" & Right("00000000" & Hex(Flags), 8) & "]"
		End If
	End Function
	
	'Returns TRUE if the specified argument was a command line switch,
	' such as -debug
	Public Function MDebug(ByVal sArg As String) As Boolean
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MDebug = InStr(1, CommandLine, StringFormat("-{0} ", sArg), CompareMethod.Text) > 0
	End Function
	
	Public Function SetCommandLine(ByRef sCommandLine As String) As Object
		On Error GoTo ERROR_HANDLER
		Dim sTemp As String
		Dim sSetting As String
		Dim sValue As String
		Dim sRet As String
		CommandLine = vbNullString
		sTemp = sCommandLine
		
		Do While Left(Trim(sTemp), 1) = "-"
			sTemp = Trim(sTemp)
			sSetting = Split(Mid(sTemp, 2) & Space(1), Space(1))(0)
			sTemp = Mid(sTemp, Len(sSetting) + 3)
			Select Case LCase(sSetting)
				Case "ppath"
					If (Left(sTemp, 1) = Chr(34)) Then
						If (InStr(2, sTemp, Chr(34), CompareMethod.Text) > 0) Then
							sValue = Mid(sTemp, 2, InStr(2, sTemp, Chr(34), CompareMethod.Text) - 2)
							sTemp = Mid(sTemp, Len(sValue) + 4)
						Else
							sValue = Mid(Split(sTemp & " -", " -")(0), 2)
							sTemp = Mid(sTemp, Len(sValue) + 3)
						End If
					Else
						sValue = Split(sTemp & " -", " -")(0)
						sTemp = Mid(sTemp, Len(sValue) + 2)
					End If
					
                    If (Len(sValue) > 0) Then
                        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                        If Dir(sValue) <> vbNullString Then
                            ChDir(sValue)
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CommandLine = StringFormat("{0}-ppath {1}{2}{1} ", CommandLine, Chr(34), sValue)
                        End If
                    End If
					
				Case "cpath"
					If (Left(sTemp, 1) = Chr(34)) Then
						If (InStr(2, sTemp, Chr(34), CompareMethod.Text) > 0) Then
							sValue = Mid(sTemp, 2, InStr(2, sTemp, Chr(34), CompareMethod.Text) - 2)
							sTemp = Mid(sTemp, Len(sValue) + 4)
						Else
							sValue = Mid(Split(sTemp & " -", " -")(0), 2)
							sTemp = Mid(sTemp, Len(sValue) + 3)
						End If
					Else
						sValue = Split(sTemp & " -", " -")(0)
						sTemp = Mid(sTemp, Len(sValue) + 2)
					End If
					
                    If (Len(sValue) > 0) Then
                        ConfigOverride = sValue
                        If (Len(GetConfigFilePath()) = 0) Then
                            ConfigOverride = vbNullString
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CommandLine = StringFormat("{0}-cpath {1}{2}{1} ", CommandLine, Chr(34), sValue)
                        End If
                    End If
					
				Case "addpath"
					If (Left(sTemp, 1) = Chr(34)) Then
						If (InStr(2, sTemp, Chr(34), CompareMethod.Text) > 0) Then
							sValue = Mid(sTemp, 2, InStr(2, sTemp, Chr(34), CompareMethod.Text) - 2)
							sTemp = Mid(sTemp, Len(sValue) + 4)
						Else
							sValue = Mid(Split(sTemp & " -", " -")(0), 2)
							sTemp = Mid(sTemp, Len(sValue) + 3)
						End If
					Else
						sValue = Split(sTemp & " -", " -")(0)
						sTemp = Mid(sTemp, Len(sValue) + 2)
					End If
					
					AddEnvPath(sValue)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CommandLine = StringFormat("{0}-addpath {1}{2}{1} ", CommandLine, Chr(34), sValue)
					
				Case "launcherver"
					If (Len(sTemp) >= 8) Then
						sValue = Left(sTemp, 8)
						'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lLauncherVersion = CInt(StringFormat("&H{0}", sValue))
						sTemp = Mid(sTemp, 9)
						
						'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CommandLine = StringFormat("{0}-launcherver {1} ", CommandLine, sValue)
					End If
					
				Case "launchererror"
					If (Left(sTemp, 1) = Chr(34)) Then
						If (InStr(2, sTemp, Chr(34), CompareMethod.Text) > 0) Then
							sValue = Mid(sTemp, 2, InStr(2, sTemp, Chr(34), CompareMethod.Text) - 2)
							sTemp = Mid(sTemp, Len(sValue) + 4)
						Else
							sValue = Mid(Split(sTemp & " -", " -")(0), 2)
							sTemp = Mid(sTemp, Len(sValue) + 3)
						End If
					Else
						sValue = Split(sTemp & " -", " -")(0)
						sTemp = Mid(sTemp, Len(sValue) + 2)
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sRet = StringFormat("{0}{1}YThe StealthBot Profile Launcher had an error!|", Chr(193), sRet)
                    If (Len(sValue) > 0) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        sRet = StringFormat("{0}{1}YOpen {2} for more information|", sRet, Chr(195), sValue)
                    End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CommandLine = StringFormat("{0}-launchererror {1}{2}{1} ", CommandLine, Chr(34), sValue)
					
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CommandLine = StringFormat("{0}-{1} ", CommandLine, sSetting)
			End Select
		Loop 
		
		If MDebug("debug") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sRet = StringFormat("{0} * Program executed in debug mode; unhandled packet information will be displayed.|", sRet)
		End If
		
		SetCommandLine = Split(sRet, "|")
		
		Exit Function
ERROR_HANDLER: 
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sRet = StringFormat("Error #{0}: {1} in modOtherCode.SetCommandLine()|CommandLine: {2}", Err.Number, Err.Description, sCommandLine)
		SetCommandLine = Split(sRet, "|")
		Err.Clear()
	End Function
	
	Private Function AddEnvPath(ByRef sPath As String) As Boolean
		On Error GoTo ERROR_HANDLER
		Dim sTemp As String
		Dim lRet As Integer
		AddEnvPath = False
		
		sTemp = New String(Chr(0), 1024)
		lRet = GetEnvironmentVariable("PATH", sTemp, Len(sTemp))
		
		If (Not lRet = 0) Then
			If (InStr(1, sTemp, sPath, CompareMethod.Text) = 0) Then
				sTemp = Left(sTemp, lRet)
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lRet = SetEnvironmentVariable("PATH", StringFormat("{0};{1}", sTemp, sPath))
				AddEnvPath = (lRet = 0)
				If (MDebug("debug")) Then
					frmChat.AddChat(RTBColors.ConsoleText, "AddEnvPath failed: Set")
					frmChat.AddChat(RTBColors.ConsoleText, StringFormat("PATH: {0}", sTemp))
					frmChat.AddChat(RTBColors.ConsoleText, StringFormat("ADD:  {0}", sPath))
				End If
			End If
		Else
			If (MDebug("debug")) Then
				frmChat.AddChat(RTBColors.ConsoleText, "AddEnvPath failed: Get")
				frmChat.AddChat(RTBColors.ConsoleText, StringFormat("Ret: {0}", lRet))
			End If
		End If
		
		Exit Function
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.{3}()", Err.Number, Err.Description, "modOtherCode", "AddEnvPath"))
		Err.Clear()
	End Function
	
	'Returns system uptime in milliseconds
	Public Function GetUptimeMS() As Double
		Dim mmt As MMTIME
		Dim lSize As Integer
		
        lSize = Runtime.InteropServices.Marshal.SizeOf(mmt)
		
		Call timeGetSystemTime(mmt, lSize)
		
		GetUptimeMS = LongToUnsigned(mmt.units)
	End Function
	
	
	'Public Function UsernameToIndex(ByVal sUsername As String) As Long
	'    Dim user        As clsUserInfo
	'    Dim FirstLetter As String * 1
	'    Dim i           As Integer
	'
	'    FirstLetter = Mid$(sUsername, 1, 1)
	'
	'    If (colUsersInChannel.Count > 0) Then
	'        For i = 1 To colUsersInChannel.Count
	'            Set user = colUsersInChannel.Item(i)
	'
	'            With user
	'                If (StrComp(Mid$(.Name, 1, 1), FirstLetter, vbTextCompare) = 0) Then
	'                    If (StrComp(sUsername, .Name, vbTextCompare) = 0) Then
	'                        UsernameToIndex = i
	'
	'                        Exit Function
	'                    End If
	'                End If
	'            End With
	'        Next i
	'
	'    End If
	'
	'    UsernameToIndex = 0
	'End Function
	
	
	Public Function checkChannel(ByVal NameToFind As String) As Short
		Dim lvItem As System.Windows.Forms.ListViewItem
		
		lvItem = frmChat.lvChannel.FindItemWithText(NameToFind)
		
		Dim i As Short
		If (lvItem Is Nothing) Then
			If BotVars.UseD2Naming Then
				checkChannel = 0
				For i = 1 To frmChat.lvChannel.Items.Count
					'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If frmChat.lvChannel.Items.Item(i).Tag = CleanUsername(ReverseConvertUsernameGateway(NameToFind)) Then
						checkChannel = i
						Exit For
					End If
				Next i
			Else
				checkChannel = 0
			End If
		Else
			checkChannel = lvItem.Index
		End If
	End Function
	
	
	Public Sub CheckPhrase(ByRef Username As String, ByRef Msg As String, ByVal mType As Byte)
		Dim i As Short
		
		If UBound(Catch_Renamed) = 0 Then
			If Catch_Renamed(0) = vbNullString Then Exit Sub
		End If
		
		For i = LBound(Catch_Renamed) To UBound(Catch_Renamed)
			If (Catch_Renamed(i) <> vbNullString) Then
				If (InStr(1, LCase(Msg), Catch_Renamed(i), CompareMethod.Text) <> 0) Then
					Call CaughtPhrase(Username, Msg, Catch_Renamed(i), mType)
					
					Exit Sub
				End If
			End If
		Next i
	End Sub
	
	
	Public Sub CaughtPhrase(ByVal Username As String, ByVal Msg As String, ByVal Phrase As String, ByVal mType As Byte)
		Dim i As Short
		Dim s As String
		
		i = FreeFile
		
		If Config.FlashOnCatchPhrases Then
			Call FlashWindow()
		End If
		
		Select Case (mType)
			Case CPTALK : s = "TALK"
			Case CPEMOTE : s = "EMOTE"
			Case CPWHISPER : s = "WHISPER"
		End Select
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir(GetFilePath(FILE_CAUGHT_PHRASES)) = vbNullString) Then
			FileOpen(i, GetFilePath(FILE_CAUGHT_PHRASES), OpenMode.Output)
			PrintLine(i, "<html>")
			FileClose(i)
		End If
		
		FileOpen(i, GetFilePath(FILE_CAUGHT_PHRASES), OpenMode.Append)
		If (LOF(i) > 10000000) Then
			FileClose(i)
			
			Call Kill(GetFilePath(FILE_CAUGHT_PHRASES))
			
			FileOpen(i, GetFilePath(FILE_CAUGHT_PHRASES), OpenMode.Output)
		End If
		
		Msg = Replace(Msg, "<", "&lt;", 1)
		Msg = Replace(Msg, ">", "&gt;", 1)
		
		PrintLine(i, "<B>" & VB6.Format(Today, "MM-dd-yyyy") & " - " & TimeOfDay & " - " & s & Space(1) & Username & ": </B>" & Replace(Msg, Phrase, "<i>" & Phrase & "</i>", 1) & "<br>")
		FileClose(i)
	End Sub
	
	
	Public Function DoReplacements(ByVal s As String, Optional ByRef Username As String = "", Optional ByRef Ping As Integer = 0) As String
		
		Dim gAcc As udtGetAccessResponse
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gAcc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gAcc = GetCumulativeAccess(Username)
		
		s = Replace(s, "%0", Username, 1)
		s = Replace(s, "%1", GetCurrentUsername, 1)
		s = Replace(s, "%c", g_Channel.Name, 1)
		s = Replace(s, "%bc", CStr(BanCount), 1)
		
		If (Ping > -2) Then
			s = Replace(s, "%p", CStr(Ping), 1)
		End If
		
		s = Replace(s, "%v", CVERSION, 1)
		s = Replace(s, "%a", IIf(gAcc.Rank >= 0, gAcc.Rank, "0"), 1)
		s = Replace(s, "%f", IIf(gAcc.Flags <> vbNullString, gAcc.Flags, "<none>"), 1)
		s = Replace(s, "%t", TimeString, 1)
		s = Replace(s, "%d", CStr(Today), 1)
		s = Replace(s, "%m", CStr(GetMailCount(Username)), 1)
		
		DoReplacements = s
	End Function
	
	Public Function ListFileLoad(ByVal sPath As String, Optional ByVal MaxItems As Short = -1) As Collection
		On Error GoTo ERROR_HANDLER
		
		Dim f As Short
		Dim i As Short
		Dim s As String
		Dim List As New Collection
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(sPath)) > 0) Then
            f = FreeFile()
            i = 0

            FileOpen(f, sPath, OpenMode.Input)
            If (LOF(f) > 0) Then
                Do
                    s = LineInput(f)

                    If Len(s) > 0 Then
                        List.Add(s)
                        i = i + 1
                    End If

                Loop Until EOF(f) Or (MaxItems >= 0 And i >= MaxItems)
            End If
            FileClose(f)
        End If
		
		ListFileLoad = List
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileLoad()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Sub ListFileAppendItem(ByVal sPath As String, ByVal Item As String)
		On Error GoTo ERROR_HANDLER
		
		Dim f As Short
		
		f = FreeFile
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(sPath)) > 0) Then
            FileOpen(f, sPath, OpenMode.Append)
            PrintLine(f, Item)
            FileClose(f)
        Else
            FileOpen(f, sPath, OpenMode.Output)
            PrintLine(f, Item)
            FileClose(f)
        End If
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileAppendItem()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub ListFileSave(ByVal sPath As String, ByVal List As Collection)
		On Error GoTo ERROR_HANDLER
		
		Dim f As Short
		Dim i As Integer
		
		f = FreeFile
		
		FileOpen(f, sPath, OpenMode.Output)
		' print each quote
		For i = 1 To List.Count()
			PrintLine(f, List.Item(i))
		Next i
		FileClose(f)
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileSave()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' Updated 4/10/06 to support millisecond pauses
	'  If using milliseconds pause for at least 100ms
	Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True, Optional ByVal milliseconds As Boolean = False)
		Dim i As Short
		
		If (AllowEvents) Then
			For i = 0 To (fSeconds * (IIf(milliseconds, 1, 1000))) \ 100
				'Debug.Print "sleeping 100ms"
				Call Sleep(100)
				
				System.Windows.Forms.Application.DoEvents()
			Next i
		Else
			Call Sleep(fSeconds * (IIf(milliseconds, 1, 1000)))
		End If
	End Sub
	
	Public Sub LogDbAction(ByVal ActionType As modEnum.enuDBActions, ByVal Caller As String, ByVal Target As String, ByVal TargetType As String, Optional ByVal Rank As Short = 0, Optional ByVal Flags As String = "", Optional ByVal Group As String = "")
		
		'Dim sPath  As String
		'Dim Action As String
		'Dim f      As Integer
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		
        If ((Len(Caller) = 0) Or (StrComp(Caller, "(console)", CompareMethod.Text) = 0)) Then
            Caller = "console"
        End If
		
		Select Case (ActionType)
			Case modEnum.enuDBActions.AddEntry
				str_Renamed = Caller & " adds " & Target
			Case modEnum.enuDBActions.ModEntry
				str_Renamed = Caller & " modifies " & Target
			Case modEnum.enuDBActions.RemEntry
				str_Renamed = Caller & " removes " & Target
		End Select
		
		If (StrComp(TargetType, "user", CompareMethod.Text) <> 0) Then
			str_Renamed = str_Renamed & " (" & LCase(TargetType) & ")"
		End If
		
		If (Rank > 0) Then
			str_Renamed = str_Renamed & " " & Rank
		End If
		
		If (Flags <> vbNullString) Then
			str_Renamed = str_Renamed & " " & Flags
		End If
		
		If (Group <> vbNullString) Then
			str_Renamed = str_Renamed & ", groups: " & Group
		End If
		
		g_Logger.WriteDatabase(str_Renamed)
		
		'f = FreeFile
		'sPath = GetProfilePath() & "\Logs\database.txt"
		
		'If (LenB(Dir$(sPath)) = 0) Then
		'    Open sPath For Output As #f
		'Else
		'    Open sPath For Append As #f
		'
		'    If ((LOF(f) > BotVars.MaxLogFileSize) And (BotVars.MaxLogFileSize > 0)) Then
		'        Close #f
		'
		'        Call Kill(sPath)
		'
		'        Open sPath For Output As #f
		'            Print #f, "Logfile cleared automatically on " & _
		''                Format(Now, "HH:MM:SS MM/DD/YY") & "."
		'    End If
		'End If
		
		'Select Case (ActionType)
		'    Case AddEntry
		'        Action = Caller & _
		''            " adds " & Target & " " & Instruction
		'
		'    Case ModEntry
		'        Action = Caller & _
		''            " modifies " & Target & " " & Instruction
		'
		'    Case RemEntry
		'        Action = Caller & " removes " & Target
		'End Select
		
		'Action = _
		''    Caller & " " & Action & Space(1) & Target & ": " & Instruction
		
		'g_Logger.WriteDatabase Action
		
		'Print #f, Action
		
		'Close #f
	End Sub
	
	Public Sub LogCommand(ByVal Caller As String, ByVal CString As String)
		On Error GoTo LogCommand_Error
		
		Dim sPath As String
		Dim Action As String
		Dim f As Short
		
        If (Len(CString) > 0) Then
            If (Len(Caller) = 0) Then
                Caller = "%console%"
            End If

            g_Logger.WriteCommand(Caller & " -> " & CString)
        End If
		
		Exit Sub
		
LogCommand_Error: 
		Debug.Print("Error " & Err.Number & " (" & Err.Description & ") in " & "Procedure; LogCommand; of; Module; modOtherCode; ")
		
		Exit Sub
	End Sub
	
	'Pos must be >0
	' Returns a single chunk of a string as if that string were Split() and that chunk
	' extracted
	' 1-based
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetStringChunk(ByVal str_Renamed As String, ByVal pos As Short) As Object
		Dim c As Short
		Dim i As Short
		Dim TargetSpace As Short
		
		'one two three
		'   1   2
		
		c = 0
		i = 1
		pos = pos
		
		' The string must have at least (pos-1) spaces to be valid
		While ((c < pos) And (i > 0))
			TargetSpace = i
			
			i = (InStr(i + 1, str_Renamed, Space(1), CompareMethod.Binary))
			
			c = (c + 1)
		End While
		
		If (c >= pos) Then
			c = InStr(TargetSpace + 1, str_Renamed, " ") ' check for another space (more afterwards)
			
			If (c > 0) Then
				GetStringChunk = Mid(str_Renamed, TargetSpace, c - (TargetSpace))
			Else
				GetStringChunk = Mid(str_Renamed, TargetSpace)
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetStringChunk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetStringChunk = ""
		End If
		
		GetStringChunk = Trim(GetStringChunk)
	End Function
	
	Function GetProductKey(Optional ByVal Product As String = "") As String
        If (Len(Product) = 0) Then
            Product = StrReverse(BotVars.Product)
        End If
		
		GetProductKey = GetProductInfo(Product).ShortCode
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Len(ReadCfg("Override", StringFormat("{0}ProdKey", Product))) > 0) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetProductKey = ReadCfg("Override", StringFormat("{0}ProdKey", Product))
        End If
	End Function
	
	Public Function InsertDummyQueueEntry() As Object
		' %%%%%blankqueuemessage%%%%%
		frmChat.AddQ("%%%%%blankqueuemessage%%%%%")
	End Function
	
	' This procedure splits a message by a specified Length, with optional line LinePostfixes
	' and split delimiters.
	' No longer puts a delim before [more] if it wasn't split by the delim -Ribose/2009-08-30
	Public Function SplitByLen(ByVal StringSplit As String, ByVal SplitLength As Integer, ByRef StringRet() As String, Optional ByVal LinePrefix As String = vbNullString, Optional ByVal LinePostfix As String = "", Optional ByVal OversizeDelimiter As String = " ") As Integer
		
		On Error GoTo ERROR_HANDLER
		
		' maximum size of battle.net messages
		Const BNET_MSG_LENGTH As Short = 223
		
		Dim lineCount As Integer ' stores line number
		Dim pos As Integer ' stores position of delimiter
		Dim strTmp As String ' stores working copy of StringSplit
		Dim length As Integer ' stores Length after LinePostfix
		Dim bln As Boolean ' stores result of delimiter split
		Dim s As String ' stores temp string for settings
		
		' check for custom line postfix
		s = Config.MultiLinePostfix
        If Len(s) > 0 Then
            If Left(s, 1) = "{" And Right(s, 1) = "}" Then
                LinePostfix = Mid(s, 2, Len(s) - 2)
            Else
                LinePostfix = s
            End If
        Else
            LinePostfix = "[more]"
        End If
		
		' initialize our array
		ReDim StringRet(0)
		
		' default our first index
		StringRet(0) = vbNullString
		
		If (SplitLength = 0) Then
			SplitLength = BNET_MSG_LENGTH
		End If
		
		If (Len(LinePrefix) >= SplitLength) Then
			Exit Function
		End If
		
		If (Len(LinePostfix) >= SplitLength) Then
			Exit Function
		End If
		
		' do loop until our string is empty
		Do While (StringSplit <> vbNullString)
			' resize array so that it can store
			' the next line
			ReDim Preserve StringRet(lineCount)
			
			' store working copy of string
			strTmp = LinePrefix & StringSplit
			
			' does our string already equal to or fall
			' below the specified Length?
			If (Len(strTmp) <= SplitLength) Then
				' assign our string to the current line
				StringRet(lineCount) = strTmp
			Else
				' Our string is over the size limit, so we're
				' going to postfix it.  Because of this, we're
				' going to have to calculate the Length after
				' the postfix has been accounted for.
				length = (SplitLength - Len(LinePostfix))
				
				' if we're going to be splitting the oversized
				' message at a specified character, we need to
				' determine the position of the character in the
				' string
				If (OversizeDelimiter <> vbNullString) Then
					' grab position of delimiter character that is the closest to our
					' specified Length
					pos = InStrRev(strTmp, OversizeDelimiter, length, CompareMethod.Text)
				End If
				
				' if the delimiter we were looking for was found,
				' and the position was greater than or equal to
				' half of the message (this check prevents breaks
				' in unecessary locations), split the message
				' accordingly.
				If ((pos) And (pos >= System.Math.Round(length / 2))) Then
					' truncate message
					strTmp = Mid(strTmp, 1, pos)
					
					' indicate that an additional
					' character will require removal
					' from official copy
					'bln = (Not KeepDelim)
				Else
					' truncate message
					strTmp = Mid(strTmp, 1, length)
				End If
				
				' store truncated message in line
				StringRet(lineCount) = strTmp & LinePostfix
			End If
			
			' remove line from official string
			StringSplit = Mid(StringSplit, (Len(strTmp) - Len(LinePrefix)) + 1)
			
			' if we need to remove an additional
			' character, lets do so now.
			'If (bln) Then
			'    StringSplit = Mid$(StringSplit, Len(OversizeDelimiter) + 1)
			'End If
			
			' increment line counter
			lineCount = (lineCount + 1)
		Loop 
		
		SplitByLen = lineCount
		
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: " & Err.Description & " in SplitByLen().")
		
		Exit Function
	End Function
	
	Public Function UsernameRegex(ByVal Username As String, ByVal sPattern As String) As Boolean
		Dim prepName As String
		Dim prepPatt As String
		
		prepName = Replace(Username, "[", "{")
		prepName = Replace(prepName, "]", "}")
		prepName = LCase(prepName)
		
		prepPatt = Replace(sPattern, "\[", "{")
		prepPatt = Replace(prepPatt, "\]", "}")
		prepPatt = LCase(prepPatt)
		
		UsernameRegex = (prepName Like prepPatt)
	End Function
	
	Public Function convertAlias(ByVal cmdName As String) As String
		On Error GoTo ERROR_HANDLER
		
		Dim commandDoc As New clsCommandDocObj
		
		
		Dim commands As MSXML2.DOMDocument60
		'UPGRADE_NOTE: Alias was upgraded to Alias_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Alias_Renamed As MSXML2.IXMLDOMNode
		If (Len(cmdName) > 0) Then
			
			commands = commandDoc.XMLDocument
			
			'If (Dir$(sCommandsPath) = vbNullString) Then
			'    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
			''        "working directory.")
			'
			'    Exit Function
			'End If
			'
			'Call commands.Load(sCommandsPath)
			
			If (InStr(1, cmdName, "'", CompareMethod.Binary) > 0) Then
				'UPGRADE_NOTE: Object commandDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				commandDoc = Nothing
				Exit Function
			End If
			
			cmdName = Replace(cmdName, "\", "\\")
			cmdName = Replace(cmdName, "'", "&apos;")
			
			'// 09/03/2008 JSM - Modified code to use the <aliases> element
			Alias_Renamed = commands.documentElement.selectSingleNode("./command/aliases/alias[translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(cmdName) & "']")
			
			'Set Alias = _
			''    commands.documentElement.selectSingleNode( _
			''        "./command/aliases/alias[contains(text(), '" & cmdName & "')]")
			
			If (Not (Alias_Renamed Is Nothing)) Then
				'// 09/03/2008 JSM - Modified code to use the <aliases> element
				convertAlias = Alias_Renamed.parentNode.parentNode.Attributes.getNamedItem("name").Text
				
				Exit Function
			End If
		End If
		
		convertAlias = cmdName
		'UPGRADE_NOTE: Object commandDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		commandDoc = Nothing
		
		Exit Function
		
ERROR_HANDLER: 
		
		Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error: XML Database Processor has encountered an error " & "during alias lookup.")
		
		convertAlias = cmdName
		
		Exit Function
		
	End Function
	
	' Fixed font issue when an element was only 1 character long -Pyro (9/28/08)
	' Fixed issue with displaying null text.
	
	' I changed the location of where fontStr was being declared. I can't figure out why it makes any difference, but _
	'when I have it declared above I get memory errors, my IDE crashes, I runtime erro 380 and _
	'- Error (#-2147417848): Method 'SelFontName' of object 'IRichText' failed in DisplayRichText().
	'I believe it's something to do with the subclassing overwriting the memory, but it only occurs when run from the IDE. - FrOzeN
	
	Public Sub DisplayRichText(ByRef rtb As System.Windows.Forms.RichTextBox, ByRef saElements() As Object)
		On Error GoTo ERROR_HANDLER
		
		Dim arr() As Object
		Dim s As String
		Dim L As Integer
		Dim lngVerticalPos As Integer
		Dim Diff As Integer
		Dim i As Integer
		Dim intRange As Integer
		Dim blUnlock As Boolean
		Dim LogThis As Boolean
		Dim length As Integer
		Dim Count As Integer
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim arrCount As Integer
		Dim SelStart As Integer
		Dim SelLength As Integer
		Dim blnHasFocus As Boolean
		Dim blnAtEnd As Boolean
		
		Static RichTextErrorCounter As Short
		
		' *****************************************
		'              SANITY CHECKS
		' *****************************************
		
		'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (StrictIsNumeric(saElements(0))) Then
			Count = 2
			
			For i = LBound(saElements) To UBound(saElements) Step 2
				'UPGRADE_ISSUE: As Variant was removed from ReDim arr(0 To Count) statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="19AFCB41-AA8E-4E6B-A441-A3E802E5FD64"'
				ReDim Preserve arr(Count)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object arr(Count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				arr(Count) = saElements(i + 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object arr(Count - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				arr(Count - 1) = saElements(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object arr(Count - 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				arr(Count - 2) = rtb.Font.Name
				
				Count = Count + 3
			Next i
			
			saElements = VB6.CopyArray(arr)
		End If
		
		rtbChatLength = Len(rtb.Text)
		
		For i = LBound(saElements) To UBound(saElements) Step 3
			If (i >= UBound(saElements)) Then
				Exit Sub
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrictIsNumeric(saElements(i + 1)) = False) Then
				Exit Sub
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			length = length + Len(KillNull(saElements(i + 2)))
		Next i
		
		If (length = 0) Then
			Exit Sub
		End If
		
		If ((BotVars.LockChat = False) Or (rtb.RTF <> frmChat.rtbChat.RTF)) Then
			
			' store rtb carat and whether rtb has focus
			With rtb
				SelStart = .SelectionStart
				SelLength = .SelectionLength
				'UPGRADE_WARNING: Control property rtb.Parent was upgraded to rtb.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
				blnHasFocus = (rtb.FindForm.ActiveControl Is rtb And rtb.FindForm.WindowState <> System.Windows.Forms.FormWindowState.Minimized)
				' whether it's at the end or within one vbCrLf of the end
				blnAtEnd = (SelStart >= rtbChatLength - 2)
			End With
			
			lngVerticalPos = IsScrolling(rtb)
			
			If (lngVerticalPos) Then
				rtb.Visible = False
				
				' below causes smooth scrolling, but also screen flickers :(
				'LockWindowUpdate rtb.hWnd
				
				blUnlock = True
			End If
			
			If (rtb.RTF = frmChat.rtbChat.RTF) Then
				LogThis = (BotVars.Logging > 0)
			ElseIf (rtb.RTF = frmChat.rtbWhispers.RTF) Then 
				LogThis = (BotVars.Logging > 0)
			End If
			
			If ((BotVars.MaxBacklogSize) And (rtbChatLength >= BotVars.MaxBacklogSize)) Then
				If (blUnlock = False) Then
					rtb.Visible = False
					
					' below causes smooth scrolling, but also screen flickers :(
					'LockWindowUpdate rtb.hWnd
				End If
				
				With rtb
					.SelectionStart = 0
					.SelectionLength = InStr(1, .Text, vbLf, CompareMethod.Binary)
					' remove line from stored selection
					SelStart = SelStart - .SelectionLength
					' if selection included part of what was removed, add negative start point
					' to length to get difference length and start selection at 0
					If SelStart < 0 Then
						SelLength = SelLength + SelStart
						SelStart = 0
						' if new length is negative, then the selection is now gone, so selection
						' length should be 0
						If SelLength < 0 Then SelLength = 0
					End If
					.SelectionFont = VB6.FontChangeName(.SelectionFont, rtb.Font.Name)
					.SelectionFont = VB6.FontChangeSize(.SelectionFont, rtb.Font.SizeInPoints)
					.SelectedText = ""
				End With
				
				If (blUnlock = False) Then
					rtb.Visible = True
					
					' below causes smooth scrolling, but also screen flickers :(
					'LockWindowUpdate &H0
				End If
			End If
			
			s = GetTimeStamp()
			
			With rtb
				.SelectionStart = Len(.Text)
				.SelectionLength = 0
				.SelectionFont = VB6.FontChangeName(.SelectionFont, rtb.Font.Name)
				.SelectionFont = VB6.FontChangeSize(.SelectionFont, rtb.Font.SizeInPoints)
				.Font = VB6.FontChangeBold(.SelectionFont, False)
				.SelectionFont = VB6.FontChangeItalic(.SelectionFont, False)
				.SelectionFont = VB6.FontChangeUnderline(.SelectionFont, False)
				.SelectionColor = System.Drawing.ColorTranslator.FromOle(RTBColors.TimeStamps)
				.SelectedText = s
				.SelectionLength = Len(.SelectedText)
			End With
			
			For i = LBound(saElements) To UBound(saElements) Step 3
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (InStr(1, saElements(i + 2), Chr(0), CompareMethod.Binary) > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					KillNull(saElements(i + 2))
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ((StrictIsNumeric(saElements(i + 1))) And (Len(saElements(i + 2)) > 0)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					L = InStr(1, saElements(i + 2), "{\rtf", CompareMethod.Text)
					
					While (L > 0)
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(i + 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mid(saElements(i + 2), L + 1, 1) = "/"
						
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						L = InStr(1, saElements(i + 2), "{\rtf", CompareMethod.Text)
					End While
					
					L = Len(rtb.Text)
					
					With rtb
						.SelectionStart = L
						.SelectionLength = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectionFont = VB6.FontChangeName(.SelectionFont, saElements(i))
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectionColor = System.Drawing.ColorTranslator.FromOle(saElements(i + 1))
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectedText = saElements(i + 2) & Left(vbCrLf, -2 * CInt((i + 2) = UBound(saElements)))
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						str_Renamed = str_Renamed & saElements(i + 2)
						.SelectionLength = Len(.SelectedText)
					End With
				End If
			Next i
			
			If (LogThis) Then
				If (rtb.RTF = frmChat.rtbChat.RTF) Then
					g_Logger.WriteChat(str_Renamed)
				ElseIf (rtb.RTF = frmChat.rtbWhispers.RTF) Then 
					g_Logger.WriteWhisper(str_Renamed)
				End If
			End If
			
			ColorModify(rtb, L)
			
			If (blUnlock) Then
				SendMessage(rtb.Handle.ToInt32, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0)
				
				rtb.Visible = True
				
				' below causes smooth scrolling, but also screen flickers :(
				'LockWindowUpdate &H0
			End If
			
			With rtb
				' if has focus
				If blnHasFocus Then
					' restore carat location and selection if not previously at end
					If Not blnAtEnd Then
						.SelectionStart = SelStart
						.SelectionLength = SelLength
					End If
					
					' restore focus
					'.SetFocus
				End If
			End With
		End If
		
		RichTextErrorCounter = 0
		
		Exit Sub
		
ERROR_HANDLER: 
		
		RichTextErrorCounter = RichTextErrorCounter + 1
		If RichTextErrorCounter > 2 Then
			RichTextErrorCounter = 0
			Exit Sub
		End If
		
		If (Err.Number = 13 Or Err.Number = 91) Then
			Exit Sub
		End If
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DisplayRichText().")
		
		Exit Sub
		
	End Sub
	
	Public Function IsScrolling(ByRef rtb As System.Windows.Forms.RichTextBox) As Integer
		
		Dim lngVerticalPos As Integer
		Dim difference As Integer
		Dim range As Short
		
		If (g_OSVersion.IsWin2000Plus()) Then
			
			GetScrollRange(rtb.Handle.ToInt32, SB_VERT, 0, range)
			
			lngVerticalPos = SendMessage(rtb.Handle.ToInt32, EM_GETTHUMB, 0, 0)
			
			If ((lngVerticalPos = 0) And (range > 0)) Then
				lngVerticalPos = 1
			End If
			
			difference = ((lngVerticalPos + (VB6.PixelsToTwipsY(rtb.Height) / VB6.TwipsPerPixelY)) - range)
			
			' In testing it appears that if the value I calcuate as Diff is negative,
			' the scrollbar is not at the bottom.
			If (difference < 0) Then
				IsScrolling = lngVerticalPos
			End If
			
		End If
		
	End Function
	
	Public Function ResolveHost(ByVal strHostName As String) As String
		Dim lServer As Integer
		Dim HostInfo As HOSTENT
		Dim ptrIP As Integer
		Dim strIP As String
		
		'Do we have an IP address or a hostname?
		If Not IsValidIPAddress(strHostName) Then
			'Resolve the IP.
			lServer = gethostbyname(strHostName)
			
			If lServer = 0 Then
				ResolveHost = vbNullString
				Exit Function
			Else
				'Copy data to HOSTENT struct.
				'UPGRADE_WARNING: Couldn't resolve default property of object HostInfo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(HostInfo, lServer, Len(HostInfo))
				
				If HostInfo.h_addrtype = 2 Then
					CopyMemory(ptrIP, HostInfo.h_addr_list, 4)
					CopyMemory(lServer, ptrIP, 4)
					ptrIP = inet_ntoa(lServer)
					strIP = Space(lstrlen(ptrIP))
					lstrcpy(strIP, ptrIP)
					
					ResolveHost = strIP
				Else
					ResolveHost = vbNullString
					Exit Function
				End If
			End If
		Else
			ResolveHost = strHostName
		End If
	End Function
	
	Public Sub CloseAllConnections(Optional ByRef ShowMessage As Boolean = True)
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNLS.CtlState <> 0) Then
			frmChat.sckBNLS.Close()
		End If
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNet.CtlState <> 0) Then
			frmChat.sckBNet.Close()
		End If
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckMCP.CtlState <> 0) Then
			frmChat.sckMCP.Close()
		End If
		
		If (ShowMessage) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "All connections closed.")
		End If
		
		BNLSAuthorized = False
		
		SetTitle("Disconnected")
		
		frmChat.UpdateTrayTooltip()
		
		g_Online = False
		
		RunInAll("Event_ServerError", "All connections closed.")
	End Sub
	
	Public Sub BuildProductInfo()
		' 4-digit code, short code, short name, long name, number of keys, BNLS ID, logon system
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(0) = CreateProductInfo("UNKW", vbNullString, "Unknown Product", 0, &H0, &H0, &H0)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(1) = CreateProductInfo(PRODUCT_STAR, "SC", "StarCraft", 1, &H1, BNCS_NLS, &HD3)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(2) = CreateProductInfo(PRODUCT_SEXP, "SC", "StarCraft Broodwar", 1, &H2, BNCS_NLS, &HD3)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(3) = CreateProductInfo(PRODUCT_W2BN, "W2", "WarCraft II: Battle.net Edition", 1, &H3, BNCS_OLS, &H4F)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(4) = CreateProductInfo(PRODUCT_D2DV, "D2", "Diablo II", 1, &H4, BNCS_NLS, &HE)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(5) = CreateProductInfo(PRODUCT_D2XP, "D2X", "Diablo II: Lord of Destruction", 2, &H5, BNCS_NLS, &HE)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(6) = CreateProductInfo(PRODUCT_WAR3, "W3", "WarCraft III: Reign of Chaos", 1, &H7, BNCS_NLS, &H1B)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(7) = CreateProductInfo(PRODUCT_W3XP, "W3", "WarCraft III: The Frozen Throne", 2, &H8, BNCS_NLS, &H1B)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(8) = CreateProductInfo(PRODUCT_DSHR, "DS", "Diablo Shareware", 0, &HA, BNCS_OLS, &H2A)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(9) = CreateProductInfo(PRODUCT_DRTL, "D1", "Diablo", 0, &H9, BNCS_OLS, &H2A)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(10) = CreateProductInfo(PRODUCT_SSHR, "SS", "StarCraft Shareware", 0, &HB, BNCS_LLS, &HA9)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(11) = CreateProductInfo(PRODUCT_JSTR, "JS", "Japanese StarCraft", 1, &H6, BNCS_LLS, &HA9)
		'UPGRADE_WARNING: Couldn't resolve default property of object ProductList(12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ProductList(12) = CreateProductInfo(PRODUCT_CHAT, "CHAT", "Telnet Chat", 0, &H0, &H0, &H0)
	End Sub
	
	Private Function CreateProductInfo(ByVal sCode As String, ByVal sShort As String, ByVal sLongName As String, ByVal iKeys As Short, ByVal iBnlsId As Integer, ByVal iLogonSystem As Integer, ByVal iVerByte As Integer) As udtProductInfo
		Dim pi As udtProductInfo
		pi.Code = UCase(sCode)
		pi.ShortCode = UCase(sShort)
		pi.FullName = sLongName
		pi.KeyCount = iKeys
		pi.BNLS_ID = iBnlsId
		pi.LogonSystem = iLogonSystem
		pi.VersionByte = iVerByte
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CreateProductInfo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CreateProductInfo = pi
	End Function
	
	Public Function GetProductInfo(ByVal sProductCode As String) As udtProductInfo
		Dim pi As udtProductInfo
		Dim Index As Short
		sProductCode = UCase(sProductCode)
		
		For Index = 0 To UBound(ProductList)
			'UPGRADE_WARNING: Couldn't resolve default property of object pi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pi = ProductList(Index)
			
			If StrComp(pi.Code, sProductCode, CompareMethod.Binary) = 0 Or StrComp(pi.Code, StrReverse(sProductCode), CompareMethod.Binary) = 0 Or StrComp(pi.ShortCode, sProductCode, CompareMethod.Binary) = 0 Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object GetProductInfo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetProductInfo = pi
				Exit Function
			End If
		Next 
		GetProductInfo = ProductList(0)
	End Function
	
	'Returns the number of monitors active on the computer.
	Public Function GetMonitorCount() As Integer
		'UPGRADE_WARNING: Add a delegate for AddressOf MonitorEnumProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		EnumDisplayMonitors(0, 0, AddressOf MonitorEnumProc, GetMonitorCount)
	End Function
	
	Private Function MonitorEnumProc(ByVal hMonitor As Integer, ByVal hDCMonitor As Integer, ByVal lprcMonitor As Integer, ByRef dwData As Integer) As Integer
		dwData = dwData + 1
		MonitorEnumProc = 1
	End Function
End Module