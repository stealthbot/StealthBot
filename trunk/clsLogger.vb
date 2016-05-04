Option Strict Off
Option Explicit On
Friend Class clsLogger
	' clsLog.Cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private Enum LOG_TYPES
		WARNING_MSG = &H1
		ERROR_MSG = &H2
		EVENT_MSG = &H3
		DEBUG_MSG = &H4
		CHAT_MSG = &H5
		WHISPER_MSG = &H6
		SCK_MSG = &H7
		COMMAND_MSG = &H8
	End Enum
	
	Private m_logsCreated As Collection
	Private m_logPath As String ' the path to the log file
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_logsCreated = New Collection
		
		m_logPath = GetFolderPath("Logs")
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub Class_Teriminate()
		If (BotVars.Logging = 1) Then
			RemoveLogsCreated()
		End If
		
		'UPGRADE_NOTE: Object m_logsCreated may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_logsCreated = Nothing
	End Sub
	
	
	Public Property LogPath() As String
		Get
			LogPath = m_logPath
		End Get
		Set(ByVal Value As String)
			m_logPath = Value
			
			If (Not (Right(m_logPath, 1) = "\")) Then
				m_logPath = m_logPath & "\"
			End If
		End Set
	End Property
	
	Public Function WriteWarning(ByVal sWarning As String, Optional ByVal source As String = "", Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		source = LCase(source)
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(source) > 0) Then
			WriteLine_Renamed(LOG_TYPES.WARNING_MSG, "warning " & source & " " & sWarning, TimeDate)
		Else
			WriteLine_Renamed(LOG_TYPES.WARNING_MSG, "warning " & sWarning, TimeDate)
		End If
	End Function
	
	Public Function WriteError(ByVal sError As String, Optional ByVal source As String = "", Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		source = LCase(source)
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(source) > 0) Then
			WriteLine_Renamed(LOG_TYPES.ERROR_MSG, "error " & source & " " & sError, TimeDate)
		Else
			WriteLine_Renamed(LOG_TYPES.ERROR_MSG, "error " & sError, TimeDate)
		End If
	End Function
	
	Public Function WriteEvent(ByVal sTitle As String, ByVal sEvent As String, Optional ByVal source As String = "", Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		source = LCase(source)
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(source) > 0) Then
			WriteLine_Renamed(LOG_TYPES.EVENT_MSG, "event " & sTitle & " " & source & " " & sEvent, TimeDate)
		Else
			WriteLine_Renamed(LOG_TYPES.EVENT_MSG, "event " & sTitle & " " & sEvent, TimeDate)
		End If
	End Function
	
	Public Function WriteCommand(ByVal sCommand As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		If (BotVars.LogCommands) Then
			WriteEvent("command", sCommand, CStr(TimeDate))
		End If
	End Function
	
	Public Function WriteDatabase(ByVal sEvent As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		If (BotVars.LogDBActions) Then
			WriteEvent("database", sEvent, CStr(TimeDate))
		End If
	End Function
	
	Public Function WriteDebug(ByVal sDebugMessage As String, Optional ByVal source As String = "", Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		source = LCase(source)
		
		If (isDebug) Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(source) > 0) Then
				WriteLine_Renamed(LOG_TYPES.DEBUG_MSG, "debug " & source & " " & sDebugMessage, TimeDate)
			Else
				WriteLine_Renamed(LOG_TYPES.DEBUG_MSG, "debug " & sDebugMessage, TimeDate)
			End If
		End If
	End Function
	
	Public Function WriteChat(ByVal sMessage As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		If (BotVars.Logging >= 1) Then
			WriteLine_Renamed(LOG_TYPES.CHAT_MSG, sMessage, TimeDate)
		End If
	End Function
	
	Public Function WriteWhisper(ByVal sMessage As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		If (BotVars.Logging >= 1) Then
			WriteLine_Renamed(LOG_TYPES.WHISPER_MSG, sMessage, TimeDate)
		End If
	End Function
	
	Public Function WriteSckData(ByVal sMessage As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		If (LogPacketTraffic) Then
			WriteLine_Renamed(LOG_TYPES.SCK_MSG, sMessage, TimeDate)
		End If
	End Function
	
	'UPGRADE_NOTE: WriteLine was upgraded to WriteLine_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function WriteLine_Renamed(ByVal ltype As LOG_TYPES, ByVal line As String, Optional ByRef TimeDate As Date = #12:00:00 AM#) As Object
		On Error GoTo ERROR_HANDLER
		
		Dim f As Short
		Dim filePath As String
		
		Select Case (ltype)
			Case LOG_TYPES.WARNING_MSG, LOG_TYPES.ERROR_MSG, LOG_TYPES.EVENT_MSG, LOG_TYPES.DEBUG_MSG
				filePath = (m_logPath & "master.txt")
				
			Case LOG_TYPES.COMMAND_MSG
				filePath = (m_logPath & "commands.txt")
				
			Case LOG_TYPES.CHAT_MSG
				filePath = (m_logPath & Datestamp() & ".txt")
				
			Case LOG_TYPES.WHISPER_MSG
				filePath = (m_logPath & Datestamp() & "-WHISPERS.txt")
				
			Case LOG_TYPES.SCK_MSG
				filePath = (m_logPath & Datestamp() & "-PACKETLOG.txt")
		End Select
		
		f = OpenLog(ltype, filePath)
		
		If (f > 0) Then
			Select Case (ltype)
				Case LOG_TYPES.CHAT_MSG, LOG_TYPES.WHISPER_MSG, LOG_TYPES.SCK_MSG
					PrintLine(f, "[" & Timestamp(TimeDate) & "] " & line)
				Case Else
					PrintLine(f, "[" & Datestamp(TimeDate) & " " & Timestamp(TimeDate) & "] " & line)
			End Select
			
			FileClose(f)
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		
		MsgBox("Error (#" & Err.Number & "): " & Err.Description & " in WriteLine().")
	End Function
	
	Private Function OpenLog(ByVal ltype As LOG_TYPES, ByVal Path As String) As Short
		On Error GoTo ERROR_HANDLER
		
		Static failed_attempts As Short
		
		Dim f As Short
		Dim i As Double
		
		f = FreeFile
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim dir_path As String
		Dim arr() As String
		Dim bln As Boolean
		Dim lineCount As Integer
		Dim offset As Short
		Dim bytes As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		If (Dir(Path) = vbNullString) Then
			
			dir_path = Mid(Path, 1, InStrRev(Path, "\"))
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If (Dir(dir_path) = vbNullString) Then
				MkDir(dir_path)
			End If
			
			FileOpen(f, Path, OpenMode.Output)
			
			If ((ltype = LOG_TYPES.CHAT_MSG) Or (ltype = LOG_TYPES.WHISPER_MSG)) Then
				For i = 1 To m_logsCreated.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object m_logsCreated(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (StrComp(m_logsCreated.Item(i), Path, CompareMethod.Text) = 0) Then
						Exit For
					End If
				Next i
				
				If (i >= m_logsCreated.Count() + 1) Then
					m_logsCreated.Add(Path)
				End If
			End If
		Else
			
			If (BotVars.MaxLogFileSize) Then
				FileOpen(f, Path, OpenMode.Input)
				If (LOF(f) >= BotVars.MaxLogFileSize) Then
					Do Until (EOF(f))
						str_Renamed = LineInput(f)
						
						ReDim Preserve arr(lineCount)
						
						arr(lineCount) = str_Renamed
						
						lineCount = (lineCount + 1)
					Loop 
					
					bln = True
				End If
				FileClose(f)
			End If
			
			If (bln) Then
				
				For i = UBound(arr) To 0 Step -1
					bytes = (bytes + Len(arr(i)))
					
					If (bytes >= BotVars.MaxLogFileSize) Then
						offset = i
						
						Exit For
					End If
				Next i
				
				FileOpen(f, Path, OpenMode.Output)
				For i = offset + 1 To UBound(arr)
					PrintLine(f, arr(i))
				Next i
				FileClose(f)
			End If
			
			FileOpen(f, Path, OpenMode.Append)
		End If
		
		OpenLog = f
		
		failed_attempts = 0
		
		Exit Function
		
ERROR_HANDLER: 
		
		' permission denied?  is someone else trying to save the file?
		If (Err.Number = 70) Then
			failed_attempts = (failed_attempts + 1)
			
			If (failed_attempts >= 3) Then
				Exit Function
			End If
			
			OpenLog(ltype, Path)
		End If
	End Function
	
	Public Sub RemoveLogsCreated()
		On Error Resume Next
		
		Dim i As Short
		
		For i = 1 To m_logsCreated.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object m_logsCreated(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Kill(m_logsCreated.Item(i))
		Next i
	End Sub
	
	Private Function Datestamp(Optional ByRef TimeDate As Date = #12:00:00 AM#) As String
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, TimeDate, CDate("00:00:00 12/30/1899")) = 0) Then
			TimeDate = Now
		End If
		
		Datestamp = VB6.Format(TimeDate, "YYYY-MM-DD")
	End Function
	
	Private Function Timestamp(Optional ByRef TimeDate As Date = #12:00:00 AM#) As String
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, TimeDate, CDate("00:00:00 12/30/1899")) = 0) Then
			TimeDate = Now
		End If
		
		Select Case (BotVars.TSSetting)
			Case 0
				Timestamp = VB6.Format(TimeDate, "HH:MM:SS AM/PM")
				
			Case 1
				Timestamp = VB6.Format(TimeDate, "HH:MM:SS")
				
			Case 2
				Timestamp = VB6.Format(TimeDate, "HH:MM:SS") & "." & Right("000" & GetCurrentMS, 3)
				
			Case Else
				Timestamp = VB6.Format(TimeDate, "HH:MM:SS AM/PM")
		End Select
	End Function
End Class