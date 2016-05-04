Option Strict Off
Option Explicit On
Module modScripting
	'/* Scripting.bas
	' * ~~~~~~~~~~~~~~~~
	' * StealthBot VBScript support module
	' * ~~~~~~~~~~~~~~~~
	' * Modified by Swent 11/06/2007
	' * ~~~~~~~~~~~~~~~~
	' * Update Ribose/2009-08-15
	' */
	
	Public Structure scObj
		Dim SCModule As MSScriptControl.Module
		Dim ObjName As String
		Dim ObjType As String
		Dim obj As Object
	End Structure
	
	Public Structure scInc
		Dim SCModule As MSScriptControl.Module
		Dim IncName As String
		Dim lineCount As Short
	End Structure
	
	Private m_srootmenu As New clsMenuObj
	Private m_arrObjs() As scObj
	Private m_objCount As Short
	Private m_arrIncs() As scInc
	Private m_incCount As Short
	Private m_sc_control As MSScriptControl.ScriptControl
	Private m_is_reloading As Boolean
	Private m_ExecutingMdl As MSScriptControl.Module
	Private m_TempMdlName As String
	Private m_IsEventError As Boolean
	Private VetoNextMessage As Boolean
	Private VetoNextPacket As Boolean
	Private m_ScriptObservers As Collection
	Private m_FunctionObservers As Collection
	Private m_SystemDisabled As Boolean
	Private m_SCInitialized As Boolean
	
	Public Sub InitScriptControl(ByVal SC As MSScriptControl.ScriptControl)
		
		' check whether the override is disabling the script system
		If m_SystemDisabled Then Exit Sub
		
		m_is_reloading = True
		
		SC.Reset()
		
		'UPGRADE_ISSUE: VBControlExtender property INet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmChat.INet.Cancel()
		frmChat.scTimer.Enabled = False
		
		DestroyObjs()
		
		If Config.ScriptingAllowUI Then
			SC.AllowUI = True
		End If
		
		'// Create scripting objects
		SC.addObject("ssc", SharedScriptSupport, True)
		SC.addObject("scTimer", frmChat.scTimer)
		SC.addObject("scINet", frmChat.INet)
		SC.addObject("BotVars", BotVars)
		
		m_sc_control = SC
		
		m_ScriptObservers = New Collection
		m_FunctionObservers = New Collection
		
		m_is_reloading = False
		
		' this will always be true after first successful load
		m_SCInitialized = True
		
	End Sub
	
	Public Sub LoadScripts()
		
		On Error GoTo ERROR_HANDLER
		
		Dim CurrentModule As MSScriptControl.Module
		Dim Paths As New Collection
		Dim strPath As String
		Dim FileName As String
		Dim fileExt As String
		Dim i As Short
		Dim j As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim tmp As String
		Dim res As Boolean
		
		' check whether the override is disabling the script system
		If m_SystemDisabled Then Exit Sub
		
		' ********************************
		'      LOAD SCRIPTS
		' ********************************
		
		' set script folder path
		strPath = GetFolderPath("Scripts")
		
		' ensure scripts folder exists
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(strPath)) > 0) Then
            ' grab initial script file name
            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FileName = Dir(strPath)

            ' grab script files
            ' note: if we don't enumerate this list prior to script loading,
            ' scripting errors can kill further script loading.
            Do While (FileName <> vbNullString)
                ' add script file to collection
                Paths.Add(FileName)

                ' grab next script file name
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileName = Dir()
            Loop

            ' Cycle through each of the files.
            For i = 1 To Paths.Count()
                ' Does the file have the extension for a script?
                'UPGRADE_WARNING: Couldn't resolve default property of object Paths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GetFileExtension(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (IsValidFileExtension(GetFileExtension(Paths.Item(i)))) Then
                    ' Add a new module to the script control.
                    CurrentModule = m_sc_control.Modules.Add(CStr(m_sc_control.Modules.Count + 1))

                    ' store the temporary name of the module for parsing errors
                    'UPGRADE_WARNING: Couldn't resolve default property of object Paths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    m_TempMdlName = Paths.Item(i)

                    ' set executing module reference for parsing errors and module-specific functions
                    m_ExecutingMdl = CurrentModule

                    ' Load the file into the module.
                    'UPGRADE_WARNING: Couldn't resolve default property of object Paths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    res = FileToModule(CurrentModule, strPath & Paths.Item(i))

                    ' set executing module to nothing
                    'UPGRADE_NOTE: Object m_ExecutingMdl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    m_ExecutingMdl = Nothing

                    ' Does the script have a valid name?
                    If (IsScriptNameValid(CurrentModule) = False) Then
                        ' No. Try to fix it.
                        'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        GetScriptDictionary(CurrentModule)("Name") = CleanFileName(m_TempMdlName)

                        ' Is it valid now?
                        If (IsScriptNameValid(CurrentModule) = False) Then
                            ' No, disable it.

                            'UPGRADE_WARNING: Couldn't resolve default property of object Paths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            frmChat.AddChat(RTBColors.ErrorMessageText, "Scripting error: " & Paths.Item(i) & " has been " & "disabled due to a naming issue.")

                            str_Renamed = strPath & "\disabled\"

                            MkDir(str_Renamed)

                            'UPGRADE_WARNING: Couldn't resolve default property of object Paths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Kill(str_Renamed & Paths.Item(i))

                            'UPGRADE_WARNING: Couldn't resolve default property of object Paths(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Rename(strPath & Paths.Item(i), str_Renamed & Paths.Item(i))

                            InitScriptControl(m_sc_control)

                            LoadScripts()

                            Exit Sub
                        End If
                    End If
                End If
            Next i
        End If
		
		InitMenus()
		
		frmChat.AddChat(RTBColors.SuccessText, "Scripts loaded.")
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (" & Err.Number & "): " & Err.Description & " in LoadScripts().")
		
	End Sub
	
	Private Function FileToModule(ByRef ScriptModule As MSScriptControl.Module, ByVal FilePath As String, Optional ByVal defaults As Boolean = True) As Boolean
		
		On Error GoTo ERROR_HANDLER
		
		Static strContent As String
		Static includes As Collection
		
		Dim strLine As String
		Dim f As Short
		Dim blnCheckOperands As Boolean
		Dim blnKeepLine As Boolean
		Dim blnScriptData As Boolean
		Dim i As Short
		Dim lineCount As Short
		Dim blnIncIsValid As Boolean
		Dim strInclude As String
		
		blnCheckOperands = True
		
		f = FreeFile
		
		If (defaults) Then
			includes = New Collection
		End If
		
		FileOpen(f, FilePath, OpenMode.Input)
		Dim strCommand As String
		Dim strPath As String
		Dim strFullPath As String
		Do While (EOF(f) = False)
			strLine = LineInput(f)
			
			strLine = Trim(strLine)
			
			' default to keep line
			blnKeepLine = True
			
			If (Len(strLine) >= 1) Then
				If ((blnCheckOperands) And (Left(strLine, 1) = "#")) Then
					' this line is a directive, parse and don't keep in code
					blnKeepLine = False
					If (InStr(1, strLine, " ") <> 0) Then
						If (Len(strLine) >= 2) Then
							
							strCommand = LCase(Mid(strLine, 2, InStr(1, strLine, " ") - 2))
							
							If (strCommand = "include") Then
								If (Len(strLine) >= 12) Then
									
									strPath = Mid(strLine, 11, Len(strLine) - 11)
									
									If (Left(strPath, 1) = "\") Then
										'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										strFullPath = StringFormat("{0}\Scripts{1}", CurDir(), strPath)
									Else
										strFullPath = strPath
									End If
									
									blnIncIsValid = True
									
									' check if file exists to include
									'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                                    If Len(Dir(strFullPath)) = 0 Then
                                        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                                        frmChat.AddChat(RTBColors.ErrorMessageText, "Scripting warning: " & Dir(FilePath) & " is trying to include " & "a file that does not exist: " & strPath)
                                        blnIncIsValid = False
                                    End If
									
									' check if file is already included by this script
									For i = 1 To includes.Count()
										'UPGRADE_WARNING: Couldn't resolve default property of object includes(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If StrComp(includes.Item(i), strFullPath, CompareMethod.Text) = 0 Then
											'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
											frmChat.AddChat(RTBColors.ErrorMessageText, "Scripting warning: " & Dir(FilePath) & " is trying to include " & "a file that has already been included: " & strPath)
											blnIncIsValid = False
										End If
									Next i
									
									If blnIncIsValid Then
										' store the file path to an include
										includes.Add(strFullPath)
									End If
								End If
							End If
						End If
					End If
                ElseIf (((Len(strLine) > 0) And (StrComp(Left(strLine, 1), "'") <> 0)) Or ((Len(strLine) = 3) And (StrComp(strLine, "rem", CompareMethod.Text) <> 0)) Or ((Len(strLine) > 3) And (StrComp(Left(strLine, 3), "rem", CompareMethod.Text) <> 0) And (InStr(1, Mid(strLine, 4, 1), "abcdefghijklmnopqrstuvwxyz01243567890_", CompareMethod.Text) = 0))) Then
                    ' this line is not a comment or blank line, stop checking for #include
                    blnCheckOperands = False
				End If
			End If
			
			' if this line is not an #include
			If (blnKeepLine) Then
				' keep it, append it to content
				strContent = strContent & strLine & vbCrLf
			Else
				' remove it, but keep line numbers consistant in errors
				strContent = strContent & vbCrLf
			End If
			
			' increment counter
			lineCount = lineCount + 1
			
			' clean up
			strLine = vbNullString
		Loop 
		FileClose(f)
		
		' if we are not loading an include, set it up
		If (defaults) Then
			' store module-level functions
			ScriptModule.ExecuteStatement(GetDefaultModuleProcs())
			
			' initialize the variables globally
			ScriptModule.ExecuteStatement("Public Script, DataBuffer")
			
			' store Script dictionary into script
			SetScriptDictionary(ScriptModule)
			
			' create default DataBuffer object
			'UPGRADE_WARNING: Couldn't resolve default property of object ScriptModule.CodeObject.DataBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ScriptModule.CodeObject.DataBuffer = SharedScriptSupport.DataBufferEx()
			
			' set the Script("Path") value
			'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetScriptDictionary(ScriptModule)("Path") = FilePath
			
			' add this as the "first" include for this script, with no name so that later it is changed to "scriptname"
			AddInclude(ScriptModule, vbNullString, lineCount)
			
			' add includes to end of code, process the same for includes in includes
			i = 1
			Do While i <= includes.Count()
				FileToModule(ScriptModule, includes.Item(i), False)
				i = i + 1
			Loop 
			
			' store the position in the arrInc array where this script's line number information starts
			' the line number information will be in order of #include directives
			' the error handler will use this information to continually subtract linecounts until
			' the line number is in the bounds of the include information where the file actually is
			' so that we know what include is causing an error
			'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetScriptDictionary(ScriptModule)("IncludeLBound") = m_incCount - includes.Count() - 1
			
			' add content
			ScriptModule.AddCode(strContent)
			
			' clean up object
			'UPGRADE_NOTE: Object includes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			includes = Nothing
			
			' clean up content for next call to this function
			strContent = vbNullString
		Else
			' this is an #include, store our information
			AddInclude(ScriptModule, "#" & Mid(FilePath, InStrRev(FilePath, "\", -1, CompareMethod.Binary) + 1), lineCount)
		End If
		
		FileToModule = True
		
		Exit Function
		
ERROR_HANDLER: 
		
		strContent = vbNullString
		
		'UPGRADE_NOTE: Object includes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		includes = Nothing
		
		FileToModule = False
		
	End Function
	
	' stores information about an include for more accurate error handling!
	Private Sub AddInclude(ByRef SCModule As MSScriptControl.Module, ByVal Name As String, ByVal Lines As Short)
		Dim inc As scInc
		
		' redefine array size
		If (m_incCount) Then
			ReDim Preserve m_arrIncs(m_incCount)
		Else
			ReDim m_arrIncs(0)
		End If
		
		' store our #include name, path, and length
		inc.IncName = Name
		inc.lineCount = Lines
		
		' store our module handle
		inc.SCModule = SCModule
		
		' store our inc
		'UPGRADE_WARNING: Couldn't resolve default property of object m_arrIncs(m_incCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_arrIncs(m_incCount) = inc
		
		' increment include counter
		m_incCount = (m_incCount + 1)
	End Sub
	
	Private Function GetDefaultModuleProcs() As String
		
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String ' storage buffer for module code
		
		' GetModuleID() module-level function
		str_Renamed = str_Renamed & "Function GetModuleID()" & vbNewLine
		str_Renamed = str_Renamed & "   GetModuleID = SSC.GetModuleID(Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' GetScriptModule() module-level function
		str_Renamed = str_Renamed & "Function GetScriptModule()" & vbNewLine
		str_Renamed = str_Renamed & "   Set GetScriptModule = SSC.GetScriptModule(Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' GetWorkingDirectory() module-level function
		str_Renamed = str_Renamed & "Function GetWorkingDirectory()" & vbNewLine
		str_Renamed = str_Renamed & "   GetWorkingDirectory = SSC.GetWorkingDirectory(Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' CreateObj() module-level function
		str_Renamed = str_Renamed & "Function CreateObj(ObjType, ObjName)" & vbNewLine
		str_Renamed = str_Renamed & "   Set CreateObj = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.CreateObj(ObjType, ObjName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' DestroyObj() module-level function
		str_Renamed = str_Renamed & "Sub DestroyObj(ObjName)" & vbNewLine
		str_Renamed = str_Renamed & "   SSC.DestroyObj ObjName, Script(""Name"")" & vbNewLine
		str_Renamed = str_Renamed & "End Sub" & vbNewLine
		
		' GetObjByName() module-level function
		str_Renamed = str_Renamed & "Function GetObjByName(ObjName)" & vbNewLine
		str_Renamed = str_Renamed & "   Set GetObjByName = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.GetObjByName(ObjName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' GetSettingsEntry() module-level function
		str_Renamed = str_Renamed & "Function GetSettingsEntry(EntryName)" & vbNewLine
		str_Renamed = str_Renamed & "   GetSettingsEntry = SSC.GetSettingsEntry(EntryName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' IsEnabled() module-level function
		str_Renamed = str_Renamed & "Function IsEnabled()" & vbNewLine
		str_Renamed = str_Renamed & "   IsEnabled = (StrComp(GetSettingsEntry(""Enabled""), ""false"", vbTextCompare) <> 0)" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' WriteSettingsEntry() module-level function
		str_Renamed = str_Renamed & "Sub WriteSettingsEntry(EntryName, EntryValue)" & vbNewLine
		str_Renamed = str_Renamed & "   SSC.WriteSettingsEntry EntryName, EntryValue, , Script(""Name"")" & vbNewLine
		str_Renamed = str_Renamed & "End Sub" & vbNewLine
		
		' CreateCommand() module-level function
		str_Renamed = str_Renamed & "Function CreateCommand(commandName)" & vbNewLine
		str_Renamed = str_Renamed & "   Set CreateCommand = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.CreateCommand(commandName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' OpenCommand() module-level function
		str_Renamed = str_Renamed & "Function OpenCommand(commandName)" & vbNewLine
		str_Renamed = str_Renamed & "   Set OpenCommand = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.OpenCommand(commandName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' DeleteCommand() module-level function
		str_Renamed = str_Renamed & "Function DeleteCommand(commandName)" & vbNewLine
		str_Renamed = str_Renamed & "   Set DeleteCommand = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.DeleteCommand(commandName, Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' GetCommands() module-level function
		str_Renamed = str_Renamed & "Function GetCommands()" & vbNewLine
		str_Renamed = str_Renamed & "   Set GetCommands = _ " & vbNewLine
		str_Renamed = str_Renamed & "         SSC.GetCommands(Script(""Name""))" & vbNewLine
		str_Renamed = str_Renamed & "End Function" & vbNewLine
		
		' store module-level coding
		GetDefaultModuleProcs = str_Renamed
		
	End Function
	
	Private Function IsScriptNameValid(ByRef CurrentModule As MSScriptControl.Module) As Boolean
		
		On Error Resume Next
		
		Dim j As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim tmp As String
		Dim nameDisallow As String
		
		str_Renamed = GetScriptName((CurrentModule.Name))
		
		If (str_Renamed = vbNullString) Then
			IsScriptNameValid = False
			
			Exit Function
		End If
		
		nameDisallow = "\/:*?<>|"""
		
		For j = 1 To Len(str_Renamed)
			If (InStr(1, nameDisallow, Mid(str_Renamed, j, 1), CompareMethod.Text) <> 0) Then
				IsScriptNameValid = False
				
				Exit Function
			End If
		Next j
		
		For j = 2 To m_sc_control.Modules.Count
			If (m_sc_control.Modules(j).Name <> CurrentModule.Name) Then
				tmp = GetScriptName(CStr(j))
				
				If (StrComp(str_Renamed, tmp, CompareMethod.Text) = 0) Then
					IsScriptNameValid = False
					
					Exit Function
				End If
			End If
		Next j
		
		IsScriptNameValid = True
		
	End Function
	
	Public Sub InitScripts()
		
		On Error Resume Next
		
		Static reloading As Boolean
		
		Dim i As Short
		Dim tmp As String
		' check whether the override is disabling the script system
		If m_SystemDisabled Then Exit Sub
		
		If (reloading = False) Then
			RunInAll("Event_FirstRun")
			
			reloading = True
		End If
		
		For i = 2 To m_sc_control.Modules.Count
			If (i > 1) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_sc_control.Modules().CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tmp = m_sc_control.Modules(i).CodeObject.GetSettingsEntry("Enabled")
			End If
			
			If (StrComp(tmp, "False", CompareMethod.Text) <> 0) Then
				InitScript(m_sc_control.Modules(i))
			End If
		Next i
		
	End Sub
	
	Public Sub InitScript(ByVal SCModule As MSScriptControl.Module)
		
		On Error Resume Next
		
		Dim i As Short
		Dim startTime As Integer
		Dim finishTime As Integer
		
		startTime = GetTickCount()
		
		RunInSingle(SCModule, "Event_Load")
		
		finishTime = GetTickCount()
		
		'// 03/27/2009 52 - added default Script property for the load time
		'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetScriptDictionary(SCModule)("InitPerf") = (finishTime - startTime)
		
		'If (g_Online) Then
		'    RunInSingle SCModule, "Event_LoggedOn", GetCurrentUsername, BotVars.Product
		'    RunInSingle SCModule, "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags
		'
		'    If (g_Channel.Users.Count > 0) Then
		'        For i = 1 To g_Channel.Users.Count
		'            With g_Channel.Users(i)
		'                 RunInSingle SCModule, "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
		''                    .Game, False
		'            End With
		'         Next i
		'     End If
		'End If
	End Sub
	
	'UPGRADE_WARNING: ParamArray Parameters was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function RunInAll(ParamArray ByVal Parameters() As Object) As Boolean
		
		On Error Resume Next
		
		Dim SC As MSScriptControl.ScriptControl
		Dim i As Short
		Dim arr() As Object
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim oldVeto As Boolean
		Dim veto As Boolean
		Dim oldEM As MSScriptControl.Module
		Dim obj As MSScriptControl.Module
		Dim Proc As MSScriptControl.Procedure
		
		' check whether the override is disabling the script system
		If m_SystemDisabled Then Exit Function
		
		SC = m_sc_control
		
		If (m_is_reloading) Then
			Exit Function
		End If
		
		oldVeto = GetVeto 'Keeps the old veto, for recursion, and Sets to false.
		veto = False 'Just to be sure
		
		oldEM = m_ExecutingMdl ' keep last reference, for recursion
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Parameters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arr = Parameters
		
		For i = 2 To SC.Modules.Count
			obj = SC.Modules(i)
			m_ExecutingMdl = obj
			
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			str_Renamed = obj.CodeObject.GetSettingsEntry("Enabled")
			
			If (StrComp(str_Renamed, "False", CompareMethod.Text) <> 0) Then
				
				' check if module has the procedure before calling it!
				'For Each Proc In obj.Procedures
				'    If StrComp(Proc.Name, arr(0), vbTextCompare) = 0 Then
				'        ' it does, count args
				'        If Proc.numArgs = UBound(arr) Then
				'            ' call it
				CallByNameEx(obj, "Run", CallType.Method, arr)
				'        End If
				'        Exit For
				'    End If
				'Next Proc
				
				veto = veto Or GetVeto 'Did they veto it, or was it already vetod?
			End If
			'UPGRADE_NOTE: Object obj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			obj = Nothing
			'UPGRADE_NOTE: Object m_ExecutingMdl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_ExecutingMdl = Nothing
		Next 
		
		m_ExecutingMdl = oldEM ' return to last reference
		
		SetVeto(oldVeto) 'Reset the old veto, this is for recursion
		RunInAll = veto 'Was this particular event vetoed?
		
		' if this is the outermost of a recursion, make sure errors clear after this
		If m_ExecutingMdl Is Nothing Then
			m_sc_control.Error.Clear()
		End If
	End Function
	
	'UPGRADE_WARNING: ParamArray Parameters was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function RunInSingle(ByRef obj As MSScriptControl.Module, ParamArray ByVal Parameters() As Object) As Boolean
		
		On Error Resume Next
		
		Dim i As Short
		Dim x As Short
		Dim arr() As Object
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim oldVeto As Boolean
		Dim oldEM As MSScriptControl.Module
		Dim Proc As MSScriptControl.Procedure
		Dim sobsers As Collection
		Dim fobsers As Collection
		Dim obser As MSScriptControl.Module
		Dim cobser As Boolean
		Dim mname As String
		
		' check whether the override is disabling the script system
		If m_SystemDisabled Then Exit Function
		
		If (m_is_reloading) Then
			Exit Function
		End If
		
		oldVeto = GetVeto 'Keeps the old veto, for recursion, and Sets to false.
		
		oldEM = m_ExecutingMdl ' keep old module reference, for recursion
		
		m_ExecutingMdl = obj
		'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_IsEventError = (StrComp(arr(0), "Event_Error", CompareMethod.Binary) = 0)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Parameters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arr = Parameters
		
		'If Obj is nothing then we are 'running' an internal event
		'This is so scriptors can observe internal events, like Internal Commands
		If (obj Is Nothing) Then
			str_Renamed = "True"
			mname = vbNullString
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			str_Renamed = obj.CodeObject.GetSettingsEntry("Enabled")
			mname = GetScriptName((obj.Name))
		End If
		
		If (Not StrComp(str_Renamed, "False", CompareMethod.Text) = 0) Then
			If (Not obj Is Nothing) Then CallByNameEx(obj, "Run", CallType.Method, arr)
			RunInSingle = GetVeto 'Was this particular event vetoed?
			
			'Call any scripts that are observing this one
			sobsers = GetScriptObservers(mname, False)
			
			For i = 1 To sobsers.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object sobsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obser = GetModuleByName(sobsers.Item(i))
				If (Not obser Is Nothing) Then 'Is the script real/loaded?
					'UPGRADE_WARNING: Couldn't resolve default property of object obser.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					str_Renamed = obser.CodeObject.GetSettingsEntry("Enabled")
					If (Not StrComp(str_Renamed, "False", CompareMethod.Text) = 0) Then 'Is it off?
						m_ExecutingMdl = obser
						CallByNameEx(obser, "Run", CallType.Method, arr)
					End If
				End If
				'UPGRADE_NOTE: Object obser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				obser = Nothing
			Next i
			
			'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fobsers = GetFunctionObservers(CStr(arr(0)))
			
			For i = 1 To fobsers.Count()
				cobser = True
				'UPGRADE_WARNING: Couldn't resolve default property of object fobsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (StrComp(fobsers.Item(i), mname, CompareMethod.Text) = 0) Then 'Dont call itself
					cobser = False
				Else
					For x = 1 To sobsers.Count() 'See if we already called it with Script Observers
						'UPGRADE_WARNING: Couldn't resolve default property of object fobsers.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object sobsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (StrComp(sobsers.Item(x), fobsers.Item(i), CompareMethod.Text) = 0) Then
							cobser = False
							Exit For
						End If
					Next x
				End If
				
				If (cobser) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object fobsers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obser = GetModuleByName(fobsers.Item(i))
					If (Not obser Is Nothing) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object obser.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						str_Renamed = obser.CodeObject.GetSettingsEntry("Enabled")
						If (Not StrComp(str_Renamed, "False", CompareMethod.Text) = 0) Then
							m_ExecutingMdl = obser
							CallByNameEx(obser, "Run", CallType.Method, arr)
						End If
					End If
					'UPGRADE_NOTE: Object obser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					obser = Nothing
				End If
			Next i
		End If
		
		m_IsEventError = False
		m_ExecutingMdl = oldEM ' got back to old reference
		SetVeto(oldVeto) 'Reset the old veto, this is for recursion
		
		' if this is the outermost of a recursion, make sure errors clear after this
		If m_ExecutingMdl Is Nothing Then
			m_sc_control.Error.Clear()
		End If
	End Function
	
	Public Sub CallByNameEx(ByRef obj As Object, ByRef ProcName As String, ByRef CallType As CallType, Optional ByRef vArgsArray As Object = Nothing)
		
		On Error GoTo ERROR_HANDLER
		
		Dim oTLI As TLI.TLIApplication
		Dim ProcID As Integer
		Dim numArgs As Integer
		Dim i As Integer
		Dim v() As Object
		
		oTLI = New TLI.TLIApplication
		
		ProcID = oTLI.InvokeID(obj, ProcName)
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If (IsNothing(vArgsArray)) Then
			Call oTLI.InvokeHook(obj, ProcID, CallType)
		End If
		
		If (IsArray(vArgsArray)) Then
			numArgs = UBound(vArgsArray)
			
			ReDim v(numArgs)
			
			For i = 0 To numArgs
				' corrected object passing -Ribose/2009-08-10
				'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If IsReference(vArgsArray(numArgs - i)) Then
					v(i) = vArgsArray(numArgs - i)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object vArgsArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v(i) = vArgsArray(numArgs - i)
				End If
			Next i
			
			Call oTLI.InvokeHookArray(obj, ProcID, CallType, v)
		End If
		
		'UPGRADE_NOTE: Object oTLI may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oTLI = Nothing
		
		Exit Sub
		
ERROR_HANDLER: 
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmChat.SControl.Error. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((frmChat.SControl.Error)) Then
			Exit Sub
		End If
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in CallByNameEx().")
		
		'UPGRADE_NOTE: Object oTLI may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oTLI = Nothing
		
	End Sub
	
	Public Function Objects(ByRef objIndex As Short) As scObj
		
		Objects = m_arrObjs(objIndex)
		
	End Function
	
	Private Function ObjCount(Optional ByRef ObjType As String = "", Optional ByVal SCModule As MSScriptControl.Module = Nothing) As Short
		
		Dim i As Short
		
		If (ObjType <> vbNullString) Then
			For i = 0 To m_objCount - 1
				If (SCModule Is Nothing) Then
					If (StrComp(ObjType, m_arrObjs(i).ObjType, CompareMethod.Text) = 0) Then
						ObjCount = (ObjCount + 1)
					End If
				Else
					If (StrComp(SCModule.Name, m_arrObjs(i).SCModule.Name) = 0) Then
						If (StrComp(ObjType, m_arrObjs(i).ObjType, CompareMethod.Text) = 0) Then
							ObjCount = (ObjCount + 1)
						End If
					End If
				End If
			Next i
		Else
			ObjCount = m_objCount
		End If
		
	End Function
	
	Public Function CreateObj(ByRef SCModule As MSScriptControl.Module, ByVal ObjType As String, ByVal ObjName As String) As Object
		
		On Error Resume Next
		
		Dim obj As scObj
		Dim scriptName As String
		
		If SCModule Is Nothing Then Exit Function
		
		'UPGRADE_NOTE: Object CreateObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CreateObj = Nothing
		If (Not ValidObjectName(ObjName)) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "Scripting error: The variable name provided to CreateObj was not valid.")
			Exit Function
		End If
		
		' redefine array size & check for duplicate controls
		Dim i As Short ' loop counter variable
		If (m_objCount) Then
			
			For i = 0 To m_objCount - 1
				If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
					If (StrComp(m_arrObjs(i).ObjType, ObjType, CompareMethod.Text) = 0) Then
						If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
							CreateObj = m_arrObjs(i).obj
							
							Exit Function
						End If
					End If
				End If
			Next i
			
			ReDim Preserve m_arrObjs(m_objCount)
		Else
			ReDim m_arrObjs(0)
		End If
		
		' store our module name & type
		obj.ObjName = ObjName
		obj.ObjType = ObjType
		
		' store our module handle
		obj.SCModule = SCModule
		
		scriptName = GetScriptName((SCModule.Name))
		
		' grab/create instance of object
		Dim tmp As clsMenuObj
		Select Case (UCase(ObjType))
			Case "TIMER"
				If (ObjCount(ObjType) > 0) Then
					frmChat.tmrScript.Load(ObjCount(ObjType))
				End If
				
				obj.obj = frmChat.tmrScript(ObjCount(ObjType))
				
			Case "LONGTIMER"
				If (ObjCount(ObjType) > 0) Then
					frmChat.tmrScriptLong.Load(ObjCount(ObjType))
				End If
				
				obj.obj = New clsSLongTimer
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.tmr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.tmr = frmChat.tmrScriptLong(ObjCount(ObjType)).Enabled
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.tmr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				With obj.obj.tmr
					'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.tmr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Interval = 1000
				End With
				
			Case "WINSOCK"
				If (ObjCount(ObjType) > 0) Then
					frmChat.sckScript.Load(ObjCount(ObjType))
				End If
				
				obj.obj = frmChat.sckScript(ObjCount(ObjType))
				
			Case "INET"
				If (ObjCount(ObjType) > 0) Then
					frmChat.itcScript.Load(ObjCount(ObjType))
				End If
				
				obj.obj = frmChat.itcScript(ObjCount(ObjType))
				
			Case "FORM"
				obj.obj = New frmScript
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.setName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.setName(ObjName)
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.setSCModule. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.setSCModule(SCModule)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				HookWindowProc(obj.obj.hWnd)
				
			Case "MENU"
				' check if there are no menus and we're adding one
				If (ObjCount("Menu", SCModule) = 0) Then
					' get the dynamic menu dash
					tmp = DynamicMenus("dashmnu" & scriptName)
					' show it
					tmp.Visible = True
					tmp.Caption = "-"
				End If
				
				obj.obj = New clsMenuObj
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.Name = ObjName
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.Parent = DynamicMenus.Item("mnu" & scriptName)
				
				DynamicMenus.Add(obj.obj)
		End Select
		
		' store object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(m_objCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_arrObjs(m_objCount) = obj
		
		' increment object counter
		m_objCount = (m_objCount + 1)
		
		' create class variable for object
		SCModule.ExecuteStatement("Set " & ObjName & " = GetObjByName(" & Chr(34) & ObjName & Chr(34) & ")")
		
		' unfortunately creating a new form triggers the scripting events
		' too early, so we have to call them manually here.
		If (UCase(ObjType) = "FORM") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Initialize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.obj.Initialize()
		End If
		
		' return object
		CreateObj = obj.obj
		
	End Function
	
	Public Sub DestroyObjs(Optional ByVal SCModule As Object = Nothing)
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		For i = m_objCount - 1 To 0 Step -1
			If (SCModule Is Nothing) Then
				DestroyObj(m_arrObjs(i).SCModule, m_arrObjs(i).ObjName)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object SCModule.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (SCModule.Name = m_arrObjs(i).SCModule.Name) Then
					DestroyObj(m_arrObjs(i).SCModule, m_arrObjs(i).ObjName)
				End If
			End If
		Next i
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DestroyObjs().")
		
		Resume Next
		
	End Sub
	
	Public Sub DestroyObj(ByVal SCModule As MSScriptControl.Module, ByVal ObjName As String)
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		Dim Index As Short
		
		If SCModule Is Nothing Then Exit Sub
		
		If (m_objCount = 0) Then
			Exit Sub
		End If
		
		Index = m_objCount
		
		For i = 0 To m_objCount - 1
			If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
				If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
					Index = i
					
					Exit For
				End If
			End If
		Next i
		
		If (Index >= m_objCount) Then
			Exit Sub
		End If
		
		Dim IsForm As Boolean
		Dim tmp As clsMenuObj
		Select Case (UCase(m_arrObjs(Index).ObjType))
			Case "TIMER"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.tmrScript.Unload(m_arrObjs(Index).obj.Index)
				Else
					frmChat.tmrScript(0).Enabled = False
				End If
				
			Case "LONGTIMER"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.tmrScriptLong.Unload(m_arrObjs(Index).obj.Index)
				Else
					frmChat.tmrScriptLong(0).Enabled = False
				End If
				
			Case "WINSOCK"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.sckScript.Unload(m_arrObjs(Index).obj.Index)
				Else
					frmChat.sckScript(0).Close()
				End If
				
			Case "INET"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.itcScript.Unload(m_arrObjs(Index).obj.Index)
				Else
					'UPGRADE_ISSUE: VBControlExtender property itcScript.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					frmChat.itcScript(0).Cancel()
				End If
				
			Case "FORM"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				UnhookWindowProc(m_arrObjs(Index).obj.hWnd)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.DestroyObjs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_arrObjs(Index).obj.DestroyObjs()
				
				'UPGRADE_ISSUE: Unload m_arrObjs().obj was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(m_arrObjs(Index).obj)
				
			Case "MENU"
				' check if there is one menu left and we're destroying it
				If (ObjCount("Menu", SCModule) = 1) Then
					' default to false
					IsForm = False
					' get the menu
					tmp = m_arrObjs(Index).obj
					' get root menu
					Do While StrComp(Right(tmp.Name, 4), "ROOT", CompareMethod.Binary) <> 0
						' check if this has a window as a parent instead
						'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If StrComp(TypeName(tmp.Parent), "frmScript", CompareMethod.Binary) = 0 Then
							' yes, get outta here!
							IsForm = True
							Exit Do
						End If
						tmp = tmp.Parent
					Loop 
					' check again so that the dash doesn't get hidden
					If Not IsForm Then
						' get dynamic dash from caption of root
						tmp = DynamicMenus("dashmnu" & tmp.Caption)
						' show it
						tmp.Visible = False
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Class_Terminate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_arrObjs(Index).obj.Class_Terminate()
		End Select
		
		'UPGRADE_NOTE: Object m_arrObjs().obj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_arrObjs(Index).obj = Nothing
		
		If (Index < m_objCount - 1) Then
			For i = Index To ((m_objCount - 1) - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_arrObjs(i) = m_arrObjs(i + 1)
			Next i
		End If
		
		If (m_objCount > 1) Then
			ReDim Preserve m_arrObjs(m_objCount - 1)
		Else
			ReDim m_arrObjs(0)
		End If
		
		SCModule.ExecuteStatement("Set " & ObjName & " = Nothing")
		
		m_objCount = (m_objCount - 1)
		
		Exit Sub
		
ERROR_HANDLER: 
		
		' scripting engine has been reset - likely due to a reload
		If (Err.Number = -2147467259) Then
			Resume Next
		End If
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DestroyObj().")
		
		Resume Next
		
	End Sub
	
	Public Function GetObjByName(ByRef SCModule As MSScriptControl.Module, ByVal ObjName As String) As Object
		
		Dim i As Short
		
		If SCModule Is Nothing Then Exit Function
		
		For i = 0 To m_objCount - 1
			If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
				If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
					GetObjByName = m_arrObjs(i).obj
					
					Exit Function
				End If
			End If
		Next i
		
	End Function
	
	Public Function GetScriptObjByMenuID(ByVal MenuID As Integer) As scObj
		
		Dim i As Short
		Dim j As Short
		
		For i = 0 To ObjCount() - 1
			If (StrComp("Menu", Objects(i).ObjType, CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(i).obj.ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(i).obj.ID = MenuID) Then
					GetScriptObjByMenuID = m_arrObjs(i)
					
					Exit Function
				End If
				'ElseIf (StrComp("Form", Objects(I).ObjType, vbTextCompare) = 0) Then
				'    For j = 0 To Objects(I).obj.ObjCount("Menu") - 1
				'        If (StrComp("Menu", Objects(I).obj.Objects(j).ObjectType, vbTextCompare) = 0) Then
				'            If (Objects(I).obj.Objects(j).ID = MenuID) Then
				'                GetScriptObjByMenuID = Objects(I)
				'
				'                Exit Function
				'            End If
				'        End If
				'    Next j
			End If
		Next i
		
	End Function
	
	Public Function GetScriptObjByIndex(ByVal ObjType As String, ByVal Index As Short) As scObj
		
		Dim i As Short
		
		For i = 0 To ObjCount() - 1
			If (StrComp(ObjType, Objects(i).ObjType, CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(i).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(i).obj.Index = Index) Then
					GetScriptObjByIndex = m_arrObjs(i)
					
					Exit For
				End If
			End If
		Next i
		
	End Function
	
	Public Function InitMenus() As Object
		
		On Error GoTo ERROR_HANDLER
		
		Dim tmp As clsMenuObj
		Dim Name As String
		Dim i As Short
		
		' destroy all the menus and start over
		DestroyMenus()
		
		' for each script add menus
		For i = 2 To frmChat.SControl.Modules.Count
			If (i = 2) Then
				frmChat.mnuScriptingDash(0).Visible = True
			End If
			
			' name is the script name at this module index
			Name = GetScriptName(CStr(i))
			
			' new menu
			tmp = New clsMenuObj
			
			' root menu, give name, hwnd, and caption
			tmp.Name = Chr(0) & Name & Chr(0) & "ROOT"
			tmp.hWnd = GetSubMenu(GetMenu(frmChat.Handle.ToInt32), 5)
			tmp.Caption = Name
			
			' add with script name as key for adding more menus
			DynamicMenus.Add(tmp, "mnu" & Name)
			
			' new menu
			tmp = New clsMenuObj
			
			' enable/disable menu, give name, parent, and caption
			tmp.Name = Chr(0) & Name & Chr(0) & "ENABLE|DISABLE"
			'UPGRADE_WARNING: Couldn't resolve default property of object tmp.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tmp.Parent = DynamicMenus.Item("mnu" & Name)
			tmp.Caption = "Enabled"
			
			'UPGRADE_WARNING: Couldn't resolve default property of object GetModuleByName().CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), "False", CompareMethod.Text) <> 0) Then
				
				tmp.Checked = True
			End If
			
			' add it
			DynamicMenus.Add(tmp)
			
			' new menu
			tmp = New clsMenuObj
			
			' view script menu item, give name, parent, and caption
			tmp.Name = Chr(0) & Name & Chr(0) & "VIEW_SCRIPT"
			'UPGRADE_WARNING: Couldn't resolve default property of object tmp.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tmp.Parent = DynamicMenus.Item("mnu" & Name)
			tmp.Caption = "View Script"
			
			' add it
			DynamicMenus.Add(tmp)
			
			' new menu
			tmp = New clsMenuObj
			
			' dash, give name, parent, caption, and hide it
			tmp.Name = Chr(0) & Name & Chr(0) & "DASH"
			'UPGRADE_WARNING: Couldn't resolve default property of object tmp.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tmp.Parent = DynamicMenus.Item("mnu" & Name)
			tmp.Caption = "-"
			tmp.Visible = False
			
			' give it a key so it can't be confused with another script's main menu or menu dash
			DynamicMenus.Add(tmp, "dashmnu" & Name)
		Next i
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in InitMenus().")
		
		Err.Clear()
		
		Resume Next
		
	End Function
	
	Public Function DestroyMenus() As Object
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		frmChat.mnuScriptingDash(0).Visible = False
		
		For i = DynamicMenus.Count() To 1 Step -1
			
			'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Len(DynamicMenus.Item(i).Name) > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Left(DynamicMenus.Item(i).Name, 1) = Chr(0)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Class_Terminate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DynamicMenus.Item(i).Class_Terminate()
					
					DynamicMenus.Remove(i)
				End If
			End If
			
		Next i
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DestroyMenus().")
		
		Err.Clear()
		
		Resume Next
		
	End Function
	
	Public Function Scripts() As Object
		
		On Error Resume Next
		
		Dim i As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim SCModule As MSScriptControl.Module
		Dim scriptName As String
		
		Scripts = New Collection
		
		For i = 2 To frmChat.SControl.Modules.Count
			scriptName = GetScriptName(CStr(i))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptModule.CodeObject.GetSettingsEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			str_Renamed = GetScriptModule.CodeObject.GetSettingsEntry("Public")
			
			If (StrComp(str_Renamed, "False", CompareMethod.Text) <> 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Scripts.Add. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Scripts.Add(frmChat.SControl.Modules(i).CodeObject, scriptName)
			End If
		Next i
		
	End Function
	
	Public Function GetModuleByName(ByVal scriptName As String) As MSScriptControl.Module
		Dim i As Short
		
		For i = 2 To frmChat.SControl.Modules.Count
			If (StrComp(GetScriptName(CStr(i)), scriptName, CompareMethod.Text) = 0) Then
				GetModuleByName = frmChat.SControl.Modules(i)
				Exit Function
			End If
		Next i
	End Function
	
	Public Sub SetVeto(ByVal B As Boolean)
		
		VetoNextMessage = B
		
	End Sub
	
	Public Function GetVeto() As Boolean
		
		GetVeto = VetoNextMessage
		
		VetoNextMessage = False
		
	End Function
	
	Private Function GetFileExtension(ByVal FileName As String) As Object
		
		Dim arr() As String
		
		arr = Split(FileName, ".")
		
		If UBound(arr) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetFileExtension. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetFileExtension = ""
		Else
			GetFileExtension = arr(UBound(arr))
		End If
		
	End Function
	
	Private Function IsValidFileExtension(ByVal ext As String) As Boolean
		
		Dim exts() As String
		Dim i As Short
		
		ReDim exts(1)
		
		'exts(0) = "dat"
		exts(0) = "txt"
		exts(1) = "vbs"
		
		For i = LBound(exts) To UBound(exts)
			If (StrComp(ext, exts(i), CompareMethod.Text) = 0) Then
				IsValidFileExtension = True
				
				Exit Function
			End If
		Next i
		
		IsValidFileExtension = False
		
	End Function
	
	Private Function CleanFileName(ByVal FileName As String) As String
		
		On Error Resume Next
		
		If (InStr(1, FileName, ".") > 1) Then
			CleanFileName = Left(FileName, InStr(1, FileName, ".") - 1)
		End If
		
	End Function
	
	'06/26/09 - Hdx Very crappy function to check if Object names are valid, a-z0-9_ and 1st chr a-z (eventually should be a regexp)
	Public Function ValidObjectName(ByRef sName As String) As Boolean
		Dim x As Short
		Dim sValid As String
		
		sValid = "abcdefghijklmnopqrstuvwxyz0123456789_"
		ValidObjectName = False
		
		For x = 1 To Len(sName)
			If (InStr(1, Left(sValid, IIf(x = 1, 26, 37)), Mid(sName, x, 1), CompareMethod.Text) = 0) Then Exit Function
		Next x
		
		ValidObjectName = True
	End Function
	
	Public Function ConvertStringArray(ByRef sArr() As String) As Object()
		Dim vArr() As Object
		Dim x As Short
		'UPGRADE_WARNING: Lower bound of array vArr was changed from LBound(sArr) to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim vArr(UBound(sArr))
		For x = LBound(sArr) To UBound(sArr)
			'UPGRADE_WARNING: Couldn't resolve default property of object vArr(x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vArr(x) = CObj(sArr(x))
		Next x
		ConvertStringArray = VB6.CopyArray(vArr)
	End Function
	
	Public Sub SC_Error()
		
		Dim Name As String
		Dim ErrType As String
		Dim Number As Integer
		Dim description As String
		Dim line As Integer
		Dim Column As Integer
		Dim source As String
		Dim Text As String
		Dim IncIndex As Short
		Dim i As Short
		Dim tmp As String
		
		' check whether the override is disabling the script system
		' (this function is being called in that case due to /exec)
		If m_SystemDisabled Then Exit Sub
		
		With m_sc_control
			Number = .Error.Number
			description = .Error.description
			line = .Error.line
			Column = .Error.Column
			source = .Error.source
			Text = .Error.Text
		End With
		
		If m_ExecutingMdl Is Nothing Then
			Exit Sub ' exec error handler will handle this
		Else
			' start at the stored include index
			'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IncIndex = GetScriptDictionary(m_ExecutingMdl)("IncludeLBound")
			' loop until we have reached a line number within this include's bounds
			For i = IncIndex To UBound(m_arrIncs)
				If Not m_ExecutingMdl Is m_arrIncs(i).SCModule Then
					' we are no longer looping through this script's includes-- line number is too large or index is invalid
					Name = GetScriptName((m_ExecutingMdl.Name))
					Exit For
				ElseIf line <= m_arrIncs(i).lineCount Then 
					' it is this script which is erroring
					Name = GetScriptName((m_ExecutingMdl.Name)) & m_arrIncs(i).IncName
					Exit For
				Else
					' this include did not contain this line
					' decrease the line number the error was found by this include's line count
					line = line - m_arrIncs(i).lineCount
				End If
			Next i
            If Len(Name) = 0 Or Left(Name, 1) = "#" Then Name = m_TempMdlName & Name
		End If
		
		' default to runtime error
		ErrType = "runtime"
		
		' check if its a parsing error
		If InStr(1, source, "compilation", CompareMethod.Binary) > 0 Then
			ErrType = "parsing"
		End If
		
		' check if the script is planning to handle errors itself, if Script("HandleErrors") = True, then call event_error
		'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If ((StrComp(GetScriptDictionary(m_ExecutingMdl)("HandleErrors"), "True", CompareMethod.Text) = 0) And (m_IsEventError = False)) Then
			' call Event_Error(Number, Description, Line, Column, Text, Source)
			If (RunInSingle(m_ExecutingMdl, "Event_Error", Number, description, line, Column, Text, source) = True) Then
				' if vetoed, exit
				Exit Sub
			End If
		End If
		
		' display error if script enabled
		If InStr(Name, "#") > 0 Then
			tmp = Left(Name, InStr(Name, "#") - 1)
		Else
			tmp = Name
		End If
		tmp = SharedScriptSupport.GetSettingsEntry("Enabled", CleanFileName(tmp))
		
		If (StrComp(tmp, "False", CompareMethod.Text) <> 0) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Scripting {0} error '{1}' in {2}: (line {3}; column {4})", ErrType, Number, Name, line, Column))
			frmChat.AddChat(RTBColors.ErrorMessageText, description)
            If Len(Trim(Text)) > 0 Then
                frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Offending line: >> {0}", Text))
            End If
		End If
		
		m_sc_control.Error.Clear()
		
	End Sub
	
	' get currently executing module (can be nothing!)
	Public Function GetScriptModule(Optional ByVal scriptName As String = vbNullString) As MSScriptControl.Module
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		' only loop if the scriptname was provided
        If Len(scriptName) > 0 Then
            ' loop through modules
            For i = 2 To m_sc_control.Modules.Count
                ' if Script("Name") = ScriptName then
                If StrComp(GetScriptName(CStr(i)), scriptName, CompareMethod.Text) = 0 Then
                    ' return this module
                    GetScriptModule = m_sc_control.Modules(i)

                    Exit Function
                End If
            Next i
        Else
            ' return currently executing
            GetScriptModule = m_ExecutingMdl
        End If
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in GetScriptModule().")
		
		Err.Clear()
		
		Resume Next
		
	End Function
	
	' get the module.Name for the provided script name
	' returns the currently executing one if none provided
	' if the currently executing one is nothing (/exec), "exec" is returned
	Public Function GetModuleID(Optional ByVal scriptName As String = vbNullString) As String
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		' only loop if the scriptname was provided
        If Len(scriptName) > 0 Then
            ' loop through modules
            For i = 2 To m_sc_control.Modules.Count
                ' if Script("Name") = ScriptName then
                If StrComp(GetScriptName(CStr(i)), scriptName, CompareMethod.Text) = 0 Then
                    ' return this ID
                    GetModuleID = m_sc_control.Modules(i).Name
                    Exit Function
                End If
            Next i
        End If
		
		' if this is the execute command, the current module is nothing
		If m_ExecutingMdl Is Nothing Then
			GetModuleID = "exec"
		Else
			' return currently executing ID
			GetModuleID = m_ExecutingMdl.Name
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in GetModuleID().")
		
		Err.Clear()
		
		Resume Next
		
	End Function
	
	' returns the script name from the Script() object for the provided module id
	' if not found returns the module ID provided
	' if none provided, uses the currently executing module ID
	Public Function GetScriptName(Optional ByVal ModuleID As String = vbNullString) As String
		
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Module_Renamed As MSScriptControl.Module
		
		' if not provided
        If (Len(ModuleID) = 0) Then
            ' use currently executing module ID
            ModuleID = GetModuleID()
        End If
		
		If StrictIsNumeric(ModuleID) = False Then Exit Function
		
		Module_Renamed = m_sc_control.Modules(ModuleID)
		
		If Module_Renamed Is Nothing Then Exit Function
		
		' get Script() value "Name"
		'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetScriptName = GetScriptDictionary(Module_Renamed)("Name")
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in GetScriptName().")
		
		Err.Clear()
		
		Resume Next
		
	End Function
	
	'Adds a Observer/Observie pair to the ScriptObservers collection.
	'Observer\x0Observie
	'Checks for duplicates, Ignoring cases of corse.
	'Also prevents Observing yourself
	Public Sub AddScriptObserver(ByVal ModuleName As String, ByVal sTargetScript As String)
		On Error GoTo ERROR_HANDLER
		Dim i As Short
		Dim observed As Collection
		
		If (StrComp(ModuleName, sTargetScript, CompareMethod.Text) = 0) Then
			Exit Sub
		End If
		
		observed = GetScriptObservers(ModuleName)
		
		For i = 1 To observed.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object observed.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(observed.Item(i), sTargetScript, CompareMethod.Text) = 0) Then
				'UPGRADE_NOTE: Object observed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				observed = Nothing
				Exit Sub
			End If
		Next i
		'UPGRADE_NOTE: Object observed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		observed = Nothing
		
		m_ScriptObservers.Add(ModuleName & Chr(0) & sTargetScript)
		
		Exit Sub
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modScripting.AddScriptObserver()")
	End Sub
	
	'Returns a collection of scripts that the passed script is currently observing, or being observed by
	'Needs a better name...
	Public Function GetScriptObservers(ByRef sScriptName As String, Optional ByRef IsObserver As Boolean = True) As Collection
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		Dim sObserver As String
		Dim sObservie As String
		
		GetScriptObservers = New Collection
		
		For i = 1 To m_ScriptObservers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object m_ScriptObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (InStr(m_ScriptObservers.Item(i), Chr(0))) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_ScriptObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sObserver = Split(m_ScriptObservers.Item(i), Chr(0))(0)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_ScriptObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sObservie = Split(m_ScriptObservers.Item(i), Chr(0))(1)
				
				
				If (StrComp(sObserver, sScriptName, CompareMethod.Text) = 0 And IsObserver) Then
					GetScriptObservers.Add(sObservie)
				ElseIf (StrComp(sObservie, sScriptName, CompareMethod.Text) = 0 And IsObserver = False) Then 
					GetScriptObservers.Add(sObserver)
				End If
			End If
		Next i
		
		Exit Function
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modScripting.GetScriptObservers()")
	End Function
	
	'Adds a Function/Observer pair to the Function Observer collection
	'Observer\x00Function
	Public Sub AddFunctionObserver(ByVal ModuleName As String, ByVal sTargetFunction As String)
		On Error GoTo ERROR_HANDLER
		Dim i As Short
		Dim sItem As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sItem = StringFormat("{0}{1}{2}", ModuleName, Chr(0), sTargetFunction)
		
		For i = 1 To m_FunctionObservers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object m_FunctionObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(m_FunctionObservers.Item(i), sItem, CompareMethod.Text) = 0) Then
				Exit Sub
			End If
		Next i
		
		m_FunctionObservers.Add(sItem)
		
		Exit Sub
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in modScripting.AddFunctionObservers()", Err.Number, Err.Description))
	End Sub
	
	'Returns a collections of scripts who are observing this event in all scripts.
	Public Function GetFunctionObservers(ByRef sFunctionName As String) As Collection
		On Error GoTo ERROR_HANDLER
		Dim i As Short
		Dim sFunction As String
		Dim sObserver As String
		
		GetFunctionObservers = New Collection
		
		For i = 1 To m_FunctionObservers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object m_FunctionObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (InStr(m_FunctionObservers.Item(i), Chr(0))) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_FunctionObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sObserver = Split(m_FunctionObservers.Item(i), Chr(0))(0)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_FunctionObservers.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sFunction = Split(m_FunctionObservers.Item(i), Chr(0))(1)
				
				If (StrComp(sFunctionName, sFunction, CompareMethod.Text) = 0) Then
					GetFunctionObservers.Add(sObserver)
				End If
			End If
		Next i
		
		Exit Function
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in modScripting.GetFunctionObservers()", Err.Number, Err.Description))
	End Function
	
	' call this during config reload to enable/disable the system
	' (all calls to the script system will exit if true!!)
	' if this call is enabling the system, load scripts (init sc, init scripts, load scripts)
	' if this call is disabling the system, clean up scripts (close)
	Public Sub SetScriptSystemDisabled(ByVal SystemDisabled As Boolean)
		
		On Error GoTo ERROR_HANDLER
		
		If m_SystemDisabled = Not SystemDisabled Then
			If SystemDisabled Then
				' system is being disabled, close (if open)
				If m_SCInitialized Then
					'RunInAll "Event_LoggedOff"
					RunInAll("Event_Close")
				End If
				' hide scripting menu
				frmChat.mnuScripting.Visible = False
				' store the state
				m_SystemDisabled = True
			Else
				' store the state
				m_SystemDisabled = False
				' show scripting menu
				frmChat.mnuScripting.Visible = True
				' system is being enabled, open
				InitScriptControl(frmChat.SControl)
				LoadScripts()
				InitScripts()
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		
		' Cannot call this method while the script is executing
		If (Err.Number = -2147467259) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "Error: Script is still executing.")
			
			Exit Sub
		End If
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in SetScriptSystemDisabled().")
	End Sub
	
	Public Function GetScriptSystemDisabled() As Boolean
		
		GetScriptSystemDisabled = m_SystemDisabled
		
	End Function
	
	' call this function to get a script module's CodeObject.Script dictionary.
	' this will make sure that the CodeObject.Script is of type Dictionary and
	' can be accessed as such
	' if not, the CodeObject.Script object will be restored (but all data was lost)
	' prevents RTE due to CodeObject.Script(key) accesses failing when Script not
	' of type Dictionary ~Ribose
	Public Function GetScriptDictionary(ByRef mdl As MSScriptControl.Module) As Scripting.Dictionary
		'UPGRADE_WARNING: Couldn't resolve default property of object mdl.CodeObject.Script. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (StrComp(TypeName(mdl.CodeObject.Script), "Dictionary") <> 0) Then
			SetScriptDictionary(mdl)
			frmChat.AddChat(RTBColors.ErrorMessageText, "Scripting error: A Script object has been reset. " & "Script module ID: " & mdl.Name)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mdl.CodeObject.Script. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetScriptDictionary = mdl.CodeObject.Script
	End Function
	
	' call this function to set or reset the contents of the CodeObject.Script
	' dictionary for the specified module
	Private Function SetScriptDictionary(ByRef mdl As MSScriptControl.Module) As Object
		' let's store a dictionary into the specified CodeObject.Script
		Dim Dict As New Scripting.Dictionary
		' make Script object keys case-insensitive
		Dict.CompareMode = Scripting.CompareMethod.TextCompare
		' store it
		'UPGRADE_WARNING: Couldn't resolve default property of object mdl.CodeObject.Script. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mdl.CodeObject.Script = Dict
	End Function
End Module