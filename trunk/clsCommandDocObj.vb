Option Strict Off
Option Explicit On
Friend Class clsCommandDocObj
	' clsCommandDocObj.cls
	' Copyright (C) 2007 Eric Evans
	
	
	
	
	Private m_database As MSXML2.DOMDocument60
	Private m_command_node As MSXML2.IXMLDOMNode
	
	Private m_aliases As Collection
	Private m_params As Collection
	Private m_name As String
	Private m_required_rank As Short
	Private m_required_flags As String
	Private m_description As String
	Private m_special_notes As String
	Private m_enabled As Boolean
	Private m_owner As String
	
	Private m_defaultXMLPath As String
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		m_database = New MSXML2.DOMDocument60
		
		m_defaultXMLPath = GetFilePath(FILE_COMMANDS)
		
		Call OpenDatabase( , True)
		
		'// 06/23/2009 JSM - initializing properties
		Call resetInstance()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object m_database may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_database = Nothing
		'UPGRADE_NOTE: Object m_params may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_params = Nothing
		'UPGRADE_NOTE: Object m_aliases may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_aliases = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub resetInstance()
		
		'UPGRADE_NOTE: Object m_command_node may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_command_node = Nothing
		'UPGRADE_NOTE: Object m_aliases may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_aliases = Nothing
		'UPGRADE_NOTE: Object m_params may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_params = Nothing
		m_name = vbNullString
		m_required_rank = -1
		m_required_flags = vbNullString
		m_description = vbNullString
		m_special_notes = vbNullString
		m_enabled = False
		m_owner = vbNullString
		
	End Sub
	
	Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString, Optional ByVal forceLoad As Boolean = False) As Object
		
		If (DatabasePath = vbNullString) Then
			DatabasePath = m_defaultXMLPath
		End If
		
		If Not FileExists(DatabasePath) And m_database.childNodes.length = 0 Then
			m_database.documentElement = m_database.createElement("commands")
		ElseIf forceLoad = True Or m_database.childNodes.length = 0 Then 
			m_database.Load(DatabasePath)
		End If
		
	End Function
	
	Private Function FileExists(ByRef FileName As String) As Boolean
		On Error GoTo ErrorHandler
		' get the attributes and ensure that it isn't a directory
		FileExists = (GetAttr(FileName) And FileAttribute.Directory) = 0
ErrorHandler: 
		' if an error occurs, this function returns False
	End Function
	
	
	'UPGRADE_NOTE: clsCommandObj was upgraded to clsCommandObj_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetCommandCount(Optional ByVal strScriptOwner As String = vbNullString) As Short
		Dim clsCommandObj_Renamed As Object
		
		On Error GoTo ERROR_HANDLER
		
		Dim nodes As MSXML2.IXMLDOMNodeList
		Dim xpath As String
		
		OpenDatabase()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strScriptOwner = clsCommandObj.CleanXPathVar(strScriptOwner)
		
		'// create xpath expression based on strScriptOwner
		If strScriptOwner = vbNullString Then
			xpath = "/commands/command[not(@owner)]"
		Else
			xpath = "/commands/command[@owner='" & strScriptOwner & "']"
		End If
		
		nodes = m_database.selectNodes(xpath)
		
		GetCommandCount = nodes.length
		
		'UPGRADE_NOTE: Object nodes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nodes = Nothing
		
		Exit Function
		
ERROR_HANDLER: 
		'UPGRADE_NOTE: Object nodes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nodes = Nothing
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.GetCommandCount().")
		
		Exit Function
		
	End Function
	
	'UPGRADE_NOTE: clsCommandObj was upgraded to clsCommandObj_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function GetCommandXPath(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional ByRef Enabled2 As Boolean = False) As String
		Dim clsCommandObj_Renamed As Object
		Dim AZ As String
		AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strCommand = clsCommandObj.CleanXPathVar(strCommand)
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strScriptOwner = clsCommandObj.CleanXPathVar(strScriptOwner)
		
		If strScriptOwner = vbNullString Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and not(@owner)]", UCase(AZ), LCase(AZ), LCase(strCommand))
		ElseIf strScriptOwner = Chr(0) Then 
			
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and ({3})]", UCase(AZ), LCase(AZ), LCase(strCommand), IIf(Enabled2, "@enabled = '1'", "not(@enabled)"))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and translate(@owner, '{0}', '{1}')='{3}']", UCase(AZ), LCase(AZ), LCase(strCommand), LCase(strScriptOwner))
		End If
		
	End Function
	
	'// this function will open a clsCommandDocObj and attempt to seek to the specified command
	Public Function OpenCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString) As Boolean
		
		Dim command_access_node As MSXML2.IXMLDOMNode
		Dim command_documentation As MSXML2.IXMLDOMNode
		Dim command_Parameters As MSXML2.IXMLDOMNodeList
		Dim command_aliases As MSXML2.IXMLDOMNodeList
		'UPGRADE_NOTE: Alias was upgraded to Alias_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Alias_Renamed As MSXML2.IXMLDOMNode
		Dim attrs As MSXML2.IXMLDOMAttribute
		Dim xpath As String
		
		Call OpenDatabase()
		Call resetInstance()
		
		If (m_database.documentElement Is Nothing) Then
			OpenCommand = False
			Exit Function
		End If
		
		xpath = GetCommandXPath(strCommand, strScriptOwner)
		
		m_command_node = m_database.selectSingleNode(xpath)
		If ((m_command_node Is Nothing) And (strScriptOwner = Chr(0))) Then
			xpath = GetCommandXPath(strCommand, strScriptOwner, True)
			m_command_node = m_database.selectSingleNode(xpath)
		End If
		
		If (m_command_node Is Nothing) Then
			'Lets try an alias
			xpath = GetCommandXPath(convertAlias(strCommand), strScriptOwner)
			m_command_node = m_database.documentElement.selectSingleNode(xpath)
			
			If ((m_command_node Is Nothing) And (strScriptOwner = Chr(0))) Then
				xpath = GetCommandXPath(convertAlias(strCommand), strScriptOwner, True)
				m_command_node = m_database.selectSingleNode(xpath)
			End If
			
			If (m_command_node Is Nothing) Then
				OpenCommand = False
				Exit Function
			End If
		End If
		
		m_aliases = getAliases(m_command_node)
		m_params = getParameters(m_command_node)
		
		OpenCommand = True
		
	End Function
	
	'// 06/13/2009 JSM - Created function
	'// 09/01/2009 JSM - Added optional parameter to control if this command should be autosaved
	Public Function CreateCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional ByRef bAutoSave As Boolean = True) As Boolean
		
		Dim l_command As MSXML2.IXMLDOMElement
		Dim l_aliases As MSXML2.IXMLDOMElement
		Dim l_documentation As MSXML2.IXMLDOMElement
		Dim l_description As MSXML2.IXMLDOMElement
		Dim l_Parameters As MSXML2.IXMLDOMElement
		Dim l_Access As MSXML2.IXMLDOMElement
		
		'UPGRADE_WARNING: Untranslated statement in CreateCommand. Please check source code.
		
		Call OpenDatabase()
		
		'// create elements
		l_command = m_database.createElement("command")
		l_aliases = m_database.createElement("aliases")
		l_documentation = m_database.createElement("documentation")
		l_description = m_database.createElement("description")
		l_Parameters = m_database.createElement("arguments")
		l_Access = m_database.createElement("access")
		
		'// set the command name
		Call l_command.setAttribute("name", strCommand)
		
		'// set the owner if necessary
		If strScriptOwner <> vbNullString Then
			Call l_command.setAttribute("owner", strScriptOwner)
		End If
		
		
		'// create heirarchy
		'UPGRADE_WARNING: Couldn't resolve default property of object l_aliases. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call l_command.appendChild(l_aliases)
		'UPGRADE_WARNING: Couldn't resolve default property of object l_documentation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call l_command.appendChild(l_documentation)
		'UPGRADE_WARNING: Couldn't resolve default property of object l_description. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call l_documentation.appendChild(l_description)
		'UPGRADE_WARNING: Couldn't resolve default property of object l_Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call l_command.appendChild(l_Parameters)
		'UPGRADE_WARNING: Couldn't resolve default property of object l_Access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call l_command.appendChild(l_Access)
		
		
		'// append it to the database
		'UPGRADE_WARNING: Couldn't resolve default property of object l_command. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call m_database.documentElement.appendChild(l_command)
		
		
	End Function
	
	Public Function NewParameter(ByRef argName As String, ByRef argIsOptional As Boolean, Optional ByRef argDataType As String = "String") As Object
		
		Dim Param As clsCommandParamsObj
		
		'// create a new parameter
		Param = New clsCommandParamsObj
		With Param
			.Name = argName
			.IsOptional = argIsOptional
			.datatype = argDataType
		End With
		
		NewParameter = Param
		
	End Function
	
	
	Public Function NewRestriction(ByRef argName As String, Optional ByRef argRequiredRank As Short = -1, Optional ByRef argRequiredFlags As String = vbNullString) As Object
		
		Dim res As clsCommandRestrictionObj
		
		'// create a new restriction
		res = New clsCommandRestrictionObj
		With res
			.Name = argName
			.RequiredRank = argRequiredRank
			.RequiredFlags = argRequiredFlags
		End With
		
		NewRestriction = res
		
	End Function
	
	
	Private Function GetXSD() As String
		Dim oFSO As Scripting.FileSystemObject
		Dim oTS As Scripting.TextStream
		Dim strXSD As String
		
		oFSO = New Scripting.FileSystemObject
		
		'// read the xsd file
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oTS = oFSO.OpenTextFile(StringFormat("{0}\Commands.xsd", My.Application.Info.DirectoryPath), Scripting.IOMode.ForReading, False)
		strXSD = oTS.ReadAll()
		Call oTS.Close()
		
		'UPGRADE_NOTE: Object oFSO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFSO = Nothing
		'UPGRADE_NOTE: Object oTS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oTS = Nothing
		
		GetXSD = strXSD
		
	End Function
	
	
	Private Sub ShowXMLErrors(ByRef colErrorList As Collection)
		
		Dim i As Short
		Dim Msg As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Msg = StringFormat("The following errors were detected with commands.xml...{0}{0}", vbNewLine)
		
		For i = 1 To colErrorList.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Msg = StringFormat("{0}{1}{2}", Msg, colErrorList.Item(i), vbNewLine)
		Next i
		
		frmChat.AddChat(RTBColors.ErrorMessageText, Msg)
		
	End Sub
	
	Public Function CommandsSanityCheck(ByRef doc As MSXML2.DOMDocument60, Optional ByRef colErrorList As Collection = Nothing) As Boolean
		
		'On Error GoTo ERROR_HANDLER
		
		colErrorList = New Collection
		
		'////////////////////////////
		'// COMMAND NODES
		'////////////////////////////
		Dim oCommands As MSXML2.IXMLDOMNodeList
		Dim oCommand As MSXML2.IXMLDOMNode
		Dim uniqueEnabledCommands As Scripting.Dictionary
		Dim uniqueCommandOwners As Scripting.Dictionary
		
		oCommands = doc.documentElement.selectNodes("/commands/command")
		uniqueEnabledCommands = New Scripting.Dictionary
		uniqueCommandOwners = New Scripting.Dictionary
		
		Dim commandName As String
		Dim ownerName As String
		Dim Enabled As String
		Dim oArguments As MSXML2.IXMLDOMNodeList
		Dim oArgument As MSXML2.IXMLDOMNode
		Dim uniqueEnabledArguments As Scripting.Dictionary
		Dim argumentName As String
		Dim datatype As String
		Dim oRestrictions As MSXML2.IXMLDOMNodeList
		Dim oRestriction As MSXML2.IXMLDOMNode
		Dim uniqueEnabledRestrictions As Scripting.Dictionary
		Dim restrictionName As String
		For	Each oCommand In oCommands
			Do 
				
				
				'// default our values
				commandName = oCommand.Attributes.getNamedItem("name").Text
				ownerName = vbNullString
				Enabled = "1"
				
				'// make sure name != ""
				If Len(commandName) = 0 Then
					colErrorList.Add("A command element is missing the name attribute.")
					Exit Do
				End If
				
				'// make sure owner != ""
				If Not (oCommand.Attributes.getNamedItem("owner") Is Nothing) Then
					ownerName = oCommand.Attributes.getNamedItem("owner").Text
					If Len(ownerName) = 0 Then
						colErrorList.Add("An owner attribute cannot be empty on command element.")
						Exit Do
					End If
				End If
				
				
				'// make sure enabled attribute is 0 or 1
				If Not (oCommand.Attributes.getNamedItem("enabled") Is Nothing) Then
					Enabled = oCommand.Attributes.getNamedItem("enabled").Text
					If Enabled <> "0" And Enabled <> "1" Then
						colErrorList.Add("If present, an enabled attribute must be equal to 0 or 1 on a command element.")
						Exit Do
					End If
				End If
				
				'// make sure only 1 command is enabled if there are similar names
				If Enabled = "1" Then
					If uniqueEnabledCommands.Exists(commandName) Then
						colErrorList.Add("Only 1 command element can have no enabled attribute or an enabled attribute equal to 1.")
						Exit Do
					End If
					uniqueEnabledCommands.Add(commandName, ownerName)
				End If
				
				'// make sure commands with the same name of separate owners
				If uniqueCommandOwners.Exists(commandName & "|" & ownerName) Then
					colErrorList.Add("Commands with equal name attributes must have different unique owner attributes.")
					Exit Do
				End If
				uniqueEnabledCommands.Add(commandName & "|" & ownerName, commandName & "|" & ownerName)
				
				'////////////////////////////
				'// ARGUEMENT NODES
				'////////////////////////////
				
				oArguments = oCommand.selectNodes("arguments/argument")
				uniqueEnabledArguments = New Scripting.Dictionary
				
				For	Each oArgument In oArguments
					Do 
						
						
						'// default our values
						argumentName = oArgument.Attributes.getNamedItem("name").Text
						
						'// make sure name != ""
						If Len(argumentName) = 0 Then
							colErrorList.Add("An argument element is missing the name attribute.")
							Exit Do
						End If
						
						'// make sure name attribute is unique
						If uniqueEnabledArguments.Exists(argumentName) Then
							colErrorList.Add("Argument elements for a command must have unique name attributes.")
							Exit Do
						End If
						uniqueEnabledArguments.Add(argumentName, commandName)
						
						'// make sure datatype attribute is string or word or numeric or number
						If Not (oArgument.Attributes.getNamedItem("datatype") Is Nothing) Then
							datatype = LCase(oArgument.Attributes.getNamedItem("datatype").Text)
							If datatype <> "string" And datatype <> "word" And datatype <> "numeric" And datatype <> "number" Then
								colErrorList.Add("If present, a datatype attribute must be equal to string, word, numeric, or number on a argument element.")
								Exit Do
							End If
						End If
						
						'// make sure match message is ok if present
						If Not (oArgument.selectSingleNode("match") Is Nothing) Then
							If Not (oArgument.selectSingleNode("match").Attributes.getNamedItem("message") Is Nothing) Then
								If Len(oArgument.selectSingleNode("match").Attributes.getNamedItem("message").Text) = 0 Then
									colErrorList.Add("If present, the message attribute of the match element for an argument must have a value.")
									Exit Do
								End If
							Else
								'// match element is present, but no message attribute
								colErrorList.Add("Match element of an argument must contain a message attribute.")
								Exit Do
							End If
							
						End If
						
						
						'////////////////////////////
						'// RESTRICTION NODES
						'////////////////////////////
						
						oRestrictions = oArgument.selectNodes("restrictions/restriction")
						uniqueEnabledRestrictions = New Scripting.Dictionary
						
						For	Each oRestriction In oRestrictions
							Do 
								
								
								'// default our values
								restrictionName = oRestriction.Attributes.getNamedItem("name").Text
								
								'// make sure name != ""
								If Len(restrictionName) = 0 Then
									colErrorList.Add("A restriction element is missing the name attribute.")
									Exit Do
								End If
								
								'// make sure name attribute is unique
								If uniqueEnabledRestrictions.Exists(restrictionName) Then
									colErrorList.Add("Restriction elements for an argument must have unique name attributes.")
									Exit Do
								End If
								uniqueEnabledRestrictions.Add(restrictionName, argumentName)
								
								Exit Do
							Loop 
						Next oRestriction
						
						Exit Do
					Loop 
				Next oArgument
				
				Exit Do
			Loop 
		Next oCommand
		
		If colErrorList.Count() > 0 Then
			ShowXMLErrors(colErrorList)
			CommandsSanityCheck = False
		Else
			CommandsSanityCheck = True
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObjStatic.CommandsSanityCheck().")
		
	End Function
	
	Public Function Save(Optional ByVal writeToFile As Boolean = True) As Boolean
		
		On Error GoTo ERROR_HANDLER
		
		Dim colErrorList As Collection
		
		setAliases(m_command_node, m_aliases)
		setParameters(m_command_node, m_params)
		
		If writeToFile = True Then
			
			'UPGRADE_WARNING: Untranslated statement in Save. Please check source code.
			
			'// 08/302009 52 - getting rid of the clsXML class since it didnt write valid XML
			Call writeDocToFile(GetFilePath(FILE_COMMANDS))
			
		End If
		
		Save = True
		
		Exit Function
		
ERROR_HANDLER: 
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.Save().")
		
	End Function
	
	
	
	Private Sub writeDocToFile(ByVal FilePath As String)
		
		On Error GoTo ERROR_HANDLER
		
		Dim reader As New MSXML2.SAXXMLReader60 '// create the SAX reader
		Dim writer As New MSXML2.MXXMLWriter60 '// create the XML writer
		
		Dim iFileNo As Short
		
		'// set properties on the XML writer
		writer.byteOrderMark = True
		writer.omitXMLDeclaration = True
		writer.indent = True
		
		'// set the XML writer to the SAX content handler
		reader.contentHandler = writer
		reader.dtdHandler = writer
		reader.ErrorHandler = writer
		reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)
		reader.putProperty("http://xml.org/sax/properties/declaration-handler", writer)
		
		'// parse the DOMDocument object
		reader.Parse(m_database)
		
		'// open the file for writing
		iFileNo = FreeFile
		FileOpen(iFileNo, FilePath, OpenMode.Output)
		
		PrintLine(iFileNo, Replace(writer.output, Chr(9), "  ",  ,  , CompareMethod.Binary))
		
		'// close the file
		FileClose(iFileNo)
		
		'UPGRADE_NOTE: Object reader may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		reader = Nothing
		'UPGRADE_NOTE: Object writer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		writer = Nothing
		
		Exit Sub
		
ERROR_HANDLER: 
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObj.writeDocToFile().")
		
	End Sub
	
	
	Public Sub Delete()
		
		If (m_command_node Is Nothing) Then
			Exit Sub
		End If
		
        If (Len(Me.Owner) = 0) Then
            frmChat.AddChat(RTBColors.ErrorMessageText, "Error: You can not delete an internal command.")
            Exit Sub
        End If
		
		m_command_node.parentNode.removeChild(m_command_node)
		
		Call Save()
		
	End Sub
	
	Private Function getParameters(ByRef commandNode As MSXML2.IXMLDOMNode) As Collection
		
		Dim parameter As MSXML2.IXMLDOMNode
		Dim Parameters As MSXML2.IXMLDOMNodeList
		Dim temp As clsCommandParamsObj
		
		getParameters = New Collection
		
		If (commandNode Is Nothing) Then
			Exit Function
		End If
		
		Parameters = commandNode.selectNodes("arguments/argument")
		
		For	Each parameter In Parameters
			temp = New clsCommandParamsObj
			
			On Error Resume Next
			With temp
				.Name = getName(parameter)
				.Restrictions = getRestrictions(parameter)
				.IsOptional = False
				If (Not parameter.Attributes.getNamedItem("optional") Is Nothing) Then
					If (parameter.Attributes.getNamedItem("optional").Text = "1") Then
						.IsOptional = True
					End If
				End If
				.datatype = "String"
				If (Not parameter.Attributes.getNamedItem("type") Is Nothing) Then
					.datatype = parameter.Attributes.getNamedItem("type").Text
				End If
				.MatchMessage = parameter.selectSingleNode("match").Attributes.getNamedItem("message").Text
				.MatchCaseSensitive = False
				If (parameter.selectSingleNode("match").Attributes.getNamedItem("case-sensitive").Text = "1") Then
					.MatchCaseSensitive = True
				End If
				.MatchError = Trim(parameter.selectSingleNode("error").Text)
				.description = Trim(parameter.selectSingleNode("documentation/description").Text)
				.SpecialNotes = Trim(parameter.selectSingleNode("documentation/specialnotes").Text)
			End With
			On Error GoTo 0
			
			getParameters.Add(temp)
			
			'UPGRADE_NOTE: Object temp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			temp = Nothing
		Next parameter
		'UPGRADE_NOTE: Object Parameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Parameters = Nothing
		
	End Function
	
	Private Function getRestrictions(ByRef parameterNode As MSXML2.IXMLDOMNode) As Collection
		
		Dim RestrictionNode As MSXML2.IXMLDOMNode
		Dim RestrictionsNode As MSXML2.IXMLDOMNode
		Dim temp As clsCommandRestrictionObj
		
		getRestrictions = New Collection
		
		If (parameterNode Is Nothing) Then
			Exit Function
		End If
		
		RestrictionsNode = parameterNode.selectSingleNode("restrictions")
		
		If (RestrictionsNode Is Nothing) Then
			Exit Function
		End If
		
		For	Each RestrictionNode In RestrictionsNode.selectNodes("restriction")
			temp = New clsCommandRestrictionObj
			With temp
				.Name = getName(RestrictionNode)
				.MatchMessage = getMatchMessage(RestrictionNode)
				If (Not RestrictionNode.selectSingleNode("error") Is Nothing) Then
					.MatchError = Trim(RestrictionNode.selectSingleNode("error").Text)
				End If
				.MatchCaseSensitive = getMatchCase(RestrictionNode)
				.RequiredFlags = getFlags(RestrictionNode)
				.RequiredRank = GetRank(RestrictionNode)
				.Fatal = GetFatal(RestrictionNode)
			End With
			getRestrictions.Add(temp)
			'UPGRADE_NOTE: Object temp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			temp = Nothing
		Next RestrictionNode
	End Function
	
	Private Function getAliases(ByRef AnyNode As MSXML2.IXMLDOMNode) As Collection
		
		'UPGRADE_NOTE: Alias was upgraded to Alias_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Alias_Renamed As MSXML2.IXMLDOMNode
		Dim aliases As MSXML2.IXMLDOMNodeList
		
		getAliases = New Collection
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		'// 09/03/2008 JSM - Modified code to use the <aliases> element
		aliases = AnyNode.selectNodes("aliases/alias")
		
		If ((aliases Is Nothing) = False) Then
			For	Each Alias_Renamed In aliases
				getAliases.Add(Alias_Renamed.Text)
			Next Alias_Renamed
		End If
		
	End Function
	
	Private Function setParameters(ByRef AnyNode As MSXML2.IXMLDOMNode, ByRef ParameterCol As Collection) As Collection
		
		Dim argumentNode As MSXML2.IXMLDOMNode
		Dim argumentsNode As MSXML2.IXMLDOMNode
		Dim argumentsOptionalNode As MSXML2.IXMLDOMNode
		Dim argumentsOptionalStringNode As MSXML2.IXMLDOMNode
		Dim argumentsOptionalNumberNode As MSXML2.IXMLDOMNode
		Dim RestrictionNode As MSXML2.IXMLDOMNode
		
		Dim newNode As MSXML2.IXMLDOMNode
		Dim node As MSXML2.IXMLDOMNode
		Dim parameter As clsCommandParamsObj
		Dim Restriction As clsCommandRestrictionObj
		Dim i As Short
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		'// get the arguments node
		argumentsNode = AnyNode.selectSingleNode("arguments")
		If Not (argumentsNode Is Nothing) Then
			For	Each node In argumentsNode.childNodes
				argumentsNode.removeChild(node)
			Next node
		End If
		
		
		
		
		'// add the the new arguments node if necessary
		If ParameterCol.Count() > 0 Then
			
			'// loop through the parameters
			For i = 1 To ParameterCol.Count()
				
				parameter = ParameterCol.Item(i)
				
				argumentNode = m_database.createNode("element", "argument", vbNullString)
				
				'// @name
				newNode = m_database.createNode("attribute", "name", vbNullString)
				newNode.Text = parameter.Name
				argumentNode.Attributes.setNamedItem(newNode)
				
				'// documentation
				newNode = m_database.createNode("element", "documentation", vbNullString)
				argumentNode.appendChild(newNode)
				
				'// create the match element if necessary
				If Len(parameter.MatchMessage) <> 0 Then
					'// add the match element
					argumentNode.appendChild(m_database.createNode("element", "match", vbNullString))
					
					'// match[@message]
					newNode = m_database.createNode("attribute", "message", vbNullString)
					newNode.Text = parameter.MatchMessage
					argumentNode.selectSingleNode("match").Attributes.setNamedItem(newNode)
					
					
					'// match[@case-sensitive]
					newNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
					newNode.Text = IIf(parameter.MatchCaseSensitive, "1", "0")
					argumentNode.selectSingleNode("match").Attributes.setNamedItem(newNode)
					
				End If
				
				'// create the error element if necessary
				If Len(parameter.MatchError) <> 0 Then
					'// add the error element
					argumentNode.appendChild(m_database.createNode("element", "error", vbNullString))
					argumentNode.selectSingleNode("error").Text = parameter.MatchError
				End If
				
				
				'// documentation/description
				newNode = m_database.createNode("element", "description", vbNullString)
				newNode.Text = parameter.description
				argumentNode.selectSingleNode("documentation").appendChild(newNode)
				
				'// documentation/specialnotes
				newNode = m_database.createNode("element", "specialnotes", vbNullString)
				newNode.Text = parameter.SpecialNotes
				argumentNode.selectSingleNode("documentation").appendChild(newNode)
				
				'// restrictions
				If parameter.Restrictions.Count() > 0 Then
					setRestrictions(argumentNode, (parameter.Restrictions))
				End If
				
				If parameter.IsOptional Then
					newNode = m_database.createNode("attribute", "optional", vbNullString)
					newNode.Text = "1"
					argumentNode.Attributes.setNamedItem(newNode)
				End If
				
				If (Not StrComp(parameter.datatype, "string", CompareMethod.Text) = 0) Then
					newNode = m_database.createNode("attribute", "type", vbNullString)
					newNode.Text = parameter.datatype
					argumentNode.Attributes.setNamedItem(newNode)
				End If
				
				argumentsNode.appendChild(argumentNode)
			Next 
		End If
	End Function
	
	Private Function setRestrictions(ByRef AnyNode As MSXML2.IXMLDOMNode, ByRef RestrictionCol As Collection) As Collection
		
		Dim RestrictionNode As MSXML2.IXMLDOMNode
		
		Dim newNode As MSXML2.IXMLDOMNode
		Dim node As MSXML2.IXMLDOMNode
		Dim Restriction As clsCommandRestrictionObj
		Dim i As Short
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		'// add our restrictions node
		AnyNode.appendChild(m_database.createNode("element", "restrictions", vbNullString))
		
		'// loop through argument restrictions
		For	Each Restriction In RestrictionCol
			
			'// create our restriction node
			RestrictionNode = m_database.createNode("element", "restriction", vbNullString)
			AnyNode.selectSingleNode("restrictions").appendChild(RestrictionNode)
			
			'// name attribute
			newNode = m_database.createNode("attribute", "name", vbNullString)
			newNode.Text = Restriction.Name
			RestrictionNode.Attributes.setNamedItem(newNode)
			
			If (Restriction.Fatal) Then
				newNode = m_database.createNode("attribute", "nonfatal", vbNullString)
				newNode.Text = IIf(Restriction.Fatal, "0", "1")
				RestrictionNode.Attributes.setNamedItem(newNode)
			End If
			
			'// create the match element if necessary
			If Restriction.MatchMessage <> "" Then
				'// add the match element
				RestrictionNode.appendChild(m_database.createNode("element", "match", vbNullString))
				
				'// set the message attribute
				newNode = m_database.createNode("attribute", "message", vbNullString)
				newNode.Text = Restriction.MatchMessage
				RestrictionNode.selectSingleNode("match").Attributes.setNamedItem(newNode)
				
				'// match[@message]
				newNode = m_database.createNode("attribute", "message", vbNullString)
				newNode.Text = Restriction.MatchMessage
				RestrictionNode.selectSingleNode("match").Attributes.setNamedItem(newNode)
				
				
				'// match[@case-sensitive]
				newNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
				newNode.Text = IIf(Restriction.MatchCaseSensitive, "1", "0")
				RestrictionNode.selectSingleNode("match").Attributes.setNamedItem(newNode)
				
				
			End If
			
			
			
			'// create the error element if necessary
			If Len(Restriction.MatchError) <> 0 Then
				'// add the error element
				RestrictionNode.appendChild(m_database.createNode("element", "error", vbNullString))
				RestrictionNode.selectSingleNode("error").Text = Restriction.MatchError
			End If
			
			
			
			'// create the access node
			RestrictionNode.appendChild(m_database.createNode("element", "access", vbNullString))
			
			'// create the rank noode
			RestrictionNode.selectSingleNode("access").appendChild(m_database.createNode("element", "rank", vbNullString))
			RestrictionNode.selectSingleNode("access/rank").Text = CStr(Restriction.RequiredRank)
			
			'// create the flags node
			RestrictionNode.selectSingleNode("access").appendChild(m_database.createNode("element", "flags", vbNullString))
			
			'// add a flag element for each flag
			For i = 1 To Len(Restriction.RequiredFlags)
				newNode = m_database.createNode("element", "flag", vbNullString)
				newNode.Text = Mid(Restriction.RequiredFlags, i, 1)
				RestrictionNode.selectSingleNode("access/flags").appendChild(newNode)
			Next i
			
		Next Restriction
		
		
	End Function
	
	
	
	
	Private Function setAliases(ByRef AnyNode As MSXML2.IXMLDOMNode, ByRef AliasCol As Collection) As Collection
		
		'UPGRADE_NOTE: Alias was upgraded to Alias_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Alias_Renamed As MSXML2.IXMLDOMNode
		Dim aliases As MSXML2.IXMLDOMNodeList
		Dim i As Short
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		'// 09/03/2008 JSM - Modified code to use the <aliases> element
		aliases = AnyNode.selectNodes("aliases/alias")
		
		If ((aliases Is Nothing) = False) Then
			For	Each Alias_Renamed In aliases
				'// 09/03/2008 JSM - Modified code to use the <aliases> element
				AnyNode.selectSingleNode("aliases").removeChild(Alias_Renamed)
			Next Alias_Renamed
		End If
		
		For i = 1 To AliasCol.Count()
			'// 09/03/2008 JSM - Modified code to use the <aliases> element
			Alias_Renamed = AnyNode.selectSingleNode("aliases").appendChild(m_database.createNode("element", "alias", vbNullString))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object AliasCol(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Alias_Renamed.Text = AliasCol.Item(i)
		Next 
		
		'Call Save
		
	End Function
	
	Private Function getName(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.Attributes.getNamedItem("name")
		If (temp Is Nothing) Then
			Exit Function
		End If
		getName = temp.Text
		
	End Function
	
	Private Function getOwner(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.Attributes.getNamedItem("owner")
		If (temp Is Nothing) Then
			Exit Function
		End If
		getOwner = temp.Text
	End Function
	
	Private Function getMatchMessage(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("match")
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		temp = temp.Attributes.getNamedItem("message")
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		getMatchMessage = temp.Text
		
	End Function
	
	Private Function getMatchCase(ByRef AnyNode As MSXML2.IXMLDOMNode) As Boolean
		
		Dim temp As MSXML2.IXMLDOMNode
		getMatchCase = False
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("match")
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		temp = temp.Attributes.getNamedItem("case-sensitive")
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		getMatchCase = (temp.Text = "1")
		
	End Function
	
	Private Function getEnabled(ByRef AnyNode As MSXML2.IXMLDOMNode) As Boolean
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		temp = AnyNode.Attributes.getNamedItem("enabled")
		getEnabled = True
		
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		If (temp.Text = "0") Then
			getEnabled = False
		End If
		
	End Function
	
	Private Function setEnabled(ByRef AnyNode As MSXML2.IXMLDOMNode, ByVal Enabled As Boolean) As Object
		
		Dim temp As MSXML2.IXMLDOMNode
		Dim attr As MSXML2.IXMLDOMAttribute
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.Attributes.getNamedItem("enabled")
		If (temp Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_database.createAttribute(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp = AnyNode.Attributes.setNamedItem(m_database.createAttribute("enabled"))
		End If
		
		If (Enabled = True) Then
			temp.Text = "1"
		Else
			temp.Text = "0"
		End If
		
		'Call Save
		
	End Function
	
	Private Function getDescription(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("documentation/description")
		If (temp Is Nothing) Then
			Exit Function
		End If
		getDescription = temp.Text
		
	End Function
	
	'Syntax: .add <username> [rank] [attributes]
	Private Function getSyntaxString(Optional ByRef IsLocal As Boolean = False) As String
		
		Dim retVal As String
		Dim P As clsCommandParamsObj
		
		If (m_command_node Is Nothing) Then
			Exit Function
		End If
		
		'// add the command name
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		retVal = StringFormat("{0} ", Me.Name)
		
		'// add the parameters
		For	Each P In Me.Parameters
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			retVal = StringFormat("{0}{2}{1}{3} ", retVal, P.Name, IIf(P.IsOptional, "[", "<"), IIf(P.IsOptional, "]", ">"))
		Next P
		
		'// add the trigger
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		retVal = StringFormat("{1}{0}", retVal, IIf((Me.RequiredRank <> -1) Or (Len(Me.RequiredFlags) > 0), IIf(IsLocal, "/", BotVars.Trigger), "/"))
		
		getSyntaxString = Trim(retVal)
		
	End Function
	
	'alias1, alias2, alias3
	Private Function getAliasString() As Object
		Dim retVal As String
		Dim i As Short
		If (aliases.Count() > 0) Then
			For i = 1 To aliases.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object aliases.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retVal = StringFormat("{0}{1}{2}", retVal, CStr(aliases.Item(i)), IIf(i = aliases.Count(), vbNullString, ", "))
			Next i
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object getAliasString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getAliasString = retVal
	End Function
	
	
	Private Function getRequirementsString(Optional ByRef bShort As Boolean = False) As Object
		
		Dim retVal As String
		
		If (m_command_node Is Nothing) Then
			Exit Function
		End If
		
		If (Me.RequiredRank <> -1) Or (Len(Me.RequiredFlags) > 0) Then
			'// available outside the console
			retVal = ""
			
			'// add access requirements if necessary
			If (Me.RequiredRank <> -1) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retVal = StringFormat("{0} Requires rank {1}", retVal, Me.RequiredRank)
				If (Len(Me.RequiredFlags) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					retVal = StringFormat("{0}.", retVal)
				End If
			End If
			
			'// add flag requirements, if necessary, taking into account the existinance of an access requirements
			If (Len(Me.RequiredFlags) <> 0) Then
				If (Me.RequiredRank <> -1) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					retVal = StringFormat("{0} or one of the following flags: {1}.", retVal, Me.RequiredFlags)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					retVal = StringFormat("{0} Requires one of the following flags: {1}.", retVal, Me.RequiredFlags)
				End If
			End If
			
		Else
			'// console only, lets show the generic blurb
			retVal = "Command is only available to the console."
			If (Not bShort) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retVal = StringFormat("{0} To allow other users to use this command, set the command's flags requirement or rank requirement.", retVal)
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object getRequirementsString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getRequirementsString = retVal
		
	End Function
	
	
	
	Private Function setDescription(ByRef AnyNode As MSXML2.IXMLDOMNode, ByVal description As String) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("documentation/description")
		If (temp Is Nothing) Then
			temp = AnyNode.selectSingleNode("documentation")
			If (temp Is Nothing) Then
				temp = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
			End If
			temp = temp.appendChild(m_database.createNode("element", "description", vbNullString))
		End If
		
		temp.Text = description
		
		'Call Save
		
	End Function
	
	Private Function getNotes(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("documentation/specialnotes")
		If (temp Is Nothing) Then
			Exit Function
		End If
		getNotes = temp.Text
		
	End Function
	
	Private Function setNotes(ByRef AnyNode As MSXML2.IXMLDOMNode, ByVal Notes As String) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("documentation/specialnotes")
		If (temp Is Nothing) Then
			temp = AnyNode.selectSingleNode("documentation")
			If (temp Is Nothing) Then
				temp = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
			End If
			
			temp = temp.appendChild(m_database.createNode("element", "specialnotes", vbNullString))
		End If
		
		temp.Text = Notes
		
		'Call Save
		
	End Function
	
	Private Function GetRank(ByRef AnyNode As MSXML2.IXMLDOMNode) As Short
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("access/rank")
		
		If (temp Is Nothing) Then
			GetRank = -1
			
			Exit Function
		End If
		
		If (temp.Text = vbNullString) Then
			GetRank = -1
			
			Exit Function
		End If
		
		GetRank = CShort(Val(temp.Text))
		
	End Function
	
	Private Function GetFatal(ByRef AnyNode As MSXML2.IXMLDOMNode) As Boolean
		
		Dim temp As MSXML2.IXMLDOMNode
		GetFatal = True
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.Attributes.getNamedItem("nonfatal")
		If (temp Is Nothing) Then
			Exit Function
		End If
		GetFatal = (Not temp.Text = "1")
	End Function
	
	Private Function setRank(ByRef AnyNode As MSXML2.IXMLDOMNode, ByVal Rank As Short) As Short
		
		Dim temp As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("access/rank")
		
		If (temp Is Nothing) Then
			temp = AnyNode.selectSingleNode("access")
			
			If (temp Is Nothing) Then
				temp = AnyNode.appendChild(m_database.createNode("element", "access", vbNullString))
			End If
			
			temp = temp.appendChild(m_database.createNode("element", "rank", vbNullString))
		End If
		
		temp.Text = CStr(Rank)
		
		'Call Save
		
	End Function
	
	Private Function getFlags(ByRef AnyNode As MSXML2.IXMLDOMNode) As String
		
		Dim temp As MSXML2.IXMLDOMNodeList
		Dim Flag As MSXML2.IXMLDOMNode
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectNodes("access/flags/flag")
		
		If (temp Is Nothing) Then
			Exit Function
		End If
		
		For	Each Flag In temp
			getFlags = getFlags & Flag.Text
		Next Flag
		
	End Function
	
	Private Function setFlags(ByRef AnyNode As MSXML2.IXMLDOMNode, ByVal Flags As String) As String
		
		Dim temp As MSXML2.IXMLDOMNode
		Dim Flag As MSXML2.IXMLDOMNode
		Dim i As Short
		
		If (AnyNode Is Nothing) Then
			Exit Function
		End If
		
		temp = AnyNode.selectSingleNode("access/flags")
		
		If (temp Is Nothing) Then
			temp = AnyNode.selectSingleNode("access")
			
			temp = temp.appendChild(m_database.createNode("element", "flags", vbNullString))
		End If
		
		For i = temp.childNodes.length - 1 To 0 Step -1
			temp.removeChild(temp.childNodes(i))
		Next i
		
		For i = 1 To Len(Flags)
			Flag = temp.appendChild(m_database.createNode("element", "flag", vbNullString))
			
			Flag.Text = Mid(Flags, i, 1)
		Next i
		
		'Call Save
		
	End Function
	
	
	Public Property Name() As String
		Get
			
			Name = getName(m_command_node)
			
		End Get
		Set(ByVal Value As String)
			
			m_name = Value
			
		End Set
	End Property
	
	
	Public Property Owner() As String
		Get
			
			Owner = getOwner(m_command_node)
			
		End Get
		Set(ByVal Value As String)
			
			m_owner = Value
			
		End Set
	End Property
	
	Public ReadOnly Property aliases() As Collection
		Get
			
			If (m_aliases Is Nothing) Then
				m_aliases = getAliases(m_command_node)
			End If
			
			aliases = m_aliases
			
		End Get
	End Property
	
	
	Public Property IsEnabled() As Boolean
		Get
			
			IsEnabled = getEnabled(m_command_node)
			
		End Get
		Set(ByVal Value As Boolean)
			
			setEnabled(m_command_node, Value)
			
		End Set
	End Property
	
	
	Public Property RequiredRank() As Short
		Get
			
			RequiredRank = GetRank(m_command_node)
			
		End Get
		Set(ByVal Value As Short)
			
			setRank(m_command_node, Value)
			
		End Set
	End Property
	
	
	Public Property RequiredFlags() As String
		Get
			
			RequiredFlags = getFlags(m_command_node)
			
		End Get
		Set(ByVal Value As String)
			
			setFlags(m_command_node, Value)
			
		End Set
	End Property
	
	
	Public Property description() As String
		Get
			
			description = getDescription(m_command_node)
			
		End Get
		Set(ByVal Value As String)
			
			setDescription(m_command_node, Value)
			
		End Set
	End Property
	
	Public ReadOnly Property AliasString() As String
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object getAliasString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AliasString = getAliasString()
			
		End Get
	End Property
	
	Public ReadOnly Property RequirementsString() As String
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object getRequirementsString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RequirementsString = getRequirementsString(False)
			
		End Get
	End Property
	
	Public ReadOnly Property RequirementsStringShort() As Object
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object getRequirementsString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RequirementsStringShort = getRequirementsString(True)
			
		End Get
	End Property
	
	
	Public Property SpecialNotes() As String
		Get
			
			SpecialNotes = getNotes(m_command_node)
			
		End Get
		Set(ByVal Value As String)
			
			setNotes(m_command_node, Value)
			
		End Set
	End Property
	
	
	Public ReadOnly Property Parameters() As Collection
		Get
			
			If (m_params Is Nothing) Then
				m_params = getParameters(m_command_node)
			End If
			
			Parameters = m_params
			
		End Get
	End Property
	
	
	Public ReadOnly Property XMLDocument() As MSXML2.DOMDocument60
		Get
			XMLDocument = m_database
		End Get
	End Property
	
	Public Function SyntaxString(Optional ByRef IsLocal As Boolean = False) As String
		
		SyntaxString = getSyntaxString(IsLocal)
		
	End Function
	
	Public Function GetParameterByName(ByVal sParamName As String) As clsCommandParamsObj
		
		Dim P As clsCommandParamsObj
		Dim col As Collection
		Dim i As Short
		
		col = Me.Parameters
		
		For i = 1 To col.Count()
			P = col.Item(i)
			If StrComp(sParamName, P.Name, CompareMethod.Text) = 0 Then
				GetParameterByName = P
				Exit Function
			End If
		Next i
		
	End Function
End Class