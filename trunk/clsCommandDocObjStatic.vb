Option Strict Off
Option Explicit On
Friend Class clsCommandDocObjStatic
	
	
	
	'// 06/24/09 Hdx - Created
	'This will return a Collection that has all of the CommandsDoc objects for all of the commands from the selected file
	'UPGRADE_NOTE: clsCommandObj was upgraded to clsCommandObj_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetCommands(Optional ByVal scriptName As String = vbNullString) As Collection
		Dim clsCommandObj_Renamed As Object
		
		Dim AZ As String
		Dim xpath As String
		Dim doc As New MSXML2.DOMDocument60
		Dim commandNodes As MSXML2.IXMLDOMNodeList
		Dim commandNode As MSXML2.IXMLDOMNode
		
		AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		
		GetCommands = New Collection
		
		doc.Load(GetFilePath(FILE_COMMANDS))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		scriptName = clsCommandObj.CleanXPathVar(scriptName)
		
		If scriptName = vbNullString Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xpath = StringFormat("./commands/command[not(@owner) and (not(@enabled) or @enable = 1)]", UCase(AZ), LCase(AZ), LCase(scriptName))
		ElseIf scriptName = Chr(0) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xpath = StringFormat("./commands/command[translate(@owner, '{0}', '{1}')='{2}' and (not(@enabled) or @enable = 1)]", UCase(AZ), LCase(AZ), LCase(scriptName))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xpath = StringFormat("./commands/command[translate(@owner, '{0}', '{1}')='{2}']", UCase(AZ), LCase(AZ), LCase(scriptName))
		End If
		
		commandNodes = doc.selectNodes(xpath)
		
		For	Each commandNode In commandNodes
			GetCommands.Add(commandNode.Attributes.getNamedItem("name").nodeValue)
		Next commandNode
		
		'UPGRADE_NOTE: Object commandNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		commandNode = Nothing
		'UPGRADE_NOTE: Object commandNodes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		commandNodes = Nothing
		'UPGRADE_NOTE: Object doc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		doc = Nothing
		
	End Function
	
	
	Public Function ValidateXMLFromFiles(ByVal strXMLPath As String, ByVal strXSDPath As String) As Object
		
		Dim oFSO As Scripting.FileSystemObject
		Dim oTS As Scripting.TextStream
		Dim strXML, strXSD As String
		
		oFSO = New Scripting.FileSystemObject
		
		'// read the xml file
		oTS = oFSO.OpenTextFile(strXMLPath, Scripting.IOMode.ForReading, False)
		strXML = oTS.ReadAll()
		Call oTS.Close()
		
		'// read the xsd file
		oTS = oFSO.OpenTextFile(strXSDPath, Scripting.IOMode.ForReading, False)
		strXSD = oTS.ReadAll()
		Call oTS.Close()
		
		ValidateXMLFromFiles = ValidateXMLFromStrings(strXML, strXSD)
		
		'UPGRADE_NOTE: Object oFSO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFSO = Nothing
		'UPGRADE_NOTE: Object oTS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oTS = Nothing
		
	End Function
	
	
	'// http://www.nonhostile.com/howto-validate-xml-xsd-in-vb6.asp
	'// 08/31/2008 JSM - Created
	Public Function ValidateXMLFromStrings(ByVal strXML As String, ByVal strXSD As String) As Boolean
		
		On Error GoTo ERROR_HANDLER
		
		Dim objSchemas As MSXML2.XMLSchemaCache60
		Dim objXML As MSXML2.DOMDocument60
		Dim objXSD As MSXML2.DOMDocument60
		Dim objErr As MSXML2.IXMLDOMParseError
		
		' load XSD as DOM to populate in Schema Cache
		objXSD = New MSXML2.DOMDocument60
		
		objXSD.async = False
		objXSD.validateOnParse = False
		objXSD.resolveExternals = False
		
		If Not objXSD.loadXML(strXSD) Then
			Err.Raise(1, "Validate", "Load XSD failed: " & objXSD.parseError.Reason)
		End If
		
		' populate schema cache
		objSchemas = New MSXML2.XMLSchemaCache60
		
		' ERROR!
		objSchemas.Add("", objXSD)
		
		' load XML file (without validation - that comes later)
		objXML = New MSXML2.DOMDocument60
		
		objXML.async = False
		objXML.validateOnParse = False
		objXML.resolveExternals = False
		
		' load XML, without any validation
		If Not objXML.loadXML(strXML) Then
			Err.Raise(1, "Validate", "Load XML failed: " & objXML.parseError.Reason)
		End If
		
		' bind Schema Cache to DOM
		objXML.schemas = objSchemas
		
		' does this XML measure up?
		objErr = objXML.Validate()
		
		' any good?
		ValidateXMLFromStrings = (objErr.ErrorCode = 0)
		If objErr.ErrorCode <> 0 Then
			Err.Raise(1, "ValidateXML", "Error (#" & objErr.ErrorCode & ") on Line " & objErr.line & ": " & objErr.Reason)
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsCommandDocObjStatic.ValidateXMLFromStrings().")
		ValidateXMLFromStrings = False
		
	End Function
End Class