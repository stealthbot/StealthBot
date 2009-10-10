VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandDocObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCommandDocObj.cls
' Copyright (C) 2007 Eric Evans


Option Explicit


Private m_database       As DOMDocument60
Private m_command_node   As IXMLDOMNode

Private m_aliases        As Collection
Private m_params         As Collection
Private m_name           As String
Private m_required_rank  As Integer
Private m_required_flags As String
Private m_description    As String
Private m_special_notes  As String
Private m_enabled        As Boolean
Private m_owner          As String

Private m_defaultXMLPath As String

Private Sub Class_Initialize()

    Set m_database = New DOMDocument60
    
    m_defaultXMLPath = GetFilePath("Commands.xml")
    
    Call OpenDatabase(, True)
    
    '// 06/23/2009 JSM - initializing properties
    Call resetInstance
    
End Sub

Private Sub Class_Terminate()

    Set m_database = Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Private Sub resetInstance()
        
    Set m_command_node = Nothing
    Set m_aliases = Nothing
    Set m_params = Nothing
    Let m_name = vbNullString
    Let m_required_rank = -1
    Let m_required_flags = vbNullString
    Let m_description = vbNullString
    Let m_special_notes = vbNullString
    Let m_enabled = False
    Let m_owner = vbNullString

End Sub

Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString, Optional ByVal forceLoad As Boolean = False)
    
    If (DatabasePath = vbNullString) Then
        DatabasePath = m_defaultXMLPath
    End If
    
    If Not FileExists(DatabasePath) And m_database.childNodes.length = 0 Then
        Set m_database.documentElement = m_database.createElement("commands")
    ElseIf forceLoad = True Or m_database.childNodes.length = 0 Then
        m_database.Load DatabasePath
    End If
    
End Function

Private Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function


Public Function GetCommandCount(Optional ByVal strScriptOwner As String = vbNullString) As Integer

    On Error GoTo ERROR_HANDLER
    
    Dim nodes                 As IXMLDOMNodeList
    Dim xpath                 As String
    
    
    OpenDatabase
    
    strScriptOwner = Replace$(strScriptOwner, "'", "&apos;")
    
    
    '// create xpath expression based on strScriptOwner
    If strScriptOwner = vbNullString Then
        xpath = "/commands/command[not(@owner)]"
    Else
        xpath = "/commands/command[@owner='" & strScriptOwner & "']"
    End If
    
    Set nodes = m_database.selectNodes(xpath)
    
    GetCommandCount = nodes.length
    
    Set nodes = Nothing
    
    Exit Function

ERROR_HANDLER:
    Set nodes = Nothing
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.GetCommandCount().")

    Exit Function
    
End Function

Private Function GetCommandXPath(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional Enabled2 As Boolean = False) As String
    Dim AZ As String
    AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    strCommand = Replace$(strCommand, "'", "&apos;")
    strScriptOwner = Replace$(strScriptOwner, "'", "&apos;")
    
    If strScriptOwner = vbNullString Then
        GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and not(@owner)]", _
            UCase$(AZ), LCase$(AZ), LCase$(strCommand))
    ElseIf strScriptOwner = Chr$(0) Then
        
        GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and ({3})]", _
            UCase$(AZ), LCase$(AZ), LCase$(strCommand), _
            IIf(Enabled2, "@enabled = '1'", "not(@enabled)"))
    Else
        GetCommandXPath = StringFormat("/commands/command[translate(@name, '{0}', '{1}')='{2}' and translate(@owner, '{0}', '{1}')='{3}']", _
            UCase$(AZ), LCase$(AZ), LCase$(strCommand), LCase$(strScriptOwner))
    End If
    
End Function

'// this function will open a clsCommandDocObj and attempt to seek to the specified command
Public Function OpenCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString) As Boolean
    
    Dim command_access_node   As IXMLDOMNode
    Dim command_documentation As IXMLDOMNode
    Dim command_Parameters    As IXMLDOMNodeList
    Dim command_aliases       As IXMLDOMNodeList
    Dim Alias                 As IXMLDOMNode
    Dim attrs                 As IXMLDOMAttribute
    Dim xpath                 As String
    
    Call OpenDatabase
    Call resetInstance

    If (m_database.documentElement Is Nothing) Then
        OpenCommand = False
        Exit Function
    End If
    
    xpath = GetCommandXPath(strCommand, strScriptOwner)
    
    Set m_command_node = m_database.selectSingleNode(xpath)
    If ((m_command_node Is Nothing) And (strScriptOwner = Chr$(0))) Then
        xpath = GetCommandXPath(strCommand, strScriptOwner, True)
        Set m_command_node = m_database.selectSingleNode(xpath)
    End If
    
    If (m_command_node Is Nothing) Then
        'Lets try an alias
        xpath = GetCommandXPath(convertAlias(strCommand), strScriptOwner)
        Set m_command_node = m_database.documentElement.selectSingleNode(xpath)
            
        If ((m_command_node Is Nothing) And (strScriptOwner = Chr$(0))) Then
            xpath = GetCommandXPath(convertAlias(strCommand), strScriptOwner, True)
            Set m_command_node = m_database.selectSingleNode(xpath)
        End If
        
        If (m_command_node Is Nothing) Then
            OpenCommand = False
            Exit Function
        End If
    End If
    
    Set m_aliases = getAliases(m_command_node)
    Set m_params = getParameters(m_command_node)
    
    OpenCommand = True
    
End Function

'// 06/13/2009 JSM - Created function
'// 09/01/2009 JSM - Added optional parameter to control if this command should be autosaved
Public Function CreateCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional bAutoSave As Boolean = True) As Boolean

    Dim l_command             As IXMLDOMElement
    Dim l_aliases             As IXMLDOMElement
    Dim l_documentation       As IXMLDOMElement
    Dim l_description         As IXMLDOMElement
    Dim l_Parameters          As IXMLDOMElement
    Dim l_Access              As IXMLDOMElement
        
    Call OpenDatabase
    
    '// create elements
    Set l_command = m_database.createElement("command")
    Set l_aliases = m_database.createElement("aliases")
    Set l_documentation = m_database.createElement("documentation")
    Set l_description = m_database.createElement("description")
    Set l_Parameters = m_database.createElement("arguments")
    Set l_Access = m_database.createElement("access")
    
    '// set the command name
    Call l_command.setAttribute("name", strCommand)
    
    '// set the owner if necessary
    If strScriptOwner <> vbNullString Then
        Call l_command.setAttribute("owner", strScriptOwner)
    End If
    
    
    '// create heirarchy
    Call l_command.appendChild(l_aliases)
    Call l_command.appendChild(l_documentation)
    Call l_documentation.appendChild(l_description)
    Call l_command.appendChild(l_Parameters)
    Call l_command.appendChild(l_Access)
    
    
    '// append it to the database
    Call m_database.documentElement.appendChild(l_command)
    
    
End Function

Public Function NewParameter(argName As String, argIsOptional As Boolean, Optional argDataType As String = "String")
    
    Dim Param As clsCommandParamsObj
    
    '// create a new parameter
    Set Param = New clsCommandParamsObj
    With Param
        .Name = argName
        .IsOptional = argIsOptional
        .datatype = argDataType
    End With
    
    Set NewParameter = Param
    
End Function


Public Function NewRestriction(argName As String, Optional argRequiredRank As Integer = -1, Optional argRequiredFlags As String = vbNullString)

    Dim res As clsCommandRestrictionObj
        
    '// create a new restriction
    Set res = New clsCommandRestrictionObj
    With res
        .Name = argName
        .RequiredRank = argRequiredRank
        .RequiredFlags = argRequiredFlags
    End With
    
    Set NewRestriction = res
        
End Function


Private Function GetXSD() As String
    Dim oFSO As FileSystemObject, oTS As TextStream
    Dim strXSD As String
    
    Set oFSO = New FileSystemObject
        
    '// read the xsd file
    Set oTS = oFSO.OpenTextFile(StringFormat("{0}\Commands.xsd", App.Path), ForReading, False)
    strXSD = oTS.ReadAll()
    Call oTS.Close
    
    Set oFSO = Nothing
    Set oTS = Nothing
    
    GetXSD = strXSD
    
End Function


Private Sub ShowXMLErrors(ByRef colErrorList As Collection)
    
    Dim i As Integer
    Dim msg As String
    
    msg = StringFormat("The following errors were detected with commands.xml...{0}{0}", vbNewLine)
    
    For i = 1 To colErrorList.Count
        msg = StringFormat("{0}{1}{2}", msg, colErrorList(i), vbNewLine)
    Next i
    
    frmChat.AddChat vbRed, msg
    
End Sub

Public Function CommandsSanityCheck(ByRef doc As DOMDocument60, Optional ByRef colErrorList As Collection) As Boolean
    
    'On Error GoTo ERROR_HANDLER
    
    Set colErrorList = New Collection
    
    '////////////////////////////
    '// COMMAND NODES
    '////////////////////////////
    Dim oCommands As IXMLDOMNodeList
    Dim oCommand As IXMLDOMNode
    Dim uniqueEnabledCommands As Dictionary
    Dim uniqueCommandOwners As Dictionary
    
    Set oCommands = doc.documentElement.selectNodes("/commands/command")
    Set uniqueEnabledCommands = New Dictionary
    Set uniqueCommandOwners = New Dictionary
    
    For Each oCommand In oCommands
        Do
           
            Dim commandName As String
            Dim ownerName As String
            Dim Enabled As String
            
            '// default our values
            commandName = oCommand.Attributes.getNamedItem("name").Text
            ownerName = vbNullString
            Enabled = "1"
            
            '// make sure name != ""
            If Len(commandName) = 0 Then
                colErrorList.Add "A command element is missing the name attribute."
                Exit Do
            End If
            
            '// make sure owner != ""
            If Not (oCommand.Attributes.getNamedItem("owner") Is Nothing) Then
                ownerName = oCommand.Attributes.getNamedItem("owner").Text
                If Len(ownerName) = 0 Then
                    colErrorList.Add "An owner attribute cannot be empty on command element."
                    Exit Do
                End If
            End If
            
            
            '// make sure enabled attribute is 0 or 1
            If Not (oCommand.Attributes.getNamedItem("enabled") Is Nothing) Then
                Enabled = oCommand.Attributes.getNamedItem("enabled").Text
                If Enabled <> "0" And Enabled <> "1" Then
                    colErrorList.Add "If present, an enabled attribute must be equal to 0 or 1 on a command element."
                    Exit Do
                End If
            End If
            
            '// make sure only 1 command is enabled if there are similar names
            If Enabled = "1" Then
                If uniqueEnabledCommands.Exists(commandName) Then
                    colErrorList.Add "Only 1 command element can have no enabled attribute or an enabled attribute equal to 1."
                    Exit Do
                End If
                uniqueEnabledCommands.Add commandName, ownerName
            End If
            
            '// make sure commands with the same name of separate owners
            If uniqueCommandOwners.Exists(commandName & "|" & ownerName) Then
                colErrorList.Add "Commands with equal name attributes must have different unique owner attributes."
                Exit Do
            End If
            uniqueEnabledCommands.Add commandName & "|" & ownerName, commandName & "|" & ownerName
            
            '////////////////////////////
            '// ARGUEMENT NODES
            '////////////////////////////
            Dim oArguments As IXMLDOMNodeList
            Dim oArgument As IXMLDOMNode
            Dim uniqueEnabledArguments As Dictionary
            
            Set oArguments = oCommand.selectNodes("arguments/argument")
            Set uniqueEnabledArguments = New Dictionary
            
            For Each oArgument In oArguments
                Do
                    
                    Dim argumentName As String
                    Dim datatype As String
                    
                    '// default our values
                    argumentName = oArgument.Attributes.getNamedItem("name").Text
                    
                    '// make sure name != ""
                    If Len(argumentName) = 0 Then
                        colErrorList.Add "An argument element is missing the name attribute."
                        Exit Do
                    End If
                    
                    '// make sure name attribute is unique
                    If uniqueEnabledArguments.Exists(argumentName) Then
                        colErrorList.Add "Argument elements for a command must have unique name attributes."
                        Exit Do
                    End If
                    uniqueEnabledArguments.Add argumentName, commandName
                    
                    '// make sure datatype attribute is string or word or numeric or number
                    If Not (oArgument.Attributes.getNamedItem("datatype") Is Nothing) Then
                        datatype = LCase(oArgument.Attributes.getNamedItem("datatype").Text)
                        If datatype <> "string" And datatype <> "word" And datatype <> "numeric" And datatype <> "number" Then
                            colErrorList.Add "If present, a datatype attribute must be equal to string, word, numeric, or number on a argument element."
                            Exit Do
                        End If
                    End If
                    
                    '// make sure match message is ok if present
                    If Not (oArgument.selectSingleNode("match") Is Nothing) Then
                        If Not (oArgument.selectSingleNode("match").Attributes.getNamedItem("message") Is Nothing) Then
                            If Len(oArgument.selectSingleNode("match").Attributes.getNamedItem("message").Text) = 0 Then
                                colErrorList.Add "If present, the message attribute of the match element for an argument must have a value."
                                Exit Do
                            End If
                        Else
                            '// match element is present, but no message attribute
                            colErrorList.Add "Match element of an argument must contain a message attribute."
                            Exit Do
                        End If
                    
                    End If
                    
                    
                    '////////////////////////////
                    '// RESTRICTION NODES
                    '////////////////////////////
                    Dim oRestrictions As IXMLDOMNodeList
                    Dim oRestriction As IXMLDOMNode
                    Dim uniqueEnabledRestrictions As Dictionary
                    
                    Set oRestrictions = oArgument.selectNodes("restrictions/restriction")
                    Set uniqueEnabledRestrictions = New Dictionary
                    
                    For Each oRestriction In oRestrictions
                        Do
                            
                            Dim restrictionName As String
        
                            '// default our values
                            restrictionName = oRestriction.Attributes.getNamedItem("name").Text
                            
                            '// make sure name != ""
                            If Len(restrictionName) = 0 Then
                                colErrorList.Add "A restriction element is missing the name attribute."
                                Exit Do
                            End If
        
                            '// make sure name attribute is unique
                            If uniqueEnabledRestrictions.Exists(restrictionName) Then
                                colErrorList.Add "Restriction elements for an argument must have unique name attributes."
                                Exit Do
                            End If
                            uniqueEnabledRestrictions.Add restrictionName, argumentName
                            
                            Exit Do
                        Loop
                    Next oRestriction
                    
                    Exit Do
                Loop
            Next oArgument

            Exit Do
        Loop
    Next oCommand
    
    If colErrorList.Count > 0 Then
        ShowXMLErrors colErrorList
        CommandsSanityCheck = False
    Else
        CommandsSanityCheck = True
    End If

    Exit Function
    
ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObjStatic.CommandsSanityCheck().")

End Function

Public Function Save(Optional ByVal writeToFile As Boolean = True) As Boolean

    On Error GoTo ERROR_HANDLER

    Dim colErrorList As Collection

    setAliases m_command_node, m_aliases
    setParameters m_command_node, m_params
    
    If writeToFile = True Then
    
        If clsCommandDocObj.ValidateXMLFromStrings(m_database.xml, GetXSD()) Then
            If Not CommandsSanityCheck(m_database, colErrorList) Then
                Save = False
                Exit Function
            End If
        Else
            Save = False
            Exit Function
        End If

        '// 08/302009 52 - getting rid of the clsXML class since it didnt write valid XML
        Call writeDocToFile(GetFilePath("Commands.xml"))
        
    End If
    
    Save = True
    
    Exit Function
    
ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.Save().")

End Function



Private Sub writeDocToFile(ByVal filePath As String)
    
    On Error GoTo ERROR_HANDLER
    
    Dim reader As New SAXXMLReader60 '// create the SAX reader
    Dim writer As New MXXMLWriter60  '// create the XML writer
    
    Dim iFileNo As Integer
    
    '// set properties on the XML writer
    writer.byteOrderMark = True
    writer.omitXMLDeclaration = True
    writer.indent = True
    
    '// set the XML writer to the SAX content handler
    Set reader.contentHandler = writer
    Set reader.dtdHandler = writer
    Set reader.ErrorHandler = writer
    reader.putProperty "http://xml.org/sax/properties/lexical-handler", writer
    reader.putProperty "http://xml.org/sax/properties/declaration-handler", writer
    
    '// parse the DOMDocument object
    reader.Parse m_database
    
    '// open the file for writing
    iFileNo = FreeFile
    Open filePath For Output As #iFileNo
    Print #iFileNo, writer.output
    
    '// close the file
    Close #iFileNo
    
    Set reader = Nothing
    Set writer = Nothing
    
    Exit Sub

ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.writeDocToFile().")

End Sub


Public Sub Delete()
    
    If (m_command_node Is Nothing) Then
        Exit Sub
    End If
    
    If (LenB(Me.Owner) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Error: You can not delete an internal command."
        Exit Sub
    End If
    
    m_command_node.parentNode.removeChild m_command_node
    
    Call Save

End Sub

Private Function getParameters(ByRef commandNode As IXMLDOMNode) As Collection

    Dim parameter  As IXMLDOMNode
    Dim Parameters As IXMLDOMNodeList
    Dim temp       As clsCommandParamsObj

    Set getParameters = New Collection
    
    If (commandNode Is Nothing) Then
        Exit Function
    End If
    
    Set Parameters = commandNode.selectNodes("arguments/argument")
    
    For Each parameter In Parameters
        Set temp = New clsCommandParamsObj
        
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
            .MatchError = parameter.selectSingleNode("error").Text
            .description = parameter.selectSingleNode("documentation/description").Text
            .SpecialNotes = parameter.selectSingleNode("documentation/specialnotes").Text
        End With
        On Error GoTo 0
        
        getParameters.Add temp
        
        Set temp = Nothing
    Next
    Set Parameters = Nothing

End Function

Private Function getRestrictions(ByRef parameterNode As IXMLDOMNode) As Collection

    Dim RestrictionNode  As IXMLDOMNode
    Dim RestrictionsNode As IXMLDOMNode
    Dim temp             As clsCommandRestrictionObj
    
    Set getRestrictions = New Collection
    
    If (parameterNode Is Nothing) Then
        Exit Function
    End If
    
    Set RestrictionsNode = parameterNode.selectSingleNode("restrictions")
    
    If (RestrictionsNode Is Nothing) Then
        Exit Function
    End If
    
    For Each RestrictionNode In RestrictionsNode.selectNodes("restriction")
        Set temp = New clsCommandRestrictionObj
        With temp
            .Name = getName(RestrictionNode)
            .MatchMessage = getMatchMessage(RestrictionNode)
            If (Not RestrictionNode.selectSingleNode("error") Is Nothing) Then
                .MatchError = RestrictionNode.selectSingleNode("error").Text
            End If
            .MatchCaseSensitive = getMatchCase(RestrictionNode)
            .RequiredFlags = getFlags(RestrictionNode)
            .RequiredRank = GetRank(RestrictionNode)
            .Fatal = GetFatal(RestrictionNode)
        End With
        getRestrictions.Add temp
        Set temp = Nothing
    Next
End Function

Private Function getAliases(ByRef AnyNode As IXMLDOMNode) As Collection

    Dim Alias   As IXMLDOMNode
    Dim aliases As IXMLDOMNodeList
    
    Set getAliases = New Collection
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// 09/03/2008 JSM - Modified code to use the <aliases> element
    Set aliases = AnyNode.selectNodes("aliases/alias")
    
    If ((aliases Is Nothing) = False) Then
        For Each Alias In aliases
            getAliases.Add Alias.Text
        Next
    End If

End Function

Private Function setParameters(ByRef AnyNode As IXMLDOMNode, ByRef ParameterCol As Collection) As Collection
    
    Dim argumentNode                  As IXMLDOMNode
    Dim argumentsNode                 As IXMLDOMNode
    Dim argumentsOptionalNode         As IXMLDOMNode
    Dim argumentsOptionalStringNode   As IXMLDOMNode
    Dim argumentsOptionalNumberNode   As IXMLDOMNode
    Dim RestrictionNode               As IXMLDOMNode
    
    Dim newNode                       As IXMLDOMNode
    Dim node                          As IXMLDOMNode
    Dim parameter                     As clsCommandParamsObj
    Dim Restriction                   As clsCommandRestrictionObj
    Dim i                             As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// get the arguments node
    Set argumentsNode = AnyNode.selectSingleNode("arguments")
    If Not (argumentsNode Is Nothing) Then
            For Each node In argumentsNode.childNodes
                argumentsNode.removeChild node
            Next node
    End If
    

    
    
    '// add the the new arguments node if necessary
    If ParameterCol.Count > 0 Then
        
        '// loop through the parameters
        For i = 1 To ParameterCol.Count
            
            Set parameter = ParameterCol(i)
            
            Set argumentNode = m_database.createNode("element", "argument", vbNullString)
            
            '// @name
            Set newNode = m_database.createNode("attribute", "name", vbNullString)
            newNode.Text = parameter.Name
            argumentNode.Attributes.setNamedItem newNode
            
            '// documentation
            Set newNode = m_database.createNode("element", "documentation", vbNullString)
            argumentNode.appendChild newNode
            
            '// create the match element if necessary
            If Len(parameter.MatchMessage) <> 0 Then
                '// add the match element
                argumentNode.appendChild m_database.createNode("element", "match", vbNullString)
            
                '// match[@message]
                Set newNode = m_database.createNode("attribute", "message", vbNullString)
                newNode.Text = parameter.MatchMessage
                argumentNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            
                '// match[@case-sensitive]
                Set newNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
                newNode.Text = IIf(parameter.MatchCaseSensitive, "1", "0")
                argumentNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            End If
            
            '// create the error element if necessary
            If Len(parameter.MatchError) <> 0 Then
                '// add the error element
                argumentNode.appendChild m_database.createNode("element", "error", vbNullString)
                argumentNode.selectSingleNode("error").Text = parameter.MatchError
            End If
            
            
            '// documentation/description
            Set newNode = m_database.createNode("element", "description", vbNullString)
            newNode.Text = parameter.description
            argumentNode.selectSingleNode("documentation").appendChild newNode
            
            '// documentation/specialnotes
            Set newNode = m_database.createNode("element", "specialnotes", vbNullString)
            newNode.Text = parameter.SpecialNotes
            argumentNode.selectSingleNode("documentation").appendChild newNode
            
            '// restrictions
            If parameter.Restrictions.Count > 0 Then
                setRestrictions argumentNode, parameter.Restrictions
            End If
            
            If parameter.IsOptional Then
                Set newNode = m_database.createNode("attribute", "optional", vbNullString)
                newNode.Text = "1"
                argumentNode.Attributes.setNamedItem newNode
            End If
            
            If (Not StrComp(parameter.datatype, "string", vbTextCompare) = 0) Then
                Set newNode = m_database.createNode("attribute", "type", vbNullString)
                newNode.Text = parameter.datatype
                argumentNode.Attributes.setNamedItem newNode
            End If
            
            argumentsNode.appendChild argumentNode
        Next
    End If
End Function

Private Function setRestrictions(ByRef AnyNode As IXMLDOMNode, ByRef RestrictionCol As Collection) As Collection
    
    Dim RestrictionNode                  As IXMLDOMNode
    
    Dim newNode                          As IXMLDOMNode
    Dim node                             As IXMLDOMNode
    Dim Restriction                      As clsCommandRestrictionObj
    Dim i                                As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// add our restrictions node
    AnyNode.appendChild m_database.createNode("element", "restrictions", vbNullString)
    
    '// loop through argument restrictions
    For Each Restriction In RestrictionCol
        
        '// create our restriction node
        Set RestrictionNode = m_database.createNode("element", "restriction", vbNullString)
        AnyNode.selectSingleNode("restrictions").appendChild RestrictionNode

        '// name attribute
        Set newNode = m_database.createNode("attribute", "name", vbNullString)
        newNode.Text = Restriction.Name
        RestrictionNode.Attributes.setNamedItem newNode
        
        If (Restriction.Fatal) Then
            Set newNode = m_database.createNode("attribute", "nonfatal", vbNullString)
            newNode.Text = IIf(Restriction.Fatal, "0", "1")
            RestrictionNode.Attributes.setNamedItem newNode
        End If
        
        '// create the match element if necessary
        If Restriction.MatchMessage <> "" Then
            '// add the match element
            RestrictionNode.appendChild m_database.createNode("element", "match", vbNullString)
            
            '// set the message attribute
            Set newNode = m_database.createNode("attribute", "message", vbNullString)
            newNode.Text = Restriction.MatchMessage
            RestrictionNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            '// match[@message]
            Set newNode = m_database.createNode("attribute", "message", vbNullString)
            newNode.Text = Restriction.MatchMessage
            RestrictionNode.selectSingleNode("match").Attributes.setNamedItem newNode
        
        
            '// match[@case-sensitive]
            Set newNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
            newNode.Text = IIf(Restriction.MatchCaseSensitive, "1", "0")
            RestrictionNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            
        End If
        
        
            
        '// create the error element if necessary
        If Len(Restriction.MatchError) <> 0 Then
            '// add the error element
            RestrictionNode.appendChild m_database.createNode("element", "error", vbNullString)
            RestrictionNode.selectSingleNode("error").Text = Restriction.MatchError
        End If
        
        
        
        '// create the access node
        RestrictionNode.appendChild m_database.createNode("element", "access", vbNullString)
        
        '// create the rank noode
        RestrictionNode.selectSingleNode("access").appendChild m_database.createNode("element", "rank", vbNullString)
        RestrictionNode.selectSingleNode("access/rank").Text = Restriction.RequiredRank
        
        '// create the flags node
        RestrictionNode.selectSingleNode("access").appendChild m_database.createNode("element", "flags", vbNullString)
        
        '// add a flag element for each flag
        For i = 1 To Len(Restriction.RequiredFlags)
            Set newNode = m_database.createNode("element", "flag", vbNullString)
            newNode.Text = Mid(Restriction.RequiredFlags, i, 1)
            RestrictionNode.selectSingleNode("access/flags").appendChild newNode
        Next i
                        
    Next Restriction
    

End Function




Private Function setAliases(ByRef AnyNode As IXMLDOMNode, ByRef AliasCol As Collection) As Collection

    Dim Alias   As IXMLDOMNode
    Dim aliases As IXMLDOMNodeList
    Dim i       As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// 09/03/2008 JSM - Modified code to use the <aliases> element
    Set aliases = AnyNode.selectNodes("aliases/alias")
    
    If ((aliases Is Nothing) = False) Then
        For Each Alias In aliases
            '// 09/03/2008 JSM - Modified code to use the <aliases> element
            AnyNode.selectSingleNode("aliases").removeChild Alias
        Next
    End If
    
    For i = 1 To AliasCol.Count
        '// 09/03/2008 JSM - Modified code to use the <aliases> element
        Set Alias = _
            AnyNode.selectSingleNode("aliases").appendChild(m_database.createNode("element", "alias", vbNullString))
        
        Alias.Text = AliasCol(i)
    Next
    
    'Call Save

End Function

Private Function getName(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.Attributes.getNamedItem("name")
    If (temp Is Nothing) Then
        Exit Function
    End If
    getName = temp.Text

End Function

Private Function getOwner(ByRef AnyNode As IXMLDOMNode) As String
    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.Attributes.getNamedItem("owner")
    If (temp Is Nothing) Then
        Exit Function
    End If
    getOwner = temp.Text
End Function

Private Function getMatchMessage(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("match")
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    Set temp = temp.Attributes.getNamedItem("message")
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    getMatchMessage = temp.Text

End Function

Private Function getMatchCase(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim temp As IXMLDOMNode
    getMatchCase = False
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("match")
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    Set temp = temp.Attributes.getNamedItem("case-sensitive")
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    getMatchCase = (temp.Text = "1")

End Function

Private Function getEnabled(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    Set temp = AnyNode.Attributes.getNamedItem("enabled")
    getEnabled = True
    
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    If (temp.Text = "0") Then
        getEnabled = False
    End If

End Function

Private Function setEnabled(ByRef AnyNode As IXMLDOMNode, ByVal Enabled As Boolean)

    Dim temp As IXMLDOMNode
    Dim attr As IXMLDOMAttribute
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    Set temp = AnyNode.Attributes.getNamedItem("enabled")
    If (temp Is Nothing) Then
        Set temp = AnyNode.Attributes.setNamedItem(m_database.createAttribute("enabled"))
    End If

    If (Enabled = True) Then
        temp.Text = "1"
    Else
        temp.Text = "0"
    End If
    
    'Call Save

End Function

Private Function getDescription(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("documentation/description")
    If (temp Is Nothing) Then
        Exit Function
    End If
    getDescription = temp.Text

End Function

'Syntax: .add <username> [rank] [attributes]
Private Function getSyntaxString(Optional IsLocal As Boolean = False) As String

    Dim retVal As String
    Dim P As clsCommandParamsObj
    
    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
    '// add the command name
    retVal = StringFormat("{0} ", Me.Name)
    
    '// add the parameters
    For Each P In Me.Parameters
        retVal = StringFormat("{0}{2}{1}{3} ", _
                    retVal, _
                    P.Name, _
                    IIf(P.IsOptional, "[", "<"), _
                    IIf(P.IsOptional, "]", ">"))
    Next
    
    '// add the trigger
    retVal = StringFormat("{1}{0}", _
        retVal, _
        IIf((Me.RequiredRank <> -1) Or (Len(Me.RequiredFlags) > 0), _
            IIf(IsLocal, "/", BotVars.Trigger), "/"))
    
    getSyntaxString = Trim$(retVal)

End Function

'alias1, alias2, alias3
Private Function getAliasString()
    Dim retVal As String
    Dim i As Integer
    If (aliases.Count > 0) Then
        For i = 1 To aliases.Count
            retVal = StringFormat("{0}{1}{2}", _
                retVal, _
                CStr(aliases.Item(i)), _
                IIf(i = aliases.Count, vbNullString, ", "))
        Next i
    End If
    getAliasString = retVal
End Function


Private Function getRequirementsString(Optional bShort As Boolean = False)

    Dim retVal As String
    
    If (m_command_node Is Nothing) Then
        Exit Function
    End If

    If (Me.RequiredRank <> -1) Or (Len(Me.RequiredFlags) > 0) Then
        '// available outside the console
        retVal = ""
        
        '// add access requirements if necessary
        If (Me.RequiredRank <> -1) Then
            retVal = StringFormat("{0} Requires rank {1}", retVal, Me.RequiredRank)
            If (Len(Me.RequiredFlags) = 0) Then
                retVal = StringFormat("{0}.", retVal)
            End If
        End If
        
        '// add flag requirements, if necessary, taking into account the existinance of an access requirements
        If (Len(Me.RequiredFlags) <> 0) Then
            If (Me.RequiredRank <> -1) Then
                retVal = StringFormat("{0} or one of the following flags: {1}.", retVal, Me.RequiredFlags)
            Else
                retVal = StringFormat("{0} Requires one of the following flags: {1}.", retVal, Me.RequiredFlags)
            End If
        End If
    
    Else
        '// console only, lets show the generic blurb
        retVal = "Command is only available to the console."
        If (Not bShort) Then
            retVal = StringFormat("{0} To allow other users to use this command, set the command's flags requirement or rank requirement.", retVal)
        End If
    End If

    getRequirementsString = retVal

End Function



Private Function setDescription(ByRef AnyNode As IXMLDOMNode, ByVal description As String) As String

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("documentation/description")
    If (temp Is Nothing) Then
        Set temp = AnyNode.selectSingleNode("documentation")
        If (temp Is Nothing) Then
            Set temp = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
        End If
        Set temp = temp.appendChild(m_database.createNode("element", "description", vbNullString))
    End If
    
    temp.Text = description
    
    'Call Save

End Function

Private Function getNotes(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("documentation/specialnotes")
    If (temp Is Nothing) Then
        Exit Function
    End If
    getNotes = temp.Text

End Function

Private Function setNotes(ByRef AnyNode As IXMLDOMNode, ByVal Notes As String) As String

    Dim temp As IXMLDOMNode
     
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
   
    Set temp = AnyNode.selectSingleNode("documentation/specialnotes")
    If (temp Is Nothing) Then
        Set temp = AnyNode.selectSingleNode("documentation")
        If (temp Is Nothing) Then
            Set temp = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
        End If
        
        Set temp = temp.appendChild(m_database.createNode("element", "specialnotes", vbNullString))
    End If
    
    temp.Text = Notes
    
    'Call Save
    
End Function

Private Function GetRank(ByRef AnyNode As IXMLDOMNode) As Integer

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("access/rank")
    
    If (temp Is Nothing) Then
        GetRank = -1
    
        Exit Function
    End If
    
    If (temp.Text = vbNullString) Then
        GetRank = -1
    
        Exit Function
    End If
    
    GetRank = CInt(Val(temp.Text))

End Function

Private Function GetFatal(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim temp As IXMLDOMNode
    GetFatal = True
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.Attributes.getNamedItem("nonfatal")
    If (temp Is Nothing) Then
        Exit Function
    End If
    GetFatal = (Not temp.Text = "1")
End Function

Private Function setRank(ByRef AnyNode As IXMLDOMNode, ByVal Rank As Integer) As Integer

    Dim temp As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("access/rank")
    
    If (temp Is Nothing) Then
        Set temp = AnyNode.selectSingleNode("access")
        
        If (temp Is Nothing) Then
            Set temp = AnyNode.appendChild(m_database.createNode("element", "access", vbNullString))
        End If
        
        Set temp = temp.appendChild(m_database.createNode("element", "rank", vbNullString))
    End If

    temp.Text = Rank
    
    'Call Save

End Function

Private Function getFlags(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNodeList
    Dim Flag As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectNodes("access/flags/flag")
    
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    For Each Flag In temp
        getFlags = getFlags & Flag.Text
    Next
    
End Function

Private Function setFlags(ByRef AnyNode As IXMLDOMNode, ByVal Flags As String) As String

    Dim temp As IXMLDOMNode
    Dim Flag As IXMLDOMNode
    Dim i    As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("access/flags")
        
    If (temp Is Nothing) Then
        Set temp = AnyNode.selectSingleNode("access")
        
        Set temp = _
            temp.appendChild(m_database.createNode("element", "flags", vbNullString))
    End If
    
    For i = temp.childNodes.length - 1 To 0 Step -1
        temp.removeChild temp.childNodes(i)
    Next i
    
    For i = 1 To Len(Flags)
        Set Flag = _
            temp.appendChild(m_database.createNode("element", "flag", vbNullString))
        
        Flag.Text = Mid$(Flags, i, 1)
    Next i
    
    'Call Save
    
End Function

Public Property Get Name() As String
    
    Name = getName(m_command_node)

End Property

Public Property Get Owner() As String
    
    Owner = getOwner(m_command_node)
    
End Property

Public Property Get aliases() As Collection

    If (m_aliases Is Nothing) Then
        Set m_aliases = getAliases(m_command_node)
    End If
    
    Set aliases = m_aliases
    
End Property

Public Property Get IsEnabled() As Boolean
    
    IsEnabled = getEnabled(m_command_node)

End Property

Public Property Let IsEnabled(bln As Boolean)
    
    setEnabled m_command_node, bln

End Property

Public Property Let Name(strName As String)

    m_name = strName

End Property

Public Property Let Owner(strOwner As String)

    m_owner = strOwner
    
End Property

Public Property Get RequiredRank() As Integer
    
    RequiredRank = GetRank(m_command_node)

End Property

Public Property Let RequiredRank(ByVal Rank As Integer)
    
    setRank m_command_node, Rank

End Property

Public Property Get RequiredFlags() As String
    
    RequiredFlags = getFlags(m_command_node)

End Property

Public Property Let RequiredFlags(ByVal str As String)
   
   setFlags m_command_node, str

End Property

Public Property Get description() As String
    
    description = getDescription(m_command_node)

End Property

Public Property Let description(ByVal strDescription As String)
    
    setDescription m_command_node, strDescription

End Property

Public Function SyntaxString(Optional IsLocal As Boolean = False) As String
    
    SyntaxString = getSyntaxString(IsLocal)
    
End Function

Public Property Get AliasString() As String
    
    AliasString = getAliasString()
    
End Property

Public Property Get RequirementsString() As String
    
    RequirementsString = getRequirementsString(False)
    
End Property

Public Property Get RequirementsStringShort()
    
    RequirementsStringShort = getRequirementsString(True)
    
End Property

Public Property Get SpecialNotes() As String
    
    SpecialNotes = getNotes(m_command_node)

End Property

Public Property Let SpecialNotes(ByVal strNotes As String)
    
    setNotes m_command_node, strNotes

End Property

Public Function GetParameterByName(ByVal sParamName As String) As clsCommandParamsObj
    
    Dim P As clsCommandParamsObj
    Dim col As Collection
    Dim i As Integer
    
    Set col = Me.Parameters
        
    For i = 1 To col.Count
        Set P = col(i)
        If StrComp(sParamName, P.Name, vbTextCompare) = 0 Then
            Set GetParameterByName = P
            Exit Function
        End If
    Next i
    
End Function


Public Property Get Parameters() As Collection
    
    If (m_params Is Nothing) Then
        Set m_params = getParameters(m_command_node)
    End If
    
    Set Parameters = m_params

End Property


Public Property Get XMLDocument() As DOMDocument60
    Set XMLDocument = m_database
End Property
