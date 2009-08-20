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

Private Sub Class_Initialize()

    Set m_database = New DOMDocument60
    
    '// 06/23/2009 JSM - initializing these as Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Private Sub Class_Terminate()

    Set m_database = Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString, Optional ByVal forceLoad As Boolean = False)

    If (DatabasePath = vbNullString) Then
        DatabasePath = App.Path & "\commands.xml"
    End If
    
    If forceLoad = True Or m_database.childNodes.length = 0 Then
        m_database.Load DatabasePath
    End If
    
End Function

Public Function GetCommandCount(Optional ByVal strScriptOwner As String = vbNullString) As Integer

    On Error GoTo ERROR_HANDLER
    
    Dim nodes                 As IXMLDOMNodeList
    Dim xpath                 As String
    
    
    OpenDatabase
    
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
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.GetCommands().")

    Exit Function
    
End Function

Private Function GetCommandXPath(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString) As String
    Dim AZ As String
    AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    strCommand = Replace$(strCommand, "'", "&apos;")
    strScriptOwner = Replace$(strScriptOwner, "'", "&apos;")
    
    If strScriptOwner = vbNullString Then
        GetCommandXPath = StringFormat("./command[translate(@name, '{0}', '{1}')='{2}' and not(@owner)]", _
            UCase$(AZ), LCase$(AZ), LCase$(strCommand))
    ElseIf strScriptOwner = Chr$(0) Then
        GetCommandXPath = StringFormat("./command[translate(@name, '{0}', '{1}')='{2}']", _
            UCase$(AZ), LCase$(AZ), LCase$(strCommand))
    Else
        GetCommandXPath = StringFormat("./command[translate(@name, '{0}', '{1}')='{2}' and translate(@name, '{0}', '{1}')='{3}']", _
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
    
    Set m_command_node = m_database.documentElement
    Set m_aliases = Nothing
    Set m_params = Nothing

    If (m_command_node Is Nothing) Then
        OpenCommand = False
        Exit Function
    End If
    
    xpath = GetCommandXPath(strCommand, strScriptOwner)
    
    Set m_command_node = m_command_node.selectSingleNode(xpath)
    
    If (m_command_node Is Nothing) Then
        'Lets try an alias
        xpath = GetCommandXPath(convertAlias(strCommand), strScriptOwner)
        Set m_command_node = m_database.documentElement.selectSingleNode(xpath)
        
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
Public Function CreateCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString) As Boolean

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
    
    '// save this crap :)
    Call Save
        
    
End Function

Public Function NewParameter(argName As String, argIsOptional As Boolean, Optional argDataType As String = "String")
    
    Dim Param As clsCommandParamsObj
    
    '// create a new parameter
    Set Param = New clsCommandParamsObj
    With Param
        .Name = argName
        .IsOptional = argIsOptional
        .DataType = argDataType
    End With
    
    Set NewParameter = Param
    
End Function


Public Function NewRestriction(argName As String, argRequiredRank As Integer, argRequiredFlags As String)

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


Public Function Save() As Boolean

    setAliases m_command_node, m_aliases
    setParameters m_command_node, m_params
    

    Dim XML As clsXML ' ...
    
    ' ...
    Set XML = New clsXML
    
    ' ...
    XML.Path = App.Path & "\commands.xml"
    
    ' ...
    XML.WriteNode m_database
    
    ' ...
    Set XML = Nothing
    
End Function

Public Sub Delete()
    
    If (m_command_node Is Nothing) Then
        Exit Sub
    End If

    m_command_node.parentNode.removeChild m_command_node
    
    Call Save

End Sub

Private Function getParameters(ByRef CommandNode As IXMLDOMNode) As Collection

    Dim Parameter  As IXMLDOMNode
    Dim Parameters As IXMLDOMNodeList
    Dim temp       As clsCommandParamsObj

    Set getParameters = New Collection
    
    If (CommandNode Is Nothing) Then
        Exit Function
    End If
    
    Set Parameters = CommandNode.selectNodes("arguments/argument")
    
    For Each Parameter In Parameters
        Set temp = New clsCommandParamsObj
        
        On Error Resume Next
        With temp
            .Name = getName(Parameter)
            .Restrictions = getRestrictions(Parameter)
            .IsOptional = False
            If (Not Parameter.Attributes.getNamedItem("optional") Is Nothing) Then
                If (Parameter.Attributes.getNamedItem("optional").Text = "1") Then
                    .IsOptional = True
                End If
            End If
            .DataType = "String"
            If (Not Parameter.Attributes.getNamedItem("type") Is Nothing) Then
                .DataType = Parameter.Attributes.getNamedItem("type").Text
            End If
            .MatchMessage = Parameter.selectSingleNode("match").Attributes.getNamedItem("message").Text
            .MatchCaseSensitive = False
            If (Parameter.selectSingleNode("match").Attributes.getNamedItem("case-sensitive").Text = "1") Then
                .MatchCaseSensitive = True
            End If
            .MatchError = Parameter.selectSingleNode("error").Text
            .description = Parameter.selectSingleNode("documentation/description").Text
            .SpecialNotes = Parameter.selectSingleNode("documentation/specialnotes").Text
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
    Dim Parameter                     As clsCommandParamsObj
    Dim Restriction                   As clsCommandRestrictionObj
    Dim I                             As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// get the arguments node
    Set argumentsNode = AnyNode.selectSingleNode("arguments")
    
    For Each node In argumentsNode.childNodes
        argumentsNode.removeChild node
    Next node
    
    
    '// add the the new arguments node if necessary
    If ParameterCol.Count > 0 Then
        
        '// loop through the parameters
        For I = 1 To ParameterCol.Count
            
            Set Parameter = ParameterCol(I)
            
            Set argumentNode = m_database.createNode("element", "argument", vbNullString)
            
            '// @name
            Set newNode = m_database.createNode("attribute", "name", vbNullString)
            newNode.Text = Parameter.Name
            argumentNode.Attributes.setNamedItem newNode
            
            '// documentation
            Set newNode = m_database.createNode("element", "documentation", vbNullString)
            argumentNode.appendChild newNode
            
            '// create the match element if necessary
            If Len(Parameter.MatchMessage) <> 0 Then
                '// add the match element
                argumentNode.appendChild m_database.createNode("element", "match", vbNullString)
            
                '// match[@message]
                Set newNode = m_database.createNode("attribute", "message", vbNullString)
                newNode.Text = Parameter.MatchMessage
                argumentNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            
                '// match[@case-sensitive]
                Set newNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
                newNode.Text = IIf(Parameter.MatchCaseSensitive, "1", "0")
                argumentNode.selectSingleNode("match").Attributes.setNamedItem newNode
            
            End If
            
            '// create the error element if necessary
            If Len(Parameter.MatchError) <> 0 Then
                '// add the error element
                Set newNode = m_database.createNode("element", "error", vbNullString)
                newNode.Text = Parameter.MatchError
                argumentNode.selectSingleNode("error").appendChild newNode
            End If
            
            
            '// documentation/description
            Set newNode = m_database.createNode("element", "description", vbNullString)
            newNode.Text = Parameter.description
            argumentNode.selectSingleNode("documentation").appendChild newNode
            
            '// documentation/specialnotes
            Set newNode = m_database.createNode("element", "specialnotes", vbNullString)
            newNode.Text = Parameter.SpecialNotes
            argumentNode.selectSingleNode("documentation").appendChild newNode
            
            '// restrictions
            If Parameter.Restrictions.Count > 0 Then
                setRestrictions argumentNode, Parameter.Restrictions
            End If
            
            If Parameter.IsOptional Then
                Set newNode = m_database.createNode("attribute", "optional", vbNullString)
                newNode.Text = "1"
                argumentNode.Attributes.setNamedItem newNode
            End If
            
            If (Not StrComp(Parameter.DataType, "string", vbTextCompare) = 0) Then
                Set newNode = m_database.createNode("attribute", "type", vbNullString)
                newNode.Text = Parameter.DataType
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
    Dim I                                As Integer
    
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
            Set newNode = m_database.createNode("element", "error", vbNullString)
            newNode.Text = Restriction.MatchError
            RestrictionNode.selectSingleNode("error").appendChild newNode
        End If
        
        
        '// create the access node
        RestrictionNode.appendChild m_database.createNode("element", "access", vbNullString)
        
        '// create the rank noode
        RestrictionNode.selectSingleNode("access").appendChild m_database.createNode("element", "rank", vbNullString)
        RestrictionNode.selectSingleNode("access/rank").Text = Restriction.RequiredRank
        
        '// create the flags node
        RestrictionNode.selectSingleNode("access").appendChild m_database.createNode("element", "flags", vbNullString)
        
        '// add a flag element for each flag
        For I = 1 To Len(Restriction.RequiredFlags)
            Set newNode = m_database.createNode("element", "flag", vbNullString)
            newNode.Text = Mid(Restriction.RequiredFlags, I, 1)
            RestrictionNode.selectSingleNode("access/flags").appendChild newNode
        Next I
                        
    Next Restriction
    

End Function




Private Function setAliases(ByRef AnyNode As IXMLDOMNode, ByRef AliasCol As Collection) As Collection

    Dim Alias   As IXMLDOMNode
    Dim aliases As IXMLDOMNodeList
    Dim I       As Integer
    
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
    
    For I = 1 To AliasCol.Count
        '// 09/03/2008 JSM - Modified code to use the <aliases> element
        Set Alias = _
            AnyNode.selectSingleNode("aliases").appendChild(m_database.createNode("element", "alias", vbNullString))
        
        Alias.Text = AliasCol(I)
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
    
    If (temp.Text = "false") Then
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
        temp.Text = "true"
    Else
        temp.Text = "false"
    End If
    
    Call Save

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
Private Function getSyntaxString() As String

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
        IIf((Me.RequiredRank <> -1) Or (Len(Me.RequiredFlags) > 0), BotVars.Trigger, "/"))
    
    getSyntaxString = Trim$(retVal)

End Function

'alias1, alias2, alias3
Private Function getAliasString()
    Dim retVal As String
    Dim I As Integer
    If (aliases.Count > 0) Then
        For I = 1 To aliases.Count
            retVal = StringFormat("{0}{1}{2}", _
                retVal, _
                CStr(aliases.Item(I)), _
                IIf(I = aliases.Count, vbNullString, ", "))
        Next I
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
    
    Call Save

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
    
    Call Save
    
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
    
    Call Save

End Function

Private Function getFlags(ByRef AnyNode As IXMLDOMNode) As String

    Dim temp As IXMLDOMNodeList
    Dim flag As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectNodes("access/flags/flag")
    
    If (temp Is Nothing) Then
        Exit Function
    End If
    
    For Each flag In temp
        getFlags = getFlags & flag.Text
    Next
    
End Function

Private Function setFlags(ByRef AnyNode As IXMLDOMNode, ByVal Flags As String) As String

    Dim temp As IXMLDOMNode
    Dim flag As IXMLDOMNode
    Dim I    As Integer
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set temp = AnyNode.selectSingleNode("access/flags")
        
    If (temp Is Nothing) Then
        Set temp = AnyNode.selectSingleNode("access")
        
        Set temp = _
            temp.appendChild(m_database.createNode("element", "flags", vbNullString))
    End If
    
    For I = temp.childNodes.length - 1 To 0 Step -1
        temp.removeChild temp.childNodes(I)
    Next I
    
    For I = 1 To Len(Flags)
        Set flag = _
            temp.appendChild(m_database.createNode("element", "flag", vbNullString))
        
        flag.Text = Mid$(Flags, I, 1)
    Next I
    
    Call Save
    
End Function

Public Property Get Name() As String
    If m_name = vbNullString Then
        m_name = getName(m_command_node)
    End If
    
    Name = m_name
End Property

Public Property Get Owner() As String
    
    If m_owner = vbNullString Then
        m_owner = getOwner(m_command_node)
    End If
    
    Owner = m_owner
    
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

Public Property Get SyntaxString() As String
    
    SyntaxString = getSyntaxString()
    
End Property

Public Property Get AliasString() As String
    
    AliasString = getAliasString()
    
End Property

Public Property Get RequirementsString() As String
    
    RequirementsString = getRequirementsString(False)
    
End Property

Public Property Get RequirementsStringShort()
    
    RequirementsStringShort = getRequirementsString(True)
    
End Property

Public Property Get Notes() As String
    
    Notes = getNotes(m_command_node)

End Property

Public Property Let Notes(ByVal strNotes As String)
    
    setNotes m_command_node, strNotes

End Property


Public Property Get Parameters() As Collection
    
    If (m_params Is Nothing) Then
        Set m_params = getParameters(m_command_node)
    End If
    
    Set Parameters = m_params

End Property
