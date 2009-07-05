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

Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString)

    If (DatabasePath = vbNullString) Then
        DatabasePath = App.Path & "\commands.xml"
    End If
    
    m_database.Load DatabasePath
    
End Function

Public Function OpenCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String) As Boolean
    
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
        Exit Function
    End If
    
    If strScriptOwner = vbNullString Then
        xpath = "./command[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(strCommand) & "']"
    Else
        xpath = "./command[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(strCommand) & "' and @owner = '" & strScriptOwner & "']"
    End If
    
    Set m_command_node = m_command_node.selectSingleNode(xpath)
    Set m_aliases = getAliases(m_command_node)
    Set m_params = getParameters(m_command_node)
    
    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
End Function

'// 06/13/2009 JSM - Created function
Public Function CreateCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String) As Boolean

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

Public Function NewParameter(argName As String, argIsOptional As Boolean, Optional argDataType As String)
    
    Dim param As clsCommandParamsObj
    
    '// create a new parameter
    Set param = New clsCommandParamsObj
    With param
        .Name = argName
        .IsOptional = argIsOptional
        If argIsOptional Then
            .DataType = argDataType
        End If
    End With
    
    Set NewParameter = param
    
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
            .description = Parameter.selectSingleNode("documentation/description").Text
            .SpecialNotes = Parameter.selectSingleNode("documentation/specialnotes").Text
        End With
        On Error GoTo 0
        
        getParameters.Add temp
        
        Set temp = Nothing
    Next
    
    Set Parameters = CommandNode.selectNodes("arguments/optional/number/argument")
    
    If (Not (Parameters Is Nothing)) Then
        
        For Each Parameter In Parameters
            
            Set temp = New clsCommandParamsObj
            
            On Error Resume Next
            With temp
                .Name = getName(Parameter)
                .DataType = "number"
                .Restrictions = getRestrictions(Parameter)
                .IsOptional = True
                .description = Parameter.selectSingleNode("documentation/description").Text
                .SpecialNotes = Parameter.selectSingleNode("documentation/specialnotes").Text
            End With
            On Error GoTo 0
            
            getParameters.Add temp
            
            Set temp = Nothing
        Next
    End If
    
    Set Parameters = CommandNode.selectNodes("arguments/optional/string/argument")
    
    If (Not (Parameters Is Nothing)) Then
        For Each Parameter In Parameters
            Set temp = New clsCommandParamsObj
            
            On Error Resume Next
            With temp
                .Name = getName(Parameter)
                .DataType = "string"
                .Restrictions = getRestrictions(Parameter)
                .IsOptional = True
                .description = Parameter.selectSingleNode("documentation/description").Text
                .SpecialNotes = Parameter.selectSingleNode("documentation/specialnotes").Text
            End With
            On Error GoTo 0
            
            getParameters.Add temp
            
            Set temp = Nothing
        Next
    End If
    
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
            
            If Parameter.IsOptional = False Then
                '// add to argumentsNode
                argumentsNode.appendChild argumentNode
            Else
                '// make sure the optional node exists
                If (argumentsNode.selectSingleNode("optional") Is Nothing) Then
                    Set argumentsOptionalNode = m_database.createNode("element", "optional", vbNullString)
                End If
            
                '// this value has already been validated in the property setter
                '// make sure the optional/<datatype> node exists
                If (argumentsOptionalNode.selectSingleNode(Parameter.DataType) Is Nothing) Then
                    argumentsOptionalNode.appendChild m_database.createNode("element", Parameter.DataType, vbNullString)
                End If
                
                '// add to optional/<datatype> node
                argumentsOptionalNode.selectSingleNode(Parameter.DataType).appendChild argumentNode
                
            End If
        Next
        
        '// add the optional node if necessary
        If Not (argumentsOptionalNode Is Nothing) Then
            argumentsNode.appendChild argumentsOptionalNode
        End If
    
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
            
            '// set he message attribute
            Set newNode = m_database.createNode("attribute", "message", vbNullString)
            newNode.Text = Restriction.MatchMessage
            RestrictionNode.selectSingleNode("match").Attributes.setNamedItem newNode
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

    Name = getName(m_command_node)
    
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
