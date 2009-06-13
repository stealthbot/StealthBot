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
    Set m_params = New Collection
    Set m_aliases = New Collection
    
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

Public Function OpenCommand(ByVal strCommand As String) As Boolean
    
    Dim command_access_node   As IXMLDOMNode
    Dim command_documentation As IXMLDOMNode
    Dim command_parameters    As IXMLDOMNodeList
    Dim command_aliases       As IXMLDOMNodeList
    Dim Alias                 As IXMLDOMNode
    Dim attrs                 As IXMLDOMAttribute
    
    Set m_command_node = m_database.documentElement

    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
    Set m_command_node = m_command_node.selectSingleNode("./command[@name='" & strCommand & "']")
    
    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
End Function

'// 06/13/2009 JSM - Created function
Public Function CreateCommand(ByVal strCommand As String) As Boolean

    Dim l_command             As IXMLDOMElement
    Dim l_aliases             As IXMLDOMElement
    Dim l_documentation       As IXMLDOMElement
    Dim l_description         As IXMLDOMElement
    Dim l_arguments           As IXMLDOMElement
    Dim l_Access              As IXMLDOMElement
        
    Call OpenDatabase
    
    '// create elements
    Set l_command = m_database.createElement("command")
    Set l_aliases = m_database.createElement("aliases")
    Set l_documentation = m_database.createElement("documentation")
    Set l_description = m_database.createElement("description")
    Set l_arguments = m_database.createElement("arguments")
    Set l_Access = m_database.createElement("access")
    
    '// set the command name
    Call l_command.setAttribute("name", strCommand)
    
    '// create heirarchy
    Call l_command.appendChild(l_aliases)
    Call l_command.appendChild(l_documentation)
    Call l_documentation.appendChild(l_description)
    Call l_command.appendChild(l_arguments)
    Call l_command.appendChild(l_Access)
    
    
    '// append it to the database
    Call m_database.documentElement.appendChild(l_command)
    
    '// save this crap :)
    Call Save
        
    
End Function



Public Function Save() As Boolean

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
        
        With temp
            .Name = getName(Parameter)
            .Restrictions = getRestrictions(Parameter)
            .IsOptional = False
        End With
        
        getParameters.Add temp
        
        Set temp = Nothing
    Next
    
    ' ...
    Set Parameters = CommandNode.selectNodes("arguments/optional/number/argument")
    
    ' ...
    If (Not (Parameters Is Nothing)) Then
        ' ...
        For Each Parameter In Parameters
            ' ...
            Set temp = New clsCommandParamsObj
            
            ' ...
            With temp
                .Name = getName(Parameter)
                .dataType = "number"
                .Restrictions = getRestrictions(Parameter)
                .IsOptional = True
            End With
            
            ' ...
            getParameters.Add temp
            
            ' ...
            Set temp = Nothing
        Next
    End If
    
    ' ...
    Set Parameters = CommandNode.selectNodes("arguments/optional/string/argument")
    
    ' ...
    If (Not (Parameters Is Nothing)) Then
        ' ...
        For Each Parameter In Parameters
            ' ...
            Set temp = New clsCommandParamsObj
            
            ' ...
            With temp
                .Name = getName(Parameter)
                .dataType = "string"
                .Restrictions = getRestrictions(Parameter)
                .IsOptional = True
            End With
            
            ' ...
            getParameters.Add temp
            
            ' ...
            Set temp = Nothing
        Next
    End If

    ' ...
    Set Parameters = Nothing

End Function

Private Function getRestrictions(ByRef ParameterNode As IXMLDOMNode) As Collection

    Dim Restriction  As IXMLDOMNode
    Dim Restrictions As IXMLDOMNodeList
    Dim temp         As clsCommandRestrictionObj
    
    Set getRestrictions = New Collection
    
    If (ParameterNode Is Nothing) Then
        Exit Function
    End If
    
    Set Restrictions = ParameterNode.selectNodes("restrictions")
    
    If (Restrictions.length = 0) Then
        Exit Function
    End If
    
    Set Restrictions = Restrictions.Item(0).selectNodes("restriction")
    
    For Each Restriction In Restrictions
        Set temp = New clsCommandRestrictionObj
        With temp
            .Name = getName(Restriction)
            .MatchMessage = getMatchMessage(Restriction)
            .RequiredFlags = getFlags(Restriction)
            .RequiredRank = GetRank(Restriction)
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
            getAliases.Add Alias.text
        Next
    End If

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
        
        Alias.text = AliasCol(I)
    Next
    
    Call Save

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
    getName = temp.text

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
    
    getMatchMessage = temp.text

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
    
    If (temp.text = "false") Then
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
        temp.text = "true"
    Else
        temp.text = "false"
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
    getDescription = temp.text

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
    
    temp.text = description
    
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
    getNotes = temp.text

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
    
    temp.text = Notes
    
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
    
    If (temp.text = vbNullString) Then
        GetRank = -1
    
        Exit Function
    End If
    
    GetRank = CInt(Val(temp.text))

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

    temp.text = Rank
    
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
        getFlags = getFlags & flag.text
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
        
        flag.text = Mid$(Flags, I, 1)
    Next I
    
    Call Save
    
End Function

Public Property Get Name() As String

    Name = getName(m_command_node)
    
End Property

Public Property Get aliases() As Collection

    Set aliases = getAliases(m_command_node)
    
End Property

Public Property Let aliases(ByVal col As Collection)

    setAliases m_command_node, col

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
    
    Set Parameters = getParameters(m_command_node)

End Property
