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
' ...

Option Explicit

' ...
Private m_database       As DOMDocument
Private m_command_node   As IXMLDOMNode

' ...
Private m_aliases        As Collection
Private m_params         As Collection
Private m_name           As String
Private m_required_rank  As Integer
Private m_required_flags As String
Private m_description    As String
Private m_special_notes  As String
Private m_enabled        As Boolean

Private Sub Class_Initialize()

    ' ...
    Set m_database = New MSXML2.DOMDocument
    Set m_params = New Collection
    Set m_aliases = New Collection
    
End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_database = Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString)

    ' ...
    If (DatabasePath = vbNullString) Then
        DatabasePath = App.Path & "\commands.xml"
    End If
    
    ' ...
    m_database.Load DatabasePath
    
End Function

Public Function OpenCommand(ByVal strCommand As String) As Boolean
    
    Dim command_access_node   As IXMLDOMNode
    Dim command_documentation As IXMLDOMNode
    Dim command_parameters    As IXMLDOMNodeList
    Dim command_aliases       As IXMLDOMNodeList
    Dim Alias                 As IXMLDOMNode
    Dim attrs                 As IXMLDOMAttribute
    
    ' ...
    Set m_command_node = m_database.documentElement

    ' ...
    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    Set m_command_node = _
            m_command_node.selectSingleNode("./command[@name='" & strCommand & "']")
            
    ' ...
    If (m_command_node Is Nothing) Then
        Exit Function
    End If
    
End Function

Public Property Get Name() As String

    Name = getName(m_command_node)

End Property

Public Property Get Aliases() As Collection

    ' ...
    Set Aliases = getAliases(m_command_node)

End Property

Public Property Let Aliases(ByVal col As Collection)

    ' ...
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

    RequiredRank = getRank(m_command_node)
    
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

Public Property Get Description() As String

    Description = getDescription(m_command_node)
    
End Property

Public Property Let Description(ByVal strDescription As String)

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

Public Function Save() As Boolean

    Dim XML As clsXML

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

    ' ...
    If (m_command_node Is Nothing) Then
        Exit Sub
    End If

    ' ...
    m_command_node.parentNode.removeChild m_command_node
    
    ' ...
    Call Save

End Sub

Private Function getParameters(ByRef CommandNode As IXMLDOMNode) As Collection

    Dim Parameter  As IXMLDOMNode
    Dim Parameters As IXMLDOMNodeList
    Dim Temp       As clsCommandParamsObj

    ' ...
    Set getParameters = New Collection
    
    ' ...
    If (CommandNode Is Nothing) Then
        Exit Function
    End If
    
    ' ..
    Set Parameters = _
        CommandNode.selectNodes("arguments/argument")
    
    ' ...
    For Each Parameter In Parameters
        ' ...
        Set Temp = New clsCommandParamsObj
        
        ' ...
        With Temp
            .Name = getName(Parameter)
            .Restrictions = getRestrictions(Parameter)
        End With
        
        ' ...
        getParameters.Add Temp
        
        ' ...
        Set Temp = Nothing
    Next

End Function

Private Function getRestrictions(ByRef ParameterNode As IXMLDOMNode) As Collection

    Dim Restriction  As IXMLDOMNode
    Dim Restrictions As IXMLDOMNodeList
    Dim Temp         As clsCommandRestrictionObj
    
    ' ...
    Set getRestrictions = New Collection
    
    ' ...
    If (ParameterNode Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    Set Restrictions = _
        ParameterNode.selectNodes("restriction")
    
    ' ...
    For Each Restriction In Restrictions
        ' ...
        Set Temp = New clsCommandRestrictionObj
        
        ' ...
        With Temp
            .Name = getName(Restriction)
            .MatchMessage = getMatchMessage(Restriction)
            .RequiredFlags = getFlags(Restriction)
            .RequiredRank = getRank(Restriction)
        End With
        
        ' ...
        getRestrictions.Add Temp
        
        ' ...
        Set Temp = Nothing
    Next
End Function

Private Function getAliases(ByRef AnyNode As IXMLDOMNode) As Collection

    Dim Alias   As IXMLDOMNode
    Dim Aliases As IXMLDOMNodeList
    
    ' ...
    Set getAliases = New Collection
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    Set Aliases = AnyNode.selectNodes("alias")
    
    ' ...
    If ((Aliases Is Nothing) = False) Then
        ' ...
        For Each Alias In Aliases
            getAliases.Add Alias.text
        Next
    End If

End Function

Private Function setAliases(ByRef AnyNode As IXMLDOMNode, ByRef AliasCol As Collection) As Collection

    Dim Alias   As IXMLDOMNode
    Dim Aliases As IXMLDOMNodeList
    Dim i       As Integer
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    Set Aliases = AnyNode.selectNodes("alias")
    
    ' ...
    If ((Aliases Is Nothing) = False) Then
        ' ...
        For Each Alias In Aliases
            AnyNode.removeChild Alias
        Next
    End If
    
    ' ...
    For i = 1 To AliasCol.Count
        ' ...
        Set Alias = _
            AnyNode.appendChild(m_database.createNode("element", "alias", vbNullString))
            
        ' ...
        Alias.text = AliasCol(i)
    Next
    
    ' ...
    Call Save

End Function

Private Function getName(ByRef AnyNode As IXMLDOMNode) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.Attributes.getNamedItem("name")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getName = Temp.text

End Function

Private Function getMatchMessage(ByRef AnyNode As IXMLDOMNode) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("match")

    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    Set Temp = Temp.Attributes.getNamedItem("message")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getMatchMessage = Temp.text

End Function


Private Function getEnabled(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.Attributes.getNamedItem("enabled")
    
    ' ...
    getEnabled = True
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    If (Temp.text = "false") Then
        getEnabled = False
    End If

End Function

Private Function setEnabled(ByRef AnyNode As IXMLDOMNode, ByVal Enabled As Boolean)

    Dim Temp As IXMLDOMNode
    Dim attr As IXMLDOMAttribute
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.Attributes.getNamedItem("enabled")
    
    ' ...
    If (Temp Is Nothing) Then
        ' ...
        Set Temp = _
            AnyNode.Attributes.setNamedItem(m_database.createAttribute("enabled"))
    End If
    
    ' ...
    If (Enabled = True) Then
        Temp.text = "true"
    Else
        Temp.text = "false"
    End If
    
    ' ...
    Call Save

End Function

Private Function getDescription(ByRef AnyNode As IXMLDOMNode) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("documentation/description")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getDescription = Temp.text

End Function

Private Function setDescription(ByRef AnyNode As IXMLDOMNode, ByVal Description As String) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("documentation/description")
    
    ' ...
    If (Temp Is Nothing) Then
        ' ...
        Set Temp = AnyNode.selectSingleNode("documentation")
        
        ' ...
        If (Temp Is Nothing) Then
            ' ...
            Set Temp = _
                AnyNode.appendChild(m_database.createNode("element", "documentation", _
                    vbNullString))
        End If
        
        ' ...
        Set Temp = _
            Temp.appendChild(m_database.createNode("element", "description", vbNullString))
    End If
    
    ' ...
    Temp.text = Description
    
    ' ...
    Call Save

End Function

Private Function getNotes(ByRef AnyNode As IXMLDOMNode) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("documentation/notes")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getNotes = Temp.text

End Function

Private Function setNotes(ByRef AnyNode As IXMLDOMNode, ByVal Notes As String) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("documentation/notes")
    
    ' ...
    If (Temp Is Nothing) Then
        ' ...
        Set Temp = AnyNode.selectSingleNode("documentation")
        
        ' ...
        If (Temp Is Nothing) Then
            ' ...
            Set Temp = _
                AnyNode.appendChild(m_database.createNode("element", "documentation", _
                    vbNullString))
        End If
        
        ' ...
        Set Temp = _
            Temp.appendChild(m_database.createNode("element", "notes", vbNullString))
    End If
    
    ' ...
    Temp.text = Notes
    
    ' ...
    Call Save
    
End Function

Private Function getRank(ByRef AnyNode As IXMLDOMNode) As Integer

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("access/rank")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getRank = CInt(Temp.text)

End Function

Private Function setRank(ByRef AnyNode As IXMLDOMNode, ByVal Rank As Integer) As Integer

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("access/rank")
    
        ' ...
    If (Temp Is Nothing) Then
        ' ...
        Set Temp = AnyNode.selectSingleNode("access")
        
        ' ...
        If (Temp Is Nothing) Then
            ' ...
            Set Temp = _
                AnyNode.appendChild(m_database.createNode("element", "access", _
                    vbNullString))
        End If
        
        ' ...
        Set Temp = _
            Temp.appendChild(m_database.createNode("element", "rank", vbNullString))
    End If

    ' ...
    Temp.text = Rank
    
    ' ...
    Call Save

End Function

Private Function getFlags(ByRef AnyNode As IXMLDOMNode) As String

    Dim Temp As IXMLDOMNode
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("access/flags")
    
    ' ...
    If (Temp Is Nothing) Then
        Exit Function
    End If
    
    ' ...
    getFlags = Temp.text
    
End Function

Private Function setFlags(ByRef AnyNode As IXMLDOMNode, ByVal Flags As String) As String

    Dim Temp As IXMLDOMNode
    Dim flag As IXMLDOMNode
    Dim i    As Integer
    
    ' ...
    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    ' ...
    Set Temp = AnyNode.selectSingleNode("access/rank")
    
        ' ...
    If (Temp Is Nothing) Then
        ' ...
        Set Temp = AnyNode.selectSingleNode("access")
        
        ' ...
        If (Temp Is Nothing) Then
            ' ...
            Set Temp = _
                AnyNode.appendChild(m_database.createNode("element", "access", _
                    vbNullString))
        End If
    End If

    ' ...
    For i = 1 To Len(Flags)
        ' ...
        Set flag = _
            Temp.appendChild(m_database.createNode("element", "flag", vbNullString))
            
        ' ...
        flag.text = Mid$(Flags, i, 1)
    Next i
    
    ' ...
    Call Save
    
End Function
