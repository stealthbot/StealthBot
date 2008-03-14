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

Private m_database       As MSXML2.DOMDocument
Private m_command_node   As MSXML2.IXMLDOMNode
Private m_name           As String
Private m_aliases        As Collection
Private m_required_rank  As Integer
Private m_required_flags As String
Private m_description    As String
Private m_special_notes  As String
Private m_params         As Collection
Private m_enabled        As Boolean
Private m_datasrc        As String

Private Sub Class_Initialize()

    Set m_database = New MSXML2.DOMDocument
    Set m_params = New Collection
    Set m_aliases = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_database = Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Public Function OpenCommand(ByVal strCommand As String, ByVal datasrc As String, Optional ByVal DatabasePath _
    As String = vbNullString) As Boolean
    
    Dim command_access_node   As IXMLDOMNode
    Dim command_documentation As IXMLDOMNode
    Dim command_parameters    As IXMLDOMNodeList
    Dim command_aliases       As IXMLDOMNodeList
    Dim alias                 As IXMLDOMNode
    Dim attrs                 As IXMLDOMAttribute
    
    ' ...
    If (DatabasePath = vbNullString) Then
        ' ...
        DatabasePath = App.Path & "\commands.xml"
    End If
    
    ' ...
    m_database.Load DatabasePath
    
    ' ...
    Set m_command_node = m_database.documentElement

    ' ...
    If (m_command_node Is Nothing) Then
        ' ...
        Exit Function
    End If
    
    If (datasrc = "internal") Then
        ' ...
        Set m_command_node = _
            m_command_node.selectSingleNode("./command[@name='" & strCommand & "']")
    Else
        ' ...
        Set m_command_node = _
            m_command_node.selectSingleNode("./command[@name='" & strCommand & "' and " & _
                "@datasrc='" & datasrc & "']")
    End If
    ' ...
    If (m_command_node Is Nothing) Then
        ' ...
        Exit Function
    End If
        
    ' ...
    m_name = m_command_node.Attributes.getNamedItem("name").text
    
    ' ...
    Set attrs = m_command_node.Attributes.getNamedItem("enabled")
    
    If (Not (attrs Is Nothing)) Then
        If (attrs.text = "true") Then
            m_enabled = True
        Else
            m_enabled = False
        End If
    Else
        m_enabled = True
    End If
    
    Set command_aliases = m_command_node.selectNodes("./alias")
    
    For Each alias In command_aliases
        m_aliases.Add alias.text
    Next
        
    ' ...
    Set command_access_node = m_command_node.selectSingleNode("access")
    
    ' ...
    If (Not (command_access_node Is Nothing)) Then
        Dim command_rank_node  As IXMLDOMNode
        Dim command_flag_nodes As IXMLDOMNodeList
        Dim command_flag_node  As IXMLDOMNode
        
        ' ...
        Set command_rank_node = command_access_node.selectSingleNode("rank")
        Set command_flag_nodes = command_access_node.selectNodes("flag")
        
        ' ...
        If (Not (command_rank_node Is Nothing)) Then
            m_required_rank = command_rank_node.text
        End If
        
        ' ...
        If (Not (command_flag_nodes Is Nothing)) Then
            ' ...
            For Each command_flag_node In command_flag_nodes
                ' ...
                m_required_flags = m_required_flags & _
                    command_flag_node.text
            Next
        End If
    End If
    
    ' ...
    Set command_documentation = m_command_node.selectSingleNode("documentation")
    
    ' ...
    If (Not (command_documentation Is Nothing)) Then
        Dim documentation_description   As IXMLDOMNode
        Dim documentation_special_notes As IXMLDOMNode
        
        ' ...
        Set documentation_description = _
            command_documentation.selectSingleNode("description")
            
        ' ...
        Set documentation_special_notes = _
            command_documentation.selectSingleNode("notes")
            
        ' ...
        If (Not (documentation_description Is Nothing)) Then
            m_description = documentation_description.text
        End If
        
        ' ...
        If (Not (documentation_special_notes Is Nothing)) Then
            m_special_notes = documentation_special_notes.text
        End If
    End If
    
    ' ...
    Set command_parameters = m_command_node.selectNodes("./arguments/argument")
    
    If (Not (command_parameters Is Nothing)) Then
        Dim paramObj       As clsCommandParamsObj
        Dim restrictionObj As clsCommandRestrictionObj
    
        Dim parameter      As IXMLDOMNode
        Dim Data           As IXMLDOMNode
        Dim restrictions   As IXMLDOMNodeList
        Dim restriction    As IXMLDOMNode
        
        For Each parameter In command_parameters
            ' ...
            Set paramObj = New clsCommandParamsObj
            
            Set attrs = parameter.Attributes.getNamedItem("name")
            
            If (Not (attrs Is Nothing)) Then
                paramObj.Name = attrs.text
            End If
            
            Set attrs = parameter.Attributes.getNamedItem("optional")
            Set Data = parameter.selectSingleNode("./data")
            Set restrictions = parameter.selectNodes("./restriction")
            
            If (Not (attrs Is Nothing)) Then
                If (attrs.text = "true") Then
                    paramObj.IsOptional = True
                End If
            End If
            
            If (Not (Data Is Nothing)) Then
                paramObj.dataType = Data.Attributes.getNamedItem("type").text
                
                If (paramObj.dataType = "string") Then
                    paramObj.pattern = Data.Attributes.getNamedItem("pattern").text
                ElseIf (paramObj.dataType = "number") Then
                
                End If
            End If
            
            For Each restriction In restrictions
                ' ...
                Set restrictionObj = New clsCommandRestrictionObj
            
                With restrictionObj
                    .Name = restriction.Attributes.getNamedItem("name").text
                    .MatchMessage = _
                        restriction.selectSingleNode("./match[@message]").Attributes.getNamedItem("message").text
                    .RequiredRank = _
                        restriction.selectSingleNode("./access/rank").text
                End With
                
                ' ...
                Call paramObj.restrictions.Add(restrictionObj)
                
                ' ...
                Set restrictionObj = Nothing
            Next
            
            ' ...
            Call m_params.Add(paramObj)
            
            ' ...
            Set paramObj = Nothing
        Next
    End If
End Function

Public Property Get Name() As String

    Name = m_name
    
End Property

Public Property Get datasrc() As String

    datasrc = m_datasrc
    
End Property

Public Property Let datasrc(str As String)

    m_datasrc = str
    
End Property

Public Property Get Aliases() As Collection

    ' ...
    Set Aliases = m_aliases

End Property

Public Property Get IsEnabled() As Boolean

    IsEnabled = m_enabled
    
End Property

Public Property Let IsEnabled(bln As Boolean)

    m_enabled = bln
    
End Property

Public Property Let Name(strName As String)

    m_name = Name
    
End Property

Public Property Get RequiredRank() As Integer

    RequiredRank = m_required_rank
    
End Property

Public Property Get RequiredFlags() As String

    RequiredFlags = m_required_flags
    
End Property

Public Property Get description() As String

    description = m_description
    
End Property

Public Property Let description(strDescription As String)

    m_description = strDescription
    
End Property

Public Property Get Notes() As String

    Notes = m_special_notes
    
End Property

Public Property Let Notes(strNotes As String)

    m_special_notes = strNotes
    
End Property

Public Property Get params() As Collection

    Set params = m_params

End Property

Public Function Save() As Boolean
    Dim node     As IXMLDOMNode
    Dim params   As IXMLDOMNode
    Dim tmp_attr As IXMLDOMAttribute
    Dim tmp_node As IXMLDOMNode
    Dim i        As Integer
    
    tmp_attr.text = m_name

    node.Attributes.setNamedItem tmp_attr
    
    'tmp_node.nodeName = "arguments"
    
    'node.appendChild "arguments"
    
    For i = 1 To params
        
        
    Next i

    ' ...
    Call Delete
    
    
    m_command_node.parentNode.appendChild node
    
End Function

Public Function Delete() As Boolean

    ' ...
    m_command_node.parentNode.removeChild m_command_node

End Function
