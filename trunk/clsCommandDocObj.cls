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
    
    m_defaultXMLPath = GetFilePath(FILE_COMMANDS)
    
    Call OpenDatabase(, True)
    
    '// 06/23/2009 JSM - initializing properties
    Call ResetInstance
    
End Sub

Private Sub Class_Terminate()

    Set m_database = Nothing
    Set m_params = Nothing
    Set m_aliases = Nothing
    
End Sub

Private Sub ResetInstance()
        
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

Public Function OpenDatabase(Optional ByVal DatabasePath As String = vbNullString, Optional ByVal forceLoad As Boolean = False, Optional ByVal Database As DOMDocument60 = Nothing)

    If Database Is Nothing Then Set Database = m_database

    If (DatabasePath = vbNullString) Then
        DatabasePath = m_defaultXMLPath
    End If

    If Not FileExists(DatabasePath) And Database.childNodes.Length = 0 Then
        Set Database.documentElement = Database.createElement("commands")
    ElseIf forceLoad = True Or Database.childNodes.Length = 0 Then
        Database.Load DatabasePath
    End If

End Function

Private Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

'// this function will open a clsCommandDocObj and attempt to seek to the specified command
Public Function OpenCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional ByVal EnabledOnly As Boolean = False, Optional ByVal Database As DOMDocument60 = Nothing) As Boolean

    Dim command_access_node   As IXMLDOMNode
    Dim command_documentation As IXMLDOMNode
    Dim command_parameters    As IXMLDOMNodeList
    Dim command_aliases       As IXMLDOMNodeList
    Dim Alias                 As IXMLDOMNode
    Dim attrs                 As IXMLDOMAttribute
    Dim xpath                 As String

    If Database Is Nothing Then Set Database = m_database

    Call OpenDatabase(, , Database)
    Call ResetInstance

    If (Database.documentElement Is Nothing) Then
        OpenCommand = False
        Exit Function
    End If

    xpath = clsCommandObj.GetCommandXPath(strCommand, strScriptOwner, EnabledOnly)

    Set m_command_node = Database.selectSingleNode(xpath)

    If (m_command_node Is Nothing) Then
        'Lets try an alias
        xpath = clsCommandObj.GetCommandXPath(ConvertAlias(strCommand, Database), strScriptOwner, EnabledOnly)
        Set m_command_node = Database.documentElement.selectSingleNode(xpath)

        If (m_command_node Is Nothing) Then
            OpenCommand = False
            Exit Function
        End If
    End If

    Set m_aliases = GetAliases(m_command_node)
    Set m_params = GetParameters(m_command_node)

    OpenCommand = True

End Function

'// 06/13/2009 JSM - Created function
Public Function CreateCommand(ByVal strCommand As String, Optional ByVal strScriptOwner As String = vbNullString, Optional ByVal Database As DOMDocument60 = Nothing) As Boolean

    Dim l_command             As IXMLDOMElement
    Dim l_aliases             As IXMLDOMElement
    Dim l_documentation       As IXMLDOMElement
    Dim l_description         As IXMLDOMElement
    Dim l_Parameters          As IXMLDOMElement
    Dim l_Access              As IXMLDOMElement

    If Database Is Nothing Then Set Database = m_database

    If (Not clsCommandObj.IsValidCommandName(strCommand)) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Error: The provided command name was not valid."
        Exit Function
    End If

    Call OpenDatabase(, , Database)

    '// create elements
    Set l_command = Database.createElement("command")
    Set l_aliases = Database.createElement("aliases")
    Set l_documentation = Database.createElement("documentation")
    Set l_description = Database.createElement("description")
    Set l_Parameters = Database.createElement("arguments")
    Set l_Access = Database.createElement("access")

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
    Call Database.documentElement.appendChild(l_command)

End Function

Private Function ConvertAlias(ByVal cmdName As String, ByVal Database As DOMDocument60) As String
    On Error GoTo ERROR_HANDLER
    
    If Database Is Nothing Then Set Database = m_database
    Dim xpath As String
    Dim AliasNode As IXMLDOMNode

    If (Len(cmdName) > 0) Then

        If (InStr(1, cmdName, "'", vbBinaryCompare) > 0) Then
            Exit Function
        End If

        cmdName = Replace(cmdName, "\", "\\")
        cmdName = Replace(cmdName, "'", "&apos;")

        '// 09/03/2008 JSM - Modified code to use the <aliases> element
        xpath = StringFormat("/commands/command/aliases/alias[translate(text(), '{0}', '{1}')='{2}']", _
                UCase$(AZ), LCase$(AZ), LCase$(cmdName))
        Set AliasNode = Database.selectSingleNode(xpath)

        If (Not (AliasNode Is Nothing)) Then
            '// 09/03/2008 JSM - Modified code to use the <aliases> element
            ConvertAlias = GetName(AliasNode.parentNode.parentNode)

            Exit Function
        End If
    End If

    ConvertAlias = cmdName

    Exit Function
ERROR_HANDLER:

    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error: XML Database Processor has encountered an error " & _
        "during alias lookup.")

    ConvertAlias = cmdName

    Exit Function

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
    Dim Msg As String

    Msg = "The following errors were detected with commands.xml:"

    For i = 1 To colErrorList.Count
        Msg = StringFormat("{0}{2}{1}", Msg, colErrorList(i), vbNewLine)
    Next i

    frmChat.AddChat RTBColors.ErrorMessageText, Msg

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

            Dim CommandName As String
            Dim ownerName As String
            Dim Enabled As String

            '// default our values
            CommandName = oCommand.Attributes.getNamedItem("name").Text
            ownerName = vbNullString
            Enabled = "1"

            '// make sure name != ""
            If Len(CommandName) = 0 Then
                colErrorList.Add "A command element is missing the name attribute."
                Exit Do
            End If

            '// make sure owner != ""
            If Not (oCommand.Attributes.getNamedItem("owner") Is Nothing) Then
                ownerName = oCommand.Attributes.getNamedItem("owner").Text
                If Len(ownerName) = 0 Then
                    colErrorList.Add StringFormat("An owner attribute cannot be empty on command element. Command name: {0}", CommandName)
                    Exit Do
                End If
            End If

            '// make sure enabled attribute is 0 or 1
            If Not (oCommand.Attributes.getNamedItem("enabled") Is Nothing) Then
                Enabled = oCommand.Attributes.getNamedItem("enabled").Text
                If Enabled <> "0" And Enabled <> "1" Then
                    colErrorList.Add StringFormat("If present, an enabled attribute must be equal to 0 or 1 on a command element. Command name: {0}", CommandName)
                    Exit Do
                End If
            End If

            '// make sure only 1 command is enabled if there are similar names
            If Enabled = "1" Then
                If uniqueEnabledCommands.Exists(CommandName) Then
                    colErrorList.Add StringFormat("Only 1 command element named the same thing can have no enabled attribute or an enabled attribute equal to 1. Command name: {0}", CommandName)
                    Exit Do
                End If
                uniqueEnabledCommands.Add CommandName, ownerName
            End If

            '// make sure commands with the same name of separate owners
            If uniqueCommandOwners.Exists(CommandName & "|" & ownerName) Then
                colErrorList.Add StringFormat("Commands with equal name attributes must have different unique owner attributes. Command name: {0}", CommandName)
                Exit Do
            End If
            uniqueEnabledCommands.Add CommandName & "|" & ownerName, CommandName & "|" & ownerName

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

                    Dim ArgumentName As String
                    Dim DataType As String

                    '// default our values
                    ArgumentName = oArgument.Attributes.getNamedItem("name").Text

                    '// make sure name != ""
                    If Len(ArgumentName) = 0 Then
                        colErrorList.Add "An argument element is missing the name attribute."
                        Exit Do
                    End If

                    '// make sure name attribute is unique
                    If uniqueEnabledArguments.Exists(ArgumentName) Then
                        colErrorList.Add "Argument elements for a command must have unique name attributes."
                        Exit Do
                    End If
                    uniqueEnabledArguments.Add ArgumentName, CommandName

                    '// make sure datatype attribute is string or word or numeric or number
                    If Not (oArgument.Attributes.getNamedItem("datatype") Is Nothing) Then
                        DataType = LCase(oArgument.Attributes.getNamedItem("datatype").Text)
                        If DataType <> "string" And DataType <> "word" And DataType <> "numeric" And DataType <> "number" Then
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

                            Dim RestrictionName As String

                            '// default our values
                            RestrictionName = oRestriction.Attributes.getNamedItem("name").Text

                            '// make sure name != ""
                            If Len(RestrictionName) = 0 Then
                                colErrorList.Add "A restriction element is missing the name attribute."
                                Exit Do
                            End If

                            '// make sure name attribute is unique
                            If uniqueEnabledRestrictions.Exists(RestrictionName) Then
                                colErrorList.Add "Restriction elements for an argument must have unique name attributes."
                                Exit Do
                            End If
                            uniqueEnabledRestrictions.Add RestrictionName, ArgumentName

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

    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObjStatic.CommandsSanityCheck().")

End Function

Public Function Save(Optional ByVal WriteToFile As Boolean = True) As Boolean

    On Error GoTo ERROR_HANDLER

    Dim colErrorList As Collection

    SetAliases m_command_node, m_aliases
    SetParameters m_command_node, m_params

    If WriteToFile = True Then

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
        Call WriteDocToFile(GetFilePath(FILE_COMMANDS))

    End If

    Save = True

    Exit Function

ERROR_HANDLER:

    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObj.Save().")

End Function

Private Sub WriteDocToFile(ByVal FilePath As String)

    On Error GoTo ERROR_HANDLER

    Dim reader As SAXXMLReader60 '// create the SAX reader
    Dim writer As MXXMLWriter60  '// create the XML writer

    Dim iFileNo As Integer
    
    Set reader = New SAXXMLReader60
    Set writer = New MXXMLWriter60

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
    Open FilePath For Output As #iFileNo

        Print #iFileNo, Replace(writer.output, Chr$(9), "  ", , , vbBinaryCompare)

    '// close the file
    Close #iFileNo

    Set reader = Nothing
    Set writer = Nothing

    Exit Sub

ERROR_HANDLER:

    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObj.writeDocToFile().")

End Sub


Public Sub Delete(Optional ByVal WriteToFile As Boolean = True)

    If (m_command_node Is Nothing) Then
        Exit Sub
    End If

    If (LenB(Me.Owner) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Error: You can not delete an internal command."
        Exit Sub
    End If

    m_command_node.parentNode.removeChild m_command_node

    Call Save(WriteToFile)

End Sub

Private Function GetParameters(ByRef CommandNode As IXMLDOMNode) As Collection

    Dim ParameterNode    As IXMLDOMNode
    Dim ParameterNodes   As IXMLDOMNodeList
    Dim Parameter        As clsCommandParamsObj
    Dim RestrictionsNode As IXMLDOMNode

    Set GetParameters = New Collection

    If (CommandNode Is Nothing) Then
        Exit Function
    End If

    Set ParameterNodes = CommandNode.selectNodes("arguments/argument")

    For Each ParameterNode In ParameterNodes
        Set Parameter = New clsCommandParamsObj

        With Parameter
            .Name = GetName(ParameterNode)
            .Restrictions = GetRestrictions(ParameterNode)

            .IsOptional = GetIsOptional(ParameterNode)
            .DataType = GetDataType(ParameterNode)

            .MatchMessage = GetMatchMessage(ParameterNode)
            .MatchError = GetError(ParameterNode)
            .MatchCaseSensitive = GetMatchCase(ParameterNode)

            .Description = GetDescription(ParameterNode)
            .SpecialNotes = GetNotes(ParameterNode)

            Set RestrictionsNode = ParameterNode.selectSingleNode("restrictions")
            If (Not RestrictionsNode Is Nothing) Then
                .RestrictionsSharedError = GetSharedError(RestrictionsNode)
            End If
            Set RestrictionsNode = Nothing
        End With

        GetParameters.Add Parameter

        Set Parameter = Nothing
    Next ParameterNode
    Set ParameterNodes = Nothing

End Function

Private Function GetRestrictions(ByRef ParameterNode As IXMLDOMNode) As Collection

    Dim RestrictionNode  As IXMLDOMNode
    Dim RestrictionsNode As IXMLDOMNode
    Dim Restriction      As clsCommandRestrictionObj

    Set GetRestrictions = New Collection

    If (ParameterNode Is Nothing) Then
        Exit Function
    End If

    Set RestrictionsNode = ParameterNode.selectSingleNode("restrictions")

    If (RestrictionsNode Is Nothing) Then
        Exit Function
    End If

    For Each RestrictionNode In RestrictionsNode.selectNodes("restriction")
        Set Restriction = New clsCommandRestrictionObj
        With Restriction
            .Name = GetName(RestrictionNode)

            .Description = GetDescription(RestrictionNode)
            .SpecialNotes = GetNotes(RestrictionNode)

            .MatchMessage = GetMatchMessage(RestrictionNode)
            .MatchError = GetError(RestrictionNode)
            .MatchCaseSensitive = GetMatchCase(RestrictionNode)

            .RequiredFlags = GetFlags(RestrictionNode)
            .RequiredRank = GetRank(RestrictionNode)

            .Fatal = GetFatal(RestrictionNode)
        End With
        GetRestrictions.Add Restriction
        Set Restriction = Nothing
    Next RestrictionNode
End Function

Private Function GetAliases(ByRef AnyNode As IXMLDOMNode) As Collection

    Dim Alias   As IXMLDOMNode
    Dim Aliases As IXMLDOMNodeList
    
    Set GetAliases = New Collection
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    '// 09/03/2008 JSM - Modified code to use the <aliases> element
    Set Aliases = AnyNode.selectNodes("aliases/alias")
    
    If ((Aliases Is Nothing) = False) Then
        For Each Alias In Aliases
            GetAliases.Add Alias.Text
        Next Alias
    End If

End Function

Private Sub SetParameters(ByRef AnyNode As IXMLDOMNode, ByRef ParameterCol As Collection)

    Dim ArgumentNode                  As IXMLDOMNode
    Dim ArgumentsNode                 As IXMLDOMNode
    Dim ArgumentsOptionalNode         As IXMLDOMNode
    Dim ArgumentsOptionalStringNode   As IXMLDOMNode
    Dim ArgumentsOptionalNumberNode   As IXMLDOMNode
    Dim RestrictionNode               As IXMLDOMNode

    Dim NewNode                       As IXMLDOMNode
    Dim MatchNode                     As IXMLDOMNode
    Dim Node                          As IXMLDOMNode
    Dim Parameter                     As clsCommandParamsObj
    Dim Restriction                   As clsCommandRestrictionObj
    Dim i                             As Integer

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    '// get the arguments node
    Set ArgumentsNode = AnyNode.selectSingleNode("arguments")
    If Not (ArgumentsNode Is Nothing) Then
        For Each Node In ArgumentsNode.childNodes
            ArgumentsNode.removeChild Node
        Next Node
    End If

    '// add the the new arguments node if necessary
    If ParameterCol.Count > 0 Then

        '// loop through the parameters
        For i = 1 To ParameterCol.Count

            Set Parameter = ParameterCol(i)

            Set ArgumentNode = m_database.createNode("element", "argument", vbNullString)

            '// @name
            Set NewNode = m_database.createNode("attribute", "name", vbNullString)
            NewNode.Text = Parameter.Name
            ArgumentNode.Attributes.setNamedItem NewNode

            '// documentation
            Set NewNode = m_database.createNode("element", "documentation", vbNullString)
            ArgumentNode.appendChild NewNode

            '// create the match element if necessary
            If Len(Parameter.MatchMessage) <> 0 Then
                '// add the match element
                Set MatchNode = ArgumentNode.appendChild(m_database.createNode("element", "match", vbNullString))

                '// match[@message]
                Set NewNode = m_database.createNode("attribute", "message", vbNullString)
                NewNode.Text = Parameter.MatchMessage
                MatchNode.Attributes.setNamedItem NewNode

                '// match[@case-sensitive]
                Set NewNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
                NewNode.Text = IIf(Parameter.MatchCaseSensitive, "1", "0")
                MatchNode.Attributes.setNamedItem NewNode

            End If

            '// create the error element if necessary
            If Len(Parameter.MatchError) <> 0 Then
                '// add the error element
                Set NewNode = ArgumentNode.appendChild(m_database.createNode("element", "error", vbNullString))
                NewNode.Text = Parameter.MatchError
            End If

            '// documentation/description
            SetDescription ArgumentNode, Parameter.Description

            '// documentation/specialnotes
            SetNotes ArgumentNode, Parameter.SpecialNotes

            '// restrictions
            If Parameter.Restrictions.Count > 0 Then
                SetRestrictions ArgumentNode, Parameter.Restrictions, Parameter.RestrictionsSharedError
            End If

            If Parameter.IsOptional Then
                Set NewNode = m_database.createNode("attribute", "optional", vbNullString)
                NewNode.Text = "1"
                ArgumentNode.Attributes.setNamedItem NewNode
            End If
            
            If (Not StrComp(Parameter.DataType, "string", vbTextCompare) = 0) Then
                Set NewNode = m_database.createNode("attribute", "type", vbNullString)
                NewNode.Text = Parameter.DataType
                ArgumentNode.Attributes.setNamedItem NewNode
            End If
            
            ArgumentsNode.appendChild ArgumentNode
        Next i
    End If
End Sub

Private Sub SetRestrictions(ByRef AnyNode As IXMLDOMNode, ByRef RestrictionCol As Collection, ByVal SharedError As String)

    Dim RestrictionsNode                 As IXMLDOMNode
    Dim RestrictionNode                  As IXMLDOMNode
    Dim SharedErrorNode                  As IXMLDOMNode

    Dim NewNode                          As IXMLDOMNode
    Dim MatchNode                        As IXMLDOMNode
    Dim Node                             As IXMLDOMNode
    Dim Restriction                      As clsCommandRestrictionObj
    Dim i                                As Integer

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    '// add our restrictions node
    Set RestrictionsNode = AnyNode.appendChild(m_database.createNode("element", "restrictions", vbNullString))

    Set SharedErrorNode = RestrictionsNode.appendChild(m_database.createNode("element", "shared-error", vbNullString))
    SharedErrorNode.Text = SharedError

    '// loop through argument restrictions
    For Each Restriction In RestrictionCol

        '// create our restriction node
        Set RestrictionNode = m_database.createNode("element", "restriction", vbNullString)
        RestrictionsNode.appendChild RestrictionNode

        '// name attribute
        Set NewNode = m_database.createNode("attribute", "name", vbNullString)
        NewNode.Text = Restriction.Name
        RestrictionNode.Attributes.setNamedItem NewNode

        If (Restriction.Fatal) Then
            Set NewNode = m_database.createNode("attribute", "nonfatal", vbNullString)
            NewNode.Text = IIf(Restriction.Fatal, "0", "1")
            RestrictionNode.Attributes.setNamedItem NewNode
        End If

        '// documentation/description
        SetDescription RestrictionNode, Restriction.Description

        '// documentation/specialnotes
        SetNotes RestrictionNode, Restriction.SpecialNotes

        '// create the match element if necessary
        If Len(Restriction.MatchMessage) > 0 Then
            '// add the match element
            Set MatchNode = RestrictionNode.appendChild(m_database.createNode("element", "match", vbNullString))

            '// set the message attribute
            '// match[@message]
            Set NewNode = m_database.createNode("attribute", "message", vbNullString)
            NewNode.Text = Restriction.MatchMessage
            MatchNode.Attributes.setNamedItem NewNode

            '// match[@case-sensitive]
            Set NewNode = m_database.createNode("attribute", "case-sensitive", vbNullString)
            NewNode.Text = IIf(Restriction.MatchCaseSensitive, "1", "0")
            MatchNode.Attributes.setNamedItem NewNode

        End If

        '// create the error element if necessary
        If Len(Restriction.MatchError) <> 0 Then
            '// add the error element
            Set NewNode = RestrictionNode.appendChild(m_database.createNode("element", "error", vbNullString))
            NewNode.Text = Restriction.MatchError
        End If

        '// create the access nodes
        SetRank RestrictionNode, Restriction.RequiredRank
        SetFlags RestrictionNode, Restriction.RequiredFlags

    Next Restriction

End Sub

Private Sub SetAliases(ByRef AnyNode As IXMLDOMNode, ByRef AliasCol As Collection)

    Dim Alias   As IXMLDOMNode
    Dim Aliases As IXMLDOMNodeList
    Dim i       As Integer

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    '// 09/03/2008 JSM - Modified code to use the <aliases> element
    Set Aliases = AnyNode.selectNodes("aliases/alias")

    If ((Aliases Is Nothing) = False) Then
        For Each Alias In Aliases
            '// 09/03/2008 JSM - Modified code to use the <aliases> element
            AnyNode.selectSingleNode("aliases").removeChild Alias
        Next Alias
    End If

    For i = 1 To AliasCol.Count
        '// 09/03/2008 JSM - Modified code to use the <aliases> element
        Set Alias = _
            AnyNode.selectSingleNode("aliases").appendChild(m_database.createNode("element", "alias", vbNullString))

        Alias.Text = AliasCol(i)
    Next i

End Sub

Private Function GetName(ByRef AnyNode As IXMLDOMNode) As String

    Dim NameNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set NameNode = AnyNode.Attributes.getNamedItem("name")
    If (NameNode Is Nothing) Then
        Exit Function
    End If
    GetName = NameNode.Text

End Function

Private Function GetOwner(ByRef AnyNode As IXMLDOMNode) As String
    Dim OwnerNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set OwnerNode = AnyNode.Attributes.getNamedItem("owner")
    If (OwnerNode Is Nothing) Then
        Exit Function
    End If
    
    GetOwner = OwnerNode.Text

End Function

Private Function GetDataType(ByRef AnyNode As IXMLDOMNode) As String
    Dim DataTypeNode As IXMLDOMNode

    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    GetDataType = "string"

    Set DataTypeNode = AnyNode.Attributes.getNamedItem("type")
    If (DataTypeNode Is Nothing) Then
        Exit Function
    End If

    GetDataType = DataTypeNode.Text

End Function

Private Function GetMatchMessage(ByRef AnyNode As IXMLDOMNode) As String

    Dim MatchNode   As IXMLDOMNode
    Dim MessageNode As IXMLDOMNode

    If (AnyNode Is Nothing) Then
        Exit Function
    End If

    Set MatchNode = AnyNode.selectSingleNode("match")
    If (MatchNode Is Nothing) Then
        Exit Function
    End If

    Set MessageNode = MatchNode.Attributes.getNamedItem("message")
    If (MessageNode Is Nothing) Then
        Exit Function
    End If

    GetMatchMessage = MessageNode.Text

End Function

Private Function GetMatchCase(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim MatchNode As IXMLDOMNode
    Dim CaseNode  As IXMLDOMNode
    GetMatchCase = False
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set MatchNode = AnyNode.selectSingleNode("match")
    If (MatchNode Is Nothing) Then
        Exit Function
    End If
    
    Set CaseNode = MatchNode.Attributes.getNamedItem("case-sensitive")
    If (CaseNode Is Nothing) Then
        Exit Function
    End If
    
    GetMatchCase = (CaseNode.Text = "1")

End Function

Private Function GetIsOptional(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim OptionalNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    Set OptionalNode = AnyNode.Attributes.getNamedItem("optional")
    GetIsOptional = False
    
    If (OptionalNode Is Nothing) Then
        Exit Function
    End If
    
    If (OptionalNode.Text = "1") Then
        GetIsOptional = True
    End If

End Function

Private Function GetEnabled(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim EnabledNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    Set EnabledNode = AnyNode.Attributes.getNamedItem("enabled")
    GetEnabled = True
    
    If (EnabledNode Is Nothing) Then
        Exit Function
    End If
    
    If (EnabledNode.Text = "0") Then
        GetEnabled = False
    End If

End Function

Private Sub SetEnabled(ByRef AnyNode As IXMLDOMNode, ByVal Enabled As Boolean)

    Dim EnabledNode As IXMLDOMNode
    Dim attr As IXMLDOMAttribute

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    Set EnabledNode = AnyNode.Attributes.getNamedItem("enabled")
    If (EnabledNode Is Nothing) Then
        Set EnabledNode = AnyNode.Attributes.setNamedItem(m_database.createAttribute("enabled"))
    End If

    If (Enabled = True) Then
        EnabledNode.Text = "1"
    Else
        EnabledNode.Text = "0"
    End If

End Sub

Private Function GetDescription(ByRef AnyNode As IXMLDOMNode) As String

    Dim DescNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set DescNode = AnyNode.selectSingleNode("documentation/description")
    If (DescNode Is Nothing) Then
        Exit Function
    End If
    
    GetDescription = Trim$(DescNode.Text)

End Function

Private Function GetError(ByRef AnyNode As IXMLDOMNode) As String

    Dim ErrorNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set ErrorNode = AnyNode.selectSingleNode("error")
    If (ErrorNode Is Nothing) Then
        Exit Function
    End If
    
    GetError = Trim$(ErrorNode.Text)

End Function

Private Function GetSharedError(ByRef AnyNode As IXMLDOMNode) As String

    Dim ErrorNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set ErrorNode = AnyNode.selectSingleNode("shared-error")
    If (ErrorNode Is Nothing) Then
        Exit Function
    End If
    
    GetSharedError = Trim$(ErrorNode.Text)

End Function

'Syntax: .add <username> [rank] [attributes]
Private Function GetSyntaxString(Optional IsLocal As Boolean = False) As String

    Dim retVal  As String
    Dim P       As clsCommandParamsObj

    If (m_command_node Is Nothing) Then
        Exit Function
    End If

    '// add the command name
    retVal = vbNullString

    '// add the parameters
    For Each P In Me.Parameters
        retVal = StringFormat("{0} {1}", retVal, P.ToString)
    Next P

    retVal = StringFormat("{0}{1}", ToString(IsLocal), retVal)

    GetSyntaxString = retVal

End Function

'alias1, alias2, alias3
Private Function GetAliasString()
    Dim retVal As String
    Dim i As Integer
    If (Aliases.Count > 0) Then
        For i = 1 To Aliases.Count
            retVal = StringFormat("{0}{1}{2}", _
                retVal, _
                CStr(Aliases.Item(i)), _
                IIf(i = Aliases.Count, vbNullString, ", "))
        Next i
    End If
    GetAliasString = retVal
End Function

Private Function GetRequirementsString(Optional bShort As Boolean = False)

    Dim retVal    As String
    Dim DetailReq As String
    Dim StateStr  As String
    
    If (m_command_node Is Nothing) Then
        Exit Function
    End If

    retVal = vbNullString
    StateStr = vbNullString

    If bShort Then
        DetailReq = "Requires"
    Else
        DetailReq = "Command requires"
    End If

    If (LenB(Me.Owner) > 0) Then
        If (modScripting.IsScriptEnabled(Me.Owner) = False) Then StateStr = " (disabled)"
        retVal = StringFormat("From script ""{0}""{1}. ", Me.Owner, StateStr)
        If (Not bShort) Then retVal = retVal & vbCrLf
    End If

    If (Me.RequiredRank > -1) Or (Len(Me.RequiredFlags) > 0) Then
        '// available outside the console
        '// add access requirements if necessary
        If (Me.RequiredRank > -1) Then
            retVal = StringFormat("{0}{2} rank {1}", retVal, Me.RequiredRank, DetailReq)
            If (Len(Me.RequiredFlags) = 0) Then
                retVal = StringFormat("{0}.", retVal)
            End If
        End If
        
        '// add flag requirements, if necessary, taking into account the existinance of an access requirements
        If (Len(Me.RequiredFlags) <> 0) Then
            If (Me.RequiredRank > -1) Then
                retVal = StringFormat("{0} or one of the following flags: {1}.", retVal, Me.RequiredFlags)
            Else
                retVal = StringFormat("{0}{2} one of the following flags: {1}.", retVal, Me.RequiredFlags, DetailReq)
            End If
        End If
    
    Else
        '// console only, lets show the generic blurb
        If (bShort) Then
            retVal = StringFormat("{0}Only available to the console.", retVal)
        Else
            retVal = StringFormat("{0}Command is only available to the console. Set a rank or flag requirement to make this command available to external sources.", retVal)
        End If
    End If

    GetRequirementsString = retVal

End Function

Private Sub SetDescription(ByRef AnyNode As IXMLDOMNode, ByVal Description As String)

    Dim DocNode  As IXMLDOMNode
    Dim DescNode As IXMLDOMNode

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    Set DescNode = AnyNode.selectSingleNode("documentation/description")
    If (DescNode Is Nothing) Then
        Set DocNode = AnyNode.selectSingleNode("documentation")
        If (DocNode Is Nothing) Then
            Set DocNode = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
        End If
        Set DescNode = DocNode.appendChild(m_database.createNode("element", "description", vbNullString))
    End If

    DescNode.Text = Description

End Sub

Private Function GetNotes(ByRef AnyNode As IXMLDOMNode) As String

    Dim NotesNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set NotesNode = AnyNode.selectSingleNode("documentation/specialnotes")
    If (NotesNode Is Nothing) Then
        Exit Function
    End If
    GetNotes = Trim$(NotesNode.Text)

End Function

Private Sub SetNotes(ByRef AnyNode As IXMLDOMNode, ByVal Notes As String)

    Dim DocNode   As IXMLDOMNode
    Dim NotesNode As IXMLDOMNode

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    Set NotesNode = AnyNode.selectSingleNode("documentation/specialnotes")
    If (NotesNode Is Nothing) Then
        Set DocNode = AnyNode.selectSingleNode("documentation")
        If (DocNode Is Nothing) Then
            Set DocNode = AnyNode.appendChild(m_database.createNode("element", "documentation", vbNullString))
        End If

        Set NotesNode = DocNode.appendChild(m_database.createNode("element", "specialnotes", vbNullString))
    End If

    NotesNode.Text = Notes

End Sub

Private Function GetRank(ByRef AnyNode As IXMLDOMNode) As Integer

    Dim RankNode As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set RankNode = AnyNode.selectSingleNode("access/rank")
    
    If (RankNode Is Nothing) Then
        GetRank = -1
        Exit Function
    End If
    
    If (RankNode.Text = vbNullString) Then
        GetRank = -1
        Exit Function
    End If
    
    GetRank = CInt(Val(RankNode.Text))

End Function

Private Function GetFatal(ByRef AnyNode As IXMLDOMNode) As Boolean

    Dim FatalNode As IXMLDOMNode
    GetFatal = True
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set FatalNode = AnyNode.Attributes.getNamedItem("nonfatal")
    If (FatalNode Is Nothing) Then
        Exit Function
    End If
    GetFatal = (Not FatalNode.Text = "1")
End Function

Private Sub SetRank(ByRef AnyNode As IXMLDOMNode, ByVal Rank As Integer)

    Dim AccessNode As IXMLDOMNode
    Dim RankNode   As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Sub
    End If
    
    Set AccessNode = AnyNode.selectSingleNode("access")
    
    If (AccessNode Is Nothing) Then
        Set AccessNode = AnyNode.appendChild(m_database.createNode("element", "access", vbNullString))
    End If
    
    Set RankNode = AccessNode.selectSingleNode("rank")
    
    If (Rank >= 0) Then
        If (RankNode Is Nothing) Then
            Set RankNode = AccessNode.appendChild(m_database.createNode("element", "rank", vbNullString))
        End If
        RankNode.Text = Rank
    Else
        If (Not RankNode Is Nothing) Then
            AccessNode.removeChild RankNode
        End If
    End If

End Sub

Private Function GetFlags(ByRef AnyNode As IXMLDOMNode) As String

    Dim FlagNodes As IXMLDOMNodeList
    Dim FlagNode  As IXMLDOMNode
    
    If (AnyNode Is Nothing) Then
        Exit Function
    End If
    
    Set FlagNodes = AnyNode.selectNodes("access/flags/flag")
    
    If (FlagNodes Is Nothing) Then
        Exit Function
    End If
    
    For Each FlagNode In FlagNodes
        GetFlags = GetFlags & FlagNode.Text
    Next FlagNode
    
End Function

Private Sub SetFlags(ByRef AnyNode As IXMLDOMNode, ByVal Flags As String)

    Dim AccessNode As IXMLDOMNode
    Dim FlagsNode  As IXMLDOMNode
    Dim FlagNode   As IXMLDOMNode
    Dim i    As Integer

    If (AnyNode Is Nothing) Then
        Exit Sub
    End If

    Set AccessNode = AnyNode.selectSingleNode("access")

    If (AccessNode Is Nothing) Then
        Set AccessNode = AnyNode.appendChild(m_database.createNode("element", "access", vbNullString))
    End If

    Set FlagsNode = AccessNode.selectSingleNode("flags")

    If (Len(Flags) > 0) Then
        If (FlagsNode Is Nothing) Then
            Set FlagsNode = AccessNode.appendChild(m_database.createNode("element", "flags", vbNullString))
        End If

        For i = FlagsNode.childNodes.Length - 1 To 0 Step -1
            FlagsNode.removeChild FlagsNode.childNodes(i)
        Next i

        For i = 1 To Len(Flags)
            Set FlagNode = _
                FlagsNode.appendChild(m_database.createNode("element", "flag", vbNullString))

            FlagNode.Text = Mid$(Flags, i, 1)
        Next i
    Else
        If (Not FlagsNode Is Nothing) Then
            AccessNode.removeChild FlagsNode
        End If
    End If

End Sub

Public Property Get Name() As String
    
    Name = GetName(m_command_node)

End Property

Public Property Get Owner() As String
    
    Owner = GetOwner(m_command_node)
    
End Property

Public Property Get Aliases() As Collection

    If (m_aliases Is Nothing) Then
        Set m_aliases = GetAliases(m_command_node)
    End If
    
    Set Aliases = m_aliases
    
End Property

Public Property Get IsEnabled() As Boolean
    
    IsEnabled = GetEnabled(m_command_node)

End Property

Public Property Let IsEnabled(bln As Boolean)
    
    SetEnabled m_command_node, bln

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
    
    SetRank m_command_node, Rank

End Property

Public Property Get RequiredFlags() As String
    
    RequiredFlags = GetFlags(m_command_node)

End Property

Public Property Let RequiredFlags(ByVal str As String)
   
   SetFlags m_command_node, str

End Property

Public Property Get Description() As String
    
    Description = GetDescription(m_command_node)

End Property

Public Property Let Description(ByVal strDescription As String)
    
    SetDescription m_command_node, strDescription

End Property

Public Function SyntaxString(Optional IsLocal As Boolean = False) As String
    
    SyntaxString = GetSyntaxString(IsLocal)
    
End Function

Public Function ToString(Optional IsLocal As Boolean = False) As String
Attribute ToString.VB_UserMemId = 0

    Dim Trigger As String

    '// add the trigger
    Trigger = BotVars.Trigger
    If ((Me.RequiredRank < 0) And (Len(Me.RequiredFlags) = 0)) Or (IsLocal) Then
        Trigger = "/"
    End If

    ToString = Trigger & Me.Name

End Function

Public Property Get AliasString() As String
    
    AliasString = GetAliasString()
    
End Property

Public Property Get RequirementsString() As String
    
    RequirementsString = GetRequirementsString(False)
    
End Property

Public Property Get RequirementsStringShort()
    
    RequirementsStringShort = GetRequirementsString(True)
    
End Property

Public Property Get SpecialNotes() As String
    
    SpecialNotes = GetNotes(m_command_node)

End Property

Public Property Let SpecialNotes(ByVal strNotes As String)
    
    SetNotes m_command_node, strNotes

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
        Set m_params = GetParameters(m_command_node)
    End If

    Set Parameters = m_params

End Property

Public Property Get XMLDocument() As DOMDocument60

    Set XMLDocument = m_database

End Property
