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

Public Function HasAccess(ByRef cmd As clsCommandObj, ByVal strUsername As String) As Boolean

    '// TODO: Move code from modCommandCode.HasAccess to here
    HasAccess = modCommandCode.HasAccess(strUsername, cmd.Name, cmd.Args)
    

End Function


Public Function HasCorrectSyntax(ByRef cmd As clsCommandObj) As Boolean
    
    '// TODO: Move code from modCommandCode.IsCorrectSyntax to here
    HasCorrectSyntax = modCommandCode.IsCorrectSyntax(cmd.Name, cmd.Args)
    
End Function


'// this function takes a string and will return a clsCommandObj object. If strOwnerName is missing or vbNullstring
'// then this function will check for an internal command. All triggers should be removed from strText
'// and Len(strText) > 0. This function does NOT consider multiple commands contained inside
'// strText. All MULTICOMMAND PARSING SHOULD TAKE PLACE PRIOR TO CALLING THIS METHOD. Since this
'// function does not have triggers
'//
'// EXAMPLE:
'// Set cmd = CreateCommandInstance("add FiftyToo 50", "someUser")
'// If cmd.IsValidCommand Then
'//     frmChat.AddChat vbGreen, cmd.Name
'// End If
Private Function CreateCommandInstance(ByRef commandString As String, ByVal strUsername As String, Optional ByVal strScriptOwner As String = vbNullString) As clsCommandObj
    
    On Error GoTo ERROR_HANDLER
    
    Dim cmd As clsCommandObj
    
    Dim commandName As String
    Dim commandArgs As String
    Dim tmp() As String
    
    '// separate the command's name and args from the command string
    tmp = Split(commandString, " ", 2)
    commandName = tmp(0)
    If UBound(tmp) = 1 Then
        commandArgs = tmp(1)
    End If
    
    '// make sure that this command exists for strScriptOwner
    Me.OpenCommand commandName, strScriptOwner
    If (m_command_node Is Nothing) Then
        '// this is not a command, we are done here
        Set CreateCommandInstance = Nothing
        Exit Function
    End If
    
    '// ok this is actually a command, lets create the object
    Set cmd = New clsCommandObj
    With cmd
        .Name = commandName
        .Args = commandArgs
        '.docs = Me
        .Arguments = SplitArguments(commandArgs)
        .Username = strUsername
    End With
    
    '// all good in the hood :)
    Set CreateCommandInstance = cmd
    Exit Function

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.CreateCommandObject().")

    Exit Function

End Function

'// this function takes the raw args string (everything after the command) and returns a
'// collection of strings. Each string is a argument that is parsed using the new
'// argument snytax.
'// EXAMPLE
'// mycommand "this is a \"single\" arg" and here are 5 more
Private Function SplitArguments(ByVal strArgString As String) As Collection

    On Error GoTo ERROR_HANDLER

    Dim I           As Integer   '// counter
    Dim L, r        As String    '// temp vars to store the left and right characters
    Dim tmp()       As String    '// array of words
    Dim Word        As String    '// stores the word
    Dim multiword   As String    '// stores the text of a multi-word argument
    Dim insideArg   As Boolean   '// used to check if a word begins a multi-word argument
    
    Set SplitArguments = New Collection
    
    '// take out any extra spaces
    strArgString = Trim$(strArgString)
    
    If Len(strArgString) = 0 Then
        '// no arguments
        Exit Function
    End If
    
    '// if there is no space then we can just strip the quotes (if present), add
    '// it to the collection, and then return
    If InStr(1, strArgString, " ") < 1 Then
    
        Word = StripQuotes(Replace(strArgString, "\""", Chr(0)))
        If InStr(1, Word, """") > 0 Then
            '// this is bad... words cannot contain unescaped "
            'Err.Raise -1, 0&, ""Words cannot contain unescaped """. Args =: " & strArgString
            Set SplitArguments = Nothing
            Set SplitArguments = New Collection
            Exit Function
        End If
        
        Word = Replace(Word, Chr(0), """")
    
        SplitArguments.Add Word
        Exit Function
    End If
    
    '// default some variables
    insideArg = False
    multiword = ""
    
    '// loop through each element and group the arguments
    tmp = Split(strArgString)
    For I = LBound(tmp) To UBound(tmp)
        Word = tmp(I)
        '// allow for escaping quotes
        Word = Replace(Word, "\""", Chr(0))
        
        '// if the length is 2 or more then then l and r should be the first and last character
        If Len(Word) > 1 Then
            L = Left$(Word, 1)
            r = Right$(Word, 1)
        '// if the length is 1, then we need to set either l or r to "" depending on insideArg
        ElseIf Len(Word) > 0 Then
            If insideArg = False Then
                L = Left$(Word, 1)
                r = ""
            Else
                L = ""
                r = Right$(Word, 1)
            End If
        Else
            L = ""
            r = ""
        End If
        
'// check if this word BEGINS with a " and we ARE NOT inside an arg
        If (L = """") And (r = """") Then
            '// this should be a single argument, if we are inside a word we have a problem
            If insideArg = True Then
                '// this is bad... words cannot contain unescaped "
                'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
                Set SplitArguments = Nothing
                Set SplitArguments = New Collection
                Exit Function
            End If
            
            '// ok this is a single word arg, but we still need to fail if it contains a "
            If InStr(1, Word, """") > 0 Then
                '// this is bad... words cannot contain unescaped "
                'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
                Set SplitArguments = Nothing
                Set SplitArguments = New Collection
                Exit Function
            End If
            
            Word = Replace(Word, Chr(0), """")
        
        ElseIf (L = """") And (insideArg = False) Then
        
            '// we are, lets start the multiword and set our bit
            multiword = Word & " "
            insideArg = True
            
'// check if this word ENDS with a " and we ARE inside an arg
        ElseIf (r = """") And (insideArg = True) Then
        
            '// we are, lets end the multiword, add it to the collect, and reset our vars
            multiword = multiword & Word
            SplitArguments.Add Replace(StripQuotes(multiword), Chr(0), """")
            insideArg = False
            multiword = ""
            
'// check if we are inside a word, if so then we append it to multi word and be done with it
        ElseIf (insideArg = True) Then
        
            multiword = multiword & Word & " "
            
'// we are not inside a word, then this must be a separate argument so we need to add it
        ElseIf (r <> """") And (L <> """") And (insideArg = False) Then
        
        
            '// make sure this word does not have any " inside it.
            If InStr(1, Word, """") > 0 Then
                '// this is bad... words cannot contain unescaped "
                'Err.Raise -1, 0&, "Words cannot contain unescaped "". Args =: " & strArgString
                Set SplitArguments = Nothing
                Set SplitArguments = New Collection
                Exit Function
            End If
            
            '// if there is no text and not inside a word, then we should ignore it
            If Len(Word) > 0 Then
                SplitArguments.Add Replace(StripQuotes(Word), Chr(0), """")
                insideArg = False
                multiword = ""
            End If
'// this should never happen with valid argument syntax
        Else
            'Err.Raise -1, 0&, "Cannot determine type of word. Args =: " & strArgString
            Set SplitArguments = Nothing
            Set SplitArguments = New Collection
            Exit Function
            
        End If
    
    Next I
    
    '// final test
    If insideArg = True Then
        '// this is bad... we ended inside an argument
        'Err.Raise -1, 0&, "Ended with an open arg string. Args =: " & strArgString
        Set SplitArguments = Nothing
        Set SplitArguments = New Collection
        Exit Function
    End If
    
    '// all good :)
    Exit Function

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.SplitArguments().")

    Exit Function

End Function

'// this function will remove the first and last double quote from a string, but only
'// if both are present
Private Function StripQuotes(ByVal strText As String) As String


    Dim retVal As String
    Dim leftStripped As Boolean
    Dim rightStripped As Boolean

    leftStripped = False
    rightStripped = False

    retVal = strText

    If Left$(retVal, 1) = """" Then
        retVal = Mid$(retVal, 2)
        leftStripped = True
    End If
    
    If Right$(retVal, 1) = """" Then
        retVal = Mid(retVal, 1, Len(retVal) - 1)
        rightStripped = True
    End If
    
    '// if these values are the same, then we can return retval, otherwise we should return
    '// whatever was passed into the function
    If leftStripped = rightStripped Then
        StripQuotes = retVal
    Else
        StripQuotes = strText
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




'// This function returns a clsCommandObj that is populated with instance of the command
'// object.
Public Function IsCommand(ByVal strText As String, ByVal strUsername As String, Optional ByVal strScriptOwner As String = vbNullString) As Collection
    
    On Error GoTo ERROR_HANDLER
    
    Const CMD_DELIMITER As String = "; "

    Dim Message        As String  '// the raw message
    Dim messageLen     As Integer '// the length of the raw message
    Dim cropLen        As Integer '// the length of the trigger
    Dim hasTrigger     As Boolean '// if true, a trigger has been found
    Dim botUsername    As String  '// stores the bot's username returned from modOtherCode.GetCurrentUsername
    Dim botRawUsername As String  '// stores the bot's username retrieved from modGlobals.CurrentUsername
    
    
    '// used for creating command instaces
    Dim command        As clsCommandObj
    Dim commandIndex   As Integer
    Dim commandString  As String
    Dim commandStrings As Collection
    Dim PublicOutput   As Boolean '// if true, then the output should be sent to the channel
    Dim IsLocal        As Boolean '// if true, then this command was enter via the chat window
    
    
    Set IsCommand = New Collection
    
    '// make sure actual text was passed for the message, otherwise we return an empty collection
    If (strText = vbNullString) Then
        '// Not a command
        Exit Function
    End If
    
    '// get the bot's username into variables. botUsername will be the bots username without any
    '// domain or product info (like * for D2 and @USEast/@Azeroth etc). botRawUsername will contain
    '// this information. The commands should trigger from both versions.
    botUsername = modGlobals.CurrentUsername
    botRawUsername = modOtherCode.GetCurrentUsername()
    
    '// if strUsername matches the bot username, it must have been entered from within the bot
    IsLocal = (0 = StrComp(strUsername, botUsername, vbTextCompare))
                
    hasTrigger = False
    Message = strText
    messageLen = Len(Message)
    
    '// If this command was entered via the bot we need to check for slashes.
    '//
    '// 0 slashes - No further processing
    '// 1 slashes - publicOutput = False
    '// 2 slashes - publicOutput = True
    '// 3 slashes - No further processing
    If (IsLocal = True) Then
    
        '// commands entered through the bot should use a /
        If (Left$(Message, 1) <> "/") Then
            '// Not a command
            Exit Function
        End If
    
        '// make sure we do no further processing if the message is nothing but 1 slash
        If (Left$(Message, 1) = "/") And messageLen = 1 Then
            '// Not a command
            Exit Function
        End If
    
        '// make sure we do no further processing if the message is nothing but 2 slashes
        If (Left$(Message, 2) = "//") And messageLen = 2 Then
            '// Not a command
            Exit Function
        End If
        
        '// make sure we do no further processing if the message begins with ///
        If (Left$(Message, 3) = "///") Then
            '// Not a command
            Exit Function
        End If
        
        '// at this point, if message begins with // than public output should be true
        If (Left$(Message, 2) = "//") Then
            PublicOutput = True
            cropLen = 3
            hasTrigger = True
        Else
            PublicOutput = False
            cropLen = 2
            hasTrigger = True
        End If
                
    End If '// (IsLocal = True)
    
    
    '// if this command was not entered via the bot, then we need to check for the bot trigger
    '// as well as for the 2 global triggers, ops and all.
    If (IsLocal = False) Then
    
        '// check for bot trigger
        '// EXAMPLE COMMAND STRING
        '// .add SammyHagar 200
        If (Left$(Message, Len(BotVars.TriggerLong)) = BotVars.TriggerLong) Then
            PublicOutput = True
            cropLen = Len(BotVars.TriggerLong) + 1
            hasTrigger = True
        End If
        
        '// check for "all: " or "all, ". These special triggers work for all bots
        '// EXAMPLE COMMAND STRING
        '// all: add SammyHagar 200
        If (hasTrigger = False) And (messageLen > 5) Then
            If (Not StrComp(Left$(Message, 3), "all", vbTextCompare)) And (Mid$(Message, 4, 2) = ": " Or Mid$(Message, 4, 2) = ", ") Then
                PublicOutput = True
                cropLen = 6
                hasTrigger = True
            End If
        End If
        
        '// check for "ops: " or "ops, ". These special triggers work for all bots that are operators
        '// EXAMPLE COMMAND STRING
        '// ops: add SammyHagar 200
        If (hasTrigger = False) And (messageLen > 5) Then
            If (Not StrComp(Left$(Message, 3), "ops", vbTextCompare)) And (Mid$(Message, 4, 2) = ": " Or Mid$(Message, 4, 2) = ", ") Then
                If (g_Channel.Self.IsOperator) Then
                    PublicOutput = True
                    cropLen = 6
                    hasTrigger = True
                End If
            End If
        End If
        
        '// check for bots name as a trigger.
        '// EXAMPLE COMMAND STRING
        '// FiftyToo: add SammyHagar 200
        If (hasTrigger = False) And (messageLen > Len(botUsername) + 2) Then
            If StrComp(Left$(Message, Len(botUsername)), botUsername, vbTextCompare) = 0 And _
                (Mid$(Message, Len(botUsername) + 1, 2) = ": " Or Mid$(Message, Len(botUsername) + 1, 2) = ", ") Then
                
                PublicOutput = True
                cropLen = Len(botUsername) + 3
                hasTrigger = True
            
            End If
        End If
                
        '// check for bots name as a trigger, with respect to product and realm
        '// EXAMPLE COMMAND STRING
        '// *FiftyToo: add SammyHagar 200
        If (hasTrigger = False) And (messageLen > Len(botRawUsername) + 2) Then
            If StrComp(Left$(Message, Len(botRawUsername)), botRawUsername, vbTextCompare) = 0 And _
                (Mid$(Message, Len(botRawUsername) + 1, 2) = ": " Or Mid$(Message, Len(botRawUsername) + 1, 2) = ", ") Then
                
                PublicOutput = True
                cropLen = Len(botRawUsername) + 3
                hasTrigger = True
            
            End If
        End If
        
        '// check for a pattern that matches the bot username
        '// EXAMPLE COMMAND STRING (matches fiftytoo followed by any 3 numbers)
        '// FiftyToo###: add SammyHagar 200
        If (hasTrigger = False) And InStr(1, Message, ": ") > 0 Then
            If (UsernameRegex(botUsername, Left(Message, InStr(1, Message, ": ") - 1)) Or _
                UsernameRegex(botRawUsername, Left$(Message, InStr(1, Message, ": ") - 1))) Then
                
                PublicOutput = True
                cropLen = InStr(1, Message, ": ") + 2
                hasTrigger = True
            End If
        End If
        If (hasTrigger = False) And InStr(1, Message, ", ") > 0 Then
            If (UsernameRegex(botUsername, Left$(Message, InStr(1, Message, ", ") - 1)) Or _
                UsernameRegex(botRawUsername, Left$(Message, InStr(1, Message, ", ") - 1))) Then
                
                PublicOutput = True
                cropLen = InStr(1, Message, ", ") + 2
                hasTrigger = True
            End If
        End If
        
        '// if we have not found a trigger, lets get out of here
        If (hasTrigger = False) Then
            '// Not a command
            Exit Function
        End If
        
    End If '// (IsLocal = False)

    
    '// get a collection of commands based on the split logic
    Set commandStrings = SplitCompleteCommandString(Mid$(Message, cropLen))
    
    '// if this command string has multiple commands, lets parse them out and process
    '// them individually.
    For commandIndex = 1 To commandStrings.Count
        commandString = commandStrings(commandIndex)
        '// lets try to parse this command and add it to the collection
        Set command = CreateCommandInstance(commandString, strUsername, strScriptOwner)
        If Not (command Is Nothing) Then
            '// we only want to add the command if it is enabled
            If command.docs.IsEnabled = True Then
                command.PublicOutput = PublicOutput
                command.IsLocal = IsLocal
                IsCommand.Add command
            End If
        End If
    Next
    
    '// all done here
    Exit Function
    
ERROR_HANDLER:
    If (Err.Number = 93) Then
        Err.Clear
        Exit Function
    End If
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObj.IsCommand().")
End Function

'// This function will split a message into separate commands. This message should NOT
'// have any triggers. This will return a collection of command strings that can be used
'// to create an instance of a command.
'//
'// TODO:
'// Fix logic to allow "; " inside a quoted argument.
Private Function SplitCompleteCommandString(ByVal completeCommandString As String) As Collection

    Dim I As Integer
    Dim commandString As String
    Dim commandStrings() As String
    
    Set SplitCompleteCommandString = New Collection
    
    '// use "; " as a delimiter for commands. Allows for /; to escape a command split
    completeCommandString = Replace(completeCommandString, "\;", Chr(0))
    commandStrings = Split(completeCommandString, "; ")
    For I = LBound(commandStrings) To UBound(commandStrings)
        '// make sure these some actual text for this command, otherwise skip it
        If Len(commandStrings(I)) > 0 Then
            SplitCompleteCommandString.Add Replace(commandStrings(I), Chr(0), ";")
        End If
    Next I

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
    
    If strScriptOwner = vbNullString Then
        xpath = "./command[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(strCommand) & "' and " & _
                          "not(@owner)]"
    ElseIf strScriptOwner = Chr$(0) Then
        xpath = "./command[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(strCommand) & "']"
    Else
        xpath = "./command[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" & LCase(strCommand) & "' and " & _
                          "translate(@owner, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = '" & LCase(strScriptOwner) & "']"
    End If
    
    Set m_command_node = m_command_node.selectSingleNode(xpath)
    
    If (m_command_node Is Nothing) Then
        Exit Function
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
