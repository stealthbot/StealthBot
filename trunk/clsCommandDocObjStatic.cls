VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandDocObjStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Public Function ValidateXMLFromFiles(ByVal strXMLPath As String, ByVal strXSDPath As String)
    
    Dim oFSO As FileSystemObject, oTS As TextStream
    Dim strXML As String, strXSD As String
    
    Set oFSO = New FileSystemObject
    
    '// read the xml file
    Set oTS = oFSO.OpenTextFile(strXMLPath, ForReading, False)
    strXML = oTS.ReadAll()
    Call oTS.Close
    
    '// read the xsd file
    Set oTS = oFSO.OpenTextFile(strXSDPath, ForReading, False)
    strXSD = oTS.ReadAll()
    Call oTS.Close
    
    ValidateXMLFromFiles = ValidateXMLFromStrings(strXML, strXSD)
    
    Set oFSO = Nothing
    Set oTS = Nothing
    
End Function


'// http://www.nonhostile.com/howto-validate-xml-xsd-in-vb6.asp
'// 08/31/2008 JSM - Created
Public Function ValidateXMLFromStrings(ByVal strXML As String, ByVal strXSD As String) As Boolean

    On Error GoTo ERROR_HANDLER

    Dim objSchemas As XMLSchemaCache60
    Dim objXML As DOMDocument60
    Dim objXSD As DOMDocument60
    Dim objErr As IXMLDOMParseError

    ' load XSD as DOM to populate in Schema Cache
    Set objXSD = New DOMDocument60
    
    objXSD.async = False
    objXSD.validateOnParse = False
    objXSD.resolveExternals = False
    
    If Not objXSD.loadXML(strXSD) Then
        Err.Raise 1, "Validate", "Load XSD failed: " & objXSD.parseError.Reason
    End If
   
    ' populate schema cache
    Set objSchemas = New XMLSchemaCache60

    ' ERROR!
    objSchemas.Add "", objXSD
   
    ' load XML file (without validation - that comes later)
    Set objXML = New DOMDocument60
    
    objXML.async = False
    objXML.validateOnParse = False
    objXML.resolveExternals = False
   
    ' load XML, without any validation
    If Not objXML.loadXML(strXML) Then
        Err.Raise 1, "Validate", "Load XML failed: " & objXML.parseError.Reason
    End If
   
    ' bind Schema Cache to DOM
    Set objXML.schemas = objSchemas
   
    ' does this XML measure up?
    Set objErr = objXML.Validate()
   
    ' any good?
    ValidateXMLFromStrings = (objErr.errorCode = 0)
    If objErr.errorCode <> 0 Then
        Err.Raise 1, "ValidateXML", "Error (#" & objErr.errorCode & ") on Line " & objErr.line & ": " & objErr.Reason
    End If

ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObjStatic.ValidateXMLFromStrings().")

End Function


'//
Public Function CommandsSanityCheck(ByRef doc As DOMDocument60, ByRef colErrorList As Collection) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
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
            Dim enabled As String
            
            '// default our values
            commandName = oCommand.Attributes.getNamedItem("name").Text
            ownerName = vbNullString
            enabled = "1"
            
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
                enabled = oCommand.Attributes.getNamedItem("enabled").Text
                If enabled <> "0" And enabled <> "1" Then
                    colErrorList.Add "If present, an enabled attribute must be equal to 0 or 1 on a command element."
                    Exit Do
                End If
            End If
            
            '// make sure only 1 command is enabled if there are similar names
            If enabled = "1" Then
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
                        If Not (oArgument.selectSingleNode("match").Attributes.getNamedItem("message")) Then
                            If Len(oArgument.selectSingleNode("match").Attributes.getNamedItem("message")) = 0 Then
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
                        Loop
                        
                    Next oRestriction
                    Exit Do
                Loop
            Next oArgument

            Exit Do
        Loop
    Next oCommand
    
    If colErrorList.Count > 0 Then
        CommandsSanityCheck = False
    Else
        CommandsSanityCheck = True
    End If



ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObjStatic.CommandsSanityCheck().")

End Function

