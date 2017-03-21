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



'// 06/24/09 Hdx - Created
'This will return a Collection that has all of the CommandsDoc objects for all of the commands from the selected file
Public Function GetCommands(Optional ByVal ScriptName As String = vbNullString, Optional ByVal Database As DOMDocument60 = Nothing) As Collection
On Error GoTo ERROR_HANDLER

    Dim xpath        As String
    Dim CommandNodes As IXMLDOMNodeList
    Dim CommandNode  As IXMLDOMNode

    Set GetCommands = New Collection

    If Database Is Nothing Then
        Set Database = New DOMDocument60
        Database.Load GetFilePath(FILE_COMMANDS)
    End If

    ScriptName = clsCommandObj.CleanXPathVar(ScriptName)

    If ScriptName = vbNullString Then
        xpath = "/commands/command[not(@owner)]"
    ElseIf ScriptName = vbNullChar Then
        xpath = "/commands/command"
    Else
        xpath = StringFormat("/commands/command[translate(@owner, '{0}', '{1}')='{2}']", _
                UCase$(AZ), LCase$(AZ), LCase$(ScriptName))
    End If

    Set CommandNodes = Database.selectNodes(xpath)

    For Each CommandNode In CommandNodes
        GetCommands.Add CommandNode.Attributes.getNamedItem("name").nodeValue
    Next CommandNode

    Set CommandNode = Nothing
    Set CommandNodes = Nothing
    Set Database = Nothing

    Exit Function

ERROR_HANDLER:
    Set CommandNode = Nothing
    Set CommandNodes = Nothing
    Set Database = Nothing
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObj.GetCommands().")
    Exit Function

End Function

Public Function GetCommandCount(Optional ByVal strScriptOwner As String = vbNullString, Optional ByVal Database As DOMDocument60 = Nothing) As Integer
On Error GoTo ERROR_HANDLER

    Dim xpath        As String
    Dim CommandNodes As IXMLDOMNodeList

    If Database Is Nothing Then
        Set Database = New DOMDocument60
        Database.Load GetFilePath(FILE_COMMANDS)
    End If

    strScriptOwner = clsCommandObj.CleanXPathVar(strScriptOwner)

    '// create xpath expression based on strScriptOwner
    If strScriptOwner = vbNullString Then
        xpath = "/commands/command[not(@owner)]"
    ElseIf strScriptOwner = vbNullChar Then
        xpath = "/commands/command"
    Else
        xpath = StringFormat("/commands/command[translate(@owner, '{0}', '{1}')='{2}']", _
                UCase$(AZ), LCase$(AZ), LCase$(strScriptOwner))
    End If

    Set CommandNodes = Database.selectNodes(xpath)

    GetCommandCount = CommandNodes.Length

    Set CommandNodes = Nothing
    Set Database = Nothing

    Exit Function

ERROR_HANDLER:
    Set CommandNodes = Nothing
    Set Database = Nothing
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObj.GetCommandCount().")
    Exit Function

End Function


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
    ValidateXMLFromStrings = (objErr.ErrorCode = 0)
    If objErr.ErrorCode <> 0 Then
        Err.Raise 1, "ValidateXML", "Error (#" & objErr.ErrorCode & ") on Line " & objErr.Line & ": " & objErr.Reason
    End If
    
    Exit Function

ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in clsCommandDocObjStatic.ValidateXMLFromStrings().")
    ValidateXMLFromStrings = False
    
End Function
