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
Public Function GetCommands(Optional ByVal scriptName As String = vbNullString) As Collection
    
    Dim AZ As String
    Dim xpath As String
    Dim doc As New DOMDocument60
    Dim commandNodes As IXMLDOMNodeList
    Dim commandNode As IXMLDOMNode
    
    AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    Set GetCommands = New Collection
  
    doc.Load GetFilePath("Commands.xml")
    
    scriptName = Replace$(scriptName, "'", "&apos;")
    
    If scriptName = vbNullString Then
         xpath = StringFormat("./commands/command[not(@owner) and (not(@enabled) or @enable = 1)]", _
            UCase$(AZ), LCase$(AZ), LCase$(scriptName))
    ElseIf scriptName = Chr$(0) Then
        xpath = StringFormat("./commands/command[translate(@owner, '{0}', '{1}')='{2}' and (not(@enabled) or @enable = 1)]", _
            UCase$(AZ), LCase$(AZ), LCase$(scriptName))
    Else
        xpath = StringFormat("./commands/command[translate(@owner, '{0}', '{1}')='{2}']", _
            UCase$(AZ), LCase$(AZ), LCase$(scriptName))
    End If
        
    Set commandNodes = doc.selectNodes(xpath)
    
    For Each commandNode In commandNodes
      GetCommands.Add commandNode.Attributes.getNamedItem("name").nodeValue
    Next commandNode
    
    Set commandNode = Nothing
    Set commandNodes = Nothing
    Set doc = Nothing
    
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
        Err.Raise 1, "ValidateXML", "Error (#" & objErr.ErrorCode & ") on Line " & objErr.line & ": " & objErr.Reason
    End If
    
    Exit Function

ERROR_HANDLER:
    
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsCommandDocObjStatic.ValidateXMLFromStrings().")
    ValidateXMLFromStrings = False
    
End Function
