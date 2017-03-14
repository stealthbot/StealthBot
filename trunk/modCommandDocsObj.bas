Attribute VB_Name = "modCommandDocsObj"
' modCommandDocsObj.mod
' Copyright (C) 2007 Eric Evans


'// 07/08/2009 JSM
'// Separated this code from the scripting system. I have modified the XML
'// schema to allow us to track the script that creates each command. This
'// change allows us to disable or delete the commands along with the script
'// itself. We will be able to better organize the commands in the treeview
'// in frmCommandManager.
'//
'// Internally, these methods will work the same. They will NOT include the
'// commands from scripts. The scripts use IOpenCommand, ICreateCommand, and
'// IDeleteCommand in the SSC. This has been implemented into the scripting
'// module as OpenCommand, CreateCommand, and DeleteCommand. The methods
'// automatically pass the script name to the SSC methods. All scripts without
'// a name will have an owner="" in the command element of the XML document.
'// Unnamed scripts that create commands will be managed in an "Unknown" group
'// in frmCommandManager.


Option Explicit

Public Function OpenCommand(ByVal strCommand As String, Optional strScriptOwner As String = vbNullString) As clsCommandDocObj
    Set OpenCommand = New clsCommandDocObj
    If (InStr(1, strCommand, "'", vbBinaryCompare) > 0) Then
        Exit Function
    End If
    strCommand = Replace(strCommand, "\", "\\")
    OpenCommand.OpenDatabase
    OpenCommand.OpenCommand strCommand, strScriptOwner
End Function

'// 06/13/2009 JSM - Created
Public Function DeleteCommand(ByVal strCommand As String)
    '// open the command and call the Delete method
    Set DeleteCommand = OpenCommand(strCommand)
    DeleteCommand.Delete
End Function

'// 06/13/2009 JSM - Created
Public Function CreateCommand(ByVal strCommand As String) As clsCommandDocObj
    Dim cmd
    Set cmd = New clsCommandDocObj
    '// create the command
    Call cmd.CreateCommand(strCommand)
    '// now lets return it
    Set CreateCommand = OpenCommand(strCommand)
End Function
