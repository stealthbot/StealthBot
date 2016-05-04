Option Strict Off
Option Explicit On
Module modCommandDocsObj
	' modCommandDocsObj.mod
	' Copyright (C) 2007 Eric Evans
	
	
	'// 07/08/2009 JSM
	'// Separated this code from the scripting system. I have modified the XML
	'// schema to allow us to track the script that creates each command. This
	'// change allows us to disable or delete the commands along with the script
	'// itself. We will be able to better organize the commands in the treeview
	'// in frmCommands.
	'//
	'// Internally, these methods will work the same. They will NOT include the
	'// commands from scripts. The scripts use IOpenCommand, ICreateCommand, and
	'// IDeleteCommand in the SSC. This has been implemented into the scripting
	'// module as OpenCommand, CreateCommand, and DeleteCommand. The methods
	'// automatically pass the script name to the SSC methods. All scripts without
	'// a name will have an owner="" in the command element of the XML document.
	'// Unnamed scripts that create commands will be managed in an "Unknown" group
	'// in frmCommands.
	
	
	
	Public Function OpenCommand(ByVal strCommand As String, Optional ByRef strScriptOwner As String = vbNullString) As clsCommandDocObj
		OpenCommand = New clsCommandDocObj
		If (InStr(1, strCommand, "'", CompareMethod.Binary) > 0) Then
			Exit Function
		End If
		strCommand = Replace(strCommand, "\", "\\")
		OpenCommand.OpenDatabase()
		OpenCommand.OpenCommand(strCommand, strScriptOwner)
	End Function
	
	'// 06/13/2009 JSM - Created
	Public Function DeleteCommand(ByVal strCommand As String) As Object
		'// open the command and call the Delete method
		DeleteCommand = OpenCommand(strCommand)
		'UPGRADE_WARNING: Couldn't resolve default property of object DeleteCommand.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DeleteCommand.Delete()
	End Function
	
	'// 06/13/2009 JSM - Created
	Public Function CreateCommand(ByVal strCommand As String) As clsCommandDocObj
		Dim cmd As Object
		cmd = New clsCommandDocObj
		'// create the command
		'UPGRADE_WARNING: Couldn't resolve default property of object cmd.CreateCommand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call cmd.CreateCommand(strCommand)
		'// now lets return it
		CreateCommand = OpenCommand(strCommand)
	End Function
End Module