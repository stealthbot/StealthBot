Attribute VB_Name = "modCommandDocsObj"
' modCommandDocsObj.mod
' Copyright (C) 2007 Eric Evans
' ...

Option Explicit

' ...
Public Function OpenCommand(ByVal strCommand As String) As clsCommandDocObj
    
    ' ...
    Set OpenCommand = New clsCommandDocObj
    
    ' ...
    If (InStr(1, strCommand, "'", vbBinaryCompare) > 0) Then
        Exit Function
    End If
    
    strCommand = Replace(strCommand, "\", "\\")
    
    ' ...
    OpenCommand.OpenDatabase
    OpenCommand.OpenCommand strCommand
    
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
