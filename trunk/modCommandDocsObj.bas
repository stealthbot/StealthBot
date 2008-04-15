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

' ...
Public Function DeleteCommand(ByVal strCommand As String)

    ' ...
    
End Function

' ...
Public Function CreateCommand(ByVal strCommand As String) As clsCommandDocObj

    ' ...
    Set CreateCommand = New clsCommandDocObj
    
    ' ...
    CreateCommand.OpenDatabase
    CreateCommand.Name = strCommand

End Function
