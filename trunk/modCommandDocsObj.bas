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
    OpenCommand.OpenCommand strCommand
    
End Function

' ...
Public Function DeleteCommand(ByVal strCommand As String)

    ' ...
    
End Function

' ...
Public Function CreateCommand(ByVal strCommand As String) As clsCommandDocObj

    ' ...

End Function
