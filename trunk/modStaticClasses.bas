Attribute VB_Name = "modStaticClasses"
Option Explicit

'// declare a single static class instance
Dim static_clsCommandObj As clsCommandObjStatic
Dim static_clsCommandDocObj As clsCommandDocObjStatic


'// create functions that match our class object that return the
'// static instance
Function clsCommandObj() As clsCommandObjStatic
    
    '// create static instance if empty
    If (static_clsCommandObj Is Nothing) Then
        Set static_clsCommandObj = New clsCommandObjStatic
    End If
    
    '// return static instance class
    Set clsCommandObj = static_clsCommandObj
    
End Function

Function clsCommandDocObj() As clsCommandDocObjStatic
    
    '// create static instance if empty
    If (static_clsCommandDocObj Is Nothing) Then
        Set static_clsCommandDocObj = New clsCommandDocObjStatic
    End If
    
    '// return static instance class
    Set clsCommandDocObj = static_clsCommandDocObj
    
End Function
