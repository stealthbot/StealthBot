Attribute VB_Name = "modStaticClasses"
'// declare a single static class instance
Dim static_clsCommandObj As clsCommandObjStatic



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

