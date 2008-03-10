Attribute VB_Name = "modCommandParamObj"
' modCommandParamObj.bas
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Public Function CreateParamater(ByVal strParamter As String) As Object
    ' ...
    Set CreateParamater = New clsCommandParamsObj
End Function
