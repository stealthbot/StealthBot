VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLog.Cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

' ...
Private Sub Class_Initialize()

    ' ...

End Sub

' ...
Private Sub Class_Terminate()

    ' ...

End Sub

' ...
Public Function WriteWarning(ByVal Source As String, ByVal sWarning As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (Source = vbNullString) Then
        Source = "stealthbot"
    End If
    
    ' ...
    WriteLine "warning " & Source & Space$(1) & sWarning

End Function

' ...
Public Function WriteError(ByVal Source As String, ByVal Procedure As String, ByVal sError As String)

    ' ...
    If (Procedure = vbNullString) Then
        Exit Function
    End If
    
    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (Source = vbNullString) Then
        Source = "stealthbot"
    End If
    
    ' ...
    WriteLine "error " & Source & Space$(1) & Procedure & Space$(1) & sError

End Function

' ...
Public Function WriteEvent(ByVal Source As String, ByVal sEventTitle As String, ByVal sEvent As String)

    ' ...
    If (sEventTitle = vbNullString) Then
        Exit Function
    End If

    ' ...
    Source = LCase$(Source)
    sEventTitle = LCase$(sEventTitle)
    
    ' ...
    If (Source = vbNullString) Then
        Source = "stealthbot"
    End If
    
    ' ...
    WriteLine "event " & Source & Space$(1) & sEventTitle & Space$(1) & sEvent

End Function

' ...
Public Function WriteDebug(ByVal Source As String, ByVal sDebugMessage As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (Source = vbNullString) Then
        Source = "stealthbot"
    End If
    
    ' ...
    WriteLine "debug " & Source & Space$(1) & sDebugMessage

End Function

' ...
Public Function WriteChat(ByVal sMessage As String)

    ' ...

End Function

' ...
Private Function WriteLine(ByVal line As String)

    Dim f        As Integer ' ...
    Dim FilePath As String ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    FilePath = (App.Path & "\logs\log.txt")
    
    ' ...
    If (Dir$(FilePath) = vbNullString) Then
        ' ...
        Open FilePath For Output As #f
        Close #f
    End If

    ' ...
    Open FilePath For Append As #f
        Print #f, Timestamp & Space$(1) & line
    Close #f

End Function

' ...
Private Function Timestamp() As String

    ' ...
    Timestamp = Format(Now, "yyyy-mm-dd HH:MM:SS")

End Function
