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

Private Enum LOG
    WARNING_MSG = &H1
    ERROR_MSG = &H2
    EVENT_MSG = &H3
    DEBUG_MSG = &H4
    CHAT_MSG = &H5
    WHISPER_MSG = &H6
    SCK_OUT_MSG = &H7
    SCK_IN_MSG = &H8
End Enum


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
    If (LenB(Source) > 0) Then
        WriteLine WARNING_MSG, "warning " & Source & Space$(1) & sWarning
    Else
        WriteLine WARNING_MSG, "warning " & sWarning
    End If

End Function

' ...
Public Function WriteError(ByVal Source As String, ByVal sError As String)
    
    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (LenB(Source) > 0) Then
        WriteLine ERROR_MSG, "error " & Source & Space$(1) & sError
    Else
        WriteLine ERROR_MSG, "error " & sError
    End If

End Function

' ...
Public Function WriteEvent(ByVal Source As String, ByVal sEvent As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (LenB(Source) > 0) Then
        WriteLine EVENT_MSG, "event " & Source & Space$(1) & sEvent
    Else
        WriteLine EVENT_MSG, "event " & sEvent
    End If

End Function

' ...
Public Function WriteDebug(ByVal Source As String, ByVal sDebugMessage As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (LenB(Source) > 0) Then
        WriteLine DEBUG_MSG, "debug " & Source & Space$(1) & sDebugMessage
    Else
        WriteLine DEBUG_MSG, "debug " & sDebugMessage
    End If

End Function

' ...
Public Function WriteChat(ByVal sMessage As String)

    ' ...
    WriteLine CHAT_MSG, sMessage

End Function

' ...
Public Function WriteWhisper(ByVal sMessage As String)

    ' ...
    WriteLine WHISPER_MSG, sMessage

End Function

' ...
Public Function WriteSckOutData(ByVal src As String, ByVal sMessage As String)

    ' ...
    WriteLine SCK_OUT_MSG, sMessage

End Function

' ...
Public Function WriteSckInData(ByVal src As String, ByVal sMessage As String)

    ' ...
    WriteLine SCK_IN_MSG, sMessage

End Function

' ...
Private Function WriteLine(ByVal ltype As LOG, ByVal line As String)

    Dim f        As Integer ' ...
    Dim FilePath As String ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    Select Case (ltype)
        Case WARNING_MSG, ERROR_MSG, EVENT_MSG, DEBUG_MSG
            FilePath = (App.path & "\logs\master.txt")
        Case CHAT_MSG
            FilePath = (App.path & "\logs\" & Timestamp() & ".txt")
        Case WHISPER_MSG
            FilePath = (App.path & "\logs\" & Timestamp() & "-WHISPERS.txt")
        Case SCK_IN_MSG, SCK_OUT_MSG
            FilePath = (App.path & "\logs\" & Timestamp() & "-PACKETLOG.txt")
    End Select
    
    ' ...
    f = OpenLogFile(FilePath)
    
    ' ...
    Print #f, Timestamp & Space$(1) & line
    
    ' ...
    CloseLogFile f

End Function

' ...
Private Function OpenLogFile(ByVal path As String) As Integer
    
    Dim f As Integer ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    If (Dir$(path) = vbNullString) Then
        Open path For Output As #f
    Else
        
    End If
    
    ' ...
    OpenLogFile = f
    
End Function

Private Function CloseLogFile(ByVal f As Integer)

    ' ...
    Close #f

End Function

' ...
Private Function Timestamp() As String

    Timestamp = Format(Now, "yyyy-mm-dd HH:MM:SS")

End Function
