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
    Dim filePath As String ' ...
    
    ' ...
    Select Case (ltype)
        Case WARNING_MSG, ERROR_MSG, EVENT_MSG, DEBUG_MSG
            filePath = (App.path & "\logs\master.txt")
        Case CHAT_MSG
            filePath = (App.path & "\logs\" & Timestamp() & ".txt")
        Case WHISPER_MSG
            filePath = (App.path & "\logs\" & Timestamp() & "-WHISPERS.txt")
        Case SCK_IN_MSG, SCK_OUT_MSG
            filePath = (App.path & "\logs\" & Timestamp() & "-PACKETLOG.txt")
    End Select
    
    ' ...
    f = OpenLog(filePath)
    
    ' ...
    Print #f, Timestamp & Space$(1) & line
    
    ' ...
    Close #f

End Function

' ...
Private Function OpenLog(ByVal path As String) As Integer
    
    Dim f As Integer ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    If (Dir$(path) = vbNullString) Then
        Open path For Output As #f
    Else
        Dim arr() As String  ' ...
        Dim i     As Integer ' ...
        Dim str   As String  ' ...
        Dim bln   As Boolean ' ...
    
        ' ...
        Open path For Input As #f
            If (BotVars.MaxLogFileSize) Then
                If (LOF(f) >= BotVars.MaxLogFileSize) Then
                    ReDim arr(LOF(f) - 1)
                    
                    For i = 0 To LOF(f)
                        Line Input #f, str
                    
                        If (i > 0) Then
                            arr(i - 1) = str
                        End If
                    Next i
                    
                    bln = True
                End If
            End If
        Close #f
    
        ' ...
        If (bln) Then
            Open path For Binary As #f
                Put #f, , arr()
            Close #f
        End If
        
        ' ...
        Open path For Append As #f
    End If
    
    ' ...
    OpenLog = f
    
End Function

' ...
Private Function Timestamp() As String

    Timestamp = _
        "[" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "]"

End Function
