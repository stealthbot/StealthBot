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

Private Enum LOG_TYPES
    WARNING_MSG = &H1
    ERROR_MSG = &H2
    EVENT_MSG = &H3
    DEBUG_MSG = &H4
    CHAT_MSG = &H5
    WHISPER_MSG = &H6
    SCK_MSG = &H7
    COMMAND_MSG = &H8
End Enum

Private m_logsCreated As Collection ' ...

Private Sub Class_Initialize()

    ' ...
    Set m_logsCreated = New Collection

End Sub

Private Sub Class_Teriminate()

    ' ...
    If (BotVars.Logging = 2) Then
        RemoveLogsCreated
    End If

    ' ...
    Set m_logsCreated = Nothing

End Sub

' ...
Public Function WriteWarning(ByVal sWarning As String, Optional ByVal Source As String)

    ' ...
    Source = LCase$(Source)

    ' ...
    If (LenB(Source) > 0) Then
        WriteLine WARNING_MSG, "warning " & Source & " " & sWarning
    Else
        WriteLine WARNING_MSG, "warning " & sWarning
    End If

End Function

' ...
Public Function WriteError(ByVal sError As String, Optional ByVal Source As String)
    
    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (LenB(Source) > 0) Then
        WriteLine ERROR_MSG, "error " & Source & " " & sError
    Else
        WriteLine ERROR_MSG, "error " & sError
    End If
        
End Function

' ...
Public Function WriteEvent(ByVal sTitle As String, ByVal sEvent As String, Optional ByVal Source As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (LenB(Source) > 0) Then
        WriteLine EVENT_MSG, "event " & sTitle & " " & Source & " " & sEvent
    Else
        WriteLine EVENT_MSG, "event " & sTitle & " " & sEvent
    End If

End Function

' ...
Public Function WriteCommand(ByVal sCommand As String)

    ' ...
    If (BotVars.LogCommands) Then
        WriteEvent "command", sCommand
    End If

End Function

' ...
Public Function WriteDatabase(ByVal sEvent As String)

    ' ...
    If (BotVars.LogDBActions) Then
        WriteEvent "database", sEvent
    End If

End Function

' ...
Public Function WriteDebug(ByVal sDebugMessage As String, Optional ByVal Source As String)

    ' ...
    Source = LCase$(Source)
    
    ' ...
    If (isDebug) Then
        If (LenB(Source) > 0) Then
            WriteLine DEBUG_MSG, "debug " & Source & " " & sDebugMessage
        Else
            WriteLine DEBUG_MSG, "debug " & sDebugMessage
        End If
    End If

End Function

' ...
Public Function WriteChat(ByVal sMessage As String)

    ' ...
    If (BotVars.Logging >= 1) Then
        WriteLine CHAT_MSG, sMessage
    End If

End Function

' ...
Public Function WriteWhisper(ByVal sMessage As String)

    ' ...
    If (BotVars.Logging >= 1) Then
        WriteLine WHISPER_MSG, sMessage
    End If

End Function

' ...
Public Function WriteSckData(ByVal sMessage As String)

    ' ...
    If (LogPacketTraffic) Then
        WriteLine SCK_MSG, sMessage
    End If

End Function

' ...
Private Function WriteLine(ByVal ltype As LOG_TYPES, ByVal line As String)

    On Error GoTo ERROR_HANDLER

    Dim f         As Integer ' ...
    Dim filePath  As String  ' ...
    
    ' ...
    Select Case (ltype)
        Case WARNING_MSG, ERROR_MSG, EVENT_MSG, DEBUG_MSG
            filePath = _
                (App.Path & "\logs\master.txt")
        
        Case COMMAND_MSG
            filePath = _
                (App.Path & "\logs\commands.txt")
                
        Case CHAT_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & ".txt")
                
        Case WHISPER_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & "-WHISPERS.txt")
                
        Case SCK_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & "-PACKETLOG.txt")
                
    End Select
    
    ' ...
    f = OpenLog(ltype, filePath)
    
    ' ...
    If (f > 0) Then
        Select Case (ltype)
            Case CHAT_MSG, WHISPER_MSG, SCK_MSG
                Print #f, "[" & Timestamp() & "] " & line
            Case Else
                Print #f, "[" & Datestamp() & " " & Timestamp() & "] " & line
        End Select
        
        ' ...
        Close #f
    End If
    
    ' ...
    Exit Function
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in WriteLine()."

End Function

' ...
Private Function OpenLog(ByVal ltype As LOG_TYPES, ByVal Path As String) As Integer

    On Error GoTo ERROR_HANDLER
    
    ' ...
    Static failed_attempts As Integer
    
    Dim f As Integer ' ...
    Dim I As Double  ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    If (Dir$(Path) = vbNullString) Then
        Dim dir_path As String  ' ...

        ' ...
        dir_path = Mid$(Path, 1, InStrRev(Path, "\"))
    
        ' ...
        If (Dir$(dir_path) = vbNullString) Then
            MkDir dir_path
        End If
    
        ' ...
        Open Path For Output As #f
        
        ' ...
        If ((ltype = CHAT_MSG) Or (ltype = WHISPER_MSG)) Then
            For I = 1 To m_logsCreated.Count
                If (StrComp(m_logsCreated(I), Path, vbTextCompare) = 0) Then
                    Exit For
                End If
            Next I
            
            ' ...
            If (I >= m_logsCreated.Count + 1) Then
                m_logsCreated.Add Path
            End If
        End If
    Else
        Dim arr()     As String  ' ...
        Dim str       As String  ' ...
        Dim bln       As Boolean ' ...
        Dim lineCount As Long    ' ...
    
        ' ...
        If (BotVars.MaxLogFileSize) Then
            Open Path For Input As #f
                If (LOF(f) >= BotVars.MaxLogFileSize) Then
                    Do Until (EOF(f))
                        Line Input #f, str
                        
                        ReDim Preserve arr(lineCount)
                    
                        arr(lineCount) = str
                        
                        lineCount = (lineCount + 1)
                    Loop
                    
                    bln = True
                End If
            Close #f
        End If
    
        ' ...
        If (bln) Then
            Dim offset As Integer ' ...
            Dim bytes  As Integer ' ...
            
            ' ...
            For I = UBound(arr) To 0 Step -1
                bytes = (bytes + Len(arr(I)))
                
                If (bytes >= BotVars.MaxLogFileSize) Then
                    offset = I
                    
                    Exit For
                End If
            Next I
        
            ' ...
            Open Path For Output As #f
                For I = offset + 1 To UBound(arr)
                    Print #f, arr(I)
                Next I
            Close #f
        End If
        
        Open Path For Append As #f
    End If
    
    ' ...
    OpenLog = f
    
    ' ...
    failed_attempts = 0
    
    ' ...
    Exit Function
    
ERROR_HANDLER:
    
    ' permission denied?  is someone else trying to save the file?
    If (Err.Number = 70) Then
        failed_attempts = (failed_attempts + 1)
        
        If (failed_attempts >= 3) Then
            Exit Function
        End If
    
        OpenLog ltype, Path
    End If
    
End Function

' ...
Public Sub RemoveLogsCreated()

    On Error Resume Next

    Dim I As Integer ' ...
    
    ' ...
    For I = 1 To m_logsCreated.Count
        Kill m_logsCreated(I)
    Next I

End Sub

Private Function Datestamp() As String

    ' ...
    Datestamp = Format(Now, "YYYY-MM-DD")

End Function

Private Function Timestamp() As String

    ' ...
    Select Case (BotVars.TSSetting)
        Case 0
            Timestamp = Format(Now, "HH:MM:SS AM/PM")
            
        Case 1
            Timestamp = Format(Now, "HH:MM:SS")
        
        Case 2
            Timestamp = _
                Format(Now, "HH:MM:SS") & "." & _
                    Right$("000" & GetCurrentMS, 3)
        
        Case Else
            Timestamp = Format(Now, "HH:MM:SS AM/PM")
    End Select

End Function
