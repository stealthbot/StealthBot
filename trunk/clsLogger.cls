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
Private m_logPath     As String     ' the path to the log file

Private Sub Class_Initialize()

    ' ...
    Set m_logsCreated = New Collection
    
    m_logPath = App.Path & "\logs\"

End Sub

Private Sub Class_Teriminate()

    ' ...
    If (BotVars.Logging = 1) Then
        RemoveLogsCreated
    End If

    ' ...
    Set m_logsCreated = Nothing

End Sub

Public Property Get LogPath() As String
    LogPath = m_logPath
End Property

Public Property Let LogPath(ByVal value As String)
    m_logPath = value
    
    If (Not (Right$(m_logPath, 1) = "\")) Then
        m_logPath = m_logPath & "\"
    End If
End Property

' ...
Public Function WriteWarning(ByVal sWarning As String, Optional ByVal source As String, Optional TimeDate As Date)

    ' ...
    source = LCase$(source)

    ' ...
    If (LenB(source) > 0) Then
        WriteLine WARNING_MSG, "warning " & source & " " & sWarning, TimeDate
    Else
        WriteLine WARNING_MSG, "warning " & sWarning, TimeDate
    End If

End Function

' ...
Public Function WriteError(ByVal sError As String, Optional ByVal source As String, Optional TimeDate As Date)
    
    ' ...
    source = LCase$(source)
    
    ' ...
    If (LenB(source) > 0) Then
        WriteLine ERROR_MSG, "error " & source & " " & sError, TimeDate
    Else
        WriteLine ERROR_MSG, "error " & sError, TimeDate
    End If
        
End Function

' ...
Public Function WriteEvent(ByVal sTitle As String, ByVal sEvent As String, Optional ByVal source As String, Optional TimeDate As Date)

    ' ...
    source = LCase$(source)
    
    ' ...
    If (LenB(source) > 0) Then
        WriteLine EVENT_MSG, "event " & sTitle & " " & source & " " & sEvent, TimeDate
    Else
        WriteLine EVENT_MSG, "event " & sTitle & " " & sEvent, TimeDate
    End If

End Function

' ...
Public Function WriteCommand(ByVal sCommand As String, Optional TimeDate As Date)

    ' ...
    If (BotVars.LogCommands) Then
        WriteEvent "command", sCommand, TimeDate
    End If

End Function

' ...
Public Function WriteDatabase(ByVal sEvent As String, Optional TimeDate As Date)

    ' ...
    If (BotVars.LogDBActions) Then
        WriteEvent "database", sEvent, TimeDate
    End If

End Function

' ...
Public Function WriteDebug(ByVal sDebugMessage As String, Optional ByVal source As String, Optional TimeDate As Date)

    ' ...
    source = LCase$(source)
    
    ' ...
    If (isDebug) Then
        If (LenB(source) > 0) Then
            WriteLine DEBUG_MSG, "debug " & source & " " & sDebugMessage, TimeDate
        Else
            WriteLine DEBUG_MSG, "debug " & sDebugMessage, TimeDate
        End If
    End If

End Function

' ...
Public Function WriteChat(ByVal sMessage As String, Optional TimeDate As Date)

    ' ...
    If (BotVars.Logging >= 1) Then
        WriteLine CHAT_MSG, sMessage, TimeDate
    End If

End Function

' ...
Public Function WriteWhisper(ByVal sMessage As String, Optional TimeDate As Date)

    ' ...
    If (BotVars.Logging >= 1) Then
        WriteLine WHISPER_MSG, sMessage, TimeDate
    End If

End Function

' ...
Public Function WriteSckData(ByVal sMessage As String, Optional TimeDate As Date)

    ' ...
    If (LogPacketTraffic) Then
        WriteLine SCK_MSG, sMessage, TimeDate
    End If

End Function

' ...
Private Function WriteLine(ByVal ltype As LOG_TYPES, ByVal line As String, Optional TimeDate As Date)

    On Error GoTo ERROR_HANDLER

    Dim f         As Integer ' ...
    Dim filePath  As String  ' ...
    
    ' ...
    Select Case (ltype)
        Case WARNING_MSG, ERROR_MSG, EVENT_MSG, DEBUG_MSG
            filePath = _
                (m_logPath & "master.txt")
        
        Case COMMAND_MSG
            filePath = _
                (m_logPath & "commands.txt")
                
        Case CHAT_MSG
            filePath = _
                (m_logPath & Datestamp() & ".txt")
                
        Case WHISPER_MSG
            filePath = _
                (m_logPath & Datestamp() & "-WHISPERS.txt")
                
        Case SCK_MSG
            filePath = _
                (m_logPath & Datestamp() & "-PACKETLOG.txt")
                
    End Select
    
    ' ...
    f = OpenLog(ltype, filePath)
    
    ' ...
    If (f > 0) Then
        Select Case (ltype)
            Case CHAT_MSG, WHISPER_MSG, SCK_MSG
                Print #f, "[" & Timestamp(TimeDate) & "] " & line
            Case Else
                Print #f, "[" & Datestamp(TimeDate) & " " & Timestamp(TimeDate) & "] " & _
                    line
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

Private Function Datestamp(Optional TimeDate As Date) As String

    ' ...
    If (DateDiff("s", TimeDate, "00:00:00 12/30/1899") = 0) Then
        TimeDate = Now
    End If

    ' ...
    Datestamp = Format(TimeDate, "YYYY-MM-DD")

End Function

Private Function Timestamp(Optional TimeDate As Date) As String

    ' ...
    If (DateDiff("s", TimeDate, "00:00:00 12/30/1899") = 0) Then
        TimeDate = Now
    End If

    ' ...
    Select Case (BotVars.TSSetting)
        Case 0
            Timestamp = Format(TimeDate, "HH:MM:SS AM/PM")
            
        Case 1
            Timestamp = Format(TimeDate, "HH:MM:SS")
        
        Case 2
            Timestamp = _
                Format(TimeDate, "HH:MM:SS") & "." & _
                    Right$("000" & GetCurrentMS, 3)
        
        Case Else
            Timestamp = Format(TimeDate, "HH:MM:SS AM/PM")
    End Select

End Function
