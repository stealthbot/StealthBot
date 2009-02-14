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

    On Error GoTo ERROR_HANDLER

    Dim f         As Integer ' ...
    Dim filePath  As String  ' ...
    
    ' ...
    Select Case (ltype)
        Case WARNING_MSG, ERROR_MSG, EVENT_MSG, DEBUG_MSG
            filePath = _
                (App.Path & "\logs\master.txt")
                
        Case CHAT_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & ".txt")
                
        Case WHISPER_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & "-WHISPERS.txt")
                
        Case SCK_IN_MSG, SCK_OUT_MSG
            filePath = _
                (App.Path & "\logs\" & Datestamp() & "-PACKETLOG.txt")
    End Select
    
    ' ...
    f = OpenLog(filePath)
    
    ' ...
    If ((ltype = CHAT_MSG) Or (ltype = WHISPER_MSG)) Then
        Print #f, "[" & Timestamp() & "] " & line
    Else
        Print #f, "[" & Datestamp() & " " & Timestamp() & "] " & line
    End If
    
    ' ...
    Close #f
    
    Exit Function
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in WriteLine()."

End Function

' ...
Private Function OpenLog(ByVal Path As String) As Integer

    On Error GoTo ERROR_HANDLER
    
    Dim f As Integer ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    If (dir$(Path) = vbNullString) Then
        Dim dir_path As String ' ...
        
        ' ...
        dir_path = Mid$(Path, 1, InStrRev(Path, "\"))
    
        ' ...
        If (dir$(dir_path) = vbNullString) Then
            MkDir dir_path
        End If
    
        ' ...
        Open Path For Output As #f
    Else
        Dim arr() As String  ' ...
        Dim I     As Integer ' ...
        Dim str   As String  ' ...
        Dim bln   As Boolean ' ...
    
        ' ...
        Open Path For Input As #f
            If (BotVars.MaxLogFileSize) Then
                If (LOF(f) >= BotVars.MaxLogFileSize) Then
                    ReDim arr(LOF(f) - 1)
                    
                    For I = 0 To LOF(f)
                        Line Input #f, str
                    
                        If (I > 0) Then
                            arr(I - 1) = str
                        End If
                    Next I
                    
                    bln = True
                End If
            End If
        Close #f
    
        ' ...
        If (bln) Then
            Open Path For Binary As #f
                Put #f, , arr()
            Close #f
        End If
        
        ' ...
        Open Path For Append As #f
    End If
    
    ' ...
    OpenLog = f
    
    Exit Function
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in OpenLog()."
    
End Function

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
