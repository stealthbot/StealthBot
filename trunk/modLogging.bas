Attribute VB_Name = "modLogging"
' modLogging
' project StealthBot
' August 2005, on a plane to Philadelphia
'  and later in a car to Providence
'  and later in a plane to Milwaukee

' Created to enhance and clean up event and text logging
Private iActiveDay As Integer

' START ENHANCED LOGGING
' Call before using any methods in this module
Public Sub StartEnhancedLogging()
    If LenB(Dir$(GetProfilePath() & "\LogsHTML\logstyle.css")) = 0 Then
        Call CreateLogfileCSS
    End If
End Sub

' GET LOG FILENAME
' Returns the appropriate filename for the CURRENT, ACTIVE StealthBot log
Public Function GetLogFilename() As String
    Static sPath As String
    
    If LenB(sPath) = 0 Or (GetActiveDay() <> iActiveDay) Then
        sPath = GetProfilePath & "\LogsHTML\" & Format(Date, "mm-dd-yy") & ".log.html"
    End If
    
    GetLogFilename = sPath
End Function

' GET ACTIVE DAY
' Returns the current DAY as an int
Public Function GetActiveDay() As Integer
    GetActiveDay = CInt(Format(Date, "dd"))
End Function

' OPEN LOGFILE
' Determines whether or not the logfile for today exists
' If it does not exist, creats and opens it and returns the filenumber
' If it exists, opens it for binary access write and returns the filenumber
Public Function OpenLogfile() As Integer
    Dim f            As Integer
    Dim sLogFilename As String
    
    sLogFilename = GetLogFilename
    f = FreeFile
    
    If LenB(Dir$(sLogFilename)) > 0 Then
        Open sLogFilename For Append As #f
    Else
        Open sLogFilename For Output As #f
        
        Print #f, "<html><head><title>"
        Print #f, "StealthBot Log: ";
        
        If (LenB(CurrentUsername) > 0) Then
            Print #f, CurrentUsername;
        Else
            If LenB(BotVars.Username) > 0 Then
                Print #f, BotVars.Username;
            Else
                Print #f, "(not configured)";
            End If
        End If
        
        Print #f, " on " & Format(Date, "m/d/yyyy") & "</title>"
        Print #f, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #f, "<LINK REL='stylesheet' HREF='logstyle.css' TYPE='text/css'>"
        Print #f, "</head><body bgcolor='#000000'><span class='title'>Log for ";
        
        If (LenB(CurrentUsername) > 0) Then
            Print #f, CurrentUsername;
        Else
            If LenB(BotVars.Username) > 0 Then
                Print #f, BotVars.Username;
            Else
                Print #f, "(not configured)";
            End If
        End If
        
        Print #f, " on " & Format(Date, "m/d/yyyy"); "</span><br>"
    End If
    
    OpenLogfile = f
End Function

Public Sub CreateLogfileCSS()
    Dim f As Integer
    
    f = FreeFile
    Open GetProfilePath() & "\LogsHTML\logstyle.css" For Output As #f
        Print #f, "<style type='text/css'>"
        Print #f, "BODY {"
        Print #f, "     background-color: #000000;"
        Print #f, "}"
        Print #f, ".title {"
        Print #f, "     font-size: 12pt;"
        Print #f, "     font-family: Tahoma, Helvetica, sans-serif;"
        Print #f, "     color: #FFFFFF;"
        Print #f, "}"
        Print #f, "</style>"
    Close #f
End Sub

Public Sub CloseLogfile(ByVal f As Integer)
    Close #f
End Sub

Public Function HTMLSanitize(ByVal sInput As String) As String
    sInput = Replace(sInput, "<", "&lt;")
    sInput = Replace(sInput, ">", "&gt;")
    sInput = Replace(sInput, Chr(34), "&quot;")
    sInput = Replace(sInput, "&", "&amp;")
    sInput = Replace(sInput, "&amp;amp;", "&amp;")
    
    HTMLSanitize = sInput
End Function

' Written 2007-06-08 to produce packet logs or do other things
'  -at
Public Sub LogPacketRaw(ByVal Server As enuPL_ServerTypes, ByVal Direction As enuPL_DirectionTypes, ByVal PacketID As Long, ByVal PacketLen As Long, ByRef PacketData As String)
    Dim l As Long
    
    If (LogPacketTraffic) Then
        'Log this packet!
        l = FreeFile
        
        Open PacketLogFilePath For Append As #l
            Print #l, GetTimestamp() & " "
        
            Select Case (Server)
                Case stBNCS
                    Print #l, "BNCS";
                Case stBNLS
                    Print #l, "BNLS";
            End Select
            
            Select Case (Direction)
                Case CtoS
                    Print #l, " C->S";
                Case StoC
                    Print #l, " S->C";
            End Select
            
            Print #l, " -- Packet ID " & Right$("00" & Hex(PacketID), 2) & _
                "h (" & PacketID & "d) length " & PacketLen
            Print #l, vbNullString
            Print #l, DebugOutput(PacketData)
            Print #l, vbCrLf
        Close #l
        
        l = 0
    End If
End Sub
