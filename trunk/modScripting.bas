Attribute VB_Name = "modScripting"
'/* Scripting.bas
' * ~~~~~~~~~~~~~~~~
' * StealthBot VBScript support module
' * ~~~~~~~~~~~~~~~~
' * Modified by Swent 11/06/2007
' */
Option Explicit

Public VetoNextMessage As Boolean
Public boolOverride As Boolean

'// Loads the Plugin System
'//   Called from Form_Load() and mnuReloadScript_Click() in frmChat
Public Sub LoadPluginSystem(ByRef SC As ScriptControl)
    Dim Path As String, intFile As Integer, strLine As String, strContent As String

   On Error GoTo LoadPluginSystem_Error
   
   ' ...
    Call InitScriptControl(SC)
    Call LoadScripts(SC)
    
    Exit Sub

    If ReadINI("Override", "DisablePS", GetConfigFilePath()) = "Y" Then
        boolOverride = True
    Else
        boolOverride = False
    End If

    '// Reset the Script Control
    SC.Reset
    
    '// Allow UI's unless they've been disabled by the user
    If ReadINI("Other", "ScriptAllowUI", GetConfigFilePath()) <> "N" Then SC.AllowUI = True

    If Not boolOverride Then
    
        '// PluginSystem.dat exists?
        Path = GetFilePath("PluginSystem.dat")
        If LenB(Dir$(Path)) = 0 Then
            Call frmChat.AddChat(vbRed, "Cannot find PluginSystem.dat. It must exist in order to load plugins!")
            Call frmChat.AddChat(vbYellow, "You may download PluginSystem.dat to your StealthBot folder using the link below.")
            Call frmChat.AddChat(vbWhite, "http://www.stealthbot.net/p/Users/Swent/index.php?file=PluginSystem.dat")
            Exit Sub
        End If
    Else
    
        '// script.txt exists?
        Path = GetFilePath("script.txt")
        If LenB(Dir$(Path)) = 0 Then
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "No script.txt file is present. It must exist, if only to #include other files!")
            Exit Sub
        End If
    End If
    
    '// Create scripting objects
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars
    
    If Not boolOverride Then
    
        '// Load PluginSystem.dat
        intFile = FreeFile
        Open Path For Input As #intFile
    
            Do While Not EOF(intFile)
                strLine = vbNullString
                Line Input #intFile, strLine
                If Len(strLine) > 1 Then strContent = strContent & strLine & vbCrLf
            Loop
        Close #intFile
        SC.AddCode strContent
    Else
        Dim strFilesToLoad() As String, i As Integer
        ReDim strFilesToLoad(0)
        strFilesToLoad(0) = "script.txt"
    
        intFile = FreeFile
        Path = GetFilePath("script.txt")
        
        '// Get names of includes (if any)
        Open Path For Input As #intFile
            Do While Not EOF(intFile)
                strLine = ""
                Line Input #intFile, strLine
                
                If Len(strLine) > 1 Then
                
                    If Left$(Trim(LCase(strLine)), 8) = "#include" And Len(strLine) > 10 Then
                        ReDim Preserve strFilesToLoad(UBound(strFilesToLoad) + 1)
                        strFilesToLoad(UBound(strFilesToLoad)) = Mid(strLine, 10)
                    ElseIf Left(Trim(LCase(strLine)), 3) = "sub" Then
                        Exit Do
                    End If
                End If
            Loop
        Close #intFile
    
        '// Load script.txt and any includes
        For i = 0 To UBound(strFilesToLoad)
            strContent = ""
            intFile = FreeFile
            Path = GetFilePath(strFilesToLoad(i))
            
            Open Path For Input As #intFile
    
                Do While Not EOF(intFile)
                    strLine = ""
                    Line Input #intFile, strLine
                    
                    If Len(strLine) > 1 Then
                        If i = 0 Then
                            If Left$(Trim(LCase(strLine)), 8) = "#include" Then strLine = ""
                        End If
                        strContent = strContent & strLine & vbCrLf
                    End If
                Loop
            Close #intFile
            SC.AddCode strContent
            Call frmChat.AddChat(vbGreen, "Script loaded: " & Replace(Path, "\\", "\"))
        Next
    End If

LoadScript_Exit:

    Exit Sub
   
LoadPluginSystem_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadPluginSystem of Module modScripting"
    Debug.Print "Using variable: " & Path
End Sub

Public Sub InitScriptControl(ByRef SC As ScriptControl)

    ' ...
    SC.Reset
    
    ' ...
    If (ReadINI("Other", "ScriptAllowUI", GetConfigFilePath()) <> "N") Then
        SC.AllowUI = True
    End If

    '// Create scripting objects
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars

End Sub

Public Sub LoadScripts(ByRef SC As ScriptControl)

    ' ...
    On Error GoTo ERROR_HANDLER

    Dim CurrentModule As Module

    Dim strPath       As String  ' ...
    Dim filename      As String  ' ...
    Dim f             As Integer ' ...
    Dim strLine       As String  ' ...
    Dim strContent    As String  ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    strPath = App.Path & "\scripts\"
    
    ' ...
    If (Dir(strPath) = vbNullString) Then
        Exit Sub
    End If

    ' ...
    filename = Dir(strPath)
    
    ' ...
    Do While (filename <> vbNullString)
        ' ...
        Set CurrentModule = SC.Modules.Add(filename)
        
        ' ...
        CreateModuleProcs SC, filename
    
        ' ...
        FileToModule CurrentModule, strPath & filename
        
        ' ...
        frmChat.AddChat vbGreen, "Script loaded: " & filename

        ' ...
        filename = Dir()
    Loop
    
    ' ...
    Exit Sub

' ...
ERROR_HANDLER:

    ' ...
    frmChat.AddChat vbRed, "Error: " & Err.Description & " in LoadScripts()."

    ' ...
    Exit Sub

End Sub

Private Sub CreateModuleProcs(ByRef SC As ScriptControl, ByVal filename As String)

    Dim CurrentModule As Module
    Dim arrCode()     As String ' ...

    ' ...
    Set CurrentModule = SC.Modules(SC.Modules.Count)
    
    ' ...
    ReDim arrCode(0 To 3)
    
    ' ...
    arrCode(0) = "Function GetModuleID()"
    arrCode(1) = "   GetModuleID = " & SC.Modules.Count
    arrCode(2) = "End Function"
    
    ' ...
    CurrentModule.AddCode Join(arrCode, vbCrLf)
    
    ' ...
    arrCode(0) = "Function CreateTimer(TimerName, TimerInterval)"
    arrCode(1) = "   Set CreateTimer = CreateTimerEx(GetModuleID, TimerName, TimerInterval)"
    arrCode(2) = "End Function"
    
    ' ...
    CurrentModule.AddCode Join(arrCode, vbCrLf)
    
    ' ...
    arrCode(0) = "Sub DeleteTimer(TimerName)"
    arrCode(1) = "   DeleteTimerEx GetModuleID, TimerName"
    arrCode(2) = "End Sub"
    
    ' ...
    CurrentModule.AddCode Join(arrCode, vbCrLf)

    ' ...
    arrCode(0) = "Function Timer(TimerName)"
    arrCode(1) = "   Set Timer = TimerEx(GetModuleID, TimerName)"
    arrCode(2) = "End Function"
    
    ' ...
    CurrentModule.AddCode Join(arrCode, vbCrLf)
    
    ' ...
    arrCode(0) = "Function GetScriptControl()"
    arrCode(1) = "   Set GetScriptControl = GetScriptControlEx(GetModuleID)"
    arrCode(2) = "End Function"
    
    ' ...
    CurrentModule.AddCode Join(arrCode, vbCrLf)

End Sub

Private Function FileToModule(ByRef ScriptModule As Module, ByVal FilePath As String)

    Dim strLine    As String  ' ...
    Dim strContent As String  ' ...
    Dim f          As Integer ' ...
    
    ' ...
    f = FreeFile

    ' ...
    Open FilePath For Input As #f
        ' ...
        Do While (Not (EOF(f)))
            ' ...
            Line Input #f, strLine
            
            ' ...
            If (Len(strLine) > 1) Then
                ' ...
                If (StrComp(Left$(Trim(LCase$(strLine)), 8), "#include", vbTextCompare) <> 0) Then
                    strContent = strContent & strLine & vbCrLf
                Else
                    ' ...
                    If (Len(Trim(LCase$(strLine))) > 8) Then
                        FileToModule ScriptModule, Mid$(Trim(LCase$(strLine)), 9)
                    End If
                End If
            End If
            
            ' ...
            strLine = vbNullString
        Loop
    Close #f
    
    ' ...
    ScriptModule.AddCode strContent

End Function

Public Sub RunInAll(ByRef SC As ScriptControl, ParamArray Parameters() As Variant)

    On Error GoTo ERROR_HANDLER

    Dim i     As Integer ' ...
    Dim arr() As Variant ' ...
    
    ' ...
    arr() = Parameters()

    ' ...
    For i = 2 To SC.Modules.Count
        CallByNameEx SC.Modules(i), "Run", VbMethod, arr()
    Next

    Exit Sub
    
ERROR_HANDLER:
    ' object does not support property or method
    If (Err.Number = 438) Then
        Exit Sub
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.Description & _
        " in RunInAll()."
    
    Exit Sub
    
End Sub

Public Sub SetVeto(ByVal b As Boolean)

    VetoNextMessage = b

End Sub


Public Function GetVeto() As Boolean

    GetVeto = VetoNextMessage
    VetoNextMessage = False
    
End Function


Public Sub ReInitScriptControl(ByRef SC As ScriptControl)

    Dim i As Integer
    Dim Message As String

    On Error GoTo ReInitScriptControl_Error

    BotLoaded = True
    RunInAll frmChat.SControl, "Event_Load"

    If g_Online Then
        RunInAll frmChat.SControl, "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        RunInAll frmChat.SControl, "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If g_Channel.Users.Count > 0 Then
            For i = 1 To g_Channel.Users.Count
                Message = ""

                With g_Channel.Users(i)
                     ParseStatstring .Statstring, Message, .Clan

                     RunInAll frmChat.SControl, "Event_UserInChannel", .DisplayName, .Flags, Message, .Ping, .game, False
                 End With
             Next i
         End If
    End If

    Exit Sub

ReInitScriptControl_Error:
    frmChat.AddChat vbRed, "Error: " & Err.Description & " in ReInitScriptControl()"
    
    Exit Sub
    
End Sub

Public Function CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray _
    As Variant)
    
    On Error GoTo ERROR_HANDLER
    
    Dim oTLI    As TLIApplication
    Dim ProcID  As Long
    Dim numArgs As Long
    Dim i       As Long
    Dim v()     As Variant
    
    Set oTLI = New TLIApplication

    ProcID = oTLI.InvokeID(obj, ProcName)

    If (IsMissing(vArgsArray)) Then
        CallByNameEx = oTLI.InvokeHook(obj, ProcID, CallType)
    End If
    
    If (IsArray(vArgsArray)) Then
        numArgs = UBound(vArgsArray)
        
        ReDim v(numArgs)
        
        For i = 0 To numArgs
            v(i) = vArgsArray(numArgs - i)
        Next i
        
        CallByNameEx = oTLI.InvokeHookArray(obj, ProcID, CallType, v)
    End If
    
    Exit Function

ERROR_HANDLER:
    ' ...
    If (frmChat.SControl.Error) Then
        Exit Function
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.Description & _
        " in CallByNameEx()."
        
    Exit Function
End Function


