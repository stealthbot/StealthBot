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

Public dictSettings As Dictionary
Public dictTimerInterval As Dictionary
Public dictTimerEnabled As Dictionary
Public dictTimerCount As Dictionary



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
        CallByNameEx SC.Modules(i), "Run", VbMethod, arr
    Next

    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.Description & " in RunInAll()."
    
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


'// Written by Swent. Sets a plugin timer's interval.
Public Sub SetPTInterval(ByVal strPrefix As String, ByVal strTimerName As String, ByVal intInterval As Integer)
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName

    dictTimerInterval(strKey) = intInterval
    dictTimerCount(strKey) = intInterval
    
    If Not dictTimerEnabled.Exists(strKey) Then
       dictTimerEnabled(strKey) = False
    End If
End Sub


'// Written by Swent. Enables or disables a plugin timer.
Public Sub SetPTEnabled(ByVal strPrefix As String, ByVal strTimerName As String, ByVal boolEnabled As Boolean)
    
    dictTimerEnabled(strPrefix & ":" & strTimerName) = boolEnabled
End Sub

Public Function CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray _
    As Variant)
    
    On Error GoTo Handler
    
    Dim oTLI    As Object
    Dim ProcID  As Long
    Dim numArgs As Long
    Dim i       As Long
    Dim v()     As Variant

    Set oTLI = CreateObject("TLI.TLIApplication")
    
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

Handler:
    frmChat.AddChat vbRed, "Error: " & Err.Description & " in CallByNameEx()"

    Exit Function
End Function


'// Written by Swent. Modifies the count in a running plugin timer.
Public Sub SetPTCount(ByVal strPrefix As String, ByVal strTimerName As String, ByVal intCount As Integer)
    
    dictTimerCount(strPrefix & ":" & strTimerName) = intCount
End Sub


'// Written by Swent. Gets the enabled status of a plugin timer.
Public Function GetPTEnabled(ByVal strPrefix As String, ByVal strTimerName As String)
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerEnabled.Exists(strKey) Then
        GetPTEnabled = dictTimerEnabled(strKey)
    Else
        GetPTEnabled = -1
    End If
End Function


'// Written by Swent. Gets a plugin timer's interval setting.
Public Function GetPTInterval(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerInterval.Exists(strKey) Then
        GetPTInterval = dictTimerInterval(strKey)
    Else
        GetPTInterval = -1
    End If
End Function


'// Written by Swent. Get's the seconds left before a plugin timer sub executes.
Public Function GetPTLeft(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName

    If dictTimerCount.Exists(strKey) Then
        GetPTLeft = dictTimerCount(strKey)
    Else
        GetPTLeft = -1
    End If
End Function


'// Written by Swent. Gets the time since a plugin timer sub was last executed.
Public Function GetPTWaiting(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerCount.Exists(strKey) Then
        GetPTWaiting = dictTimerInterval(strKey) - dictTimerCount(strKey) + 1
    Else
        GetPTWaiting = -1
    End If
End Function


'// Written by Swent. Gets keys for the timer dictionaries.
Public Function GetPTKeys() As String

    GetPTKeys = Join(dictTimerEnabled.Keys)
End Function


