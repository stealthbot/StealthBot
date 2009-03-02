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

    Dim strPath  As String  ' ...
    Dim fileName As String  ' ...
    Dim i        As Integer ' ...
    
    ' ********************************
    '      LOAD REGULAR SCRIPTS
    ' ********************************

    ' ...
    strPath = App.Path & "\scripts\"
    
    ' ...
    If (Dir(strPath) <> vbNullString) Then
        ' ...
        fileName = Dir(strPath)
        
        ' ...
        Do While (fileName <> vbNullString)
            ' ...
            Set CurrentModule = SC.Modules.Add(fileName)
            
            ' ...
            FileToModule CurrentModule, strPath & fileName
    
            ' ...
            fileName = Dir()
        Loop
    End If
    
    ' ********************************
    '      LOAD PLUGIN SYSTEM
    ' ********************************

    ' ...
    If (ReadINI("Override", "DisablePS", GetConfigFilePath()) <> "Y") Then
        ' ...
        strPath = GetFilePath("PluginSystem.dat")
        
        ' ...
        If (LenB(Dir$(strPath)) = 0) Then
            Call frmChat.AddChat(vbRed, "Cannot find PluginSystem.dat. It must exist in order to load plugins!")
            Call frmChat.AddChat(vbYellow, "You may download PluginSystem.dat to your StealthBot folder using the link below.")
            Call frmChat.AddChat(vbWhite, "http://www.stealthbot.net/p/Users/Swent/index.php?file=PluginSystem.dat")
        Else
            FileToModule SC.Modules(1), strPath
        End If
    End If
        
    ' ...
    Exit Sub

' ...
ERROR_HANDLER:

    ' ...
    frmChat.AddChat vbRed, "Error: " & Err.description & " in LoadScripts()."

    ' ...
    Exit Sub

End Sub

Public Function InitScripts()

    Dim i As Integer ' ...

    ' ...
    RunInAll "Event_Load"

    ' ...
    If (g_Online) Then
        RunInAll "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        RunInAll "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If (g_Channel.Users.Count > 0) Then
            For i = 1 To g_Channel.Users.Count
                With g_Channel.Users(i)
                     RunInAll "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
                        .game, False
                End With
             Next i
         End If
    End If

End Function

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

Public Sub RunInAll(ParamArray Parameters() As Variant)

    On Error GoTo ERROR_HANDLER

    Dim SC    As ScriptControl
    Dim i     As Integer ' ...
    Dim arr() As Variant ' ...
    
    ' ...
    Set SC = frmChat.SControl
    
    ' ...
    arr() = Parameters()

    ' ...
    For i = 1 To SC.Modules.Count
        CallByNameEx SC.Modules(i), "Run", VbMethod, arr()
    Next

    Exit Sub
    
ERROR_HANDLER:
    ' object does not support property or method
    If (Err.Number = 438) Then
        Resume Next
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in RunInAll()."
    
    Exit Sub
    
End Sub

Public Function CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray _
    As Variant)
    
    On Error GoTo ERROR_HANDLER
    
    Dim oTLI    As TLI.TLIApplication
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
    
    Set oTLI = Nothing
    
    Exit Function

ERROR_HANDLER:
    ' ...
    If (frmChat.SControl.Error) Then
        Exit Function
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in CallByNameEx()."
        
    Set oTLI = Nothing
        
    Exit Function
    
End Function

'// Loads the Plugin System
'//   Called from Form_Load() and mnuReloadScript_Click() in frmChat
Public Sub LoadPluginSystem(ByRef SC As ScriptControl)
    Dim Path As String, intFile As Integer, strLine As String, strContent As String, strOverride As String

   On Error GoTo LoadPluginSystem_Error

    strOverride = ReadINI("Override", "DisablePS", GetConfigFilePath())
    
    If ReadINI("Override", "DisablePS", GetConfigFilePath()) = "Y" Then
        boolOverride = True
    Else
        If Len(strOverride) = 0 Then WriteINI "Override", "DisablePS", "N"
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
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "Failed to load script.txt (file does not exist).")
            Exit Sub
        End If
    End If
    
    '// Create scripting objects
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars
    
    'Exit Sub
    
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

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure LoadPluginSystem of Module modScripting"
    Debug.Print "Using variable: " & Path
End Sub


Public Sub SetVeto(ByVal B As Boolean)
    VetoNextMessage = B
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
    SC.Run "Event_Load"

    If g_Online Then
        SC.Run "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        SC.Run "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If g_Channel.Users.Count > 0 Then
            For i = 1 To g_Channel.Users.Count
                Message = ""

                With g_Channel.Users(i)
                     'ParseStatstring .Statstring, Message, .Clan

                     SC.Run "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, .game, False
                 End With
             Next i
         End If
    End If

    On Error GoTo 0
    Exit Sub

ReInitScriptControl_Error:

    'Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ReInitScriptControl of Module modScripting"
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
