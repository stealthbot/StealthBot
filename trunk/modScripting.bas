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
            AddChat vbRed, "Cannot find PluginSystem.dat. It must exist in order to load plugins!"
            Exit Sub
        End If
    Else
    
        '// script.txt exists?
        Path = GetFilePath("script.txt")
        If LenB(Dir$(Path)) = 0 Then
            AddChat vbRed, "No script.txt file is present. It must exist, if only to #include other files!"
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
    
        '// Create timer dictionaries
        Set dictTimerInterval = New Dictionary
        Set dictTimerEnabled = New Dictionary
        Set dictTimerCount = New Dictionary
        dictTimerInterval.CompareMode = TextCompare
        dictTimerEnabled.CompareMode = TextCompare
        dictTimerCount.CompareMode = TextCompare
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
            AddChat vbGreen, "Script loaded: " & Replace(Path, "\\", "\")
        Next
    End If

LoadScript_Exit:

    Exit Sub
   
LoadPluginSystem_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadPluginSystem of Module modScripting"
    Debug.Print "Using variable: " & Path
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
    SC.Run "Event_Load"

    If g_Online Then
        SC.Run "Event_LoggedOn", CurrentUsername, BotVars.Product
        SC.Run "Event_ChannelJoin", gChannel.Current, 0

        If colUsersInChannel.Count > 0 Then
            For i = 1 To colUsersInChannel.Count
                Message = ""

                With colUsersInChannel.Item(i)
                     ParseStatstring .Statstring, Message, .Clan

                     SC.Run "Event_UserInChannel", .Username, .Flags, Message, .Ping, .Product, False
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


'// Written by Swent. Gets the number of plugin timers.
Public Function GetNumPT() As Integer

    GetNumPT = dictTimerEnabled.Count
End Function
