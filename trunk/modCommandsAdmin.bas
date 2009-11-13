Attribute VB_Name = "modCommandsAdmin"
Option Explicit
'This module will hold all the commands that relate to Andmistering the bot, Changing settings
'Editing the database, etc..

'This is a stub function for now, it still calls the old uber complicated OnAddOld function, but hey :/
Public Sub OnAdd(Command As clsCommandObj)
    Dim dbAccess   As udtGetAccessResponse
    Dim response() As String
    Dim i          As Integer
    ReDim Preserve response(0)
    
    dbAccess = GetCumulativeAccess(Command.Username)
    If (Command.IsLocal) Then
        dbAccess.Rank = 201
        dbAccess.Flags = "A"
    End If
    
    Call OnAddOld(Command.Username, dbAccess, Command.Args, Command.IsLocal, response())
    
    For i = LBound(response) To UBound(response)
        Command.Respond response(i)
    Next i
End Sub

Public Sub OnClear(Command As clsCommandObj)
    frmChat.mnuClear_Click
End Sub

Public Sub OnDisable(Command As clsCommandObj)
    Dim Module As Module
    Dim Name   As String
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled via the override."
        Exit Sub
    End If
    
    If (Command.IsValid) Then
        Set Module = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Module Is Nothing) Then
            Command.Respond "Error: Could not find specified script."
        Else
            Name = modScripting.GetScriptName(Module.Name)
            If (StrComp(SharedScriptSupport.GetSettingsEntry("Enabled", Name), "False", vbTextCompare) = 0) Then
                Command.Respond Name & " is already disabled."
            Else
                RunInSingle Module, "Event_Close"
                SharedScriptSupport.WriteSettingsEntry "Enabled", "False", , Name
                modScripting.DestroyObjs Module
                Command.Respond Name & " has been disabled."
            End If
        End If
    Else
        Command.Respond "Error: You must specify a script."
    End If
End Sub

Public Sub OnDump(Command As clsCommandObj)
    Call DumpPacketCache
End Sub

Public Sub OnEnable(Command As clsCommandObj)
    Dim Module As Module
    Dim Name   As String
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled via the override."
        Exit Sub
    End If
    
    If (Command.IsValid) Then
        Set Module = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Module Is Nothing) Then
            Command.Respond "Error: Could not find specified script."
        Else
            Name = modScripting.GetScriptName(Module.Name)
            If (StrComp(SharedScriptSupport.GetSettingsEntry("Enabled", Name), "True", vbTextCompare) = 0) Then
                Command.Respond Name & " is already enabled."
            Else
                SharedScriptSupport.WriteSettingsEntry "Enabled", "True", , Name
                modScripting.InitScript Module
                Command.Respond Name & " has been enabled."
            End If
        End If
    Else
        Command.Respond "Error: You must specify a script."
    End If
End Sub

Public Sub OnLockText(Command As clsCommandObj)
    Call frmChat.mnuLock_Click
End Sub

Public Sub OnQuit(Command As clsCommandObj)
    BotIsClosing = True
    Unload frmChat
    Set frmChat = Nothing
End Sub

'This is a stub function for now, it still calls the old uber complicated OnRemOld function, but hey :/
Public Sub OnRem(Command As clsCommandObj)
    Dim dbAccess   As udtGetAccessResponse
    Dim response() As String
    Dim i          As Integer
    ReDim Preserve response(0)
    
    dbAccess = GetCumulativeAccess(Command.Username)
    If (Command.IsLocal) Then
        dbAccess.Rank = 201
        dbAccess.Flags = "A"
    End If
    
    Call OnRemOld(Command.Username, dbAccess, Command.Args, Command.IsLocal, response())
    
    For i = LBound(response) To UBound(response)
        Command.Respond response(i)
    Next i
End Sub

Public Sub OnSetCommandLine(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call SetCommandLine(Command.Argument("CommandLine"))
        If (LenB(CommandLine) > 0) Then
            Command.Respond StringFormat("Command Line set to: {0}", CommandLine)
        Else
            Command.Respond "Command Line cleared."
        End If
    Else
        Command.Respond "You must specify a new command line."
    End If
End Sub

Public Sub OnSetExpKey(Command As clsCommandObj)
    Dim strKey As String
    If (Command.IsValid) Then
        strKey = Replace$(Command.Argument("Key"), "-", vbNullString)
        strKey = Replace$(strKey, Space$(1), vbNullString)
        
        If (Not StrComp(ReadCfg("Override", "SetKeyIgnoreLength"), "Y", vbTextCompare) = 0) Then
            If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
                strKey = vbNullString
            End If
        End If
        If (LenB(strKey) > 0) Then
            Call WriteINI("Main", "ExpKey", strKey)
            BotVars.ExpKey = strKey
            Command.Respond "New expansion cdkey set."
        Else
            Command.Respond "The cdkey you specified was invalid."
        End If
    Else
        Command.Respond "You must specify a cdkey."
    End If
End Sub

Public Sub OnSetHome(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call WriteINI("Main", "HomeChan", Command.Argument("Channel"))
        BotVars.HomeChannel = Command.Argument("Channel")
        Command.Respond StringFormat("New home channel set to {0}{1}{0}.", Chr$(34), BotVars.HomeChannel)
    Else
        Command.Respond "You must specify a channel."
    End If
End Sub

Public Sub OnSetKey(Command As clsCommandObj)
    Dim strKey As String
    If (Command.IsValid) Then
        strKey = Replace$(Command.Argument("Key"), "-", vbNullString)
        strKey = Replace$(strKey, Space$(1), vbNullString)
        
        If (Not StrComp(ReadCfg("Override", "SetKeyIgnoreLength"), "Y", vbTextCompare) = 0) Then
            If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
                strKey = vbNullString
            End If
        End If
        If (LenB(strKey) > 0) Then
            Call WriteINI("Main", "CDKey", strKey)
            BotVars.CDKey = strKey
            Command.Respond "New cdkey set."
        Else
            Command.Respond "The cdkey you specified was invalid."
        End If
    Else
        Command.Respond "You must specify a cdkey."
    End If
End Sub

Public Sub OnSetName(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call WriteINI("Main", "Username", Command.Argument("Username"))
        BotVars.Username = Command.Argument("Username")
        Command.Respond "New username set to " & BotVars.Username
    Else
        Command.Respond "You must specify a username."
    End If
End Sub

Public Sub OnSetPass(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call WriteINI("Main", "Password", Command.Argument("Password"))
        BotVars.Password = Command.Argument("Password")
        Command.Respond "New password set."
    Else
        Command.Respond "You must specify a password."
    End If
End Sub

Public Sub OnSetPMsg(Command As clsCommandObj)
    If (Command.IsValid) Then
        ProtectMsg = Command.Argument("Message")
        Call WriteINI("Other", "ProtectMsg", Command.Argument("Message"))
        Command.Respond "Channel protection message set."
    Else
        Command.Respond "You must specify a message."
    End If
End Sub

Public Sub OnSetServer(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call WriteINI("Main", "Server", Command.Argument("Server"))
        BotVars.Server = Command.Argument("Server")
        Command.Respond "New server set to " & BotVars.Server
    Else
        Command.Respond "You must specify a server."
    End If
End Sub

Public Sub OnSetTrigger(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call WriteINI("Main", "Trigger", StringFormat("{{0}}", Command.Argument("Trigger")))
        BotVars.Trigger = Command.Argument("Trigger")
        Command.Respond StringFormat("The new trigger is {0}{1}{0}.", Chr$(34), BotVars.Trigger)
    Else
        Command.Respond "You must specify a trigger."
    End If
End Sub

Public Sub OnWhisperCmds(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on":
            BotVars.WhisperCmds = True
            Call WriteINI("Main", "WhisperBack", "Y")
            Command.Respond "Command responses will now be whispered back."
        
        Case "off":
            BotVars.WhisperCmds = False
            Call WriteINI("Main", "WhisperBack", "N")
            Command.Respond "Command responses will now be displayed publicly."
        
        Case Else:
            Command.Respond StringFormat("Command responses will be {0}.", _
                IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly"))
    End Select
End Sub
