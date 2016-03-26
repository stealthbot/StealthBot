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
    
    ' special case: d2 naming conventions
    If (BotVars.UseD2Naming) Then
        Dim Username
        Username = Command.Argument("Username")
        If (Len(Username) > 1) Then
            If (Left$(Username, 1) = "*") Then
                If (InStr(2, Username, "*") = 0 And InStr(2, Username, "?") = 0) Then
                    ' format: *user
                    ' assume user means: user
                    Command.Args = Mid$(Command.Args, 2)
                End If
            End If
            
            If (InStr(Username, "*") = 0 And InStr(Username, "?") = 0) Then
                ' format: charname
                Dim User As clsUserObj, IsChar As Boolean, Acct As String
                For i = 1 To g_Channel.Users.Count
                    Set User = g_Channel.Users(i)
                    If (StrComp(User.CharacterName, Username, vbTextCompare) = 0) Then
                        ' the user provided is a character in the channel
                        IsChar = True
                        Acct = User.DisplayName
                    End If
                    If (StrComp(User.DisplayName, Username, vbTextCompare) = 0) Then
                        ' if the user provided is ALSO an accountname, assume account name
                        IsChar = False
                        Exit For
                    End If
                Next i
                If (IsChar) Then
                    Command.Args = Replace$(Command.Args, Username, Acct, 1, 1, vbBinaryCompare)
                End If
            End If
        End If
    End If
    
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
    frmChat.ClearChatScreen
End Sub

Public Sub OnDisable(Command As clsCommandObj)
    Dim Module As Module
    Dim Name   As String
    Dim i      As Integer
    Dim mnu    As clsMenuObj
    
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
                For i = 1 To DynamicMenus.Count
                    Set mnu = DynamicMenus(i)
                    If (StrComp(mnu.Name, Chr$(0) & Name & Chr$(0) & "ENABLE|DISABLE", vbTextCompare) = 0) Then
                        mnu.Checked = False
                        Exit For
                    End If
                Next i
                Set mnu = Nothing
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
    Dim i      As Integer
    Dim mnu    As clsMenuObj
    
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
                For i = 1 To DynamicMenus.Count
                    Set mnu = DynamicMenus(i)
                    If (StrComp(mnu.Name, Chr$(0) & Name & Chr$(0) & "ENABLE|DISABLE", vbTextCompare) = 0) Then
                        mnu.Checked = True
                        Exit For
                    End If
                Next i
                Set mnu = Nothing
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

Public Sub OnSetBnlsServer(Command As clsCommandObj)
    If (Command.IsValid) Then
        Config.BnlsServer = Command.Argument("Server")
        Call Config.Save
        
        BotVars.BnlsServer = Config.BnlsServer
        Command.Respond StringFormat("New BNLS server set to {0}{1}{0}.", Chr$(34), BotVars.BnlsServer)
    Else
        Command.Respond "You must specify a server."
    End If
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
        
        If Config.SetKeyIgnoreLength Then
            If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
                strKey = vbNullString
            End If
        End If
        If (LenB(strKey) > 0) Then
            Config.ExpKey = strKey
            Call Config.Save
            
            BotVars.ExpKey = Config.ExpKey
            Command.Respond "New expansion cdkey set."
        Else
            Command.Respond "The cdkey you specified was invalid."
        End If
    Else
        Command.Respond "You must specify a cdkey."
    End If
End Sub

Public Sub OnSetHome(Command As clsCommandObj)
    Dim Channel As String
    If (Command.IsValid) Then
        Channel = Command.Argument("Channel")
        Config.HomeChannel = Channel
        Call Config.Save
        
        BotVars.HomeChannel = Config.HomeChannel
        If LenB(Channel) = 0 Then
            Command.Respond StringFormat("Reset home channel to server default.", Chr$(34), BotVars.HomeChannel)
        Else
            Command.Respond StringFormat("New home channel set to {0}{1}{0}.", Chr$(34), BotVars.HomeChannel)
        End If
    Else
        Command.Respond "Homechannel command invalid."
    End If
End Sub

Public Sub OnSetKey(Command As clsCommandObj)
    Dim strKey As String
    If (Command.IsValid) Then
        strKey = Replace$(Command.Argument("Key"), "-", vbNullString)
        strKey = Replace$(strKey, Space$(1), vbNullString)
        
        If Config.SetKeyIgnoreLength Then
            If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
                strKey = vbNullString
            End If
        End If
        If (LenB(strKey) > 0) Then
            Config.CdKey = strKey
            Call Config.Save
            
            BotVars.CdKey = Config.CdKey
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
        Config.Username = Command.Argument("Username")
        Call Config.Save
        
        BotVars.Username = Config.Username
        Command.Respond StringFormat("New username set to {0}{1}{0}.", Chr$(34), BotVars.Username)
    Else
        Command.Respond "You must specify a username."
    End If
End Sub

Public Sub OnSetPass(Command As clsCommandObj)
    If (Command.IsValid) Then
        Config.Password = Command.Argument("Password")
        Call Config.Save
        
        BotVars.Password = Config.Password
        Command.Respond "New password set."
    Else
        Command.Respond "You must specify a password."
    End If
End Sub

Public Sub OnSetPMsg(Command As clsCommandObj)
    If (Command.IsValid) Then
        ProtectMsg = Command.Argument("Message")
        
        Config.ChannelProtectionMessage = ProtectMsg
        Call Config.Save
        
        Command.Respond "Channel protection message set."
    Else
        Command.Respond "You must specify a message."
    End If
End Sub

Public Sub OnSetServer(Command As clsCommandObj)
    If (Command.IsValid) Then
        Config.Server = Command.Argument("Server")
        Call Config.Save
        
        BotVars.Server = Config.Server
        Command.Respond StringFormat("New server set to {0}{1}{0}.", Chr$(34), BotVars.Server)
    Else
        Command.Respond "You must specify a server."
    End If
End Sub

Public Sub OnSetTrigger(Command As clsCommandObj)
    If (Command.IsValid) Then
        
        Config.Trigger = Command.Argument("Trigger")
        Call Config.Save
        
        BotVars.Trigger = Config.Trigger
        Command.Respond StringFormat("The new trigger is {0}{1}{0}.", Chr$(34), BotVars.Trigger)
    Else
        Command.Respond "You must specify a trigger."
    End If
End Sub

Public Sub OnWhisperCmds(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on":
            BotVars.WhisperCmds = True
            Config.WhisperResponses = True
            Call Config.Save
            
            Command.Respond "Command responses will now be whispered back."
        
        Case "off":
            BotVars.WhisperCmds = False
            Config.WhisperResponses = False
            Call Config.Save

            Command.Respond "Command responses will now be displayed publicly."
        
        Case Else:
            Command.Respond StringFormat("Command responses are currently {0}.", _
                IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly"))
    End Select
End Sub
