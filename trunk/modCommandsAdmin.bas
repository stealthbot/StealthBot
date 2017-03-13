Attribute VB_Name = "modCommandsAdmin"
Option Explicit
'This module will hold all the commands that relate to Andmistering the bot, Changing settings
'Editing the database, etc..

Public Sub OnAdd(Command As clsCommandObj)
    Dim sNameToAdd      As String           ' The name of the user being added
    Dim aOptions()      As String           ' Broken out command options (designated by --)
    
    Dim iRank           As Integer          ' The rank being given to the user
    Dim sFlags          As String           ' Flags being assigned to the user
    Dim sType           As String           ' The type of entry being added
    Dim sGroup          As String           ' The current group being processed.
    Dim sBanMessage     As String           ' The ban message being assigned to this entry.
    
    Dim i               As Integer          ' counter
    Dim x               As Integer          ' position holder
    Dim aTemp()         As String           ' Array to temporarily hold split
    
    If ((Not Command.IsValid) Or LenB(Trim$(Command.Argument("username"))) = 0) Then
        Command.Respond "You must specify a user to add."
        Exit Sub
    End If
    
    ' If using D2 conventions, check for character names.
    '   If a character name is found and no account with the same name is found, assume the input
    '   was meant to target the account for the specified character.
    If (BotVars.UseD2Naming) Then
        Dim bIsCharacter    As Boolean      ' TRUE if the supplied name is a D2 character name.
        Dim sAccountName    As String       ' The account name of the targeted character.
        
        sNameToAdd = Command.Argument("username")
        If (Len(sNameToAdd) > 1) Then
            If (Left$(sNameToAdd, 1) = "*") Then
                If (InStr(2, sNameToAdd, "*") = 0 And InStr(2, sNameToAdd, "?") = 0) Then
                    ' format: *user
                    ' assume user means: user
                    Command.Args = Mid$(Command.Args, 2)
                End If
            End If
            
            If (InStr(sNameToAdd, "*") = 0 And InStr(sNameToAdd, "?") = 0) Then
                ' format: charname
                
                For i = 1 To g_Channel.Users.Count
                    With g_Channel.Users(i)
                        If (StrComp(.CharacterName, sNameToAdd, vbTextCompare) = 0) Then
                            ' the user provided is a character in the channel
                            bIsCharacter = True
                            sAccountName = .DisplayName
                        End If
                        If (StrComp(.DisplayName, sNameToAdd, vbTextCompare) = 0) Then
                            ' if the user provided is ALSO an accountname, assume account name
                            bIsCharacter = False
                            Exit For
                        End If
                    End With
                Next i
                If (bIsCharacter) Then
                    Command.Args = Replace$(Command.Args, sNameToAdd, sAccountName, 1, 1, vbBinaryCompare)
                End If
            End If
        End If
    Else
        sNameToAdd = Command.Argument("username")
    End If
    
    ' Set default values
    iRank = -1
    sFlags = vbNullString
    sType = DB_TYPE_USER     ' Default entry type: user
    sGroup = vbNullString
    sBanMessage = vbNullString

    ' Check for special parameters
    x = InStr(1, Command.Args, Space(1) & CMD_PARAM_PREFIX)
    If x > 0 Then
        ' Split up the parameters into an array.
        aOptions = Split(Mid(Command.Args, x), Space(1) & CMD_PARAM_PREFIX)
    Else
        ReDim aOptions(0)
    End If
        
    ' Process arguments
    If Len(Command.Argument("rank")) > 0 Then
        iRank = Int(Command.Argument("rank"))
    End If
    
    If Len(Command.Argument("attributes")) > 0 Then
        sFlags = Command.Argument("attributes")
        
        ' Make sure this isn't part of the special param system
        If Left(sFlags, Len(CMD_PARAM_PREFIX)) = CMD_PARAM_PREFIX Then
            sFlags = vbNullString
        End If
    End If

    ' Extra parameters
    If UBound(aOptions) > 0 Then

        For i = 0 To UBound(aOptions)
            If Len(aOptions(i)) > 0 Then
                aTemp = Split(aOptions(i), Space(1), 2)
            
                If UBound(aTemp) > 0 Then
                    ' Check parameter
                    Select Case UCase(aTemp(0))
                        Case "TYPE":                ' Defined type (default: user)
                            sType = UCase(aTemp(1))
                            
                        Case "GROUP"
                            sGroup = Split(aTemp(1), Space(1))(0)
                            
                        Case "BANMSG"
                            sBanMessage = aTemp(1)
                            
                    End Select
                Else
                    Command.Respond StringFormat("Invalid arguments. No value given for parameter: {0}", aTemp(0))
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    Call Database.HandleAddCommand(Command, sNameToAdd, sType, iRank, sFlags, sGroup, sBanMessage)
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

Public Sub OnRem(Command As clsCommandObj)
    Dim dbCallingUser   As udtUserAccess    ' Access of the requesting user
    Dim dbCurrentEntry  As clsDBEntryObj    ' Entry being modified
    Dim dbCurrentAccess As udtUserAccess    ' Cumulative access of the entry bheing modified.
    
    Dim aOptions()      As String           ' Broken out command options (designated by --)
    Dim sNameToRemove   As String           ' The name of the entry being removed.
    Dim sType           As String           ' The type of the entry being removed.
    
    Dim aTemp()         As String           ' Temporary array for processing
    Dim i               As Integer          ' counter
    Dim x               As Integer          ' position

    If (Not Command.IsValid) Then
        Command.Respond "You must specify a user to remove."
        Exit Sub
    End If
    
    If Command.IsLocal Then
        dbCallingUser = Database.GetConsoleAccess()
    Else
        dbCallingUser = Database.GetUserAccess(Command.Username)
    End If
    
    sNameToRemove = Command.Argument("username")
    sType = vbNullString
    
    ' Check for parameters
    x = InStr(1, Command.Args, CMD_PARAM_PREFIX)
    If x > 0 Then
        aOptions = Split(Mid(Command.Args, x), CMD_PARAM_PREFIX)
        
        For i = 1 To UBound(aOptions)
            aTemp = Split(aOptions(i), Space(1), 2)
            
            If UBound(aTemp) > 0 Then
                Select Case UCase(aTemp(0))
                    Case "TYPE"
                        sType = UCase(aTemp(1))
                    Case Else
                        Command.Respond "The specified parameter is not recognized."
                        Exit Sub
                End Select
            End If
        Next i
    End If
    
    ' Get the existing entry.
    Set dbCurrentEntry = Database.GetEntry(sNameToRemove, sType)
    If dbCurrentEntry Is Nothing Then
        Command.Respond "The specified entry could not be found."
        Exit Sub
    End If
    
    dbCurrentAccess = Database.GetEntryAccess(dbCurrentEntry)
    
    If ((Not Command.IsLocal) And (Not Database.CanUserModifyEntry(Command.Username, dbCurrentEntry))) Then
        Command.Respond "You do not have sufficient access to modify that entry."
        Exit Sub
    End If
    
    Call Database.RemoveEntry(dbCurrentEntry)
    Call Database.Save
                
    Command.Respond StringFormat("{0}{1}{0} has been removed from the database.", Chr(34), dbCurrentEntry.ToString())

End Sub

Public Sub OnSetBnlsServer(Command As clsCommandObj)
    If (Command.IsValid) Then
        Config.BNLSServer = Command.Argument("Server")
        Call Config.Save
        
        BotVars.BNLSServer = Config.BNLSServer
        Command.Respond StringFormat("New BNLS server set to {0}{1}{0}.", Chr$(34), BotVars.BNLSServer)
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
        
        If Config.IgnoreCDKeyLength Then
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
    
    Channel = Command.Argument("Channel")
    Config.HomeChannel = Channel
    Call Config.Save
    
    BotVars.HomeChannel = Config.HomeChannel
    PrepareHomeChannelMenu
    If LenB(Channel) = 0 Then
        Command.Respond "Home channel set to server default."
    Else
        Command.Respond StringFormat("New home channel set to {0}{1}{0}.", Chr$(34), Config.HomeChannel)
    End If
End Sub

Public Sub OnSetKey(Command As clsCommandObj)
    Dim strKey As String
    If (Command.IsValid) Then
        strKey = Replace$(Command.Argument("Key"), "-", vbNullString)
        strKey = Replace$(strKey, Space$(1), vbNullString)
        
        If Config.IgnoreCDKeyLength Then
            If ((Not Len(strKey) = 13) And (Not Len(strKey) = 16) And (Not Len(strKey) = 26)) Then
                strKey = vbNullString
            End If
        End If
        If (LenB(strKey) > 0) Then
            Config.CDKey = strKey
            Call Config.Save
            
            BotVars.CDKey = Config.CDKey
            Command.Respond "New CD key set."
        Else
            Command.Respond "The CD key you specified was invalid."
        End If
    Else
        Command.Respond "You must specify a CD key."
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
            
            Command.Respond "Command responses will now be whispered back."
        
        Case "off":
            BotVars.WhisperCmds = False

            Command.Respond "Command responses will now be displayed publicly."
        
        Case Else:
            Command.Respond StringFormat("Command responses are currently {0}.", _
                IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly"))
    End Select
    
    If Config.WhisperCommands <> BotVars.WhisperCmds Then
        Config.WhisperCommands = BotVars.WhisperCmds
        Call Config.Save
    End If
End Sub
