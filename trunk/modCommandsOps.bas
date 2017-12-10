Attribute VB_Name = "modCommandsOps"
Option Explicit
'This module will hold all commands that have to deal with holding operator over a channel

Private Const ERROR_NOT_OPS As String = "Error: This command requires channel operator status."
Public Enum CacheChanneListEnum
    enRetrieve = 0
    enAdd = 1
    enReset = 255
End Enum


Public Sub OnAddPhrase(Command As clsCommandObj)
    Dim sPhrase As String
    Dim i       As Integer
    Dim iFile   As Integer
    
    ' grab free file handle
    iFile = FreeFile
    If (Command.IsValid) Then
        sPhrase = Command.Argument("Phrase")
        
        For i = LBound(Phrases) To UBound(Phrases)
            If (StrComp(sPhrase, Phrases(i), vbTextCompare) = 0) Then
                Exit For
            End If
        Next i
        
        If (i > UBound(Phrases)) Then
            'Thats a lot of crap.. It check if the last item in Phrases is not just whitespace
            If (LenB(Trim$(Phrases(UBound(Phrases)))) > 0) Then
                ReDim Preserve Phrases(0 To UBound(Phrases) + 1)
            End If
            
            Phrases(UBound(Phrases)) = sPhrase
            
            Open GetFilePath(FILE_PHRASE_BANS) For Output As #iFile
                For i = LBound(Phrases) To UBound(Phrases)
                    If (LenB(Trim$(Phrases(i))) > 0) Then
                        Print #iFile, Phrases(i)
                    End If
                Next i
            Close #iFile
            
            Command.Respond StringFormat("Phraseban {0}{1}{0} added.", Chr$(34), sPhrase)
        Else
            Command.Respond "Error: That phrase is already banned."
        End If
    End If
End Sub

Public Sub OnBan(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            If (InStr(1, Command.Argument("Username"), "*", vbBinaryCompare) = 0) Then
                If (Command.IsLocal) Then
                    frmChat.AddQ "/ban " & Command.Args
                Else
                    Dim dbAccess As udtUserAccess
                    dbAccess = Database.GetUserAccess(Command.Username)
                    
                    Command.Respond Ban(Command.Args, dbAccess.Rank)
                End If
            Else
                Command.Respond WildCardBan(Command.Argument("Username"), Command.Argument("Message"), 1)
            End If
        Else
            Command.Respond ERROR_NOT_OPS
        End If
    Else
        Command.Respond "Error: You must specify a username to ban."
    End If
End Sub

Public Sub OnCAdd(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a game code to ban."
        Exit Sub
    End If

    Call Database.HandleAddCommand(Command, UCase(Command.Argument("game")), DB_TYPE_GAME, , "+B", , IIf(Len(Command.Argument("message")) = 0, "Client Ban", Command.Argument("message")))

End Sub

Public Sub OnCDel(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a game code to unban."
        Exit Sub
    End If
    
    Call Database.HandleAddCommand(Command, UCase(Command.Argument("game")), DB_TYPE_GAME, , "-B")

End Sub

Public Sub OnChPw(Command As clsCommandObj)
    Dim delay As Integer
    If (Command.IsValid) Then
        Select Case (LCase$(Command.Argument("SubCommand")))
            Case "on", "true", "enable", "enabled":
                If (LenB(Command.Argument("Value")) > 0) Then
                
                    BotVars.ChannelPassword = Command.Argument("Value")
                    If (BotVars.ChannelPasswordDelay <= 0) Then BotVars.ChannelPasswordDelay = 30
                    
                    Command.Respond StringFormat("Channel password protection enabled, delay set to {0}.", BotVars.ChannelPasswordDelay)
                Else
                    Command.Respond "Error: You must supply a password."
                End If
            
            Case "off", "false", "disable", "disabled":
                BotVars.ChannelPassword = vbNullString
                BotVars.ChannelPasswordDelay = 0
                Command.Respond "Channel password protection disabled."
            
            Case "delay":
                If (StrictIsNumeric(Command.Argument("Value"))) Then
                    delay = Val(Command.Argument("Value"))
                    If ((delay < 256) And (delay > 0)) Then
                        BotVars.ChannelPasswordDelay = CByte(delay)
                        Command.Respond StringFormat("Channel password delay set to {0}.", delay)
                    Else
                        Command.Respond "Error: Invalid channel delay."
                    End If
                Else
                    Command.Respond "Error: Invalid channel delay."
                End If
            
            Case Else:
            
                If ((LenB(BotVars.ChannelPassword) = 0) Or (BotVars.ChannelPasswordDelay = 0)) Then
                    Command.Respond "Channel password protection is currently disabled."
                Else
                    Command.Respond StringFormat("Channel password protection is currently enabled. Password [{0}], Delay [{1}].", _
                        BotVars.ChannelPassword, BotVars.ChannelPasswordDelay)
                End If
        End Select
    End If
End Sub

Public Sub OnClearBanList(Command As clsCommandObj)
    g_Channel.ClearBanlist
    Command.Respond "Banned user list cleared."
End Sub

Public Sub OnD2LevelBan(Command As clsCommandObj)
    Dim Level As Integer
    If (LenB(Command.Argument("Level")) > 0) Then
        Level = Command.Argument("Level")
        If (Level > 0) Then
            If (Level < 256) Then
                Command.Respond StringFormat("Banning Diablo II users under level {0}.", Level)
                BotVars.BanD2UnderLevel = CByte(Level)
            Else
                Command.Respond "Error: Invalid level specified."
            End If
        Else
            BotVars.BanD2UnderLevel = 0
            Command.Respond "Diablo II level bans disabled."
        End If
        
        If Config.LevelBanD2 <> BotVars.BanD2UnderLevel Then
            Config.LevelBanD2 = BotVars.BanD2UnderLevel
            Call Config.Save
        End If
    Else
        If (BotVars.BanD2UnderLevel = 0) Then
            Command.Respond "Currently not banning Diablo II users by level."
        Else
            Command.Respond StringFormat("Currently banning Diablo II users under level {0}.", BotVars.BanD2UnderLevel)
        End If
    End If
End Sub

Public Sub OnDelPhrase(Command As clsCommandObj)
    Dim iFile   As Integer
    Dim sPhrase As String
    Dim bFound  As Boolean
    Dim i       As Integer
    
    If (Command.IsValid) Then
        sPhrase = Command.Argument("Phrase")
        
        iFile = FreeFile
        
        Open GetFilePath(FILE_PHRASE_BANS) For Output As #iFile
            For i = LBound(Phrases) To UBound(Phrases)
                If (Not StrComp(Phrases(i), sPhrase, vbTextCompare) = 0) Then
                    Print #iFile, Phrases(i)
                Else
                    bFound = True
                End If
            Next i
        Close #iFile
        
        ReDim Phrases(0)
        Call frmChat.LoadArray(LOAD_PHRASES, Phrases())
        
        If (bFound) Then
            Command.Respond StringFormat("Phrase {0}{1}{0} deleted.", Chr$(34), sPhrase)
        Else
            Command.Respond "Error: That phrase is not banned."
        End If
    End If
End Sub

Public Sub OnDes(Command As clsCommandObj)
    If (g_Channel.Self.IsOperator) Then
        If (LenB(Command.Argument("Username")) > 0) Then
            Call frmChat.AddQ("/designate " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
            Command.Respond StringFormat("I have designated {0}.", Command.Argument("Username"))
        Else
            If (LenB(g_Channel.OperatorHeir) > 0) Then
                Command.Respond StringFormat("I have designated {0}.", g_Channel.OperatorHeir)
            Else
                Command.Respond "No user has been designated."
            End If
        End If
    Else
        Command.Respond ERROR_NOT_OPS
    End If
End Sub

Public Sub OnExile(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call OnShitAdd(Command)
        Call OnIPBan(Command)
    Else
        Command.Respond "Error: You must specify a username."
    End If
End Sub

Public Sub OnGiveUp(Command As clsCommandObj)
    ' This command will allow a user to designate a specified user using
    ' Battle.net's "designate" command, and will then make the bot resign
    ' its status as a channel moderator.  This command is useful if you are
    ' lazy and you just wish to designate someone as quickly as possible.
    
    Dim i         As Integer
    Dim opsCount  As Integer
    Dim sUsername As String
    Dim colUsers  As Collection
    
    If (Command.IsValid) Then
        sUsername = Command.Argument("Username")
        If (g_Channel.GetUserIndex(sUsername) > 0) Then
            If (g_Channel.Self.IsOperator) Then
                opsCount = GetOpsCount
                
                Set colUsers = New Collection
                
                If (StrComp(g_Channel.Name, "Clan " & g_Clan.Name, vbTextCompare) = 0) Then
                    If (g_Clan.Self.Rank >= clrankChieftain) Then
                        'Lets get a count of Shamans that are in the channel
                        For i = 1 To g_Clan.Shamans.Count
                            If (g_Channel.GetUserIndexEx(g_Clan.Shamans(i).Name) > 0) Then
                                colUsers.Add g_Clan.Shamans(i).Name
                            End If
                        Next i
                        
                        If (opsCount > (colUsers.Count + 1)) Then 'colUser.Count is present shamans, +1 for the bot being ops
                            Command.Respond "Error: There is currently a channel moderator present that cannot be removed from his or her position."
                            Exit Sub
                        End If
                        
                        'Lets demote the shamans
                        For i = 1 To colUsers.Count
                            g_Clan.Members(g_Clan.GetMemberIndexEx(colUsers.Item(i))).Demote
                        Next i
                        
                        opsCount = GetOpsCount
                    End If
                End If
                
                If (StrComp(Left$(g_Channel.Name, 3), "Op ", vbTextCompare) = 0) Then
                    If (opsCount >= 2) Then
                        Command.Respond "Error: There is currently a channel moderator present that cannot be removed from his or her position."
                        Exit Sub
                    End If
                ElseIf (StrComp(Left$(g_Channel.Name, 5), "Clan ", vbTextCompare) = 0) Then
                    If ((g_Clan.Self.Rank < 4) Or (Not StrComp(g_Channel.Name, "Clan " & g_Clan.Name, vbTextCompare) = 0)) Then
                        If (opsCount >= 2) Then
                            Command.Respond "Error: There is currently a channel moderator present that cannot be removed from his or her position."
                            Exit Sub
                        End If
                    End If
                End If
                
                Call bnetSend("/designate " & ReverseConvertUsernameGateway(sUsername))
                Call Pause(2)
                Call bnetSend("/resign")
                
                For i = 1 To colUsers.Count
                    g_Clan.Members(g_Clan.GetUserIndexEx(colUsers.Item(i))).Promote
                Next i
            Else
                Command.Respond ERROR_NOT_OPS
            End If
        Else
            Command.Respond "Error: The specified user is not present within the channel."
        End If
    End If
End Sub

Public Sub OnIdleBans(Command As clsCommandObj)
    If (Command.IsValid) Then
        Select Case (LCase$(Command.Argument("SubCommand")))
            Case "on", "true", "enable", "enabled":
                BotVars.IB_On = BTRUE
                
                If (StrictIsNumeric(Command.Argument("Value"))) Then
                    BotVars.IB_Wait = Val(Command.Argument("Value"))
                    Command.Respond StringFormat("IdleBans activated with a delay of {0}.", BotVars.IB_Wait)
                Else
                    BotVars.IB_Wait = 400
                    Command.Respond "IdleBans activated using the default delay of 400."
                End If
                
                Config.IdleBan = True
                Config.IdleBanDelay = BotVars.IB_Wait
                
            Case "off", "false", "disable", "disabled":
                BotVars.IB_On = BFALSE
                Config.IdleBan = False
                Command.Respond "IdleBans deactivated."
            
            Case "kick":
                Select Case LCase$(Command.Argument("Value"))
                    Case "on", "true", "enable", "enabled":
                        BotVars.IB_Kick = True
                    Case "off", "false", "disable", "disabled":
                        BotVars.IB_Kick = False
                    Case "toggle":
                        BotVars.IB_Kick = Not BotVars.IB_Kick
                    Case Else:
                        Command.Respond StringFormat("IdleBan action is set to {0}.", IIf(Config.IdleBanKick, "kick", "ban"))
                        Exit Sub
                End Select
                    
                If Config.IdleBanKick <> BotVars.IB_Kick Then
                    Config.IdleBanKick = BotVars.IB_Kick
                    Call Config.Save
                        
                    Command.Respond StringFormat("Idle users will now be {0}.", IIf(Config.IdleBanKick, "kicked", "banned"))
                    Exit Sub
                End If
                
            Case "delay":
                If (StrictIsNumeric(Command.Argument("Value"))) Then
                    BotVars.IB_Wait = CInt(Command.Argument("Value"))
                    Config.IdleBanDelay = BotVars.IB_Wait
                    Command.Respond StringFormat("IdleBan delay set to {0}.", BotVars.IB_Wait)
                Else
                    Command.Respond "Error: IdleBan delays require a numeric value."
                End If
                
            Case Else:
                If (BotVars.IB_On = BTRUE) Then
                    Command.Respond StringFormat("Idle {0} is currently enabled with a delay of {1} seconds.", _
                        IIf(BotVars.IB_Kick, "kicking", "banning"), BotVars.IB_Wait)
                Else
                    Command.Respond "IdleBan is currently disabled."
                End If
                Exit Sub
        End Select
        
        Call Config.Save
    End If
End Sub

Public Sub OnIPBan(Command As clsCommandObj)
    Dim dbAccess As udtUserAccess
    Dim dbTarget As udtUserAccess
    Dim sTarget  As String
    
    If (Command.IsValid) Then
        
        If (Not g_Channel.Self.IsOperator) Then
            Command.Respond "The bot does not currently have ops."
            Exit Sub
        End If

        If (Command.IsLocal) Then
            dbAccess = Database.GetConsoleAccess()
        Else
            dbAccess = Database.GetUserAccess(Command.Username)
        End If
        
        sTarget = StripInvalidNameChars(Command.Argument("Username"))
        
        If (LenB(sTarget) > 0) Then
            If (InStr(1, sTarget, Config.GatewayDelimiter) > 0) Then sTarget = StripRealm(sTarget)
            
            If (dbAccess.Rank < 101) Then
                If (GetSafelist(sTarget) Or GetSafelist(Command.Argument("Username"))) Then
                    Command.Respond "Error: That user is safelisted."
                    Exit Sub
                End If
            End If
            
            dbTarget = Database.GetUserAccess(Command.Argument("Username"))
            
            If ((dbTarget.Rank >= dbAccess.Rank) Or _
                ((InStr(1, dbTarget.Flags, "A", vbTextCompare) > 0) And (dbAccess.Rank < 101))) Then
                Command.Respond "Error: You do not have enough access to do that."
            Else
                Call frmChat.AddQ(StringFormat("/ban {0} {1}", Command.Argument("Username"), Command.Argument("Message")), , Command.Username)
                Call frmChat.AddQ(StringFormat("/squelch {0}", Command.Argument("Username")), , Command.Username)
                Command.Respond StringFormat("User {0}{1}{0} IPBanned.", Chr$(34), Command.Argument("Username"))
            End If
        End If
    End If
End Sub

Public Sub OnKick(Command As clsCommandObj)
    Dim dbAccess As udtUserAccess
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            
            If (InStr(1, Command.Argument("Username"), "*", vbTextCompare) > 0) Then
                Command.Respond WildCardBan(Command.Argument("Username"), Command.Argument("Message"), 0)
            Else
                If (Command.IsLocal) Then
                    frmChat.AddQ "/kick " & Command.Args
                Else
                    dbAccess = Database.GetUserAccess(Command.Username)
                    Command.Respond Ban(Command.Args, dbAccess.Rank, 1)
                End If
            End If
        Else
            Command.Respond ERROR_NOT_OPS
        End If
    End If
End Sub

Public Sub OnIPBans(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            BotVars.IPBans = True
            Command.Respond "IP banning activated."
            
            g_Channel.CheckUsers
            
        Case "off", "false", "disable", "disabled":
            BotVars.IPBans = False
            Command.Respond "IP banning deactivated."
        
        Case Else:
            Command.Respond StringFormat("IP banning is currently {0}activated.", _
                IIf(BotVars.IPBans, vbNullString, "de"))
    End Select
    
    If Config.IPBans <> BotVars.IPBans Then
        Config.IPBans = BotVars.IPBans
        Call Config.Save
    End If
End Sub

Public Sub OnKickOnYell(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            BotVars.KickOnYell = 1
            Command.Respond "Kick-on-yell enabled."
            
        Case "off", "false", "disable", "disabled":
            BotVars.KickOnYell = 0
            Command.Respond "Kick-on-yell disabled."
        
        Case Else:
            Command.Respond StringFormat("Kick-on-yell is currently {0}.", _
                IIf(BotVars.KickOnYell = 1, "enabled", "disabled"))
    End Select
    
    If Config.KickOnYell <> BotVars.KickOnYell Then
        Config.KickOnYell = BotVars.KickOnYell
        Call Config.Save
    End If
End Sub


Public Sub OnLevelBan(Command As clsCommandObj)
    Dim Level As Integer
    If (LenB(Command.Argument("Level")) > 0) Then
        Level = Command.Argument("Level")
        If (Level > 0) Then
            If (Level < 256) Then
                Command.Respond StringFormat("Banning Warcraft III users under level {0}.", Level)
                BotVars.BanUnderLevel = CByte(Level)
            Else
                Command.Respond "Error: Invalid level specified."
            End If
        Else
            BotVars.BanUnderLevel = 0
            Command.Respond "Levelbans disabled."
        End If
        
        If Config.LevelBanW3 <> BotVars.BanUnderLevel Then
            Config.LevelBanW3 = BotVars.BanUnderLevel
            Call Config.Save
        End If
    Else
        If (BotVars.BanUnderLevel = 0) Then
            Command.Respond "Currently not banning Warcraft III users by level."
        Else
            Command.Respond StringFormat("Currently banning Warcraft III users under level {0}.", BotVars.BanUnderLevel)
        End If
    End If
End Sub

Public Sub OnPeonBan(Command As clsCommandObj)
    ' This command will enable, disable, or check the status of, WarCraft III peon
    ' banning.  The "Peon" class is defined by Battle.net, and is currently the lowest
    ' ranking WarCraft III user classification, in which, users have less than twenty-five
    ' wins on record for any given race.
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            BotVars.BanPeons = True
            Command.Respond "Peon banning activated."
            
        Case "off", "false", "disable", "disabled":
            BotVars.BanPeons = False
            Command.Respond "Peon banning deactivated."
        
        Case Else:
            Command.Respond StringFormat("The bot is currently {0}banning peons.", _
                IIf(BotVars.BanPeons, vbNullString, "not "))
    End Select
    
    If Config.PeonBan <> BotVars.BanPeons Then
        Config.PeonBan = BotVars.BanPeons
        Call Config.Save
    End If
End Sub

Public Sub OnPhraseBans(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            PhraseBans = True
            Command.Respond "Phrasebans have been enabled."
            
        Case "off", "false", "disable", "disabled":
            PhraseBans = False
            Command.Respond "Phrasebans have been disabled."
        
        Case "kick":
            Select Case LCase$(Command.Argument("Value"))
                Case "on":
                    Config.PhraseKick = True
                Case "off":
                    Config.PhraseKick = False
                Case "toggle":
                    Config.PhraseKick = Not Config.PhraseKick
                Case Else:
                    Command.Respond StringFormat("Phraseban punishment is set to {0}.", IIf(Config.PhraseKick, "kick", "ban"))
                    Exit Sub
            End Select
            
            Command.Respond StringFormat("The punishment for saying a banned phrase is set to: {0}", _
                IIf(Config.PhraseKick, "kick", "ban"))
            
            Call Config.Save
            Exit Sub
        Case Else:
            Command.Respond StringFormat("Phrasebans are currently {0}.", _
                IIf(PhraseBans, "enabled", "disabled"))
    End Select
    
    If (Config.PhraseBans <> PhraseBans) Then
        Config.PhraseBans = PhraseBans
        Call Config.Save
    End If
End Sub

Public Sub OnPingBan(Command As clsCommandObj)
    Dim sValue As String
    sValue = Command.Argument("value")
    
    If Command.IsValid Then
        Select Case LCase$(sValue)
            Case "on", "true", "enable", "enabled":
                Config.PingBan = True
            Case "off", "false", "disable", "disabled":
                Config.PingBan = False
            Case "toggle":
                Config.PingBan = Not Config.PingBan
            Case Else:
                If IsNumeric(sValue) Then
                    Config.PingBanLevel = CLng(sValue)
                    Call Config.Save
                    
                    Command.Respond "PingBan level set to: " & Config.PingBanLevel
                    If Config.PingBan Then Call g_Channel.CheckUsers
                    Exit Sub
                Else
                    If Config.PingBan Then
                        Command.Respond StringFormat("PingBan is enabled. Users with a ping {0} {1} will be banned.", _
                            IIf(Config.PingBanLevel <= 0, "equal to", "greater than"), Config.PingBanLevel)
                    Else
                        Command.Respond "PingBan is currently disabled."
                    End If
                    Exit Sub
                End If
        End Select
        Call Config.Save
        
        Command.Respond StringFormat("PingBan has been {0}.", IIf(Config.PingBan, "enabled", "disabled"))
        
        If Config.PingBan Then Call g_Channel.CheckUsers
    End If
End Sub

Public Sub OnPlugBan(Command As clsCommandObj)
    ' This command will enable, disable, or check the status of, UDP plug bans.
    ' UDP plugs were traditionally used, in place of lag bars, to signifiy
    ' that a user was incapable of hosting (or possibly even joining) a game.
    ' However, as bot development became more popular, the emulation of such
    ' a connectivity issue became fairly common, and the UDP plug began to
    ' represent that a user was using a bot.  This feature allows for the
    ' banning of both, potential bots, and users unlikely to be capable of
    ' creating and/or joining games based on the UDP protocol.
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            If (BotVars.PlugBan) Then
                Command.Respond "PlugBan is already activated."
            Else
                BotVars.PlugBan = True
                Command.Respond "PlugBan activated."
                Call g_Channel.CheckUsers
            End If
            
        Case "off", "false", "disable", "disabled":
            If (Not BotVars.PlugBan) Then
                Command.Respond "PlugBan is already deactivated."
            Else
                BotVars.PlugBan = False
                Command.Respond "PlugBan deactivated."
            End If
        
        Case Else:
            Command.Respond StringFormat("The bot is currently {0}banning people with the UDP plug.", _
                IIf(BotVars.PlugBan, vbNullString, "not "))
    End Select
    
    If Config.UDPBan <> BotVars.PlugBan Then
        Config.UDPBan = BotVars.PlugBan
        Call Config.Save
    End If
    
    If Config.UDPBan Then
        Call g_Channel.CheckUsers
    End If
End Sub

Public Sub OnPOff(Command As clsCommandObj)
    PhraseBans = False
    If Config.PhraseBans Then
        Config.PhraseBans = False
        Call Config.Save
    End If
    Command.Respond "Phrasebans deactivated."
End Sub

Public Sub OnPOn(Command As clsCommandObj)
    PhraseBans = True
    If Not Config.PhraseBans Then
        Config.PhraseBans = True
        Call Config.Save
    End If
    Command.Respond "Phrasebans activated."
End Sub

Public Sub OnProtect(Command As clsCommandObj)
    Select Case (LCase$(Command.Argument("SubCommand")))
        Case "on", "true", "enable", "enabled":
            If (g_Channel.Self.IsOperator) Then
                Protect = True
                
                Call WildCardBan("*", ProtectMsg, 1)
                
                If (LenB(Command.Username) > 0) Then
                    Command.Respond StringFormat("Lockdown activated by {0}.", Command.Username)
                Else
                    Command.Respond "Lockdown activated."
                End If
            Else
                Command.Respond ERROR_NOT_OPS
            End If
        
        Case "off", "false", "disable", "disabled":
            If (Protect) Then
                Protect = False
                
                Command.Respond "Lockdown deactivated."
            Else
                Command.Respond "Lockdown was not enabled."
            End If
            
        Case Else:
            Command.Respond StringFormat("Lockdown is currently {0}active.", IIf(Protect, vbNullString, "not "))
    End Select
    
    If Protect <> Config.ChannelProtection Then
        Config.ChannelProtection = Protect
        Call Config.Save
    End If
End Sub

Public Sub OnPStatus(Command As clsCommandObj)
    Command.Respond StringFormat("Phrasebans are currently {0}.", _
        IIf(PhraseBans, "enabled", "disabled"))
End Sub

Public Sub OnQuietTime(Command As clsCommandObj)
    ' This command will enable, disable or check the status of, quiet time.
    ' Quiet time is a feature that will ban non-safelisted users from the
    ' channel when they speak publicly within the channel.  This is useful
    ' when a channel wishes to have a discussion while allowing public
    ' attendance, but disallowing public participation.
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on", "true", "enable", "enabled":
            BotVars.QuietTime = True
            Command.Respond "Quiet-time enabled."
            
        Case "off", "false", "disable", "disabled":
            BotVars.QuietTime = False
            Command.Respond "Quiet-time disabled."
        
        Case "kick":
            Select Case LCase$(Command.Argument("Value"))
                Case "on", "true", "enable", "enabled":
                    Config.QuietTimeKick = True
                Case "off", "false", "disable", "disabled":
                    Config.QuietTimeKick = False
                Case "toggle"
                    Config.QuietTimeKick = Not Config.QuietTimeKick
                Case Else
                    Command.Respond StringFormat("The QuietTime action is set to {0}", IIf(Config.QuietTimeKick, "kick", "ban"))
                    Exit Sub
            End Select
            
            Command.Respond StringFormat("QuietTime action set to {0}.", IIf(Config.QuietTimeKick, "kick", "ban"))
            Call Config.Save
        Case Else:
            Command.Respond StringFormat("QuietTime is currently {0}.", _
                IIf(BotVars.QuietTime, "enabled", "disabled"))
    End Select
    
    If Config.QuietTime <> BotVars.QuietTime Then
        Config.QuietTime = BotVars.QuietTime
        Call Config.Save
    End If
End Sub

Public Sub OnResign(Command As clsCommandObj)
    If (Not g_Channel.Self.IsOperator) Then Exit Sub
    Call frmChat.AddQ("/resign", enuPriority.SPECIAL_MESSAGE, Command.Username)
End Sub

Public Sub OnSafeAdd(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a username to add to the safelist."
        Exit Sub
    End If
    
    Dim bGroup  As Boolean      ' Is a safelist group configured?
    bGroup = False
    
    ' Is our safelist a group?
    If Len(Config.SafelistGroup) > 0 Then
        bGroup = True
    End If
    
    Call Database.HandleAddCommand(Command, Command.Argument("username"), DB_TYPE_USER, , IIf(bGroup, vbNullString, "+S"), IIf(bGroup, "+" & Config.SafelistGroup, vbNullString))
End Sub

Public Sub OnSafeDel(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must supply a name to remove from the safelist."
        Exit Sub
    End If
    
    Dim bGroup As Boolean
    bGroup = False
    
    If Len(Config.SafelistGroup) > 0 Then
        bGroup = True
    End If

    Call Database.HandleAddCommand(Command, Command.Argument("username"), DB_TYPE_USER, , IIf(bGroup, vbNullString, "-S"), IIf(bGroup, "-" & Config.SafelistGroup, vbNullString))
    
End Sub

Public Sub OnShitAdd(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a name to add to the shitlist."
        Exit Sub
    End If
 
    Dim bGroup  As Boolean
    bGroup = False
    
    ' Is shitlist a group?
    If Len(Config.ShitlistGroup) > 0 Then
        bGroup = True
    End If
    
    Call Database.HandleAddCommand(Command, Command.Argument("username"), DB_TYPE_USER, , IIf(bGroup, vbNullString, "+B"), IIf(bGroup, "+" & Config.ShitlistGroup, vbNullString), Command.Argument("message"))
End Sub

Public Sub OnShitDel(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must supply a name to remove from the shitlist."
        Exit Sub
    End If
    
    Dim bGroup  As Boolean
    bGroup = False

    If Len(Config.ShitlistGroup) > 0 Then
        bGroup = True
    End If
    
    Call Database.HandleAddCommand(Command, Command.Argument("username"), DB_TYPE_USER, , IIf(bGroup, vbNullString, "-B"), IIf(bGroup, "-" & Config.ShitlistGroup, vbNullString))
End Sub

Public Sub OnSweepBan(Command As clsCommandObj)
    ' This command will grab the listing of users in the specified channel
    ' using Battle.net's "who" command, and will then begin banning each
    ' user from the current channel using Battle.net's "ban" command.
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            ' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
            'Call CacheChannelList(vbNullString, 255, "ban ")
            Call CacheChannelList(enReset, "ban ")
            Call frmChat.AddQ("/who " & Command.Argument("Channel"), enuPriority.CHANNEL_MODERATION_MESSAGE, Command.Username, "request_receipt")
        Else
            Command.Respond ERROR_NOT_OPS
        End If
    End If
End Sub

Public Sub OnSweepIgnore(Command As clsCommandObj)
    ' This command will grab the listing of users in the specified channel
    ' using Battle.net's "who" command, and will then begin ignoring each
    ' user using Battle.net's "squelch" command.  This command is often used
    ' instead of "sweepban" to temporarily ban all users on a given ip address
    ' without actually immediately banning them from the channel.  This is
    ' useful if a user wishes to stay below Battle.net's limitations on bans,
    ' and still prevent a number of users from joining the channel for a
    ' temporary amount of time.
    If (Command.IsValid) Then
        ' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to be removed to script
        'Call CacheChannelList(vbNullString, 255, "squelch ")
        Call CacheChannelList(enReset, "squelch ")
        Call frmChat.AddQ("/who " & Command.Argument("Channel"), enuPriority.CHANNEL_MODERATION_MESSAGE, Command.Username, "request_receipt")
    End If
End Sub

Public Sub OnTagAdd(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a tag to ban."
        Exit Sub
    End If
    
    Dim sTag       As String       ' Tag being banned
    Dim bGroup     As Boolean      ' TRUE if using a tagban group.
    bGroup = False
    
    ' Tags that contain an asterisk (*) are treated as users, otherwise clans.
    
    sTag = Command.Argument("tag")
    
    ' Is taglist a group?
    If Len(Config.TagbanGroup) > 0 Then
        bGroup = True
    End If
    
    Call Database.HandleAddCommand(Command, sTag, IIf(InStr(1, sTag, "*", vbBinaryCompare) = 0, DB_TYPE_CLAN, DB_TYPE_USER), , IIf(bGroup, vbNullString, "+B"), IIf(bGroup, "+" & Config.TagbanGroup, vbNullString), Command.Argument("message"))
End Sub

Public Sub OnTagDel(Command As clsCommandObj)
    If Not Command.IsValid Then
        Command.Respond "You must specify a tag to unban."
        Exit Sub
    End If
    
    Dim sTag    As String
    Dim bGroup  As Boolean
    bGroup = False
    
    sTag = Command.Argument("tag")

    If Len(Config.TagbanGroup) > 0 Then
        bGroup = True
    End If
    
    Call Database.HandleAddCommand(Command, sTag, IIf(InStr(1, sTag, "*", vbBinaryCompare) = 0, DB_TYPE_CLAN, DB_TYPE_USER), , IIf(bGroup, vbNullString, "-B"), IIf(bGroup, "-" & Config.TagbanGroup, vbNullString))
End Sub

Public Sub OnUnBan(Command As clsCommandObj)
    Dim sTargetUser As String
    
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            sTargetUser = Command.Argument("Username")
            
            ' If no user was specified, unban the last banned user.
            If (LenB(sTargetUser) = 0) Then
                If (g_Channel.Banlist.Count > 0) Then
                    sTargetUser = g_Channel.Banlist(g_Channel.Banlist.Count).Name
                End If
            End If
            
            If (InStr(1, sTargetUser, "*", vbBinaryCompare) <> 0) Then
                Call WildCardBan(sTargetUser, vbNullString, 2)
            Else
                Call frmChat.AddQ("/unban " & sTargetUser, enuPriority.CHANNEL_MODERATION_MESSAGE, Command.Username)
            End If
        Else
            Command.Respond ERROR_NOT_OPS
        End If
    End If
End Sub

Public Sub OnUnExile(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call OnShitDel(Command)
        Call OnUnIPBan(Command)
    Else
        Command.Respond "Error: You must specify a user to unban."
    End If
End Sub

Public Sub OnUnIPBan(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            Call frmChat.AddQ("/unsquelch " & Command.Argument("Username"), , Command.Username)
            Call frmChat.AddQ("/unban " & Command.Argument("Username"), , Command.Username)
            Command.Respond StringFormat("User {0}{1}{0} has been Un-IPBanned.", Chr$(34), Command.Argument("Username"))
        Else
            Command.Respond ERROR_NOT_OPS
        End If
    End If
End Sub

Public Sub OnVoteBan(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (VoteDuration = -1) Then
            If (g_Channel.Self.IsOperator) Then
                Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, Command.Argument("Username"))
                VoteDuration = 30
                If (Command.IsLocal) Then
                    VoteInitiator = Database.GetConsoleAccess()
                Else
                    VoteInitiator = Database.GetUserAccess(Command.Username)
                End If
            
                Command.Respond StringFormat("30-second VoteBan vote started. Type YES to ban {0}, NO to acquit him/her.", Command.Argument("Username"))
            Else
                Command.Respond ERROR_NOT_OPS
            End If
        Else
            Command.Respond "A vote is currently in progress."
        End If
    Else
        Command.Respond "You must specify a user to kick."
    End If
End Sub

Public Sub OnVoteKick(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (VoteDuration = -1) Then
            If (g_Channel.Self.IsOperator) Then
                Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, Command.Argument("Username"))
                VoteDuration = 30
                If (Command.IsLocal) Then
                    VoteInitiator = Database.GetConsoleAccess()
                Else
                    VoteInitiator = Database.GetUserAccess(Command.Username)
                End If
            
                Command.Respond StringFormat("30-second VoteKick vote started. Type YES to kick {0}, NO to acquit him/her.", Command.Argument("Username"))
            Else
                Command.Respond ERROR_NOT_OPS
            End If
        Else
            Command.Respond "A vote is currently in progress."
        End If
    Else
        Command.Respond "You must specify a user to kick."
    End If
End Sub

Private Function GetOpsCount(Optional ByVal strIgnore As String = vbNullString) As Integer
    Dim i As Integer
    For i = 1 To g_Channel.Users.Count
        If (Not StrComp(g_Channel.Users(i).DisplayName, strIgnore, vbBinaryCompare) = 0) Then
            If (g_Channel.Users(i).IsOperator) Then GetOpsCount = GetOpsCount + 1
        End If
    Next i
End Function


'This is called in ChannelBan/Squelch to reset, and cache the users in a channel
'Then again in frmChat.tmrCache, to get a list of users and ban/squelch them
'The CS/CB commands AddQ "/Who Channel" and have a "request_receipt" message
'Then in Event_ServerInfo() it checks if it has a "request_receipt" if it does it call this function
'This is vary fucking ugly, we NEED to move this to a script whenever possible.
'Public Function CacheChannelList(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
Public Function CacheChannelList(ByVal eMode As CacheChanneListEnum, ByRef Data As String) As String
    Static colData             As New Collection
    Static sToken              As String
    Static bChannelListFollows As Boolean 'renamed this variable for clarity
    
    Dim sTemp As String
    
    'If(Mode = 255) Then
    If (eMode = enReset) Then
        Do While colData.Count > 0
            colData.Remove 1
        Loop
        sToken = Data
        bChannelListFollows = False
    End If
    
    If (Not InStr(1, LCase$(Data), "users in channel ", vbTextCompare) = 0) Then
        ' we weren't expecting a channel list, but we are now
        bChannelListFollows = True
    Else
        If (bChannelListFollows = True) Then ' if we're expecting a channel list, process it
            Select Case (eMode)
                Case enRetrieve ' RETRIEVE
                    ' Merge all the cache array items into one comma-delimited string
                    Do While colData.Count > 0
                        sTemp = StringFormat("{0}{1}, ", sTemp, colData.Item(1))
                        colData.Remove 1
                    Loop
                    Data = sToken
                    CacheChannelList = sTemp
                Case enAdd  ' ADD
                    colData.Add Data
            End Select
        End If
    End If
End Function

'This must be public becase The old OnAdd command uses it -.-
Public Function WildCardBan(ByVal sMatch As String, ByVal sBanMsg As String, ByVal Banning As Byte) As String
    'Values for Banning byte:
    '0 = Kick
    '1 = Ban
    '2 = Unban
    
    Dim i        As Integer
    Dim iSafe    As Integer
    Dim sCommand As String
    Dim sName    As String
    
    If (g_Channel.Self.IsOperator) Then
        If (g_Channel.Users.Count < 1) Then Exit Function
        
        If (LenB(sBanMsg) = 0) Then sBanMsg = sMatch
        
        sMatch = PrepareCheck(sMatch)
        
        Select Case (Banning)
            Case 1: sCommand = "/ban"
            Case 2: sCommand = "/unban"
            Case Else: sCommand = "/kick"
        End Select
        
        
        If (Not Banning = 2) Then
            ' Kicking or Banning
            For i = 1 To g_Channel.Users.Count
                With g_Channel.Users(i)
                    If (Not StrComp(.DisplayName, GetCurrentUsername, vbBinaryCompare) = 0) Then
                        sName = PrepareCheck(.DisplayName)
                        
                        If (sName Like sMatch) Then
                            If (GetSafelist(.DisplayName) = False) Then
                                If (Not .IsOperator) Then
                                    Call frmChat.AddQ(StringFormat("{0} {1} {2}", sCommand, .DisplayName, sBanMsg))
                                End If
                            Else
                                iSafe = (iSafe + 1)
                            End If
                        End If
                    End If
                End With
            Next i
            
            If (iSafe > 0) Then
                If (Not StrComp(sBanMsg, ProtectMsg, vbTextCompare) = 0) Then
                    WildCardBan = StringFormat("Encountered {0} safelisted user{1}.", iSafe, IIf(iSafe > 1, "s", vbNullString))
                End If
            End If
            
        Else
            For i = 1 To g_Channel.Banlist.Count
                With g_Channel.Banlist(i)
                    If ((.IsActive) And (LenB(.DisplayName) > 0)) Then
                        If (sMatch = "*") Then
                            Call frmChat.AddQ(StringFormat("{0} {1}", sCommand, .DisplayName))
                        Else
                            sName = PrepareCheck(.DisplayName)
                            If (sName Like sMatch) Then
                                Call frmChat.AddQ(StringFormat("{0} {1}", sCommand, .DisplayName))
                            End If
                        End If
                    End If
                End With
            Next i
        End If
    End If
End Function

