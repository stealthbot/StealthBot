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
            
            Open GetFilePath("PhraseBans.txt") For Output As #iFile
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
                    Dim dbAccess As udtGetAccessResponse
                    dbAccess = GetCumulativeAccess(Command.Username)
                    
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
    Dim sArgs As String
    If (Command.IsValid) Then
        If (LenB(Command.Argument("Message")) > 0) Then
            sArgs = "--banmsg " & Command.Argument("Message")
        Else
            sArgs = "--banmsg Client Ban"
        End If
        
        Command.Args = StringFormat("{0} +B --type GAME {1}", Command.Argument("Game"), sArgs)
        Call OnAdd(Command)
    End If
End Sub

Public Sub OnCDel(Command As clsCommandObj)
    If (Command.IsValid) Then
        Command.Args = Command.Argument("Game") & " -B --type GAME"
        Call OnAdd(Command)
    End If
End Sub

Public Sub OnChPw(Command As clsCommandObj)
    Dim delay As Integer
    If (Command.IsValid) Then
        Select Case (LCase$(Command.Argument("SubCommand")))
            Case "on":
                If (LenB(Command.Argument("Value")) > 0) Then
                
                    BotVars.ChannelPassword = Command.Argument("Value")
                    If (BotVars.ChannelPasswordDelay <= 0) Then BotVars.ChannelPasswordDelay = 30
                    
                    Command.Respond StringFormat("Channel password protection enabled, delay set to {0}.", BotVars.ChannelPasswordDelay)
                Else
                    Command.Respond "Error: You must supply a password."
                End If
            
            Case "off":
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
        Call WriteINI("Other", "BanD2UnderLevel", BotVars.BanUnderLevel)
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
        
        Open GetFilePath("PhraseBans.txt") For Output As #iFile
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
            Call frmChat.AddQ("/designate " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
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
    Dim colUsers  As New Collection
    
    If (Command.IsValid) Then
        sUsername = Command.Argument("Username")
        If (g_Channel.GetUserIndex(sUsername) > 0) Then
            If (g_Channel.Self.IsOperator) Then
                opsCount = GetOpsCount
                
                If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
                    If (g_Clan.Self.Rank >= 4) Then
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
                    If ((g_Clan.Self.Rank < 4) Or (Not StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0)) Then
                        If (opsCount >= 2) Then
                            Command.Respond "Error: There is currently a channel moderator present that cannot be removed from his or her position."
                            Exit Sub
                        End If
                    End If
                End If
                
                Call bnetSend("/designate " & reverseUsername(sUsername))
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
            Case "on":
                BotVars.IB_On = BTRUE
                
                If (StrictIsNumeric(Command.Argument("Value"))) Then
                    BotVars.IB_Wait = Val(Command.Argument("Value"))
                    Command.Respond StringFormat("IdleBans activated, with a delay of {0}.", BotVars.IB_Wait)
                Else
                    BotVars.IB_Wait = 400
                    Command.Respond "IdleBans activated, using the default delay of 400."
                End If
                
                Call WriteINI("Other", "IdleBans", "Y")
                Call WriteINI("Other", "IdleBanDelay", BotVars.IB_Wait)
                
            Case "off":
                BotVars.IB_On = BFALSE
                Call WriteINI("Other", "IdleBans", "N")
                Command.Respond "IdleBans deactivated."
            
            Case "kick":
                If (LenB(Command.Argument("Value")) > 0) Then
                    Select Case LCase$(Command.Argument("Value"))
                        Case "on":
                            BotVars.IB_Kick = True
                            Call WriteINI("Other", "KickIdle", "Y")
                            Command.Respond "Idle users will now be kicked instead of banned."
                        
                        Case "off":
                            BotVars.IB_Kick = False
                            Call WriteINI("Other", "KickIdle", "N")
                            Command.Respond "Idle users will now be banned instead of kicked."
                    
                    End Select
                End If
                
            Case "delay":
                If (StrictIsNumeric(Command.Argument("Value"))) Then
                    BotVars.IB_Wait = CInt(Command.Argument("Value"))
                    Call WriteINI("Other", "IdleBanDelay", BotVars.IB_Wait)
                    Command.Respond StringFormat("IdleBan delay set to {0}.", BotVars.IB_Wait)
                Else
                    Command.Respond "Error: IdleBan delays require a numeric value."
                End If
                
            Case Else:
                If (BotVars.IB_On = BTRUE) Then
                    Command.Respond StringFormat("Idle {0} is currently enabled with a delay of {1} seconds.", _
                        IIf(BotVars.IB_Kick, "kicking", "banning"), BotVars.IB_Wait)
                Else
                    Command.Respond "Idlebanning is currently disabled."
                End If
        End Select
    End If
End Sub

Public Sub OnIPBan(Command As clsCommandObj)
    Dim dbAccess As udtGetAccessResponse
    Dim dbTarget As udtGetAccessResponse
    Dim sTarget  As String
    
    If (Command.IsValid) Then
        
        If (Not g_Channel.Self.IsOperator) Then
            Command.Respond "The bot does not currently have ops."
            Exit Sub
        End If
        
        dbAccess = GetCumulativeAccess(Command.Username)
        If (Command.IsLocal) Then
            dbAccess.Rank = 201
            dbAccess.Flags = "A"
        End If
        
        sTarget = StripInvalidNameChars(Command.Argument("Username"))
        
        If (LenB(sTarget) > 0) Then
            If (InStr(1, sTarget, "@") > 0) Then sTarget = StripRealm(sTarget)
            
            If (dbAccess.Rank < 101) Then
                If (GetSafelist(sTarget) Or GetSafelist(Command.Argument("Username"))) Then
                    Command.Respond "Error: That user is safelisted."
                    Exit Sub
                End If
            End If
            
            dbTarget = GetCumulativeAccess(Command.Argument("Username"))
            
            If ((dbTarget.Rank >= dbAccess.Rank) Or _
                ((InStr(1, dbTarget.Flags, "A", vbTextCompare) > 0) And (dbAccess.Rank < 101))) Then
                Command.Respond "Error: You do not have enought access to do that."
            Else
                Call frmChat.AddQ(StringFormat("/ban {0} {1}", Command.Argument("Username"), Command.Argument("Message")), , Command.Username)
                Call frmChat.AddQ(StringFormat("/squelch {0}", Command.Argument("Username")), , Command.Username)
                Command.Respond StringFormat("User {0}{1}{0} IPBanned.", Chr$(34), Command.Argument("Username"))
            End If
        End If
    End If
End Sub

Public Sub OnKick(Command As clsCommandObj)
    Dim dbAccess As udtGetAccessResponse
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            
            If (InStr(1, Command.Argument("Username"), "*", vbTextCompare) > 0) Then
                Command.Respond WildCardBan(Command.Argument("Username"), Command.Argument("Message"), 0)
            Else
                If (Command.IsLocal) Then
                    frmChat.AddQ "/kick " & Command.Args
                Else
                    dbAccess = GetCumulativeAccess(Command.Username)
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
        Case "on":
            BotVars.IPBans = True
            Call WriteINI("Other", "IPBans", "Y")
            Command.Respond "IP banning activated."
            g_Channel.CheckUsers
            
        Case "off":
            BotVars.IPBans = False
            Call WriteINI("Other", "IPBans", "N")
            Command.Respond "IP banning deactivated."
        
        Case Else:
            Command.Respond StringFormat("IP banning is currently {0}activated.", _
                IIf(BotVars.IPBans, vbNullString, "de"))
    End Select
End Sub

Public Sub OnKickOnYell(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on":
            BotVars.KickOnYell = 1
            Call WriteINI("Other", "KickOnYell", "Y")
            Command.Respond "Kick-on-yell enabled."
            
        Case "off":
            BotVars.KickOnYell = 0
            Call WriteINI("Other", "KickOnYell", "N")
            Command.Respond "Kick-on-yell disabled."
        
        Case Else:
            Command.Respond StringFormat("Kick-on-yell is currently {0}.", _
                IIf(BotVars.KickOnYell = 1, "enabled", "disabled"))
    End Select
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
        Call WriteINI("Other", "BanUnderLevel", BotVars.BanUnderLevel)
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
        Case "on":
            BotVars.BanPeons = True
            Call WriteINI("Other", "PeonBans", "Y")
            Command.Respond "Peon banning activated."
            
        Case "off":
            BotVars.BanPeons = False
            Call WriteINI("Other", "PeonBans", "N")
            Command.Respond "Peon banning deactivated."
        
        Case Else:
            Command.Respond StringFormat("The bot is currently {0}banning peons.", _
                IIf(BotVars.BanPeons, vbNullString, "not "))
    End Select
End Sub

Public Sub OnPhraseBans(Command As clsCommandObj)
    Select Case LCase$(Command.Argument("SubCommand"))
        Case "on":
            PhraseBans = True
            Call WriteINI("Other", "PhraseBans", "Y")
            Command.Respond "Phrasebans activated."
            
        Case "off":
            PhraseBans = False
            Call WriteINI("Other", "PhraseBans", "N")
            Command.Respond "Phrasebans deactivated."
        
        Case Else:
            Command.Respond StringFormat("Phrasebans are currently {0}.", _
                IIf(PhraseBans, "enabled", "disabled"))
    End Select
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
        Case "on":
            If (BotVars.PlugBan) Then
                Command.Respond "PlugBan is already activated."
            Else
                BotVars.PlugBan = True
                Command.Respond "PlugBan activated."
                Call WriteINI("Other", "PlugBans", "Y")
                Call g_Channel.CheckUsers
            End If
            
        Case "off":
            If (Not BotVars.PlugBan) Then
                Command.Respond "PlugBan is already deactivated."
            Else
                BotVars.PlugBan = False
                Command.Respond "PlugBan deactivated."
                Call WriteINI("Other", "PlugBans", "N")
            End If
        
        Case Else:
            Command.Respond StringFormat("The bot is currently {0}banning people with the UDP plug.", _
                IIf(BotVars.PlugBan, vbNullString, "not "))
    End Select
End Sub

Public Sub OnPOff(Command As clsCommandObj)
    PhraseBans = False
    Call WriteINI("Other", "PhraseBans", "N")
    Command.Respond "Phrasebans deactivated."
End Sub

Public Sub OnPOn(Command As clsCommandObj)
    PhraseBans = True
    Call WriteINI("Other", "PhraseBans", "Y")
    Command.Respond "Phrasebans activated."
End Sub

Public Sub OnProtect(Command As clsCommandObj)
    Select Case (LCase$(Command.Argument("SubCommand")))
        Case "on":
            If (g_Channel.Self.IsOperator) Then
                Protect = True
                
                Call WildCardBan("*", ProtectMsg, 1)
                Call WriteINI("Main", "Protect", "Y")
                
                If (LenB(Command.Username) > 0) Then
                    Command.Respond StringFormat("Lockdown activated by {0}.", Command.Username)
                Else
                    Command.Respond "Lockdown activated."
                End If
            Else
                Command.Respond ERROR_NOT_OPS
            End If
        
        Case "off":
            If (Protect) Then
                Protect = False
                
                Call WriteINI("Main", "Protect", "N")
                Command.Respond "Lockdown deactivated."
            Else
                Command.Respond "Lockdown was not enabled."
            End If
            
        Case Else:
            Command.Respond StringFormat("Lockdown is currently {0}active.", IIf(Protect, vbNullString, "not "))
    End Select
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
        Case "on":
            BotVars.QuietTime = True
            Call WriteINI("Main", "QuietTime", "Y")
            Command.Respond "Quiet-time enabled."
            
        Case "off":
            BotVars.QuietTime = False
            Call WriteINI("Main", "QuietTime", "N")
            Command.Respond "Quiet-time disabled."
        
        Case Else:
            Command.Respond StringFormat("Quiet-time is currently {0}.", _
                IIf(BotVars.QuietTime, "enabled", "disabled"))
    End Select
End Sub

Public Sub OnResign(Command As clsCommandObj)
    If (Not g_Channel.Self.IsOperator) Then Exit Sub
    Call frmChat.AddQ("/resign", PRIORITY.SPECIAL_MESSAGE, Command.Username)
End Sub

Public Sub OnSafeAdd(Command As clsCommandObj)
    Dim sArgs As String
    If (Command.IsValid) Then
        If (LenB(BotVars.DefaultSafelistGroup) > 0) Then
            Dim dbAccess As udtGetAccessResponse
            dbAccess = GetAccess(BotVars.DefaultSafelistGroup, "GROUP")
            
            If (LenB(dbAccess.Username) > 0) Then sArgs = "--group " & BotVars.DefaultSafelistGroup
        End If
        
        If (LenB(sArgs) = 0) Then sArgs = "+S"
        sArgs = StringFormat("{0} {1} --type USER", Command.Argument("Username"), sArgs)
        
        Command.Args = sArgs
        Call OnAdd(Command)
    Else
        Command.Respond "Error: You must specify a username."
    End If
End Sub

Public Sub OnSafeDel(Command As clsCommandObj)
    If (Command.IsValid) Then
        Command.Args = Command.Argument("Username") & " -S --type USER"
        Call OnAdd(Command)
    Else
        Command.Respond "Error: You must supply a username."
    End If
End Sub

Public Sub OnShitAdd(Command As clsCommandObj)
    Dim sArgs    As String
    
    If (Command.IsValid) Then
        If (LenB(BotVars.DefaultShitlistGroup) > 0) Then
            Dim dbAccess As udtGetAccessResponse
            dbAccess = GetAccess(BotVars.DefaultShitlistGroup, "GROUP")
            If (LenB(dbAccess.Username) > 0) Then
                sArgs = "--group " & BotVars.DefaultShitlistGroup
            End If
        End If
        
        If (LenB(sArgs) = 0) Then sArgs = "+B"
        sArgs = StringFormat("{0} {1} --type USER", Command.Argument("Username"), sArgs)
        
        If (LenB(Command.Argument("Message")) > 0) Then
            sArgs = StringFormat("{0} --banmsg {1}", sArgs, Command.Argument("Message"))
        End If
        
        Command.Args = sArgs
        Call OnAdd(Command)
    End If
End Sub

Public Sub OnShitDel(Command As clsCommandObj)
    If (Command.IsValid) Then
        Command.Args = Command.Argument("Username") & " -B --type USER"
        Call OnAdd(Command)
    Else
        Command.Respond "Error: You must specify a username."
    End If
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
            Call frmChat.AddQ("/who " & Command.Argument("Channel"), PRIORITY.CHANNEL_MODERATION_MESSAGE, Command.Username, "request_receipt")
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
        Call frmChat.AddQ("/who " & Command.Argument("Channel"), PRIORITY.CHANNEL_MODERATION_MESSAGE, Command.Username, "request_receipt")
    End If
End Sub

Public Sub OnTagAdd(Command As clsCommandObj)
    Dim sArgs As String
    If (Command.IsValid) Then
        If (LenB(BotVars.DefaultTagbansGroup) > 0) Then
            Dim dbAccess As udtGetAccessResponse
            dbAccess = GetAccess(BotVars.DefaultTagbansGroup, "GROUP")
            
            If (LenB(dbAccess.Username) > 0) Then
                sArgs = "--group " & BotVars.DefaultTagbansGroup
            End If
        End If
        
        If (LenB(sArgs) = 0) Then sArgs = "+B"
        
        If (LenB(Command.Argument("Message")) > 0) Then
            sArgs = StringFormat("{0} --banmsg {1}", sArgs, Command.Argument("Message"))
        End If
        
        If (InStr(Command.Argument("Tag"), "*") = 0) Then
            sArgs = StringFormat("{0} {1} --type CLAN", Command.Argument("Tag"), sArgs)
        Else
            sArgs = StringFormat("{0} {1} --type USER", Command.Argument("Tag"), sArgs)
        End If
        
        Command.Args = sArgs
        Call OnAdd(Command)
    End If
End Sub

Public Sub OnTagDel(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (Not InStr(Command.Argument("Tag"), "*") = 0) Then
            Command.Args = Command.Argument("Tag") & " -B --Type USER"
            Call OnAdd(Command)
            Command.Args = Command.Argument("Tag") & " -B --Type CLAN"
            Call OnAdd(Command)
        Else
            Command.Args = StringFormat("*{0}* -B --Type USER", Command.Argument("Tag"))
            Call OnAdd(Command)
            Command.Args = StringFormat("*{0}* -B --Type CLAN", Command.Argument("Tag"))
            Call OnAdd(Command)
        End If
    End If
End Sub

Public Sub OnUnBan(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (g_Channel.Self.IsOperator) Then
            
            ' what the hell is a flood cap?
            If (bFlood) Then
                If (floodCap < 45) Then
                    floodCap = (floodCap + 15)
                    Call frmChat.AddQ("/unban " & Command.Argument("UserName"), , Command.Username)
                End If
            Else
                If (InStr(1, Command.Argument("Username"), "*", vbBinaryCompare) <> 0) Then
                    Call WildCardBan(Command.Argument("Username"), vbNullString, 2)
                Else
                    Call frmChat.AddQ("/unban " & Command.Argument("Username"), PRIORITY.CHANNEL_MODERATION_MESSAGE, Command.Username)
                End If
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
                    With VoteInitiator
                        .Rank = 201
                        .Flags = "A"
                        .Username = "(Console)"
                    End With
                Else
                    VoteInitiator = GetCumulativeAccess(Command.Username)
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
                    With VoteInitiator
                        .Rank = 201
                        .Flags = "A"
                        .Username = "(Console)"
                    End With
                Else
                    VoteInitiator = GetCumulativeAccess(Command.Username)
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

Private Function GetOpsCount(Optional strIgnore As String = vbNullString) As Integer
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
