Attribute VB_Name = "modCommandCode"
' modCommandCode.bas
' Copyright (C) 2002 - 2007 Stealth & Eric Evans

' This module checks a message for the presence of a command identifer and, if
' found, the message is then sent to the secondary processor, which then finally
' sends the message to an individual command handler specific to each command.
' The secondary processor however first ensures that the command is valid and that
' the user has sufficient access to execute the command.  The ProcessCommand()
' function contains a default error handler that will catch any errors thrown within
' that function, or any of the functions that it calls, in case of an unhandled
' exception.

Option Explicit
Private m_waswhispered  As Boolean ' ...
Private m_DisplayOutput As Boolean ' ...

Public flood    As String ' ...?
Public floodCap As Byte   ' ...?

' prepares commands for processing, and calls helper functions associated with processing
Public Function ProcessCommand(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As _
        Boolean = False, Optional ByVal WasWhispered As Boolean = False, Optional DisplayOutput As Boolean = _
                True) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim commands         As Collection
    Dim Command          As clsCommandObj
    Dim dbAccess         As udtGetAccessResponse
    Dim I                As Integer
    Dim Count            As Integer
    Dim bln              As Boolean
    Dim command_return() As String
    Dim outbuf           As String
    Dim execCommand      As Boolean
    
    ' ...
    ReDim Preserve command_return(0)
    
    ' store file scope copy of whisper status
    m_waswhispered = WasWhispered
    
    ' replace message variables
    Message = Replace(Message, "%me", IIf(IsLocal, GetCurrentUsername, Username), 1, -1, vbTextCompare)
    
    If ((IsLocal) And (Left$(Message, 3) = "///")) Then
        AddQ Mid$(Message, 3)
        Exit Function
    End If

    ' 08/17/2009 - 52 - using static class method
    Set commands = clsCommandObj.IsCommand(Message, IIf(IsLocal, modGlobals.CurrentUsername, Username), Chr$(0))

    For Each Command In commands
        m_DisplayOutput = Command.PublicOutput
        
        If (Command.HasAccess) Then
            Command.WasWhispered = WasWhispered
            If (IsLocal) Then
                With dbAccess
                    .Rank = 201
                    .Flags = "A"
                End With
            Else
                dbAccess = GetCumulativeAccess(Username)
            End If
            
            If (LenB(Command.docs.Owner) = 0) Then 'Is it a built in command?
                If (Not executeCommand(Command.Username, dbAccess, Command.Name, Command.Args, Command.IsLocal, command_return)) Then
                    Call DispatchCommand(Command)
                    Call RunInSingle(Nothing, "Event_Command", Command)
                    Command.SendResponse
                End If
            Else
                Call RunInSingle(modScripting.GetModuleByName(Command.docs.Owner), "Event_Command", Command)
                Command.SendResponse
            End If
                    
            If (DisplayOutput) Then
            
                'Ignore code as bot is closing
                If BotIsClosing Then Exit Function
            
                If (command_return(0) <> vbNullString) Then
                    For I = LBound(command_return) To UBound(command_return)
                        If (IsLocal) Then
                            If (Command.PublicOutput) Then
                                AddQ command_return(I), PRIORITY.CONSOLE_MESSAGE
                            Else
                                frmChat.AddChat RTBColors.ConsoleText, command_return(I)
                            End If
                        Else
                            If ((BotVars.WhisperCmds) Or (WasWhispered)) Then
                                AddQ "/w " & Username & Space$(1) & command_return(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
                            Else
                                AddQ command_return(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
                            End If
                        End If
                    Next I
                End If
            End If
        End If
        
        execCommand = False
    Next
        
    If (IsLocal) Then
        If ((bln = False) And (commands.Count = 0)) Then AddQ Message
    End If
    
    'Unload memory - FrOzeN
    Set Command = Nothing
    Set commands = Nothing
    
    Exit Function
    
' default (if all else fails) error handler to keep erroneous
' commands and/or input formats from killing me
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: #" & Err.Number & ": " & Err.description & _
        " in modCommandCode.ProcessCommand().")

    Set Command = Nothing
    ProcessCommand = False
End Function

'This is the replacement for ExecuteCommand, Uses the new clsCommandObj, Should be cleaner.
Public Function DispatchCommand(Command As clsCommandObj)
    DispatchCommand = True
    Select Case LCase(Command.Name)
        'Bot information commands
        Case "about":         Call modCommandsInfo.OnAbout(Command)
        Case "accountinfo":   Call modCommandsInfo.OnAccountInfo(Command)
        Case "allseen":       Call modCommandsInfo.OnAllSeen(Command)
        Case "detail":        Call modCommandsInfo.OnDetail(Command)
        Case "find":          Call modCommandsInfo.OnFind(Command)
        Case "findattr":      Call modCommandsInfo.OnFindAttr(Command)
        Case "findgrp":       Call modCommandsInfo.OnFindGrp(Command)
        Case "help":          Call modCommandsInfo.OnHelp(Command)
        Case "helpattr":      Call modCommandsInfo.OnHelpAttr(Command)
        Case "helprank":      Call modCommandsInfo.OnHelpRank(Command)
        Case "info":          Call modCommandsInfo.OnInfo(Command)
        Case "initperf":      Call modCommandsInfo.OnInitPerf(Command)
        Case "lastwhisper":   Call modCommandsInfo.OnLastWhisper(Command)
        Case "localip":       Call modCommandsInfo.OnLocalIp(Command)
        Case "owner":         Call modCommandsInfo.OnOwner(Command)
        Case "ping":          Call modCommandsInfo.OnPing(Command)
        Case "pingme":        Call modCommandsInfo.OnPingMe(Command)
        Case "profile":       Call modCommandsInfo.OnProfile(Command)
        Case "scriptdetail":  Call modCommandsInfo.OnScriptDetail(Command)
        Case "scripts":       Call modCommandsInfo.OnScripts(Command)
        Case "server":        Call modCommandsInfo.OnServer(Command)
        Case "time":          Call modCommandsInfo.OnTime(Command)
        Case "trigger":       Call modCommandsInfo.OnTrigger(Command)
        Case "uptime":        Call modCommandsInfo.OnUptime(Command)
        Case "where":         Call modCommandsInfo.OnWhere(Command)
        Case "whoami":        Call modCommandsInfo.OnWhoAmI(Command)
        Case "whois":         Call modCommandsInfo.OnWhoIs(Command)
        
        'Clan related commands
        Case "clan":          Call modCommandsClan.OnClan(Command)
        Case "demote":        Call modCommandsClan.OnDemote(Command)
        Case "disbandclan":   Call modCommandsClan.OnDisbandClan(Command)
        Case "invite":        Call modCommandsClan.OnInvite(Command)
        Case "makechieftain": Call modCommandsClan.OnMakeChieftain(Command)
        Case "motd":          Call modCommandsClan.OnMOTD(Command)
        Case "promote":       Call modCommandsClan.OnPromote(Command)
        Case "setmotd":       Call modCommandsClan.OnSetMOTD(Command)
        
        'Media Player comands
        Case "allowmp3":      Call modCommandsMP3.OnAllowMp3(Command)
        Case "fos":           Call modCommandsMP3.OnFOS(Command)
        Case "loadwinamp":    Call modCommandsMP3.OnLoadWinamp(Command)
        Case "mp3":           Call modCommandsMP3.OnMP3(Command)
        Case "next":          Call modCommandsMP3.OnNext(Command)
        Case "pause":         Call modCommandsMP3.OnPause(Command)
        Case "play":          Call modCommandsMP3.OnPlay(Command)
        Case "previous":      Call modCommandsMP3.OnPrevious(Command)
        Case "repeat":        Call modCommandsMP3.OnRepeat(Command)
        Case "setvol":        Call modCommandsMP3.OnSetVol(Command)
        Case "shuffle":       Call modCommandsMP3.OnShuffle(Command)
        Case "stop":          Call modCommandsMP3.OnStop(Command)
        Case "useitunes":     Call modCommandsMP3.OnUseiTunes(Command)
        Case "usewinamp":     Call modCommandsMP3.OnUseWinamp(Command)
        
        'Chat Commands
        Case "away":          Call modCommandsChat.OnAway(Command)
        Case "back":          Call modCommandsChat.OnBack(Command)
        Case "block":         Call modCommandsChat.OnBlock(Command)
        Case "connect":       Call modCommandsChat.OnConnect(Command)
        Case "cq":            Call modCommandsChat.OnCQ(Command)
        Case "disconnect":    Call modCommandsChat.OnDisconnect(Command)
        Case "expand":        Call modCommandsChat.OnExpand(Command)
        Case "fadd":          Call modCommandsChat.OnFAdd(Command)
        Case "filter":        Call modCommandsChat.OnFilter(Command)
        Case "forcejoin":     Call modCommandsChat.OnForceJoin(Command)
        Case "frem":          Call modCommandsChat.OnFRem(Command)
        Case "home":          Call modCommandsChat.OnHome(Command)
        Case "ignore":        Call modCommandsChat.OnIgnore(Command)
        Case "igpriv":        Call modCommandsChat.OnIgPriv(Command)
        Case "join":          Call modCommandsChat.OnJoin(Command)
        Case "unblock":       Call modCommandsChat.OnUnBlock(Command)
        Case "unfilter":      Call modCommandsChat.OnUnFilter(Command)
        Case "unignore":      Call modCommandsChat.OnUnIgnore(Command)
        Case "unigpriv":      Call modCommandsChat.OnUnIgPriv(Command)
        Case "quickrejoin":   Call modCommandsChat.OnQuickRejoin(Command)
        Case "reconnect":     Call modCommandsChat.OnReconnect(Command)
        Case "rejoin":        Call modCommandsChat.OnReJoin(Command)
        Case "say":           Call modCommandsChat.OnSay(Command)
        Case "scq":           Call modCommandsChat.OnSCQ(Command)
        Case "shout":         Call modCommandsChat.OnShout(Command)
        Case "watch":         Call modCommandsChat.OnWatch(Command)
        Case "watchoff":      Call modCommandsChat.OnWatchOff(Command)
        
        'Admin Commands
        Case "add":           Call modCommandsAdmin.OnAdd(Command)
        Case "clear":         Call modCommandsAdmin.OnClear(Command)
        Case "disable":       Call modCommandsAdmin.OnDisable(Command)
        Case "dump":          Call modCommandsAdmin.OnDump(Command)
        Case "enable":        Call modCommandsAdmin.OnEnable(Command)
        Case "locktext":      Call modCommandsAdmin.OnLockText(Command)
        Case "quit":          Call modCommandsAdmin.OnQuit(Command)
        Case "setexpkey":     Call modCommandsAdmin.OnSetExpKey(Command)
        Case "sethome":       Call modCommandsAdmin.OnSetHome(Command)
        Case "setkey":        Call modCommandsAdmin.OnSetKey(Command)
        Case "setname":       Call modCommandsAdmin.OnSetName(Command)
        Case "setpass":       Call modCommandsAdmin.OnSetPass(Command)
        Case "setserver":     Call modCommandsAdmin.OnSetServer(Command)
        Case "settrigger":    Call modCommandsAdmin.OnSetTrigger(Command)
        Case "whispercmds":   Call modCommandsAdmin.OnWhisperCmds(Command)
        
        'Ops Commands
        Case "cancel":        Call modCommandsOps.OnCancel(Command)
        Case "chpw":          Call modCommandsOps.OnChPw(Command)
        Case "clearbanlist":  Call modCommandsOps.OnClearBanList(Command)
        Case "d2levelban":    Call modCommandsOps.OnD2LevelBan(Command)
        Case "des":           Call modCommandsOps.OnDes(Command)
        Case "exile":         Call modCommandsOps.OnExile(Command)
        Case "giveup":        Call modCommandsOps.OnGiveUp(Command)
        Case "ipban":         Call modCommandsOps.OnIPBan(Command)
        Case "ipbans":        Call modCommandsOps.OnIPBans(Command)
        Case "kickonyell":    Call modCommandsOps.OnKickOnYell(Command)
        Case "levelban":      Call modCommandsOps.OnLevelBan(Command)
        Case "peonban":       Call modCommandsOps.OnPeonBan(Command)
        Case "phrasebans":    Call modCommandsOps.OnPhraseBans(Command)
        Case "plugban":       Call modCommandsOps.OnPlugBan(Command)
        Case "poff":          Call modCommandsOps.OnPOff(Command)
        Case "pon":           Call modCommandsOps.OnPOn(Command)
        Case "pstatus":       Call modCommandsOps.OnPStatus(Command)
        Case "quiettime":     Call modCommandsOps.OnQuietTime(Command)
        Case "resign":        Call modCommandsOps.OnResign(Command)
        Case "shitadd":       Call modCommandsOps.OnShitAdd(Command)
        Case "shitdel":       Call modCommandsOps.OnShitDel(Command)
        Case "sweepban":      Call modCommandsOps.OnSweepBan(Command)
        Case "sweepignore":   Call modCommandsOps.OnSweepIgnore(Command)
        Case "tally":         Call modCommandsOps.OnTally(Command)
        Case "unexile":       Call modCommandsOps.OnUnExile(Command)
        Case "unipban":       Call modCommandsOps.OnUnIPBan(Command)
        
        'Misc Commands
        Case "bmail":         Call modCommandsMisc.OnBMail(Command)
        Case "checkmail":     Call modCommandsMisc.OnCheckMail(Command)
        Case "exec":          Call modCommandsMisc.OnExec(Command)
        Case "flip":          Call modCommandsMisc.OnFlip(Command)
        Case "math":          Call modCommandsMisc.OnMath(Command)
        Case "mmail":         Call modCommandsMisc.OnMMail(Command)
        Case "roll":          Call modCommandsMisc.OnRoll(Command)
        Case "readfile":      Call modCommandsMisc.OnReadFile(Command)
        
        Case Else: DispatchCommand = False
    End Select
End Function

Public Function executeCommand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal cmdName As String, ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    LogCommand IIf(InBot, vbNullString, Username), cmdName & Space(1) & msgData
    executeCommand = True
    Select Case (cmdName)
        'Case "add":           Call OnAddOld(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idlebans":      Call OnIdleBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clientbans":    Call OnClientBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cadd":          Call OnCAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cdel":          Call OnCDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banned":        Call OnBanned(Username, dbAccess, msgData, InBot, cmdRet())
        Case "protect":       Call OnProtect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "rem":           Call OnRem(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletime":      Call OnIdleTime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idle":          Call OnIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safedel":       Call OnSafeDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagdel":        Call OnTagDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setidle":       Call OnSetIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletype":      Call OnIdleType(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setpmsg":       Call OnSetPMsg(Username, dbAccess, msgData, InBot, cmdRet())
        Case "phrases":       Call OnPhrases(Username, dbAccess, msgData, InBot, cmdRet())
        Case "addphrase":     Call OnAddPhrase(Username, dbAccess, msgData, InBot, cmdRet())
        Case "delphrase":     Call OnDelPhrase(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagadd":        Call OnTagAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safelist":      Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safeadd":       Call OnSafeAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safecheck":     Call OnSafeCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitlist":      Call OnShitList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagbans":       Call OnTagBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "bancount":      Call OnBanCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banlistcount":  Call OnBanListCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagcheck":      Call OnTagCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "slcheck":       Call OnSLCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "greet":         Call OnGreet(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ban":           Call OnBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unban":         Call OnUnban(Username, dbAccess, msgData, InBot, cmdRet())
        Case "kick":          Call OnKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "voteban":       Call OnVoteBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "votekick":      Call OnVoteKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "vote":          Call OnVote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "addquote":      Call OnAddQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quote":         Call OnQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "inbox":         Call OnInbox(Username, dbAccess, msgData, InBot, cmdRet())
        Case Else:            executeCommand = False
    End Select
End Function

' handle idlebans command
Private Function OnIdleBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String ' ...
    Dim tmpbuf     As String ' temporary output buffer
    Dim subcmd     As String ' ...
    Dim Index      As Long   ' ...
    Dim tmpData    As String ' ...
    
    tmpData = msgData
    
    If (Len(tmpData) > 0) Then
        Index = InStr(1, tmpData, Space$(1), vbBinaryCompare)
    
        If (Index <> 0) Then
            subcmd = Mid$(tmpData, 1, Index - 1)
        Else
            subcmd = tmpData
        End If
        
        subcmd = LCase$(subcmd)
        
        If (Index) Then
            tmpData = Mid$(msgData, Index + 1)
        End If
    
        Select Case (subcmd)
            Case "on"
                BotVars.IB_On = BTRUE
                
                If (Len(tmpData) > 0) Then
                    If (StrictIsNumeric(tmpData)) Then
                        BotVars.IB_Wait = Val(tmpData)
                    End If
                End If
                
                If (BotVars.IB_Wait > 0) Then
                    tmpbuf = "IdleBans activated, with a delay of " & BotVars.IB_Wait & "."
                    
                    Call WriteINI("Other", "IdleBans", "Y")
                    Call WriteINI("Other", "IdleBanDelay", BotVars.IB_Wait)
                Else
                    BotVars.IB_Wait = 400
                    
                    tmpbuf = "IdleBans activated, using the default delay of 400."
                    
                    Call WriteINI("Other", "IdleBanDelay", "400")
                    Call WriteINI("Other", "IdleBans", "Y")
                End If
                
            Case "off"
                BotVars.IB_On = BFALSE
                
                tmpbuf = "IdleBans deactivated."
                
                Call WriteINI("Other", "IdleBans", "N")
                
            Case "kick"
                If (Len(tmpData) > 0) Then
                    Select Case (LCase$(tmpData))
                        Case "on"
                            tmpbuf = "Idle users will now be kicked instead of banned."
                            
                            Call WriteINI("Other", "KickIdle", "Y")
                            
                            BotVars.IB_Kick = True
                            
                        Case "off"
                            tmpbuf = "Idle users will now be banned instead of kicked."
                            
                            Call WriteINI("Other", "KickIdle", "N")
                            
                            BotVars.IB_Kick = False
                            
                        Case Else
                            tmpbuf = "Error: Unknown idle kick setting."
                    End Select
                Else
                    tmpbuf = "Error: Too few arguments."
                End If
        
            Case "wait", "delay"
                If (StrictIsNumeric(tmpData)) Then
                    BotVars.IB_Wait = CInt(tmpData)
                    
                    tmpbuf = "IdleBan delay set to " & BotVars.IB_Wait & "."
                    
                    Call WriteINI("Other", "IdleBanDelay", CInt(tmpData))
                Else
                    tmpbuf = "Error: IdleBan delays require a numeric value."
                End If
                
            Case "status"
                If (BotVars.IB_On = BTRUE) Then
                    tmpbuf = IIf(BotVars.IB_Kick, "Kicking", "Banning") & _
                        " users who are idle for " & BotVars.IB_Wait & "+ seconds."
                Else
                    tmpbuf = "IdleBans are disabled."
                End If
                
            Case Else
                tmpbuf = "Error: Invalid command."
        End Select
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIdleBans

' handle clientbans command
Private Function OnClientBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase2(tmpbuf(), , , , "GAME", , , "B")
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnClientBans

' handle cadd command
Private Function OnCAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    Dim Index    As Integer

    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' ...
    Index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (Index) Then
        Dim user As String ' ...
        
        ' ...
        user = Mid$(msgData, 1, Index - 1)
        
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            tmpbuf(0) = "Error: The specified game name is invalid."
        Else
            Dim bmsg As String ' ...
            
            ' ...
            bmsg = Mid$(msgData, Index + 1)
        
            ' ...
            Call OnAddOld(Username, dbAccess, user & " +B --type GAME --banmsg " & bmsg, True, tmpbuf())
        End If
    Else
        ' ...
        Call OnAddOld(Username, dbAccess, msgData & " +B --type GAME", True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnCAdd

' handle cdel command
Private Function OnCDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpbuf(0) = "Error: The specified game name is invalid."
    Else
        ' remove user from shitlist using "add" command
        Call OnAddOld(Username, dbAccess, msgData & " -B --type GAME", True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnCDel

' handle banned command
Private Function OnBanned(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will display a listing of all of the users that have been
    ' banned from the channel since the time of having joined the channel.
    
    Dim tmpbuf()  As String ' temporary output buffer
    Dim tmpCount  As Integer
    Dim BanCount  As Integer
    Dim I         As Integer
    Dim j         As Integer ' ...
    Dim userCount As Integer ' ...
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    If (g_Channel.Banlist.Count = 0) Then
        cmdRet(0) = "There are presently no users on the bot's internal banlist."
        Exit Function
    End If

    tmpbuf(tmpCount) = "User(s) banned: "
    
    For I = 1 To g_Channel.Banlist.Count
        If (g_Channel.Banlist(I).IsDuplicateBan = False) Then
            For j = 1 To g_Channel.Banlist.Count
                If (StrComp(g_Channel.Banlist(j).DisplayName, g_Channel.Banlist(I).DisplayName, vbTextCompare) = 0) Then
                    userCount = (userCount + 1)
                End If
            Next j
            
            tmpbuf(tmpCount) = StringFormatA("{0}, {1}", tmpbuf(tmpCount), g_Channel.Banlist(I).DisplayName)
                    
            If (userCount > 1) Then
                tmpbuf(tmpCount) = StringFormatA("{0} ({1})", tmpbuf(tmpCount), userCount)
            End If
                    
            If ((Len(tmpbuf(tmpCount)) > 90) And (I <> g_Channel.Banlist.Count)) Then
                ReDim Preserve tmpbuf(tmpCount + 1) ' increase array size
                tmpbuf(tmpCount) = Replace(tmpbuf(tmpCount), " , ", Space$(1)) & " [more]" ' apply postfix to previous line
                tmpbuf(tmpCount + 1) = "User(s) banned: " ' apply prefix to new line
                tmpCount = (tmpCount + 1) ' incrememnt counter
            End If
    
            tmpbuf(tmpCount) = Replace(tmpbuf(tmpCount), " , ", Space$(1))
        End If
        
        ' ...
        userCount = 0
    Next I
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnBanned

' handle protect command
Private Function OnProtect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    Select Case (LCase$(msgData))
        Case "on"
            If ((MyFlags = 2) Or (MyFlags = 18)) Then
                Protect = True
                
                tmpbuf = "Lockdown activated by " & Username & "."
                
                Call WildCardBan("*", ProtectMsg, 1)
                
                Call WriteINI("Main", "Protect", "Y")
            Else
                tmpbuf = "The bot does not have ops."
            End If
        
        Case "off"
            If (Protect) Then
                Protect = False
                
                tmpbuf = "Lockdown deactivated."
                
                Call WriteINI("Main", "Protect", "N")
            Else
                tmpbuf = "Protection was not enabled."
            End If
            
        Case "status"
            Select Case (Protect)
                Case True: tmpbuf = "Lockdown is currently active."
                Case Else: tmpbuf = "Lockdown is currently disabled."
            End Select
    End Select
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnProtect

' handle rem command
Private Function OnRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U          As String  ' ...
    Dim tmpbuf     As String  ' temporary output buffer
    Dim user       As udtGetAccessResponse
    Dim dbType     As String  ' ...
    Dim Index      As Long    ' ...
    Dim params     As String  ' ...
    Dim strArray() As String  ' ...
    Dim I          As Integer ' ...

    ' check for presence of optional add command
    ' parameters
    Index = InStr(1, msgData, " --", vbBinaryCompare)

    ' did we find such parameters?
    If (Index > 0) Then
        ' grab parameters
        params = Mid$(msgData, Index - 1)

        ' remove paramaters from message
        msgData = Mid$(msgData, 1, Index)
    End If
    
    ' do we have any special paramaters?
    If (Len(params) > 0) Then
        ' split message by paramter
        strArray() = Split(params, " --")
        
        ' loop through paramter list
        For I = 1 To UBound(strArray)
            Dim Parameter As String ' ...
            Dim pmsg      As String ' ...
            
            ' check message for a space
            Index = InStr(1, strArray(I), Space(1), vbBinaryCompare)
            
            ' did our search find a space?
            If (Index > 0) Then
                ' grab parameter
                Parameter = Mid$(strArray(I), 1, Index - 1)
                
                ' grab parameter message
                pmsg = Mid$(strArray(I), Index + 1)
            Else
                ' grab parameter
                Parameter = strArray(I)
            End If
            
            ' convert parameter to lowercase
            Parameter = LCase$(Parameter)
            
            ' handle parameters
            Select Case (Parameter)
                Case "type" ' ...
                    ' do we have a valid parameter Length?
                    If (Len(pmsg) > 0) Then
                        Dim recType As String ' ...
                    
                        ' grab database entry type
                        recType = UCase$(pmsg)
                        
                        ' ...
                        If (recType = "USER") Then
                            dbType = "USER"
                        ElseIf (recType = "GROUP") Then
                            dbType = "GROUP"
                        ElseIf (recType = "CLAN") Then
                            dbType = "CLAN"
                        ElseIf (recType = "GAME") Then
                            dbType = "GAME"
                        Else
                            dbType = "USER"
                        End If
                    End If
            End Select
        Next I
    End If

    U = msgData
    user = GetAccess(U, dbType)
    
    If (Len(U) > 0) Then
        If ((GetAccess(U, dbType).Rank = -1) And _
            (GetAccess(U, dbType).Flags = vbNullString)) Then
            
            tmpbuf = "User not found."
        ElseIf (GetAccess(U, dbType).Rank >= dbAccess.Rank) Then
            tmpbuf = "That user has higher or equal access."
        ElseIf ((InStr(1, GetAccess(U, dbType).Flags, "L") <> 0) And _
                (Not (InBot)) And _
                (InStr(1, GetAccess(Username, dbType).Flags, "A") = 0) And _
                (GetAccess(Username, dbType).Rank <= 99)) Then
            
                tmpbuf = "Error: That user is Locked."
        Else
            Dim res As Boolean ' ...
        
            res = DB_remove(U, dbType)
            
            If (res) Then
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(RemEntry, IIf(InBot, "console", Username), U, dbType)
                End If
                
                tmpbuf = "Successfully removed database entry " & Chr$(34) & _
                    U & "." & Chr$(34)
            Else
                tmpbuf = "Error: There was a problem removing that entry from the database."
            End If
            
            'Call LoadDatabase
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnRem

' handle idletime command
Private Function OnIdleTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer

    U = msgData
        
    If ((Not (StrictIsNumeric(U))) Or (Val(U) > 50000)) Then
        tmpbuf = "Error setting idle wait time."
    Else
        Call WriteINI("Main", "IdleWait", 2 * Int(U))
        
        tmpbuf = "Idle wait time set to " & Int(U) & " minutes."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIdleTime

' handle idle command
Private Function OnIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
        
    U = LCase$(msgData)
    
    If (U = "on") Then
        Call WriteINI("Main", "Idles", "Y")
        
        tmpbuf = "Idles activated."
    ElseIf (U = "off") Then
        Call WriteINI("Main", "Idles", "N")
        
        tmpbuf = "Idles deactivated."
    ElseIf (U = "kick") Then
        If (InStr(1, msgData, Space(1), vbBinaryCompare) = 0) Then
            tmpbuf = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
        Else
            U = LCase$(Mid$(msgData, InStr(1, msgData, Space(1)) + 1))
            
            If (U = "on") Then
                BotVars.IB_Kick = True
                
                tmpbuf = "Idle kick is now enabled."
            ElseIf (U = "off") Then
                BotVars.IB_Kick = False
                
                tmpbuf = "Idle kick disabled."
            Else
                tmpbuf = "Error: Unknown idle kick command."
            End If
        End If
    Else
        tmpbuf = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIdle

' handle safedel command
Private Function OnSafeDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer

    ReDim Preserve tmpbuf(0)
    
    U = msgData
    
    If (InStr(1, U, Space(1), vbBinaryCompare) <> 0) Then
        tmpbuf(0) = "Error: The specified username is invalid."
    Else
        Call OnAddOld(Username, dbAccess, U & " -S --type USER", True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnSafeDel

' handle tagdel command
Private Function OnTagDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpbuf(0) = "Error: The specified tag is invalid."
    ElseIf (InStr(1, msgData, "*", vbBinaryCompare) <> 0) Then
        ' remove user from shitlist using "add" command
        Call OnAddOld(Username, dbAccess, msgData & " -B --type USER", True, tmpbuf())
        Call OnAddOld(Username, dbAccess, msgData & " -B --type CLAN", True, tmpbuf())
    Else
        ' remove user from shitlist using "add" command
        Call OnAddOld(Username, dbAccess, "*" & msgData & "*" & " -B --type USER", _
            True, tmpbuf())
        Call OnAddOld(Username, dbAccess, "*" & msgData & "*" & " -B --type CLAN", _
            True, tmpbuf())
    End If
        
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnTagDel

' handle setidle command
Private Function OnSetIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String ' ...
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    U = msgData
    
    ' ...
    If (LenB(U) > 0) Then
        ' ...
        Do While (Left$(U, 1) = "/")
            U = Mid$(U, 2)
            If LenB(U) = 0 Then Exit Function
        Loop
        
        ' ...
        Call WriteINI("Main", "IdleMsg", U)
        
        ' ...
        tmpbuf = "Idle message set."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetIdle

' handle idletype command
Private Function OnIdleType(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
        
    U = msgData
    
    If ((LCase$(U) = "msg") Or (LCase$(U) = "message")) Then
        Call WriteINI("Main", "IdleType", "msg")
        
        tmpbuf = "Idle type set to [ msg ]."
    ElseIf ((LCase$(U) = "quote") Or (LCase$(U) = "quotes")) Then
        Call WriteINI("Main", "IdleType", "quote")
        
        tmpbuf = "Idle type set to [ quote ]."
    ElseIf (LCase$(U) = "uptime") Then
        Call WriteINI("Main", "IdleType", "uptime")
        
        tmpbuf = "Idle type set to [ uptime ]."
    ElseIf (LCase$(U) = "mp3") Then
        Call WriteINI("Main", "IdleType", "mp3")
        
        tmpbuf = "Idle type set to [ mp3 ]."
    Else
        tmpbuf = "Error setting idle type. The types are [ message quote uptime mp3 ]."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIdleType

' handle setpmsg command
Private Function OnSetPMsg(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer

    U = msgData
    
    ProtectMsg = U
    
    Call WriteINI("Other", "ProtectMsg", U)
    
    tmpbuf = "Channel protection message set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetPMsg

' handle phrases command
Private Function OnPhrases(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    Dim I        As Integer
    Dim found    As Integer
    Dim temp     As String
    
    ' ...
    For I = LBound(Phrases) To UBound(Phrases)
        ' ...
        If ((Phrases(I) <> Space$(1)) And (Phrases(I) <> vbNullString)) Then
            ' ...
            temp = temp & Phrases(I) & ", "
            
            ' ...
            found = (found + 1)
        End If
    Next I
    
    ' ..
    If (found > 0) Then
        SplitByLen Mid$(temp, 1, Len(temp) - Len(", ")), 180, tmpbuf, "Phrasebans: ", _
            " [more]", ", "
    Else
        tmpbuf(0) = "There are no phrasebans."
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnPhrases

' handle addphrase command
Private Function OnAddPhrase(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim c      As Integer
    Dim tmpbuf As String ' temporary output buffer
    Dim U      As String
    Dim I      As Integer
    
    ' grab free file handle
    f = FreeFile
    
    U = msgData
    
    For I = LBound(Phrases) To UBound(Phrases)
        If (StrComp(U, Phrases(I), vbTextCompare) = 0) Then
            Exit For
        End If
    Next I

    If (I > (UBound(Phrases))) Then
        If ((Phrases(UBound(Phrases)) <> vbNullString) Or _
            (Phrases(UBound(Phrases)) <> " ")) Then
            
            ReDim Preserve Phrases(0 To UBound(Phrases) + 1)
        End If
        
        Phrases(UBound(Phrases)) = U
        
        Open GetFilePath("phrasebans.txt") For Output As #f
            For c = LBound(Phrases) To UBound(Phrases)
                If (Len(Phrases(c)) > 0) Then
                    Print #f, Phrases(c)
                End If
            Next c
        Close #f
        
        tmpbuf = "Phraseban " & Chr(34) & U & Chr(34) & " added."
    Else
        tmpbuf = "Error: That phrase is already banned."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAddPhrase

' handle delphrase command
Private Function OnDelPhrase(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim U      As String
    Dim Y      As String
    Dim tmpbuf As String ' temporary output buffer
    Dim c      As Integer
    
    U = msgData
    
    f = FreeFile
    
    Open GetFilePath("phrasebans.txt") For Output As #f
        Y = vbNullString
    
        For c = LBound(Phrases) To UBound(Phrases)
            If (StrComp(Phrases(c), LCase$(U), vbTextCompare) <> 0) Then
                Print #f, Phrases(c)
            Else
                Y = "x"
            End If
        Next c
    Close #f
    
    ReDim Phrases(0)
    
    Call frmChat.LoadArray(LOAD_PHRASES, Phrases())
    
    If (Len(Y) > 0) Then
        tmpbuf = "Phrase " & Chr(34) & U & Chr(34) & " deleted."
    Else
        tmpbuf = "Error: That phrase is not banned."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDelPhrase

' handle tagadd command
Private Function OnTagAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf() As String  ' ...
    Dim Index    As Integer ' ...
    Dim tag_msg  As String  ' ...
    Dim user     As String  ' ...
    Dim bmsg     As String  ' ...
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' ...
    If (BotVars.DefaultTagbansGroup <> vbNullString) Then
        Dim default_group_access As udtGetAccessResponse
        
        ' ...
        default_group_access = _
                GetAccess(BotVars.DefaultTagbansGroup, "GROUP")
        
        ' ...
        If (default_group_access.Username <> vbNullString) Then
            tag_msg = " --group " & BotVars.DefaultTagbansGroup
        End If
    End If
    
    ' ...
    If (tag_msg = vbNullString) Then
        tag_msg = " +B"
    End If
    
    ' ...
    Index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (Index) Then
        ' ...
        user = Mid$(msgData, 1, Index - 1)
        
        ' ...
        bmsg = Mid$(msgData, Index + 1)
        
        ' ...
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            ' ...
            tmpbuf(0) = "Error: The specified username is invalid."
        End If
    Else
        ' ...
        user = msgData
    End If
    
    ' ..
    If (InStr(1, user, "*", vbBinaryCompare) = 0) Then
        ' ...
        Call OnAddOld(Username, dbAccess, user & tag_msg & " --type CLAN", True, tmpbuf())
    Else
        ' ...
        Call OnAddOld(Username, dbAccess, user & tag_msg & " --type USER", True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnTagAdd

' handle safelist command
Private Function OnSafeList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase2(tmpbuf(), , , , , , , "S")
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnSafeList

' handle safeadd command
Private Function OnSafeAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    Dim safe_msg As String ' ...
    
    ReDim Preserve tmpbuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpbuf(0) = "Error: The specified username is invalid."
    Else
        ' ...
        If (BotVars.DefaultSafelistGroup <> vbNullString) Then
            Dim default_group_access As udtGetAccessResponse
            
            ' ...
            default_group_access = _
                    GetAccess(BotVars.DefaultSafelistGroup, "GROUP")
            
            ' ...
            If (default_group_access.Username <> vbNullString) Then
                safe_msg = " --group " & BotVars.DefaultSafelistGroup
            End If
        End If
        
        ' ...
        If (safe_msg = vbNullString) Then
            safe_msg = " +S"
        End If
    
        Call OnAddOld(Username, dbAccess, msgData & safe_msg, True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnSafeAdd

' handle safecheck command
Private Function OnSafeCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String ' ...
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    Y = msgData
            
    ' ...
    If (LenB(Y) > 0) Then
        ' ...
        If (GetSafelist(Y)) Then
            tmpbuf = Y & " is on the bot's safelist."
        Else
            tmpbuf = "That user is not safelisted."
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSafeCheck

' handle shitlist command
Private Function OnShitList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase2(tmpbuf(), , "!*[*]*", , , , , "B")
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnShitList

' handle tagbans command
Private Function OnTagBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase2(tmpbuf(), , "*[*]*", , , , , "B")
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnTagBans

' handle bancount command
Private Function OnBanCount(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (g_Channel.BanCount = 0) Then
        tmpbuf = "No users have been banned since I joined this channel."
    Else
        tmpbuf = "Since I joined this channel, " & g_Channel.BanCount & " user(s) have " & _
            "been banned."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnBanCount

' handle banlistcount command
Private Function OnBanListCount(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (g_Channel.BanCount = 0) Then
        ' ...
        tmpbuf = "There are currently no users on the internal ban list."
    Else
        Dim bCount As Integer ' ...
        Dim I      As Integer ' ...
    
        ' ...
        tmpbuf = "There are currently " & g_Channel.Banlist.Count & " user(s) on the internal ban list"
                        
        ' ...
        If (g_Channel.Self.IsOperator) Then
            'tmpBuf = tmpBuf & ", " & g_Channel.Self.Banlist.Count & " of which users were " & _
            '            "banned by me."
        Else
            tmpbuf = tmpbuf & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnBanCount

' handle tagcheck command
Private Function OnTagCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim tmpbuf As String ' temporary output buffer

    Y = msgData
            
    If (Len(Y) > 0) Then
        gAcc = GetCumulativeAccess(Y)
        
        If (InStr(1, gAcc.Flags, "B") <> 0) Then
            tmpbuf = Y & " has been matched to one or more tagbans"
        
            If (InStr(1, gAcc.Flags, "S") <> 0) Then
                tmpbuf = tmpbuf & "; however, " & Y & " has also been found on the bot's " & _
                    "safelist and therefore will not be banned"
            End If
        Else
            tmpbuf = "That user matches no tagbans"
        End If
        
        tmpbuf = tmpbuf & "."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnTagCheck

' handle slcheck command
Private Function OnSLCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim tmpbuf As String ' temporary output buffer

    Y = msgData
            
    If (Len(Y) > 0) Then
        gAcc = GetCumulativeAccess(Y)
        
        If (InStr(1, gAcc.Flags, "B") <> 0) Then
            tmpbuf = Y & " is on the bot's shitlist"
        
            If (InStr(1, gAcc.Flags, "S") <> 0) Then
                tmpbuf = tmpbuf & "; however, " & Y & " is also on the " & _
                    "bot's safelist and therefore will not be banned"
            End If
        Else
            tmpbuf = "That user is not shitlisted"
        End If
        
        tmpbuf = tmpbuf & "."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSLCheck

' TO DO:
' handle greet command
Private Function OnGreet(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf       As String ' temporary output buffer
    Dim strSplit()   As String
    Dim greetCommand As String
    
    ' do we have parameters?
    If (msgData = vbNullString) Then
        tmpbuf = "Greet messages are currently "
        tmpbuf = tmpbuf & IIf(BotVars.UseGreet, "enabled", "disabled") & "."
        cmdRet(0) = tmpbuf
        Exit Function
    End If
    
    ' split string by spaces
    strSplit() = Split(msgData, Space(1), 2)

    ' grab greet command
    greetCommand = strSplit(0)

    ' ...
    If (greetCommand = "on") Then
        BotVars.UseGreet = True
        
        tmpbuf = "Greet messages enabled."
        
        Call WriteINI("Other", "UseGreets", "Y")
    ElseIf (greetCommand = "off") Then
        BotVars.UseGreet = False
        
        tmpbuf = "Greet messages disabled."
        
        Call WriteINI("Other", "UseGreets", "N")
    ElseIf (greetCommand = "whisper") Then
        Dim greetSubCommand As String ' ...
        
        ' grab greet sub command
        greetSubCommand = strSplit(1)
        
        ' ...
        If (greetSubCommand = "on") Then
            BotVars.WhisperGreet = True
            
            tmpbuf = "Greet messages will now be whispered."
            
            Call WriteINI("Other", "WhisperGreet", "Y")
        ElseIf (greetSubCommand = "off") Then
            BotVars.WhisperGreet = False
            
            tmpbuf = "Greet messages will no longer be whispered."
            
            Call WriteINI("Other", "WhisperGreet", "N")
        End If
    Else
        Dim GreetMessage As String ' ...
        
        ' grab greet message
        GreetMessage = msgData
    
        ' ...
        If (Left$(GreetMessage, 1) = "/") Then
            tmpbuf = "Error: Invalid greet message specified."
        Else
            tmpbuf = "Greet message set."
            
            BotVars.GreetMsg = GreetMessage
            
            Call WriteINI("Other", "GreetMsg", GreetMessage)
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnGreet

' handle ban command
Private Function OnBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim U       As String
    Dim tmpbuf  As String ' temporary output buffer
    Dim banmsg  As String
    Dim Y       As String
    Dim I       As Integer

    If ((MyFlags And USER_CHANNELOP&) <> USER_CHANNELOP&) Then
        If (InBot) Then
            tmpbuf = "Error: I am not currently a channel operator."
        End If
    Else
        U = msgData
        
        If (U <> vbNullString) Then
            I = InStr(1, U, Space(1), vbBinaryCompare)
            
            If (I > 0) Then
                banmsg = Mid$(U, I + 1)
                
                U = Left$(U, I - 1)
            End If
            
            If (InStr(1, U, "*", vbBinaryCompare) <> 0) Then
                tmpbuf = WildCardBan(U, banmsg, 1)
            Else
                If (InBot) Then
                    frmChat.AddQ "/ban " & msgData
                Else
                    Y = Ban(U & IIf(banmsg <> vbNullString, Space$(1) & banmsg, _
                        vbNullString), dbAccess.Rank)
                End If
            End If
            
            If (Len(Y) > 2) Then
                tmpbuf = Y
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnBan

' handle unban command
Private Function OnUnban(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    If ((MyFlags And USER_CHANNELOP&) <> USER_CHANNELOP&) Then
       If (InBot) Then
           tmpbuf = "Error: I am not currently a channel operator."
       End If
    Else
        U = msgData
        
        If (U <> vbNullString) Then
            If (Dii = True) Then
                If (Not (Mid$(U, 1, 1) = "*")) Then
                    U = "*" & U
                End If
            End If
            
            ' what the hell is a flood cap?
            If (bFlood) Then
                If (floodCap < 45) Then
                    floodCap = (floodCap + 15)
                    
                    ' bnetSend?
                    Call AddQ("/unban " & U)
                End If
            Else
                If (InStr(1, msgData, "*", vbBinaryCompare) <> 0) Then
                    Call WildCardBan(U, vbNullString, 2)
                Else
                    Call AddQ("/unban " & U, 1, Username)
                End If
            End If
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUnBan

' handle kick command
Private Function OnKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim I      As Integer
    Dim banmsg As String
    Dim tmpbuf As String ' temporary output buffer
    Dim Y      As String
    
    If ((MyFlags And USER_CHANNELOP&) <> USER_CHANNELOP&) Then
       tmpbuf = "Error: I am not currently a channel operator."
    Else
        U = msgData
        
        If (Len(U) > 0) Then
            I = InStr(1, U, " ", vbTextCompare)
            
            If (I > 0) Then
                banmsg = Mid$(U, I + 1)
                
                U = Left$(U, I - 1)
            End If
            
            If (InStr(1, U, "*", vbTextCompare) > 0) Then
                If (dbAccess.Rank >= 100) Then
                    tmpbuf = WildCardBan(U, banmsg, 0)
                Else
                    tmpbuf = WildCardBan(U, banmsg, 0)
                End If
            Else
                If (InBot) Then
                    frmChat.AddQ "/kick " & msgData
                Else
                    Y = Ban(U & IIf(Len(banmsg) > 0, Space$(1) & banmsg, vbNullString), _
                        dbAccess.Rank, 1)
                    
                    If (Len(Y) > 1) Then
                        tmpbuf = Y
                    End If
                End If
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnKick

' handle voteban command
Private Function OnVoteBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    Dim user   As String
    Dim dur    As String
    
    user = msgData
    
    If (VoteDuration = -1) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, user)
        
        VoteDuration = 30
        
        VoteInitiator = dbAccess
        
        tmpbuf = "30-second VoteBan vote started. Type YES to ban " & user & ", NO to acquit him/her."
    Else
        tmpbuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnVoteBan

' handle votekick command
Private Function OnVoteKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf   As String ' temporary output buffer
    Dim user     As String
    
    user = msgData
    
    If (VoteDuration = -1) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, user)
        
        VoteDuration = 30
        
        VoteInitiator = dbAccess
        
        tmpbuf = "30-second VoteKick vote started. Type YES to kick " & user & _
            ", NO to acquit him/her."
    Else
        tmpbuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnVoteKick

' handle vote command
Private Function OnVote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
 
    Dim tmpbuf      As String ' temporary output buffer
    Dim tmpDuration As Long
    
    If (Len(msgData) > 0) Then
        If (VoteDuration = -1) Then
            ' ensure that tmpDuration is an integer
            tmpDuration = Val(msgData)
        
            ' check for proper duration time and call for vote
            If ((tmpDuration > 0) And (tmpDuration <= 32000)) Then
                ' set vote duration
                VoteDuration = tmpDuration
                
                ' set vote initiator
                VoteInitiator = dbAccess
                
                ' execute vote
                Call Voting(BVT_VOTE_START, BVT_VOTE_STD)
                
                tmpbuf = "Vote initiated. Type YES or NO to vote; your vote will " & _
                    "be counted only once."
            Else
                ' duration entered is either negative, is too large, or is a string
                tmpbuf = "Please enter a number of seconds for your vote to last."
            End If
        Else
            tmpbuf = "A vote is currently in progress."
        End If
    Else
        ' duration not entered
        tmpbuf = "Please enter a number of seconds for your vote to last."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnVote

' handle addquote command
Private Function OnAddQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (g_Quotes Is Nothing) Then
        Set g_Quotes = New Collection
    End If
    
    g_Quotes.Add (msgData)
    tmpbuf = "Quote added!"
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAddQuote

' handle quote command
Private Function OnQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = "Quote: " & _
        g_Quotes.GetRandomQuote
    
    ' this was len() = 0, which doesn't work cause we add "Quote:" above -andy
    If (Len(tmpbuf) < 8) Then
        tmpbuf = "Error reading your quotes, or no quote file exists."
    ElseIf (Len(tmpbuf) > 223) Then
        ' try one more time
        tmpbuf = "Quote: " & _
            g_Quotes.GetRandomQuote
        
        If (Len(tmpbuf) > 223) Then
            'too long? too bad. truncate
            tmpbuf = Left$(tmpbuf, 223)
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnQuote

' handle getmail command
Private Function OnInbox(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Msg      As udtMail
    Dim tmpbuf() As String ' temporary output buffer
    Dim mcount   As Integer
    Dim Index    As Integer
            
    If (InBot) Then
        If (g_Online) Then
            Username = GetCurrentUsername
        Else
            Username = BotVars.Username
        End If
    End If
    
    ' ...
    mcount = GetMailCount(Username)
    
    ' ...
    If (mcount > 1) Then
        ReDim tmpbuf(mcount - 1)
    Else
        ReDim tmpbuf(0)
    End If
    
    ' ...
    If (mcount > 0) Then
        ' ...
        Do
            ' ...
            GetMailMessage Username, Msg
            
            ' ...
            If (Len(RTrim(Msg.To)) > 0) Then
                ' ...
                tmpbuf(Index) = "Message from " & RTrim$(Msg.From) & ": " & _
                    RTrim$(Msg.Message)
                    
                ' ...
                Index = Index + 1
            End If
        Loop While (GetMailCount(Username) > 0)
    Else
        If (dbAccess.Rank > 0) Then
            tmpbuf(0) = "You do not currently have any messages " & _
                "in your inbox."
        End If
    End If
    
    ' return message
    If (InBot) Then
        cmdRet() = tmpbuf()
    Else
        Dim I As Integer ' ...
        
        ' ...
        For I = 0 To UBound(tmpbuf)
            If (tmpbuf(I) <> vbNullString) Then
                AddQ "/w " & Username & " " & tmpbuf(I)
            End If
        Next I
    End If
End Function ' end function OnGetMail

Public Function OnAddOld(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    ' ...
    Dim gAcc       As udtGetAccessResponse

    Dim strArray() As String  ' ...
    Dim I          As Integer ' ...
    Dim tmpbuf     As String  ' temporary output buffer
    Dim dbPath     As String  ' ...
    Dim user       As String  ' ...
    Dim Rank       As Integer ' ...
    Dim Flags      As String  ' ...
    Dim found      As Boolean ' ...
    Dim params     As String  ' ...
    Dim Index      As Integer ' ...
    Dim sGrp       As String  ' ...
    Dim dbType     As String  ' ...
    Dim banmsg     As String  ' ...

    ' check for presence of optional add command
    ' parameters
    Index = InStr(1, msgData, " --", vbBinaryCompare)
    
    ' did we find such parameters, and if so,
    ' do they begin after an entry name?
    If (Index > 1) Then
        ' grab parameters
        params = Mid$(msgData, Index - 1)

        ' remove paramaters from message
        msgData = Mid$(msgData, 1, Index)
    End If
    
    ' does our message contain an entry name? rank? flags?
    ' anything? we don't want to error out if not.
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        ' split message
        strArray() = Split(msgData, Space(1))
    Else
        Exit Function
    End If
    
    If (UBound(strArray) > 0) Then
        ' grab username
        user = strArray(0)
        
        ' ...
        If (user = vbNullString) Then
            cmdRet(0) = "Error: You have specified an invalid entry name."
            
            Exit Function
        End If
        
        ' grab rank & flags
        If (StrictIsNumeric(strArray(1))) Then
            ' grab rank
            Rank = strArray(1)
            
            ' grab flags
            If (UBound(strArray) >= 2) Then
                Flags = strArray(2)
            End If
        Else
            ' grab flags
            Flags = strArray(1)
        End If
        
        ' ...
        If (BotVars.CaseSensitiveFlags = False) Then
            Flags = UCase$(Flags)
        End If
        
        ' do we have any special paramaters?
        If (Len(params)) Then
            ' split message by paramter
            strArray() = Split(params, " --")
            
            ' loop through paramter list
            For I = 1 To UBound(strArray)
                Dim Parameter As String ' ...
                Dim pmsg      As String ' ...
                
                ' check message for a space
                Index = InStr(1, strArray(I), Space(1), vbBinaryCompare)
                
                ' did our search find a space?
                If (Index > 0) Then
                    ' grab parameter
                    Parameter = Mid$(strArray(I), 1, Index - 1)
                    
                    ' grab parameter message
                    pmsg = Mid$(strArray(I), Index + 1)
                Else
                    ' grab parameter
                    Parameter = strArray(I)
                End If
                
                ' convert parameter to lowercase
                Parameter = LCase$(Parameter)
                
                ' handle parameters
                Select Case (Parameter)
                    Case "type" ' ...
                        ' do we have a valid parameter Length?
                        If (Len(pmsg)) Then
                            ' grab database entry type
                            dbType = UCase$(pmsg)
                            
                            ' ...
                            If (dbType = "USER") Then
                                ' ...
                            ElseIf (dbType = "GROUP") Then
                                ' check for presence of space in name
                                If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
                                    cmdRet(0) = "Error: The specified group name contains one or more " & _
                                        "invalid characters."
                                                                
                                    Exit Function
                                End If
                            ElseIf (dbType = "CLAN") Then
                                ' check for invalid clan entry
                                If ((Len(user) < 2) Or (Len(user) > 4)) Then
                                    ' return message
                                    cmdRet(0) = "Error: The clan name specified is of an " & _
                                        "incorrect Length."
                                        
                                    Exit Function
                                End If
                            ElseIf (dbType = "GAME") Then
                                ' convert entry to uppercase
                                user = UCase$(user)
                                
                                ' check for invalid game entry
                                Select Case (user)
                                    Case "CHAT" ' Chat Client
                                    Case "DRTL" ' Diablo I: Retail
                                    Case "DSHR" ' Diablo I: Shareware
                                    Case "W2BN" ' WarCraft II: Battle.net Edition
                                    Case "STAR" ' StarCraft
                                    Case "SSHR" ' StarCraft: Shareware
                                    Case "JSTR" ' StarCraft: Japanese
                                    Case "SEXP" ' StarCraft: Brood War
                                    Case "D2DV" ' Diablo II
                                    Case "D2XP" ' Diablo II: Lord of Destruction
                                    Case "WAR3" ' WarCraft III: Reign of Chaos
                                    Case "W3XP" ' WarCraft III: The Frozen Throne
                                    Case Else
                                        ' return message
                                        cmdRet(0) = "Error: The game specified is invalid."
                                        
                                        Exit Function
                                End Select
                            End If
                        End If
                
                    Case "banmsg" ' ...
                        ' do we have a valid parameter Length?
                        If (Len(pmsg)) Then
                            banmsg = pmsg
                        End If
                        
                    Case "group" ' ...
                        ' do we have a valid parameter Length?
                        If (Len(pmsg)) Then
                            Dim Splt() As String
                            Dim j      As Integer
                        
                            If (InStr(1, pmsg, ",", vbBinaryCompare) <> 0) Then
                                ' we no longer officially support the use of multiple
                                ' user groupings; however, manual database modifications
                                ' will still allow users to do so if the need ever arises.
                                'Splt() = Split(pmsg, ",")
                                
                                cmdRet(0) = "Error: The specified group name contains one or more " & _
                                        "invalid characters."
                                        
                                Exit Function
                            Else
                                ReDim Preserve Splt(0)
                                
                                Splt(0) = pmsg
                            End If
                            
                            For j = 0 To UBound(Splt)
                                Dim tmp As udtGetAccessResponse ' ...
                                
                                ' ...
                                tmp = GetAccess(Splt(j), "GROUP")
                            
                                If (dbAccess.Rank < tmp.Rank) Then
                                    cmdRet(0) = "Error: You do not have sufficient access to " & _
                                        "add a member to the specified group."
                                        
                                    Exit Function
                                End If
                                
                                If ((StrComp(Splt(j), user, vbTextCompare) = 0) And _
                                    (dbType = "GROUP")) Then
                                    
                                    cmdRet(0) = "Error: You cannot make a group a member of " & _
                                        "itself."
                                        
                                    Exit Function
                                Else
                                    If (tmp.Username = vbNullString) Then
                                        Exit For
                                    Else
                                        ' we need to check to make sure that we aren't allowing
                                        ' two groups to be members of each other, potentially
                                        ' causing a stack overflow when doing recursion in
                                        ' GetCumulativeAccess().
                                        If ((Len(tmp.Groups)) And (tmp.Groups <> "%")) Then
                                            If (CheckGroup(tmp.Username, user)) Then
                                                cmdRet(0) = "Error: " & Chr$(34) & tmp.Username & _
                                                    Chr$(34) & " is already a member of group " & _
                                                        Chr$(34) & user & "." & Chr$(34)
                                        
                                                    Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                            
                            If (j < (UBound(Splt) + 1)) Then
                                cmdRet(0) = "Error: The specified group(s) could " & _
                                    "not be found."
                                    
                                Exit Function
                            Else
                                sGrp = pmsg
                            End If
                        End If
                End Select
            Next I
        End If
        
        ' we want to ensure that we have a default
        ' entry type if none is specified explicitly
        If (dbType = vbNullString) Then
            dbType = "USER"
        End If
        
        ' grab access for entry
        gAcc = GetAccess(user, dbType)
        
        ' if we've found a matching user, lets correct
        ' the casing of the name that we've entered
        If (Len(gAcc.Username) > 0) Then
            If (StrComp(gAcc.Type, dbType, vbTextCompare) = 0) Then
                user = gAcc.Username
            End If
        End If
        
        ' grab access for entry
        gAcc = GetCumulativeAccess(user, dbType)

        ' is rank valid?
        If ((Rank <= 0) And (Flags = vbNullString) And (sGrp = vbNullString)) Then
            
            If ((Rank = 0) And ((gAcc.Rank > 0) Or (gAcc.Flags <> vbNullString) Or _
                (gAcc.Groups <> vbNullString))) Then
                Call OnRem(Username, dbAccess, user, InBot, cmdRet)
            Else
                cmdRet(0) = "Error: You have specified an invalid rank."
            End If
            
            Exit Function
            
        ' is rank higher than user's rank?
        ElseIf ((Rank) And (Rank >= dbAccess.Rank)) Then
            cmdRet(0) = "Error: You do not have sufficient access to assign an entry with the " & _
                "specified rank."
            Exit Function
        ' can we modify specified user?
        ElseIf ((gAcc.Rank) And (gAcc.Rank >= dbAccess.Rank)) Then
            cmdRet(0) = "Error: You do not have sufficient access to modify the specified entry."
            Exit Function
        Else
            ' did we specify flags?
            If (Len(Flags)) Then
                Dim currentCharacter As String ' ...
            
                For I = 1 To Len(Flags)
                    currentCharacter = Mid$(Flags, I, 1)
                
                    If ((currentCharacter <> "+") And (currentCharacter <> "-")) Then
                        'Select Case (currentCharacter)
                        '    Case "A" ' administrator
                        '        If (dbAccess.Rank <= 100) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "B" ' banned
                        '        If (dbAccess.Rank < 70) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "D" ' designated
                        '        If (dbAccess.Rank < 100) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "L" ' locked
                        '        If (dbAccess.Rank < 70) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "S" ' safelisted
                        '        If (dbAccess.Rank < 70) Then
                        '            Exit For
                        '        End If
                        'End Select
                    End If
                Next I
                
                If (I < (Len(Flags) + 1)) Then
                    ' return message
                    cmdRet(0) = "Error: You do not have sufficient access to add one or " & _
                        "more flags specified."
                    
                    Exit Function
                Else
                    ' are we adding flags?
                    If (Left$(Flags, 1) = "+") Then
                        ' remove "+" prefix
                        Flags = Mid$(Flags, 2)
                    
                        ' ...
                        If (Len(Flags) > 0) Then
                            ' set user flags & check for duplicate entries
                            For I = 1 To Len(Flags)
                                currentCharacter = Mid$(Flags, I, 1)
                            
                                ' is flag valid (alphabetic)?
                                If (((Asc(currentCharacter) >= Asc("A")) And (Asc(currentCharacter) <= Asc("Z"))) Or _
                                    ((Asc(currentCharacter) >= Asc("a")) And (Asc(currentCharacter) <= Asc("z")))) Then
                                    
                                    If (InStr(1, gAcc.Flags, currentCharacter, vbBinaryCompare) = 0) Then
                                        gAcc.Flags = gAcc.Flags & currentCharacter
                                    End If
                                End If
                            Next I
                            
                            ' ...
                            If (Len(gAcc.Flags) = 0) Then
                                ' return message
                                cmdRet(0) = "Error: The flag(s) that you have specified are invalid."
                            
                                Exit Function
                            End If
                        Else
                            ' return message
                            cmdRet(0) = "Error: You must specify at least one flag for addition."
                            
                            Exit Function
                        End If

                    ' are we removing flags?
                    ElseIf (Left$(Flags, 1) = "-") Then
                        Dim tmpFlags As String
                    
                        ' remove "-" prefix
                        tmpFlags = Mid$(Flags, 2)
                        
                        ' are we modifying an existing user? we better be!
                        If (gAcc.Username <> vbNullString) Then
                            ' ...
                            If (Len(tmpFlags) > 0) Then
                                ' check for special flags
                                If (InStr(1, tmpFlags, "B", vbBinaryCompare) <> 0) Then
                                    If (InStr(1, user, "*", vbBinaryCompare) <> 0) Then
                                        Call WildCardBan(user, vbNullString, 2)
                                    Else
                                        ' ...
                                        If (g_Channel.IsOnBanList(user)) Then
                                            Call AddQ("/unban " & user)
                                        End If
                                    End If
                                End If
                                
                                ' remove specified flags
                                For I = 1 To Len(tmpFlags)
                                    gAcc.Flags = Replace(gAcc.Flags, Mid$(tmpFlags, I, 1), _
                                        vbNullString)
                                Next I
                            Else
                                ' return message
                                cmdRet(0) = "Error: You must specify at least one flag " & _
                                    "for removal."
                            
                                Exit Function
                            End If
                        Else
                            ' return message
                            cmdRet(0) = "Error: The specified database entry was not found."
                        
                            Exit Function
                        End If
                        
                        ' does this entry have any remaining access?
                        If ((gAcc.Rank = 0) And (gAcc.Flags = vbNullString) And _
                            ((gAcc.Groups = vbNullString) Or (gAcc.Groups = "%"))) Then
                            
                            Dim res As Boolean ' ...
                           
                            ' with no access a database entry is
                            ' pointless, so lets remove it
                            res = DB_remove(user, gAcc.Type)
                            
                            If (res) Then
                                cmdRet(0) = Chr(34) & user & Chr(34) & " has been removed " & _
                                    "from the database."
                            Else
                                cmdRet(0) = "Error: There was a problem removing that entry " & _
                                    "from the database."
                            End If
                                
                            Exit Function
                        End If
                    Else
                        ' if we're adding with no flag indicator ('+' or '-'),
                        ' then we need to remove the previous entry from the database.
                        'Call DB_remove(user, dbType)
                    
                        ' clear user flags
                        gAcc.Flags = vbNullString
                        
                        ' set rank to specified
                        gAcc.Rank = Rank
                    
                        ' set user flags & check for duplicate entries
                        For I = 1 To Len(Flags)
                            currentCharacter = Mid$(Flags, I, 1)
                        
                            ' is flag valid (alphabetic)?
                            If (((Asc(currentCharacter) >= Asc("A")) And (Asc(currentCharacter) <= Asc("Z"))) Or _
                                ((Asc(currentCharacter) >= Asc("a")) And (Asc(currentCharacter) <= Asc("z")))) Then
                                
                                If (InStr(1, gAcc.Flags, currentCharacter, vbBinaryCompare) = 0) Then
                                    gAcc.Flags = gAcc.Flags & currentCharacter
                                End If
                            End If
                        Next I
                        
                        ' ...
                        If (Len(gAcc.Flags) = 0) Then
                            ' return message
                            cmdRet(0) = "Error: The flag(s) that you have specified are invalid."
                        
                            Exit Function
                        End If
                    End If
                End If
            Else
                ' if we're adding with no flag indicator ('+' or '-'),
                ' then we need to remove the previous entry from the database.
                'Call DB_remove(user, dbType)

                ' clear flags
                gAcc.Flags = vbNullString
            
                ' set rank to specified
                gAcc.Rank = Rank
            End If

            ' grab path to database
            dbPath = GetFilePath("users.txt")

            ' does user already exist in database?
            For I = LBound(DB) To UBound(DB)
                If ((StrComp(DB(I).Username, user, vbTextCompare) = 0) And _
                    (StrComp(DB(I).Type, gAcc.Type, vbTextCompare) = 0)) Then
                    
                    ' modify database entry
                    With DB(I)
                        .Username = user
                        .Rank = gAcc.Rank
                        .Flags = gAcc.Flags
                        .ModifiedBy = Username
                        .ModifiedOn = Now
                        .Type = dbType
                        .Groups = sGrp
                        .BanMessage = banmsg
                    End With
                
                    ' commit modifications
                    Call WriteDatabase(dbPath)
                    
                    ' log actions
                    If (BotVars.LogDBActions) Then
                        Call LogDBAction(ModEntry, IIf(InBot, "console", Username), DB(I).Username, _
                            DB(I).Type, DB(I).Rank, DB(I).Flags, DB(I).Groups)
                    End If
                    
                    ' we have found the
                    ' specified user
                    found = True
                    
                    Exit For
                End If
            Next I
            
            ' did we find a matching entry or not?
            If (found = False) Then

                ' redefine array size
                If (DB(0).Username = vbNullString) Then
                    ReDim Preserve DB(0)
                Else
                    ReDim Preserve DB(UBound(DB) + 1)
                End If

                With DB(UBound(DB))
                    .Username = user
                    .Rank = IIf((gAcc.Rank >= 0), _
                        gAcc.Rank, 0)
                    .Flags = gAcc.Flags
                    .ModifiedBy = Username
                    .ModifiedOn = Now
                    .AddedBy = Username
                    .AddedOn = Now
                    .Type = IIf(((dbType <> vbNullString) And (dbType <> "%")), _
                        dbType, "USER")
                    .Groups = sGrp
                    .BanMessage = banmsg
                End With
                
                'MsgBox dbPath
                
                ' commit modifications
                Call WriteDatabase(dbPath)
                
                ' log actions
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(AddEntry, IIf(InBot, "console", Username), DB(UBound(DB)).Username, _
                        DB(UBound(DB)).Type, DB(UBound(DB)).Rank, DB(UBound(DB)).Flags, DB(UBound(DB)).Groups)
                End If
            End If
            
            ' check for errors & create message
            If (gAcc.Rank > 0) Then
                tmpbuf = Chr(34) & user & Chr(34) & " has been given access " & _
                    gAcc.Rank
                
                ' was the user given the specified flags, too?
                If (Len(gAcc.Flags)) Then
                    ' lets make sure we don't use
                    ' improper grammar because of groups!
                    If (Len(sGrp)) Then
                        tmpbuf = tmpbuf & ", flags " & gAcc.Flags
                    Else
                        tmpbuf = tmpbuf & " and flags " & gAcc.Flags
                    End If
                End If
            Else
                ' was the user given the specified flags?
                If (Len(gAcc.Flags)) Then
                    tmpbuf = Chr(34) & user & Chr(34) & " has been given flags " & _
                        gAcc.Flags
                End If
            End If
            
            ' was the user assigned to a group?
            If (Len(sGrp)) Then
                If (Len(tmpbuf)) Then
                    tmpbuf = tmpbuf & ", and has been made a member of " & _
                        "the group(s): " & sGrp
                Else
                    tmpbuf = Chr(34) & user & Chr(34) & " has been made a member of " & _
                        "the group(s): " & sGrp
                End If
            End If
            
            ' terminate sentence
            ' with period
            tmpbuf = tmpbuf & "."
        End If
        
        ' ...
        Call g_Channel.CheckUsers
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAdd

Private Function AddQ(ByVal s As String, Optional msg_priority As Integer = -1, Optional ByVal user As String = vbNullString, Optional ByVal Tag As String = vbNullString) As Integer
    AddQ = frmChat.AddQ(s, msg_priority, user, Tag)
End Function

Private Function WildCardBan(ByVal sMatch As String, ByVal smsgData As String, ByVal Banning As Byte) As String
    'Values for Banning byte:
    '0 = Kick
    '1 = Ban
    '2 = Unban
    
    Dim I     As Integer
    Dim Typ   As String
    Dim z     As String
    Dim iSafe As Integer
    
    If ((MyFlags = 2) Or (MyFlags = 18)) Then
        If (smsgData = vbNullString) Then
            smsgData = sMatch
        End If
        
        sMatch = PrepareCheck(sMatch)
        
        Select Case (Banning)
            Case 1: Typ = "ban "
            Case 2: Typ = "unban "
            Case Else: Typ = "kick "
        End Select
        
        If (Dii) Then
            Typ = Typ & "*"
        End If
        
        If (g_Channel.Users.Count < 1) Then
            Exit Function
        End If
        
        If (Banning <> 2) Then
            ' Kicking or Banning
        
            For I = 1 To g_Channel.Users.Count
                With g_Channel.Users(I)
                    If (StrComp(g_Channel.Users(I).DisplayName, GetCurrentUsername, vbBinaryCompare) <> 0) Then
                        z = PrepareCheck(.DisplayName)
                        
                        If (z Like sMatch) Then
                            If (GetSafelist(.DisplayName) = False) Then
                                If (.IsOperator = False) Then
                                    Call AddQ("/" & Typ & .DisplayName & Space(1) & smsgData)
                                End If
                            Else
                                iSafe = (iSafe + 1)
                            End If
                        End If
                    End If
                End With
            Next I
            
            If (iSafe) Then
                If (StrComp(smsgData, ProtectMsg, vbTextCompare) <> 0) Then
                    WildCardBan = "Encountered " & iSafe & " safelisted user(s)."
                End If
            End If
            
        Else '// unbanning
        
            For I = 1 To g_Channel.Banlist.Count
                If ((g_Channel.Banlist(I).IsActive) And (g_Channel.Banlist(I).DisplayName <> vbNullString)) Then
                    If (sMatch = "*") Then
                        Call AddQ("/" & Typ & g_Channel.Banlist(I).DisplayName)
                    Else
                        z = PrepareCheck(g_Channel.Banlist(I).DisplayName)
                        
                        If (z Like sMatch) Then
                            Call AddQ("/" & Typ & g_Channel.Banlist(I).DisplayName)
                        End If
                    End If
                End If
            Next I
        End If
    End If
End Function

Private Function searchDatabase2(ByRef arrReturn() As String, Optional user As String = vbNullString, _
    Optional ByVal match As String = vbNullString, Optional Group As String = vbNullString, _
        Optional dbType As String = vbNullString, Optional lowerBound As Integer = -1, _
            Optional upperBound As Integer = -1, Optional Flags As String = vbNullString) As Integer
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    Dim I        As Integer
    Dim found    As Integer
    Dim tmpbuf   As String
    
    If (user <> vbNullString) Then
        ' store GetAccess() response
        Dim gAcc As udtGetAccessResponse
    
        ' grab user access
        gAcc = GetAccess(user, dbType)
        
        If ((Not gAcc.Type = "%") And (StrComp(gAcc.Type, "USER", vbTextCompare) <> 0)) Then
            gAcc.Username = StringFormatA("{0} ({1})", gAcc.Username, gAcc.Type)
        End If
        
        If (gAcc.Rank > 0) Then
            If (LenB(gAcc.Flags) > 0) Then
                tmpbuf = StringFormatA("Found user {0}, who holds rank {1} and flags {2}.", _
                    gAcc.Username, gAcc.Rank, gAcc.Flags)
            Else
                tmpbuf = StringFormatA("Found user {0} who holds rank {1}.", gAcc.Username, gAcc.Rank)
            End If
        ElseIf (LenB(gAcc.Flags) > 0) Then
            tmpbuf = StringFormatA("Found user {0}, with flags {1}.", gAcc.Username, gAcc.Flags)
        Else
            tmpbuf = "No such user(s) found."
        End If
    Else
        For I = LBound(DB) To UBound(DB)
            Dim res        As Boolean ' store result of access check
            Dim blnChecked As Boolean ' ...
        
            If (LenB(DB(I).Username) > 0) Then
                If (match <> vbNullString) Then
                    If (Left$(match, 1) = "!") Then
                        res = (Not (LCase$(PrepareCheck(DB(I).Username)) Like (LCase$(Mid$(match, 2)))))
                    Else
                        res = (LCase$(PrepareCheck(DB(I).Username)) Like (LCase$(match)))
                    End If
                    
                    blnChecked = True
                End If
                
                If (LenB(Group) > 0) Then
                    If (StrComp(DB(I).Groups, Group, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If

                ' ...
                If (dbType <> vbNullString) Then
                    ' ...
                    If (StrComp(DB(I).Type, dbType, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If
                
                ' ...
                If ((lowerBound >= 0) And (upperBound >= 0)) Then
                    If ((DB(I).Rank >= lowerBound) And _
                        (DB(I).Rank <= upperBound)) Then
                        
                        ' ...
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                ElseIf (lowerBound >= 0) Then
                    If (DB(I).Rank = lowerBound) Then
                        ' ...
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If
                
                ' ...
                If (Flags <> vbNullString) Then
                    Dim j As Integer ' ...
                
                    For j = 1 To Len(Flags)
                        If (InStr(1, DB(I).Flags, Mid$(Flags, j, 1), _
                            vbBinaryCompare) = 0) Then
                            
                            Exit For
                        End If
                    Next j
                    
                    If (j = (Len(Flags) + 1)) Then
                        ' ...
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If
                
                ' ...
                If (res = True) Then
                    ' ...
                    tmpbuf = tmpbuf & DB(I).Username & _
                        IIf(((DB(I).Type <> "%") And _
                                (StrComp(DB(I).Type, "USER", vbTextCompare) <> 0)), _
                            " (" & LCase$(DB(I).Type) & ")", vbNullString) & _
                        IIf(DB(I).Rank > 0, "\" & DB(I).Rank, vbNullString) & _
                        IIf(DB(I).Flags <> vbNullString, "\" & DB(I).Flags, vbNullString) & ", "
                    
                    ' increment found counter
                    found = (found + 1)
                End If
            End If
            
            ' reset booleans
            res = False
            blnChecked = False
        Next I

        If (found = 0) Then
            ' return message
            arrReturn(0) = "No such user(s) found."
        Else
            ' ...
            Call SplitByLen(Mid$(tmpbuf, 1, Len(tmpbuf) - Len(", ")), 180, arrReturn(), _
                "User(s) found: ", " [more]", ", ")
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.description & " in searchDatabase2()."
    
    Exit Function
End Function

Public Function RemoveItem(ByVal rItem As String, File As String, Optional ByVal dbType As String = vbNullString) As String
    
    Dim s()        As String
    Dim f          As Integer
    Dim Counter    As Integer
    Dim strCompare As String
    Dim strAdd     As String
    
    f = FreeFile
    
    If Dir$(GetFilePath(File & ".txt")) = vbNullString Then
        RemoveItem = "No %msgex% file found. Create one using .add, .addtag, or .shitlist."
        Exit Function
    End If
    
    Open (GetFilePath(File & ".txt")) For Input As #f
    If LOF(f) < 2 Then
        RemoveItem = "The %msgex% file is empty."
        Close #f
        Exit Function
    End If
    
    ReDim s(0)
    
    Do
        Line Input #f, strAdd
        s(UBound(s)) = strAdd
        ReDim Preserve s(0 To UBound(s) + 1)
    Loop Until EOF(f)
    
    Close #f
    
    For Counter = LBound(s) To UBound(s)
        strCompare = s(Counter)
        If strCompare <> vbNullString And strCompare <> " " Then
            If InStr(1, strCompare, " ", vbTextCompare) <> 0 Then
                strCompare = Left$(strCompare, InStr(1, strCompare, " ", vbTextCompare) - 1)
            End If
            
            If StrComp(LCase$(rItem), LCase$(strCompare), vbTextCompare) = 0 Then GoTo Successful
        End If
    Next Counter
    
    RemoveItem = "No such user found."
    
Successful:
    Close #f
    
    s(Counter) = vbNullString
    
    RemoveItem = "Successfully removed %msgex% " & Chr(34) & rItem & Chr(34) & "."
    
    Open (GetFilePath(File & ".txt")) For Output As #f
        For Counter = LBound(s) To UBound(s)
            If s(Counter) <> vbNullString And s(Counter) <> " " Then Print #f, s(Counter)
        Next Counter
theEnd:
    Close #f
End Function

Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER

    Dim I     As Integer ' ...
    Dim found As Boolean ' ...
    
    For I = LBound(DB) To UBound(DB)
        If (StrComp(DB(I).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean ' ...
        
            If (Len(dbType)) Then
                If (StrComp(DB(I).Type, dbType, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
            
            If (bln) Then
                found = True
                
                Exit For
            End If
        End If
        
        bln = False
    Next I
    
    If (found) Then
        Dim bak As udtDatabase ' ...
        
        Dim j   As Integer ' ...
        
        ' ...
        bak = DB(I)

        ' we aren't removing the last array
        ' element, are we?
        If (UBound(DB) = 0) Then
            ' redefine array size
            ReDim DB(0)
            
            ' ...
            With DB(0)
                .Username = vbNullString
                .Flags = vbNullString
                .Rank = 0
                .Groups = vbNullString
                .AddedBy = vbNullString
                .ModifiedBy = vbNullString
                .AddedOn = Now
                .ModifiedOn = Now
            End With
        Else
            ' ...
            For j = I To UBound(DB) - 1
                DB(j) = DB(j + 1)
            Next j
            
            ' redefine array size
            ReDim Preserve DB(UBound(DB) - 1)
    
            ' if we're removing a group, we need to also fix our
            ' group memberships, in case anything is broken now
            If (StrComp(bak.Type, "GROUP", vbBinaryCompare) = 0) Then
                Dim res As Boolean ' ...
           
                ' if we remove a user from the database during the
                ' execution of the inner loop, we have to reset our
                ' inner loop variables, otherwise we create errors
                ' due to incorrect database indexes.  Because of this,
                ' we have to dual-loop until our inner loop runs out
                ' of matching users.
                Do
                    ' reset loop variable
                    res = False
                
                    ' loop through database checking for users that
                    ' were members of the group that we just removed
                    For I = LBound(DB) To UBound(DB)
                        If (Len(DB(I).Groups) And DB(I).Groups <> "%") Then
                            If (InStr(1, DB(I).Groups, ",", vbBinaryCompare) <> 0) Then
                                Dim Splt()     As String ' ...
                                Dim innerfound As Boolean ' ...
                                
                                Splt() = Split(DB(I).Groups, ",")
                                
                                For j = LBound(Splt) To UBound(Splt)
                                    If (StrComp(bak.Username, Splt(j), vbTextCompare) = 0) Then
                                        innerfound = True
                                    
                                        Exit For
                                    End If
                                Next j
                            
                                If (innerfound) Then
                                    Dim k As Integer ' ...
                                    
                                    For k = (j + 1) To UBound(Splt)
                                        Splt(k - 1) = Splt(k)
                                    Next k
                                    
                                    ReDim Preserve Splt(UBound(Splt) - 1)
                                    
                                    DB(I).Groups = Join(Splt(), vbNullString)
                                End If
                            Else
                                If (StrComp(bak.Username, DB(I).Groups, vbTextCompare) = 0) Then
                                    res = DB_remove(DB(I).Username, DB(I).Type)
                                    
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                Loop While (res)
            End If
        End If
        
        ' commit modifications
        Call WriteDatabase(GetFilePath("users.txt"))
        
        DB_remove = True
        
        Exit Function
    End If
    
    DB_remove = False
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.description & " in DB_remove().")
    DB_remove = False
End Function

' requires public
Public Function GetSafelist(ByVal Username As String) As Boolean

    Dim I As Long ' ...
    
    If (bFlood = False) Then
        Dim gAcc As udtGetAccessResponse
        
        gAcc = GetCumulativeAccess(Username, "USER")
        
        If (Not InStr(1, gAcc.Flags, "S", vbBinaryCompare) = 0) Then
            GetSafelist = True
        ElseIf (gAcc.Rank >= 20) Then
            GetSafelist = True
        End If
    Else
        For I = 0 To (UBound(gFloodSafelist) - 1)
            If PrepareCheck(Username) Like gFloodSafelist(I) Then
                GetSafelist = True
                Exit For
            End If
        Next I
    End If
    
End Function

Public Function GetShitlist(ByVal Username As String) As String
    Dim gAcc As udtGetAccessResponse
    Dim Ban  As Boolean
    
    gAcc = GetCumulativeAccess(Username, "USER")
    
    If (Not InStr(1, gAcc.Flags, "Z", vbBinaryCompare) = 0) Then
        Ban = True
    ElseIf (Not InStr(1, gAcc.Flags, "B", vbBinaryCompare) = 0) Then
        If (GetSafelist(Username) = False) Then Ban = True
    End If
    
    If (Ban) Then
        If ((Len(gAcc.BanMessage) > 0) And (gAcc.BanMessage <> "%")) Then
            GetShitlist = Username & Space(1) & gAcc.BanMessage
        Else
            GetShitlist = Username & Space(1) & "Shitlisted"
        End If
    End If
End Function

' requires public
Public Function PrepareCheck(ByVal toCheck As String) As String
    toCheck = Replace(toCheck, "[", "ÿ")
    toCheck = Replace(toCheck, "]", "ö")
    toCheck = Replace(toCheck, "~", "Ü")
    toCheck = Replace(toCheck, "#", "¢")
    toCheck = Replace(toCheck, "-", "£")
    toCheck = Replace(toCheck, "&", "¥")
    toCheck = Replace(toCheck, "@", "¤")
    toCheck = Replace(toCheck, "{", "ƒ")
    toCheck = Replace(toCheck, "}", "á")
    toCheck = Replace(toCheck, "^", "í")
    toCheck = Replace(toCheck, "`", "ó")
    toCheck = Replace(toCheck, "_", "ú")
    toCheck = Replace(toCheck, "+", "ñ")
    toCheck = Replace(toCheck, "$", "÷")
    PrepareCheck = LCase$(toCheck)
End Function

' requires public
Public Function ReversePrepareCheck(ByVal toCheck As String) As String
    toCheck = Replace(toCheck, "ÿ", "[")
    toCheck = Replace(toCheck, "ö", "]")
    toCheck = Replace(toCheck, "Ü", "~")
    toCheck = Replace(toCheck, "¢", "#")
    toCheck = Replace(toCheck, "£", "-")
    toCheck = Replace(toCheck, "¥", "&")
    toCheck = Replace(toCheck, "¤", "@")
    toCheck = Replace(toCheck, "ƒ", "{")
    toCheck = Replace(toCheck, "á", "}")
    toCheck = Replace(toCheck, "í", "^")
    toCheck = Replace(toCheck, "ó", "`")
    toCheck = Replace(toCheck, "ú", "_")
    toCheck = Replace(toCheck, "ñ", "+")
    toCheck = Replace(toCheck, "÷", "$")
    ReversePrepareCheck = LCase$(toCheck)
End Function

' requires public
Public Sub LoadDatabase()
    On Error Resume Next

    Dim gA     As udtDatabase
    
    Dim s      As String
    Dim x()    As String
    Dim Path   As String
    Dim I      As Integer
    Dim f      As Integer
    Dim found  As Boolean
    Dim SaveDB As Boolean
    
    ReDim DB(0)
    Path = GetFilePath("users.txt")
    
    If Dir$(Path) <> vbNullString Then
        f = FreeFile
        Open Path For Input As #f
            
        If LOF(f) > 1 Then
            Do
                
                Line Input #f, s
                
                If InStr(1, s, " ", vbTextCompare) > 0 Then
                    x() = Split(s, " ", 10)
                    
                    If UBound(x) > 0 Then
                        ReDim Preserve DB(I)
                        
                        With DB(I)
                            .Username = x(0)
                            
                            .Rank = 0
                            .AddedOn = Now
                            .AddedBy = "2.6r3Import"
                            .BanMessage = vbNullString
                            .Flags = vbNullString
                            .Groups = vbNullString
                            .ModifiedBy = "2.6r3Import"
                            .ModifiedOn = Now
                            .Type = "USER"
                            
                            If StrictIsNumeric(x(1)) Then
                                .Rank = Val(x(1))
                            Else
                                If x(1) <> "%" Then
                                    .Flags = x(1)
                                    
                                    'If InStr(X(1), "S") > 0 Then
                                    '    AddToSafelist .Name
                                    '    .Flags = Replace(.Flags, "S", "")
                                    'End If
                                End If
                            End If
                            
                            If UBound(x) > 1 Then
                                If StrictIsNumeric(x(2)) Then
                                    .Rank = Int(x(2))
                                Else
                                    If x(2) <> "%" Then
                                        .Flags = x(2)
                                    End If
                                End If
                                
                                '  0        1       2       3      4        5          6       7     8
                                ' username access flags addedby addedon modifiedby modifiedon type banmsg
                                If UBound(x) > 2 Then
                                    .AddedBy = x(3)
                                    
                                    If UBound(x) > 3 Then
                                        .AddedOn = CDate(Replace(x(4), "_", " "))
                                        
                                        If UBound(x) > 4 Then
                                            .ModifiedBy = x(5)
                                            
                                            If UBound(x) > 5 Then
                                                .ModifiedOn = CDate(Replace(x(6), "_", " "))

                                                If UBound(x) > 6 Then
                                                    .Type = x(7)
                                                    
                                                    If UBound(x) > 7 Then
                                                        .Groups = x(8)
                                                        
                                                        If UBound(x) > 8 Then
                                                            .BanMessage = x(9)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                
                            End If
                            
                            If .Rank > 200 Then
                                .Rank = 200
                            End If
                            
                            If .Type = "" Or .Type = "%" Then
                                .Type = "USER"
                            End If
                            SaveDB = (.AddedOn = Now) Or SaveDB
                        End With

                        I = I + 1
                    End If
                End If
                
            Loop While Not EOF(f)
        End If
        
        Close #f
    End If

    ' 9/13/06: Add the bot owner 200
    If (LenB(BotVars.BotOwner) > 0) Then
        For I = 0 To UBound(DB)
            If (StrComp(DB(I).Username, BotVars.BotOwner, vbTextCompare) = 0) Then
                found = True
                
                Exit For
            End If
        Next I
        
        If (found = False) Then
            If (UBound(DB)) Then
                ReDim Preserve DB(UBound(DB) + 1)
            End If
            
            With DB(UBound(DB))
                .Username = BotVars.BotOwner
                .Type = "USER"
                .Rank = 200
                .AddedBy = "(console)"
                .AddedOn = Now
                .ModifiedBy = "(console)"
                .ModifiedOn = Now
            End With
            SaveDB = True
        End If
    End If
    
    If (SaveDB) Then Call WriteDatabase(Path)
End Sub

'08/15/09 - Hdx - Converted to use clsCommandObj to check if the command has valid syntax
'Removed outBuff ... What use was it?
Public Function IsCorrectSyntax(ByVal commandName As String, ByVal commandArgs As String, Optional scriptOwner As String = vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim docs As clsCommandDocObj
    Dim Command As New clsCommandObj
    
    Set docs = OpenCommand(commandName, scriptOwner)
    If (docs Is Nothing) Then
        Set docs = OpenCommand(convertAlias(commandName), scriptOwner)
        If (docs Is Nothing) Then
            IsCorrectSyntax = False
            Exit Function
        End If
    End If
    
    With Command
        .Name = docs.Name
        .Args = commandArgs
        IsCorrectSyntax = .IsValid
    End With
    
    Set docs = Nothing
    Set Command = Nothing
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in IsCorrectSyntax().")
End Function

'08/15/09 - Hdx - Converted to use clsCommandObj to check if the user has enough access
'Removed outBuff ... What use was it?
Public Function HasAccess(ByVal Username As String, ByVal commandName As String, Optional ByVal commandArgs As _
    String = vbNullString, Optional scriptOwner As String = vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim docs As clsCommandDocObj
    Dim Command As clsCommandObj
    
    Set docs = OpenCommand(commandName, scriptOwner)
    If (docs Is Nothing) Then
        Set docs = OpenCommand(convertAlias(commandName), scriptOwner)
        If (docs Is Nothing) Then
            HasAccess = False
            Exit Function
        End If
    End If
    
    With Command
        .Name = docs.Name
        .Args = commandArgs
        .Username = Username
        HasAccess = .HasAccess
    End With
    
    Set docs = Nothing
    Set Command = Nothing
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.description & " in HasAccess()."
End Function

' Writes database to disk
' Updated 9/13/06 for new features
Public Sub WriteDatabase(ByVal U As String)
    Dim f As Integer
    Dim I As Integer
    
    On Error GoTo WriteDatabase_Exit

    f = FreeFile
    
    Open U For Output As #f
        For I = LBound(DB) To UBound(DB)
            ' ...
            If (LenB(DB(I).Username) > 0) Then
                Print #f, DB(I).Username;
                Print #f, " " & DB(I).Rank;
                Print #f, " " & IIf(Len(DB(I).Flags) > 0, DB(I).Flags, "%");
                Print #f, " " & IIf(Len(DB(I).AddedBy) > 0, DB(I).AddedBy, "%");
                Print #f, " " & IIf(DB(I).AddedOn > 0, DateCleanup(DB(I).AddedOn), "%");
                Print #f, " " & IIf(Len(DB(I).ModifiedBy) > 0, DB(I).ModifiedBy, "%");
                Print #f, " " & IIf(DB(I).ModifiedOn > 0, DateCleanup(DB(I).ModifiedOn), "%");
                Print #f, " " & IIf(Len(DB(I).Type) > 0, DB(I).Type, "USER");
                Print #f, " " & IIf(Len(DB(I).Groups) > 0, DB(I).Groups, "%");
                Print #f, " " & IIf(Len(DB(I).BanMessage) > 0, DB(I).BanMessage, "%");
                Print #f, vbCr
            End If
        Next I

WriteDatabase_Exit:
    Close #f
    
    Exit Sub

WriteDatabase_Error:

    Debug.Print "Error #" & Err.Number & " (" & Err.description & ") in procedure " & _
        "WriteDatabase of Module modCommandCode"
    
    Resume WriteDatabase_Exit
End Sub

' requires public
Private Function DateCleanup(ByVal TDate As Date) As String
    Dim T As String
    
    T = Format(TDate, "dd-MM-yyyy_HH:MM:SS")
    DateCleanup = Replace(T, " ", "_")
End Function

Private Function GetAccessINIValue(ByVal sKey As String, Optional ByVal Default As Long) As Long
    Dim s As String, L As Long
    
    s = ReadINI("Numeric", sKey, "access.ini")
    L = Val(s)
    
    If L > 0 Then
        GetAccessINIValue = L
    Else
        If Default > 0 Then
            GetAccessINIValue = Default
        Else
            GetAccessINIValue = 100
        End If
    End If
End Function

Private Function CheckUser(ByVal user As String, Optional ByVal _
    allow_illegal As Boolean = False) As Boolean
    
    Dim I       As Integer ' ...
    Dim bln     As Boolean ' ...
    Dim illegal As Boolean ' ...
    Dim invalid As Boolean ' ...
    
    ' ...
    If (Left$(user, 1) = "*") Then
        user = Mid$(user, 2)
    End If
    
    ' ...
    user = Replace(user, "@USWest", vbNullString, 1)
    user = Replace(user, "@USEast", vbNullString, 1)
    user = Replace(user, "@Asia", vbNullString, 1)
    user = Replace(user, "@Europe", vbNullString, 1)
    
    ' ...
    user = Replace(user, "@Lordaeron", vbNullString, 1)
    user = Replace(user, "@Azeroth", vbNullString, 1)
    user = Replace(user, "@Kalimdor", vbNullString, 1)
    user = Replace(user, "@Northrend", vbNullString, 1)
    
    If (Len(user) = 0) Then
        invalid = True
    ElseIf (Len(user) > 15) Then
        invalid = True
    Else
        ' 95 (a)
        ' 65 (A)
        ' 48 (0)
        ' 57 (9)
    
        For I = 1 To Len(user)
            Dim currentCharacter As String
            currentCharacter = Mid$(user, I, 1)

            ' is the character between A-Z or a-z?
            If (Asc(currentCharacter) < Asc("A")) Or (Asc(currentCharacter) > Asc("z")) Then
                ' is the character between 0 - 9?
                If ((Asc(currentCharacter) < Asc("0")) Or (Asc(currentCharacter) > Asc("9"))) Then
                
                    ' !@$(){}[]=+`~^-’.:;_|
                    ' is the character a valid special character?
                    If ((Asc(currentCharacter) = Asc("[")) Or _
                        (Asc(currentCharacter) = Asc("]")) Or _
                        (Asc(currentCharacter) = Asc("(")) Or _
                        (Asc(currentCharacter) = Asc(")")) Or _
                        (Asc(currentCharacter) = Asc(".")) Or _
                        (Asc(currentCharacter) = Asc("-")) Or _
                        (Asc(currentCharacter) = Asc("_"))) Then
                    
                        If (bln) Then
                            illegal = True
                        Else
                            bln = True
                        End If
                    Else
                        ' check for illegal characters, and for
                        ' characters that have always been invalid
                        Select Case (Asc(currentCharacter))
                            Case Asc("{"): illegal = True
                            Case Asc("}"): illegal = True
                            Case Asc("="): illegal = True
                            Case Asc("+"): illegal = True
                            Case Asc("`"): illegal = True
                            Case Asc("~"): illegal = True
                            Case Asc("^"): illegal = True
                            Case Asc(":"): illegal = True
                            Case Asc(";"): illegal = True
                            Case Asc("|"): illegal = True
                            Case Asc("@"): illegal = True
                            Case Asc("$"): illegal = True
                            Case Asc("!"): illegal = True
                            Case Asc("#"): illegal = True
                            Case Else:
                                invalid = True
                                
                                ' break loop
                                Exit For
                        End Select
                    End If
                End If
            End If
        Next I
    End If
    
    ' is our user valid?
    If (Not (invalid)) Then
        ' does our user contain illegal characters?
        If (illegal) Then
            ' do we allow illegal characters?
            If (allow_illegal) Then
                CheckUser = True
            Else
                CheckUser = False
            End If
        Else
            CheckUser = True
        End If
    Else
        CheckUser = False
    End If
End Function

Public Function usingGameConventions() As Boolean
    Select Case (StrReverse$(BotVars.Product))
        Case "D2DV", "D2XP"
            If ((BotVars.UseGameConventions = True) And ((BotVars.UseD2GameConventions = True))) Then
                usingGameConventions = True
            End If
            
        Case "WAR3", "W3XP"
            If ((BotVars.UseGameConventions = True)) And ((BotVars.UseW3GameConventions = True)) Then
                usingGameConventions = True
            End If
    End Select
End Function

Public Function convertUsername(ByVal Username As String) As String
    Dim Index As Long ' ...
    
    If (Len(Username) < 1) Then
        convertUsername = Username
    Else
        If ((StrReverse$(BotVars.Product) = "D2DV") Or _
                (StrReverse$(BotVars.Product) = "D2XP")) Then
            
            If ((BotVars.UseGameConventions = False) Or _
                    ((BotVars.UseD2GameConventions = False))) Then
               
                Index = InStr(1, Username, "*", vbBinaryCompare)
            
                If (Index > 0) Then
                    convertUsername = Mid$(Username, Index + 1)
                Else
                    convertUsername = Username
                End If
            Else
                Index = InStr(1, Username, "*", vbBinaryCompare)
            
                If (Index > 1) Then
                    convertUsername = _
                        Left$(Username, Index - 1) & " (*" & Mid$(Username, Index + 1) & ")"
                Else
                    If (Index = 0) Then
                        convertUsername = "*" & Username
                    End If
                End If
            End If
            
        ElseIf ((StrReverse$(BotVars.Product) = "WAR3") Or _
                    (StrReverse$(BotVars.Product) = "W3XP")) Then
                
            If ((BotVars.UseGameConventions = False)) Or _
                    ((BotVars.UseW3GameConventions = False)) Then
    
                If (BotVars.Gateway <> vbNullString) Then
                    Select Case (BotVars.Gateway)
                        Case "Lordaeron": Index = InStr(1, Username, "@USWest", vbTextCompare)
                        Case "Azeroth":   Index = InStr(1, Username, "@USEast", vbTextCompare)
                        Case "Kalimdor":  Index = InStr(1, Username, "@Asia", vbTextCompare)
                        Case "Northrend": Index = InStr(1, Username, "@Europe", vbTextCompare)
                    End Select
                    
                    If (Index > 1) Then
                        convertUsername = Left$(Username, Index - 1)
                    Else
                        convertUsername = Username & "@" & BotVars.Gateway
                    End If
                End If
            End If
        End If
        
        If (convertUsername = vbNullString) Then
            convertUsername = Username
        End If
    End If
End Function

Public Function reverseUsername(ByVal Username As String) As String
    Dim Index As Long ' ...
    
    If (Len(Username) < 1) Then
        Exit Function
    End If

    If ((StrReverse$(BotVars.Product) = "D2DV") Or _
        (StrReverse$(BotVars.Product) = "D2XP")) Then
        
        If ((BotVars.UseGameConventions = False)) Or _
                ((BotVars.UseD2GameConventions = False)) Then
            
            ' With reverseUsername() now being called from AddQ(), usernames
            ' in procedures called prior to AddQ() will no longer require
            ' prefixes; however, we want to ensure that a '*' was not already
            ' specified before conversion to allow older scripts and procedures
            ' to continue functioning correctly.  This check may be removed in
            ' future releases.
            If (Left$(Username, 1) <> "*") Then
                reverseUsername = ("*" & Username)
            End If
        End If
    ElseIf ((StrReverse$(BotVars.Product) = "WAR3") Or _
            (StrReverse$(BotVars.Product) = "W3XP")) Then
            
        If ((BotVars.UseGameConventions = False)) Or _
                ((BotVars.UseW3GameConventions = False)) Then
            
            If (BotVars.Gateway <> vbNullString) Then
                Index = InStr(1, Username, ("@" & BotVars.Gateway), vbTextCompare)
    
                If (Index > 0) Then
                    reverseUsername = Left$(Username, Index - 1)
                Else
                    Select Case (BotVars.Gateway)
                        Case "Lordaeron": reverseUsername = Username & "@USWest"
                        Case "Azeroth":   reverseUsername = Username & "@USEast"
                        Case "Kalimdor":  reverseUsername = Username & "@Asia"
                        Case "Northrend": reverseUsername = Username & "@Europe"
                        Case Else:        reverseUsername = Username
                    End Select
                End If
            End If
        End If
    End If
    
    If (reverseUsername = vbNullString) Then
        reverseUsername = Username
    End If
End Function
