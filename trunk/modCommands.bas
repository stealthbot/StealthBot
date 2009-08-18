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

' *******************************************************************************
' * This module, or any related functions outside of this module, should not be *
' * modified without consultation prior to the modifications.                   *
' *******************************************************************************

Option Explicit

' Winamp Constants
'Private Const WA_PREVTRACK   As Long = 40044 ' ...
'Private Const WA_NEXTTRACK   As Long = 40048 ' ...
'Private Const WA_PLAY        As Long = 40045 ' ...
'Private Const WA_PAUSE       As Long = 40046 ' ...
'Private Const WA_STOP        As Long = 40047 ' ...
'Private Const WA_FADEOUTSTOP As Long = 40147 ' ...

'Private m_dbAccess     As udtGetAccessResponse
'Private m_username     As String  ' ...
'Private m_IsLocal      As Boolean ' ...
Private m_waswhispered  As Boolean ' ...
Private m_DisplayOutput As Boolean ' ...

Public flood    As String ' ...?
Public floodCap As Byte   ' ...?

' prepares commands for processing, and calls helper functions associated with
' processing
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
    
    ' ...
    If ((IsLocal) And (Left$(Message, 3) = "///")) Then
        ' ...
        AddQ Mid$(Message, 3)
        
        ' ...
        Exit Function
    End If

    ' 08/17/2009 - 52 - using static class method
    Set commands = clsCommandObj.IsCommand(Message, IIf(IsLocal, modGlobals.CurrentUsername, Username), Chr$(0))

    For Each Command In commands
        m_DisplayOutput = Command.PublicOutput
        
        ' ...
        If (Command.HasAccess) Then
            ' ...
            Command.WasWhispered = WasWhispered
            If (IsLocal) Then
                With dbAccess
                    .Rank = 201
                    .Flags = "A"
                End With
            Else
                dbAccess = GetCumulativeAccess(Username)
            End If
            
            ' ...
            If (LenB(Command.docs.Owner) = 0) Then 'Is it a built in command?
                If (Not executeCommand(Username, dbAccess, Command.Name & Space$(1) & Command.Args, IsLocal, command_return)) Then
                    Call DispatchCommand(Command)
                    Call RunInSingle(Nothing, "Event_Command", Command)
                    Command.SendResponse
                End If
            Else
                Call RunInSingle(modScripting.GetModuleByName(Command.docs.Owner), "Event_Command", Command)
                Command.SendResponse
            End If
                    
            ' ...
            If (DisplayOutput) Then
            
                'Ignore code as bot is closing
                If BotIsClosing Then
                    Exit Function
                End If
            
                ' ...
                If (command_return(0) <> vbNullString) Then
                    ' ...
                    For I = LBound(command_return) To UBound(command_return)
                        ' ...
                        If (IsLocal) Then
                            ' ...
                            If (Command.PublicOutput) Then
                                AddQ command_return(I), PRIORITY.CONSOLE_MESSAGE
                            Else
                                frmChat.AddChat RTBColors.ConsoleText, command_return(I)
                            End If
                        Else
                            ' ...
                            If ((BotVars.WhisperCmds) Or (WasWhispered)) Then
                                AddQ "/w " & Username & Space$(1) & command_return(I), _
                                        PRIORITY.COMMAND_RESPONSE_MESSAGE
                            Else
                                AddQ command_return(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
                            End If
                        End If
                    Next I
                End If
            End If
        End If
        
        ' ...
        execCommand = False
    Next
        
    ' ...
    If (IsLocal) Then
        ' ...
        If ((bln = False) And (commands.Count = 0)) Then
            AddQ Message
        End If
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

    'Unload memory - FrOzeN
    Set Command = Nothing

    ' return command failure result
    ProcessCommand = False
    
    Exit Function
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
        
        Case Else: DispatchCommand = False
    End Select
End Function

' command processing helper function
Public Function executeCommand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal Message As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpmsg   As String  ' stores copy of message
    Dim cmdName  As String  ' stores command name
    Dim msgData  As String  ' stores unparsed command parameters
    Dim blnNoCmd As Boolean ' stores result of command switch (true = no command found)
    'Dim I        As Integer ' loop counter
    
    ' create single command data array element for safe bounds checking
    ' and to help aide in a reduction of command function overhead
    ReDim Preserve cmdRet(0)
    
    ' store local copy of message
    tmpmsg = Message

    ' grab command name & message data
    If (InStr(1, tmpmsg, Space(1), vbBinaryCompare) <> 0) Then
        ' grab command name
        cmdName = Left$(tmpmsg, (InStr(1, tmpmsg, Space(1), _
            vbBinaryCompare) - 1))
        
        ' remove command name (and space) from message
        tmpmsg = Mid$(tmpmsg, Len(cmdName) + 2)
        
        ' grab message data
        msgData = tmpmsg
    Else
        ' grab command name
        cmdName = tmpmsg
    End If

    ' ...
    LogCommand IIf(InBot, vbNullString, Username), Message
            
    ' command switch
    Select Case (cmdName)
        Case "dump":          Call OnDump(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quit":          Call OnQuit(Username, dbAccess, msgData, InBot, cmdRet())
        Case "locktext":      Call OnLockText(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "efp":           Call OnEfp(Username, dbAccess, msgData, InBot, cmdRet())
        Case "peonban":       Call OnPeonBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quiettime":     Call OnQuietTime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "roll":          Call OnRoll(Username, dbAccess, msgData, InBot, cmdRet())
        Case "sweepban":      Call OnSweepBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "sweepignore":   Call OnSweepIgnore(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setname":       Call OnSetName(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setpass":       Call OnSetPass(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setkey":        Call OnSetKey(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setexpkey":     Call OnSetExpKey(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setserver":     Call OnSetServer(Username, dbAccess, msgData, InBot, cmdRet())
        Case "giveup":        Call OnGiveUp(Username, dbAccess, msgData, InBot, cmdRet())
        Case "math":          Call OnMath(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idlebans":      Call OnIdleBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "chpw":          Call OnChPw(Username, dbAccess, msgData, InBot, cmdRet())
        Case "sethome":       Call OnSetHome(Username, dbAccess, msgData, InBot, cmdRet())
        Case "resign":        Call OnResign(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clearbanlist":  Call OnClearBanList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "kickonyell":    Call OnKickOnYell(Username, dbAccess, msgData, InBot, cmdRet())
        Case "plugban":       Call OnPlugBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clientbans":    Call OnClientBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cadd":          Call OnCAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cdel":          Call OnCDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banned":        Call OnBanned(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ipbans":        Call OnIPBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ipban":         Call OnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unipban":       Call OnUnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "designate":     Call OnDesignate(Username, dbAccess, msgData, InBot, cmdRet())
        Case "protect":       Call OnProtect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "whispercmds":   Call OnWhisperCmds(Username, dbAccess, msgData, InBot, cmdRet())
        Case "rem":           Call OnRem(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletime":      Call OnIdleTime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idle":          Call OnIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitdel":       Call OnShitDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safedel":       Call OnSafeDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagdel":        Call OnTagDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setidle":       Call OnSetIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletype":      Call OnIdleType(Username, dbAccess, msgData, InBot, cmdRet())
        Case "settrigger":    Call OnSetTrigger(Username, dbAccess, msgData, InBot, cmdRet())
        Case "levelban":      Call OnLevelBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "d2levelban":    Call OnD2LevelBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "phrasebans":    Call OnPhraseBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "pon":           Call OnPhraseBans(Username, dbAccess, "on", InBot, cmdRet())
        Case "poff":          Call OnPhraseBans(Username, dbAccess, "off", InBot, cmdRet())
        Case "pstatus":       Call OnPhraseBans(Username, dbAccess, vbNullString, InBot, cmdRet())
        Case "setpmsg":       Call OnSetPMsg(Username, dbAccess, msgData, InBot, cmdRet())
        Case "phrases":       Call OnPhrases(Username, dbAccess, msgData, InBot, cmdRet())
        Case "addphrase":     Call OnAddPhrase(Username, dbAccess, msgData, InBot, cmdRet())
        Case "delphrase":     Call OnDelPhrase(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagadd":        Call OnTagAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safelist":      Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safeadd":       Call OnSafeAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safecheck":     Call OnSafeCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "exile":         Call OnExile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unexile":       Call OnUnExile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitlist":      Call OnShitList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagbans":       Call OnTagBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitadd":       Call OnShitAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "bancount":      Call OnBanCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banlistcount":  Call OnBanListCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagcheck":      Call OnTagCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "slcheck":       Call OnSLCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "readfile":      Call OnReadFile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "greet":         Call OnGreet(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ban":           Call OnBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unban":         Call OnUnban(Username, dbAccess, msgData, InBot, cmdRet())
        Case "kick":          Call OnKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "voteban":       Call OnVoteBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "votekick":      Call OnVoteKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "vote":          Call OnVote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tally":         Call OnTally(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cancel":        Call OnCancel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "addquote":      Call OnAddQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quote":         Call OnQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "checkmail":     Call OnCheckMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "inbox":         Call OnInbox(Username, dbAccess, msgData, InBot, cmdRet())
        Case "add":           Call OnAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "mmail":         Call OnMMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "bmail":         Call OnBMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "designated":    Call OnDesignated(Username, dbAccess, msgData, InBot, cmdRet())
        Case "flip":          Call OnFlip(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clear":         Call OnClear(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "monitor":       Call OnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "unmonitor":     Call OnUnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "online":        Call OnOnline(Username, dbAccess, msgData, InBot, cmdRet())
        Case "enable":        Call OnEnable(Username, dbAccess, msgData, InBot, cmdRet())
        Case "disable":       Call OnDisable(Username, dbAccess, msgData, InBot, cmdRet())
        Case "exec":          Call OnExec(Username, dbAccess, msgData, InBot, cmdRet())
        Case Else
            blnNoCmd = True
    End Select
    
    ' is the bot unloading?
    If (BotIsClosing) Then
        Exit Function
    End If
    
    ' was a command found? return.
    executeCommand = (Not (blnNoCmd))
End Function

' handle dump command
Private Function OnDump(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Call DumpPacketCache
    
End Function ' end function OnDump

' handle quit command
Private Function OnQuit(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will initiate the bot's termination sequence.
    
    BotIsClosing = True
    
    Unload frmChat
    Set frmChat = Nothing
End Function ' end function OnQuit

' handle locktext command
Private Function OnLockText(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will prevent chat messages from displaying on the bot's screen.
    
    Call frmChat.mnuLock_Click
End Function ' end function OnLockText

' handle efp command
'Private Function OnEfp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'    ' This command will enable, disable, or check the status of, the Effective Floodbot
'    ' Protection system.  EFP is a system designed to combat floodbot attacks by reducing
'    ' the number of allowable commands, and enhancing the strength of the message queue.
'
'    Dim tmpbuf As String ' temporary output buffer
'
'    ' ...
'    msgData = LCase$(msgData)
'
'    If (msgData = "on") Then
'        ' enable efp
'        Call frmChat.SetFloodbotMode(1)
'
'        tmpbuf = "Emergency floodbot protection enabled."
'    ElseIf (msgData = "off") Then
'        ' disable efp
'        Call frmChat.SetFloodbotMode(0)
'
'        tmpbuf = "Emergency floodbot protection disabled."
'    ElseIf (msgData = "status") Then
'        If (bFlood) Then
'            frmChat.AddChat RTBColors.TalkBotUsername, "Emergency floodbot protection is " & _
'                "enabled. (No messages can be sent to battle.net.)"
'        Else
'            tmpbuf = "Emergency floodbot protection is disabled."
'        End If
'    End If
'
'    ' return message
'    cmdRet(0) = tmpbuf
'End Function ' end function OnEfp

' handle peonban command
Private Function OnPeonBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable, or check the status of, WarCraft III peon
    ' banning.  The "Peon" class is defined by Battle.net, and is currently the lowest
    ' ranking WarCraft III user classification, in which, users have less than twenty-five
    ' wins on record for any given race.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)
    
    If (msgData = "on") Then
        ' enable peon banning
        BotVars.BanPeons = 1
        
        ' write configuration entry
        Call WriteINI("Other", "PeonBans", "Y")
        
        tmpbuf = "Peon banning activated."
    ElseIf (msgData = "off") Then
        ' disable peon banning
        BotVars.BanPeons = 0
        
        ' write configuration entry
        Call WriteINI("Other", "PeonBans", "N")
        
        tmpbuf = "Peon banning deactivated."
    ElseIf (msgData = "status") Then
        tmpbuf = "The bot is currently "
        
        If (BotVars.BanPeons = 0) Then
            tmpbuf = tmpbuf & "not banning peons."
        Else
            tmpbuf = tmpbuf & "banning peons."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPeonBan

' handle quiettime command
Private Function OnQuietTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable or check the status of, quiet time.
    ' Quiet time is a feature that will ban non-safelisted users from the
    ' channel when they speak publicly within the channel.  This is useful
    ' when a channel wishes to have a discussion while allowing public
    ' attendance, but disallowing public participation.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)
    
    If (msgData = "on") Then
        ' enable quiettime
        BotVars.QuietTime = True
    
        ' write configuration entry
        Call WriteINI("Main", "QuietTime", "Y")
        
        tmpbuf = "Quiet-time enabled."
    ElseIf (msgData = "off") Then
        ' disable quiettime
        BotVars.QuietTime = False
        
        ' write configuration entry
        Call WriteINI("Main", "QuietTime", "N")
        
        tmpbuf = "Quiet-time disabled."
    ElseIf (msgData = "status") Then
        If (BotVars.QuietTime) Then
            tmpbuf = "Quiet-time is currently enabled."
        Else
            tmpbuf = "Quiet-time is currently disabled."
        End If
    Else
        tmpbuf = "Error: Invalid arguments."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnQuietTime

' handle roll command
Private Function OnRoll(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will state a random number between a range of zero to
    ' one-hundred, or from zero to any specified number.
    
    Dim tmpbuf  As String ' temporary output buffer
    Dim iWinamp As Long
    Dim Track   As Long

    If (Len(msgData) = 0) Then
        Randomize
        
        iWinamp = CLng(Rnd * 100)
        
        tmpbuf = "Random number (0-100): " & iWinamp
    Else
        Randomize
        
        If (StrictIsNumeric(msgData)) Then
            If (Val(msgData) < 100000000) Then
                Track = CLng(Rnd * CLng(msgData))
                
                tmpbuf = "Random number (0-" & msgData & "): " & Track
            Else
                tmpbuf = "Invalid value."
            End If
        Else
            tmpbuf = "Error: Invalid value."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnRoll

' handle sweepban command
Private Function OnSweepBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will grab the listing of users in the specified channel
    ' using Battle.net's "who" command, and will then begin banning each
    ' user from the current channel using Battle.net's "ban" command.
    
    Dim U      As String ' ...
    Dim Y      As String ' ...
    Dim tmpbuf As String ' ...

    ' ...
    If (g_Channel.Self.IsOperator) Then
        ' ...
        Call cache(vbNullString, 255, "ban ")
        
        ' ...
        Call AddQ("/who " & msgData, PRIORITY.CHANNEL_MODERATION_MESSAGE, _
            Username, "request_receipt")
    Else
        ' ...
        tmpbuf = "Error: The bot is not currently a channel operator."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSweepBan

' handle sweepignore command
Private Function OnSweepIgnore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will grab the listing of users in the specified channel
    ' using Battle.net's "who" command, and will then begin ignoring each
    ' user using Battle.net's "squelch" command.  This command is often used
    ' instead of "sweepban" to temporarily ban all users on a given ip address
    ' without actually immediately banning them from the channel.  This is
    ' useful if a user wishes to stay below Battle.net's limitations on bans,
    ' and still prevent a number of users from joining the channel for a
    ' temporary amount of time.
    
    Dim U      As String ' ...
    Dim Y      As String ' ...
    Dim tmpbuf As String ' ...

    ' ...
    Call cache(vbNullString, 255, "squelch ")
    
    ' ...
    Call AddQ("/who " & msgData, PRIORITY.CHANNEL_MODERATION_MESSAGE, _
        Username, "request_receipt")
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSweepIgnore

' handle setname command
Private Function OnSetName(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the username that the bot uses to connect with
    ' to the specified value.
    
    Dim tmpbuf As String ' temporary output buffer

    ' are we using a beta?
    '#If BETA = 1 Then
    '    ' only allow use of setname command while on-line to prevent beta
    '    ' authorization bypassing
    '    If ((Not (g_Online = True)) Or (g_Connected = False)) Then
    '        Exit Function
    '    End If
    '#End If

    ' write configuration entry
    Call WriteINI("Main", "Username", msgData)
    
    ' set username
    BotVars.Username = msgData
    
    tmpbuf = "New username set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetName

' handle setpass command
Private Function OnSetPass(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the password that the bot uses to connect with
    ' to the specified value.
    
    Dim tmpbuf As String ' temporary output buffer

    ' write configuration entry
    Call WriteINI("Main", "Password", msgData)
    
    ' set password
    BotVars.Password = msgData
    
    tmpbuf = "New password set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetPass

' handle math command
Private Function OnMath(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will execute a specified mathematical statement using the
    ' restricted script control, SCRestricted, on frmChat.  The execution
    ' of any code through direct user-interaction can become quite error-prone
    ' and, as such, this command requires its own error handler.  The input
    ' of this command must also be properly sanitized to ensure that no
    ' harmful statements are inadvertently allowed to launch on the user's
    ' machine.
    
    ' default error handler for math command
    On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf As String ' temporary output buffer

    If (Len(msgData)) Then
        If (InStr(1, msgData, "CreateObject", vbTextCompare)) Then
            ' CreateObject() is a no no, because of
            ' its use in exploits.
            
            tmpbuf = "Evaluation error."
        Else
            Dim res As String ' stores result of Eval()
        
            ' disable dangerous exploits
            With frmChat.SCRestricted
                .AllowUI = False
                .UseSafeSubset = True
            End With
            
            ' keep the bot from echoing specified phrases
            msgData = Replace(msgData, Chr$(34), vbNullString)
            
            ' evaluate expression
            res = frmChat.SCRestricted.Eval(msgData)
            
            ' check for scripting object errors
            If (res <> vbNullString) Then
                tmpbuf = "The statement " & Chr$(34) & _
                    msgData & Chr$(34) & " evaluates to: " & _
                        res & "."
            Else
                tmpbuf = "Evaluation error."
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
    
    Exit Function
    
ERROR_HANDLER:
    tmpbuf = "Evaluation error."

    ' return message
    cmdRet(0) = tmpbuf
    
    Exit Function
End Function ' end function OnMath

' handle setkey command
Private Function OnSetKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the bot's CD-Key to the CD-Key specified.
    
    Dim tmpbuf As String ' temporary output buffer

    ' clean data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)

    ' write configuration information
    Call WriteINI("Main", "CDKey", msgData)
    
    ' set CD-Key
    BotVars.CDKey = msgData
    
    tmpbuf = "New cdkey set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetKey

' handle setexpkey command
Private Function OnSetExpKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the bot's expansion CD-Key to the expansion
    ' CD-Key specified.
    
    Dim tmpbuf As String ' temporary output buffer

    ' sanitize data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)
    
    ' write configuration entry
    Call WriteINI("Main", "ExpKey", msgData)
    
    ' set expansion CD-Key
    BotVars.ExpKey = msgData
    
    tmpbuf = "New expansion CD-key set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetExpKey

' handle setserver command
Private Function OnSetServer(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the server that the bot connects to to the value
    ' specified.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        cmdRet(0) = "Error: The specified server contains " & _
            "invalid character(s)."
    
        Exit Function
    End If

    ' write configuration information
    Call WriteINI("Main", "Server", msgData)
    
    ' set server
    BotVars.Server = msgData
    
    tmpbuf = "New server set."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetServer

' handle giveup command
Private Function OnGiveUp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will allow a user to designate a specified user using
    ' Battle.net's "designate" command, and will then make the bot resign
    ' its status as a channel moderator.  This command is useful if you are
    ' lazy and you just wish to designate someone as quickly as possible.
    
    ' ...
    If (g_Channel.GetUserIndex(msgData) > 0) Then
        Dim I          As Integer ' ...
        Dim userCount  As Integer ' ...
        Dim opsCount   As Integer ' ...
        Dim arrUsers() As String ' ...
    
        ' ...
        If (g_Channel.Self.IsOperator = False) Then
            ' ...
            cmdRet(0) = "Error: This command requires channel operator status."
        
            ' ...
            Exit Function
        End If
        
        ' ...
        For I = 1 To g_Channel.Users.Count
            ' ...
            If (StrComp(g_Channel.Users(I).DisplayName, GetCurrentUsername, vbBinaryCompare) <> 0) Then
                ' ...
                If (g_Channel.Users(I).IsOperator) Then
                    ' ...
                    opsCount = (opsCount + 1)
                End If
            End If
        Next I
    
        ' ...
        If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
            ' ...
            ReDim Preserve arrUsers(0)
            
            ' ...
            If (g_Clan.Self.Rank >= 4) Then
                ' ...
                frmChat.cboSend.Text = vbNullString
            
                ' ...
                For I = 1 To g_Clan.Shamans.Count
                    ' ...
                    If (g_Channel.GetUserIndexEx(g_Clan.Shamans(I).Name) > 0) Then
                        ' ...
                        arrUsers(userCount) = g_Clan.Shamans(I).Name
    
                        ' ...
                        userCount = (userCount + 1)
    
                        ' ...
                        ReDim Preserve arrUsers(0 To userCount)
                    End If
                Next I
                
                ' ...
                If (opsCount > userCount) Then
                    ' ...
                    cmdRet(0) = "Error: There is currently a channel moderator present that cannot be " & _
                            "removed from his or her position."
                        
                    Exit Function
                End If
                
                ' ...
                If (userCount > 0) Then
                    ' demote shamans
                    For I = 0 To (userCount - 1)
                        ' ...
                        g_Clan.Members(g_Clan.GetUserIndexEx(arrUsers(I))).Demote
                        
                        ' ...
                        'frmChat.AddChat vbRed, _
                        '    "DEBUG: DEMOTE " & g_Clan.Members(g_Clan.GetUserIndexEx(arrUsers(I))).DisplayName
    
                        ' ...
                        'Call Pause(200, True, True)
                    Next I
                End If
                
                ' ...
                opsCount = 0
                
                ' ...
                For I = 1 To g_Channel.Users.Count
                    ' ...
                    If (StrComp(g_Channel.Users(I).DisplayName, GetCurrentUsername, vbBinaryCompare) <> 0) Then
                        ' ...
                        If (g_Channel.Users(I).IsOperator) Then
                            ' ...
                            opsCount = (opsCount + 1)
                        End If
                    End If
                Next I
            End If
        End If
        
        ' ...
        If (StrComp(Left$(g_Channel.Name, 3), "Op ", vbTextCompare) = 0) Then
            ' ...
            If (opsCount >= 2) Then
                ' ...
                cmdRet(0) = "Error: There is currently a channel moderator present that cannot be " & _
                    "removed from his or her position."
                
                ' ...
                Exit Function
            End If
        ElseIf (StrComp(Left$(g_Channel.Name, 5), "Clan ", vbTextCompare) = 0) Then
            ' ...
            If ((g_Clan.Self.Rank < 4) Or _
                    (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) <> 0)) Then
                    
                ' ...
                If (opsCount >= 2) Then
                    ' ...
                    cmdRet(0) = "Error: There is currently a channel moderator present that " & _
                            "cannot be removed from his or her position."
                    
                    ' ...
                    Exit Function
                End If
            End If
        End If
        
        ' ...
        'If (QueueLoad > 0) Then
        '    Call Pause(2, True, False)
        'End If
        
        ' designate user
        Call bnetSend("/designate " & reverseUsername(msgData))
        
        ' ...
        'frmChat.AddChat vbRed, "DEBUG: DESIGNATE " & msgData
        
        ' ...
        Call Pause(3, True, False)
        
        ' rejoin channel
        Call bnetSend("/resign")
        
        ' ...
        'frmChat.AddChat vbRed, "DEBUG: RESIGN"

        ' ...
        If (userCount > 0) Then
            ' promote shamans again
            For I = 0 To (userCount - 1)
                ' ...
                g_Clan.Members(g_Clan.GetUserIndexEx(arrUsers(I))).Promote
                
                ' ...
                'frmChat.AddChat vbRed, _
                '    "DEBUG: PROMOTE " & g_Clan.Members(g_Clan.GetUserIndexEx(arrUsers(I))).DisplayName
                
                ' ...
                'Call Pause(200, True, True)
            Next I
        End If

        ' ...
        ReDim arrUsers(0)
    Else
        ' ...
        cmdRet(0) = "Error: The specified user is not present within the channel."
    End If
End Function ' end function OnGiveUp

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

' handle chpw command
Private Function OnChPw(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String ' ...
    Dim tmpbuf     As String ' temporary output buffer
    
    ' ...
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        strArray = Split(msgData, " ", 2)
    Else
        ' ...
        ReDim Preserve strArray(0)
        
        ' ...
        strArray(0) = msgData
    End If
    
    Select Case (LCase$(strArray(0)))
        Case "on", "set"
            If ((UBound(strArray) >= 1) And Len(strArray(1))) Then
                If (BotVars.ChannelPasswordDelay <= 0) Then
                    BotVars.ChannelPasswordDelay = 30
                End If
                
                tmpbuf = "Channel password protection enabled, delay set to " & _
                    BotVars.ChannelPasswordDelay & "."
            Else
                tmpbuf = "Error: Invalid channel password."
            End If
            
        Case "off", "kill", "clear"
            BotVars.ChannelPassword = vbNullString
            
            BotVars.ChannelPasswordDelay = 0
            
            tmpbuf = "Channel password protection disabled."
            
        Case "time", "delay", "wait"
            If (StrictIsNumeric(strArray(1))) Then
                If ((Val(strArray(1)) <= 255) And _
                    (Val(strArray(1)) >= 1)) Then
                   
                    BotVars.ChannelPasswordDelay = CByte(Val(strArray(1)))
                    
                    tmpbuf = "Channel password delay set to " & strArray(1) & "."
                Else
                    tmpbuf = "Error: Invalid channel delay."
                End If
            Else
                tmpbuf = "Error: Invalid channel delay."
            End If
            
        Case "info", "status"
            If ((BotVars.ChannelPassword = vbNullString) Or _
                (BotVars.ChannelPasswordDelay = 0)) Then
                
                tmpbuf = "Channel password protection is disabled."
            Else
                tmpbuf = "Channel password protection is enabled. Password [" & _
                    BotVars.ChannelPassword & "], Delay [" & _
                        BotVars.ChannelPasswordDelay & "]."
            End If
            
        Case Else
            tmpbuf = "Error: Unknown channel password command."
    End Select
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnChPw

' handle sethome command
Private Function OnSetHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the home channel to the channel specified.
    ' The home channel is the channel that the bot joins immediately
    ' following a completion of the connection procedure.
    
    Dim tmpbuf As String ' temporary output buffer

    Call WriteINI("Main", "HomeChan", msgData)
    
    BotVars.HomeChannel = msgData
    
    tmpbuf = "Home channel set to """ & msgData & """."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetHome

' handle resign command
Private Function OnResign(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot resign from the role of operator
    ' by rejoining the channel through the use of Battle.net's /resign
    ' command.
    
    If ((g_Channel.Self.IsOperator) = False) Then
        Exit Function
    End If
    
    Call AddQ("/resign", PRIORITY.SPECIAL_MESSAGE, Username)
End Function ' end function OnResign

' handle clearbanlist
Private Function OnClearBanList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will clear the bot's internal list of banned users.
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    g_Channel.ClearBanlist
    
    tmpbuf = "Banned user list cleared."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnClearBanList

' handle kickonyell command
Private Function OnKickOnYell(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    Select Case (LCase$(msgData))
        Case "on"
            BotVars.KickOnYell = 1
            
            tmpbuf = "Kick-on-yell enabled."
            
            Call WriteINI("Other", "KickOnYell", "Y")
            
        Case "off"
            BotVars.KickOnYell = 0
            
            tmpbuf = "Kick-on-yell disabled."
            
            Call WriteINI("Other", "KickOnYell", "N")
            
        Case "status"
            tmpbuf = "Kick-on-yell is "
            tmpbuf = tmpbuf & IIf(BotVars.KickOnYell = 1, "enabled", "disabled") & "."
    End Select
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnKickOnYell

' handle plugban command
Private Function OnPlugBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable, or check the status of, UDP plug bans.
    ' UDP plugs were traditionally used, in place of lag bars, to signifiy
    ' that a user was incapable of hosting (or possibly even joining) a game.
    ' However, as bot development became more popular, the emulation of such
    ' a connectivity issue became fairly common, and the UDP plug began to
    ' represent that a user was using a bot.  This feature allows for the
    ' banning of both, potential bots, and users unlikely to be capable of
    ' creating and/or joining games based on the UDP protocl.
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    msgData = LCase$(msgData)

    Select Case (msgData)
        Case "on"
            Dim I As Integer
        
            If (BotVars.PlugBan) Then
                tmpbuf = "PlugBan is already activated."
            Else
                BotVars.PlugBan = True
                
                tmpbuf = "PlugBan activated."
                
                Call g_Channel.CheckUsers
                
                Call WriteINI("Other", "PlugBans", "Y")
            End If
            
        Case "off"
            If (BotVars.PlugBan) Then
                BotVars.PlugBan = False
                
                tmpbuf = "PlugBan deactivated."
                
                Call WriteINI("Other", "PlugBans", "N")
            Else
                tmpbuf = "PlugBan is already deactivated."
            End If
            
        Case "status"
            If (BotVars.PlugBan) Then
                tmpbuf = "PlugBan is activated."
            Else
                tmpbuf = "PlugBan is deactivated."
            End If
    End Select

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPlugBan

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
            Call OnAdd(Username, dbAccess, user & " +B --type GAME --banmsg " & bmsg, True, tmpbuf())
        End If
    Else
        ' ...
        Call OnAdd(Username, dbAccess, msgData & " +B --type GAME", True, tmpbuf())
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
        Call OnAdd(Username, dbAccess, msgData & " -B --type GAME", True, tmpbuf())
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
    
    ' ...
    If (g_Channel.Banlist.Count = 0) Then
        ' ...
        cmdRet(0) = "There are presently no users on the bot's internal banlist."
    
        ' ...
        Exit Function
    End If

    ' ...
    tmpbuf(tmpCount) = "User(s) banned: "
    
    ' ...
    For I = 1 To g_Channel.Banlist.Count
        ' ...
        If (g_Channel.Banlist(I).IsDuplicateBan = False) Then
            ' ...
            For j = 1 To g_Channel.Banlist.Count
                ' ...
                If (StrComp(g_Channel.Banlist(j).DisplayName, g_Channel.Banlist(I).DisplayName, _
                        vbTextCompare) = 0) Then
                
                    ' ...
                    userCount = (userCount + 1)
                End If
            Next j
            
            ' ...
            tmpbuf(tmpCount) = _
                    tmpbuf(tmpCount) & ", " & g_Channel.Banlist(I).DisplayName
                    
            ' ...
            If (userCount > 1) Then
                tmpbuf(tmpCount) = _
                        tmpbuf(tmpCount) & " (" & userCount & ") "
            End If
                    
            ' ...
            If ((Len(tmpbuf(tmpCount)) > 90) And (I <> g_Channel.Banlist.Count)) Then
                ' increase array size
                ReDim Preserve tmpbuf(tmpCount + 1)
            
                ' apply postfix to previous line
                tmpbuf(tmpCount) = Replace(tmpbuf(tmpCount), " , ", Space$(1)) & " [more]"
                
                ' apply prefix to new line
                tmpbuf(tmpCount + 1) = "User(s) banned: "
                
                ' incrememnt counter
                tmpCount = (tmpCount + 1)
            End If
    
            tmpbuf(tmpCount) = Replace(tmpbuf(tmpCount), " , ", Space$(1))
        End If
        
        ' ...
        userCount = 0
    Next I
    
    'For i = LBound(gBans) To UBound(gBans)
    '    If (gBans(i).userName <> vbNullString) Then
    '        tmpBuf(tmpCount) = tmpBuf(tmpCount) & ", " & gBans(i).userName
    '
    '        If ((Len(tmpBuf(tmpCount)) > 90) And (i <> UBound(gBans))) Then
    '            ' increase array size
    '            ReDim Preserve tmpBuf(tmpCount + 1)
    '
    '            ' apply postfix to previous line
    '            tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1)) & " [more]"
    '
    '            ' apply prefix to new line
    '            tmpBuf(tmpCount + 1) = "Banned users: "
    '
    '            ' incrememnt counter
    '            tmpCount = (tmpCount + 1)
    '        End If
    '
    '        ' incrememnt counter
    '        BanCount = (BanCount + 1)
    '    End If
    'Next i
    '
    ' has anyone been banned?
    'If (BanCount = 0) Then
    '    tmpBuf(tmpCount) = "No users have been banned."
    'Else
    '    tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1))
    'End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnBanned

' handle ipbans command
Private Function OnIPBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim I      As Integer
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)

    If (Left$(msgData, 2)) = "on" Then
        BotVars.IPBans = True
        
        Call WriteINI("Other", "IPBans", "Y")
        
        tmpbuf = "IPBanning activated."
        
        Call g_Channel.CheckUsers
        
        'If ((MyFlags = 2) Or (MyFlags = 18)) Then
        '    For i = 1 To colUsersInChannel.Count
        '        Select Case colUsersInChannel.Item(i).Flags
        '            Case 20, 30, 32, 48
        '                Call AddQ("/ban " & colUsersInChannel.Item(i).Name & _
        '                    " IPBanned.")
        '        End Select
        '    Next i
        'End If
    ElseIf (Left$(msgData, 3) = "off") Then
        BotVars.IPBans = False
        
        Call WriteINI("Other", "IPBans", "N")
        
        tmpbuf = "IPBanning deactivated."
        
    ElseIf (Left$(msgData, 6) = "status") Then
        If (BotVars.IPBans) Then
            tmpbuf = "IPBanning is currently active."
        Else
            tmpbuf = "IPBanning is currently disabled."
        End If
    Else
        tmpbuf = "Error: Unrecognized IPBan command. Use 'on', 'off' or 'status'."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIPBans

' handle ipban command
Private Function OnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim gAcc   As udtGetAccessResponse

    Dim tmpbuf As String ' temporary output buffer
    Dim tmpAcc As String ' ...
    
    Dim msgFirstPart As String ' ...

    ' Get the first part of the message. (the username given)
    msgFirstPart = Split(msgData, " ")(0)
    
    ' ...
    tmpAcc = StripInvalidNameChars(msgFirstPart)

    ' ...
    If (Len(tmpAcc) > 0) Then
        ' ...
        If (InStr(1, tmpAcc, "@") > 0) Then
            tmpAcc = StripRealm(tmpAcc)
        End If
        
        ' ...
        If (dbAccess.Rank <= 100) Then
            If ((GetSafelist(tmpAcc)) Or (GetSafelist(msgFirstPart))) Then
                ' return message
                cmdRet(0) = "Error: That user is safelisted."
                
                Exit Function
            End If
        End If
        
        ' ...
        gAcc = GetAccess(msgFirstPart)
        
        ' ...
        If ((gAcc.Rank >= dbAccess.Rank) Or _
            ((InStr(gAcc.Flags, "A") > 0) And (dbAccess.Rank <= 100))) Then

            tmpbuf = "Error: You do not have enough access to do that."
        Else
            Call AddQ("/ban " & msgData, , Username)
            Call AddQ("/squelch " & msgFirstPart, , Username)
        
            tmpbuf = "User " & Chr(34) & msgFirstPart & Chr(34) & " IPBanned."
        End If
    Else
        ' return message
        tmpbuf = "Error: You do not have enough access to do that."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIPBan

' handle unipban command
Private Function OnUnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (LenB(msgData) > 0) Then
        ' ...
        Call AddQ("/unsquelch " & msgData, , Username)
        
        ' ...
        Call AddQ("/unban " & msgData, , Username)
        
        ' ...
        tmpbuf = "User " & Chr$(34) & msgData & Chr$(34) & _
            " Un-IPBanned."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUnIPBan

' handle designate command
Private Function OnDesignate(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (LenB(msgData) > 0) Then
        ' ...
        If (g_Channel.Self.IsOperator) Then
            ' ...
            Call AddQ("/designate " & msgData, , Username)
            
            ' ...
            tmpbuf = "I have designated " & msgData & "."
        Else
            ' ...
            tmpbuf = "Error: The bot does not have ops."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDesignate

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

' handle whispercmds command
Private Function OnWhisperCmds(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (StrComp(msgData, "status", vbTextCompare) = 0) Then
        tmpbuf = "Command responses will be " & _
            IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly") & "."
    Else
        If (BotVars.WhisperCmds) Then
            BotVars.WhisperCmds = False
            
            Call WriteINI("Main", "WhisperBack", "N")
            
            tmpbuf = "Command responses will now be displayed publicly."
        Else
            BotVars.WhisperCmds = True

            Call WriteINI("Main", "WhisperBack", "Y")
            
            tmpbuf = "Command responses will now be whispered back."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWhisperCmds

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

' handle shitdel command
Private Function OnShitDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpbuf(0) = "Error: The specified username is invalid."
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, msgData & " -B --type USER", _
            True, tmpbuf())
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnShitDel

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
        Call OnAdd(Username, dbAccess, U & " -S --type USER", True, tmpbuf())
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
        Call OnAdd(Username, dbAccess, msgData & " -B --type USER", True, tmpbuf())
        Call OnAdd(Username, dbAccess, msgData & " -B --type CLAN", True, tmpbuf())
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, "*" & msgData & "*" & " -B --type USER", _
            True, tmpbuf())
        Call OnAdd(Username, dbAccess, "*" & msgData & "*" & " -B --type CLAN", _
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

' handle settrigger command
Private Function OnSetTrigger(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf     As String ' temporary output buffer
    Dim newTrigger As String ' ...
    
    ' ...
    newTrigger = msgData
    
    ' ...
    If (LenB(newTrigger) > 0) Then
        ' ...
        'If (Left$(newTrigger, 1) = "/") Then
        '    ' ...
        '    cmdRet(0) = "Error: Trigger may not begin with a " & _
        '        "forward slash."
        '
        '    ' ...
        '    Exit Function
        'End If
        
        ' ...
        'If ((Left$(newTrigger, 1) = Space$(1)) Or (Right$(newTrigger, 1) = Space$(1))) Then
        '
        '    ' ...
        '    cmdRet(0) = "Error: Trigger may not begin or end with a " & _
        '        "space."
        '
        '    ' ...
        '    Exit Function
        'If (Left$(newTrigger, 1) = "/") Then
        '    ' ...
        '    cmdRet(0) = "Error: Trigger may not begin with a " & _
        '        "forward slash."
        '
        '    ' ...
        '    Exit Function
        'End If
        
        ' set new trigger
        BotVars.Trigger = newTrigger
    
        ' write trigger to configuration
        Call WriteINI("Main", "Trigger", "{" & newTrigger & "}")
    
        ' ...
        tmpbuf = "The new trigger is " & Chr$(34) & newTrigger & _
            Chr$(34) & "." & ""
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetTrigger

' handle levelban command
Private Function OnLevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim I      As Integer ' ...
    Dim tmpbuf As String  ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (StrictIsNumeric(msgData)) Then
            I = Val(msgData)
            
            If (I >= 1) Then
                If (I <= 255) Then
                    tmpbuf = "Banning Warcraft III users under level " & I & "."
                    
                    BotVars.BanUnderLevel = CByte(I)
                Else
                    tmpbuf = "Error: Invalid level specified."
                End If
            Else
                tmpbuf = "Levelbans disabled."
                
                BotVars.BanUnderLevel = 0
            End If
        Else
            BotVars.BanUnderLevel = 0
            
            tmpbuf = "Levelbans disabled."
        End If
        
        Call WriteINI("Other", "BanUnderLevel", BotVars.BanUnderLevel)
    Else
        If (BotVars.BanUnderLevel = 0) Then
           tmpbuf = "Currently not banning Warcraft III users by level."
        Else
            tmpbuf = "Currently banning Warcraft III users under level " & _
                BotVars.BanUnderLevel & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnLevelBan

' handle d2levelban command
Private Function OnD2LevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim I      As Integer
    Dim tmpbuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (StrictIsNumeric(msgData)) Then
            I = Val(msgData)
                
            If (I >= 1) Then
                If (I <= 255) Then
                    BotVars.BanD2UnderLevel = CByte(I)
            
                    tmpbuf = "Banning Diablo II characters under level " & I & "."
                Else
                    tmpbuf = "Error: Invalid level specified."
                End If
            Else
                tmpbuf = "Diablo II Levelbans disabled."
                
                BotVars.BanD2UnderLevel = 0
            End If
        Else
            tmpbuf = "Diablo II Levelbans disabled."
            
            BotVars.BanD2UnderLevel = 0
        End If
        
        Call WriteINI("Other", "BanD2UnderLevel", BotVars.BanD2UnderLevel)
    Else
        If (BotVars.BanD2UnderLevel = 0) Then
           tmpbuf = "Currently not banning Diablo II users by level."
        Else
           tmpbuf = "Currently banning Diablo II users under level " & BotVars.BanD2UnderLevel & "."
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnD2LevelBans

' handle phrasebans command
Private Function OnPhraseBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    Dim tmpbuf As String ' temporary output buffer
  
    If (Len(msgData) > 0) Then
        ' ...
        msgData = LCase$(msgData)
    
        If (msgData = "on") Then
            Call WriteINI("Other", "Phrasebans", "Y")
            
            PhraseBans = True
            
            tmpbuf = "Phrasebans activated."
        Else
            Call WriteINI("Other", "Phrasebans", "N")
            
            PhraseBans = False
            
            tmpbuf = "Phrasebans deactivated."
        End If
    Else
        If (PhraseBans = True) Then
            tmpbuf = "Phrasebans are enabled."
        Else
            tmpbuf = "Phrasebans are disabled."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPhraseBans

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
        Call OnAdd(Username, dbAccess, user & tag_msg & " --type CLAN", True, tmpbuf())
    Else
        ' ...
        Call OnAdd(Username, dbAccess, user & tag_msg & " --type USER", True, tmpbuf())
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
    
        Call OnAdd(Username, dbAccess, msgData & safe_msg, True, tmpbuf())
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

' handle exile command
Private Function OnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf     As String ' temporary output buffer
    Dim saCmdRet() As String ' ...
    Dim ibCmdRet() As String ' ...
    Dim U          As String ' ...
    Dim Y          As String ' ...
    
    ' ...
    ReDim Preserve saCmdRet(0)
    ReDim Preserve ibCmdRet(0)

    ' ...
    U = msgData
    
    ' ...
    Call OnShitAdd(Username, dbAccess, U, InBot, saCmdRet())
    
    ' ...
    Call OnIPBan(Username, dbAccess, U, InBot, ibCmdRet())
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnExile

' handle unexile command
Private Function OnUnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf     As String ' temporary output buffer
    Dim U          As String ' ...
    Dim sdCmdRet() As String ' ...
    Dim uiCmdRet() As String ' ...
    
    ' declare index zero of array
    ReDim Preserve sdCmdRet(0)
    ReDim Preserve uiCmdRet(0)

    ' ...
    U = msgData
    
    ' ...
    Call OnShitDel(Username, dbAccess, U, InBot, sdCmdRet())
    
    ' ...
    Call OnUnIPBan(Username, dbAccess, U, InBot, uiCmdRet())
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUnExile

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

' handle shitadd command
Private Function OnShitAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String  ' ...
    Dim Index    As Integer ' ...
    Dim shit_msg As String  ' ...
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' ...
    If (BotVars.DefaultShitlistGroup <> vbNullString) Then
        Dim default_group_access As udtGetAccessResponse
        
        ' ...
        default_group_access = _
                GetAccess(BotVars.DefaultShitlistGroup, "GROUP")
        
        ' ...
        If (default_group_access.Username <> vbNullString) Then
            shit_msg = " --group " & BotVars.DefaultShitlistGroup
        End If
    End If
    
    ' ...
    If (shit_msg = vbNullString) Then
        shit_msg = " +B"
    End If
    
    ' ...
    Index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (Index) Then
        Dim user As String ' ...
        
        ' ...
        user = Mid$(msgData, 1, Index - 1)
        
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            tmpbuf(0) = "Error: The specified username is invalid."
        Else
            Dim Msg As String ' ...
            
            ' ...
            Msg = Mid$(msgData, Index + 1)
        
            ' ...
            shit_msg = user & shit_msg & " --type USER --banmsg " & Msg
        End If
    Else
        ' ...
        shit_msg = msgData & shit_msg & " --type USER"
    End If
    
    ' ...
    Call OnAdd(Username, dbAccess, shit_msg, True, tmpbuf())
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnShitAdd

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
' handle readfile command
Private Function OnReadFile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    
    ' redefine array size
    ReDim Preserve tmpbuf(tmpCount)
    
    U = msgData
    
    If (Len(U)) Then
        If (InStr(1, U, "..", vbBinaryCompare) <> 0) Then
            tmpbuf(tmpCount) = "Error: You may only specify a file within the program " & _
                "directory or subdirectories."
        ElseIf (InStr(1, U, ".ini", vbTextCompare) <> 0) Then
            tmpbuf(tmpCount) = "Error: You may not read configuration files."
        Else
            Dim Y As String  ' ...
            Dim f As Integer ' ...
        
            ' grab a file number
            f = FreeFile
        
            If (InStr(1, U, ".", vbBinaryCompare) > 0) Then
                Y = Left$(U, InStr(1, U, ".", vbBinaryCompare) - 1)
            Else
                Y = U
            End If
            
            Select Case (UCase$(Y))
                Case "CON", "PRN", "AUX", "CLOCK$", "NUL", "COM1", _
                     "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", _
                     "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", _
                     "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            
                    cmdRet(0) = "Error: You cannot read the specified file."
            
                    Exit Function
            End Select
            
            ' get absolute file path
            U = Dir$(App.Path & "\" & U)
            
            If (U = vbNullString) Then
                tmpbuf(tmpCount) = "Error: The specified file could not " & _
                    "be found."
            Else
                ' store line in buffer
                tmpbuf(tmpCount) = "Contents of file " & _
                    msgData & ":"
                
                ' increment counter
                tmpCount = (tmpCount + 1)
            
                ' open file
                Open U For Input As #f
                    ' read until end-of-line
                    Do While (Not (EOF(f)))
                        Dim tmp As String ' ...
                        
                        ' read line into tmp
                        Line Input #f, tmp
                        
                        If (tmp <> vbNullString) Then
                            ' redefine array size
                            ReDim Preserve tmpbuf(tmpCount)
                        
                            ' store line in buffer
                            tmpbuf(tmpCount) = "Line " & tmpCount & ": " & _
                                tmp
                            
                            ' increment counter
                            tmpCount = (tmpCount + 1)
                        End If
                    Loop
                Close #f
                
                ' redefine array size
                ReDim Preserve tmpbuf(tmpCount)
                
                ' store line in buffer
                tmpbuf(tmpCount) = "End of File."
            End If
        End If
    End If
    
    ' return message
    cmdRet() = tmpbuf()
    
    Exit Function
    
ERROR_HANDLER:
    cmdRet(0) = "There was an error reading the specified file."

    Exit Function
End Function ' end function OnReadFile

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

' handle tally command
Private Function OnTally(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
     
    Dim tmpbuf As String ' temporary output buffer
     
    If (VoteDuration > 0) Then
        tmpbuf = Voting(BVT_VOTE_TALLY)
    Else
        tmpbuf = "No vote is currently in progress."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnTally

' handle cancel command
Private Function OnCancel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (VoteDuration > 0) Then
        tmpbuf = Voting(BVT_VOTE_END, BVT_VOTE_CANCEL)
    Else
        tmpbuf = "No vote in progress."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnCancel

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

' handle checkmail command
Private Function OnCheckMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Track  As Long
    Dim tmpbuf As String ' temporary output buffer
    
    If (InBot) Then
        Track = GetMailCount(GetCurrentUsername)
    Else
        Track = GetMailCount(Username)
    End If
    
    If (Track > 0) Then
        tmpbuf = "You have " & Track & " new messages."
        
        If (InBot) Then
            tmpbuf = tmpbuf & " Type /getmail to retrieve them."
        Else
            tmpbuf = tmpbuf & " Type !inbox to retrieve them."
        End If
    Else
        tmpbuf = "You have no mail."
    End If
     
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnCheckMail

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

' TO DO:
' handle add command
Public Function OnAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
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

' handle mmail command
Private Function OnMMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim temp       As udtMail
    
    Dim strArray() As String
    Dim tmpbuf     As String ' temporary output buffer
    Dim c          As Integer
    Dim f          As Integer
    Dim Track      As Long
    
    If (InBot) Then
        If (g_Online) Then
            Username = GetCurrentUsername
        Else
            Username = BotVars.Username
        End If
    End If
    
    strArray = Split(msgData, " ", 2)
            
    If (UBound(strArray) > 0) Then
        Dim gAcc As udtGetAccessResponse ' ...
    
        tmpbuf = "Mass mailing "

        With temp
            .From = Username
            .Message = strArray(1)
            
            If (StrictIsNumeric(strArray(0))) Then
                'number games
                Track = Val(strArray(0))
                
                For c = 0 To UBound(DB)
                    If (StrComp(DB(c).Type, "USER", vbTextCompare) = 0) Then
                        gAcc = GetCumulativeAccess(DB(c).Username)
                        
                        If (gAcc.Rank = Track) Then
                            .To = DB(c).Username
                            
                            Call AddMail(temp)
                        End If
                    End If
                Next c
                
                tmpbuf = tmpbuf & "to users with access " & Track
            Else
                For c = 0 To UBound(DB)
                    If (StrComp(DB(c).Type, "USER", vbTextCompare) = 0) Then
                        gAcc = GetCumulativeAccess(DB(c).Username)
                    
                        For f = 1 To Len(gAcc.Flags)
                            If (InStr(1, gAcc.Flags, Mid$(strArray(0), f, 1), _
                                vbBinaryCompare) > 0) Then
                                
                                .To = DB(c).Username
                                
                                Call AddMail(temp)
                                
                                Exit For
                            End If
                        Next f
                    End If
                Next c
                
                tmpbuf = tmpbuf & "to users with any of the flags " & strArray(0)
            End If
        End With
        
        tmpbuf = tmpbuf & " complete."
    Else
        tmpbuf = "Format: .mmail <flag(s)> <message> OR .mmail <access> <message>"
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnMMail

' handle bmail command
Private Function OnBMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim temp       As udtMail ' ...

    Dim strArray() As String ' ...
    Dim tmpbuf     As String ' temporary output buffer
    
    ' ...
    If (InBot) Then
        If (g_Online) Then
            Username = GetCurrentUsername
        Else
            Username = BotVars.Username
        End If
    End If
    
    ' ...
    strArray = Split(msgData, " ", 2)
    
    If (UBound(strArray) > 0) Then
        ' ...
        With temp
            .To = strArray(0)
            .From = Username
            .Message = strArray(1)
        End With
        
        If (Len(temp.To) = 0) Then
            tmpbuf = "Error: Invalid user."
        Else
            ' ...
            Call AddMail(temp)
            
            tmpbuf = "Added mail for " & strArray(0) & "."
        End If
    Else
        tmpbuf = "Error: Too few arguments."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnBMail

' handle designated command
Private Function OnDesignated(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (g_Channel.Self.IsOperator = False) Then
        tmpbuf = "The bot does not currently have ops."
    ElseIf (g_Channel.OperatorHeir = vbNullString) Then
        tmpbuf = "No users have been designated."
    Else
        tmpbuf = "I have designated """ & g_Channel.OperatorHeir & """."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDesignated

' handle flip command
Private Function OnFlip(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim I      As Integer
    Dim tmpbuf As String ' temporary output buffer

    Randomize
    
    I = (Rnd * 2)
    
    If (I = 0) Then
        tmpbuf = "Tails."
    Else
        tmpbuf = "Heads."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnFlip

' handle clear command
Private Function OnClear(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    frmChat.mnuClear_Click
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnClear

' handle monitor command
'Private Function OnMonitor(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'
'    Dim tmpBuf As String ' temporary output buffer
'
'    If (Len(msgData) > 0) Then
'        If (LCase$(msgData) = "on") Then
'            If (Not (MonitorExists)) Then
'                InitMonitor
'                If (MonitorForm.Connect(False)) Then
'                    tmpBuf = "User monitor connecting."
'                Else
'                    tmpBuf = "User monitor login information not filled in."
'                End If
'            Else
'                tmpBuf = "User montor already enabled."
'            End If
'        ElseIf (LCase$(msgData) = "off") Then
'            If (Not MonitorExists) Then
'                tmpBuf = "User monitor is not running."
'            Else
'                MonitorForm.ShutdownMonitor
'                tmpBuf = "User monitor disabled."
'            End If
'        Else
'            If (Not (MonitorExists())) Then
'                Call InitMonitor
'            End If
'
'            If (MonitorForm.AddUser(msgData)) Then
'                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & _
'                    " added to the monitor list."
'            Else
'                tmpBuf = "Failed to add user " & Chr$(&H22) & msgData & Chr$(&H22) & _
'                    " to the monitor list. (Contains Spaces, or already in the list)"
'            End If
'        End If
'    End If
'
'    ' return message
'    cmdRet(0) = tmpBuf
'End Function ' end function OnMonitor

' handle unmonitor command
'Private Function OnUnMonitor(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'
'    Dim tmpBuf As String ' temporary output buffer
'
'    If (Len(msgData) > 0) Then
'        If (MonitorExists) Then
'            If (MonitorForm.RemoveUser(msgData)) Then
'                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & " was removed from the monitor list."
'            Else
'                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & " was not found in the monitor list."
'            End If
'        Else
'            tmpBuf = "User monitor is not enabled."
'        End If
'    End If
'
'    ' return message
'    cmdRet(0) = tmpBuf
'End Function ' end function OnUnMonitor

' handle online command
'Private Function OnOnline(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'
'    Dim tmpBuf As String ' temporary output buffer
'
'    If (MonitorExists) Then
'        tmpBuf = MonitorForm.OnlineUsers
'    Else
'        tmpBuf = "User monitor is not enabled."
'    End If
'
'    ' return message
'    cmdRet = Split(tmpBuf, vbNewLine)
'End Function ' end function OnOnline

' handle enable command
Private Function OnEnable(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error Resume Next
    
    Dim Module As Module
    Dim Name   As String  ' ...
    Dim I      As Integer ' ...
    Dim str    As String  ' ...

    ' ...
    If (frmChat.SControl.Modules.Count > 1) Then
        For I = 2 To frmChat.SControl.Modules.Count
            Set Module = frmChat.SControl.Modules(I)
            Name = _
                modScripting.GetScriptName(CStr(I))
                
            If (StrComp(Name, msgData, vbTextCompare) = 0) Then
                str = _
                    Module.CodeObject.GetSettingsEntry("Enabled")
            
                If (StrComp(str, "True", vbTextCompare) = 0) Then
                    cmdRet(0) = Name & " is already enabled."
                Else
                    Module.CodeObject.WriteSettingsEntry "Enabled", "True"
                    
                    InitScript Module
                        
                    cmdRet(0) = Name & " has been enabled."
                End If
            
                Exit Function
            End If
        Next I
    End If
    
    cmdRet(0) = "Error: Could not find specified script."
    
End Function

' handle disable command
Private Function OnDisable(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error Resume Next
    
    Dim Module As Module
    Dim Name   As String  ' ...
    Dim I      As Integer ' ...
    
    ' ...
    If (frmChat.SControl.Modules.Count > 1) Then
        For I = 2 To frmChat.SControl.Modules.Count
            Set Module = frmChat.SControl.Modules(I)
            
            Name = _
                modScripting.GetScriptName(CStr(I))
                
            If (StrComp(Name, msgData, vbTextCompare) = 0) Then
                RunInSingle Module, "Event_Close"
                
                Module.CodeObject.WriteSettingsEntry "Enabled", "False"
                    
                DestroyObjs Module

                cmdRet(0) = Name & " has been disabled."
                    
                Exit Function
            End If
        Next I
    End If
    
    cmdRet(0) = "Error: Could not find specified script."
    
End Function

' handle exec command
Private Function OnExec(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim ErrType As String

    frmChat.SControl.ExecuteStatement msgData
    
    Exit Function
    
ERROR_HANDLER:
    
    With frmChat.SControl
        ErrType = "runtime"
        
        If InStr(1, .Error.source, "compilation", vbBinaryCompare) > 0 Then ErrType = "parsing"
        
        If InBot Then
            frmChat.AddChat RTBColors.ErrorMessageText, _
                "Execution " & ErrType & " error " & Chr(39) & .Error.Number & Chr(39) & _
                ": (column " & .Error.Column & ")"
            frmChat.AddChat RTBColors.ErrorMessageText, .Error.description
        Else
            cmdRet(0) = "Execution " & ErrType & " error " & Chr(39) & .Error.Number & Chr(39) & _
                ": " & .Error.description
        End If
        .Error.Clear
    End With
    
    Resume Next
    
End Function

' requires public
Public Function cache(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
    Static s()  As String
    Static sTyp As String
    Static bChannelListFollows  As Boolean 'renamed this variable for clarity
    
    Dim I       As Integer
    
    ' Mode=255 means we're resetting to get ready for a sweepban. [ugly]-andy
    If (Mode = 255) Then
        ' ...
        ReDim s(0)
        
        ' ...
        sTyp = Typ
        
        ' ...
        bChannelListFollows = False
    End If
    
    If (InStr(1, LCase$(Inpt), "in channel ", vbTextCompare) <> 0) Then
        ' we weren't expecting a channel list, but we are now
        bChannelListFollows = True
    Else
        If (bChannelListFollows = True) Then
            ' if we're expecting a channel list, process it
            Select Case (Mode)
                Case 0 ' RETRIEVE
                    ' Merge all the cache array items into one space-delimited string
                    For I = 0 To UBound(s)
                        cache = cache & s(I) & ", "
                    Next I
        
                    ' Clear the cache array
                    ReDim s(0)
                    
                    Typ = sTyp
                    
                Case 1 ' ADD
                    ' Expand the cache array out one value
                    ReDim Preserve s(UBound(s) + 1)
                    
                    ' Add this item to the cache array
                    ' With some extra processing for D2 realm characters...
                    
                    s(UBound(s)) = Inpt
            End Select
        End If
    End If
End Function

Private Function AddQ(ByVal s As String, Optional msg_priority As Integer = -1, Optional ByVal user As String = _
    vbNullString, Optional ByVal Tag As String = vbNullString) As Integer
    
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
        
        'frmchat.addchat rtbcolors.ConsoleText, "Fired."
        'frmchat.addchat rtbcolors.ConsoleText, "Initial smsgData: " & smsgData
        'frmchat.addchat rtbcolors.ConsoleText, "Initial sMatch: " & sMatch
        
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
                        ' unipban user
                        'If (BotVars.IPBans = True) Then
                        '    Call AddQ("/unsquelch " & gBans(i).userNameActual, 1)
                        'End If
                    
                        Call AddQ("/" & Typ & g_Channel.Banlist(I).DisplayName)
                    Else
                        ' unipban user
                        'If (BotVars.IPBans = True) Then
                        '    Call AddQ("/unsquelch " & gBans(i).userNameActual, 1)
                        'End If
                    
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
        
        ' ...
        If ((gAcc.Type <> "%") And _
            (StrComp(gAcc.Type, "USER", vbTextCompare) <> 0)) Then
            
            ' ...
            gAcc.Username = gAcc.Username & _
                " (" & LCase$(gAcc.Type) & ")"
        End If
        
        If (gAcc.Rank > 0) Then
            If (gAcc.Flags <> vbNullString) Then
                tmpbuf = "Found user " & gAcc.Username & ", with access " & gAcc.Rank & _
                    " and flags " & gAcc.Flags & "."
            Else
                tmpbuf = "Found user " & gAcc.Username & ", with access " & gAcc.Rank & "."
            End If
        ElseIf (gAcc.Flags <> vbNullString) Then
            tmpbuf = "Found user " & gAcc.Username & ", with flags " & gAcc.Flags & "."
        Else
            tmpbuf = "No such user(s) found."
        End If
    Else
        For I = LBound(DB) To UBound(DB)
            Dim res        As Boolean ' store result of access check
            Dim blnChecked As Boolean ' ...
        
            If (DB(I).Username <> vbNullString) Then
                ' ...
                If (match <> vbNullString) Then
                    If (Left$(match, 1) = "!") Then
                        If (Not (LCase$(PrepareCheck(DB(I).Username)) Like _
                                (LCase$(Mid$(match, 2))))) Then

                            res = True
                        Else
                            res = False
                        End If
                    Else
                        If (LCase$(PrepareCheck(DB(I).Username)) Like _
                           (LCase$(match))) Then
                           
                            res = True
                        Else
                            res = False
                        End If
                    End If
                    
                    blnChecked = True
                End If
                
                ' ...
                If (Group <> vbNullString) Then
                    ' ...
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

Public Function RemoveItem(ByVal rItem As String, File As String, Optional ByVal dbType As String = _
    vbNullString) As String
    
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

Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = _
    vbNullString) As Boolean
    
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
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.description & " in DB_remove().")
        
    DB_remove = False
    
    Exit Function
End Function

' requires public
Public Function GetSafelist(ByVal Username As String) As Boolean

    Dim I As Long ' ...
    
    ' ...
    If (bFlood = False) Then
        ' ...
        Dim gAcc As udtGetAccessResponse
        
        ' ...
        gAcc = GetCumulativeAccess(Username, "USER")
        
        ' ...
        If (InStr(1, gAcc.Flags, "S", vbBinaryCompare) <> 0) Then
            GetSafelist = True
        ElseIf (gAcc.Rank >= 20) Then
            GetSafelist = True
        End If
    Else
        ' ...
        For I = 0 To (UBound(gFloodSafelist) - 1)
            If PrepareCheck(Username) Like gFloodSafelist(I) Then
                ' ...
                GetSafelist = True
                
                ' ...
                Exit For
            End If
        Next I
    End If
    
End Function

' requires public
Public Function GetShitlist(ByVal Username As String) As String

    Dim gAcc As udtGetAccessResponse
    Dim Ban  As Boolean
    
    ' ...
    gAcc = GetCumulativeAccess(Username, "USER")
    
    ' ...
    If (InStr(1, gAcc.Flags, "Z", vbBinaryCompare) <> 0) Then
        ' ...
        Ban = True
    ElseIf (InStr(1, gAcc.Flags, "B", vbBinaryCompare) <> 0) Then
        ' ...
        If (GetSafelist(Username) = False) Then
            ' ...
            Ban = True
        End If
    End If
    
    ' ...
    If (Ban) Then
        ' ...
        If ((Len(gAcc.BanMessage) > 0) And (gAcc.BanMessage <> "%")) Then
            GetShitlist = Username & Space(1) & gAcc.BanMessage
        Else
            GetShitlist = Username & Space(1) & "Shitlisted"
        End If
    End If
    
End Function

' requires public
Public Function GetPing(ByVal Username As String) As Long
    Dim I As Integer
    
    I = g_Channel.GetUserIndex(Username)
    
    If I > 0 Then
        GetPing = g_Channel.Users(I).Ping
    Else
        GetPing = -3
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

'Private Sub DBRemove(ByVal s As String)
'    Dim T()  As udtDatabase
'
'    Dim I    As Integer
'    Dim C    As Integer
'    Dim n    As Integer
'    Dim temp As String
'
'    s = LCase$(s)
'
'    For I = LBound(DB) To UBound(DB)
'        If StrComp(DB(I).Username, s, vbTextCompare) = 0 Then
'            ReDim T(0 To UBound(DB) - 1)
'            For C = LBound(DB) To UBound(DB)
'                If C <> I Then
'                    T(n) = DB(C)
'                    n = n + 1
'                End If
'            Next C
'
'            ReDim DB(UBound(T))
'            For C = LBound(T) To UBound(T)
'                DB(C) = T(C)
'            Next C
'            Exit Sub
'        End If
'    Next I
'
'    n = FreeFile
'
'    temp = GetFilePath("users.txt")
'
'    Open temp For Output As #n
'        For I = LBound(DB) To UBound(DB)
'            Print #n, DB(I).Username & Space(1) & DB(I).Rank & Space(1) & DB(I).Flags
'        Next I
'    Close #n
'End Sub

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

Private Function ValidateAccess(ByRef gAcc As udtGetAccessResponse, ByVal CWord As String, _
    Optional ByVal Argument As String = vbNullString, Optional ByVal restrictionName As String = _
        vbNullString) As Boolean
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    ' ...
    If (Len(CWord) > 0) Then
        Dim commands As DOMDocument60
        Dim Command  As IXMLDOMNode
        
        ' ...
        Set commands = New DOMDocument60
        
        ' ...
        If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
            Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
                "working directory.")
                
            Exit Function
        End If

        ' ...
        Call commands.Load(App.Path & "\commands.xml")
        
        ' ...
        For Each Command In commands.documentElement.childNodes
            Dim accessGroup As IXMLDOMNode
            Dim Access      As IXMLDOMNode
            Dim flag        As IXMLDOMNode
        
            ' ...
            If (StrComp(Command.Attributes.getNamedItem("name").Text, _
                CWord, vbTextCompare) = 0) Then
                
                ' ...
                Set accessGroup = Command.selectSingleNode("access")
                
                ' ...
                For Each Access In accessGroup.childNodes
                    If (LCase$(Access.nodeName) = "rank") Then
                        If ((gAcc.Rank) >= (Val(Access.Text))) Then
                            ValidateAccess = True
                        
                            Exit For
                        End If
                    '// 09/03/2008 JSM - Modified code to use the <flags> element
                    ElseIf (LCase$(Access.nodeName = "flags")) Then
                        For Each flag In Access.childNodes
                            If (InStr(1, gAcc.Flags, flag.Text, vbBinaryCompare) <> 0) Then
                                ValidateAccess = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                
                ' ...
                If (Argument <> vbNullString) Then
                    ' ...
                End If
                
                ' ...
                If (restrictionName <> vbNullString) Then
                    Dim Restrictions As IXMLDOMNodeList
                    Dim Restriction  As IXMLDOMNode
                    
                    ' ...
                    Set Restrictions = Command.selectNodes("restrictions/restriction")
                    
                    ' ...
                    For Each Restriction In Restrictions
                        If (StrComp(Restriction.Attributes.getNamedItem("name").Text, _
                            restrictionName, vbTextCompare) = 0) Then
                            
                            ' ...
                            Set accessGroup = Restriction.selectSingleNode("access")
                            
                            ' ...
                            For Each Access In accessGroup.childNodes
                                If (LCase$(Access.nodeName) = "rank") Then
                                    If ((gAcc.Rank) >= (Val(Access.Text))) Then
                                        ValidateAccess = True
                                    
                                        Exit For
                                    End If
                                '// 09/03/2008 JSM - Modified code to use the <flags> element
                                ElseIf (LCase$(Access.nodeName = "flags")) Then
                                    For Each flag In Access.childNodes
                                        If (InStr(1, gAcc.Flags, flag.Text, vbBinaryCompare) <> 0) Then
                                            ValidateAccess = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
                           
                Exit For
            End If
        Next
    End If
    
    Exit Function

' ...
ERROR_HANDLER::
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: XML Database Processor has encountered an error " & _
        "during access validation.")
    
    ' ...
    ValidateAccess = False

    Exit Function
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

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure " & _
        "WriteDatabase of Module modCommandCode"
    
    Resume WriteDatabase_Exit
End Sub

' requires public
Public Function DateCleanup(ByVal TDate As Date) As String
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
    
        ' ...
        For I = 1 To Len(user)
            ' ...
            Dim currentCharacter As String
            
            ' ...
            currentCharacter = Mid$(user, I, 1)

            ' is the character between A-Z or a-z?
            If (Asc(currentCharacter) < Asc("A")) Or (Asc(currentCharacter) > Asc("z")) Then
                MsgBox currentCharacter
            
                ' is the character between 0 - 9?
                If ((Asc(currentCharacter) < Asc("0")) Or (Asc(currentCharacter) > Asc("9"))) Then
                    'MsgBox Asc(currentCharacter)
                
                    ' !@$(){}[]=+`~^-’.:;_|
                
                    ' is the character a valid special
                    ' character?
                    
                    If ((Asc(currentCharacter) = Asc("[")) Or _
                        (Asc(currentCharacter) = Asc("]")) Or _
                        (Asc(currentCharacter) = Asc("(")) Or _
                        (Asc(currentCharacter) = Asc(")")) Or _
                        (Asc(currentCharacter) = Asc(".")) Or _
                        (Asc(currentCharacter) = Asc("-")) Or _
                        (Asc(currentCharacter) = Asc("_"))) Then
                    
                        ' ...
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
        ' does our user contain illegal
        ' characters?
        If (illegal) Then
            ' do we allow illegal
            ' characters?
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
            If ((BotVars.UseGameConventions = True) And _
                    ((BotVars.UseD2GameConventions = True))) Then
                    
                usingGameConventions = True
            End If
            
        Case "WAR3", "W3XP"
            If ((BotVars.UseGameConventions = True)) And _
                    ((BotVars.UseW3GameConventions = True)) Then
                    
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

Public Function SearchDatabase(ByRef arrReturn() As String, Optional Username As String = vbNullString, _
    Optional ByVal match As String = vbNullString, Optional Group As String = vbNullString, _
        Optional dbType As String = vbNullString, Optional lowerBound As Integer = -1, _
            Optional upperBound As Integer = -1, Optional Flags As String = vbNullString) As Integer
    
    On Error GoTo ERROR_HANDLER
    
    Dim I        As Integer
    Dim found    As Integer
    Dim tmpbuf   As String
    
    If (LenB(Username) > 0) Then
        Dim dbAccess As udtGetAccessResponse
        dbAccess = GetAccess(Username, dbType)
        
        If (Not (dbAccess.Type = "%") And (Not StrComp(dbAccess.Type, "USER", vbTextCompare) = 0)) Then
            dbAccess.Username = dbAccess.Username & " (" & LCase$(dbAccess.Type) & ")"
        End If
        
        If (dbAccess.Rank > 0) Then
            tmpbuf = "Found user " & dbAccess.Username & ", who holds rank " & dbAccess.Rank & _
                IIf(Len(dbAccess.Flags) > 0, " and flags " & dbAccess.Flags, vbNullString) & "."
        ElseIf (LenB(dbAccess.Flags) > 0) Then
            tmpbuf = "Found user " & dbAccess.Username & ", with flags " & dbAccess.Flags & "."
        Else
            tmpbuf = "No such user(s) found."
        End If
    Else
        For I = LBound(DB) To UBound(DB)
            Dim res        As Boolean
            Dim blnChecked As Boolean
        
            If (LenB(DB(I).Username) > 0) Then
                If (LenB(match) > 0) Then
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

                If (LenB(dbType) > 0) Then
                    If (StrComp(DB(I).Type, dbType, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If ((lowerBound >= 0) And (upperBound >= 0)) Then
                    If ((DB(I).Rank >= lowerBound) And (DB(I).Rank <= upperBound)) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                ElseIf (lowerBound >= 0) Then
                    If (DB(I).Rank = lowerBound) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If (LenB(Flags) > 0) Then
                    Dim j As Integer
                
                    For j = 1 To Len(Flags)
                        If (InStr(1, DB(I).Flags, Mid$(Flags, j, 1), vbBinaryCompare) = 0) Then
                            Exit For
                        End If
                    Next j
                    
                    If (j = (Len(Flags) + 1)) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If (res = True) Then
                    tmpbuf = tmpbuf & DB(I).Username
                    If (Not (DB(I).Type = "%") And (Not StrComp(DB(I).Type, "USER", vbTextCompare) = 0)) Then
                        tmpbuf = tmpbuf & " (" & LCase$(DB(I).Type) & ")"
                    End If
                    tmpbuf = tmpbuf & _
                        IIf(DB(I).Rank > 0, "\" & DB(I).Rank, vbNullString) & _
                        IIf(LenB(DB(I).Flags) > 0, "\" & DB(I).Flags, vbNullString) & ", "
                    
                    ' increment found counter
                    found = (found + 1)
                End If
            End If
            
            ' reset booleans
            res = False
            blnChecked = False
        Next I

        If (found = 0) Then
            arrReturn(0) = "No such user(s) found."
        Else
            Call SplitByLen(Mid$(tmpbuf, 1, Len(tmpbuf) - Len(", ")), 180, arrReturn(), _
                "User(s) found: ", " [more]", ", ")
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandCode.SearchDatabase()."
End Function

