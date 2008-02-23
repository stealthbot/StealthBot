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
Private Const WA_PREVTRACK   As Long = 40044 ' ...
Private Const WA_NEXTTRACK   As Long = 40048 ' ...
Private Const WA_PLAY        As Long = 40045 ' ...
Private Const WA_PAUSE       As Long = 40046 ' ...
Private Const WA_STOP        As Long = 40047 ' ...
Private Const WA_FADEOUTSTOP As Long = 40147 ' ...

Private m_dbAccess As udtGetAccessResponse
Private m_Username As String  ' ...
Private m_Console  As Boolean ' ...
Private m_Whisper  As Boolean ' ...

Public flood    As String ' ...?
Public floodCap As Byte   ' ...?

' prepares commands for processing, and calls helper functions associated with
' processing
Public Function ProcessCommand(ByVal Username As String, ByVal Message As String, _
    Optional ByVal InBot As Boolean = False, Optional ByVal WhisperedIn As Boolean = False) As Boolean
    
    ' default error response for commands
    On Error GoTo ERROR_HANDLER
    
    ' stores the access response for use when commands are
    ' issued via console
    Dim ConsoleAccessResponse As udtGetAccessResponse

    Dim i            As Integer ' loop counter
    Dim tmpmsg       As String  ' stores local copy of message
    Dim cmdRet()     As String  ' stores output of commands
    Dim PublicOutput As Boolean ' stores result of public command
                                ' output check (used for displaying command
                                ' output when issuing via console)
    
    ' create single command data array element for safe bounds checking
    ReDim Preserve cmdRet(0)
    
    ' create console access response structure
    With ConsoleAccessResponse
        .Access = 201
        .Flags = "A"
    End With
    
    m_Username = Username
    m_Console = InBot
    m_Whisper = WhisperedIn

    If (m_Console) Then
        m_dbAccess = ConsoleAccessResponse
    Else
        m_dbAccess = GetCumulativeAccess(Username, "USER")
    End If

    ' store local copy of message
    tmpmsg = Message
    
    ' replace message variables
    tmpmsg = Replace(tmpmsg, "%me", IIf((InBot), CurrentUsername, Username), 1)
    
    ' check for command identifier when command
    ' is not issued from within console
    If (Not (InBot)) Then
        ' we're going to mute our own access for now to prevent users from
        ' using the "say" command to gain full control over the bot.
        If (StrComp(Username, CurrentUsername, vbBinaryCompare) = 0) Then
            Exit Function
        End If
    
        ' check for commands using universal command identifier (?)
        If (StrComp(Left$(tmpmsg, Len("?trigger")), "?trigger", vbTextCompare) = 0) Then
            ' remove universal command identifier from message
            tmpmsg = Mid$(tmpmsg, 2)
        
        ' check for commands using command identifier
        ElseIf ((Len(tmpmsg) >= Len(BotVars.TriggerLong)) And _
                (Left$(tmpmsg, Len(BotVars.TriggerLong)) = BotVars.TriggerLong)) Then
            
            ' remove command identifier from message
            tmpmsg = Mid$(tmpmsg, Len(BotVars.TriggerLong) + 1)
        
            ' check for command identifier and name combination
            ' (e.g., .Eric[nK] say hello)
            If (Len(tmpmsg) >= (Len(CurrentUsername) + 1)) Then
                If (StrComp(Left$(tmpmsg, Len(CurrentUsername) + 1), _
                    CurrentUsername & Space(1), vbTextCompare) = 0) Then
                        
                    ' remove username (and space) from message
                    tmpmsg = Mid$(tmpmsg, Len(CurrentUsername) + 2)
                End If
            End If
        
        ' check for commands using either name and colon (and space),
        ' or name and comma (and space)
        ' (e.g., Eric[nK]: say hello; and, Eric[nK], say hello)
        ElseIf ((Len(tmpmsg) >= (Len(CurrentUsername) + 2)) And _
                ((StrComp(Left$(tmpmsg, Len(CurrentUsername) + 2), CurrentUsername & ": ", _
                  vbTextCompare) = 0) Or _
                 (StrComp(Left$(tmpmsg, Len(CurrentUsername) + 2), CurrentUsername & ", ", _
                  vbTextCompare) = 0))) Then
            
            ' remove username (and colon/comma) from message
            tmpmsg = Mid$(tmpmsg, Len(CurrentUsername) + 3)
        Else
            ' allow commands without any command identifier if
            ' commands are sent via whisper
            'If (Not (WhisperedIn)) Then
            '    ' return negative result indicating that message does not contain
            '    ' a valid command identifier
            '    ProcessCommand = False
            '
            '    ' exit function
            '    Exit Function
            'End If
            
            ' return negative result indicating that message does not contain
            ' a valid command identifier
            ProcessCommand = False
        
            ' exit function
            Exit Function
        End If
    Else
        ' remove slash (/) from in-console message
        tmpmsg = Mid$(tmpmsg, 2)
        
        ' check for second slash indicating
        ' public output
        If (Left$(tmpmsg, 1) = "/") Then
            ' enable public display of command
            PublicOutput = True
        
            ' remove second slash (/) from in-console
            ' message
            tmpmsg = Mid$(tmpmsg, 2)
        End If
    End If

    ' check for multiple command syntax if not issued from
    ' within the console
    If ((Not (InBot)) And _
        (InStr(1, tmpmsg, "; ", vbBinaryCompare) > 0)) Then
       
        Dim X() As String  ' ...
    
        ' split message
        X = Split(tmpmsg, "; ")
        
        ' loop through commands
        For i = 0 To UBound(X)
            Dim tmpX As String ' ...
            
            ' store local copy of message
            tmpX = X(i)
        
            ' can we check for a command identifier without
            ' causing an rte?
            If (Len(tmpX) >= Len(BotVars.TriggerLong)) Then
                ' check for presence of command identifer
                If (Left$(tmpX, Len(BotVars.TriggerLong)) = BotVars.TriggerLong) Then
                    ' remove command identifier from message
                    tmpX = Mid$(tmpX, Len(BotVars.TriggerLong) + 1)
                End If
            End If
        
            ' execute command
            ProcessCommand = ExecuteCommand(Username, GetCumulativeAccess(Username, _
                "USER"), tmpX, InBot, cmdRet())
            
            If (ProcessCommand) Then
                ' display command response
                If (cmdRet(0) <> vbNullString) Then
                    Dim j As Integer ' ...
                
                    ' loop through command response
                    For j = 0 To UBound(cmdRet)
                        If ((InBot) And (Not (PublicOutput))) Then
                            ' display message on screen
                            Call frmChat.AddChat(RTBColors.ConsoleText, cmdRet(j))
                        Else
                            ' send message to battle.net
                            If (WhisperedIn) Then
                                ' whisper message
                                Call AddQ("/w " & Username & _
                                    Space(1) & cmdRet(j), 2, Username)
                            Else
                                Call AddQ(cmdRet(j), 2, Username)
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
    Else
        ' send command to main processor
        If (InBot) Then
            ' execute command
            ProcessCommand = ExecuteCommand(Username, ConsoleAccessResponse, tmpmsg, _
                InBot, cmdRet())
        Else
            ' execute command
            ProcessCommand = ExecuteCommand(Username, GetCumulativeAccess(Username, _
                "USER"), tmpmsg, InBot, cmdRet())
        End If
        
        If (ProcessCommand) Then
            ' display command response
            If (cmdRet(0) <> vbNullString) Then
                ' loop through command response
                For i = 0 To UBound(cmdRet)
                    If ((InBot) And (Not (PublicOutput))) Then
                        ' display message on screen
                        Call frmChat.AddChat(RTBColors.ConsoleText, cmdRet(i))
                    Else
                        ' display message
                        If ((WhisperedIn) Or _
                           ((BotVars.WhisperCmds) And (Not (InBot)))) Then
                           
                            ' whisper message
                            Call AddQ("/w " & Username & _
                                Space(1) & cmdRet(i), 2, Username)
                        Else
                            Call AddQ(cmdRet(i), 2, Username)
                        End If
                    End If
                Next i
            End If
        Else
            ' send command directly to Battle.net if
            ' command is found to be invalid and issued
            ' internally
            If (InBot) Then
                Call AddQ(Message, 0, "(console)")
            End If
        End If
    End If
    
    ' break out of function before reaching error
    ' handler
    Exit Function
    
' default (if all else fails) error handler to keep erroneous
' commands and/or input formats from killing me
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: Command processor has encountered an error.")

    ' return command failure result
    ProcessCommand = False
    
    Exit Function
End Function ' end function ProcessCommand

' command processing helper function
Public Function ExecuteCommand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal Message As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpmsg   As String  ' stores copy of message
    Dim cmdName  As String  ' stores command name
    Dim msgData  As String  ' stores unparsed command parameters
    Dim blnNoCmd As Boolean ' stores result of command switch (true = no command found)
    Dim i        As Integer ' loop counter
    
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
    
    ' convert command name to lcase
    cmdName = LCase$(cmdName)
    
    ' check for & convert aliases to officially
    ' supported command names
    cmdName = convertAlias(cmdName)

    ' initial access check
    If ((ValidateAccess(dbAccess, cmdName) = True) Or _
        (InBot = True)) Then
        
        ' command switch
        Select Case (cmdName)
            Case "quit":         Call OnQuit(Username, dbAccess, msgData, InBot, cmdRet())
            Case "locktext":     Call OnLockText(Username, dbAccess, msgData, InBot, cmdRet())
            Case "allowmp3":     Call OnAllowMp3(Username, dbAccess, msgData, InBot, cmdRet())
            Case "loadwinamp":   Call OnLoadWinamp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "efp":          Call OnEfp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "home":         Call OnHome(Username, dbAccess, msgData, InBot, cmdRet())
            Case "clan":         Call OnClan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "peonban":      Call OnPeonBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "invite":       Call OnInvite(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setmotd":      Call OnSetMotd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "where":        Call OnWhere(Username, dbAccess, msgData, InBot, cmdRet())
            Case "quiettime":    Call OnQuietTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "roll":         Call OnRoll(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sweepban":     Call OnSweepBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sweepignore":  Call OnSweepIgnore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setname":      Call OnSetName(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setpass":      Call OnSetPass(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setkey":       Call OnSetKey(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setexpkey":    Call OnSetExpKey(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setserver":    Call OnSetServer(Username, dbAccess, msgData, InBot, cmdRet())
            Case "giveup":       Call OnGiveUp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "math":         Call OnMath(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idlebans":     Call OnIdleBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "chpw":         Call OnChPw(Username, dbAccess, msgData, InBot, cmdRet())
            Case "join":         Call OnJoin(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sethome":      Call OnSetHome(Username, dbAccess, msgData, InBot, cmdRet())
            Case "resign":       Call OnResign(Username, dbAccess, msgData, InBot, cmdRet())
            Case "clearbanlist": Call OnClearBanList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "kickonyell":   Call OnKickOnYell(Username, dbAccess, msgData, InBot, cmdRet())
            Case "rejoin":       Call OnRejoin(Username, dbAccess, msgData, InBot, cmdRet())
            Case "plugban":      Call OnPlugBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "clientbans":   Call OnClientBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setvol":       Call OnSetVol(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cadd":         Call OnCAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cdel":         Call OnCDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "banned":       Call OnBanned(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ipbans":       Call OnIPBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ipban":        Call OnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unipban":      Call OnUnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "designate":    Call OnDesignate(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shuffle":      Call OnShuffle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "repeat":       Call OnRepeat(Username, dbAccess, msgData, InBot, cmdRet())
            Case "next":         Call OnNext(Username, dbAccess, msgData, InBot, cmdRet())
            Case "protect":      Call OnProtect(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whispercmds":  Call OnWhisperCmds(Username, dbAccess, msgData, InBot, cmdRet())
            Case "stop":         Call OnStop(Username, dbAccess, msgData, InBot, cmdRet())
            Case "play":         Call OnPlay(Username, dbAccess, msgData, InBot, cmdRet())
            Case "useitunes":    Call OnUseiTunes(Username, dbAccess, msgData, InBot, cmdRet())
            Case "usewinamp":    Call OnUseWinamp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "pause":        Call OnPause(Username, dbAccess, msgData, InBot, cmdRet())
            Case "fos":          Call OnFos(Username, dbAccess, msgData, InBot, cmdRet())
            Case "rem":          Call OnRem(Username, dbAccess, msgData, InBot, cmdRet())
            Case "reconnect":    Call OnReconnect(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unigpriv":     Call OnUnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
            Case "igpriv":       Call OnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
            Case "block":        Call OnBlock(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idletime":     Call OnIdleTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idle":         Call OnIdle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitdel":      Call OnShitDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safedel":      Call OnSafeDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagdel":       Call OnTagDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setidle":      Call OnSetIdle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idletype":     Call OnIdleType(Username, dbAccess, msgData, InBot, cmdRet())
            Case "filter":       Call OnFilter(Username, dbAccess, msgData, InBot, cmdRet())
            Case "trigger":      Call OnTrigger(Username, dbAccess, msgData, InBot, cmdRet())
            Case "settrigger":   Call OnSetTrigger(Username, dbAccess, msgData, InBot, cmdRet())
            Case "levelban":     Call OnLevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "d2levelban":   Call OnD2LevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "phrasebans":   Call OnPhraseBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "pon":          Call OnPhraseBans(Username, dbAccess, "on", InBot, cmdRet())
            Case "poff":         Call OnPhraseBans(Username, dbAccess, "off", InBot, cmdRet())
            Case "pstatus":      Call OnPhraseBans(Username, dbAccess, vbNullString, InBot, cmdRet())
            Case "setpmsg":      Call OnSetPMsg(Username, dbAccess, msgData, InBot, cmdRet())
            Case "phrases":      Call OnPhrases(Username, dbAccess, msgData, InBot, cmdRet())
            Case "addphrase":    Call OnAddPhrase(Username, dbAccess, msgData, InBot, cmdRet())
            Case "delphrase":    Call OnDelPhrase(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagadd":       Call OnTagAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "fadd":         Call OnFAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "frem":         Call OnFRem(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safelist":     Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safeadd":      Call OnSafeAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safecheck":    Call OnSafeCheck(Username, dbAccess, msgData, InBot, cmdRet())
            Case "exile":        Call OnExile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unexile":      Call OnUnExile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitlist":     Call OnShitList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagbans":      Call OnTagBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitadd":      Call OnShitAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "dnd":          Call OnDND(Username, dbAccess, msgData, InBot, cmdRet())
            Case "bancount":     Call OnBanCount(Username, dbAccess, msgData, InBot, cmdRet())
            Case "banlistcount": Call OnBanListCount(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagcheck":     Call OnTagCheck(Username, dbAccess, msgData, InBot, cmdRet())
            Case "slcheck":      Call OnSLCheck(Username, dbAccess, msgData, InBot, cmdRet())
            Case "readfile":     Call OnReadFile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "greet":        Call OnGreet(Username, dbAccess, msgData, InBot, cmdRet())
            Case "allseen":      Call OnAllSeen(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ban":          Call OnBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unban":        Call OnUnban(Username, dbAccess, msgData, InBot, cmdRet())
            Case "kick":         Call OnKick(Username, dbAccess, msgData, InBot, cmdRet())
            Case "lastwhisper":  Call OnLastWhisper(Username, dbAccess, msgData, InBot, cmdRet())
            Case "say":          Call OnSay(Username, dbAccess, msgData, InBot, cmdRet())
            Case "expand":       Call OnExpand(Username, dbAccess, msgData, InBot, cmdRet())
            Case "detail":       Call OnDetail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "info":         Call OnInfo(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shout":        Call OnShout(Username, dbAccess, msgData, InBot, cmdRet())
            Case "voteban":      Call OnVoteBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "votekick":     Call OnVoteKick(Username, dbAccess, msgData, InBot, cmdRet())
            Case "vote":         Call OnVote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tally":        Call OnTally(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cancel":       Call OnCancel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "back":         Call OnBack(Username, dbAccess, msgData, InBot, cmdRet())
            Case "prev":         Call OnPrev(Username, dbAccess, msgData, InBot, cmdRet())
            Case "uptime":       Call OnUptime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "away":         Call OnAway(Username, dbAccess, msgData, InBot, cmdRet())
            Case "mp3":          Call OnMP3(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ping":         Call OnPing(Username, dbAccess, msgData, InBot, cmdRet())
            Case "addquote":     Call OnAddQuote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "owner":        Call OnOwner(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ignore":       Call OnIgnore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "quote":        Call OnQuote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unignore":     Call OnUnignore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cq":           Call OnCQ(Username, dbAccess, msgData, InBot, cmdRet())
            Case "scq":          Call OnSCQ(Username, dbAccess, msgData, InBot, cmdRet())
            Case "time":         Call OnTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "getping":      Call OnGetPing(Username, dbAccess, msgData, InBot, cmdRet())
            Case "checkmail":    Call OnCheckMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "getmail":      Call OnGetMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whoami":       Call OnWhoAmI(Username, dbAccess, msgData, InBot, cmdRet())
            Case "add":          Call OnAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "mmail":        Call OnMMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "bmail":        Call OnBMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "designated":   Call OnDesignated(Username, dbAccess, msgData, InBot, cmdRet())
            Case "flip":         Call OnFlip(Username, dbAccess, msgData, InBot, cmdRet())
            Case "about":        Call OnAbout(Username, dbAccess, msgData, InBot, cmdRet())
            Case "server":       Call OnServer(Username, dbAccess, msgData, InBot, cmdRet())
            Case "find":         Call OnFind(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whois":        Call OnWhoIs(Username, dbAccess, msgData, InBot, cmdRet())
            Case "findattr":     Call OnFindAttr(Username, dbAccess, msgData, InBot, cmdRet())
            Case "findgrp":      Call OnFindGrp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "monitor":      Call OnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unmonitor":    Call OnUnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
            Case "online":       Call OnOnline(Username, dbAccess, msgData, InBot, cmdRet())
            Case "help":         Call OnHelp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "promote":      Call OnPromote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "demote":       Call OnDemote(Username, dbAccess, msgData, InBot, cmdRet())
            Case Else
                blnNoCmd = True
        End Select
    
        ' ...
        If (Not (blnNoCmd)) Then
            ' append entry to command log
            Call LogCommand(Username, Message)
        End If
        
        ' was a command found? return.
        ExecuteCommand = (Not (blnNoCmd))
    Else
        ' return false result, as user does not have sufficient
        ' access to issue the requested command
        ExecuteCommand = False
    End If
End Function

' handle quit command
Private Function OnQuit(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will initiate the bot's termination sequence.
    
    Call frmChat.Form_Unload(0)
End Function ' end function OnQuit

' handle locktext command
Private Function OnLockText(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will prevent chat messages from displaying on the bot's screen.
    
    Call frmChat.mnuLock_Click
End Function ' end function OnLockText

' handle allowmp3 command
Private Function OnAllowMp3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable or disable the use of media player-related commands.
    
    Dim tmpBuf As String ' temporary output buffer

    If (BotVars.DisableMP3Commands) Then
        tmpBuf = "Allowing MP3 commands."
        
        BotVars.DisableMP3Commands = False
    Else
        tmpBuf = "MP3 commands are now disabled."
        
        BotVars.DisableMP3Commands = True
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAllowMp3

' handle loadwinamp command
Private Function OnLoadWinamp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will run Winamp from the default directory, or the directory
    ' specified within the configuration file.
    
    Dim clsWinamp As clsWinamp
    Dim tmpBuf    As String  ' temporary output buffer
    Dim bln       As Boolean ' ...
    
    ' ...
    Set clsWinamp = New clsWinamp

    ' ...
    bln = clsWinamp.OpenPlayer(ReadCFG("Other", "WinampPath"))
       
    ' ...
    If (bln) Then
        tmpBuf = "Winamp loaded."
    Else
        tmpBuf = "There was an error loading Winamp."
    End If
    
    ' ...
    Set clsWinamp = Nothing
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLoadWinamp

' handle efp command
Private Function OnEfp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable, or check the status of, the Effective Floodbot
    ' Protection system.  EFP is a system designed to combat floodbot attacks by reducing
    ' the number of allowable commands, and enhancing the strength of the message queue.
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)

    If (msgData = "on") Then
        ' enable efp
        Call frmChat.SetFloodbotMode(1)
        
        tmpBuf = "Emergency floodbot protection enabled."
    ElseIf (msgData = "off") Then
        ' disable efp
        Call frmChat.SetFloodbotMode(0)
        
        tmpBuf = "Emergency floodbot protection disabled."
    ElseIf (msgData = "status") Then
        If (bFlood) Then
            frmChat.AddChat RTBColors.TalkBotUsername, "Emergency floodbot protection is " & _
                "enabled. (No messages can be sent to battle.net.)"
        Else
            tmpBuf = "Emergency floodbot protection is disabled."
        End If
    End If
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnEfp

' handle home command
Private Function OnHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot join its home channel.
    
    Call AddQ("/join " & BotVars.HomeChannel, 1, Username)
End Function ' end function OnHome

' handle clan command
Private Function OnClan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will allow the use of Battle.net's /clan command without requiring
    ' users be given the ability to use the bot's say command.
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' is bot a channel operator?
    If ((MyFlags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        Select Case (LCase$(msgData))
            Case "public", "pub"
                tmpBuf = "Clan channel is now public."
                
                ' set clan channel to public
                Call AddQ("/clan public", 1, Username)
                
            Case "private", "priv"
                tmpBuf = "Clan channel is now private."
                
                ' set clan channel to private
                Call AddQ("/clan private", 1, Username)
                
            Case Else
               ' set clan channel to specified
               Call AddQ("/clan " & msgData, , Username)
        End Select
    Else
        tmpBuf = "The bot must have ops to change clan privacy status."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnClan

' handle peonban command
Private Function OnPeonBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable, or check the status of, WarCraft III peon
    ' banning.  The "Peon" class is defined by Battle.net, and is currently the lowest
    ' ranking WarCraft III user classification, in which, users have less than twenty-five
    ' wins on record for any given race.
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)
    
    If (msgData = "on") Then
        ' enable peon banning
        BotVars.BanPeons = 1
        
        ' write configuration entry
        Call WriteINI("Other", "PeonBans", "Y")
        
        tmpBuf = "Peon banning activated."
    ElseIf (msgData = "off") Then
        ' disable peon banning
        BotVars.BanPeons = 0
        
        ' write configuration entry
        Call WriteINI("Other", "PeonBans", "N")
        
        tmpBuf = "Peon banning deactivated."
    ElseIf (msgData = "status") Then
        tmpBuf = "The bot is currently "
        
        If (BotVars.BanPeons = 0) Then
            tmpBuf = tmpBuf & "not banning peons."
        Else
            tmpBuf = tmpBuf & "banning peons."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPeonBan

' handle invite command
Private Function OnInvite(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will send an invitation to the specified user to join the
    ' clan that the bot is currently either a Shaman or Chieftain of.  This
    ' command will only work if the bot is logged on using WarCraft III, and
    ' is either a Shaman, or a Chieftain of the clan in question.
    
    Dim tmpBuf As String ' temporary output buffer

    ' are we using warcraft iii?
    If (IsW3) Then
        ' is my ranking sufficient to issue
        ' an invitation?
        If (Clan.MyRank >= 3) Then
            Call InviteToClan(msgData)
            
            tmpBuf = msgData & ": Clan invitation sent."
        Else
            tmpBuf = "Error: The bot must hold Shaman or Chieftain rank to invite users."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnInvite

' handle setmotd command
Private Function OnSetMotd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the clan channel's Message Of The Day.  This
    ' command will only work if the bot is logged on using WarCraft III,
    ' and is either a Shaman or a Chieftain of the clan in question.
    
    Dim tmpBuf As String ' temporary output buffer

    If (IsW3) Then
        If (Clan.MyRank >= 3) Then
            Call SetClanMOTD(msgData)
            
            tmpBuf = "Clan MOTD set."
        Else
            tmpBuf = "Error: Shaman or Chieftain rank is required to set the MOTD."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetMotd

' handle where command
Private Function OnWhere(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will state the channel that the bot is currently
    ' residing in.  Battle.net uses this command to display basic
    ' user data, such as game type, and channel or game name.
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' if sent from within the bot, send "where" command
    ' directly to Battle.net
    If (InBot) Then
        Call AddQ("/where " & msgData, , "(console)")
    End If

    tmpBuf = "I am currently in channel " & gChannel.Current & " (" & _
        colUsersInChannel.Count & " users present)"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhere

' handle quiettime command
Private Function OnQuietTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable, disable or check the status of, quiet time.
    ' Quiet time is a feature that will ban non-safelisted users from the
    ' channel when they speak publicly within the channel.  This is useful
    ' when a channel wishes to have a discussion while allowing public
    ' attendance, but disallowing public participation.
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)
    
    If (msgData = "on") Then
        ' enable quiettime
        BotVars.QuietTime = True
    
        ' write configuration entry
        Call WriteINI("Main", "QuietTime", "Y")
        
        tmpBuf = "Quiet-time enabled."
    ElseIf (msgData = "off") Then
        ' disable quiettime
        BotVars.QuietTime = False
        
        ' write configuration entry
        Call WriteINI("Main", "QuietTime", "N")
        
        tmpBuf = "Quiet-time disabled."
    ElseIf (msgData = "status") Then
        If (BotVars.QuietTime) Then
            tmpBuf = "Quiet-time is currently enabled."
        Else
            tmpBuf = "Quiet-time is currently disabled."
        End If
    Else
        tmpBuf = "Error: Invalid arguments."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnQuietTime

' handle roll command
Private Function OnRoll(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will state a random number between a range of zero to
    ' one-hundred, or from zero to any specified number.
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim iWinamp As Long
    Dim Track   As Long

    If (Len(msgData) = 0) Then
        Randomize
        
        iWinamp = CLng(Rnd * 100)
        
        tmpBuf = "Random number (0-100): " & iWinamp
    Else
        Randomize
        
        If (StrictIsNumeric(msgData)) Then
            If (Val(msgData) < 100000000) Then
                Track = CLng(Rnd * CLng(msgData))
                
                tmpBuf = "Random number (0-" & msgData & "): " & Track
            Else
                tmpBuf = "Invalid value."
            End If
        Else
            tmpBuf = "Error: Invalid value."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnRoll

' handle sweepban command
Private Function OnSweepBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will grab the listing of users in the specified channel
    ' using Battle.net's "who" command, and will then begin banning each
    ' user from the current channel using Battle.net's "ban" command.
    
    Dim u As String
    Dim Y As String

    Caching = True
    
    Call Cache(vbNullString, 255, "ban ")
    
    Call AddQ("/who " & msgData, 1, Username)
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
    
    Dim u As String
    Dim Y As String
    
    Caching = True
    
    Call Cache(vbNullString, 255, "squelch ")
    
    Call AddQ("/who " & msgData, 1, Username)
End Function ' end function OnSweepIgnore

' handle setname command
Private Function OnSetName(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the username that the bot uses to connect with
    ' to the specified value.
    
    Dim tmpBuf As String ' temporary output buffer

    ' are we using a beta?
    #If BETA = 1 Then
        ' only allow use of setname command while on-line to prevent beta
        ' authorization bypassing
        If ((Not (g_Online = True)) Or (g_Connected = False)) Then
            Exit Function
        End If
    #End If

    ' write configuration entry
    Call WriteINI("Main", "Username", msgData)
    
    ' set username
    BotVars.Username = msgData
    
    tmpBuf = "New username set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetName

' handle setpass command
Private Function OnSetPass(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the password that the bot uses to connect with
    ' to the specified value.
    
    Dim tmpBuf As String ' temporary output buffer

    ' write configuration entry
    Call WriteINI("Main", "Password", msgData)
    
    ' set password
    BotVars.Password = msgData
    
    tmpBuf = "New password set."
    
    ' return message
    cmdRet(0) = tmpBuf
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
    
    Dim tmpBuf As String ' temporary output buffer

    If (Len(msgData)) Then
        If (InStr(1, msgData, "CreateObject", vbTextCompare)) Then
            ' CreateObject() is a no no, because of
            ' its use in exploits.
            
            tmpBuf = "Evaluation error."
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
                tmpBuf = "The statement " & Chr$(34) & _
                    msgData & Chr$(34) & " evaluates to: " & _
                        res & "."
            Else
                tmpBuf = "Evaluation error."
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
    
    Exit Function
    
ERROR_HANDLER:
    tmpBuf = "Evaluation error."

    ' return message
    cmdRet(0) = tmpBuf
    
    Exit Function
End Function ' end function OnMath

' handle setkey command
Private Function OnSetKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the bot's CD-Key to the CD-Key specified.
    
    Dim tmpBuf As String ' temporary output buffer

    ' clean data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)

    ' write configuration information
    Call WriteINI("Main", "CDKey", msgData)
    
    ' set CD-Key
    BotVars.CDKey = msgData
    
    tmpBuf = "New cdkey set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetKey

' handle setexpkey command
Private Function OnSetExpKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the bot's expansion CD-Key to the expansion
    ' CD-Key specified.
    
    Dim tmpBuf As String ' temporary output buffer

    ' sanitize data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)
    
    ' write configuration entry
    Call WriteINI("Main", "ExpKey", msgData)
    
    ' set expansion CD-Key
    BotVars.ExpKey = msgData
    
    tmpBuf = "New expansion CD-key set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetExpKey

' handle setserver command
Private Function OnSetServer(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the server that the bot connects to to the value
    ' specified.
    
    Dim tmpBuf As String ' temporary output buffer
    
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
    
    tmpBuf = "New server set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetServer

' handle giveup command
Private Function OnGiveUp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will allow a user to designate a specified user using
    ' Battle.net's "designate" command, and will then make the bot resign
    ' its status as a channel moderator.  This command is useful if you are
    ' lazy and you just wish to designate someone as quickly as possible.
    
    ' ...
    If (checkChannel(msgData)) Then
        Dim i          As Integer ' ...
        Dim arrUsers() As String  ' ...
        Dim userCount  As Integer ' ...
    
        ' ...
        If ((MyFlags And USER_CHANNELOP&) <> USER_CHANNELOP&) Then
            ' ...
            cmdRet(0) = "Error: This command requires channel " & _
                "operator status."
        
            Exit Function
        End If
    
        ' ...
        If (StrComp(gChannel.Current, "Clan " & Clan.Name, vbTextCompare) = 0) Then
            ' ...
            ReDim Preserve arrUsers(0)
            
            ' ...
            If (Clan.MyRank >= 4) Then
                ' ...
                For i = 1 To frmChat.lvClanList.ListItems.Count
                    ' ...
                    If (StrComp(frmChat.lvClanList.ListItems(i).text, _
                        CurrentUsername, vbTextCompare) <> 0) Then
                        
                        ' ...
                        If (frmChat.lvClanList.ListItems(i).SmallIcon = 3) Then
                            ' ...
                            arrUsers(userCount) = _
                                frmChat.lvClanList.ListItems(i).text
        
                            ' ...
                            userCount = (userCount + 1)
        
                            ' ...
                            ReDim Preserve arrUsers(0 To userCount)
                        End If
                    End If
                Next i
            End If
            
            ' ...
            If (userCount) Then
                ' demote shamans
                For i = 0 To (userCount - 1)
                    ' ...
                    Call frmChat.AddChat(vbRed, "DEMOTE: " & arrUsers(i))
                
                    ' ...
                    With PBuffer
                        .InsertDWord &H1
                        .InsertNTString arrUsers(i)
                        .InsertByte &H2 ' General member (Grunt)
                        .SendPacket &H7A
                    End With
                Next i
            End If
        End If
        
        ' designate user
        Call AddQ("/designate " & msgData, 0, Username)
        
        ' rejoin channel
        Call AddQ("/resign", 0, Username)
        
        ' ...
        If (userCount) Then
            ' promote shamans again
            For i = 0 To (userCount - 1)
                ' ...
                Call frmChat.AddChat(vbRed, "PROMOTE: " & arrUsers(i))
            
                ' ...
                With PBuffer
                    .InsertDWord &H3
                    .InsertNTString arrUsers(i)
                    .InsertByte &H3 ' Officer (Shaman)
                    .SendPacket &H7A
                End With
            Next i
        End If
        
        ' ...
        ReDim arrUsers(0)
    Else
        ' ...
        cmdRet(0) = "Error: The specified user is not present " & _
            "within the channel."
    End If
End Function ' end function OnGiveUp

' handle idlebans command
Private Function OnIdleBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String ' ...
    Dim tmpBuf     As String ' temporary output buffer
    Dim subcmd     As String ' ...
    Dim index      As Long   ' ...
    Dim tmpData    As String ' ...
    
    tmpData = msgData
    
    If (Len(tmpData) > 0) Then
        index = InStr(1, tmpData, Space$(1), vbBinaryCompare)
    
        If (index <> 0) Then
            subcmd = Mid$(tmpData, 1, index - 1)
        Else
            subcmd = tmpData
        End If
        
        subcmd = LCase$(subcmd)
        
        If (index) Then
            tmpData = Mid$(msgData, index + 1)
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
                    tmpBuf = "IdleBans activated, with a delay of " & BotVars.IB_Wait & "."
                    
                    Call WriteINI("Other", "IdleBans", "Y")
                    Call WriteINI("Other", "IdleBanDelay", BotVars.IB_Wait)
                Else
                    BotVars.IB_Wait = 400
                    
                    tmpBuf = "IdleBans activated, using the default delay of 400."
                    
                    Call WriteINI("Other", "IdleBanDelay", "400")
                    Call WriteINI("Other", "IdleBans", "Y")
                End If
                
            Case "off"
                BotVars.IB_On = BFALSE
                
                tmpBuf = "IdleBans deactivated."
                
                Call WriteINI("Other", "IdleBans", "N")
                
            Case "kick"
                If (Len(tmpData) > 0) Then
                    Select Case (LCase$(tmpData))
                        Case "on"
                            tmpBuf = "Idle users will now be kicked instead of banned."
                            
                            Call WriteINI("Other", "KickIdle", "Y")
                            
                            BotVars.IB_Kick = True
                            
                        Case "off"
                            tmpBuf = "Idle users will now be banned instead of kicked."
                            
                            Call WriteINI("Other", "KickIdle", "N")
                            
                            BotVars.IB_Kick = False
                            
                        Case Else
                            tmpBuf = "Error: Unknown idle kick setting."
                    End Select
                Else
                    tmpBuf = "Error: Too few arguments."
                End If
        
            Case "wait", "delay"
                If (StrictIsNumeric(tmpData)) Then
                    BotVars.IB_Wait = CInt(tmpData)
                    
                    tmpBuf = "IdleBan delay set to " & BotVars.IB_Wait & "."
                    
                    Call WriteINI("Other", "IdleBanDelay", CInt(tmpData))
                Else
                    tmpBuf = "Error: IdleBan delays require a numeric value."
                End If
                
            Case "status"
                If (BotVars.IB_On = BTRUE) Then
                    tmpBuf = IIf(BotVars.IB_Kick, "Kicking", "Banning") & _
                        " users who are idle for " & BotVars.IB_Wait & "+ seconds."
                Else
                    tmpBuf = "IdleBans are disabled."
                End If
                
            Case Else
                tmpBuf = "Error: Invalid command."
        End Select
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdleBans

' handle chpw command
Private Function OnChPw(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String ' ...
    Dim tmpBuf     As String ' temporary output buffer
    
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
                
                tmpBuf = "Channel password protection enabled, delay set to " & _
                    BotVars.ChannelPasswordDelay & "."
            Else
                tmpBuf = "Error: Invalid channel password."
            End If
            
        Case "off", "kill", "clear"
            BotVars.ChannelPassword = vbNullString
            
            BotVars.ChannelPasswordDelay = 0
            
            tmpBuf = "Channel password protection disabled."
            
        Case "time", "delay", "wait"
            If (StrictIsNumeric(strArray(1))) Then
                If ((Val(strArray(1)) <= 255) And _
                    (Val(strArray(1)) >= 1)) Then
                   
                    BotVars.ChannelPasswordDelay = CByte(Val(strArray(1)))
                    
                    tmpBuf = "Channel password delay set to " & strArray(1) & "."
                Else
                    tmpBuf = "Error: Invalid channel delay."
                End If
            Else
                tmpBuf = "Error: Invalid channel delay."
            End If
            
        Case "info", "status"
            If ((BotVars.ChannelPassword = vbNullString) Or _
                (BotVars.ChannelPasswordDelay = 0)) Then
                
                tmpBuf = "Channel password protection is disabled."
            Else
                tmpBuf = "Channel password protection is enabled. Password [" & _
                    BotVars.ChannelPassword & "], Delay [" & _
                        BotVars.ChannelPasswordDelay & "]."
            End If
            
        Case Else
            tmpBuf = "Error: Unknown channel password command."
    End Select
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnChPw

' handle join command
Private Function OnJoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot join the specified channel.
    
    Dim tmpBuf As String ' temporary output buffer

    If (LenB(msgData) > 0) Then
        AddQ "/join " & msgData, 1, Username
    End If
End Function ' end function OnJoin

' handle sethome command
Private Function OnSetHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the home channel to the channel specified.
    ' The home channel is the channel that the bot joins immediately
    ' following a completion of the connection procedure.
    
    Dim tmpBuf As String ' temporary output buffer

    Call WriteINI("Main", "HomeChan", msgData)
    
    BotVars.HomeChannel = msgData
    
    tmpBuf = "Home channel set to [ " & msgData & " ]"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetHome

' handle resign command
Private Function OnResign(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot resign from the role of operator
    ' by rejoining the channel through the use of Battle.net's /resign
    ' command.
    
    Call AddQ("/resign", 1, Username)
End Function ' end function OnResign

' handle clearbanlist
Private Function OnClearBanList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will clear the bot's internal list of banned users.
    
    Dim tmpBuf As String ' temporary output buffer

    ' redefine array size
    ReDim gBans(0)
    
    tmpBuf = "Banned user list cleared."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnClearBanList

' handle kickonyell command
Private Function OnKickOnYell(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    Select Case (LCase$(msgData))
        Case "on"
            BotVars.KickOnYell = 1
            
            tmpBuf = "Kick-on-yell enabled."
            
            Call WriteINI("Other", "KickOnYell", "Y")
            
        Case "off"
            BotVars.KickOnYell = 0
            
            tmpBuf = "Kick-on-yell disabled."
            
            Call WriteINI("Other", "KickOnYell", "N")
            
        Case "status"
            tmpBuf = "Kick-on-yell is "
            tmpBuf = tmpBuf & IIf(BotVars.KickOnYell = 1, "enabled", "disabled") & "."
    End Select
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnKickOnYell

' handle rejoin command
Private Function OnRejoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot rejoin the current channel.
    
    ' join temporary channel
    Call AddQ("/join " & CurrentUsername & " Rejoin", 1, Username)
    
    ' rejoin previous channel
    Call AddQ("/join " & gChannel.Current, 1, Username)
End Function ' end function OnRejoin

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
    
    Dim tmpBuf As String ' temporary output buffer

    ' ...
    msgData = LCase$(msgData)

    Select Case (msgData)
        Case "on"
            Dim i As Integer
        
            If (BotVars.PlugBan) Then
                tmpBuf = "PlugBan is already activated."
            Else
                BotVars.PlugBan = True
                
                tmpBuf = "PlugBan activated."
                
                For i = 1 To colUsersInChannel.Count
                    With colUsersInChannel.Item(i)
                        If ((.Flags = 16) And (Not .Safelisted)) Then
                            AddQ "/ban " & .Username & " PlugBan", 1
                        End If
                    End With
                Next i
                
                Call WriteINI("Other", "PlugBans", "Y")
            End If
            
        Case "off"
            If (BotVars.PlugBan) Then
                BotVars.PlugBan = False
                
                tmpBuf = "PlugBan deactivated."
                
                Call WriteINI("Other", "PlugBans", "N")
            Else
                tmpBuf = "PlugBan is already deactivated."
            End If
            
        Case "status"
            If (BotVars.PlugBan) Then
                tmpBuf = "PlugBan is activated."
            Else
                tmpBuf = "PlugBan is deactivated."
            End If
    End Select

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPlugBan

' handle clientbans command
Private Function OnClientBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase(tmpBuf(), , , , "GAME", , , "B")
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnClientBans

' handle setvol command
Private Function OnSetVol(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the volume of the media player to the level
    ' specified by the user.
    
    Dim tmpBuf As String ' temporary output buffer
    Dim lngVol As Long   ' ...

    If (BotVars.DisableMP3Commands = False) Then
        If (StrictIsNumeric(msgData)) Then
            lngVol = CLng(msgData)
        
            If (lngVol > 100) Then
                lngVol = 100
            End If
            
            MediaPlayer.Volume = lngVol

            tmpBuf = "Volume set to " & msgData & "%."
        Else
            tmpBuf = "Error: Invalid volume level (0-100)."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetVol

' handle cadd command
Private Function OnCAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim index    As Integer

    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' ...
    index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (index) Then
        Dim user As String ' ...
        
        ' ...
        user = Mid$(msgData, 1, index - 1)
        
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            tmpBuf(0) = "Error: The specified game name is invalid."
        Else
            Dim bmsg As String ' ...
            
            ' ...
            bmsg = Mid$(msgData, index + 1)
        
            ' ...
            Call OnAdd(Username, dbAccess, user & " +B --type GAME --banmsg " & bmsg, True, tmpBuf())
        End If
    Else
        ' ...
        Call OnAdd(Username, dbAccess, msgData & " +B --type GAME", True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnCAdd

' handle cdel command
Private Function OnCDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpBuf(0) = "Error: The specified game name is invalid."
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, msgData & " -B --type GAME", True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnCDel

' handle banned command
Private Function OnBanned(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will display a listing of all of the users that have been
    ' banned from the channel since the time of having joined the channel.
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim BanCount As Integer
    Dim i        As Integer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)

    tmpBuf(tmpCount) = "Banned users: "
    
    For i = LBound(gBans) To UBound(gBans)
        If (gBans(i).Username <> vbNullString) Then
            tmpBuf(tmpCount) = tmpBuf(tmpCount) & ", " & gBans(i).Username
            
            If ((Len(tmpBuf(tmpCount)) > 90) And (i <> UBound(gBans))) Then
                ' increase array size
                ReDim Preserve tmpBuf(tmpCount + 1)
            
                ' apply postfix to previous line
                tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1)) & " [more]"
                
                ' apply prefix to new line
                tmpBuf(tmpCount + 1) = "Banned users: "
                
                ' incrememnt counter
                tmpCount = (tmpCount + 1)
            End If
            
            ' incrememnt counter
            BanCount = (BanCount + 1)
        End If
    Next i

    ' has anyone been banned?
    If (BanCount = 0) Then
        tmpBuf(tmpCount) = "No users have been banned."
    Else
        tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1))
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnBanned

' handle ipbans command
Private Function OnIPBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    msgData = LCase$(msgData)

    If (Left$(msgData, 2)) = "on" Then
        BotVars.IPBans = True
        
        Call WriteINI("Other", "IPBans", "Y")
        
        tmpBuf = "IPBanning activated."
        
        If ((MyFlags = 2) Or (MyFlags = 18)) Then
            For i = 1 To colUsersInChannel.Count
                Select Case colUsersInChannel.Item(i).Flags
                    Case 20, 30, 32, 48
                        Call AddQ("/ban " & colUsersInChannel.Item(i).Username & _
                            " IPBanned.", 1)
                End Select
            Next i
        End If
    ElseIf (Left$(msgData, 3) = "off") Then
        BotVars.IPBans = False
        
        Call WriteINI("Other", "IPBans", "N")
        
        tmpBuf = "IPBanning deactivated."
        
    ElseIf (Left$(msgData, 6) = "status") Then
        If (BotVars.IPBans) Then
            tmpBuf = "IPBanning is currently active."
        Else
            tmpBuf = "IPBanning is currently disabled."
        End If
    Else
        tmpBuf = "Error: Unrecognized IPBan command. Use 'on', 'off' or 'status'."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIPBans

' handle ipban command
Private Function OnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim gAcc   As udtGetAccessResponse

    Dim tmpBuf As String ' temporary output buffer
    Dim tmpAcc As String ' ...

    ' ...
    tmpAcc = StripInvalidNameChars(msgData)

    ' ...
    If (Len(tmpAcc) > 0) Then
        ' ...
        If (InStr(1, tmpAcc, "@") > 0) Then
            tmpAcc = StripRealm(msgData)
        End If
        
        ' ...
        If (dbAccess.Access <= 100) Then
            If ((GetSafelist(tmpAcc)) Or (GetSafelist(msgData))) Then
                ' return message
                cmdRet(0) = "Error: That user is safelisted."
                
                Exit Function
            End If
        End If
        
        ' ...
        gAcc = GetAccess(msgData)
        
        ' ...
        If ((gAcc.Access >= dbAccess.Access) Or _
            ((InStr(gAcc.Flags, "A") > 0) And (dbAccess.Access <= 100))) Then

            tmpBuf = "Error: You do not have enough access to do that."
        Else
            Call AddQ("/squelch " & msgData, 1, Username)
        
            tmpBuf = "User " & Chr(34) & msgData & Chr(34) & " IPBanned."
        End If
    Else
        ' return message
        tmpBuf = "Error: You do not have enough access to do that."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIPBan

' handle unipban command
Private Function OnUnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (Len(msgData) > 0) Then
        Call AddQ("/unsquelch " & msgData, 1, Username)
        Call AddQ("/unban " & msgData, 1, Username)
        
        tmpBuf = "User " & Chr(34) & msgData & Chr(34) & " Un-IPBanned."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnIPBan

' handle designate command
Private Function OnDesignate(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If ((MyFlags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
            'diablo 2 handling
            If (Dii = True) Then
                If (Not (Mid$(msgData, 1, 1) = "*")) Then
                    msgData = "*" & msgData
                End If
            End If
            
            Call AddQ("/designate " & msgData, 1, Username)
            
            tmpBuf = "I have designated [ " & msgData & " ]"
        Else
            tmpBuf = "Error: The bot does not have ops."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDesignate

' handle shuffle command
Private Function OnShuffle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will toggle the usage of the selected media player's
    ' shuffling feature.
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long
    
    If (Not (BotVars.DisableMP3Commands)) Then
        tmpBuf = "Winamp's Shuffle feature has been toggled."
        
        'hWndWA = GetWinamphWnd()
        
        If (hWndWA = 0) Then
            tmpBuf = "Winamp is not loaded."
        Else
            'Call SendMessage(hWndWA, WM_COMMAND, WA_TOGGLESHUFFLE, 0)
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShuffle

' handle repeat command
Private Function OnRepeat(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will toggle the usage of the selected media player's
    ' repeat feature.
    
    Dim hWndWA As Long
    Dim tmpBuf As String ' temporary output buffer
    
    If (Not (BotVars.DisableMP3Commands)) Then
        tmpBuf = "Winamp's Repeat feature has been toggled."
        
        'hWndWA = GetWinamphWnd()
        
        If (hWndWA = 0) Then
            tmpBuf = "Winamp is not loaded."
        Else
            'Call SendMessage(hWndWA, WM_COMMAND, WA_TOGGLEREPEAT, 0)
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnRepeat

' handle next command
Private Function OnNext(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long

    If (BotVars.DisableMP3Commands = False) Then
        Dim pos As Integer ' ...
        
        ' ...
        pos = MediaPlayer.PlaylistPosition
    
        ' ...
        Call MediaPlayer.PlayTrack(pos + 1)
        
        ' ...
        tmpBuf = "Skipped forwards."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnNext

' handle protect command
Private Function OnProtect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    Select Case (LCase$(msgData))
        Case "on"
            If ((MyFlags = 2) Or (MyFlags = 18)) Then
                Protect = True
                
                tmpBuf = "Lockdown activated by " & Username & "."
                
                Call WildCardBan("*", ProtectMsg, 1)
                
                Call WriteINI("Main", "Protect", "Y")
            Else
                tmpBuf = "The bot does not have ops."
            End If
        
        Case "off"
            If (Protect) Then
                Protect = False
                
                tmpBuf = "Lockdown deactivated."
                
                Call WriteINI("Main", "Protect", "N")
            Else
                tmpBuf = "Protection was not enabled."
            End If
            
        Case "status"
            Select Case (Protect)
                Case True: tmpBuf = "Lockdown is currently active."
                Case Else: tmpBuf = "Lockdown is currently disabled."
            End Select
    End Select
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnProtect

' handle whispercmds command
Private Function OnWhisperCmds(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (StrComp(msgData, "status", vbTextCompare) = 0) Then
        tmpBuf = "Command responses will be " & _
            IIf(BotVars.WhisperCmds, "whispered back", "displayed publicly") & "."
    Else
        If (BotVars.WhisperCmds) Then
            BotVars.WhisperCmds = False
            
            Call WriteINI("Main", "WhisperBack", "N")
            
            tmpBuf = "Command responses will now be displayed publicly."
        Else
            BotVars.WhisperCmds = True

            Call WriteINI("Main", "WhisperBack", "Y")
            
            tmpBuf = "Command responses will now be whispered back."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhisperCmds

' handle stop command
Private Function OnStop(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        Call MediaPlayer.QuitPlayback
    
        tmpBuf = "Stopped play."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnStop

' handle play command
Private Function OnPlay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim Track  As Long
    
    If (BotVars.DisableMP3Commands = False) Then
        If (Len(msgData) > 0) Then
            If (StrictIsNumeric(msgData)) Then
                ' ...
                Track = CLng(msgData)
                
                ' ...
                Call MediaPlayer.PlayTrack(Track)
                
                ' ...
                tmpBuf = "Skipped to track " & Track & "."
            Else
                'Call WinampJumpToFile(msgData)
            End If
        Else
            ' ...
            Call MediaPlayer.PlayTrack
            
            ' ...
            tmpBuf = "Play started."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPlay

' handle useitunes command
Private Function OnUseiTunes(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    If (StrComp(BotVars.MediaPlayer, "iTunes", vbTextCompare) = 0) Then
        ' ...
    End If
    
    ' ...
    tmpBuf = "iTunes is ready."
    
    ' ...
    BotVars.MediaPlayer = "iTunes"
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUseiTunes

' handle usewinamp command
Private Function OnUseWinamp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' ...
    If (StrComp(BotVars.MediaPlayer, "Winamp", vbTextCompare) = 0) Then
        ' ...
    End If
    
    ' ...
    tmpBuf = "Winamp is ready."
    
    ' ...
    BotVars.MediaPlayer = "Winamp"

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUseWinamp

' handle pause command
Private Function OnPause(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (BotVars.DisableMP3Commands = False) Then
        Call MediaPlayer.PausePlayback
        
        tmpBuf = "Paused/resumed play."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPause

' handle fos command
Private Function OnFos(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim hWndWA As Long
    Dim tmpBuf As String ' temporary output buffer

   If (BotVars.DisableMP3Commands = False) Then
        'Call SendMessage(hWndWA, WM_COMMAND, WA_FADEOUTSTOP, 0)
        
        tmpBuf = "Fade-out stop."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFos

' handle rem command
Private Function OnRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u          As String  ' ...
    Dim tmpBuf     As String  ' temporary output buffer
    Dim dbType     As String  ' ...
    Dim index      As Long    ' ...
    Dim Params     As String  ' ...
    Dim strArray() As String  ' ...
    Dim i          As Integer ' ...

    ' check for presence of optional add command
    ' parameters
    index = InStr(1, msgData, " --", vbBinaryCompare)

    ' did we find such parameters?
    If (index > 0) Then
        ' grab parameters
        Params = Mid$(msgData, index - 1)

        ' remove paramaters from message
        msgData = Mid$(msgData, 1, index)
    End If
    
    ' do we have any special paramaters?
    If (Len(Params) > 0) Then
        ' split message by paramter
        strArray() = Split(Params, " --")
        
        ' loop through paramter list
        For i = 1 To UBound(strArray)
            Dim parameter As String ' ...
            Dim pmsg      As String ' ...
            
            ' check message for a space
            index = InStr(1, strArray(i), Space(1), vbBinaryCompare)
            
            ' did our search find a space?
            If (index > 0) Then
                ' grab parameter
                parameter = Mid$(strArray(i), 1, index - 1)
                
                ' grab parameter message
                pmsg = Mid$(strArray(i), index + 1)
            Else
                ' grab parameter
                parameter = strArray(i)
            End If
            
            ' convert parameter to lowercase
            parameter = LCase$(parameter)
            
            ' handle parameters
            Select Case (parameter)
                Case "type" ' ...
                    ' do we have a valid parameter length?
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
        Next i
    End If

    u = msgData
    
    If (Len(u) > 0) Then
        If ((GetAccess(u, dbType).Access = -1) And _
            (GetAccess(u, dbType).Flags = vbNullString)) Then
            
            tmpBuf = "User not found."
        ElseIf (GetAccess(u, dbType).Access >= dbAccess.Access) Then
            tmpBuf = "That user has higher or equal access."
        ElseIf ((InStr(1, GetAccess(u, dbType).Flags, "L") <> 0) And _
                (Not (InBot)) And _
                (InStr(1, GetAccess(Username, dbType).Flags, "A") = 0) And _
                (GetAccess(Username, dbType).Access <= 99)) Then
            
                tmpBuf = "Error: That user is Locked."
        Else
            Dim res As Boolean ' ...
        
            res = DB_remove(u, dbType)
            
            If (res) Then
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(RemEntry, Username, u, msgData)
                End If
                
                tmpBuf = "Successfully removed database entry " & Chr$(34) & _
                    u & "." & Chr$(34)
            Else
                tmpBuf = "Error: There was a problem removing that entry from the database."
            End If
            
            'Call LoadDatabase
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnRem

' handle reconnect command
Private Function OnReconnect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    If (g_Online) Then
        BotVars.HomeChannel = gChannel.Current
        
        Call frmChat.DoDisconnect
        
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Reconnecting by command, " & _
            "please wait..."
        
        Pause 1
        
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    Else
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    End If
End Function ' end function OnReconnect

' handle unigpriv command
Private Function OnUnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    Call AddQ("/o unigpriv", 1, Username)
    
    tmpBuf = "Recieving text from non-friends."
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnIgPriv

' handle igpriv command
Private Function OnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    Call AddQ("/o igpriv", 1, Username)
    
    tmpBuf = "Ignoring text from non-friends."
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIgPriv

' handle block command
Private Function OnBlock(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    Dim z      As String
    Dim i      As Integer

    u = msgData
    
    z = ReadINI("BlockList", "Total", "filters.ini")
    
    If (StrictIsNumeric(z)) Then
        i = z
    Else
        Call WriteINI("BlockList", "Total", "Total=0", "filters.ini")
        
        i = 0
    End If
    
    Call WriteINI("BlockList", "Filter" & (i + 1), u, "filters.ini")
    Call WriteINI("BlockList", "Total", i + 1, "filters.ini")
    
    tmpBuf = "Added """ & u & """ to the username block list."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBlock

' handle idletime command
Private Function OnIdleTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = msgData
        
    If ((Not (StrictIsNumeric(u))) Or (Val(u) > 50000)) Then
        tmpBuf = "Error setting idle wait time."
    Else
        Call WriteINI("Main", "IdleWait", 2 * Int(u))
        
        tmpBuf = "Idle wait time set to " & Int(u) & " minutes."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdleTime

' handle idle command
Private Function OnIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
        
    u = LCase$(msgData)
    
    If (u = "on") Then
        Call WriteINI("Main", "Idles", "Y")
        
        tmpBuf = "Idles activated."
    ElseIf (u = "off") Then
        Call WriteINI("Main", "Idles", "N")
        
        tmpBuf = "Idles deactivated."
    ElseIf (u = "kick") Then
        If (InStr(1, msgData, Space(1), vbBinaryCompare) = 0) Then
            tmpBuf = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
        Else
            u = LCase$(Mid$(msgData, InStr(1, msgData, Space(1)) + 1))
            
            If (u = "on") Then
                BotVars.IB_Kick = True
                
                tmpBuf = "Idle kick is now enabled."
            ElseIf (u = "off") Then
                BotVars.IB_Kick = False
                
                tmpBuf = "Idle kick disabled."
            Else
                tmpBuf = "Error: Unknown idle kick command."
            End If
        End If
    Else
        tmpBuf = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdle

' handle shitdel command
Private Function OnShitDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpBuf(0) = "Error: The specified username is invalid."
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, msgData & " -B --type USER", _
            True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnShitDel

' handle safedel command
Private Function OnSafeDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer

    ReDim Preserve tmpBuf(0)
    
    u = msgData
    
    If (InStr(1, u, Space(1), vbBinaryCompare) <> 0) Then
        tmpBuf(0) = "Error: The specified username is invalid."
    Else
        Call OnAdd(Username, dbAccess, u & " -S --type USER", True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnSafeDel

' handle tagdel command
Private Function OnTagDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    If (InStr(1, msgData, Space(1), vbBinaryCompare) <> 0) Then
        tmpBuf(0) = "Error: The specified tag is invalid."
    ElseIf (InStr(1, msgData, "*", vbBinaryCompare) <> 0) Then
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, msgData & " -B --type USER", True, tmpBuf())
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, "*" & msgData & "*" & " -B --type USER", _
            True, tmpBuf())
    End If
        
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnTagDel

' TO DO:
' handle profile command
Private Function OnProfile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim PPL    As Boolean
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(u) > 0) Then
        PPL = True
    
        'If ((BotVars.WhisperCmds Or WhisperedIn) And _
        '    (Not (PublicOutput))) Then
        '
        '    PPLRespondTo = Username
        'End If
        
        Call RequestProfile(u)
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnProfile

' handle setidle command
Private Function OnSetIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(u) > 0) Then
        If (Left$(u, 1) = "/") Then
            u = " " & u
        End If
        
        Call WriteINI("Main", "IdleMsg", u)
        
        tmpBuf = "Idle message set."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetIdle

' handle idletype command
Private Function OnIdleType(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
        
    u = msgData
    
    If ((LCase$(u) = "msg") Or (LCase$(u) = "message")) Then
        Call WriteINI("Main", "IdleType", "msg")
        
        tmpBuf = "Idle type set to [ msg ]."
    ElseIf ((LCase$(u) = "quote") Or (LCase$(u) = "quotes")) Then
        Call WriteINI("Main", "IdleType", "quote")
        
        tmpBuf = "Idle type set to [ quote ]."
    ElseIf (LCase$(u) = "uptime") Then
        Call WriteINI("Main", "IdleType", "uptime")
        
        tmpBuf = "Idle type set to [ uptime ]."
    ElseIf (LCase$(u) = "mp3") Then
        Call WriteINI("Main", "IdleType", "mp3")
        
        tmpBuf = "Idle type set to [ mp3 ]."
    Else
        tmpBuf = "Error setting idle type. The types are [ message quote uptime mp3 ]."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdleType

' handle filter command
Private Function OnFilter(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
    Dim z      As String

    u = msgData
    
    z = ReadINI("TextFilters", "Total", "filters.ini")
    
    If (StrictIsNumeric(z)) Then
        i = z
    Else
        Call WriteINI("TextFilters", "Total", "Total=0", "filters.ini")
        
        i = 0
    End If
    
    Call WriteINI("TextFilters", "Filter" & (i + 1), u, "filters.ini")
    Call WriteINI("TextFilters", "Total", i + 1, "filters.ini")
    
    ReDim Preserve gFilters(UBound(gFilters) + 1)
    
    gFilters(UBound(gFilters)) = u
    
    tmpBuf = "Added " & Chr(34) & u & Chr(34) & " to the text message filter list."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFilter

' handle trigger command
Private Function OnTrigger(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (Len(BotVars.TriggerLong) = 1) Then
        tmpBuf = "The bot's current trigger is " & Chr(34) & Space(1) & _
            BotVars.TriggerLong & Space(1) & Chr(34) & " (Alt + 0" & Asc(BotVars.TriggerLong) & ")"
    Else
        tmpBuf = "The bot's current trigger is " & Chr(34) & Space(1) & _
            BotVars.TriggerLong & Space(1) & Chr(34) & " (Length: " & Len(BotVars.TriggerLong) & ")"
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTrigger

' handle settrigger command
Private Function OnSetTrigger(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim newTrigger As String
    
    newTrigger = msgData
    
    If (Len(newTrigger) > 0) Then
        If (Left$(newTrigger, 1) <> "/") Then
            If ((Left$(newTrigger, 1) <> Space(1)) And _
                (Right$(newTrigger, 1) <> Space(1))) Then
                
                ' set new trigger
                BotVars.Trigger = newTrigger
            
                ' write trigger to configuration
                Call WriteINI("Main", "Trigger", newTrigger)
            
                tmpBuf = "The new trigger is " & Chr(34) & newTrigger & Chr(34) & "."
            Else
                tmpBuf = "Error: Trigger may not begin or end with a space."
            End If
        Else
            tmpBuf = "Error: Trigger may not begin with a forward slash."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetTrigger

' handle levelban command
Private Function OnLevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (StrictIsNumeric(msgData)) Then
            i = Val(msgData)
            
            If (i >= 1) Then
                If (i <= 255) Then
                    tmpBuf = "Banning Warcraft III users under level " & i & "."
                    
                    BotVars.BanUnderLevel = CByte(i)
                Else
                    tmpBuf = "Error: Invalid level specified."
                End If
            Else
                tmpBuf = "Levelbans disabled."
                
                BotVars.BanUnderLevel = 0
            End If
        Else
            BotVars.BanUnderLevel = 0
            
            tmpBuf = "Levelbans disabled."
        End If
        
        Call WriteINI("Other", "BanUnderLevel", BotVars.BanUnderLevel)
    Else
        If (BotVars.BanUnderLevel = 0) Then
           tmpBuf = "Currently not banning Warcraft III users by level."
        Else
           tmpBuf = "Currently banning Warcraft III users under level " & BotVars.BanUnderLevel & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLevelBan

' handle d2levelban command
Private Function OnD2LevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (StrictIsNumeric(msgData)) Then
            i = Val(msgData)
                
            If (i >= 1) Then
                If (i <= 255) Then
                    BotVars.BanD2UnderLevel = CByte(i)
            
                    tmpBuf = "Banning Diablo II characters under level " & i & "."
                Else
                    tmpBuf = "Error: Invalid level specified."
                End If
            Else
                tmpBuf = "Diablo II Levelbans disabled."
                
                BotVars.BanD2UnderLevel = 0
            End If
        Else
            tmpBuf = "Diablo II Levelbans disabled."
            
            BotVars.BanD2UnderLevel = 0
        End If
        
        Call WriteINI("Other", "BanD2UnderLevel", BotVars.BanD2UnderLevel)
    Else
        If (BotVars.BanD2UnderLevel = 0) Then
           tmpBuf = "Currently not banning Diablo II users by level."
        Else
           tmpBuf = "Currently banning Diablo II users under level " & BotVars.BanD2UnderLevel & "."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnD2LevelBans

' handle phrasebans command
Private Function OnPhraseBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    Dim tmpBuf As String ' temporary output buffer
  
    If (Len(msgData) > 0) Then
        ' ...
        msgData = LCase$(msgData)
    
        If (msgData = "on") Then
            Call WriteINI("Other", "Phrasebans", "Y")
            
            Phrasebans = True
            
            tmpBuf = "Phrasebans activated."
        Else
            Call WriteINI("Other", "Phrasebans", "N")
            
            Phrasebans = False
            
            tmpBuf = "Phrasebans deactivated."
        End If
    Else
        If (Phrasebans = True) Then
            tmpBuf = "Phrasebans are enabled."
        Else
            tmpBuf = "Phrasebans are disabled."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPhraseBans

' handle mimic command
'Private Function OnMimic(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'
'    Dim u      As String
'    Dim tmpBuf As String ' temporary output buffer
'
'    u = msgData
'
'    If (Len(u) > 0) Then
'        Mimic = LCase$(u)
'
'        tmpBuf = "Mimicking [ " & u & " ]"
'    End If
'
'    ' return message
'    cmdRet(0) = tmpBuf
'End Function ' end function OnMimic

' handle nomimic command
'Private Function OnNoMimic(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
'    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
'
'    Dim tmpBuf As String ' temporary output buffer
'
'    Mimic = vbNullString
'
'    tmpBuf = "Mimic off."
'
'    ' return message
'    cmdRet(0) = tmpBuf
'End Function ' end function OnNoMimic

' handle setpmsg command
Private Function OnSetPMsg(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = msgData
    
    ProtectMsg = u
    
    Call WriteINI("Other", "ProtectMsg", u)
    
    tmpBuf = "Channel protection message set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetPMsg

' handle phrases command
Private Function OnPhrases(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim response As String
    Dim i        As Integer
    Dim found    As Integer
    
    ReDim Preserve tmpBuf(tmpCount)

    tmpBuf(tmpCount) = "Phraseban(s): "
    
    For i = LBound(Phrases) To UBound(Phrases)
        If ((Phrases(i) <> " ") And (Phrases(i) <> vbNullString)) Then
            tmpBuf(tmpCount) = tmpBuf(tmpCount) & ", " & Phrases(i)
            
            If (Len(tmpBuf(tmpCount)) > 89) Then
                ReDim Preserve tmpBuf(tmpCount + 1)
                
                tmpBuf(tmpCount + 1) = "Phraseban(s): "
            
                tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), ", ", " ") & " [more]"
                
                tmpCount = (tmpCount + 1)
            End If
            
            found = (found + 1)
        End If
    Next i
    
    If (found > 0) Then
        tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), ", ", " ")
    Else
        tmpBuf(0) = "There are no phrasebans."
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnPhrases

' handle addphrase command
Private Function OnAddPhrase(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim c      As Integer
    Dim tmpBuf As String ' temporary output buffer
    Dim u      As String
    Dim i      As Integer
    
    ' grab free file handle
    f = FreeFile
    
    u = msgData
    
    For i = LBound(Phrases) To UBound(Phrases)
        If (StrComp(u, Phrases(i), vbTextCompare) = 0) Then
            Exit For
        End If
    Next i

    If (i > (UBound(Phrases))) Then
        If ((Phrases(UBound(Phrases)) <> vbNullString) Or _
            (Phrases(UBound(Phrases)) <> " ")) Then
            
            ReDim Preserve Phrases(0 To UBound(Phrases) + 1)
        End If
        
        Phrases(UBound(Phrases)) = u
        
        Open GetFilePath("phrasebans.txt") For Output As #f
            For c = LBound(Phrases) To UBound(Phrases)
                If (Len(Phrases(c)) > 0) Then
                    Print #f, Phrases(c)
                End If
            Next c
        Close #f
        
        tmpBuf = "Phraseban " & Chr(34) & u & Chr(34) & " added."
    Else
        tmpBuf = "Error: That phrase is already banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAddPhrase

' handle delphrase command
Private Function OnDelPhrase(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim u      As String
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer
    Dim c      As Integer
    
    u = msgData
    
    f = FreeFile
    
    Open GetFilePath("phrasebans.txt") For Output As #f
        Y = vbNullString
    
        For c = LBound(Phrases) To UBound(Phrases)
            If (StrComp(Phrases(c), LCase$(u), vbTextCompare) <> 0) Then
                Print #f, Phrases(c)
            Else
                Y = "x"
            End If
        Next c
    Close #f
    
    ReDim Phrases(0)
    
    Call frmChat.LoadArray(LOAD_PHRASES, Phrases())
    
    If (Len(Y) > 0) Then
        tmpBuf = "Phrase " & Chr(34) & u & Chr(34) & " deleted."
    Else
        tmpBuf = "Error: That phrase is not banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDelPhrase

' handle tagadd command
Private Function OnTagAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf() As String  ' ...
    Dim index    As Integer ' ...
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' ...
    index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (index) Then
        Dim user As String ' ...
        Dim bmsg As String ' ...
        
        ' ...
        user = Mid$(msgData, 1, index - 1)
        
        ' ...
        bmsg = Mid$(msgData, index + 1)
        
        ' ...
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            tmpBuf(0) = "Error: The specified username is invalid."
        ElseIf (InStr(1, user, "*", vbBinaryCompare) <> 0) Then
            ' ...
            Call OnAdd(Username, dbAccess, user & " +B --type USER --banmsg " & _
                bmsg, True, tmpBuf())
        Else
            If (Len(user) = 0) Then
                tmpBuf(0) = "Error: The specified tag is invalid."
            Else
                ' ...
                Call OnAdd(Username, dbAccess, "*" & user & "*" & " +B --type USER --banmsg " _
                    & bmsg, True, tmpBuf())
            End If
        End If
    Else
        If (InStr(1, msgData, "*", vbBinaryCompare) <> 0) Then
            ' ...
            Call OnAdd(Username, dbAccess, msgData & " +B --type USER", True, tmpBuf())
        Else
            If (Len(msgData) = 0) Then
               tmpBuf(0) = "Error: The specified tag is invalid."
            Else
                ' ...
                Call OnAdd(Username, dbAccess, "*" & msgData & "*" & " +B --type USER", True, _
                    tmpBuf())
            End If
        End If
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnTagAdd

' handle fadd command
Private Function OnFAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(u) > 0) Then
        Call AddQ("/f a " & u, 1, Username)
        
        tmpBuf = "Added user " & Chr(34) & u & Chr(34) & " to this account's friends list."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFAdd

' handle frem command
Private Function OnFRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(u) > 0) Then
        Call AddQ("/f r " & u, 1, Username)
        
        tmpBuf = "Removed user " & Chr(34) & u & Chr(34) & " from this account's friends list."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFRem

' handle safelist command
Private Function OnSafeList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase(tmpBuf(), , , , , , , "S")
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnSafeList

' handle safeadd command
Private Function OnSafeAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim u        As String
    
    ReDim Preserve tmpBuf(0)
    
    u = msgData
        
    If (InStr(1, u, Space(1), vbBinaryCompare) <> 0) Then
        tmpBuf(0) = "Error: The specified username is invalid."
    Else
        Call OnAdd(Username, dbAccess, u & " +S", True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnSafeAdd

' handle safecheck command
Private Function OnSafeCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer

    Y = msgData
            
    If (Len(Y)) Then
        If (GetSafelist(Y)) Then
            tmpBuf = Y & " is on the bot's safelist."
        Else
            tmpBuf = "That user is not safelisted."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSafeCheck

' handle exile command
Private Function OnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim saCmdRet() As String
    Dim ibCmdRet() As String
    Dim u          As String
    Dim Y          As String
    
    ReDim Preserve saCmdRet(0)
    ReDim Preserve ibCmdRet(0)

    u = msgData
    
    Call OnShitAdd(Username, dbAccess, u, InBot, saCmdRet())
    Call OnIPBan(Username, dbAccess, u, InBot, ibCmdRet())
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnExile

' handle unexile command
Private Function OnUnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim u          As String
    Dim sdCmdRet() As String
    Dim uiCmdRet() As String
    
    ' declare index zero of array
    ReDim Preserve sdCmdRet(0)
    ReDim Preserve uiCmdRet(0)

    u = msgData
    
    Call OnShitDel(Username, dbAccess, u, InBot, sdCmdRet())
    Call OnUnignore(Username, dbAccess, u, InBot, uiCmdRet())
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnExile

' handle shitlist command
Private Function OnShitList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase(tmpBuf(), , "!*[*]*", , , , , "B")
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnShitList

' handle tagbans command
Private Function OnTagBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase(tmpBuf(), , "*[*]*", , , , , "B")
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnTagBans

' handle shitadd command
Private Function OnShitAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String  ' ...
    Dim index    As Integer ' ...
    
    ' redefine array size
    ReDim Preserve tmpBuf(0)
    
    ' ...
    index = InStr(1, msgData, Space(1), vbBinaryCompare)
    
    ' ...
    If (index) Then
        Dim user As String ' ...
        
        ' ...
        user = Mid$(msgData, 1, index - 1)
        
        If (InStr(1, user, Space(1), vbBinaryCompare) <> 0) Then
            tmpBuf(0) = "Error: The specified username is invalid."
        Else
            Dim Msg As String ' ...
            
            ' ...
            Msg = Mid$(msgData, index + 1)
        
            ' ...
            Call OnAdd(Username, dbAccess, user & " +B --type USER --banmsg " & _
                Msg, True, tmpBuf())
        End If
    Else
        ' ...
        Call OnAdd(Username, dbAccess, msgData & " +B --type USER", True, tmpBuf())
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnShitAdd

' handle dnd command
Private Function OnDND(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim DNDMsg As String
    
    If (Len(msgData) = 0) Then
        AddQ "/dnd", 1
    Else
        DNDMsg = msgData
    
        AddQ "/dnd " & DNDMsg, 1, Username
    End If
End Function ' end function OnDND

' handle bancount command
Private Function OnBanCount(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (BanCount = 0) Then
        tmpBuf = "No users have been banned since I joined this channel."
    Else
        tmpBuf = "Since I joined this channel, " & BanCount & " user(s) " & _
            "have been banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBanCount

' handle banlistcount command
Private Function OnBanListCount(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (BanCount = 0) Then
        tmpBuf = "There are currently no users on the internal ban list."
    Else
        Dim bCount As Integer ' ...
        Dim i      As Integer ' ...
    
        tmpBuf = "There are currently " & UBound(gBans) & _
            " user(s) on the internal ban list"
        
        For i = 0 To UBound(gBans)
            If (StrComp(gBans(i).cOperator, CurrentUsername, vbTextCompare) = 0) Then
                bCount = (bCount + 1)
            End If
        Next i
        
        If ((MyFlags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
            tmpBuf = tmpBuf & ", " & _
                bCount & " of which users were banned by me."
        Else
            tmpBuf = tmpBuf & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBanCount

' handle tagcheck command
Private Function OnTagCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer

    Y = msgData
            
    If (Len(Y) > 0) Then
        gAcc = GetCumulativeAccess(Y)
        
        If (InStr(1, gAcc.Flags, "B") <> 0) Then
            tmpBuf = Y & " has been matched to one or more tagbans"
        
            If (InStr(1, gAcc.Flags, "S") <> 0) Then
                tmpBuf = tmpBuf & "; however, " & Y & " has also been found on the bot's " & _
                    "safelist and therefore will not be banned"
            End If
        Else
            tmpBuf = "That user matches no tagbans"
        End If
        
        tmpBuf = tmpBuf & "."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTagCheck

' handle slcheck command
Private Function OnSLCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer

    Y = msgData
            
    If (Len(Y) > 0) Then
        gAcc = GetCumulativeAccess(Y)
        
        If (InStr(1, gAcc.Flags, "B") <> 0) Then
            tmpBuf = Y & " is on the bot's shitlist"
        
            If (InStr(1, gAcc.Flags, "S") <> 0) Then
                tmpBuf = tmpBuf & "; however, " & Y & " is also on the " & _
                    "bot's safelist and therefore will not be banned"
            End If
        Else
            tmpBuf = "That user is not shitlisted"
        End If
        
        tmpBuf = tmpBuf & "."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSLCheck

' TO DO:
' handle readfile command
Private Function OnReadFile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    
    ' redefine array size
    ReDim Preserve tmpBuf(tmpCount)
    
    u = msgData
    
    If (Len(u)) Then
        If (InStr(1, u, "..", vbBinaryCompare) <> 0) Then
            tmpBuf(tmpCount) = "Error: You may only specify a file within the program " & _
                "directory or subdirectories."
        ElseIf (InStr(1, u, ".ini", vbTextCompare) <> 0) Then
            tmpBuf(tmpCount) = "Error: You may not read configuration files."
        Else
            Dim Y As String  ' ...
            Dim f As Integer ' ...
        
            ' grab a file number
            f = FreeFile
        
            If (InStr(1, u, ".", vbBinaryCompare) > 0) Then
                Y = Left$(u, InStr(1, u, ".", vbBinaryCompare) - 1)
            Else
                Y = u
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
            u = Dir$(App.Path & "\" & u)
            
            If (u = vbNullString) Then
                tmpBuf(tmpCount) = "Error: The specified file could not " & _
                    "be found."
            Else
                ' store line in buffer
                tmpBuf(tmpCount) = "Contents of file " & _
                    msgData & ":"
                
                ' increment counter
                tmpCount = (tmpCount + 1)
            
                ' open file
                Open u For Input As #f
                    ' read until end-of-line
                    Do While (Not (EOF(f)))
                        Dim tmp As String ' ...
                        
                        ' read line into tmp
                        Line Input #f, tmp
                        
                        If (tmp <> vbNullString) Then
                            ' redefine array size
                            ReDim Preserve tmpBuf(tmpCount)
                        
                            ' store line in buffer
                            tmpBuf(tmpCount) = "Line " & tmpCount & ": " & _
                                tmp
                            
                            ' increment counter
                            tmpCount = (tmpCount + 1)
                        End If
                    Loop
                Close #f
                
                ' redefine array size
                ReDim Preserve tmpBuf(tmpCount)
                
                ' store line in buffer
                tmpBuf(tmpCount) = "End of File."
            End If
        End If
    End If
    
    ' return message
    cmdRet() = tmpBuf()
    
    Exit Function
    
ERROR_HANDLER:
    cmdRet(0) = "There was an error reading the specified file."

    Exit Function
End Function ' end function OnReadFile

' TO DO:
' handle greet command
Private Function OnGreet(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf       As String ' temporary output buffer
    Dim strSplit()   As String
    Dim greetCommand As String
    
    ' split string by spaces
    strSplit() = Split(msgData, Space(1), 2)
    
    ' grab greet command
    greetCommand = strSplit(0)
    
    ' ...
    If (greetCommand = "on") Then
        BotVars.UseGreet = True
        
        tmpBuf = "Greet messages enabled."
        
        Call WriteINI("Other", "UseGreets", "Y")
    ElseIf (greetCommand = "off") Then
        BotVars.UseGreet = False
        
        tmpBuf = "Greet messages disabled."
        
        Call WriteINI("Other", "UseGreets", "N")
    ElseIf (greetCommand = "whisper") Then
        Dim greetSubCommand As String ' ...
        
        ' grab greet sub command
        greetSubCommand = strSplit(1)
        
        ' ...
        If (greetSubCommand = "on") Then
            BotVars.WhisperGreet = True
            
            tmpBuf = "Greet messages will now be whispered."
            
            Call WriteINI("Other", "WhisperGreet", "Y")
        ElseIf (greetSubCommand = "off") Then
            BotVars.WhisperGreet = False
            
            tmpBuf = "Greet messages will no longer be whispered."
            
            Call WriteINI("Other", "WhisperGreet", "N")
        End If
    Else
        Dim greetMessage As String ' ...
        
        ' grab greet message
        greetMessage = msgData
    
        ' ...
        If (Left$(greetMessage, 1) = "/") Then
            tmpBuf = "Error: Invalid greet message specified."
        Else
            tmpBuf = "Greet message set."
            
            BotVars.GreetMsg = greetMessage
            
            Call WriteINI("Other", "GreetMsg", greetMessage)
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGreet

' handle allseen command
Private Function OnAllSeen(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim i        As Integer

    ' redefine array size
    ReDim Preserve tmpBuf(tmpCount)

    ' prefix message with "Last 15 users seen"
    tmpBuf(tmpCount) = "Last 15 users seen: "
    
    ' were there any users seen?
    If (colLastSeen.Count = 0) Then
        tmpBuf(tmpCount) = tmpBuf(tmpCount) & "(list is empty)"
    Else
        For i = 1 To colLastSeen.Count
            ' append user to list
            tmpBuf(tmpCount) = tmpBuf(tmpCount) & _
                colLastSeen.Item(i) & ", "
            
            If (Len(tmpBuf(tmpCount)) > 90) Then
                If (i < colLastSeen.Count) Then
                    ' redefine array size
                    ReDim Preserve tmpBuf(tmpCount + 1)
                    
                    ' clear new array index
                    tmpBuf(tmpCount + 1) = vbNullString
                    
                    ' remove ending comma from index
                    tmpBuf(tmpCount) = Mid$(tmpBuf(tmpCount), 1, _
                        Len(tmpBuf(tmpCount)) - Len(", "))
                
                    ' postfix [more] to end of entry
                    tmpBuf(tmpCount) = tmpBuf(tmpCount) & " [more]"
                    
                    ' increment loop counter
                    tmpCount = (tmpCount + 1)
                End If
            End If
        Next i
        
        ' check for ending comma
        If (Right$(tmpBuf(tmpCount), 2) = ", ") Then
            ' remove ending comma from index
            tmpBuf(tmpCount) = Mid$(tmpBuf(tmpCount), 1, _
                Len(tmpBuf(tmpCount)) - Len(", "))
        End If
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnAllSeen

' handle ban command
Private Function OnBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim u       As String
    Dim tmpBuf  As String ' temporary output buffer
    Dim banmsg  As String
    Dim Y       As String
    Dim i       As Integer

    If ((MyFlags <> 2) And (MyFlags <> 18)) Then
        If (InBot) Then
            tmpBuf = "You are not a channel operator."
        End If
    Else
        u = msgData
        
        If (Len(u) > 0) Then
            i = InStr(1, u, " ")
            
            If (i > 0) Then
                banmsg = Mid$(u, i + 1)
                
                u = Left$(u, i - 1)
            End If
            
            If (InStr(1, u, "*", vbTextCompare) > 0) Then
                Call WildCardBan(u, banmsg, 1)
            Else
                If (banmsg <> vbNullString) Then
                    Y = Ban(u & IIf(Len(banmsg) > 0, " " & banmsg, _
                        vbNullString), dbAccess.Access)
                Else
                    Y = Ban(u & IIf(Len(banmsg) > 0, " " & banmsg, _
                        vbNullString), dbAccess.Access)
                End If
            End If
            
            If (Len(Y) > 2) Then
                tmpBuf = Y
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBan

' handle unban command
Private Function OnUnban(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    If ((MyFlags <> 2) And (MyFlags <> 18)) Then
       If (InBot) Then
           tmpBuf = "You are not a channel operator."
       End If
    Else
        u = msgData
        
        If (Len(u) > 0) Then
            If (Dii = True) Then
                If (Not (Mid$(u, 1, 1) = "*")) Then
                    u = "*" & u
                End If
            End If
            
            ' what the hell is a flood cap?
            If (bFlood) Then
                If (floodCap < 45) Then
                    floodCap = (floodCap + 15)
                    
                    Call bnetSend("/unban " & u)
                End If
            Else
                If (InStr(1, msgData, "*", vbTextCompare) <> 0) Then
                    Call WildCardBan(u, vbNullString, 2)
                Else
                    ' unipban user
                    'If (BotVars.IPBans = True) Then
                    '    Call AddQ("/unsquelch " & u, 1)
                    'End If
                
                    Call AddQ("/unban " & u, 1, Username)
                End If
            End If
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnBan

' handle kick command
Private Function OnKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim i      As Integer
    Dim banmsg As String
    Dim tmpBuf As String ' temporary output buffer
    Dim Y      As String
    
    If ((MyFlags <> 2) And (MyFlags <> 18)) Then
       If (InBot) Then
           tmpBuf = "You are not a channel operator."
       End If
    Else
        u = msgData
        
        If (Len(u) > 0) Then
            i = InStr(1, u, " ", vbTextCompare)
            
            If (i > 0) Then
                banmsg = Mid$(u, i + 1)
                
                u = Left$(u, i - 1)
            End If
            
            If (InStr(1, u, "*", vbTextCompare) > 0) Then
                If (dbAccess.Access > 99) Then
                    Call WildCardBan(u, banmsg, 0)
                Else
                    Call WildCardBan(u, banmsg, 0)
                End If
            Else
                Y = Ban(u & IIf(Len(banmsg) > 0, " " & banmsg, vbNullString), _
                    dbAccess.Access, 1)
                
                If (Len(Y) > 1) Then
                    tmpBuf = Y
                End If
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnKick

' handle lastwhisper command
Private Function OnLastWhisper(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (LastWhisper <> vbNullString) Then
        tmpBuf = "The last whisper to this bot was from " & LastWhisper & " at " & _
            FormatDateTime(LastWhisperFromTime, vbLongTime) & " on " & _
                FormatDateTime(LastWhisperFromTime, vbLongDate) & "."
        
    Else
        tmpBuf = "The bot has not been whispered since it logged on."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLastWhisper

' handle say command
Private Function OnSay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf  As String ' temporary output buffer
    Dim tmpSend As String ' ...
    
    If (Len(msgData)) Then
        If (dbAccess.Access >= GetAccessINIValue("say70", 70)) Then
            If (dbAccess.Access >= GetAccessINIValue("say90", 90)) Then
                tmpSend = msgData
            Else
                tmpSend = Replace(msgData, "/", "")
            End If
        Else
            tmpSend = Username & " says: " & _
                msgData
        End If
        
        Call AddQ(tmpSend, , Username)
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSay

' handle expand command
Private Function OnExpand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim tmpSend As String

    If (Len(msgData)) Then
        If (dbAccess.Access >= GetAccessINIValue("say70", 70)) Then
            If (dbAccess.Access >= GetAccessINIValue("say90", 90)) Then
                tmpSend = msgData
            Else
                tmpSend = Replace(msgData, "/", "")
            End If
            
            tmpSend = Expand(msgData)
        Else
            tmpSend = Username & " says: " & _
                Expand(msgData)
        End If
        
        If (Len(tmpSend) > 220) Then
            tmpSend = Mid$(tmpSend, 1, 220)
        End If
        
        tmpBuf = tmpSend
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnExpand

' handle detail command
Private Function OnDetail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = GetDBDetail(msgData)
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDetail

' handle info command
Private Function OnInfo(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim user      As String
    Dim userIndex As Integer
    Dim tmpBuf()  As String ' temporary output buffer

    user = msgData
    
    userIndex = UsernameToIndex(user)
    
    If (userIndex > 0) Then
        ReDim Preserve tmpBuf(0 To 1)
    
        With colUsersInChannel.Item(userIndex)
            tmpBuf(0) = "User " & .Username & " is logged on using " & _
                ProductCodeToFullName(.Product)
            
            If ((.Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
                tmpBuf(0) = tmpBuf(0) & " with ops, and a ping time of " & .Ping & "ms."
            Else
                tmpBuf(0) = tmpBuf(0) & " with a ping time of " & .Ping & "ms."
            End If
            
            tmpBuf(1) = "He/she has been present in the channel for " & _
                ConvertTime(.TimeInChannel(), 1) & "."
        End With
    Else
        tmpBuf(0) = "No such user is present."
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnInfo

' handle shout command
Private Function OnShout(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf  As String ' temporary output buffer
    Dim tmpSend As String

    If (Len(msgData)) Then
        If (dbAccess.Access > 69) Then
            If (dbAccess.Access > 89) Then
                tmpSend = msgData
            Else
                tmpSend = Replace(msgData, "/", vbNullString, 1)
            End If
            
            tmpSend = UCase$(tmpSend)
        Else
            tmpSend = Username & " shouts: " & _
                UCase$(msgData)
        End If
        
        tmpBuf = tmpSend
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShout

' handle voteban command
Private Function OnVoteBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim user   As String
    
    user = msgData
    
    If (VoteDuration = -1) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, user)
        
        VoteDuration = 30
        
        tmpBuf = "30-second VoteBan vote started. Type YES to ban " & user & ", NO to acquit him/her."
    Else
        tmpBuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVoteBan

' handle votekick command
Private Function OnVoteKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf   As String ' temporary output buffer
    Dim user     As String
    
    user = msgData
    
    If (VoteDuration = -1) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, user)
        
        VoteDuration = 30
        
        VoteInitiator = dbAccess
        
        tmpBuf = "30-second VoteKick vote started. Type YES to kick " & user & _
            ", NO to acquit him/her."
    Else
        tmpBuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVoteKick

' handle vote command
Private Function OnVote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
 
    Dim tmpBuf      As String ' temporary output buffer
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
                
                tmpBuf = "Vote initiated. Type YES or NO to vote; your vote will " & _
                    "be counted only once."
            Else
                ' duration entered is either negative, is too large, or is a string
                tmpBuf = "Please enter a number of seconds for your vote to last."
            End If
        Else
            tmpBuf = "A vote is currently in progress."
        End If
    Else
        ' duration not entered
        tmpBuf = "Please enter a number of seconds for your vote to last."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVote

' handle tally command
Private Function OnTally(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
     
    Dim tmpBuf As String ' temporary output buffer
     
    If (VoteDuration > 0) Then
        tmpBuf = Voting(BVT_VOTE_TALLY)
    Else
        tmpBuf = "No vote is currently in progress."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTally

' handle cancel command
Private Function OnCancel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (VoteDuration > 0) Then
        tmpBuf = Voting(BVT_VOTE_END, BVT_VOTE_CANCEL)
    Else
        tmpBuf = "No vote in progress."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCancel

' handle back command
Private Function OnBack(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long
    
    If (AwayMsg <> vbNullString) Then
        Call AddQ("/away", 1, Username)
        
        If (Not (InBot)) Then
            ' alert users of status change
            Call AddQ("/me is back from " & AwayMsg & ".")
            
            ' set away message
            AwayMsg = vbNullString
        End If
    Else
        ' ...
        Call OnPrev(Username, dbAccess, msgData, InBot, _
            cmdRet())
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBack

' handle prev command
Private Function OnPrev(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long   ' ...
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        Dim pos As Integer ' ...
        
        ' ...
        pos = MediaPlayer.PlaylistPosition
    
        ' ...
        Call MediaPlayer.PlayTrack(pos - 1)
        
        ' ...
        tmpBuf = "Skipped backwards."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPrev

' handle uptime command
Private Function OnUptime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = "System uptime " & ConvertTime(GetUptimeMS) & _
        ", connection uptime " & ConvertTime(uTicks) & "."
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUptime

' handle away command
Private Function OnAway(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(AwayMsg) > 0) Then
        ' send away command to battle.net
        Call AddQ("/away")
        
        ' alert users of status change
        If (Not (InBot)) Then
            Call AddQ("/me is back from (" & AwayMsg & ")")
        End If
        
        ' set away message
        AwayMsg = vbNullString
    Else
        If (Len(msgData) > 0) Then
            ' set away message
            AwayMsg = msgData
            
            ' send away command to battle.net
            Call AddQ("/away " & AwayMsg)
            
            ' alert users of status change
            If (Not (InBot)) Then
                Call AddQ("/me is away (" & AwayMsg & ")")
            End If
        Else
            ' set away message
            AwayMsg = " - "
        
            ' send away command to battle.net
            Call AddQ("/away")
            
            ' alert users of status change
            If (Not (InBot)) Then
                Call AddQ("/me is away.")
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAway

' handle mp3 command
Private Function OnMP3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf       As String  ' temporary output buffer
    Dim TrackName    As String  ' ...
    Dim ListPosition As Integer ' ...
    Dim ListCount    As Integer ' ...
    Dim TrackTime    As Integer ' ...
    Dim TrackLength  As Integer ' ...
    
    TrackName = MediaPlayer.TrackName
    ListPosition = MediaPlayer.PlaylistPosition
    ListCount = MediaPlayer.PlaylistCount
    TrackTime = MediaPlayer.TrackTime
    TrackLength = MediaPlayer.TrackLength
    
    If (TrackName = vbNullString) Then
        tmpBuf = "Winamp is not loaded."
    Else
        tmpBuf = "Current MP3 " & _
            "[" & ListPosition & "/" & ListCount & "]: " & _
                TrackName & " (" & SecondsToString(TrackTime) & _
                    "/" & SecondsToString(TrackLength) & ")"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMP3

' handle ping command
Private Function OnPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim Latency As Long
    Dim user    As String
    
    user = msgData
    
    If (Len(user)) Then
        Latency = GetPing(user)
        
        If (Latency < -1) Then
            tmpBuf = "I can't see " & user & " in the channel."
        Else
            tmpBuf = user & "'s ping at login was " & Latency & "ms."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPing

' handle addquote command
Private Function OnAddQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim u      As String
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer
    
    f = FreeFile
    
    u = msgData
    
    If (Len(u)) Then
        Y = Dir$(GetFilePath("quotes.txt"))
        
        If (Len(Y) = 0) Then
            Open GetFilePath("quotes.txt") For Output As #f
                Print #f, u
            Close #f
        Else
            Open Y For Append As #f
                Print #f, u
            Close #f
        End If
            
        tmpBuf = "Quote added!"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAddQuote

' handle owner command
Private Function OnOwner(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (LenB(BotVars.BotOwner)) Then
        tmpBuf = "This bot's owner is " & _
            BotVars.BotOwner & "."
    Else
        tmpBuf = "There is no owner currently set."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnOwner

' handle ignore command
Private Function OnIgnore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
        
    u = msgData
    
    If (Len(u)) Then
        If ((GetAccess(u).Access >= dbAccess.Access) Or _
            (InStr(GetAccess(u).Flags, "A"))) Then
            
            tmpBuf = "That user has equal or higher access."
        Else
            Call AddQ("/ignore " & u, 1)
            
            tmpBuf = "Ignoring messages from " & Chr(34) & u & Chr(34) & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIgnore

' handle quote command
Private Function OnQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = "Quote: " & _
        GetRandomQuote
    
    If (Len(tmpBuf) = 0) Then
        tmpBuf = "Error reading quotes, or no quote file exists."
    ElseIf (Len(tmpBuf) > 220) Then
        ' try one more time
        tmpBuf = "Quote: " & _
            GetRandomQuote
        
        If (Len(tmpBuf) > 220) Then
            'too long? too bad. truncate
            tmpBuf = Left$(tmpBuf, 220)
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnQuote

' handle unignore command
Private Function OnUnignore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(msgData)) Then
        AddQ "/unignore " & u, 1
        
        tmpBuf = "Receiving messages from """ & u & """."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnignore

' handle cq command
Private Function OnCQ(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' ...
    If (msgData = vbNullString) Then
        ' ...
        Call g_Queue.Clear
    
        ' ...
        tmpBuf = "Queue cleared."
    Else
        ' ...
        Call g_Queue.RemoveLines(vbNullString, msgData)
        
        ' ...
        tmpBuf = "Queue entries for the user " & Chr$(34) & _
            msgData & Chr$(34) & " have been removed."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCQ

' handle scq command
Private Function OnSCQ(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' ...
    If (msgData = vbNullString) Then
        ' ...
        Call g_Queue.Clear
    Else
        ' ...
        Call g_Queue.RemoveLines(vbNullString, msgData)
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCQ

' handle time command
Private Function OnTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = "The current time on this computer is " & Time & " on " & _
        Format(Date, "MM-dd-yyyy") & "."
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTime

' handle getping command
Private Function OnGetPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim Latency As Long

    If (InBot) Then
        If (g_Online) Then
            ' grab current latency
            Latency = GetPing(CurrentUsername)
        
            tmpBuf = "Your ping at login was " & Latency & "ms."
        Else
            tmpBuf = "You are not connected."
        End If
    Else
        Latency = GetPing(Username)
    
        If (Latency > -2) Then
            tmpBuf = "Your ping at login was " & Latency & "ms."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGetPing

' handle checkmail command
Private Function OnCheckMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Track  As Long
    Dim tmpBuf As String ' temporary output buffer
    
    If (InBot) Then
        Track = GetMailCount(CurrentUsername)
    Else
        Track = GetMailCount(Username)
    End If
    
    If (Track > 0) Then
        tmpBuf = "You have " & Track & " new messages."
        
        If (InBot) Then
            tmpBuf = tmpBuf & " Type /getmail to retrieve them."
        Else
            tmpBuf = tmpBuf & " Type !inbox to retrieve them."
        End If
    Else
        tmpBuf = "You have no mail."
    End If
     
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCheckMail

' handle getmail command
Private Function OnGetMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Msg    As udtMail
    
    Dim tmpBuf As String ' temporary output buffer
            
    If (InBot) Then
        Username = CurrentUsername
    End If
    
    If (GetMailCount(Username) > 0) Then
        Call GetMailMessage(Username, Msg)
        
        If (Len(RTrim(Msg.To)) > 0) Then
            tmpBuf = "Message from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message)
        End If
    Else
        tmpBuf = "You do not currently have any messages " & _
            "in your inbox."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGetMail

' handle whoami command
Private Function OnWhoAmI(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpBuf As String ' temporary output buffer

    If (InBot) Then
        tmpBuf = "You are the bot console."
        
        If (g_Online) Then
            Call AddQ("/whoami")
        End If
    ElseIf (dbAccess.Access = 1000) Then
        tmpBuf = "You are the bot owner, " & Username & "."
    Else
        If (dbAccess.Access > 0) Then
            If (dbAccess.Flags <> vbNullString) Then
                tmpBuf = dbAccess.Username & " has access " & dbAccess.Access & _
                    " and flags " & dbAccess.Flags & "."
            Else
                tmpBuf = dbAccess.Username & " has access " & dbAccess.Access & "."
            End If
        Else
            If (dbAccess.Flags <> vbNullString) Then
                tmpBuf = dbAccess.Username & " has flags " & dbAccess.Flags & "."
            End If
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhoAmI

' TO DO:
' handle add command
Private Function OnAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    ' ...
    Dim gAcc       As udtGetAccessResponse

    Dim strArray() As String  ' ...
    Dim i          As Integer ' ...
    Dim tmpBuf     As String  ' temporary output buffer
    Dim dbPath     As String  ' ...
    Dim user       As String  ' ...
    Dim Rank       As Integer ' ...
    Dim Flags      As String  ' ...
    Dim found      As Boolean ' ...
    Dim Params     As String  ' ...
    Dim index      As Integer ' ...
    Dim sGrp       As String  ' ...
    Dim dbType     As String  ' ...
    Dim banmsg     As String  ' ...

    ' check for presence of optional add command
    ' parameters
    index = InStr(1, msgData, " --", vbBinaryCompare)
    
    ' did we find such parameters, and if so,
    ' do they begin after an entry name?
    If (index > 1) Then
        ' grab parameters
        Params = Mid$(msgData, index - 1)

        ' remove paramaters from message
        msgData = Mid$(msgData, 1, index)
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
        
        ' do we have any special paramaters?
        If (Len(Params)) Then
            ' split message by paramter
            strArray() = Split(Params, " --")
            
            ' loop through paramter list
            For i = 1 To UBound(strArray)
                Dim parameter As String ' ...
                Dim pmsg      As String ' ...
                
                ' check message for a space
                index = InStr(1, strArray(i), Space(1), vbBinaryCompare)
                
                ' did our search find a space?
                If (index > 0) Then
                    ' grab parameter
                    parameter = Mid$(strArray(i), 1, index - 1)
                    
                    ' grab parameter message
                    pmsg = Mid$(strArray(i), index + 1)
                Else
                    ' grab parameter
                    parameter = strArray(i)
                End If
                
                ' convert parameter to lowercase
                parameter = LCase$(parameter)
                
                ' handle parameters
                Select Case (parameter)
                    Case "type" ' ...
                        ' do we have a valid parameter length?
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
                                        "incorrect length."
                                        
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
                        ' do we have a valid parameter length?
                        If (Len(pmsg)) Then
                            banmsg = pmsg
                        End If
                        
                    Case "group" ' ...
                        ' do we have a valid parameter length?
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
                            
                                If (dbAccess.Access < tmp.Access) Then
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
            Next i
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
        If (Len(gAcc.Username)) Then
            user = gAcc.Username
        End If

        ' is rank valid?
        If ((Rank <= 0) And (Flags = vbNullString) And _
            (sGrp = vbNullString)) Then
            
            tmpBuf = "Error: You have specified an invalid rank."
            
        ' is rank higher than user's rank?
        ElseIf ((Rank) And (Rank >= dbAccess.Access)) Then
            tmpBuf = "Error: You do not have sufficient access to assign an entry with the " & _
                "specified rank."
            
        ' can we modify specified user?
        ElseIf ((gAcc.Access) And (gAcc.Access >= dbAccess.Access)) Then
            tmpBuf = "Error: You do not have sufficient access to modify the specified entry."
        Else
            ' did we specify flags?
            If (Len(Flags)) Then
                Dim currentCharacter As String ' ...
            
                For i = 1 To Len(Flags)
                    currentCharacter = Mid$(Flags, i, 1)
                
                    If ((currentCharacter <> "+") And (currentCharacter <> "-")) Then
                        Select Case (currentCharacter)
                            Case "A" ' administrator
                                If (dbAccess.Access <= 100) Then
                                    Exit For
                                End If
                                
                            Case "B" ' banned
                                If (dbAccess.Access < 70) Then
                                    Exit For
                                End If
                                
                            Case "D" ' designated
                                If (dbAccess.Access < 100) Then
                                    Exit For
                                End If
                            
                            Case "L" ' locked
                                If (dbAccess.Access < 70) Then
                                    Exit For
                                End If
                            
                            Case "S" ' safelisted
                                If (dbAccess.Access < 70) Then
                                    Exit For
                                End If
                        End Select
                    End If
                Next i
                
                If (i < (Len(Flags) + 1)) Then
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
                            For i = 1 To Len(Flags)
                                currentCharacter = Mid$(Flags, i, 1)
                            
                                ' is flag valid (alphabetic)?
                                If (((Asc(currentCharacter) >= Asc("A")) And (Asc(currentCharacter) <= Asc("Z"))) Or _
                                    ((Asc(currentCharacter) >= Asc("a")) And (Asc(currentCharacter) <= Asc("z")))) Then
                                    
                                    If (InStr(1, gAcc.Flags, currentCharacter, vbBinaryCompare) = 0) Then
                                        gAcc.Flags = gAcc.Flags & currentCharacter
                                    End If
                                End If
                            Next i
                            
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
                                        ' unban user if found in banlist
                                        For i = LBound(gBans) To UBound(gBans)
                                            If (StrComp(gBans(i).Username, user, _
                                                    vbTextCompare) = 0) Then
                                                
                                                Call AddQ("/unban " & user)
                                            End If
                                        Next i
                                    End If
                                End If
                                
                                ' remove specified flags
                                For i = 1 To Len(tmpFlags)
                                    gAcc.Flags = Replace(gAcc.Flags, Mid$(tmpFlags, i, 1), _
                                        vbNullString)
                                Next i
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
                        If ((gAcc.Access = 0) And (gAcc.Flags = vbNullString) And _
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
                        gAcc.Access = Rank
                    
                        ' set user flags & check for duplicate entries
                        For i = 1 To Len(Flags)
                            currentCharacter = Mid$(Flags, i, 1)
                        
                            ' is flag valid (alphabetic)?
                            If (((Asc(currentCharacter) >= Asc("A")) And (Asc(currentCharacter) <= Asc("Z"))) Or _
                                ((Asc(currentCharacter) >= Asc("a")) And (Asc(currentCharacter) <= Asc("z")))) Then
                                
                                If (InStr(1, gAcc.Flags, currentCharacter, vbBinaryCompare) = 0) Then
                                    gAcc.Flags = gAcc.Flags & currentCharacter
                                End If
                            End If
                        Next i
                        
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
                gAcc.Access = Rank
            End If

            ' grab path to database
            dbPath = GetFilePath("users.txt")

            ' does user already exist in database?
            For i = LBound(DB) To UBound(DB)
                If ((StrComp(DB(i).Username, user, vbTextCompare) = 0) And _
                    (StrComp(DB(i).Type, gAcc.Type, vbTextCompare) = 0)) Then
                    
                    ' modify database entry
                    With DB(i)
                        .Username = user
                        .Access = gAcc.Access
                        .Flags = gAcc.Flags
                        .ModifiedBy = Username
                        .ModifiedOn = Now
                        .Type = dbType
                        .Groups = sGrp
                        
                        If (banmsg <> vbNullString) Then
                            .BanMessage = banmsg
                        End If
                    End With
                
                    ' commit modifications
                    Call WriteDatabase(dbPath)
                    
                    ' log actions
                    If (BotVars.LogDBActions) Then
                        Call LogDBAction(ModEntry, Username, gAcc.Username, msgData)
                    End If
                    
                    ' we have found the
                    ' specified user
                    found = True
                    
                    Exit For
                End If
            Next i
            
            ' did we find a matching entry or not?
            If (found = False) Then
                ' redefine array size
                ReDim Preserve DB(UBound(DB) + 1)
                
                With DB(UBound(DB))
                    .Username = user
                    .Access = IIf((gAcc.Access >= 0), _
                        gAcc.Access, 0)
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
                
                ' commit modifications
                Call WriteDatabase(dbPath)
                
                ' log actions
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(AddEntry, Username, gAcc.Username, msgData)
                End If
            End If
            
            ' check for errors & create message
            If (gAcc.Access > 0) Then
                tmpBuf = Chr(34) & user & Chr(34) & " has been given access " & _
                    gAcc.Access
                
                ' was the user given the specified flags, too?
                If (Len(gAcc.Flags)) Then
                    ' lets make sure we don't use
                    ' improper grammar because of groups!
                    If (Len(sGrp)) Then
                        tmpBuf = tmpBuf & ", flags " & gAcc.Flags
                    Else
                        tmpBuf = tmpBuf & " and flags " & gAcc.Flags
                    End If
                End If
            Else
                ' was the user given the specified flags?
                If (Len(gAcc.Flags)) Then
                    tmpBuf = Chr(34) & user & Chr(34) & " has been given flags " & _
                        gAcc.Flags
                End If
            End If
            
            ' was the user assigned to a group?
            If (Len(sGrp)) Then
                If (Len(tmpBuf)) Then
                    tmpBuf = tmpBuf & ", and has been made a member of " & _
                        "the group(s): " & sGrp
                Else
                    tmpBuf = Chr(34) & user & Chr(34) & " has been made a member of " & _
                        "the group(s): " & sGrp
                End If
            End If
            
            ' terminate sentence
            ' with period
            tmpBuf = tmpBuf & "."
        End If
        
        ' ...
        Call checkUsers
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAdd

' handle mmail command
Private Function OnMMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Temp       As udtMail
    
    Dim strArray() As String
    Dim tmpBuf     As String ' temporary output buffer
    Dim c          As Integer
    Dim f          As Integer
    Dim Track      As Long
    
    strArray = Split(msgData, " ", 2)
            
    If (UBound(strArray) > 0) Then
        Dim gAcc As udtGetAccessResponse ' ...
    
        tmpBuf = "Mass mailing "

        With Temp
            .From = Username
            .Message = strArray(1)
            
            If (StrictIsNumeric(strArray(0))) Then
                'number games
                Track = Val(strArray(0))
                
                For c = 0 To UBound(DB)
                    gAcc = GetCumulativeAccess(DB(c).Username)
                    
                    If (StrComp(gAcc.Type, "USER", vbTextCompare) = 0) Then
                        If (gAcc.Access = Track) Then
                            .To = DB(c).Username
                            
                            Call AddMail(Temp)
                        End If
                    End If
                Next c
                
                tmpBuf = tmpBuf & "to users with access " & Track
            Else
                For c = 0 To UBound(DB)
                    gAcc = GetCumulativeAccess(DB(c).Username)
                
                    For f = 1 To Len(strArray(0))
                        If (StrComp(gAcc.Type, "USER", vbTextCompare) = 0) Then
                            If (InStr(1, gAcc.Flags, Mid$(strArray(0), f, 1), _
                                vbBinaryCompare) > 0) Then
                                
                                .To = DB(c).Username
                                
                                Call AddMail(Temp)
                                
                                Exit For
                            End If
                        End If
                    Next f
                Next c
                
                tmpBuf = tmpBuf & "to users with any of the flags " & strArray(0)
            End If
        End With
        
        tmpBuf = tmpBuf & " complete."
    Else
        tmpBuf = "Format: .mmail <flag(s)> <message> OR .mmail <access> <message>"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMMail

' handle bmail command
Private Function OnBMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Temp       As udtMail ' ...

    Dim strArray() As String ' ...
    Dim tmpBuf     As String ' temporary output buffer
    
    ' ...
    strArray = Split(msgData, " ", 2)
    
    If (UBound(strArray) > 0) Then
        ' ...
        With Temp
            .To = strArray(0)
            .From = Username
            .Message = strArray(1)
        End With
        
        If (Len(Temp.To) = 0) Then
            tmpBuf = "Error: Invalid user."
        Else
            ' ...
            Call AddMail(Temp)
            
            tmpBuf = "Added mail for " & strArray(0) & "."
        End If
    Else
        tmpBuf = "Error: Too few arguments."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBMail

' handle designated command
Private Function OnDesignated(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If ((MyFlags <> 2) And (MyFlags <> 18)) Then
        tmpBuf = "The bot does not currently have ops."
    ElseIf (gChannel.Designated = vbNullString) Then
        tmpBuf = "No users have been designated."
    Else
        tmpBuf = "I have designated """ & gChannel.Designated & """."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDesignated

' handle flip command
Private Function OnFlip(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer

    Randomize
    
    i = (Rnd * 2)
    
    If (i = 0) Then
        tmpBuf = "Tails."
    Else
        tmpBuf = "Heads."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFlip

' handle about command
Private Function OnAbout(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = ".: " & CVERSION & " by Stealth."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAbout

' handle server command
Private Function OnServer(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf       As String ' temporary output buffer
    Dim RemoteHost   As String ' ...
    Dim RemoteHostIP As String ' ...
    
    ' ...
    RemoteHost = frmChat.sckBNet.RemoteHost
    
    ' ...
    RemoteHostIP = frmChat.sckBNet.RemoteHostIP
    
    ' ...
    If (StrComp(RemoteHost, RemoteHostIP, vbBinaryCompare) = 0) Then
        tmpBuf = "I am currently connected to " & _
            frmChat.sckBNet.RemoteHostIP & "."
    Else
        tmpBuf = "I am currently connected to " & _
            frmChat.sckBNet.RemoteHost & " (" & _
                frmChat.sckBNet.RemoteHostIP & ")."
    End If
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnServer

' handle find command
Private Function OnFind(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc     As udtGetAccessResponse

    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    
    ReDim Preserve tmpBuf(0)

    u = GetFilePath("users.txt")
            
    If (Dir$(u) = vbNullString) Then
        tmpBuf(0) = "No userlist available. Place a users.txt file" & _
            "in the bot's root directory."
    End If
    
    u = msgData
    
    If (Len(u) > 0) Then
        If (StrictIsNumeric(u)) Then
            ' execute search
            Call searchDatabase(tmpBuf(), , , , , Val(u))
        ElseIf (InStr(1, u, Space(1), vbBinaryCompare) <> 0) Then
            Dim lowerBound As String ' ...
            Dim upperBound As String ' ...
            
            ' grab range values
            If (InStr(1, u, " - ", vbBinaryCompare) <> 0) Then
                lowerBound = Mid$(u, 1, InStr(1, u, " - ", vbBinaryCompare) - 1)
                upperBound = Mid$(u, InStr(1, u, " - ", vbBinaryCompare) + Len(" - "))
            Else
                lowerBound = Mid$(u, 1, InStr(1, u, Space(1), vbBinaryCompare) - 1)
                upperBound = Mid$(u, InStr(1, u, Space(1), vbBinaryCompare) + 1)
            End If
            
            If ((StrictIsNumeric(lowerBound)) And _
                (StrictIsNumeric(upperBound))) Then
            
                ' execute search
                Call searchDatabase(tmpBuf(), , , , , CInt(Val(lowerBound)), CInt(Val(upperBound)))
            Else
                tmpBuf(0) = "Error: You have specified an invalid range."
            End If
        ElseIf ((InStr(1, u, "*", vbBinaryCompare) <> 0) Or _
                (InStr(1, u, "?", vbBinaryCompare) <> 0)) Then
            
            ' execute search
            Call searchDatabase(tmpBuf(), , PrepareCheck(u))
        Else
            ' execute search
            Call searchDatabase(tmpBuf(), , PrepareCheck(u))
        End If
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnFind

' handle whois command
Private Function OnWhoIs(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc     As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer
    Dim u        As String

    u = msgData
            
    If (InBot) Then
        Call AddQ("/whois " & u, 1)
    End If

    If (Len(u)) Then
        gAcc = GetCumulativeAccess(u)
        
        If (gAcc.Username <> vbNullString) Then
            If (gAcc.Access > 0) Then
                If (gAcc.Flags <> vbNullString) Then
                    tmpBuf = gAcc.Username & " has access " & gAcc.Access & _
                        " and flags " & gAcc.Flags & "."
                Else
                    tmpBuf = gAcc.Username & " has access " & gAcc.Access & "."
                End If
            Else
                If (gAcc.Flags <> vbNullString) Then
                    tmpBuf = gAcc.Username & " has flags " & gAcc.Flags & "."
                End If
            End If
        Else
            tmpBuf = "There was no such user found."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhoIs

' handle findattr command
Private Function OnFindAttr(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim i        As Integer
    Dim found    As Integer
    
    ReDim Preserve tmpBuf(tmpCount)
    
    ' ...
    u = msgData

    If (Len(u) > 0) Then
        ' execute search
        Call searchDatabase(tmpBuf(), , , , , , , u)
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnFindAttr

' handle findgrp command
Private Function OnFindGrp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u        As String
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim i        As Integer
    Dim found    As Integer
    
    ReDim Preserve tmpBuf(tmpCount)
    
    ' ...
    u = msgData

    If (Len(u) > 0) Then
        ' execute search
        Call searchDatabase(tmpBuf(), , , u)
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnFindAttr

' handle monitor command
Private Function OnMonitor(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (LCase$(msgData) = "on") Then
            If (Not (MonitorExists)) Then
                InitMonitor
                If (MonitorForm.Connect(False)) Then
                    tmpBuf = "User monitor connecting."
                Else
                    tmpBuf = "User monitor login information not filled in."
                End If
            Else
                tmpBuf = "User montor already enabled."
            End If
        ElseIf (LCase$(msgData) = "off") Then
            If (Not MonitorExists) Then
                tmpBuf = "User monitor is not running."
            Else
                MonitorForm.ShutdownMonitor
                tmpBuf = "User monitor disabled."
            End If
        Else
            If (Not (MonitorExists())) Then
                Call InitMonitor
            End If
            
            If (MonitorForm.AddUser(msgData)) Then
                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & _
                    " added to the monitor list."
            Else
                tmpBuf = "Failed to add user " & Chr$(&H22) & msgData & Chr$(&H22) & _
                    " to the monitor list. (Contains Spaces, or already in the list)"
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMonitor

' handle unmonitor command
Private Function OnUnMonitor(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (MonitorExists) Then
            If (MonitorForm.RemoveUser(msgData)) Then
                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & " was removed from the monitor list."
            Else
                tmpBuf = "User " & Chr$(&H22) & msgData & Chr$(&H22) & " was not found in the monitor list."
            End If
        Else
            tmpBuf = "User monitor is not enabled."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnMonitor

' handle online command
Private Function OnOnline(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (MonitorExists) Then
        tmpBuf = MonitorForm.OnlineUsers
    Else
        tmpBuf = "User monitor is not enabled."
    End If
    
    ' return message
    cmdRet = Split(tmpBuf, vbNewLine)
End Function ' end function OnOnline

' handle help command
Private Function OnHelp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    
    ' ...
    Call grabCommandData(msgData, tmpBuf())
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnHelp

' handle promote command
Private Function OnPromote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    If (Len(msgData) > 0) Then
        Dim liUser As ListItem ' ...
        
        ' ...
        Set liUser = _
            frmChat.lvClanList.FindItem(reverseUsername(msgData))
    
        ' ...
        If (Not (liUser Is Nothing)) Then
            With PBuffer
                .InsertDWord &HB
                .InsertNTString liUser.text
                .InsertWord (liUser.SmallIcon + 1)
                .SendPacket &H7A
            End With
        End If
    End If
End Function

' handle demote command
Private Function OnDemote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    If (Len(msgData) > 0) Then
        Dim liUser As ListItem ' ...
        
        ' ...
        Set liUser = _
            frmChat.lvClanList.FindItem(reverseUsername(msgData))
    
        ' ...
        If (Not (liUser Is Nothing)) Then
            With PBuffer
                .InsertDWord &H1
                .InsertNTString liUser.text
                .InsertByte (liUser.SmallIcon - 1)
                .SendPacket &H7A
            End With
        End If
    End If
End Function

' requires public
Public Function Cache(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
    Static s()  As String
    Static sTyp As String
    Dim i       As Integer
    
    'Debug.Print "cache input: " & Inpt
    
    If InStr(1, LCase$(Inpt), "in channel ", vbTextCompare) = 0 Then
        Select Case Mode
            Case 0
                For i = 0 To UBound(s)
                    Cache = Cache & Replace(s(i), ",", "") & Space(1)
                Next i
                
                'Debug.Print "cache output: " & Cache
                
                ReDim s(0)
                Typ = sTyp
                
            Case 1
                ReDim Preserve s(UBound(s) + 1)
                s(UBound(s)) = Inpt
                'Debug.Print "-> added " & Inpt & " to cache"
            Case 255
                ReDim s(0)
                sTyp = Typ
                
        End Select
    End If
End Function

Private Function Expand(ByVal s As String) As String
    Dim i As Integer
    Dim Temp As String
    
    If Len(s) > 1 Then
        For i = 1 To Len(s)
            Temp = Temp & Mid(s, i, 1) & Space(1)
        Next i
        Expand = Trim(Temp)
    Else
        Expand = s
    End If
End Function

Private Sub AddQ(ByVal s As String, Optional Priority As Byte = 100, Optional ByVal user As String = _
    vbNullString, Optional ByVal Tag As String = vbNullString)
    
    Call frmChat.AddQ(s, Priority, user, Tag)
End Sub

Private Function WildCardBan(ByVal sMatch As String, ByVal smsgData As String, ByVal Banning As Byte) ', Optional ExtraMode As Byte)
    'Values for Banning byte:
    '0 = Kick
    '1 = Ban
    '2 = Unban
    
    Dim i     As Integer
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
        
        If (colUsersInChannel.Count < 1) Then
            Exit Function
        End If
        
        If (Banning <> 2) Then
            ' Kicking or Banning
        
            For i = 1 To colUsersInChannel.Count
                With colUsersInChannel.Item(i)
                    If (Not (.IsSelf())) Then
                        z = PrepareCheck(.Username)
                        
                        If (z Like sMatch) Then
                            If (GetSafelist(.Username) = False) Then
                                If ((LenB(.Username) > 0) And _
                                   ((.Flags <> 2) And (.Flags <> 18))) Then
                                   
                                    Call AddQ("/" & Typ & .Username & Space(1) & _
                                        smsgData, 1)
                                End If
                            Else
                                iSafe = (iSafe + 1)
                            End If
                        End If
                    End If
                End With
            Next i
            
            If (iSafe) Then
                If (StrComp(smsgData, ProtectMsg, vbTextCompare) <> 0) Then
                    Call AddQ("Encountered " & iSafe & " safelisted user(s).")
                End If
            End If
            
        Else '// unbanning
        
            For i = 0 To UBound(gBans)
                If (gBans(i).Username <> vbNullString) Then
                    If (sMatch = "*") Then
                        ' unipban user
                        'If (BotVars.IPBans = True) Then
                        '    Call AddQ("/unsquelch " & gBans(i).UsernameActual, 1)
                        'End If
                    
                        Call AddQ("/" & Typ & gBans(i).UsernameActual, 1)
                    Else
                        ' unipban user
                        'If (BotVars.IPBans = True) Then
                        '    Call AddQ("/unsquelch " & gBans(i).UsernameActual, 1)
                        'End If
                    
                        z = PrepareCheck(gBans(i).UsernameActual)
                        
                        If (z Like sMatch) Then
                            Call AddQ("/" & Typ & gBans(i).UsernameActual, 1)
                        End If
                    End If
                End If
            Next i
        End If
    End If
End Function

Private Function searchDatabase(ByRef arrReturn() As String, Optional user As String = vbNullString, _
    Optional ByVal match As String = vbNullString, Optional Group As String = vbNullString, _
    Optional dbType As String = vbNullString, Optional lowerBound As Integer = -1, _
    Optional upperBound As Integer = -1, Optional Flags As String = vbNullString) As Integer
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    Dim i        As Integer
    Dim found    As Integer
    Dim tmpBuf   As String
    
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
        
        If (gAcc.Access > 0) Then
            If (gAcc.Flags <> vbNullString) Then
                tmpBuf = "Found user " & gAcc.Username & ", with access " & gAcc.Access & _
                    " and flags " & gAcc.Flags & "."
            Else
                tmpBuf = "Found user " & gAcc.Username & ", with access " & gAcc.Access & "."
            End If
        ElseIf (gAcc.Flags <> vbNullString) Then
            tmpBuf = "Found user " & gAcc.Username & ", with flags " & gAcc.Flags & "."
        Else
            tmpBuf = "No such user(s) found."
        End If
    Else
        For i = LBound(DB) To UBound(DB)
            Dim res        As Boolean ' store result of access check
            Dim blnChecked As Boolean ' ...
        
            If (DB(i).Username <> vbNullString) Then
                ' ...
                If (match <> vbNullString) Then
                    If (Left$(match, 1) = "!") Then
                        If (Not (LCase$(PrepareCheck(DB(i).Username)) Like _
                                (LCase$(Mid$(match, 2))))) Then

                            res = True
                        Else
                            res = False
                        End If
                    Else
                        If (LCase$(PrepareCheck(DB(i).Username)) Like _
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
                    If (StrComp(DB(i).Groups, Group, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If

                ' ...
                If (dbType <> vbNullString) Then
                    ' ...
                    If (StrComp(DB(i).Type, dbType, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                End If
                
                ' ...
                If ((lowerBound >= 0) And (upperBound >= 0)) Then
                    If ((DB(i).Access >= lowerBound) And _
                        (DB(i).Access <= upperBound)) Then
                        
                        ' ...
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                ElseIf (lowerBound >= 0) Then
                    If (DB(i).Access = lowerBound) Then
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
                        If (InStr(1, DB(i).Flags, Mid$(Flags, j, 1), _
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
                    tmpBuf = tmpBuf & DB(i).Username & _
                        IIf(((DB(i).Type <> "%") And _
                                (StrComp(DB(i).Type, "USER", vbTextCompare) <> 0)), _
                            " (" & LCase$(DB(i).Type) & ")", vbNullString) & _
                        IIf(DB(i).Access > 0, "\" & DB(i).Access, vbNullString) & _
                        IIf(DB(i).Flags <> vbNullString, "\" & DB(i).Flags, vbNullString) & ", "
                    
                    ' increment found counter
                    found = (found + 1)
                End If
            End If
            
            ' reset booleans
            res = False
            blnChecked = False
        Next i

        If (found = 0) Then
            ' return message
            arrReturn(0) = _
                "No such user(s) found."
        Else
            Dim prefix As String ' ...
            Dim arr()  As String ' ...
            
            ' ...
            prefix = "User(s) found: "
        
            ' ...
            Call SplitByLen(tmpBuf, 80 - _
                Len(prefix), arr(), " [more]", ", ")
            
            ' ...
            For i = 0 To UBound(arr)
                ' ...
                arr(i) = prefix & arr(i)
            Next i
            
            ' return message
            arrReturn() = arr()
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    MsgBox Err.description
    
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

    Dim i     As Integer ' ...
    Dim found As Boolean ' ...
    
    For i = LBound(DB) To UBound(DB)
        If (StrComp(DB(i).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean ' ...
        
            If (Len(dbType)) Then
                If (StrComp(DB(i).Type, dbType, vbBinaryCompare) = 0) Then
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
    Next i
    
    If (found) Then
        Dim bak As udtDatabase ' ...
        
        Dim j   As Integer ' ...
        
        ' ...
        bak = DB(i)

        ' we aren't removing the last array
        ' element, are we?
        If (i < UBound(DB)) Then
            For j = (i + 1) To UBound(DB)
                DB(j - 1) = DB(j)
            Next j
        End If
        
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
                For i = LBound(DB) To UBound(DB)
                    If (Len(DB(i).Groups) And DB(i).Groups <> "%") Then
                        If (InStr(1, DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Dim Splt()     As String ' ...
                            Dim innerfound As Boolean ' ...
                            
                            Splt() = Split(DB(i).Groups, ",")
                            
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
                                
                                DB(i).Groups = Join(Splt(), vbNullString)
                            End If
                        Else
                            If (StrComp(bak.Username, DB(i).Groups, vbTextCompare) = 0) Then
                                res = DB_remove(DB(i).Username, DB(i).Type)
                                
                                Exit For
                            End If
                        End If
                    End If
                Next i
            Loop While (res)
        End If
        
        ' commit modifications
        Call WriteDatabase(GetFilePath("users.txt"))
        
        DB_remove = True
        
        Exit Function
    End If
    
    DB_remove = False
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error: DB_remove() has encountered an error while " & _
        "removing a database entry.")
        
    DB_remove = False
    
    Exit Function
End Function

' requires public
Public Function GetSafelist(ByVal Username As String) As Boolean
    Dim i As Long ' ...
    
    ' ...
    If (Not (bFlood)) Then
        ' ...
        Dim gAcc As udtGetAccessResponse
        
        ' ...
        gAcc = GetCumulativeAccess(Username, "USER")
        
        ' ...
        If (InStr(1, gAcc.Flags, "S", vbBinaryCompare) <> 0) Then
            GetSafelist = True
        ElseIf (gAcc.Access >= 20) Then
            GetSafelist = True
        End If
    Else
        ' ...
        For i = 0 To (UBound(gFloodSafelist) - 1)
            If PrepareCheck(Username) Like gFloodSafelist(i) Then
                GetSafelist = True
                
                ' ...
                Exit For
            End If
        Next i
    End If
End Function

' requires public
Public Function GetShitlist(ByVal Username As String) As String
    ' ...
    Dim gAcc As udtGetAccessResponse
    
    ' ...
    gAcc = GetCumulativeAccess(Username, "USER")
    
    ' ...
    If ((InStr(1, gAcc.Flags, "B", vbBinaryCompare) <> 0) And _
        (InStr(1, gAcc.Flags, "S", vbBinaryCompare) = 0) And _
        (gAcc.Access < 20)) Then
        
        If ((Len(gAcc.BanMessage) > 0) And (gAcc.BanMessage <> "%")) Then
            GetShitlist = Username & Space(1) & _
                gAcc.BanMessage
        Else
            GetShitlist = Username & Space(1) & _
                "Shitlisted"
        End If
    End If
End Function

' requires public
Public Function GetPing(ByVal Username As String) As Long
    Dim i As Integer
    
    i = UsernameToIndex(Username)
    
    If i > 0 Then
        GetPing = colUsersInChannel.Item(i).Ping
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

Private Sub DBRemove(ByVal s As String)
    Dim t()  As udtDatabase
    
    Dim i    As Integer
    Dim c    As Integer
    Dim n    As Integer
    Dim Temp As String
    
    s = LCase$(s)
    
    For i = LBound(DB) To UBound(DB)
        If StrComp(DB(i).Username, s, vbTextCompare) = 0 Then
            ReDim t(0 To UBound(DB) - 1)
            For c = LBound(DB) To UBound(DB)
                If c <> i Then
                    t(n) = DB(c)
                    n = n + 1
                End If
            Next c
            
            ReDim DB(UBound(t))
            For c = LBound(t) To UBound(t)
                DB(c) = t(c)
            Next c
            Exit Sub
        End If
    Next i
    
    n = FreeFile
    
    Temp = GetFilePath("users.txt")
    
    Open Temp For Output As #n
        For i = LBound(DB) To UBound(DB)
            Print #n, DB(i).Username & Space(1) & DB(i).Access & Space(1) & DB(i).Flags
        Next i
    Close #n
End Sub

' requires public
Public Sub LoadDatabase()
    On Error Resume Next

    Dim gA    As udtDatabase
    
    Dim s     As String
    Dim X()   As String
    Dim Path  As String
    Dim i     As Integer
    Dim f     As Integer
    Dim found As Boolean
    
    ReDim DB(0)
    
    Path = GetFilePath("users.txt")
    
    If Dir$(Path) <> vbNullString Then
        f = FreeFile
        Open Path For Input As #f
            
        If LOF(f) > 1 Then
            Do
                
                Line Input #f, s
                
                If InStr(1, s, " ", vbTextCompare) > 0 Then
                    X() = Split(s, " ", 10)
                    
                    If UBound(X) > 0 Then
                        ReDim Preserve DB(i)
                        
                        With DB(i)
                            .Username = X(0)
                            
                            If StrictIsNumeric(X(1)) Then
                                .Access = Val(X(1))
                            Else
                                If X(1) <> "%" Then
                                    .Flags = X(1)
                                    
                                    'If InStr(X(1), "S") > 0 Then
                                    '    AddToSafelist .Username
                                    '    .Flags = Replace(.Flags, "S", "")
                                    'End If
                                End If
                            End If
                            
                            If UBound(X) > 1 Then
                                If StrictIsNumeric(X(2)) Then
                                    .Access = Int(X(2))
                                Else
                                    If X(2) <> "%" Then
                                        .Flags = X(2)
                                    End If
                                End If
                                
                                '  0        1       2       3      4        5          6       7     8
                                ' username access flags addedby addedon modifiedby modifiedon type banmsg
                                If UBound(X) > 2 Then
                                    .AddedBy = X(3)
                                    
                                    If UBound(X) > 3 Then
                                        .AddedOn = CDate(Replace(X(4), "_", " "))
                                        
                                        If UBound(X) > 4 Then
                                            .ModifiedBy = X(5)
                                            
                                            If UBound(X) > 5 Then
                                                .ModifiedOn = CDate(Replace(X(6), "_", " "))

                                                If UBound(X) > 6 Then
                                                    .Type = X(7)
                                                    
                                                    If UBound(X) > 7 Then
                                                        .Groups = X(8)
                                                        
                                                        If UBound(X) > 8 Then
                                                            .BanMessage = X(9)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                
                            End If
                            
                            If .Access > 200 Then: .Access = 200
                        End With

                        i = i + 1
                    End If
                End If
                
            Loop While Not EOF(f)
        End If
        
        Close #f
    End If

    ' 9/13/06: Add the bot owner 200
    If (LenB(BotVars.BotOwner) > 0) Then
        For i = 0 To UBound(DB)
            If (StrComp(DB(i).Username, BotVars.BotOwner, vbTextCompare) = 0) Then
                found = True
                
                Exit For
            End If
        Next i
        
        If (found = False) Then
            If (UBound(DB)) Then
                ReDim Preserve DB(UBound(DB) + 1)
            End If
            
            With DB(UBound(DB))
                .Username = BotVars.BotOwner
                .Type = "USER"
                .Access = 200
                .AddedBy = "(console)"
                .AddedOn = Now
                .ModifiedBy = "(console)"
                .ModifiedOn = Now
            End With
            
            Call WriteDatabase(Path)
        End If
    End If
End Sub

Private Function ValidateAccess(ByRef gAcc As udtGetAccessResponse, ByVal CWord As String, _
    Optional ByVal ARGUMENT As String = vbNullString, Optional ByVal restrictionName As String = _
    vbNullString) As Boolean
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    ' ...
    If (Len(CWord) > 0) Then
        Dim Commands As MSXML2.DOMDocument
        Dim command  As MSXML2.IXMLDOMNode
        
        ' ...
        Set Commands = New MSXML2.DOMDocument
        
        ' ...
        If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
            Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
                "working directory.")
                
            Exit Function
        End If

        ' ...
        Call Commands.Load(App.Path & "\commands.xml")
        
        ' ...
        For Each command In Commands.documentElement.childNodes
            Dim accessGroup As MSXML2.IXMLDOMNode
            Dim Access      As MSXML2.IXMLDOMNode
        
            ' ...
            If (StrComp(command.Attributes.getNamedItem("name").text, _
                CWord, vbTextCompare) = 0) Then
                
                ' ...
                Set accessGroup = command.selectSingleNode("access")
                
                ' ...
                For Each Access In accessGroup.childNodes
                    If (LCase$(Access.nodeName) = "rank") Then
                        If ((gAcc.Access) >= (Val(Access.text))) Then
                            ValidateAccess = True
                        
                            Exit For
                        End If
                    ElseIf (LCase$(Access.nodeName = "flag")) Then
                        If (InStr(1, gAcc.Flags, Access.text, vbBinaryCompare) <> 0) Then
                            ValidateAccess = True
                        
                            Exit For
                        End If
                    End If
                Next
                
                ' ...
                If (ARGUMENT <> vbNullString) Then
                    ' ...
                End If
                
                ' ...
                If (restrictionName <> vbNullString) Then
                    Dim restrictions As MSXML2.IXMLDOMNodeList
                    Dim RESTRICTION  As MSXML2.IXMLDOMNode
                    
                    ' ...
                    Set restrictions = command.selectNodes("restriction")
                    
                    ' ...
                    For Each RESTRICTION In restrictions
                        If (StrComp(RESTRICTION.Attributes.getNamedItem("name").text, _
                            restrictionName, vbTextCompare) = 0) Then
                            
                            ' ...
                            Set accessGroup = RESTRICTION.selectSingleNode("access")
                            
                            ' ...
                            For Each Access In accessGroup.childNodes
                                If (LCase$(Access.nodeName) = "rank") Then
                                    If ((gAcc.Access) >= (Val(Access.text))) Then
                                        ValidateAccess = True
                                    
                                        Exit For
                                    End If
                                ElseIf (LCase$(Access.nodeName = "flag")) Then
                                    If (InStr(1, gAcc.Flags, Access.text, vbBinaryCompare) <> 0) Then
                                        ValidateAccess = True
                                    
                                        Exit For
                                    End If
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

' ...
Private Function convertAlias(ByVal cmdName As String) As String
    ' ...
    On Error GoTo ERROR_HANDLER

    ' ...
    If (Len(cmdName) > 0) Then
        Dim Commands As MSXML2.DOMDocument
        Dim command  As MSXML2.IXMLDOMNode
        
        ' ...
        Set Commands = New MSXML2.DOMDocument
        
        ' ...
        If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
            Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
                "working directory.")
                
            Exit Function
        End If
        
        ' ...
        Call Commands.Load(App.Path & "\commands.xml")
        
        ' ...
        For Each command In Commands.documentElement.childNodes
            Dim aliases As MSXML2.IXMLDOMNodeList
            Dim alias   As MSXML2.IXMLDOMNode
            
            ' ...
            Set aliases = command.selectNodes("alias")

            ' ...
            For Each alias In aliases
                ' ...
                If (StrComp(alias.text, cmdName, vbTextCompare) = 0) Then
                    ' ...
                    convertAlias = command.Attributes.getNamedItem("name").text
                    
                    Exit Function
                End If
            Next
        Next
    End If
    
    ' ...
    convertAlias = cmdName
    
    Exit Function
    
' ...
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: XML Database Processor has encountered an error " & _
        "during alias lookup.")
        
    ' ...
    convertAlias = False

    Exit Function
End Function

Public Sub grabCommandData(ByVal cmdName As String, cmdRet() As String)
    ' ...
    On Error GoTo ERROR_HANDLER
    
    Dim tmpBuf() As String  ' ...
    Dim tmpCount As Integer ' ...
    Dim found    As Integer ' ...

    ' redefine array size
    ReDim Preserve tmpBuf(tmpCount)
    
    ' ...
    cmdName = convertAlias(cmdName)
    
    ' ...
    If (Len(cmdName) > 0) Then
        Dim Commands As MSXML2.DOMDocument
        Dim command  As MSXML2.IXMLDOMNode
        
        ' ...
        Set Commands = New MSXML2.DOMDocument
        
        ' ...
        If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
            Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
                "working directory.")
                
            Exit Sub
        End If

        ' ...
        Call Commands.Load(App.Path & "\commands.xml")
        
        ' ...
        For Each command In Commands.documentElement.childNodes
            Dim blnFound As Boolean ' ...
        
            ' ...
            If (StrComp(command.Attributes.getNamedItem("name").text, _
                cmdName, vbTextCompare) = 0) Then
                
                ' ...
                blnFound = True
            ElseIf ((PrepareCheck(command.Attributes.getNamedItem("name").text)) Like _
                (PrepareCheck(cmdName))) Then
                
                ' ...
                blnFound = True
            End If
            
            ' ...
            If (blnFound = True) Then
                ' ...
                Dim docs    As MSXML2.IXMLDOMNode
                Dim Access  As MSXML2.IXMLDOMNode
                Dim args    As MSXML2.IXMLDOMNode
                Dim arg     As MSXML2.IXMLDOMNode
                
                Dim tmp     As String ' ...
                
                ' ...
                Set docs = command.selectSingleNode("documentation")
                Set Access = command.selectSingleNode("access")
                Set args = command.selectSingleNode("arguments")
        
                ' ...
                If (found >= 1) Then
                    tmpBuf(tmpCount) = tmpBuf(tmpCount) & _
                        " [more]"
                
                    ' ...
                    tmpCount = (tmpCount + 1)
                    
                    ' ...
                    ReDim Preserve tmpBuf(tmpCount)
                End If
                
                ' ...
                tmpBuf(tmpCount) = docs.selectSingleNode("description").text

                ' ...
                If (Len(tmpBuf(tmpCount)) > 0) Then
                    ' ...
                    If (Right$(tmpBuf(tmpCount), 1) = ".") Then
                        ' ...
                        tmpBuf(tmpCount) = Left$(tmpBuf(tmpCount), _
                            Len(tmpBuf(tmpCount)) - 1)
                    End If
                    
                    ' ...
                    tmpBuf(tmpCount) = tmpBuf(tmpCount) & _
                        ", and requires "
                Else
                    ' ...
                    tmpBuf(tmpCount) = "This command requires "
                End If
                
                ' ...
                If (Not (Access.selectSingleNode("rank") Is Nothing)) Then
                    tmp = Access.selectSingleNode("rank").text & _
                        " access"
                End If
                
                ' ...
                If (Not (Access.selectSingleNode("flag") Is Nothing)) Then
                    Dim Flags As MSXML2.IXMLDOMNodeList
                    Dim flag  As MSXML2.IXMLDOMNode
                
                    ' ...
                    Set Flags = Access.selectNodes("flag")
                    
                    ' ...
                    tmp = "either " & _
                        tmp & " or flag "
                
                    ' ...
                    For Each flag In Flags
                        Dim flagCount As Integer ' ...
                    
                        ' ...
                        If (Not (flag.nextSibling Is Nothing)) Then
                            tmp = tmp & _
                                "'" & flag.text & "', "
                        Else
                            ' ...
                            If (flagCount = 0) Then
                                tmp = tmp & _
                                    "'" & flag.text & "'"
                            Else
                                tmp = tmp & _
                                    "or '" & flag.text & "'"
                            End If
                        End If
                        
                        ' ...
                        flagCount = (flagCount + 1)
                    Next
                End If
                
                ' ...
                tmpBuf(tmpCount) = tmpBuf(tmpCount) & tmp & "."
                
                ' ...
                tmpBuf(tmpCount) = tmpBuf(tmpCount) & _
                    " [Syntax: <trigger>" & _
                    command.Attributes.getNamedItem("name").text & _
                        Space(1)
                
                ' ...
                If (Not (arg Is Nothing)) Then
                    For Each arg In args.childNodes
                        tmpBuf(tmpCount) = tmpBuf(tmpCount) & _
                            arg.Attributes.getNamedItem("name").text & Space(1)
                    Next
                End If
                
                ' ...
                tmpBuf(tmpCount) = tmpBuf(tmpCount) & "]"
            
                ' ...
                found = (found + 1)
            End If
            
            ' ...
            blnFound = False
        Next
    End If
    
    ' ...
    If (found = 0) Then
        tmpBuf(tmpCount) = "Sorry, but no related documentation could be found."
    End If
    
    ' redefine array size
    ReDim Preserve cmdRet(0 To tmpCount)
    
    ' return result
    cmdRet() = tmpBuf()
    
    Exit Sub
    
' ...
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: XML Database Processor has encountered an error " & _
        "during documentation lookup.")
    
    Exit Sub
End Sub

' ...
Public Sub checkUsers()
    If ((MyFlags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        Dim i       As Integer ' ...
        Dim tmp     As String  ' ...
        Dim doCheck As Boolean ' ...
    
        ' ...
        doCheck = True
    
        For i = 1 To colUsersInChannel.Count
            If ((colUsersInChannel(i).Flags And USER_CHANNELOP&) <> _
                 USER_CHANNELOP&) Then
                 
                If (GetSafelist(colUsersInChannel.Item(i).Username) = False) Then
                    If (Protect) Then
                        ' ...
                        Call Ban(colUsersInChannel.Item(i).Username & _
                            Space(1) & ProtectMsg, (AutoModSafelistValue - 1))
                    Else
                        ' ...
                        tmp = GetShitlist(colUsersInChannel.Item(i).Username)
                       
                        ' ...
                        If (tmp <> vbNullString) Then
                            ' ...
                            Call AddQ("/ban " & tmp)
                        Else
                            Dim j As Integer ' ...
                
                            ' ...
                            If ((doCheck) And (BotVars.BanEvasion)) Then
                                For j = 0 To UBound(gBans)
                                    If (StrComp(colUsersInChannel.Item(i).Username, _
                                            gBans(j).Username, vbTextCompare) = 0) Then
                                        
                                        Call Ban(colUsersInChannel.Item(i).Username & _
                                            " Ban Evasion", (AutoModSafelistValue - 1))
                                        
                                        ' ...
                                        doCheck = False
                                    End If
                                Next j
                            End If
                            
                            ' ...
                            If ((doCheck) And (BotVars.IPBans)) Then
                                If ((colUsersInChannel.Item(i).Flags And USER_SQUELCHED) = _
                                     USER_SQUELCHED) Then
                                    
                                    Call Ban(colUsersInChannel.Item(i).Username & _
                                        " IPBanned.", (AutoModSafelistValue - 1))
                                    
                                    ' ...
                                    doCheck = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            ' ...
            tmp = vbNullString
            
            ' ...
            doCheck = True
        Next i
    End If
End Sub

' requires public
Public Function GetRandomQuote() As String
    On Error GoTo GetRandomQuote_Error

    Dim colQuotes As Collection
    
    Dim Rand      As Integer
    Dim f         As Integer
    Dim s         As String
    
    Set colQuotes = New Collection

    If LenB(Dir$(GetFilePath("quotes.txt"))) > 0 Then
    
        f = FreeFile
        Open (GetFilePath("quotes.txt")) For Input As #f
        
        If LOF(f) > 1 Then
        
            Do
                Line Input #f, s
                
                s = Trim(s)
                
                If LenB(s) > 0 Then
                    colQuotes.Add s
                End If
            Loop Until EOF(f)
            
            Randomize
            Rand = Rnd * colQuotes.Count
            
            If Rand <= 0 Then
                Rand = 1
            End If
            
            If Len(colQuotes.Item(Rand)) < 1 Then
                Randomize
                Rand = Rnd * colQuotes.Count
                
                If Rand <= 0 Then
                    Rand = 1
                End If
            End If
            
            GetRandomQuote = colQuotes.Item(Rand)

        End If
        
        Close #f
        
    End If
    
    If Left$(GetRandomQuote, 1) = "/" Then GetRandomQuote = " " & GetRandomQuote

GetRandomQuote_Exit:
    Set colQuotes = Nothing
    
    Exit Function

GetRandomQuote_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure GetRandomQuote of Module modCommandCode"
    Resume GetRandomQuote_Exit
End Function

' Writes database to disk
' Updated 9/13/06 for new features
Public Sub WriteDatabase(ByVal u As String)
    Dim f As Integer
    Dim i As Integer
    
    On Error GoTo WriteDatabase_Exit

    f = FreeFile
    
    Open u For Output As #f
        For i = LBound(DB) To UBound(DB)
            ' ...
            If ((DB(i).Access > 0) Or _
                (Len(DB(i).Flags) > 0) Or _
                (Len(DB(i).Groups) > 0)) Then
                
                ' ...
                Print #f, DB(i).Username;
                Print #f, " " & DB(i).Access;
                Print #f, " " & IIf(Len(DB(i).Flags) > 0, DB(i).Flags, "%");
                Print #f, " " & IIf(Len(DB(i).AddedBy) > 0, DB(i).AddedBy, "%");
                Print #f, " " & IIf(DB(i).AddedOn > 0, DateCleanup(DB(i).AddedOn), "%");
                Print #f, " " & IIf(Len(DB(i).ModifiedBy) > 0, DB(i).ModifiedBy, "%");
                Print #f, " " & IIf(DB(i).ModifiedOn > 0, DateCleanup(DB(i).ModifiedOn), "%");
                Print #f, " " & IIf(Len(DB(i).Type) > 0, DB(i).Type, "%");
                Print #f, " " & IIf(Len(DB(i).Groups) > 0, DB(i).Groups, "%");
                Print #f, " " & IIf(Len(DB(i).BanMessage) > 0, DB(i).BanMessage, "%");
                Print #f, vbCr
            End If
        Next i

WriteDatabase_Exit:
    Close #f
    
    Exit Sub

WriteDatabase_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure " & _
        "WriteDatabase of Module modCommandCode"
    
    Resume WriteDatabase_Exit
End Sub

Private Function GetDBDetail(ByVal Username As String) As String
    Dim sRetAdd As String, sRetMod As String
    Dim i As Integer
    
    For i = 0 To UBound(DB)
        With DB(i)
            If (StrComp(Username, .Username, vbTextCompare) = 0) Then
                If .AddedBy <> "%" And LenB(.AddedBy) > 0 Then
                    sRetAdd = " was added by " & .AddedBy & " on " & _
                        .AddedOn & "."
                End If
                
                If ((.ModifiedBy <> "%") And (LenB(.ModifiedBy) > 0)) Then
                    If ((.AddedOn <> .ModifiedOn) Or (.AddedBy <> .ModifiedBy)) Then
                        sRetMod = " was last modified by " & .ModifiedBy & _
                            " on " & .ModifiedOn & "."
                    Else
                        sRetMod = " have not been modified since they were added."
                    End If
                End If
                
                If ((LenB(sRetAdd) > 0) Or (LenB(sRetMod) > 0)) Then
                    If (LenB(sRetAdd) > 0) Then
                        GetDBDetail = DB(i).Username & sRetAdd & " They" & _
                            sRetMod
                    Else
                        'no add, but we could have a modify
                        GetDBDetail = DB(i).Username & sRetMod
                    End If
                Else
                    GetDBDetail = "No detailed information is available for that user."
                End If
                
                Exit Function
            End If
        End With
    Next i
    
    GetDBDetail = "That user was not found in the database."
End Function

' requires public
Public Function DateCleanup(ByVal tDate As Date) As String
    Dim t As String
    
    t = Format(tDate, "dd-MM-yyyy_HH:MM:SS")
    
    DateCleanup = Replace(t, " ", "_")
End Function

Private Function GetAccessINIValue(ByVal sKey As String, Optional ByVal Default As Long) As Long
    Dim s As String, l As Long
    
    s = ReadINI("Numeric", sKey, "access.ini")
    l = Val(s)
    
    If l > 0 Then
        GetAccessINIValue = l
    Else
        If Default > 0 Then
            GetAccessINIValue = Default
        Else
            GetAccessINIValue = 100
        End If
    End If
End Function

Private Function checkUser(ByVal user As String, Optional ByVal _
    allow_illegal As Boolean = False) As Boolean
    
    Dim i       As Integer ' ...
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
        For i = 1 To Len(user)
            ' ...
            Dim currentCharacter As String
            
            ' ...
            currentCharacter = Mid$(user, i, 1)

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
        Next i
    End If
    
    ' is our user valid?
    If (Not (invalid)) Then
        ' does our user contain illegal
        ' characters?
        If (illegal) Then
            ' do we allow illegal
            ' characters?
            If (allow_illegal) Then
                checkUser = True
            Else
                checkUser = False
            End If
        Else
            checkUser = True
        End If
    Else
        checkUser = False
    End If
End Function

Public Function convertUsername(ByVal Username As String) As String
    Dim index As Long ' ...
    
    If (Len(Username) < 1) Then
        convertUsername = Username
    
        Exit Function
    End If

    If ((StrReverse$(BotVars.Product) = "D2DV") Or _
        (StrReverse$(BotVars.Product) = "D2XP")) Then
        
        If ((Not (BotVars.UseGameConventions)) Or _
            (Not (BotVars.UseD2GameConventions))) Then
           
            index = InStr(1, Username, "*", vbBinaryCompare)
        
            If (index <> 0) Then
                convertUsername = Mid$(Username, index + 1)
            End If
        Else
            index = InStr(1, Username, "*", vbBinaryCompare)
        
            If (index > 1) Then
                convertUsername = Left$(Username, index - 1) & _
                    " (" & Mid$(Username, index) & ")"
            Else
                convertUsername = Username
            End If
        End If
    ElseIf ((StrReverse$(BotVars.Product) = "WAR3") Or _
            (StrReverse$(BotVars.Product) = "W3XP")) Then
            
        If ((Not (BotVars.UseGameConventions)) Or _
            (Not (BotVars.UseW3GameConventions))) Then

            If (BotVars.Gateway <> vbNullString) Then
                Select Case (BotVars.Gateway)
                    Case "Lordaeron": index = InStr(1, Username, "@USWest", vbBinaryCompare)
                    Case "Azeroth":   index = InStr(1, Username, "@USEast", vbBinaryCompare)
                    Case "Kalimdor":  index = InStr(1, Username, "@Asia", vbBinaryCompare)
                    Case "Northrend": index = InStr(1, Username, "@Europe", vbBinaryCompare)
                End Select
                
                If (index <> 0) Then
                    convertUsername = Left$(Username, index - 1)
                Else
                    convertUsername = Username & "@" & _
                        BotVars.Gateway
                End If
            End If
        End If
    End If
    
    If (convertUsername = vbNullString) Then
        convertUsername = Username
    End If
End Function

Public Function reverseUsername(ByVal Username As String) As String
    Dim index As Long ' ...
    
    If (Len(Username) < 1) Then
        Exit Function
    End If

    If ((StrReverse$(BotVars.Product) = "D2DV") Or _
        (StrReverse$(BotVars.Product) = "D2XP")) Then
        
        If ((Not (BotVars.UseGameConventions)) Or _
            (Not (BotVars.UseD2GameConventions))) Then
            
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
            
        If ((Not (BotVars.UseGameConventions)) Or _
            (Not (BotVars.UseW3GameConventions))) Then
            
            If (BotVars.Gateway <> vbNullString) Then
                index = InStr(1, Username, ("@" & BotVars.Gateway), vbBinaryCompare)
    
                If (index <> 0) Then
                    reverseUsername = Left$(Username, index - 1)
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

Public Function SecondsToString(ByVal seconds As Long) As String
    Dim Temp  As Long ' ...
    Dim secs  As Long ' ...
    Dim mins  As Long ' ...
    Dim hours As Long ' ...
    
    ' ...
    Temp = seconds
    
    ' ...
    Do While (Temp > 0)
        If (Temp - 3600 >= 0) Then
            ' ...
            Temp = (Temp - 3600)
            
            ' ...
            hours = (hours + 1)
        ElseIf (Temp - 60 >= 0) Then
            ' ...
            Temp = (Temp - 60)
            
            ' ...
            mins = (mins + 1)
        Else
            ' ...
            secs = Temp
                   
            ' ...
            Temp = 0
        End If
    Loop
    
    ' ...
    SecondsToString = IIf(hours, Right$("00" & hours, 2) & ":", vbNullString) & _
        Right$("00" & mins, 2) & ":" & Right$("00" & secs, 2)
End Function
