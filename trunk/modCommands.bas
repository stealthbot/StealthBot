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
Private m_WasWhispered  As Boolean ' ...
Private m_DisplayOutput As Boolean ' ...

Public flood    As String ' ...?
Public floodCap As Byte   ' ...?

' prepares commands for processing, and calls helper functions associated with
' processing
Public Function ProcessCommand(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As _
        Boolean = False, Optional ByVal WasWhispered As Boolean = False, Optional DisplayOutput As Boolean = _
                True) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
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
    m_WasWhispered = WasWhispered
    
    ' replace message variables
    Message = Replace(Message, "%me", IIf(IsLocal, GetCurrentUsername, Username), 1, -1, vbTextCompare)
    
    ' ...
    If (Left$(Message, 3) = "///") Then
        ' ...
        AddQ Mid$(Message, 3)
        
        ' ...
        Exit Function
    End If

    ' ...
    Set Command = IsCommand(Message, IsLocal)

    ' ...
    Do While (Command.Name <> vbNullString)
        ' ...
        If (Command.IsLocal) Then
            execCommand = True
        ElseIf (HasAccess(Username, Command.Name, Command.Args, outbuf)) Then
            execCommand = True
        End If
        
        ' ...
        m_DisplayOutput = Command.PublicOutput
        
        ' ...
        If (execCommand) Then
            ' ...
            If (IsLocal) Then
                With dbAccess
                    .Access = 201
                End With
            Else
                dbAccess = GetCumulativeAccess(Username)
            End If
            
            ' ...
            Call executeCommand(Username, dbAccess, Command.Name & Space$(1) & Command.Args, _
                    IsLocal, command_return)
                    
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
        Else
            ' ...
            If ((DisplayOutput) And (LenB(outbuf))) Then
                ' ...
                If ((BotVars.WhisperCmds) Or (WasWhispered)) Then
                    AddQ "/w " & Username & Space$(1) & command_return(I), _
                            PRIORITY.COMMAND_RESPONSE_MESSAGE
                Else
                    AddQ command_return(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
                End If
            End If
        End If
        
        ' ...
        Set Command = IsCommand(vbNullString, IsLocal)
        
        ' ...
        Count = (Count + 1)
        
        ' ...
        execCommand = False
    Loop
        
    ' ...
    If (IsLocal) Then
        ' ...
        If ((bln = False) And (Count = 0)) Then
            AddQ Message
        End If
    End If
    
    'Unload memory - FrOzeN
    Set Command = Nothing
    
    Exit Function
    
' default (if all else fails) error handler to keep erroneous
' commands and/or input formats from killing me
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: " & Err.description & _
        " in ProcessCommand().")

    'Unload memory - FrOzeN
    Set Command = Nothing

    ' return command failure result
    ProcessCommand = False
    
    Exit Function
End Function

' prepares commands for processing, and calls helper functions associated with
' processing
'Public Function ProcessCommand3(ByVal Username As String, ByVal Message As String, _
'    Optional ByVal InBot As Boolean = False, Optional ByVal WhisperedIn As Boolean = False) As Boolean
'
'    ' default error response for commands
'    On Error GoTo ERROR_HANDLER
'
'    ' stores the access response for use when commands are
'    ' issued via console
'    Dim ConsoleAccessResponse As udtGetAccessResponse
'
'    Dim i            As Integer ' loop counter
'    Dim tmpmsg       As String  ' stores local copy of message
'    Dim cmdRet()     As String  ' stores output of commands
'    Dim PublicOutput As Boolean ' stores result of public command
'                                ' output check (used for displaying command
'                                ' output when issuing via console)
'
'    ' create single command data array element for safe bounds checking
'    ReDim Preserve cmdRet(0)
'
'    ' create console access response structure
'    With ConsoleAccessResponse
'        .Access = 201
'        .Flags = "A"
'    End With
'
'    m_Username = Username
'    m_IsLocal = InBot
'    m_WasWhispered = WhisperedIn
'
'    If (m_console) Then
'        m_dbAccess = ConsoleAccessResponse
'    Else
'        m_dbAccess = GetCumulativeAccess(Username, "USER")
'    End If
'
'    ' store local copy of message
'    tmpmsg = Message
'
'    ' replace message variables
'    tmpmsg = Replace(tmpmsg, "%me", IIf((InBot), CurrentUsername, Username), 1)
'
'    ' check for command identifier when command
'    ' is not issued from within console
'    If (Not (InBot)) Then
'        ' we're going to mute our own access for now to prevent users from
'        ' using the "say" command to gain full control over the bot.
'        If (StrComp(Username, CurrentUsername, vbBinaryCompare) = 0) Then
'            Exit Function
'        End If
'
'        ' check for commands using universal command identifier (?)
'        If (StrComp(Left$(tmpmsg, Len("?trigger")), "?trigger", vbTextCompare) = 0) Then
'            ' remove universal command identifier from message
'            tmpmsg = Mid$(tmpmsg, 2)
'
'        ' check for commands using command identifier
'        ElseIf ((Len(tmpmsg) >= Len(BotVars.TriggerLong)) And _
'                (Left$(tmpmsg, Len(BotVars.TriggerLong)) = BotVars.TriggerLong)) Then
'
'            ' remove command identifier from message
'            tmpmsg = Mid$(tmpmsg, Len(BotVars.TriggerLong) + 1)
'
'            ' check for command identifier and name combination
'            ' (e.g., .Eric[nK] say hello)
'            If (Len(tmpmsg) >= (Len(CurrentUsername) + 1)) Then
'                If (StrComp(Left$(tmpmsg, Len(CurrentUsername) + 1), _
'                    CurrentUsername & Space(1), vbTextCompare) = 0) Then
'
'                    ' remove username (and space) from message
'                    tmpmsg = Mid$(tmpmsg, Len(CurrentUsername) + 2)
'                End If
'            End If
'
'        ' check for commands using either name and colon (and space),
'        ' or name and comma (and space)
'        ' (e.g., Eric[nK]: say hello; and, Eric[nK], say hello)
'        ElseIf ((Len(tmpmsg) >= (Len(CurrentUsername) + 2)) And _
'                ((StrComp(Left$(tmpmsg, Len(CurrentUsername) + 2), CurrentUsername & ": ", _
'                  vbTextCompare) = 0) Or _
'                 (StrComp(Left$(tmpmsg, Len(CurrentUsername) + 2), CurrentUsername & ", ", _
'                  vbTextCompare) = 0))) Then
'
'            ' remove username (and colon/comma) from message
'            tmpmsg = Mid$(tmpmsg, Len(CurrentUsername) + 3)
'        Else
'            ' allow commands without any command identifier if
'            ' commands are sent via whisper
'            'If (Not (WhisperedIn)) Then
'            '    ' return negative result indicating that message does not contain
'            '    ' a valid command identifier
'            '    ProcessCommand = False
'            '
'            '    ' exit function
'            '    Exit Function
'            'End If
'
'            ' return negative result indicating that message does not contain
'            ' a valid command identifier
'            ProcessCommand3 = False
'
'            ' exit function
'            Exit Function
'        End If
'    Else
'        ' remove slash (/) from in-console message
'        tmpmsg = Mid$(tmpmsg, 2)
'
'        ' check for second slash indicating
'        ' public output
'        If (Left$(tmpmsg, 1) = "/") Then
'            ' enable public display of command
'            PublicOutput = True
'
'            ' remove second slash (/) from in-console
'            ' message
'            tmpmsg = Mid$(tmpmsg, 2)
'        End If
'    End If
'
'    ' check for multiple command syntax if not issued from
'    ' within the console
'    If ((Not (InBot)) And _
'        (InStr(1, tmpmsg, "; ", vbBinaryCompare) > 0)) Then
'
'        Dim X() As String  ' ...
'
'        ' split message
'        X = Split(tmpmsg, "; ")
'
'        ' loop through commands
'        For i = 0 To UBound(X)
'            Dim tmpX As String ' ...
'
'            ' store local copy of message
'            tmpX = X(i)
'
'            ' can we check for a command identifier without
'            ' causing an rte?
'            If (Len(tmpX) >= Len(BotVars.TriggerLong)) Then
'                ' check for presence of command identifer
'                If (Left$(tmpX, Len(BotVars.TriggerLong)) = BotVars.TriggerLong) Then
'                    ' remove command identifier from message
'                    tmpX = Mid$(tmpX, Len(BotVars.TriggerLong) + 1)
'                End If
'            End If
'
'            ' execute command
'            ProcessCommand3 = ExecuteCommand(Username, GetCumulativeAccess(Username, _
'                "USER"), tmpX, InBot, cmdRet())
'
'            If (ProcessCommand3) Then
'                ' display command response
'                If (cmdRet(0) <> vbNullString) Then
'                    Dim j As Integer ' ...
'
'                    ' loop through command response
'                    For j = 0 To UBound(cmdRet)
'                        If ((InBot) And (Not (PublicOutput))) Then
'                            ' display message on screen
'                            Call frmChat.AddChat(RTBColors.ConsoleText, cmdRet(j))
'                        Else
'                            ' send message to battle.net
'                            If (WhisperedIn) Then
'                                ' whisper message
'                                Call AddQ("/w " & Username & Space$(1) & cmdRet(j), _
'                                    PRIORITY.COMMAND_RESPONSE_MESSAGE, Username)
'                            Else
'                                ' send standard message
'                                Call AddQ(cmdRet(j), PRIORITY.COMMAND_RESPONSE_MESSAGE, _
'                                    Username)
'                            End If
'                        End If
'                    Next j
'                End If
'            End If
'        Next i
'    Else
'        ' send command to main processor
'        If (InBot) Then
'            ' execute command
'            ProcessCommand3 = ExecuteCommand(Username, ConsoleAccessResponse, tmpmsg, _
'                InBot, cmdRet())
'        Else
'            ' execute command
'            ProcessCommand3 = ExecuteCommand(Username, GetCumulativeAccess(Username, _
'                "USER"), tmpmsg, InBot, cmdRet())
'        End If
'
'        If (ProcessCommand3) Then
'            ' display command response
'            If (cmdRet(0) <> vbNullString) Then
'                ' loop through command response
'                For i = 0 To UBound(cmdRet)
'                    If ((InBot) And (Not (PublicOutput))) Then
'                        ' display message on screen
'                        Call frmChat.AddChat(RTBColors.ConsoleText, cmdRet(i))
'                    Else
'                        ' display message
'                        If ((WhisperedIn) Or _
'                           ((BotVars.WhisperCmds) And (Not (InBot)))) Then
'
'                            ' whisper message
'                            Call AddQ("/w " & Username & _
'                                Space(1) & cmdRet(i), PRIORITY.COMMAND_RESPONSE_MESSAGE, _
'                                    Username)
'                        Else
'                            ' send standard message
'                            Call AddQ(cmdRet(i), PRIORITY.COMMAND_RESPONSE_MESSAGE, _
'                                Username)
'                        End If
'                    End If
'                Next i
'            End If
'        Else
'            ' send command directly to Battle.net if
'            ' command is found to be invalid and issued
'            ' internally
'            If (InBot) Then
'                Call AddQ(Message, PRIORITY.CONSOLE_MESSAGE, "(console)")
'            End If
'        End If
'    End If
'
'    ' break out of function before reaching error
'    ' handler
'    Exit Function
'
'' default (if all else fails) error handler to keep erroneous
'' commands and/or input formats from killing me
'ERROR_HANDLER:
'    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: " & Err.description & _
'        " in ProcessCommand().")
'
'    ' return command failure result
'    ProcessCommand3 = False
'
'    Exit Function
'End Function ' end function ProcessCommand

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
        Case "allowmp3":      Call OnAllowMp3(Username, dbAccess, msgData, InBot, cmdRet())
        Case "loadwinamp":    Call OnLoadWinamp(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "efp":           Call OnEfp(Username, dbAccess, msgData, InBot, cmdRet())
        Case "home":          Call OnHome(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clan":          Call OnClan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "peonban":       Call OnPeonBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "invite":        Call OnInvite(Username, dbAccess, msgData, InBot, cmdRet())
        Case "createclan":    Call OnCreateClan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "disbandclan":   Call OnDisbandClan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "makechieftain": Call OnMakeChieftain(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setmotd":       Call OnSetMotd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "where":         Call OnWhere(Username, dbAccess, msgData, InBot, cmdRet())
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
        Case "join":          Call OnJoin(Username, dbAccess, msgData, InBot, cmdRet())
        Case "sethome":       Call OnSetHome(Username, dbAccess, msgData, InBot, cmdRet())
        Case "resign":        Call OnResign(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clearbanlist":  Call OnClearBanList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "kickonyell":    Call OnKickOnYell(Username, dbAccess, msgData, InBot, cmdRet())
        Case "rejoin":        Call OnRejoin(Username, dbAccess, msgData, InBot, cmdRet())
        Case "forcejoin":     Call OnForceJoin(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quickrejoin":   Call OnQuickRejoin(Username, dbAccess, msgData, InBot, cmdRet())
        Case "plugban":       Call OnPlugBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clientbans":    Call OnClientBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setvol":        Call OnSetVol(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cadd":          Call OnCAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cdel":          Call OnCDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banned":        Call OnBanned(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ipbans":        Call OnIPBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ipban":         Call OnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unipban":       Call OnUnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "designate":     Call OnDesignate(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shuffle":       Call OnShuffle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "repeat":        Call OnRepeat(Username, dbAccess, msgData, InBot, cmdRet())
        Case "next":          Call OnNext(Username, dbAccess, msgData, InBot, cmdRet())
        Case "protect":       Call OnProtect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "whispercmds":   Call OnWhisperCmds(Username, dbAccess, msgData, InBot, cmdRet())
        Case "stop":          Call OnStop(Username, dbAccess, msgData, InBot, cmdRet())
        Case "play":          Call OnPlay(Username, dbAccess, msgData, InBot, cmdRet())
        Case "useitunes":     Call OnUseiTunes(Username, dbAccess, msgData, InBot, cmdRet())
        Case "usewinamp":     Call OnUseWinamp(Username, dbAccess, msgData, InBot, cmdRet())
        Case "pause":         Call OnPause(Username, dbAccess, msgData, InBot, cmdRet())
        Case "fos":           Call OnFos(Username, dbAccess, msgData, InBot, cmdRet())
        Case "rem":           Call OnRem(Username, dbAccess, msgData, InBot, cmdRet())
        Case "reconnect":     Call OnReconnect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unigpriv":      Call OnUnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
        Case "igpriv":        Call OnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
        Case "block":         Call OnBlock(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletime":      Call OnIdleTime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idle":          Call OnIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitdel":       Call OnShitDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safedel":       Call OnSafeDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagdel":        Call OnTagDel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "setidle":       Call OnSetIdle(Username, dbAccess, msgData, InBot, cmdRet())
        Case "idletype":      Call OnIdleType(Username, dbAccess, msgData, InBot, cmdRet())
        Case "filter":        Call OnFilter(Username, dbAccess, msgData, InBot, cmdRet())
        Case "trigger":       Call OnTrigger(Username, dbAccess, msgData, InBot, cmdRet())
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
        Case "fadd":          Call OnFAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "frem":          Call OnFRem(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safelist":      Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safeadd":       Call OnSafeAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "safecheck":     Call OnSafeCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "exile":         Call OnExile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unexile":       Call OnUnExile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitlist":      Call OnShitList(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagbans":       Call OnTagBans(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shitadd":       Call OnShitAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "dnd":           Call OnDND(Username, dbAccess, msgData, InBot, cmdRet())
        Case "bancount":      Call OnBanCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "banlistcount":  Call OnBanListCount(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tagcheck":      Call OnTagCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "slcheck":       Call OnSLCheck(Username, dbAccess, msgData, InBot, cmdRet())
        Case "readfile":      Call OnReadFile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "greet":         Call OnGreet(Username, dbAccess, msgData, InBot, cmdRet())
        Case "allseen":       Call OnAllSeen(Username, dbAccess, msgData, InBot, cmdRet())
        Case "profile":       Call OnProfile(Username, dbAccess, msgData, InBot, cmdRet())
        Case "accountinfo":   Call OnAccountInfo(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ban":           Call OnBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unban":         Call OnUnban(Username, dbAccess, msgData, InBot, cmdRet())
        Case "kick":          Call OnKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "lastwhisper":   Call OnLastWhisper(Username, dbAccess, msgData, InBot, cmdRet())
        Case "say":           Call OnSay(Username, dbAccess, msgData, InBot, cmdRet())
        Case "expand":        Call OnExpand(Username, dbAccess, msgData, InBot, cmdRet())
        Case "detail":        Call OnDetail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "info":          Call OnInfo(Username, dbAccess, msgData, InBot, cmdRet())
        Case "shout":         Call OnShout(Username, dbAccess, msgData, InBot, cmdRet())
        Case "voteban":       Call OnVoteBan(Username, dbAccess, msgData, InBot, cmdRet())
        Case "votekick":      Call OnVoteKick(Username, dbAccess, msgData, InBot, cmdRet())
        Case "vote":          Call OnVote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "tally":         Call OnTally(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cancel":        Call OnCancel(Username, dbAccess, msgData, InBot, cmdRet())
        Case "back":          Call OnBack(Username, dbAccess, msgData, InBot, cmdRet())
        Case "previous":      Call OnPrevious(Username, dbAccess, msgData, InBot, cmdRet())
        Case "uptime":        Call OnUptime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "away":          Call OnAway(Username, dbAccess, msgData, InBot, cmdRet())
        Case "mp3":           Call OnMP3(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ping":          Call OnPing(Username, dbAccess, msgData, InBot, cmdRet())
        Case "addquote":      Call OnAddQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "owner":         Call OnOwner(Username, dbAccess, msgData, InBot, cmdRet())
        Case "ignore":        Call OnIgnore(Username, dbAccess, msgData, InBot, cmdRet())
        Case "quote":         Call OnQuote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "unignore":      Call OnUnignore(Username, dbAccess, msgData, InBot, cmdRet())
        Case "cq":            Call OnCQ(Username, dbAccess, msgData, InBot, cmdRet())
        Case "scq":           Call OnSCQ(Username, dbAccess, msgData, InBot, cmdRet())
        Case "time":          Call OnTime(Username, dbAccess, msgData, InBot, cmdRet())
        Case "getping":       Call OnGetPing(Username, dbAccess, msgData, InBot, cmdRet())
        Case "checkmail":     Call OnCheckMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "inbox":         Call OnInbox(Username, dbAccess, msgData, InBot, cmdRet())
        Case "whoami":        Call OnWhoAmI(Username, dbAccess, msgData, InBot, cmdRet())
        Case "add":           Call OnAdd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "mmail":         Call OnMMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "bmail":         Call OnBMail(Username, dbAccess, msgData, InBot, cmdRet())
        Case "designated":    Call OnDesignated(Username, dbAccess, msgData, InBot, cmdRet())
        Case "flip":          Call OnFlip(Username, dbAccess, msgData, InBot, cmdRet())
        Case "about":         Call OnAbout(Username, dbAccess, msgData, InBot, cmdRet())
        Case "watch":         Call OnWatch(Username, dbAccess, msgData, InBot, cmdRet())
        Case "watchoff":      Call OnWatchOff(Username, dbAccess, msgData, InBot, cmdRet())
        Case "clear":         Call OnClear(Username, dbAccess, msgData, InBot, cmdRet())
        Case "server":        Call OnServer(Username, dbAccess, msgData, InBot, cmdRet())
        Case "find":          Call OnFind(Username, dbAccess, msgData, InBot, cmdRet())
        Case "whois":         Call OnWhoIs(Username, dbAccess, msgData, InBot, cmdRet())
        Case "findattr":      Call OnFindAttr(Username, dbAccess, msgData, InBot, cmdRet())
        Case "findgrp":       Call OnFindGrp(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "monitor":       Call OnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "unmonitor":     Call OnUnMonitor(Username, dbAccess, msgData, InBot, cmdRet())
        'Case "online":        Call OnOnline(Username, dbAccess, msgData, InBot, cmdRet())
        Case "help":          Call OnHelp(Username, dbAccess, msgData, InBot, cmdRet())
        Case "helpattr":      Call OnHelpAttr(Username, dbAccess, msgData, InBot, cmdRet())
        Case "helprank":      Call OnHelpRank(Username, dbAccess, msgData, InBot, cmdRet())
        Case "promote":       Call OnPromote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "demote":        Call OnDemote(Username, dbAccess, msgData, InBot, cmdRet())
        Case "connect":       Call OnConnect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "disconnect":    Call OnDisconnect(Username, dbAccess, msgData, InBot, cmdRet())
        Case "motd":          Call OnMotd(Username, dbAccess, msgData, InBot, cmdRet())
        Case "scripts":       Call OnScripts(Username, dbAccess, msgData, InBot, cmdRet())
        Case "enable":        Call OnEnable(Username, dbAccess, msgData, InBot, cmdRet())
        Case "disable":       Call OnDisable(Username, dbAccess, msgData, InBot, cmdRet())
        Case "sdetail":       Call OnSDetail(Username, dbAccess, msgData, InBot, cmdRet())
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

' handle connect command
Private Function OnConnect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    frmChat.DoConnect
    
End Function ' end function OnConnect

' handle disconnect command
Private Function OnDisconnect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    frmChat.DoDisconnect
    
End Function ' end function OnDisconnect

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

' handle allowmp3 command
Private Function OnAllowMp3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will enable or disable the use of media player-related commands.
    
    Dim tmpbuf As String ' temporary output buffer

    If (BotVars.DisableMP3Commands) Then
        tmpbuf = "Allowing MP3 commands."
        
        BotVars.DisableMP3Commands = False
    Else
        tmpbuf = "MP3 commands are now disabled."
        
        BotVars.DisableMP3Commands = True
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAllowMp3

' handle loadwinamp command
Private Function OnLoadWinamp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will run Winamp from the default directory, or the directory
    ' specified within the configuration file.
    
    Dim clsWinamp As clsWinamp
    Dim tmpbuf    As String  ' temporary output buffer
    Dim bln       As Boolean ' ...
    
    ' ...
    Set clsWinamp = New clsWinamp

    ' ...
    bln = clsWinamp.Start(ReadCfg("Other", "WinampPath"))
       
    ' ...
    If (bln) Then
        tmpbuf = "Winamp loaded."
    Else
        tmpbuf = "There was an error loading Winamp."
    End If
    
    ' ...
    Set clsWinamp = Nothing
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnLoadWinamp

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

' handle home command
Private Function OnHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot join its home channel.
    
    ' ...
    Call AddQ("/join " & BotVars.HomeChannel, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
        Username)
End Function ' end function OnHome

' handle clan command
Private Function OnClan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will allow the use of Battle.net's /clan command without requiring
    ' users be given the ability to use the bot's say command.
    
    Dim tmpbuf As String ' temporary output buffer

    Select Case (LCase$(msgData))
        Case "public", "pub"
            ' is bot a channel operator?
            If (g_Channel.Self.IsOperator) Then
                ' ...
                tmpbuf = "Clan channel is now public."
                
                ' set clan channel to public
                Call AddQ("/clan public", PRIORITY.CHANNEL_MODERATION_MESSAGE, _
                    Username)
            Else
                tmpbuf = "Error: The bot must have ops to change clan privacy status."
            End If
            
        Case "private", "priv"
            ' is bot a channel operator?
            If (g_Channel.Self.IsOperator) Then
                ' ...
                tmpbuf = "Clan channel is now private."
                
                ' set clan channel to private
                Call AddQ("/clan private", PRIORITY.CHANNEL_MODERATION_MESSAGE, _
                    Username)
            Else
                tmpbuf = "Error: The bot must have ops to change clan privacy status."
            End If
            
        Case Else
            ' set clan channel to specified
            Call AddQ("/clan " & msgData, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
                Username)
    End Select
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnClan

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

' handle OnMOTD command
Private Function OnMotd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String

    If (g_Clan.Self.Name <> vbNullString) Then
        tmpbuf = _
            "Clan " & g_Clan.Name & " MOTD: " & g_Clan.MOTD
    Else
        tmpbuf = "Error: I am not a member of a clan."
    End If
    
    cmdRet(0) = tmpbuf
End Function ' end function onMOTD

' handle invite command
Private Function OnInvite(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will send an invitation to the specified user to join the
    ' clan that the bot is currently either a Shaman or Chieftain of.  This
    ' command will only work if the bot is logged on using WarCraft III, and
    ' is either a Shaman, or a Chieftain of the clan in question.
    
    Dim tmpbuf As String ' temporary output buffer

    ' are we using warcraft iii?
    If (IsW3) Then
        ' is my ranking sufficient to issue
        ' an invitation?
        If (g_Clan.Self.Rank >= 3) Then
            Call InviteToClan(msgData)
            
            tmpbuf = msgData & ": Clan invitation sent."
        Else
            tmpbuf = "Error: The bot must hold Shaman or Chieftain rank to invite users."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnInvite

' handle createclan command
Private Function OnCreateClan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' return message
    cmdRet(0) = vbNullString
End Function ' end function OnCreateClan

' handle disbandclan command
Private Function OnDisbandClan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (g_Clan.Self.Rank >= 4) Then
        ' ...
        Call DisbandClan
    Else
        tmpbuf = "Error: You must be a chieftain to execute this command."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDisbandClan

' handle makechieftain command
Private Function OnMakeChieftain(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (Len(msgData) > 0) Then
        ' ...
        If (g_Clan.Self.Rank >= 4) Then
            ' ...
            Call MakeMemberChieftain(reverseUsername(msgData))
        Else
            tmpbuf = "Error: You must be a chieftain to execute this command."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnMakeChieftain

' handle setmotd command
Private Function OnSetMotd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the clan channel's Message Of The Day.  This
    ' command will only work if the bot is logged on using WarCraft III,
    ' and is either a Shaman or a Chieftain of the clan in question.
    
    Dim tmpbuf As String ' temporary output buffer

    If (IsW3) Then
        If (g_Clan.Self.Rank >= 3) Then
            Call SetClanMOTD(msgData)
            
            tmpbuf = "Clan MOTD set."
        Else
            tmpbuf = "Error: Shaman or Chieftain rank is required to set the MOTD."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetMotd

' handle where command
Private Function OnWhere(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will state the channel that the bot is currently
    ' residing in.  Battle.net uses this command to display basic
    ' user data, such as game type, and channel or game name.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' if sent from within the bot, send "where" command
    ' directly to Battle.net
    If (InBot) Then
        Call AddQ("/where " & msgData, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            "(console)")
    End If

    ' ...
    tmpbuf = "I am currently in channel " & g_Channel.Name & " (" & g_Channel.Users.Count & _
        " users present)"
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWhere

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
                frmChat.cboSend.text = vbNullString
            
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
        If (QueueLoad > 0) Then
            Call Pause(2, True, False)
        End If
        
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
        QueueLoad = (QueueLoad + 1)

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

' handle join command
Private Function OnJoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot join the specified channel.
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (LenB(msgData) > 0) Then
        ' ...
        Call AddQ("/join " & msgData, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            Username)
    End If
End Function ' end function OnJoin

' handle sethome command
Private Function OnSetHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the home channel to the channel specified.
    ' The home channel is the channel that the bot joins immediately
    ' following a completion of the connection procedure.
    
    Dim tmpbuf As String ' temporary output buffer

    Call WriteINI("Main", "HomeChan", msgData)
    
    BotVars.HomeChannel = msgData
    
    tmpbuf = "Home channel set to [ " & msgData & " ]"
    
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

' handle rejoin command
Private Function OnRejoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot rejoin the current channel.
    
    ' join temporary channel
    Call AddQ("/join " & GetCurrentUsername & " Rejoin", PRIORITY.COMMAND_RESPONSE_MESSAGE, _
        Username)
    
    ' rejoin previous channel
    Call AddQ("/join " & g_Channel.Name, PRIORITY.COMMAND_RESPONSE_MESSAGE, Username)
End Function ' end function OnRejoin

' handle quickrejoin command
Private Function OnQuickRejoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot rejoin the current channel.
    
    ' ...
    Call RejoinChannel(g_Channel.Name)
End Function ' end function OnRejoin

' handle forcejoin command
Private Function OnForceJoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will make the bot rejoin the current channel.
    
    ' ...
    If (Len(msgData) = 0) Then
        Exit Function
    End If
    
    ' ...
    Call FullJoin(msgData)
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
    Call searchDatabase(tmpbuf(), , , , "GAME", , , "B")
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnClientBans

' handle setvol command
Private Function OnSetVol(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will set the volume of the media player to the level
    ' specified by the user.
    
    Dim tmpbuf As String ' temporary output buffer
    Dim lngVol As Long   ' ...
    Dim strVol As String

    strVol = msgData
    
    If (BotVars.DisableMP3Commands = False) Then
        If (StrictIsNumeric(strVol)) Then
            lngVol = CLng(strVol)
            
        
            If (lngVol > 100) Then
                lngVol = 100
                strVol = 100
            End If
            
            MediaPlayer.Volume = lngVol

            tmpbuf = "Volume set to " & strVol & "%."
        Else
            tmpbuf = "Error: Invalid volume level (0-100)."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSetVol

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
        If (dbAccess.Access <= 100) Then
            If ((GetSafelist(tmpAcc)) Or (GetSafelist(msgFirstPart))) Then
                ' return message
                cmdRet(0) = "Error: That user is safelisted."
                
                Exit Function
            End If
        End If
        
        ' ...
        gAcc = GetAccess(msgFirstPart)
        
        ' ...
        If ((gAcc.Access >= dbAccess.Access) Or _
            ((InStr(gAcc.Flags, "A") > 0) And (dbAccess.Access <= 100))) Then

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
            tmpbuf = "I have designated [ " & msgData & " ]"
        Else
            ' ...
            tmpbuf = "Error: The bot does not have ops."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDesignate

' handle shuffle command
Private Function OnShuffle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will toggle the usage of the selected media player's
    ' shuffling feature.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        ' ...
        If (MediaPlayer.Shuffle) Then
            ' ...
            MediaPlayer.Shuffle = False
            
            ' ...
            tmpbuf = "The shuffle option has been disabled for the selected " & _
                "media player."
        Else
            ' ...
            MediaPlayer.Shuffle = True
            
            ' ...
            tmpbuf = "The shuffle option has been enabled for the selected " & _
                "media player."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnShuffle

' handle repeat command
Private Function OnRepeat(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    ' This command will toggle the usage of the selected media player's
    ' repeat feature.
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        ' ...
        If (MediaPlayer.Repeat) Then
            ' ...
            MediaPlayer.Repeat = False
            
            ' ...
            tmpbuf = "The repeat option has been disabled for the selected " & _
                "media player."
        Else
            ' ...
            MediaPlayer.Repeat = True
            
            ' ...
            tmpbuf = "The repeat option has been enabled for the selected " & _
                "media player."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnRepeat

' handle next command
Private Function OnNext(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        'Dim pos As Integer ' ...
        
        ' ...
        'pos = MediaPlayer.PlaylistPosition
    
        ' ...
        'Call MediaPlayer.PlayTrack(pos + 1)
        
        ' ...
        Call MediaPlayer.NextTrack
        
        ' ...
        tmpbuf = "Skipped forwards."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnNext

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

' handle stop command
Private Function OnStop(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        ' ...
        Call MediaPlayer.QuitPlayback
    
        ' ...
        tmpbuf = "Stopped playback."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnStop

' handle play command
Private Function OnPlay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    Dim Track  As Long   ' ...
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        ' ...
        If (MediaPlayer.IsLoaded() = False) Then
            MediaPlayer.Start
        End If

        ' ...
        MediaPlayer.PlayTrack msgData
        
        ' ...
        tmpbuf = "Playback started."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPlay

' handle useitunes command
Private Function OnUseiTunes(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    BotVars.MediaPlayer = "iTunes"
    
    ' ...
    tmpbuf = "iTunes is ready."
    
    ' ...
    Call WriteINI("Other", "MediaPlayer", "iTunes")
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUseiTunes

' handle usewinamp command
Private Function OnUseWinamp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    BotVars.MediaPlayer = "Winamp"
    
    ' ...
    tmpbuf = "Winamp is ready."
    
    ' ...
    Call WriteINI("Other", "MediaPlayer", "Winamp")

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUseWinamp

' handle pause command
Private Function OnPause(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (BotVars.DisableMP3Commands = False) Then
        If (MediaPlayer.IsLoaded()) Then
            Call MediaPlayer.PausePlayback
        
            tmpbuf = "Paused/resumed play."
        Else
            tmpbuf = MediaPlayer.Name & " is not loaded."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPause

' handle fos command
Private Function OnFos(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim hWndWA As Long
    Dim tmpbuf As String ' temporary output buffer

   If (BotVars.DisableMP3Commands = False) Then
        If (MediaPlayer.IsLoaded()) Then
            MediaPlayer.FadeOutToStop
        
            tmpbuf = "Fade-out stop."
        Else
            tmpbuf = MediaPlayer.Name & " is not loaded."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnFos

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
        If ((GetAccess(U, dbType).Access = -1) And _
            (GetAccess(U, dbType).Flags = vbNullString)) Then
            
            tmpbuf = "User not found."
        ElseIf (GetAccess(U, dbType).Access >= dbAccess.Access) Then
            tmpbuf = "That user has higher or equal access."
        ElseIf ((InStr(1, GetAccess(U, dbType).Flags, "L") <> 0) And _
                (Not (InBot)) And _
                (InStr(1, GetAccess(Username, dbType).Flags, "A") = 0) And _
                (GetAccess(Username, dbType).Access <= 99)) Then
            
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

' handle reconnect command
Private Function OnReconnect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmp As String ' ...
    
    If (g_Online) Then
        tmp = BotVars.HomeChannel
    
        BotVars.HomeChannel = g_Channel.Name
        
        Call frmChat.DoDisconnect
        
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Reconnecting by command, " & _
            "please wait..."
        
        Pause 1
        
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
        
        Pause 3, True
        
        BotVars.HomeChannel = tmp
    Else
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    End If
End Function ' end function OnReconnect

' handle unigpriv command
Private Function OnUnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    Call AddQ("/o unigpriv", PRIORITY.COMMAND_RESPONSE_MESSAGE, _
        Username)
    
    ' ...
    tmpbuf = "Recieving text from non-friends."
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUnIgPriv

' handle igpriv command
Private Function OnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    Call AddQ("/o igpriv", PRIORITY.COMMAND_RESPONSE_MESSAGE, _
        Username)
    
    ' ...
    tmpbuf = "Ignoring text from non-friends."
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIgPriv

' handle block command
Private Function OnBlock(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    Dim z      As String
    Dim I      As Integer

    U = msgData
    
    z = ReadINI("BlockList", "Total", "filters.ini")
    
    If (StrictIsNumeric(z)) Then
        I = z
    Else
        Call WriteINI("BlockList", "Total", "Total=0", "filters.ini")
        
        I = 0
    End If
    
    Call WriteINI("BlockList", "Filter" & (I + 1), U, "filters.ini")
    Call WriteINI("BlockList", "Total", I + 1, "filters.ini")
    
    tmpbuf = "Added """ & U & """ to the username block list."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnBlock

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
    Else
        ' remove user from shitlist using "add" command
        Call OnAdd(Username, dbAccess, "*" & msgData & "*" & " -B --type USER", _
            True, tmpbuf())
    End If
        
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnTagDel

' handle profile command
Private Function OnProfile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    U = msgData
    
    If (Len(U) > 0) Then
        If ((InBot = False) Or (m_DisplayOutput)) Then
            PPL = True
    
            ' ...
            If (BotVars.WhisperCmds Or m_WasWhispered) Then
                PPLRespondTo = Username
            End If
        Else
            frmProfile.lblUsername = U
        End If
        
        Call RequestProfile(U)
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnProfile

' handle accountinfo command
Private Function OnAccountInfo(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    If ((InBot = False) Or (m_DisplayOutput)) Then
        PPL = True

        ' ...
        If (BotVars.WhisperCmds Or m_WasWhispered) Then
            PPLRespondTo = Username
        End If
    End If
    
    RequestSystemKeys
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAccountInfo

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
        If (Left$(U, 1) = "/") Then
            U = Mid$(U, 2)
        End If
        
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

' handle filter command
Private Function OnFilter(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim I      As Integer
    Dim tmpbuf As String ' temporary output buffer
    Dim z      As String

    U = msgData
    
    z = ReadINI("TextFilters", "Total", "filters.ini")
    
    If (StrictIsNumeric(z)) Then
        I = z
    Else
        Call WriteINI("TextFilters", "Total", "Total=0", "filters.ini")
        
        I = 0
    End If
    
    Call WriteINI("TextFilters", "Filter" & (I + 1), U, "filters.ini")
    Call WriteINI("TextFilters", "Total", I + 1, "filters.ini")
    
    ReDim Preserve gFilters(UBound(gFilters) + 1)
    
    gFilters(UBound(gFilters)) = U
    
    tmpbuf = "Added " & Chr(34) & U & Chr(34) & " to the text message filter list."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnFilter

' handle trigger command
Private Function OnTrigger(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (LenB(BotVars.TriggerLong) = 1) Then
        ' ...
        tmpbuf = "The bot's current trigger is " & _
            Chr$(34) & Space$(1) & BotVars.TriggerLong & Space$(1) & Chr$(34) & _
                " (Alt + 0" & Asc(BotVars.TriggerLong) & ")"
    Else
        ' ...
        tmpbuf = "The bot's current trigger is " & _
            Chr$(34) & Space$(1) & BotVars.TriggerLong & Space$(1) & Chr$(34) & _
                " (Length: " & Len(BotVars.TriggerLong) & ")"
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnTrigger

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
        If ((Left$(newTrigger, 1) = Space$(1)) Or (Right$(newTrigger, 1) = Space$(1))) Then
            
            ' ...
            cmdRet(0) = "Error: Trigger may not begin or end with a " & _
                "space."
            
            ' ...
            Exit Function
        ElseIf (Left$(newTrigger, 1) = "/") Then
            ' ...
            cmdRet(0) = "Error: Trigger may not begin with a " & _
                "forward slash."
            
            ' ...
            Exit Function
        End If
        
        ' set new trigger
        BotVars.Trigger = newTrigger
    
        ' write trigger to configuration
        Call WriteINI("Main", "Trigger", newTrigger)
    
        ' ...
        tmpbuf = "The new trigger is " & Chr$(34) & newTrigger & _
            Chr$(34) & "."
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
        If (Len(user) = 0) Then
            ' ...
            tmpbuf(0) = "Error: The specified tag is invalid."
        Else
            ' ...
            user = "*" & user & "*"
        End If
    End If
    
    ' ...
    tag_msg = tag_msg & " --type USER"
    
    ' ...
    Call OnAdd(Username, dbAccess, user & tag_msg, True, tmpbuf())
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnTagAdd

' handle fadd command
Private Function OnFAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    U = msgData
    
    If (LenB(U) > 0) Then
        ' ...
        Call AddQ("/f a " & U, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            Username)
        
        ' ...
        tmpbuf = "Added user " & Chr(34) & U & Chr(34) & " to this account's friends list."
    End If
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnFAdd

' handle frem command
Private Function OnFRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    U = msgData
    
    If (Len(U) > 0) Then
        ' ...
        Call AddQ("/f r " & U, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            Username)
        
        ' ...
        tmpbuf = "Removed user " & Chr(34) & U & Chr(34) & " from this account's friends list."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnFRem

' handle safelist command
Private Function OnSafeList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    
    ' redefine array size
    ReDim Preserve tmpbuf(0)
    
    ' search database for shitlisted users
    Call searchDatabase(tmpbuf(), , , , , , , "S")
    
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
    Call searchDatabase(tmpbuf(), , "!*[*]*", , , , , "B")
    
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
    Call searchDatabase(tmpbuf(), , "*[*]*", , , , , "B")
    
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

' handle dnd command
Private Function OnDND(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim DNDMsg As String ' ...
    
    ' ...
    If (LenB(msgData) = 0) Then
        ' ...
        Call AddQ("/dnd", PRIORITY.COMMAND_RESPONSE_MESSAGE)
    Else
        ' ...
        DNDMsg = msgData
    
        ' ...
        Call AddQ("/dnd " & DNDMsg, PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            Username)
    End If
End Function ' end function OnDND

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

' handle allseen command
Private Function OnAllSeen(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim I        As Integer

    ' redefine array size
    ReDim Preserve tmpbuf(tmpCount)

    ' prefix message with "Last 15 users seen"
    tmpbuf(tmpCount) = "Last 15 users seen: "
    
    ' were there any users seen?
    If (colLastSeen.Count = 0) Then
        tmpbuf(tmpCount) = tmpbuf(tmpCount) & "(list is empty)"
    Else
        For I = 1 To colLastSeen.Count
            ' append user to list
            tmpbuf(tmpCount) = tmpbuf(tmpCount) & _
                colLastSeen.Item(I) & ", "
            
            If (Len(tmpbuf(tmpCount)) > 90) Then
                If (I < colLastSeen.Count) Then
                    ' redefine array size
                    ReDim Preserve tmpbuf(tmpCount + 1)
                    
                    ' clear new array index
                    tmpbuf(tmpCount + 1) = vbNullString
                    
                    ' remove ending comma from index
                    tmpbuf(tmpCount) = Mid$(tmpbuf(tmpCount), 1, _
                        Len(tmpbuf(tmpCount)) - Len(", "))
                
                    ' postfix [more] to end of entry
                    tmpbuf(tmpCount) = tmpbuf(tmpCount) & " [more]"
                    
                    ' increment loop counter
                    tmpCount = (tmpCount + 1)
                End If
            End If
        Next I
        
        ' check for ending comma
        If (Right$(tmpbuf(tmpCount), 2) = ", ") Then
            ' remove ending comma from index
            tmpbuf(tmpCount) = Mid$(tmpbuf(tmpCount), 1, _
                Len(tmpbuf(tmpCount)) - Len(", "))
        End If
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnAllSeen

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
                        vbNullString), dbAccess.Access)
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
                If (dbAccess.Access >= 100) Then
                    tmpbuf = WildCardBan(U, banmsg, 0)
                Else
                    tmpbuf = WildCardBan(U, banmsg, 0)
                End If
            Else
                If (InBot) Then
                    frmChat.AddQ "/kick " & msgData
                Else
                    Y = Ban(U & IIf(Len(banmsg) > 0, Space$(1) & banmsg, vbNullString), _
                        dbAccess.Access, 1)
                    
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

' handle lastwhisper command
Private Function OnLastWhisper(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (LastWhisper <> vbNullString) Then
        tmpbuf = "The last whisper to this bot was from " & LastWhisper & " at " & _
            FormatDateTime(LastWhisperFromTime, vbLongTime) & " on " & _
                FormatDateTime(LastWhisperFromTime, vbLongDate) & "."
        
    Else
        tmpbuf = "The bot has not been whispered since it logged on."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnLastWhisper

' handle say command
Private Function OnSay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf  As String ' temporary output buffer
    Dim tmpSend As String ' ...
    
    ' ...
    If (Len(msgData) > 0) Then
        ' ...
        If (Len(msgData) > 223) Then
            msgData = Mid$(msgData, 1, 223)
        End If
    
        ' ...
        Call AddQ(msgData, PRIORITY.COMMAND_RESPONSE_MESSAGE, Username)
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnSay

' handle expand command
Private Function OnExpand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf  As String ' temporary output buffer
    Dim tmpSend As String

    ' ...
    If (Len(msgData) > 0) Then
        ' ...
        tmpSend = Expand(msgData)
        
        ' ...
        If (Len(tmpSend) > 223) Then
            tmpSend = Mid$(tmpSend, 1, 223)
        End If
        
        ' ...
        Call AddQ(tmpSend, PRIORITY.COMMAND_RESPONSE_MESSAGE)
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnExpand

' handle detail command
Private Function OnDetail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = GetDBDetail(msgData)
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnDetail

' handle info command
Private Function OnInfo(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim user      As String
    Dim UserIndex As Integer
    Dim tmpbuf()  As String ' temporary output buffer
    
    ' ...
    user = msgData
    
    ' ...
    If (Len(user) > 0) Then
        ' ...
        ReDim Preserve tmpbuf(0 To 1)
    
        ' ...
        UserIndex = g_Channel.GetUserIndex(user)
        
        ' ...
        If (UserIndex > 0) Then
            ' ...
            With g_Channel.Users(UserIndex)
                ' ...
                tmpbuf(0) = "User " & .DisplayName & " is logged on using " & _
                    ProductCodeToFullName(.game)
                
                ' ...
                If (.IsOperator) Then
                    tmpbuf(0) = tmpbuf(0) & " with ops, and a ping time of " & .Ping & "ms."
                Else
                    tmpbuf(0) = tmpbuf(0) & " with a ping time of " & .Ping & "ms."
                End If
                
                ' ...
                tmpbuf(1) = "He/she has been present in the channel for " & _
                    ConvertTime(.TimeInChannel(), 1) & "."
            End With
        Else
            ' ...
            ReDim Preserve tmpbuf(0)
        
            ' ...
            tmpbuf(0) = "No such user is present."
        End If
    Else
        Exit Function
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnInfo

' handle shout command
Private Function OnShout(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf  As String ' temporary output buffer
    Dim tmpSend As String

    ' ...
    If (Len(msgData) > 0) Then
        ' ...
        tmpSend = UCase$(msgData)
        
        ' ...
        If (Len(tmpSend) > 223) Then
            tmpSend = Mid$(tmpSend, 1, 223)
        End If
        
        ' ...
        Call AddQ(tmpSend, PRIORITY.COMMAND_RESPONSE_MESSAGE)
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnShout

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

' handle back command
Private Function OnBack(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim hWndWA As Long   ' ...
    
    ' ...
    If (AwayMsg <> vbNullString) Then
        ' ...
        Call AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE, _
            Username)
        
        ' ...
        If (InBot = False) Then
            ' alert users of status change
            If (AwayMsg = " - ") Then
                Call AddQ("/me is back.", _
                    PRIORITY.COMMAND_RESPONSE_MESSAGE)
            Else
                Call AddQ("/me is back from (" & AwayMsg & ")", _
                    PRIORITY.COMMAND_RESPONSE_MESSAGE)
            End If
            
            ' set away message
            AwayMsg = vbNullString
        End If
    Else
        ' ...
        Call OnPrevious(Username, dbAccess, msgData, InBot, cmdRet())
    End If

End Function ' end function OnBack

' handle prev command
Private Function OnPrevious(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    Dim hWndWA As Long   ' ...
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        If (MediaPlayer.IsLoaded()) Then
            ' ...
            MediaPlayer.PreviousTrack
        
            ' ...
            tmpbuf = "Skipped backwards."
        Else
            tmpbuf = MediaPlayer.Name & " is not loaded."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPrev

' handle uptime command
Private Function OnUptime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = "System uptime " & ConvertTime(GetUptimeMS) & ", connection uptime " & _
        ConvertTime(uTicks) & "."
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUptime

' handle away command
Private Function OnAway(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    ' ...
    If (LenB(AwayMsg) > 0) Then
        ' send away command to battle.net
        Call AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE)
        
        ' alert users of status change
        If (InBot = False) Then
            If (AwayMsg = " - ") Then
                Call AddQ("/me is back.", _
                    PRIORITY.COMMAND_RESPONSE_MESSAGE)
            Else
                Call AddQ("/me is back from (" & AwayMsg & ")", _
                    PRIORITY.COMMAND_RESPONSE_MESSAGE)
            End If
        End If
        
        ' set away message
        AwayMsg = vbNullString
    Else
        If (LenB(msgData) > 0) Then
            ' set away message
            AwayMsg = msgData
            
            ' send away command to battle.net
            Call AddQ("/away " & AwayMsg, PRIORITY.COMMAND_RESPONSE_MESSAGE)
            
            ' alert users of status change
            If (InBot = False) Then
                ' ...
                Call AddQ("/me is away (" & AwayMsg & ")")
            End If
        Else
            ' set away message
            AwayMsg = " - "
        
            ' send away command to battle.net
            Call AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            
            ' alert users of status change
            If (InBot = False) Then
                Call AddQ("/me is away.")
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAway

' handle mp3 command
Private Function OnMP3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf       As String  ' temporary output buffer
    Dim TrackName    As String  ' ...
    Dim ListPosition As Long    ' ...
    Dim ListCount    As Long    ' ...
    Dim TrackTime    As Long    ' ...
    Dim TrackLength  As Long    ' ...
    
    ' ...
    If (BotVars.DisableMP3Commands = False) Then
        If (MediaPlayer.IsLoaded = False) Then
            tmpbuf = MediaPlayer.Name & " is not loaded."
        Else
            ' ...
            TrackName = MediaPlayer.TrackName
            ListPosition = MediaPlayer.PlaylistPosition
            ListCount = MediaPlayer.PlaylistCount
            TrackTime = MediaPlayer.TrackTime
            TrackLength = MediaPlayer.TrackLength
            
            ' ...
            If (TrackName = vbNullString) Then
                tmpbuf = MediaPlayer.Name & " is not currently playing any media."
            Else
                tmpbuf = "Current MP3 " & _
                    "[" & ListPosition & "/" & ListCount & "]: " & _
                        TrackName & " (" & SecondsToString(TrackTime) & _
                            "/" & SecondsToString(TrackLength)
                
                ' ...
                If (MediaPlayer.IsPaused) Then
                    tmpbuf = tmpbuf & ", paused)"
                Else
                    tmpbuf = tmpbuf & ")"
                End If
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnMP3

' handle ping command
Private Function OnPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf  As String ' temporary output buffer
    Dim Latency As Long
    Dim user    As String
    
    user = msgData
    
    If (user <> vbNullString) Then
        Latency = GetPing(user)
        
        If (Latency >= -1) Then
            tmpbuf = user & "'s ping at login was " & Latency & "ms."
        Else
            tmpbuf = "I can't see " & user & " in the channel."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnPing

' handle addquote command
Private Function OnAddQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    Dim U      As String
    Dim Y      As String
    Dim tmpbuf As String ' temporary output buffer
    
    f = FreeFile
    
    U = msgData
    
    If (Len(U)) Then
        Y = Dir$(GetFilePath("quotes.txt"))
        
        If (Len(Y) = 0) Then
            Open GetFilePath("quotes.txt") For Output As #f
                Print #f, U
            Close #f
        Else
            Open Y For Append As #f
                Print #f, U
            Close #f
        End If
            
        tmpbuf = "Quote added!"
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAddQuote

' handle owner command
Private Function OnOwner(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    If (LenB(BotVars.BotOwner)) Then
        tmpbuf = "This bot's owner is " & _
            BotVars.BotOwner & "."
    Else
        tmpbuf = "There is no owner currently set."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnOwner

' handle ignore command
Private Function OnIgnore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
        
    U = Split(msgData, " ")(0)
    
    If (U <> vbNullString) Then
        ' ...
        If (GetAccess(U).Access >= dbAccess.Access) Then
            ' ...
            tmpbuf = "That user has equal or higher access."
        Else
            ' ...
            AddQ "/ignore " & U
            
            ' ...
            tmpbuf = "Ignoring messages from " & Chr(34) & U & Chr(34) & "."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnIgnore

' handle quote command
Private Function OnQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = "Quote: " & _
        GetRandomQuote
    
    If (Len(tmpbuf) = 0) Then
        tmpbuf = "Error reading quotes, or no quote file exists."
    ElseIf (Len(tmpbuf) > 223) Then
        ' try one more time
        tmpbuf = "Quote: " & _
            GetRandomQuote
        
        If (Len(tmpbuf) > 223) Then
            'too long? too bad. truncate
            tmpbuf = Left$(tmpbuf, 223)
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnQuote

' handle unignore command
Private Function OnUnignore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim U      As String
    Dim tmpbuf As String ' temporary output buffer
    
    U = Split(msgData, " ")(0)
    
    ' ...
    If (msgData <> vbNullString) Then
        ' ...
        AddQ "/unignore " & U
        
        ' ...
        tmpbuf = "Receiving messages from """ & U & """."
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnUnignore

' handle cq command
Private Function OnCQ(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (msgData = vbNullString) Then
        ' ...
        g_Queue.Clear
    
        ' ...
        tmpbuf = "Queue cleared."
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnCQ

' handle scq command
Private Function OnSCQ(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    ' ...
    If (msgData = vbNullString) Then
        Call g_Queue.Clear
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnCQ

' handle time command
Private Function OnTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = "The current time on this computer is " & Time & " on " & _
        Format(Date, "MM-dd-yyyy") & "."
            
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnTime

' handle getping command
Private Function OnGetPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf  As String ' temporary output buffer
    Dim Latency As Long

    If (InBot) Then
        If (g_Online) Then
            ' grab current latency
            Latency = GetPing(GetCurrentUsername)
        
            ' ...
            tmpbuf = "Your ping at login was " & Latency & "ms."
        Else
            ' ...
            tmpbuf = "Error: You are not connected."
        End If
    Else
        ' ...
        Latency = GetPing(Username)
    
        ' ...
        If (Latency >= -1) Then
            tmpbuf = "Your ping at login was " & Latency & "ms."
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnGetPing

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
        If (dbAccess.Access > 0) Then
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

' handle whoami command
Private Function OnWhoAmI(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpbuf As String ' temporary output buffer

    If (InBot) Then
        tmpbuf = "You are the bot console."
        
        If (g_Online) Then
            Call AddQ("/whoami", PRIORITY.CONSOLE_MESSAGE)
        End If
    ElseIf (dbAccess.Access = 1000) Then
        tmpbuf = "You are the bot owner, " & Username & "."
    Else
        If (dbAccess.Access > 0) Then
            If (dbAccess.Flags <> vbNullString) Then
                tmpbuf = dbAccess.Username & " has access " & dbAccess.Access & _
                    " and flags " & dbAccess.Flags & "."
            Else
                tmpbuf = dbAccess.Username & " has access " & dbAccess.Access & "."
            End If
        Else
            If (dbAccess.Flags <> vbNullString) Then
                tmpbuf = dbAccess.Username & " has flags " & dbAccess.Flags & "."
            End If
        End If
    End If

    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWhoAmI

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
        If ((Rank <= 0) And (Flags = vbNullString) And _
                (sGrp = vbNullString) And (dbType = vbNullString)) Then
            
            tmpbuf = "Error: You have specified an invalid rank."
            
        ' is rank higher than user's rank?
        ElseIf ((Rank) And (Rank >= dbAccess.Access)) Then
            tmpbuf = "Error: You do not have sufficient access to assign an entry with the " & _
                "specified rank."
            
        ' can we modify specified user?
        ElseIf ((gAcc.Access) And (gAcc.Access >= dbAccess.Access)) Then
            tmpbuf = "Error: You do not have sufficient access to modify the specified entry."
        Else
            ' did we specify flags?
            If (Len(Flags)) Then
                Dim currentCharacter As String ' ...
            
                For I = 1 To Len(Flags)
                    currentCharacter = Mid$(Flags, I, 1)
                
                    If ((currentCharacter <> "+") And (currentCharacter <> "-")) Then
                        'Select Case (currentCharacter)
                        '    Case "A" ' administrator
                        '        If (dbAccess.Access <= 100) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "B" ' banned
                        '        If (dbAccess.Access < 70) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "D" ' designated
                        '        If (dbAccess.Access < 100) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "L" ' locked
                        '        If (dbAccess.Access < 70) Then
                        '            Exit For
                        '        End If
                        '
                        '    Case "S" ' safelisted
                        '        If (dbAccess.Access < 70) Then
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
                gAcc.Access = Rank
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
                        .Access = gAcc.Access
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
                            DB(I).Type, DB(I).Access, DB(I).Flags, DB(I).Groups)
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
                
                'MsgBox dbPath
                
                ' commit modifications
                Call WriteDatabase(dbPath)
                
                ' log actions
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(AddEntry, IIf(InBot, "console", Username), DB(UBound(DB)).Username, _
                        DB(UBound(DB)).Type, DB(UBound(DB)).Access, DB(UBound(DB)).Flags, DB(UBound(DB)).Groups)
                End If
            End If
            
            ' check for errors & create message
            If (gAcc.Access > 0) Then
                tmpbuf = Chr(34) & user & Chr(34) & " has been given access " & _
                    gAcc.Access
                
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
                    gAcc = GetCumulativeAccess(DB(c).Username)
                    
                    If (StrComp(gAcc.Type, "USER", vbTextCompare) = 0) Then
                        If (gAcc.Access = Track) Then
                            .To = DB(c).Username
                            
                            Call AddMail(temp)
                        End If
                    End If
                Next c
                
                tmpbuf = tmpbuf & "to users with access " & Track
            Else
                For c = 0 To UBound(DB)
                    gAcc = GetCumulativeAccess(DB(c).Username)
                
                    For f = 1 To Len(strArray(0))
                        If (StrComp(gAcc.Type, "USER", vbTextCompare) = 0) Then
                            If (InStr(1, gAcc.Flags, Mid$(strArray(0), f, 1), _
                                vbBinaryCompare) > 0) Then
                                
                                .To = DB(c).Username
                                
                                Call AddMail(temp)
                                
                                Exit For
                            End If
                        End If
                    Next f
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

' handle about command
Private Function OnAbout(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer

    tmpbuf = ".: " & CVERSION & " by Stealth."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnAbout

' handle watch command
Private Function OnWatch(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    WatchUser = msgData
    
    tmpbuf = "Watching " & WatchUser
        
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWatch

' handle watchoff command
Private Function OnWatchOff(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    WatchUser = vbNullString
    
    tmpbuf = "Watch off."
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWatchOff

' handle clear command
Private Function OnClear(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf As String ' temporary output buffer
    
    frmChat.mnuClear_Click
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnClear

' handle server command
Private Function OnServer(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf       As String ' temporary output buffer
    Dim RemoteHost   As String ' ...
    Dim RemoteHostIP As String ' ...
    
    ' ...
    RemoteHost = frmChat.sckBNet.RemoteHost
    
    ' ...
    RemoteHostIP = frmChat.sckBNet.RemoteHostIP
    
    ' ...
    If (StrComp(RemoteHost, RemoteHostIP, vbBinaryCompare) = 0) Then
        tmpbuf = "I am currently connected to " & _
            frmChat.sckBNet.RemoteHostIP & "."
    Else
        tmpbuf = "I am currently connected to " & _
            frmChat.sckBNet.RemoteHost & " (" & _
                frmChat.sckBNet.RemoteHostIP & ")."
    End If
            
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnServer

' handle find command
Private Function OnFind(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc     As udtGetAccessResponse

    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    
    ReDim Preserve tmpbuf(0)

    U = GetFilePath("users.txt")
            
    If (Dir$(U) = vbNullString) Then
        tmpbuf(0) = "No userlist available. Place a users.txt file" & _
            "in the bot's root directory."
    End If
    
    U = msgData
    
    If (Len(U) > 0) Then
        If (StrictIsNumeric(U)) Then
            ' execute search
            Call searchDatabase(tmpbuf(), , , , , Val(U))
        ElseIf (InStr(1, U, Space(1), vbBinaryCompare) <> 0) Then
            Dim lowerBound As String ' ...
            Dim upperBound As String ' ...
            
            ' grab range Values()
            If (InStr(1, U, " - ", vbBinaryCompare) <> 0) Then
                lowerBound = Mid$(U, 1, InStr(1, U, " - ", vbBinaryCompare) - 1)
                upperBound = Mid$(U, InStr(1, U, " - ", vbBinaryCompare) + Len(" - "))
            Else
                lowerBound = Mid$(U, 1, InStr(1, U, Space(1), vbBinaryCompare) - 1)
                upperBound = Mid$(U, InStr(1, U, Space(1), vbBinaryCompare) + 1)
            End If
            
            If ((StrictIsNumeric(lowerBound)) And _
                (StrictIsNumeric(upperBound))) Then
            
                ' execute search
                Call searchDatabase(tmpbuf(), , , , , CInt(Val(lowerBound)), CInt(Val(upperBound)))
            Else
                tmpbuf(0) = "Error: You have specified an invalid range."
            End If
        ElseIf ((InStr(1, U, "*", vbBinaryCompare) <> 0) Or _
                (InStr(1, U, "?", vbBinaryCompare) <> 0)) Then
            
            ' execute search
            Call searchDatabase(tmpbuf(), , PrepareCheck(U))
        Else
            ' execute search
            Call searchDatabase(tmpbuf(), , PrepareCheck(U))
        End If
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnFind

' handle whois command
Private Function OnWhoIs(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    Dim gAcc     As udtGetAccessResponse
    
    Dim tmpbuf   As String ' temporary output buffer
    Dim U        As String

    U = msgData
            
    If (InBot) Then
        Call AddQ("/whois " & U, PRIORITY.CONSOLE_MESSAGE)
    End If

    If (Len(U)) Then
        gAcc = GetCumulativeAccess(U)
        
        If (gAcc.Username <> vbNullString) Then
            If (gAcc.Access > 0) Then
                If (gAcc.Flags <> vbNullString) Then
                    tmpbuf = gAcc.Username & " has access " & gAcc.Access & _
                        " and flags " & gAcc.Flags & "."
                Else
                    tmpbuf = gAcc.Username & " has access " & gAcc.Access & "."
                End If
            Else
                If (gAcc.Flags <> vbNullString) Then
                    tmpbuf = gAcc.Username & " has flags " & gAcc.Flags & "."
                End If
            End If
        Else
            tmpbuf = "There was no such user found."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpbuf
End Function ' end function OnWhoIs

' handle findattr command
Private Function OnFindAttr(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim I        As Integer
    Dim found    As Integer
    
    ReDim Preserve tmpbuf(tmpCount)
    
    ' ...
    U = msgData

    If (Len(U) > 0) Then
        ' execute search
        Call searchDatabase(tmpbuf(), , , , , , , U)
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnFindAttr

' handle findgrp command
Private Function OnFindGrp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim U        As String
    Dim tmpbuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim I        As Integer
    Dim found    As Integer
    
    ReDim Preserve tmpbuf(tmpCount)
    
    ' ...
    U = msgData

    If (Len(U) > 0) Then
        ' execute search
        Call searchDatabase(tmpbuf(), , , U)
    End If
    
    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnFindAttr

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

' handle help command
Private Function OnHelp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpbuf()    As String ' temporary output buffer
    Dim CommandDocs As clsCommandDocObj
    Dim FindCommand As String
    Dim spaceIndex  As Integer
    Dim I           As Integer
    
    ' ...
    ReDim Preserve tmpbuf(0)
    
    ' ...
    spaceIndex = InStr(1, msgData, Space$(1), vbBinaryCompare)
    
    ' ...
    If (spaceIndex <> 0) Then
        ' ...
        FindCommand = Mid$(msgData, 1, spaceIndex - 1)
    Else
        ' ...
        FindCommand = msgData
    End If
    
    ' ...
    Set CommandDocs = OpenCommand(FindCommand)
    
    ' ...
    If (CommandDocs.Name = vbNullString) Then
        ' ...
        Set CommandDocs = OpenCommand(convertAlias(FindCommand))
    
        ' ...
        If (CommandDocs.Name = vbNullString) Then
            ' ...
            cmdRet(0) = "Sorry, but no related documentation could be found."
        
            ' ...
            Exit Function
        End If
    End If
    
    tmpbuf(0) = "[" & CommandDocs.Name
    
    If (CommandDocs.Aliases.Count) Then
        tmpbuf(0) = tmpbuf(0) & " (aliases: "
    Else
        tmpbuf(0) = tmpbuf(0) & " (aliases: none"
    End If
    
    If (CommandDocs.Aliases.Count) Then
        For I = 1 To CommandDocs.Aliases.Count
            tmpbuf(0) = tmpbuf(0) & CommandDocs.Aliases(I) & ", "
        Next I
        
        tmpbuf(0) = Mid$(tmpbuf(0), 1, Len(tmpbuf(0)) - Len(", "))
    End If

    ' ...
    tmpbuf(0) = tmpbuf(0) & ")]: " & CommandDocs.description
    
    ' ...
    tmpbuf(0) = tmpbuf(0) & Space$(1) & "(Syntax: " & "<trigger>" & CommandDocs.Name
            
    If (CommandDocs.Parameters.Count) Then
        For I = 1 To CommandDocs.Parameters.Count
            If (CommandDocs.Parameters(I).IsOptional) Then
                tmpbuf(0) = tmpbuf(0) & " [" & CommandDocs.Parameters(I).Name & "]"
            Else
                tmpbuf(0) = tmpbuf(0) & " <" & CommandDocs.Parameters(I).Name & ">"
            End If
        Next I
    End If
    
    tmpbuf(0) = tmpbuf(0) & "). "
    
    If (CommandDocs.IsEnabled = False) Then
        tmpbuf(0) = tmpbuf(0) & " Command is currently disabled"
    Else
        If ((CommandDocs.RequiredRank = -1) And _
                (CommandDocs.RequiredFlags = vbNullString)) Then
        
            tmpbuf(0) = tmpbuf(0) & " Command is only available to the console"
        ElseIf (CommandDocs.RequiredRank = 0) Then
            tmpbuf(0) = tmpbuf(0) & " Command is available to everyone"
        Else
            tmpbuf(0) = tmpbuf(0) & " Requires "
            
            If (CommandDocs.RequiredRank > 0) Then
                tmpbuf(0) = _
                    tmpbuf(0) & CommandDocs.RequiredRank & " access"
                    
                If (CommandDocs.RequiredFlags <> vbNullString) Then
                    tmpbuf(0) = tmpbuf(0) & " or "
                End If
            End If
            
            If (CommandDocs.RequiredFlags <> vbNullString) Then
                tmpbuf(0) = tmpbuf(0) & "flags "
                
                For I = 1 To Len(CommandDocs.RequiredFlags)
                    tmpbuf(0) = _
                        tmpbuf(0) & Mid$(CommandDocs.RequiredFlags, I, 1) & ", "
                            
                    If (I + 1 = Len(CommandDocs.RequiredFlags)) Then
                        tmpbuf(0) = tmpbuf(0) & "or "
                    End If
                Next I
                
                tmpbuf(0) = Mid$(tmpbuf(0), 1, Len(tmpbuf(0)) - Len(", "))
            End If
        End If
    End If
    
    tmpbuf(0) = tmpbuf(0) & "."

    ' return message
    cmdRet() = tmpbuf()
End Function ' end function OnHelp

' handle helpattr command
Private Function OnHelpAttr(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String  ' temporary output buffer
    Dim I           As Integer ' ...
    Dim xmldoc      As DOMDocument60
    Dim commands    As IXMLDOMNodeList
    Dim flagstr     As String
    Dim lastCommand As String
    Dim thisCommand As String
        
    ' ...
    Set xmldoc = New DOMDocument60
    
    ' ...
    If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
        Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
            "working directory.")
            
        Exit Function
    End If
    
    ' ...
    xmldoc.Load App.Path & "\commands.xml"
    
    ' ...
    If (InStr(1, msgData, "'", vbBinaryCompare) > 0) Then
        Exit Function
    End If

    ' ...
    msgData = Replace(msgData, "\", "\\")
    
    ' ...
    If (Len(msgData) = 0) Then
        Exit Function
    End If
    
    ' ...
    If (BotVars.CaseSensitiveFlags = False) Then
        msgData = UCase$(msgData)
    End If
    
    ' ...
    For I = 1 To Len(msgData)
        flagstr = flagstr & _
            "'" & Mid$(msgData, I, 1) & "' or "
    Next I
    
    ' ...
    flagstr = _
        Left$(flagstr, Len(flagstr) - 3)
        
    ' ...
    Set commands = _
        xmldoc.documentElement.selectNodes( _
            "./command/access/flags/flag[text()=" & flagstr & "]")
    
    ' ...
    If (commands.Length > 0) Then
        For I = 0 To commands.Length - 1
            thisCommand = commands(I).parentNode.parentNode.parentNode. _
                Attributes.getNamedItem("name").text
                
            If (StrComp(thisCommand, lastCommand, vbTextCompare) <> 0) Then
                tmpbuf = tmpbuf & thisCommand & ", "
            End If
            
            lastCommand = thisCommand
        Next I
        
        ' ...
        tmpbuf = _
            Left$(tmpbuf, Len(tmpbuf) - 2)
        
        tmpbuf = "Commands available to specified flag(s): " & tmpbuf
    Else
        tmpbuf = "No commands are available to the given flag(s)."
    End If
    
    ' ...
    cmdRet(0) = tmpbuf
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in OnHelpAttr()."

    Exit Function

End Function ' end function OnHelpAttr

' handle helprank command
Private Function OnHelpRank(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String  ' temporary output buffer
    Dim I           As Integer ' ...
    Dim xmldoc      As DOMDocument60
    Dim commands    As IXMLDOMNodeList
    Dim flagstr     As String
    Dim lastCommand As String
    Dim thisCommand As String
        
    ' ...
    Set xmldoc = New DOMDocument60
    
    ' ...
    If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
        Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
            "working directory.")
            
        Exit Function
    End If
    
    ' ...
    xmldoc.Load App.Path & "\commands.xml"
    
    ' ...
    If (InStr(1, msgData, "'", vbBinaryCompare) > 0) Then
        Exit Function
    End If

    ' ...
    msgData = Replace(msgData, "\", "\\")
    
    ' ...
    If (Len(msgData) = 0) Then
        Exit Function
    End If
        
    ' ...
    Set commands = _
        xmldoc.documentElement.selectNodes( _
            "./command/access/rank[number() <= " & msgData & "]")
    
    ' ...
    If (commands.Length > 0) Then
        For I = 0 To commands.Length - 1
            thisCommand = commands(I).parentNode.parentNode.Attributes. _
                getNamedItem("name").text
                
            If (StrComp(thisCommand, lastCommand, vbTextCompare) <> 0) Then
                tmpbuf = tmpbuf & thisCommand & ", "
            End If
            
            lastCommand = thisCommand
        Next I
        
        ' ...
        tmpbuf = _
            Left$(tmpbuf, Len(tmpbuf) - 2)
        
        tmpbuf = "Commands available to specified rank: " & tmpbuf
    Else
        tmpbuf = "No commands are available to the given rank."
    End If
    
    ' ...
    cmdRet(0) = tmpbuf
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in OnHelpAttr()."

    Exit Function

End Function ' end function OnHelpRank

' handle promote command
Private Function OnPromote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    ' ...
    If (Len(msgData) > 0) Then
        Dim liUser As ListItem ' ...
        
        ' ...
        Set liUser = _
            frmChat.lvClanList.FindItem(msgData)
    
        ' ...
        If (liUser Is Nothing) Then
            ' ...
            cmdRet(0) = "Error: The specified user is not currently a member of " & _
                    "this clan."
                    
            ' ...
            Exit Function
        End If
        
        ' ...
        If (liUser.SmallIcon >= 3) Then
            ' ...
            cmdRet(0) = "Error: The specified user is already at the highest promotable " & _
                    "ranking."
        
            ' ...
            Exit Function
        End If
    
        ' ...
        Call PromoteMember(liUser.text, liUser.SmallIcon + 1)
        
        ' ...
        'If (InBot = False) Then
        '    ' ...
        '    cmdRet(0) = Chr$(34) & liUser.text & Chr$(34) & " has been promoted to " & _
        '            GetRank(liUser.SmallIcon + 1) & "."
        'End If
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
            frmChat.lvClanList.FindItem(msgData)
    
        ' ...
        If (liUser Is Nothing) Then
            ' ...
            cmdRet(0) = "Error: The specified user is not currently a member of " & _
                    "this clan."
                    
            ' ...
            Exit Function
        End If
        
        ' ...
        If (liUser.SmallIcon <= 1) Then
            ' ...
            cmdRet(0) = "Error: The specified user is already at the lowest demoteable " & _
                    "ranking."
        
            ' ...
            Exit Function
        End If
        
        ' ...
        Call DemoteMember(liUser.text, liUser.SmallIcon - 1)
        
        ' ...
        'If (InBot = False) Then
        '    ' ...
        '    cmdRet(0) = Chr$(34) & liUser.text & Chr$(34) & " has been demoted to " & _
        '            GetRank(liUser.SmallIcon - 1) & "."
        'End If
    End If
End Function

' handle scripts command
Private Function OnScripts(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error Resume Next
    
    Dim tmpbuf As String  ' ...
    Dim I      As Integer ' ...
    Dim str    As String  ' ...
    Dim Name   As String  ' ...
    
    ' ...
    If (frmChat.SControl.Modules.Count) Then
        tmpbuf = _
            "Loaded Scripts (" & frmChat.SControl.Modules.Count & "): "
                 
        ' ...
        For I = 1 To frmChat.SControl.Modules.Count
            Name = _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
            
            If (Name = "PluginSystem") Then
                str = _
                    ReadINI("Override", "DisablePS", GetConfigFilePath())
            
                If (StrComp(str, "Y", vbTextCompare) = 0) Then
                    Name = "(" & Name & "), "
                Else
                    Name = Name & ", "
                End If
            Else
                str = _
                    frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Enabled")
            
                If (StrComp(str, "False", vbTextCompare) = 0) Then
                    Name = "(" & Name & "), "
                Else
                    Name = Name & ", "
                End If
            End If
            
            tmpbuf = tmpbuf & Name
        Next I
        
        ' ...
        tmpbuf = Mid$(tmpbuf, 1, Len(tmpbuf) - 2)
    Else
        tmpbuf = "There are no scripts currently loaded."
    End If
    
    ' ...
    cmdRet(0) = tmpbuf
    
End Function

' handle enable command
Private Function OnEnable(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Name As String  ' ...
    Dim I    As Integer ' ...
    
    ' ...
    If (frmChat.SControl.Modules.Count) Then
        For I = 1 To frmChat.SControl.Modules.Count
            Name = _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
                
            If (StrComp(Name, msgData, vbTextCompare) = 0) Then
                frmChat.SControl.Modules(I).CodeObject.WriteSettingsEntry _
                    "Enabled", "True"
                    
                InitScript frmChat.SControl.Modules(I)
                    
                cmdRet(0) = Name & " has been enabled."
            
                Exit Function
            End If
        Next I
    End If
    
    cmdRet(0) = "Error: Could not find specified script."
    
End Function

' handle disable command
Private Function OnDisable(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Name As String  ' ...
    Dim I    As Integer ' ...
    
    ' ...
    If (frmChat.SControl.Modules.Count) Then
        For I = 1 To frmChat.SControl.Modules.Count
            Name = _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
                
            If (StrComp(Name, msgData, vbTextCompare) = 0) Then
                frmChat.SControl.Modules(I).CodeObject.WriteSettingsEntry _
                    "Enabled", "False"
                    
                DestroyObjs frmChat.SControl.Modules(I)
                    
                cmdRet(0) = Name & " has been disabled."
            
                Exit Function
            End If
        Next I
    End If
    
    cmdRet(0) = "Error: Could not find specified script."
    
End Function

' handle sdetail command
Private Function OnSDetail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Name As String  ' ...
    Dim I    As Integer ' ...
    
    ' ...
    If (frmChat.SControl.Modules.Count) Then
        For I = 1 To frmChat.SControl.Modules.Count
            Name = _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
                
            If (StrComp(Name, msgData, vbTextCompare) = 0) Then
                Dim version As String ' ...
                Dim author  As String ' ...
                
                version = _
                    frmChat.SControl.Modules(I).CodeObject.Script("Major")
                    
                version = version & "." & _
                    frmChat.SControl.Modules(I).CodeObject.Script("Minor")
                    
                version = version & " Revision " & _
                    frmChat.SControl.Modules(I).CodeObject.Script("Revision")
                    
                author = _
                    frmChat.SControl.Modules(I).CodeObject.Script("Author")
                    
                cmdRet(0) = Name & " v" & version & _
                    IIf(LenB(author) > 0, " by " & author, "")
            
                Exit Function
            End If
        Next I
    End If
    
    cmdRet(0) = "Error: Could not find specified script."
    
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

Private Function Expand(ByVal s As String) As String
    Dim I As Integer
    Dim temp As String
    
    If Len(s) > 1 Then
        For I = 1 To Len(s)
            temp = temp & Mid(s, I, 1) & Space(1)
        Next I
        Expand = Trim(temp)
    Else
        Expand = s
    End If
End Function

Private Sub AddQ(ByVal s As String, Optional msg_priority As Integer = -1, Optional ByVal user As String = _
    vbNullString, Optional ByVal Tag As String = vbNullString)
    
    Call frmChat.AddQ(s, msg_priority, user, Tag)
End Sub

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

Private Function searchDatabase(ByRef arrReturn() As String, Optional user As String = vbNullString, _
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
        
        If (gAcc.Access > 0) Then
            If (gAcc.Flags <> vbNullString) Then
                tmpbuf = "Found user " & gAcc.Username & ", with access " & gAcc.Access & _
                    " and flags " & gAcc.Flags & "."
            Else
                tmpbuf = "Found user " & gAcc.Username & ", with access " & gAcc.Access & "."
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
                    If ((DB(I).Access >= lowerBound) And _
                        (DB(I).Access <= upperBound)) Then
                        
                        ' ...
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    
                    blnChecked = True
                ElseIf (lowerBound >= 0) Then
                    If (DB(I).Access = lowerBound) Then
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
                        IIf(DB(I).Access > 0, "\" & DB(I).Access, vbNullString) & _
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
    frmChat.AddChat vbRed, "Error: " & Err.description & " in searchDatabase()."
    
    Exit Function
End Function

Public Function RemoveItem(ByVal rItem As String, file As String, Optional ByVal dbType As String = _
    vbNullString) As String
    
    Dim s()        As String
    Dim f          As Integer
    Dim Counter    As Integer
    Dim strCompare As String
    Dim strAdd     As String
    
    f = FreeFile
    
    If Dir$(GetFilePath(file & ".txt")) = vbNullString Then
        RemoveItem = "No %msgex% file found. Create one using .add, .addtag, or .shitlist."
        Exit Function
    End If
    
    Open (GetFilePath(file & ".txt")) For Input As #f
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
    
    Open (GetFilePath(file & ".txt")) For Output As #f
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
                .Access = 0
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
        ElseIf (gAcc.Access >= 20) Then
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
'            Print #n, DB(I).Username & Space(1) & DB(I).Access & Space(1) & DB(I).Flags
'        Next I
'    Close #n
'End Sub

' requires public
Public Sub LoadDatabase()
    On Error Resume Next

    Dim gA    As udtDatabase
    
    Dim s     As String
    Dim X()   As String
    Dim Path  As String
    Dim I     As Integer
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
                        ReDim Preserve DB(I)
                        
                        With DB(I)
                            .Username = X(0)
                            
                            If StrictIsNumeric(X(1)) Then
                                .Access = Val(X(1))
                            Else
                                If X(1) <> "%" Then
                                    .Flags = X(1)
                                    
                                    'If InStr(X(1), "S") > 0 Then
                                    '    AddToSafelist .Name
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
                            
                            If .Access > 200 Then
                                .Access = 200
                            End If
                            
                            If .Type = "" Or .Type = "%" Then
                                .Type = "USER"
                            End If
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

Public Function IsCorrectSyntax(ByVal CommandName As String, ByVal CommandArgs As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim Command As clsCommandDocObj
    Dim regex   As RegExp
    Dim matches As MatchCollection
    
    ' ...
    Set Command = OpenCommand(CommandName)

    ' ...
    If (Command.Name = vbNullString) Then
        Exit Function
    End If
    
    If (Command.Parameters.Count) Then
        Dim Parameter   As clsCommandParamsObj
        Dim Restriction As clsCommandRestrictionObj
        Dim Splt()      As String
        Dim loopCount   As Integer
        Dim bln         As Boolean
        Dim I           As Integer
        Dim spaceIndex  As Integer
        
        ' ...
        spaceIndex = InStr(1, CommandArgs, Space$(1), vbBinaryCompare)
        
        If ((spaceIndex <> 0) And (Command.Parameters.Count > 1)) Then
            Splt() = Split(CommandArgs, Space$(1), Command.Parameters.Count)
        Else
            If (CommandArgs = vbNullString) Then
                IsCorrectSyntax = False
                
                Exit Function
            End If
        
            ReDim Preserve Splt(0)
        
            Splt(0) = CommandArgs
        End If
        
        For I = 1 To Command.Parameters.Count
            Set Parameter = Command.Parameters(I)

            If (Parameter.IsOptional) Then
                If (Command.Parameters.Count > I) Then
                    If (Command.Parameters.Item(I + 1).IsOptional) Then
                        If (Parameter.dataType = "number") Then
                            If (StrictIsNumeric(Splt(loopCount)) = False) Then
                                bln = True
                            End If
                        Else
                            IsCorrectSyntax = False
                            
                            Exit Function
                        End If
                    End If
                End If
            End If
        
            If (bln = False) Then
                ' ...
                If (Parameter.dataType = "number") Then
                    Dim lVal As Long

                    If (StrictIsNumeric(Splt(loopCount)) = False) Then
                        IsCorrectSyntax = False
                        
                        Exit Function
                    End If

                    'lVal = val(splt(loopCount))
                    
                    'If ((lVal < parameter.min) Or (lVal > parameter.max)) Then
                    '    IsCorrectSyntax = False
                    '
                    '    Exit Function
                    'End If
                    
                ElseIf (Parameter.dataType = "string") Then
                    Set regex = New RegExp
                
                    With regex
                        .pattern = Parameter.pattern
                        .Global = True
                    End With
                    
                    Set matches = regex.Execute(Splt(loopCount))
                    
                    If (matches.Count = 0) Then
                        IsCorrectSyntax = False
                        
                        Exit Function
                    Else
                        If (matches.Item(0).Value <> Splt(loopCount)) Then
                            IsCorrectSyntax = False
                            
                            Exit Function
                        End If
                    End If
                    
                    Set regex = Nothing
                End If
                
                ' ...
                If (Parameter.IsOptional) Then
                    Exit For
                End If
            
                ' ...
                loopCount = (loopCount + 1)
            End If
            
            ' ...
            bln = False
        Next I
    End If
    
    IsCorrectSyntax = True
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in IsCorrectSyntax().")
    
    Exit Function
End Function

Public Function HasAccess(ByVal Username As String, ByVal CommandName As String, Optional ByVal CommandArgs As _
    String = vbNullString, Optional ByRef outbuf As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim Command     As clsCommandDocObj
    Dim user        As clsDBEntryObj
    Dim regex       As RegExp
    Dim matches     As MatchCollection
    Dim FailedCheck As Boolean
    
    ' ...
    Set Command = OpenCommand(CommandName)

    ' ...
    If (Command.Name = vbNullString) Then
        Exit Function
    End If
    
    ' console-only access
    If ((Command.RequiredRank = -1) And _
            (Command.RequiredFlags = vbNullString)) Then
    
        HasAccess = False
    
        Exit Function
    End If
    
    ' ...
    Set user = SharedScriptSupport.GetDBEntry(Username, , , "USER")

    ' ...
    If ((user.Rank >= Command.RequiredRank) = False) Then
        ' ...
        If (user.HasAnyFlag(Command.RequiredFlags) = False) Then
            HasAccess = False
            
            Exit Function
        End If
    End If
    
    ' ...
    If (Command.Parameters.Count) Then
        Dim Parameter   As clsCommandParamsObj
        Dim Restriction As clsCommandRestrictionObj
        Dim Splt()      As String
        Dim loopCount   As Integer
        Dim bln         As Boolean
        Dim I           As Integer
        
        If (InStr(1, CommandArgs, Space$(1), vbBinaryCompare) <> 0) Then
            Splt() = Split(CommandArgs, Space$(1))
        Else
            ReDim Preserve Splt(0)
            
            Splt(0) = CommandArgs
        End If
        
        For I = 1 To Command.Parameters.Count
            ' ...
            If (loopCount > UBound(Splt)) Then
                Exit For
            End If
        
            ' ...
            Set Parameter = Command.Parameters(I)
            
            'frmChat.AddChat vbRed, Parameter.dataType
            'frmChat.AddChat vbRed, StrictIsNumeric(splt(loopCount))
            
            ' ...
            If (Parameter.IsOptional) Then
                ' ...
                If (Parameter.dataType = "number") Then
                    ' ...
                    If (StrictIsNumeric(Splt(loopCount)) = False) Then
                        bln = True
                    End If
                End If
            End If
        
            If (bln = False) Then
                If (Parameter.Restrictions.Count) Then
                    Set regex = New RegExp
                
                    For Each Restriction In Parameter.Restrictions
                        With regex
                            .pattern = Restriction.MatchMessage
                            .Global = True
                        End With

                        Set matches = regex.Execute(Splt(loopCount))

                        If (matches.Count > 0) Then
                            If ((Restriction.RequiredRank = -1) And _
                                    (Restriction.RequiredFlags = vbNullString)) Then
                                    
                                ' ...
                                FailedCheck = True
                            Else
                                If ((user.Rank >= Restriction.RequiredRank) = False) Then
                                    If (user.HasAnyFlag(Restriction.RequiredFlags) = False) Then
                                        ' ...
                                        FailedCheck = True
                                    End If
                                End If
                            End If
                            
                            If (FailedCheck) Then
                                AddQ "Error: You do not have sufficient access to perform the specified " & _
                                    "action."
                                
                                Exit Function
                            End If
                        End If
                    Next
                    
                    Set regex = Nothing
                End If
                
                ' ...
                loopCount = (loopCount + 1)
            End If
            
            ' ...
            bln = False
            FailedCheck = False
        Next I
    End If
    
    HasAccess = True
    
    Exit Function

ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.description & " in HasAccess()."
    
    Exit Function
End Function

Private Function ValidateAccess(ByRef gAcc As udtGetAccessResponse, ByVal CWord As String, _
    Optional ByVal ARGUMENT As String = vbNullString, Optional ByVal restrictionName As String = _
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
            If (StrComp(Command.Attributes.getNamedItem("name").text, _
                CWord, vbTextCompare) = 0) Then
                
                ' ...
                Set accessGroup = Command.selectSingleNode("access")
                
                ' ...
                For Each Access In accessGroup.childNodes
                    If (LCase$(Access.nodeName) = "rank") Then
                        If ((gAcc.Access) >= (Val(Access.text))) Then
                            ValidateAccess = True
                        
                            Exit For
                        End If
                    '// 09/03/2008 JSM - Modified code to use the <flags> element
                    ElseIf (LCase$(Access.nodeName = "flags")) Then
                        For Each flag In Access.childNodes
                            If (InStr(1, gAcc.Flags, flag.text, vbBinaryCompare) <> 0) Then
                                ValidateAccess = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                
                ' ...
                If (ARGUMENT <> vbNullString) Then
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
                        If (StrComp(Restriction.Attributes.getNamedItem("name").text, _
                            restrictionName, vbTextCompare) = 0) Then
                            
                            ' ...
                            Set accessGroup = Restriction.selectSingleNode("access")
                            
                            ' ...
                            For Each Access In accessGroup.childNodes
                                If (LCase$(Access.nodeName) = "rank") Then
                                    If ((gAcc.Access) >= (Val(Access.text))) Then
                                        ValidateAccess = True
                                    
                                        Exit For
                                    End If
                                '// 09/03/2008 JSM - Modified code to use the <flags> element
                                ElseIf (LCase$(Access.nodeName = "flags")) Then
                                    For Each flag In Access.childNodes
                                        If (InStr(1, gAcc.Flags, flag.text, vbBinaryCompare) <> 0) Then
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

Public Function LoadQuotes(Optional strPath As String = vbNullString)
    Dim f As Integer
    Dim s As String

    Set g_Quotes = New Collection
    
    If (strPath = vbNullString) Then
        strPath = App.Path & "\quotes.txt"
    End If
    
    If (LenB(Dir$(strPath)) > 0) Then
        f = FreeFile
        
        Open (GetFilePath("quotes.txt")) For Input As #f
            If (LOF(f) > 1) Then
                Do
                    Line Input #f, s
                    
                    s = Trim(s)
                    
                    If (LenB(s) > 0) Then
                        g_Quotes.Add s
                    End If
                Loop Until EOF(f)
            End If
        Close #f
    End If
End Function

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
                Print #f, " " & DB(I).Access;
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

Private Function GetDBDetail(ByVal Username As String) As String
    Dim sRetAdd As String, sRetMod As String
    Dim I As Integer
    
    For I = 0 To UBound(DB)
        With DB(I)
            If (StrComp(Username, .Username, vbTextCompare) = 0) Then
                If .AddedBy <> "%" And LenB(.AddedBy) > 0 Then
                    sRetAdd = .Username & " was added by " & .AddedBy & " on " & _
                        .AddedOn & "."
                End If
                
                If ((.ModifiedBy <> "%") And (LenB(.ModifiedBy) > 0)) Then
                    If ((.AddedOn <> .ModifiedOn) Or (.AddedBy <> .ModifiedBy)) Then
                        sRetMod = " The entry was last modified by " & .ModifiedBy & _
                            " on " & .ModifiedOn & "."
                    Else
                        sRetMod = " The entry has not been modified since it was added."
                    End If
                End If
                
                If ((LenB(sRetAdd) > 0) Or (LenB(sRetMod) > 0)) Then
                    If (LenB(sRetAdd) > 0) Then
                        GetDBDetail = sRetAdd & sRetMod
                    Else
                        'no add, but we could have a modify
                        GetDBDetail = sRetMod
                    End If
                Else
                    GetDBDetail = "No detailed information is available for that user."
                End If
                
                Exit Function
            End If
        End With
    Next I
    
    GetDBDetail = "That user was not found in the database."
End Function

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

Public Function SecondsToString(ByVal seconds As Long) As String
    Dim temp  As Long ' ...
    Dim secs  As Long ' ...
    Dim mins  As Long ' ...
    Dim hours As Long ' ...
    
    ' ...
    temp = seconds
    
    ' ...
    Do While (temp > 0)
        If (temp - 3600 >= 0) Then
            ' ...
            temp = (temp - 3600)
            
            ' ...
            hours = (hours + 1)
        ElseIf (temp - 60 >= 0) Then
            ' ...
            temp = (temp - 60)
            
            ' ...
            mins = (mins + 1)
        Else
            ' ...
            secs = temp
                   
            ' ...
            temp = 0
        End If
    Loop
    
    ' ...
    SecondsToString = IIf(hours, Right$("00" & hours, 2) & ":", vbNullString) & _
        Right$("00" & mins, 2) & ":" & Right$("00" & secs, 2)
End Function
