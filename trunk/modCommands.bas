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
    
    ' replace message variables
    Message = Replace(Message, "%me", IIf(IsLocal, GetCurrentUsername, Username), 1, -1, vbTextCompare)
    
    If ((IsLocal) And (Left$(Message, 3) = "///")) Then
        frmChat.AddQ Mid$(Message, 3)
        Exit Function
    End If

    ' 08/17/2009 - 52 - using static class method
    Set commands = clsCommandObj.IsCommand(Message, IIf(IsLocal, modGlobals.CurrentUsername, Username), Chr$(0))

    For Each Command In commands
        Command.WasWhispered = WasWhispered
        
        If (Command.HasAccess) Then
            If (IsLocal) Then
                With dbAccess
                    .Rank = 201
                    .Flags = "A"
                End With
            Else
                dbAccess = GetCumulativeAccess(Username)
            End If
            
            'this is eww but i'll change it later
            LogCommand IIf(Command.IsLocal, vbNullString, Username), Command.IsLocal & Space(1) & Command.Args
            
            If (LenB(Command.docs.Owner) = 0) Then 'Is it a built in command?
                Call DispatchCommand(Command)
                Call RunInSingle(Nothing, "Event_Command", Command)
            Else
                Call RunInSingle(modScripting.GetModuleByName(Command.docs.Owner), "Event_Command", Command)
            End If
        End If
        If (DisplayOutput) Then Command.SendResponse
    Next
        
    If (IsLocal) Then
        If (commands.Count = 0) Then frmChat.AddQ Message
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
        Case "bancount":      Call modCommandsInfo.OnBanCount(Command)
        Case "banlistcount":  Call modCommandsInfo.OnBanListCount(Command)
        Case "banned":        Call modCommandsInfo.OnBanned(Command)
        Case "clientbans":    Call modCommandsInfo.OnClientBans(Command)
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
        Case "phrases":       Call modCommandsInfo.OnPhrases(Command)
        Case "ping":          Call modCommandsInfo.OnPing(Command)
        Case "pingme":        Call modCommandsInfo.OnPingMe(Command)
        Case "profile":       Call modCommandsInfo.OnProfile(Command)
        Case "safecheck":     Call modCommandsInfo.OnSafeCheck(Command)
        Case "safelist":      Call modCommandsInfo.OnSafeList(Command)
        Case "scriptdetail":  Call modCommandsInfo.OnScriptDetail(Command)
        Case "scripts":       Call modCommandsInfo.OnScripts(Command)
        Case "server":        Call modCommandsInfo.OnServer(Command)
        Case "shitcheck":     Call modCommandsInfo.OnShitCheck(Command)
        Case "shitlist":      Call modCommandsInfo.OnShitList(Command)
        Case "tagbans":       Call modCommandsInfo.OnTagBans(Command)
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
        Case "rem":           Call modCommandsAdmin.OnRem(Command)
        Case "setexpkey":     Call modCommandsAdmin.OnSetExpKey(Command)
        Case "sethome":       Call modCommandsAdmin.OnSetHome(Command)
        Case "setkey":        Call modCommandsAdmin.OnSetKey(Command)
        Case "setname":       Call modCommandsAdmin.OnSetName(Command)
        Case "setpass":       Call modCommandsAdmin.OnSetPass(Command)
        Case "setpmsg":       Call modCommandsAdmin.OnSetPMsg(Command)
        Case "setserver":     Call modCommandsAdmin.OnSetServer(Command)
        Case "settrigger":    Call modCommandsAdmin.OnSetTrigger(Command)
        Case "whispercmds":   Call modCommandsAdmin.OnWhisperCmds(Command)
        
        'Ops Commands
        Case "addphrase":     Call modCommandsOps.OnAddPhrase(Command)
        Case "ban":           Call modCommandsOps.OnBan(Command)
        Case "cadd":          Call modCommandsOps.OnCAdd(Command)
        Case "cdel":          Call modCommandsOps.OnCDel(Command)
        Case "chpw":          Call modCommandsOps.OnChPw(Command)
        Case "clearbanlist":  Call modCommandsOps.OnClearBanList(Command)
        Case "d2levelban":    Call modCommandsOps.OnD2LevelBan(Command)
        Case "des":           Call modCommandsOps.OnDes(Command)
        Case "delphrase":     Call modCommandsOps.OnDelPhrase(Command)
        Case "exile":         Call modCommandsOps.OnExile(Command)
        Case "giveup":        Call modCommandsOps.OnGiveUp(Command)
        Case "idlebans":      Call modCommandsOps.OnIdleBans(Command)
        Case "ipban":         Call modCommandsOps.OnIPBan(Command)
        Case "ipbans":        Call modCommandsOps.OnIPBans(Command)
        Case "kick":          Call modCommandsOps.OnKick(Command)
        Case "kickonyell":    Call modCommandsOps.OnKickOnYell(Command)
        Case "levelban":      Call modCommandsOps.OnLevelBan(Command)
        Case "peonban":       Call modCommandsOps.OnPeonBan(Command)
        Case "phrasebans":    Call modCommandsOps.OnPhraseBans(Command)
        Case "plugban":       Call modCommandsOps.OnPlugBan(Command)
        Case "poff":          Call modCommandsOps.OnPOff(Command)
        Case "pon":           Call modCommandsOps.OnPOn(Command)
        Case "protect":       Call modCommandsOps.OnProtect(Command)
        Case "pstatus":       Call modCommandsOps.OnPStatus(Command)
        Case "quiettime":     Call modCommandsOps.OnQuietTime(Command)
        Case "resign":        Call modCommandsOps.OnResign(Command)
        Case "safeadd":       Call modCommandsOps.OnSafeAdd(Command)
        Case "safedel":       Call modCommandsOps.OnSafeDel(Command)
        Case "shitadd":       Call modCommandsOps.OnShitAdd(Command)
        Case "shitdel":       Call modCommandsOps.OnShitDel(Command)
        Case "sweepban":      Call modCommandsOps.OnSweepBan(Command)
        Case "sweepignore":   Call modCommandsOps.OnSweepIgnore(Command)
        Case "tagadd":        Call modCommandsOps.OnTagAdd(Command)
        Case "tagdel":        Call modCommandsOps.OnTagDel(Command)
        Case "unban":         Call modCommandsOps.OnUnBan(Command)
        Case "unexile":       Call modCommandsOps.OnUnExile(Command)
        Case "unipban":       Call modCommandsOps.OnUnIPBan(Command)
        Case "voteban":       Call modCommandsOps.OnVoteBan(Command)
        Case "votekick":      Call modCommandsOps.OnVoteKick(Command)
        
        'Misc Commands
        Case "addquote":      Call modCommandsMisc.OnAddQuote(Command)
        Case "bmail":         Call modCommandsMisc.OnBMail(Command)
        Case "cancel":        Call modCommandsMisc.OnCancel(Command)
        Case "checkmail":     Call modCommandsMisc.OnCheckMail(Command)
        Case "exec":          Call modCommandsMisc.OnExec(Command)
        Case "flip":          Call modCommandsMisc.OnFlip(Command)
        Case "greet":         Call modCommandsMisc.OnGreet(Command)
        Case "idle":          Call modCommandsMisc.OnIdle(Command)
        Case "idletime":      Call modCommandsMisc.OnIdleTime(Command)
        Case "idletype":      Call modCommandsMisc.OnIdleType(Command)
        Case "inbox":         Call modCommandsMisc.OnInbox(Command)
        Case "math":          Call modCommandsMisc.OnMath(Command)
        Case "mmail":         Call modCommandsMisc.OnMMail(Command)
        Case "quote":         Call modCommandsMisc.OnQuote(Command)
        Case "roll":          Call modCommandsMisc.OnRoll(Command)
        Case "readfile":      Call modCommandsMisc.OnReadFile(Command)
        Case "setidle":       Call modCommandsMisc.OnSetIdle(Command)
        Case "tally":         Call modCommandsMisc.OnTally(Command)
        Case "vote":          Call modCommandsMisc.OnVote(Command)
        
        Case Else: DispatchCommand = False
    End Select
End Function

Public Function OnRemOld(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim sUsername  As String  ' ...
    Dim tmpbuf     As String  ' temporary output buffer
    Dim user       As udtGetAccessResponse
    Dim dbType     As String  ' ...
    Dim Index      As Long    ' ...
    Dim params     As String  ' ...
    Dim strArray() As String  ' ...
    Dim I          As Integer ' ...

    ' check for presence of optional add command parameters
    Index = InStr(1, msgData, " --", vbBinaryCompare)

    ' did we find such parameters?
    If (Index > 0) Then
        params = Mid$(msgData, Index - 1) ' grab parameters
        msgData = Mid$(msgData, 1, Index) ' remove paramaters from message
    End If
    
    ' do we have any special paramaters?
    If (Len(params) > 0) Then
        strArray() = Split(params, " --") ' split message by paramter
        
        ' loop through paramter list
        For I = 1 To UBound(strArray)
            Dim Parameter As String ' ...
            Dim pmsg      As String ' ...
            
            Index = InStr(1, strArray(I), Space(1), vbBinaryCompare)
            If (Index > 0) Then
                Parameter = Mid$(strArray(I), 1, Index - 1)
                pmsg = Mid$(strArray(I), Index + 1)
            Else
                Parameter = strArray(I)
            End If
            
            Parameter = LCase$(Parameter)
            
            ' handle parameters
            Select Case (Parameter)
                Case "type"
                    If (Len(pmsg) > 0) Then
                        Select Case UCase$(pmsg)
                            Case "USER":
                            Case "GROUP":
                            Case "CLAN":
                            Case "GAME":
                            Case Else: pmsg = "USER"
                        End Select
                        dbType = UCase$(pmsg)
                    End If
            End Select
        Next I
    End If

    sUsername = msgData
    user = GetAccess(sUsername, dbType)
    
    If (Len(sUsername) > 0) Then
        If ((GetAccess(sUsername, dbType).Rank = -1) And (LenB(GetAccess(sUsername, dbType).Flags) = 0)) Then
            tmpbuf = "User not found."
        ElseIf (GetAccess(sUsername, dbType).Rank >= dbAccess.Rank) Then
            tmpbuf = "That user has higher or equal access."
        ElseIf ((Not InStr(1, GetAccess(sUsername, dbType).Flags, "L") = 0) And _
                (Not (InBot)) And _
                (InStr(1, GetAccess(Username, dbType).Flags, "A") = 0) And _
                (GetAccess(Username, dbType).Rank <= 99)) Then
            
                tmpbuf = "Error: That user is Locked."
        Else
            Dim res As Boolean ' ...
        
            res = DB_remove(sUsername, dbType)
            
            If (res) Then
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(RemEntry, IIf(InBot, "console", Username), sUsername, dbType)
                End If
                tmpbuf = StringFormat("Successfully removed database entry {0}{1}{0}.", Chr$(34), sUsername)
            Else
                tmpbuf = "Error: There was a problem removing that entry from the database."
            End If
        End If
    End If
    cmdRet(0) = tmpbuf
End Function

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
                Call OnRemOld(Username, dbAccess, user, InBot, cmdRet)
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
                Dim currentCharacter As String
            
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
                                    Call modCommandsOps.WildCardBan(user, vbNullString, 2)
                                Else
                                    ' ...
                                    If (g_Channel.IsOnBanList(user)) Then
                                        frmChat.AddQ ("/unban " & user)
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
    Dim s      As String
    Dim X()    As String
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
                    X() = Split(s, " ", 10)
                    
                    If UBound(X) > 0 Then
                        ReDim Preserve DB(I)
                        
                        With DB(I)
                            .Username = X(0)
                            
                            .Rank = 0
                            .AddedOn = Now
                            .AddedBy = "2.6r3Import"
                            .BanMessage = vbNullString
                            .Flags = vbNullString
                            .Groups = vbNullString
                            .ModifiedBy = "2.6r3Import"
                            .ModifiedOn = Now
                            .Type = "USER"
                            
                            If StrictIsNumeric(X(1)) Then
                                .Rank = Val(X(1))
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
                                    .Rank = Int(X(2))
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
'
''08/15/09 - Hdx - Converted to use clsCommandObj to check if the command has valid syntax
''Removed outBuff ... What use was it?
'Public Function IsCorrectSyntax(ByVal commandName As String, ByVal commandArgs As String, Optional scriptOwner As String = vbNullString) As Boolean
'
'    On Error GoTo ERROR_HANDLER
'
'    Dim docs As clsCommandDocObj
'    Dim Command As New clsCommandObj
'
'    Set docs = OpenCommand(commandName, scriptOwner)
'    If (docs Is Nothing) Then
'        Set docs = OpenCommand(convertAlias(commandName), scriptOwner)
'        If (docs Is Nothing) Then
'            IsCorrectSyntax = False
'            Exit Function
'        End If
'    End If
'
'    With Command
'        .Name = docs.Name
'        .Args = commandArgs
'        IsCorrectSyntax = .IsValid
'    End With
'
'    Set docs = Nothing
'    Set Command = Nothing
'    Exit Function
'
'ERROR_HANDLER:
'    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in IsCorrectSyntax().")
'End Function

'08/15/09 - Hdx - Converted to use clsCommandObj to check if the user has enough access
'Removed outBuff ... What use was it?
'Public Function HasAccess(ByVal Username As String, ByVal commandName As String, Optional ByVal commandArgs As _
'    String = vbNullString, Optional scriptOwner As String = vbNullString) As Boolean
'
'    On Error GoTo ERROR_HANDLER
'    Dim docs As clsCommandDocObj
'    Dim Command As clsCommandObj
'
'    Set docs = OpenCommand(commandName, scriptOwner)
'    If (docs Is Nothing) Then
'        Set docs = OpenCommand(convertAlias(commandName), scriptOwner)
'        If (docs Is Nothing) Then
'            HasAccess = False
'            Exit Function
'        End If
'    End If
'
'    With Command
'        .Name = docs.Name
'        .Args = commandArgs
'        .Username = Username
'        HasAccess = .HasAccess
'    End With
'
'    Set docs = Nothing
'    Set Command = Nothing
'    Exit Function
'
'ERROR_HANDLER:
'    frmChat.AddChat vbRed, "Error: " & Err.description & " in HasAccess()."
'End Function

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


Private Function CheckUser(ByVal user As String, Optional ByVal allow_illegal As Boolean = False) As Boolean
    
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
                                Exit For
                        End Select
                    End If
                End If
            End If
        Next I
    End If
    
    ' is our user valid?
    If (Not (invalid)) Then
        If (illegal) Then ' does our user contain illegal characters?
            CheckUser = allow_illegal ' do we allow illegal characters?
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
        If ((StrReverse$(BotVars.Product) = "D2DV") Or (StrReverse$(BotVars.Product) = "D2XP")) Then
            If ((BotVars.UseGameConventions = False) Or ((BotVars.UseD2GameConventions = False))) Then
               
                Index = InStr(1, Username, "*", vbBinaryCompare)
            
                If (Index > 0) Then
                    convertUsername = Mid$(Username, Index + 1)
                Else
                    convertUsername = Username
                End If
            Else
                Index = InStr(1, Username, "*", vbBinaryCompare)
            
                If (Index > 1) Then
                    convertUsername = StringFormat("{0} (*{1})", Left$(Username, Index - 1), Mid$(Username, Index + 1))
                Else
                    If (Index = 0) Then
                        convertUsername = "*" & Username
                    End If
                End If
            End If
            
        ElseIf ((StrReverse$(BotVars.Product) = "WAR3") Or (StrReverse$(BotVars.Product) = "W3XP")) Then
            If ((BotVars.UseGameConventions = False)) Or ((BotVars.UseW3GameConventions = False)) Then
    
                If (LenB(BotVars.Gateway) > 0) Then
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

    If ((StrReverse$(BotVars.Product) = "D2DV") Or (StrReverse$(BotVars.Product) = "D2XP")) Then
        If ((BotVars.UseGameConventions = False)) Or ((BotVars.UseD2GameConventions = False)) Then
            
            ' With reverseUsername() now being called from AddQ(), usernames
            ' in procedures called prior to AddQ() will no longer require
            ' prefixes; however, we want to ensure that a '*' was not already
            ' specified before conversion to allow older scripts and procedures
            ' to continue functioning correctly.  This check may be removed in
            ' future releases.
            If (Not Left$(Username, 1) = "*") Then
                reverseUsername = ("*" & Username)
            End If
        End If
    ElseIf ((StrReverse$(BotVars.Product) = "WAR3") Or (StrReverse$(BotVars.Product) = "W3XP")) Then
        If ((BotVars.UseGameConventions = False)) Or ((BotVars.UseW3GameConventions = False)) Then
            
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
