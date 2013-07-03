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

Public flood    As String
Public floodCap As Byte

' prepares commands for processing, and calls helper functions associated with processing
Public Function ProcessCommand(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As _
        Boolean = False, Optional ByVal WasWhispered As Boolean = False, Optional DisplayOutput As Boolean = _
                True) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Dim commands         As Collection
    Dim Command          As clsCommandObj
    
    ' replace message variables
    Message = Replace(Message, "%me", IIf(IsLocal, GetCurrentUsername, Username), 1, -1, vbTextCompare)
    
    If ((IsLocal) And (Left$(Message, 3) = "///")) Then
        frmChat.AddQ Mid$(Message, 3)
        Exit Function
    End If

    Set commands = clsCommandObj.IsCommand(Message, IIf(IsLocal, modGlobals.CurrentUsername, CleanUsername(Username)), _
            IsLocal, WasWhispered, Chr$(0))

    For Each Command In commands
        If (Command.HasAccess) Then
            'this is eww but i'll change it later
            LogCommand IIf(Command.IsLocal, vbNullString, Username), Command.IsLocal & Space(1) & Command.Args
            
            If (LenB(Command.docs.Owner) = 0) Then
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
        Case "about":          Call modCommandsInfo.OnAbout(Command)
        Case "accountinfo":    Call modCommandsInfo.OnAccountInfo(Command)
        Case "bancount":       Call modCommandsInfo.OnBanCount(Command)
        Case "banlistcount":   Call modCommandsInfo.OnBanListCount(Command)
        Case "banned":         Call modCommandsInfo.OnBanned(Command)
        Case "clientbans":     Call modCommandsInfo.OnClientBans(Command)
        Case "detail":         Call modCommandsInfo.OnDetail(Command)
        Case "find":           Call modCommandsInfo.OnFind(Command)
        Case "findattr":       Call modCommandsInfo.OnFindAttr(Command)
        Case "findgrp":        Call modCommandsInfo.OnFindGrp(Command)
        Case "help":           Call modCommandsInfo.OnHelp(Command)
        Case "helpattr":       Call modCommandsInfo.OnHelpAttr(Command)
        Case "helprank":       Call modCommandsInfo.OnHelpRank(Command)
        Case "info":           Call modCommandsInfo.OnInfo(Command)
        Case "initperf":       Call modCommandsInfo.OnInitPerf(Command)
        Case "lastseen":       Call modCommandsInfo.OnLastSeen(Command)
        Case "lastwhisper":    Call modCommandsInfo.OnLastWhisper(Command)
        Case "localip":        Call modCommandsInfo.OnLocalIp(Command)
        Case "owner":          Call modCommandsInfo.OnOwner(Command)
        Case "phrases":        Call modCommandsInfo.OnPhrases(Command)
        Case "ping":           Call modCommandsInfo.OnPing(Command)
        Case "pingme":         Call modCommandsInfo.OnPingMe(Command)
        Case "profile":        Call modCommandsInfo.OnProfile(Command)
        Case "safecheck":      Call modCommandsInfo.OnSafeCheck(Command)
        Case "safelist":       Call modCommandsInfo.OnSafeList(Command)
        Case "scriptdetail":   Call modCommandsInfo.OnScriptDetail(Command)
        Case "scripts":        Call modCommandsInfo.OnScripts(Command)
        Case "server":         Call modCommandsInfo.OnServer(Command)
        Case "shitcheck":      Call modCommandsInfo.OnShitCheck(Command)
        Case "shitlist":       Call modCommandsInfo.OnShitList(Command)
        Case "tagbans":        Call modCommandsInfo.OnTagBans(Command)
        Case "time":           Call modCommandsInfo.OnTime(Command)
        Case "trigger":        Call modCommandsInfo.OnTrigger(Command)
        Case "uptime":         Call modCommandsInfo.OnUptime(Command)
        Case "where":          Call modCommandsInfo.OnWhere(Command)
        Case "whoami":         Call modCommandsInfo.OnWhoAmI(Command)
        Case "whois":          Call modCommandsInfo.OnWhoIs(Command)
        
        'Clan related commands
        Case "clan":           Call modCommandsClan.OnClan(Command)
        Case "demote":         Call modCommandsClan.OnDemote(Command)
        Case "disbandclan":    Call modCommandsClan.OnDisbandClan(Command)
        Case "invite":         Call modCommandsClan.OnInvite(Command)
        Case "makechieftain":  Call modCommandsClan.OnMakeChieftain(Command)
        Case "motd":           Call modCommandsClan.OnMOTD(Command)
        Case "promote":        Call modCommandsClan.OnPromote(Command)
        Case "setmotd":        Call modCommandsClan.OnSetMOTD(Command)
        
        'Media Player comands
        Case "allowmp3":       Call modCommandsMP3.OnAllowMp3(Command)
        Case "fos":            Call modCommandsMP3.OnFOS(Command)
        Case "loadwinamp":     Call modCommandsMP3.OnLoadWinamp(Command)
        Case "mp3":            Call modCommandsMP3.OnMP3(Command)
        Case "next":           Call modCommandsMP3.OnNext(Command)
        Case "pause":          Call modCommandsMP3.OnPause(Command)
        Case "play":           Call modCommandsMP3.OnPlay(Command)
        Case "previous":       Call modCommandsMP3.OnPrevious(Command)
        Case "repeat":         Call modCommandsMP3.OnRepeat(Command)
        Case "setvol":         Call modCommandsMP3.OnSetVol(Command)
        Case "shuffle":        Call modCommandsMP3.OnShuffle(Command)
        Case "stop":           Call modCommandsMP3.OnStop(Command)
        Case "useitunes":      Call modCommandsMP3.OnUseiTunes(Command)
        Case "usewinamp":      Call modCommandsMP3.OnUseWinamp(Command)
        'Case "usewmp":         Call modCommandsMP3.OnUseWMPlayer(Command)
        
        'Chat Commands
        Case "away":           Call modCommandsChat.OnAway(Command)
        Case "back":           Call modCommandsChat.OnBack(Command)
        Case "block":          Call modCommandsChat.OnBlock(Command)
        Case "connect":        Call modCommandsChat.OnConnect(Command)
        Case "cq":             Call modCommandsChat.OnCQ(Command)
        Case "disconnect":     Call modCommandsChat.OnDisconnect(Command)
        Case "expand":         Call modCommandsChat.OnExpand(Command)
        Case "fadd":           Call modCommandsChat.OnFAdd(Command)
        Case "filter":         Call modCommandsChat.OnFilter(Command)
        Case "forcejoin":      Call modCommandsChat.OnForceJoin(Command)
        Case "frem":           Call modCommandsChat.OnFRem(Command)
        Case "home":           Call modCommandsChat.OnHome(Command)
        Case "ignore":         Call modCommandsChat.OnIgnore(Command)
        Case "igpriv":         Call modCommandsChat.OnIgPriv(Command)
        Case "join":           Call modCommandsChat.OnJoin(Command)
        Case "unblock":        Call modCommandsChat.OnUnBlock(Command)
        Case "unfilter":       Call modCommandsChat.OnUnFilter(Command)
        Case "unignore":       Call modCommandsChat.OnUnIgnore(Command)
        Case "unigpriv":       Call modCommandsChat.OnUnIgPriv(Command)
        Case "quickrejoin":    Call modCommandsChat.OnQuickRejoin(Command)
        Case "reconnect":      Call modCommandsChat.OnReconnect(Command)
        Case "rejoin":         Call modCommandsChat.OnReJoin(Command)
        Case "say":            Call modCommandsChat.OnSay(Command)
        Case "scq":            Call modCommandsChat.OnSCQ(Command)
        Case "shout":          Call modCommandsChat.OnShout(Command)
        Case "watch":          Call modCommandsChat.OnWatch(Command)
        Case "watchoff":       Call modCommandsChat.OnWatchOff(Command)
        
        'Admin Commands
        Case "add":            Call modCommandsAdmin.OnAdd(Command)
        Case "clear":          Call modCommandsAdmin.OnClear(Command)
        Case "disable":        Call modCommandsAdmin.OnDisable(Command)
        Case "dump":           Call modCommandsAdmin.OnDump(Command)
        Case "enable":         Call modCommandsAdmin.OnEnable(Command)
        Case "locktext":       Call modCommandsAdmin.OnLockText(Command)
        Case "quit":           Call modCommandsAdmin.OnQuit(Command)
        Case "rem":            Call modCommandsAdmin.OnRem(Command)
        Case "setcommandline": Call modCommandsAdmin.OnSetCommandLine(Command)
        Case "setexpkey":      Call modCommandsAdmin.OnSetExpKey(Command)
        Case "sethome":        Call modCommandsAdmin.OnSetHome(Command)
        Case "setkey":         Call modCommandsAdmin.OnSetKey(Command)
        Case "setname":        Call modCommandsAdmin.OnSetName(Command)
        Case "setpass":        Call modCommandsAdmin.OnSetPass(Command)
        Case "setpmsg":        Call modCommandsAdmin.OnSetPMsg(Command)
        Case "setserver":      Call modCommandsAdmin.OnSetServer(Command)
        Case "setbnlsserver":  Call modCommandsAdmin.OnSetBnlsServer(Command)
        Case "settrigger":     Call modCommandsAdmin.OnSetTrigger(Command)
        Case "whispercmds":    Call modCommandsAdmin.OnWhisperCmds(Command)
        
        'Ops Commands
        Case "addphrase":      Call modCommandsOps.OnAddPhrase(Command)
        Case "ban":            Call modCommandsOps.OnBan(Command)
        Case "cadd":           Call modCommandsOps.OnCAdd(Command)
        Case "cdel":           Call modCommandsOps.OnCDel(Command)
        Case "chpw":           Call modCommandsOps.OnChPw(Command)
        Case "clearbanlist":   Call modCommandsOps.OnClearBanList(Command)
        Case "d2levelban":     Call modCommandsOps.OnD2LevelBan(Command)
        Case "des":            Call modCommandsOps.OnDes(Command)
        Case "delphrase":      Call modCommandsOps.OnDelPhrase(Command)
        Case "exile":          Call modCommandsOps.OnExile(Command)
        Case "giveup":         Call modCommandsOps.OnGiveUp(Command)
        Case "idlebans":       Call modCommandsOps.OnIdleBans(Command)
        Case "ipban":          Call modCommandsOps.OnIPBan(Command)
        Case "ipbans":         Call modCommandsOps.OnIPBans(Command)
        Case "kick":           Call modCommandsOps.OnKick(Command)
        Case "kickonyell":     Call modCommandsOps.OnKickOnYell(Command)
        Case "levelban":       Call modCommandsOps.OnLevelBan(Command)
        Case "peonban":        Call modCommandsOps.OnPeonBan(Command)
        Case "phrasebans":     Call modCommandsOps.OnPhraseBans(Command)
        Case "plugban":        Call modCommandsOps.OnPlugBan(Command)
        Case "poff":           Call modCommandsOps.OnPOff(Command)
        Case "pon":            Call modCommandsOps.OnPOn(Command)
        Case "protect":        Call modCommandsOps.OnProtect(Command)
        Case "pstatus":        Call modCommandsOps.OnPStatus(Command)
        Case "quiettime":      Call modCommandsOps.OnQuietTime(Command)
        Case "resign":         Call modCommandsOps.OnResign(Command)
        Case "safeadd":        Call modCommandsOps.OnSafeAdd(Command)
        Case "safedel":        Call modCommandsOps.OnSafeDel(Command)
        Case "shitadd":        Call modCommandsOps.OnShitAdd(Command)
        Case "shitdel":        Call modCommandsOps.OnShitDel(Command)
        Case "sweepban":       Call modCommandsOps.OnSweepBan(Command)
        Case "sweepignore":    Call modCommandsOps.OnSweepIgnore(Command)
        Case "tagadd":         Call modCommandsOps.OnTagAdd(Command)
        Case "tagdel":         Call modCommandsOps.OnTagDel(Command)
        Case "unban":          Call modCommandsOps.OnUnBan(Command)
        Case "unexile":        Call modCommandsOps.OnUnExile(Command)
        Case "unipban":        Call modCommandsOps.OnUnIPBan(Command)
        Case "voteban":        Call modCommandsOps.OnVoteBan(Command)
        Case "votekick":       Call modCommandsOps.OnVoteKick(Command)
        
        'Misc Commands
        Case "addquote":       Call modCommandsMisc.OnAddQuote(Command)
        Case "bmail":          Call modCommandsMisc.OnBMail(Command)
        Case "cancel":         Call modCommandsMisc.OnCancel(Command)
        Case "checkmail":      Call modCommandsMisc.OnCheckMail(Command)
        Case "exec":           Call modCommandsMisc.OnExec(Command)
        Case "flip":           Call modCommandsMisc.OnFlip(Command)
        Case "greet":          Call modCommandsMisc.OnGreet(Command)
        Case "idle":           Call modCommandsMisc.OnIdle(Command)
        Case "idletime":       Call modCommandsMisc.OnIdleTime(Command)
        Case "idletype":       Call modCommandsMisc.OnIdleType(Command)
        Case "inbox":          Call modCommandsMisc.OnInbox(Command)
        Case "math":           Call modCommandsMisc.OnMath(Command)
        Case "mmail":          Call modCommandsMisc.OnMMail(Command)
        Case "quote":          Call modCommandsMisc.OnQuote(Command)
        Case "roll":           Call modCommandsMisc.OnRoll(Command)
        Case "readfile":       Call modCommandsMisc.OnReadFile(Command)
        Case "setidle":        Call modCommandsMisc.OnSetIdle(Command)
        Case "tally":          Call modCommandsMisc.OnTally(Command)
        Case "vote":           Call modCommandsMisc.OnVote(Command)
        
        Case Else: DispatchCommand = False
    End Select
End Function

Public Function DBUserToString(ByVal User As String, ByVal dbType As String) As String
    
    Dim TypeStr As String
    
    If (Len(dbType) > 0 And StrComp(dbType, "%") <> 0 And StrComp(dbType, "user", vbTextCompare) <> 0) Then
        TypeStr = StringFormat(" ({0})", LCase$(dbType))
    Else
        TypeStr = vbNullString
    End If
    
    DBUserToString = User & TypeStr
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.description & " in DBUserToString().")
End Function

Public Function OnRemOld(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim sUsername  As String
    Dim tmpbuf     As String  ' temporary output buffer
    Dim User       As udtGetAccessResponse
    Dim dbType     As String
    Dim Index      As Long
    Dim params     As String
    Dim strArray() As String
    Dim i          As Integer

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
        For i = 1 To UBound(strArray)
            Dim parameter As String
            Dim pmsg      As String
            
            Index = InStr(1, strArray(i), Space(1), vbBinaryCompare)
            If (Index > 0) Then
                parameter = Mid$(strArray(i), 1, Index - 1)
                pmsg = Mid$(strArray(i), Index + 1)
            Else
                parameter = strArray(i)
            End If
            
            parameter = LCase$(parameter)
            
            ' handle parameters
            Select Case (parameter)
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
        Next i
    End If

    sUsername = msgData
    User = GetAccess(sUsername, dbType)
    
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
            Dim res As Boolean
            
            dbType = User.Type
        
            res = DB_remove(sUsername, dbType)
            
            If (res) Then
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(RemEntry, IIf(InBot, "console", Username), sUsername, dbType)
                End If
                tmpbuf = StringFormat("{0} has been removed from the database.", DBUserToString(sUsername, dbType))
            Else
                tmpbuf = "Error: There was a problem removing that entry from the database."
            End If
        End If
    End If
    cmdRet(0) = tmpbuf
End Function

Public Function OnAddOld(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim gAcc       As udtGetAccessResponse

    Dim strArray() As String
    Dim i          As Integer
    Dim tmpbuf     As String  ' temporary output buffer
    Dim dbPath     As String
    Dim User       As String
    Dim Rank       As Integer
    Dim Flags      As String
    Dim found      As Boolean
    Dim params     As String
    Dim Index      As Integer
    Dim sGrp       As String
    Dim dbType     As String
    Dim banmsg     As String

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
        User = strArray(0)
        
        If (User = vbNullString) Then
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
        
        If (BotVars.CaseSensitiveFlags = False) Then
            Flags = UCase$(Flags)
        End If
        
        ' do we have any special paramaters?
        If (Len(params)) Then
            ' split message by paramter
            strArray() = Split(params, " --")
            
            ' loop through paramter list
            For i = 1 To UBound(strArray)
                Dim parameter As String
                Dim pmsg      As String
                
                ' check message for a space
                Index = InStr(1, strArray(i), Space(1), vbBinaryCompare)
                
                ' did our search find a space?
                If (Index > 0) Then
                    ' grab parameter
                    parameter = Mid$(strArray(i), 1, Index - 1)
                    
                    ' grab parameter message
                    pmsg = Mid$(strArray(i), Index + 1)
                Else
                    ' grab parameter
                    parameter = strArray(i)
                End If
                
                ' convert parameter to lowercase
                parameter = LCase$(parameter)
                
                ' handle parameters
                Select Case (Trim$(parameter))
                    Case "type"
                        ' do we have a valid parameter Length?
                        If (Len(pmsg)) Then
                            ' grab database entry type
                            dbType = UCase$(pmsg)
                            
                            If (dbType = "USER") Then
                                ' Do nothing
                            ElseIf (dbType = "GROUP") Then
                                ' check for presence of space in name
                                If (InStr(1, User, Space(1), vbBinaryCompare) <> 0) Then
                                    cmdRet(0) = "Error: The specified group name contains one or more " & _
                                        "invalid characters."
                                                                
                                    Exit Function
                                End If
                            ElseIf (dbType = "CLAN") Then
                                ' check for invalid clan entry
                                If ((Len(User) < 2) Or (Len(User) > 4)) Then
                                    ' return message
                                    cmdRet(0) = "Error: The clan name specified is of an " & _
                                        "incorrect Length."
                                        
                                    Exit Function
                                End If
                            ElseIf (dbType = "GAME") Then
                                ' convert entry to uppercase
                                User = UCase$(User)
                                
                                ' check for invalid game entry
                                Select Case (User)
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
                
                    Case "banmsg"
                        ' do we have a valid parameter Length?
                        If (Len(pmsg)) Then
                            banmsg = pmsg
                        End If
                        
                    Case "group"
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
                                Dim tmp As udtGetAccessResponse
                                
                                tmp = GetAccess(Splt(j), "GROUP")
                            
                                If (dbAccess.Rank < tmp.Rank) Then
                                    cmdRet(0) = "Error: You do not have sufficient access to " & _
                                        "add a member to the specified group."
                                        
                                    Exit Function
                                End If
                                
                                If ((StrComp(Splt(j), User, vbTextCompare) = 0) And _
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
                                            If (CheckGroup(tmp.Username, User)) Then
                                                cmdRet(0) = "Error: " & Chr$(34) & tmp.Username & _
                                                    Chr$(34) & " is already a member of group " & _
                                                        Chr$(34) & User & "." & Chr$(34)
                                        
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
        gAcc = GetAccess(User, dbType)
        
        ' if we've found a matching user, lets correct
        ' the casing of the name that we've entered
        If (Len(gAcc.Username) > 0) Then
            If (StrComp(gAcc.Type, dbType, vbTextCompare) = 0) Then
                User = gAcc.Username
            End If
        End If
        
        ' grab access for entry
        gAcc = GetCumulativeAccess(User, dbType)

        ' is rank valid?
        If ((Rank <= 0) And (Flags = vbNullString) And (sGrp = vbNullString)) Then
            
            If ((Rank = 0) And ((gAcc.Rank > 0) Or (gAcc.Flags <> vbNullString) Or _
                (gAcc.Groups <> vbNullString))) Then
                Call OnRemOld(Username, dbAccess, User, InBot, cmdRet)
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
                        If (Len(tmpFlags) > 0) Then
                            ' check for special flags
                            If (InStr(1, tmpFlags, "B", vbBinaryCompare) <> 0) Then
                                If (InStr(1, User, "*", vbBinaryCompare) <> 0) Then
                                    Call modCommandsOps.WildCardBan(User, vbNullString, 2)
                                Else
                                    If (g_Channel.IsOnBanList(User)) Then
                                        frmChat.AddQ ("/unban " & User)
                                    End If
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
                    If ((gAcc.Rank = 0) And (gAcc.Flags = vbNullString) And _
                        ((gAcc.Groups = vbNullString) Or (gAcc.Groups = "%"))) Then
                        
                        Dim res As Boolean
                       
                        ' with no access a database entry is
                        ' pointless, so lets remove it
                        res = DB_remove(User, gAcc.Type)
                        
                        If (res) Then
                            cmdRet(0) = DBUserToString(User, dbType) & " has been removed " & _
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
            dbPath = GetFilePath("Users.txt")

            ' does user already exist in database?
            For i = LBound(DB) To UBound(DB)
                If ((StrComp(DB(i).Username, User, vbTextCompare) = 0) And _
                    (StrComp(DB(i).Type, gAcc.Type, vbTextCompare) = 0)) Then
                    
                    ' modify database entry
                    With DB(i)
                        .Username = User
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
                        Call LogDBAction(ModEntry, IIf(InBot, "console", Username), DB(i).Username, _
                            DB(i).Type, DB(i).Rank, DB(i).Flags, DB(i).Groups)
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
                If (DB(0).Username = vbNullString) Then
                    ReDim Preserve DB(0)
                Else
                    ReDim Preserve DB(UBound(DB) + 1)
                End If

                With DB(UBound(DB))
                    .Username = User
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
                tmpbuf = DBUserToString(User, dbType) & " has been given rank " & _
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
                    tmpbuf = DBUserToString(User, dbType) & " has been given flags " & _
                        gAcc.Flags
                End If
            End If
            
            ' was the user assigned to a group?
            If (Len(sGrp)) Then
                If (Len(tmpbuf)) Then
                    tmpbuf = tmpbuf & ", and has been made a member of " & _
                        "the group(s): " & sGrp
                Else
                    tmpbuf = DBUserToString(User, dbType) & " has been made a member of " & _
                        "the group(s): " & sGrp
                End If
                
            End If
            
            ' terminate sentence
            ' with period
            tmpbuf = tmpbuf & "."
        End If
        
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

    Dim i     As Integer
    Dim found As Boolean
    
    For i = LBound(DB) To UBound(DB)
        If (StrComp(DB(i).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean
        
            If (Len(dbType)) Then
                If (StrComp(DB(i).Type, dbType, vbTextCompare) = 0) Then
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
        Dim bak As udtDatabase
        
        Dim j   As Integer
        
        bak = DB(i)

        ' we aren't removing the last array
        ' element, are we?
        If (UBound(DB) = 0) Then
            ' redefine array size
            ReDim DB(0)
            
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
            For j = i To UBound(DB) - 1
                DB(j) = DB(j + 1)
            Next j
            
            ' redefine array size
            ReDim Preserve DB(UBound(DB) - 1)
    
            ' if we're removing a group, we need to also fix our
            ' group memberships, in case anything is broken now
            If (StrComp(bak.Type, "GROUP", vbBinaryCompare) = 0) Then
                Dim res As Boolean
           
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
                                Dim Splt()     As String
                                Dim innerfound As Boolean
                                
                                Splt() = Split(DB(i).Groups, ",")
                                
                                For j = LBound(Splt) To UBound(Splt)
                                    If (StrComp(bak.Username, Splt(j), vbTextCompare) = 0) Then
                                        innerfound = True
                                    
                                        Exit For
                                    End If
                                Next j
                            
                                If (innerfound) Then
                                    Dim K As Integer
                                    
                                    For K = (j + 1) To UBound(Splt)
                                        Splt(K - 1) = Splt(K)
                                    Next K
                                    
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
        End If
        
        ' commit modifications
        Call WriteDatabase(GetFilePath("Users.txt"))
        
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
    Dim i As Long
    
    If (bFlood = False) Then
        Dim gAcc As udtGetAccessResponse
        
        gAcc = GetCumulativeAccess(Username, "USER")
        
        If (Not InStr(1, gAcc.Flags, "S", vbBinaryCompare) = 0) Then
            GetSafelist = True
        ElseIf (gAcc.Rank >= AutoModSafelistValue) Then
            GetSafelist = True
        End If
    Else
        For i = 0 To (UBound(gFloodSafelist) - 1)
            If PrepareCheck(Username) Like gFloodSafelist(i) Then
                GetSafelist = True
                Exit For
            End If
        Next i
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
            GetShitlist = Username & Space$(1) & gAcc.BanMessage
        ElseIf InStr(1, gAcc.Username, " (clan)", vbBinaryCompare) > 0 Then
            GetShitlist = Username & Space$(1) & "Clanban: " & Mid$(gAcc.Username, 2, InStr(1, gAcc.Username, " (clan)", vbBinaryCompare) - 2)
        ElseIf InStr(1, gAcc.Username, " (game)", vbBinaryCompare) > 0 Then
            GetShitlist = Username & Space$(1) & "Clientban: " & Mid$(gAcc.Username, 2, InStr(1, gAcc.Username, " (game)", vbBinaryCompare) - 2)
        ElseIf InStr(1, gAcc.Username, "*", vbBinaryCompare) > 0 Then
            GetShitlist = Username & Space$(1) & "Tagban: " & Mid$(gAcc.Username, 2, Len(gAcc.Username) - 2)
        Else
            GetShitlist = Username & Space$(1) & "Shitlisted"
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
    Dim x()    As String
    Dim Path   As String
    Dim i      As Integer
    Dim f      As Integer
    Dim found  As Boolean
    Dim SaveDB As Boolean
    
    ReDim DB(0)
    Path = GetFilePath("Users.txt")
    
    If LenB(Dir$(Path)) > 0 Then
        f = FreeFile
        Open Path For Input As #f
            
        If LOF(f) > 1 Then
            Do
                
                Line Input #f, s
                
                If InStr(1, s, " ", vbTextCompare) > 0 Then
                    x() = Split(s, " ", 10)
                    
                    If UBound(x) > 0 Then
                        ReDim Preserve DB(i)
                        
                        With DB(i)
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

' Writes database to disk
' Updated 9/13/06 for new features
Public Sub WriteDatabase(ByVal U As String)
    Dim f As Integer
    Dim i As Integer
    
    On Error GoTo WriteDatabase_Exit

    f = FreeFile
    
    Open U For Output As #f
        For i = LBound(DB) To UBound(DB)
            If (LenB(DB(i).Username) > 0) Then
                Print #f, DB(i).Username;
                Print #f, " " & DB(i).Rank;
                Print #f, " " & IIf(Len(DB(i).Flags) > 0, DB(i).Flags, "%");
                Print #f, " " & IIf(Len(DB(i).AddedBy) > 0, DB(i).AddedBy, "%");
                Print #f, " " & IIf(DB(i).AddedOn > 0, DateCleanup(DB(i).AddedOn), "%");
                Print #f, " " & IIf(Len(DB(i).ModifiedBy) > 0, DB(i).ModifiedBy, "%");
                Print #f, " " & IIf(DB(i).ModifiedOn > 0, DateCleanup(DB(i).ModifiedOn), "%");
                Print #f, " " & IIf(Len(DB(i).Type) > 0, DB(i).Type, "USER");
                Print #f, " " & IIf(Len(DB(i).Groups) > 0, DB(i).Groups, "%");
                Print #f, " " & IIf(Len(DB(i).BanMessage) > 0, DB(i).BanMessage, "%")
            End If
        Next i

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


Private Function CheckUser(ByVal User As String, Optional ByVal allow_illegal As Boolean = False) As Boolean
    
    Dim i       As Integer
    Dim bln     As Boolean
    Dim illegal As Boolean
    Dim invalid As Boolean
    
    If (Left$(User, 1) = "*") Then
        User = Mid$(User, 2)
    End If
    
    User = Replace(User, "@USWest", vbNullString, 1)
    User = Replace(User, "@USEast", vbNullString, 1)
    User = Replace(User, "@Asia", vbNullString, 1)
    User = Replace(User, "@Europe", vbNullString, 1)
    
    User = Replace(User, "@Lordaeron", vbNullString, 1)
    User = Replace(User, "@Azeroth", vbNullString, 1)
    User = Replace(User, "@Kalimdor", vbNullString, 1)
    User = Replace(User, "@Northrend", vbNullString, 1)
    
    If (Len(User) = 0) Then
        invalid = True
    ElseIf (Len(User) > 15) Then
        invalid = True
    Else
        ' 95 (a)
        ' 65 (A)
        ' 48 (0)
        ' 57 (9)
    
        For i = 1 To Len(User)
            Dim currentCharacter As String
            currentCharacter = Mid$(User, i, 1)

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
        Next i
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

' this fully converts a username based on naming conventions
Public Function ConvertUsername(ByVal Username As String) As String
    If (LenB(Username) = 0) Then
        ConvertUsername = Username
    Else
        ' handle namespace conversions (@gateways)
        ConvertUsername = ConvertUsernameGateway(Username)
        
        ' handle D2 naming conventions
        ConvertUsername = ConvertUsernameD2(ConvertUsername, Username)
    End If
End Function

' this converts a username only on gateway conventions
Public Function ConvertUsernameGateway(ByVal Username As String) As String
    Dim Index          As Long    ' index of substring in string
    Dim Gateways(5, 2) As String  ' list of known namespaces
    Dim blnIsW3        As Boolean ' whether this bot is on WC3
    Dim MyGateway      As String  ' the bot's gateway
    Dim intConvert     As Integer ' convert type, 0=none, 1=wc3->legacy, 2=legacy->wc3, 3=show all
    Dim blnConvert     As Boolean ' whether we are converting
    Dim i              As Integer ' loop iterator
    Dim blnOnOther     As Boolean ' whether the user is on the other namespace
    Dim MyGatewayIndex As Integer ' the bot's namespace index, 0=legacy 1=wc3

    ' populate list
    Gateways(0, 0) = "USWest"
    Gateways(0, 1) = "Lordaeron"
    Gateways(1, 0) = "USEast"
    Gateways(1, 1) = "Azeroth"
    Gateways(2, 0) = "Asia"
    Gateways(2, 1) = "Kalimdor"
    Gateways(3, 0) = "Europe"
    Gateways(3, 1) = "Northrend"
    Gateways(4, 0) = "Beta"
    Gateways(4, 1) = "Westfall"
    
    ' store whether we are on WC3
    blnIsW3 = _
        ((StrReverse$(BotVars.Product) = "WAR3") Or _
         (StrReverse$(BotVars.Product) = "W3XP"))
    
    ' store how we will be converting namespaces
    intConvert = BotVars.GatewayConventions
    
    ' get my gateway
    MyGateway = BotVars.Gateway
    
    ' handle not having gateway yet
    If (LenB(MyGateway) = 0) Then
        ConvertUsernameGateway = Username
        Exit Function
    End If
    
    ' get my namespace index
    For i = 0 To 4
        If (StrComp(Gateways(i, 0), MyGateway, vbTextCompare) = 0) Then
            MyGatewayIndex = 0
            Exit For
        ElseIf (StrComp(Gateways(i, 1), MyGateway, vbTextCompare) = 0) Then
            MyGatewayIndex = 1
            Exit For
        End If
    Next i
    
    ' is user on other namespace?
    Index = InStr(1, Username, "@" & Gateways(i, IIf(MyGatewayIndex = 0, 1, 0)), vbTextCompare)
    
    ' store whether user is on other namespace
    ' (whether the other @gateway was found)
    blnOnOther = (Index > 0)
    
    ' choose action
    Select Case intConvert
        Case 0: ' default, no conversions
            blnConvert = False
        Case 1: ' legacy, convert if we are on wc3
            blnConvert = (blnIsW3)
        Case 2: ' wc3, convert if we are not on wc3 (we are on legacy)
            blnConvert = (Not blnIsW3)
        Case 3: ' show all, convert if they are on this namespace
            blnConvert = (Not blnOnOther)
        Case Else: ' default
            blnConvert = False
    End Select
    
    If (blnConvert) Then
        ' we are converting
        If (blnOnOther) Then
            ' return username without other namespace
            ConvertUsernameGateway = Left$(Username, Index - 1)
        Else
            ' return username with our namespace
            ConvertUsernameGateway = Username & "@" & BotVars.Gateway
        End If
    Else
        ' we are not converting, leave it alone
        ConvertUsernameGateway = Username
    End If
    
End Function

' this converts a username only on D2 naming conventions
Public Function ConvertUsernameD2(ByVal Username As String, Optional ByVal RealUsername As String) As String
    Dim Index          As Long       ' index of substring in string
    Dim strFormat      As String     ' D2 naming format
    Dim Title          As String     ' D2 character title
    Dim UserObj        As clsUserObj ' user object to get more D2 information
    Dim Char           As String     ' D2 character name
    Dim Name           As String     ' D2 account name
    
    If IsMissing(RealUsername) Then
        RealUsername = Username
    End If
    
    If (BotVars.UseD2Naming = False) Then
        ' D2 naming disabled
        Index = InStr(1, Username, "*", vbBinaryCompare)
        
        If (Index > 0) Then
            ' user has star in name
            ConvertUsernameD2 = Mid$(Username, Index + 1)
        Else
            ' user has no star in name
            ConvertUsernameD2 = Username
        End If
    Else
        ' d2 naming enabled
        Index = InStr(1, Username, "*", vbBinaryCompare)
        
        If (Index > 1) Then
            ' user has star in name after position 1
            ' (we are on D2 and the character is provided)
            
            ' get d2 naming format
            strFormat = BotVars.D2NamingFormat
            strFormat = Replace$(strFormat, "title ", "{0}", 1, 1, vbTextCompare)
            strFormat = Replace$(strFormat, "char", "{1}", 1, 1, vbTextCompare)
            strFormat = Replace$(strFormat, "name", "{2}", 1, 1, vbTextCompare)
            
            ' get char and name from username
            Char = Left$(Username, Index - 1)
            Name = Mid$(Username, Index + 1)
            
            ' get D2 character title, if available
            Set UserObj = g_Channel.GetUserEx(RealUsername)
            Title = UserObj.Stats.CharacterTitle
            If (LenB(Title) > 0) Then Title = Title & " "
            Set UserObj = Nothing
            
            ' return formatted name
            ConvertUsernameD2 = StringFormat(strFormat, Title, Char, Name)
        ElseIf (Index = 0) Then
            ' user has no star in name
            If ((StrReverse$(BotVars.Product) = "D2DV") Or (StrReverse$(BotVars.Product) = "D2XP")) Then
                ' if on D2, add star and return
                ConvertUsernameD2 = "*" & Username
            Else
                ' if not on D2, get any D2 info we can anyway
                
                ' get d2 naming format
                strFormat = BotVars.D2NamingFormat
                strFormat = Replace$(strFormat, "title ", "{0}", 1, 1, vbTextCompare)
                strFormat = Replace$(strFormat, "char", "{1}", 1, 1, vbTextCompare)
                strFormat = Replace$(strFormat, "name", "{2}", 1, 1, vbTextCompare)
                
                ' get name from username
                Name = Username
                
                ' get D2 character name and title, if available
                Set UserObj = g_Channel.GetUserEx(RealUsername)
                Title = UserObj.Stats.CharacterTitle
                If (LenB(Title) > 0) Then Title = Title & " "
                Char = UserObj.Stats.CharacterName
                Set UserObj = Nothing
                
                ' if character name found
                If (LenB(Char) = 0) Then
                    ' if no character name, return *name
                    ConvertUsernameD2 = "*" & Username
                Else
                    ' if character name, return formatted name
                    ConvertUsernameD2 = StringFormat(strFormat, Title, Char, Name)
                End If
            End If
        Else
            ' if user has star in name at position 1, keep *name format
            ConvertUsernameD2 = Username
        End If
    End If
End Function

' reverses converting username gateways
Public Function ReverseConvertUsernameGateway(ByVal Username As String) As String
    Dim Index          As Long    ' index of substring in string
    Dim Gateways(5, 2) As String  ' list of known namespaces
    Dim blnIsW3        As Boolean ' whether this bot is on WC3
    Dim MyGateway      As String  ' the bot's gateway
    Dim intConvert     As Integer ' convert type, 0=none, 1=wc3->legacy, 2=legacy->wc3, 3=show all
    Dim blnConvert     As Boolean ' whether we are converting
    Dim i              As Integer ' loop iterator
    Dim blnOnThis      As Boolean ' whether the user is on this namespace
    Dim MyGatewayIndex As Integer ' the bot's namespace index, 0=legacy 1=wc3
    
    If (LenB(Username) = 0) Then
        Exit Function
    End If

    ' add * to D2 if not using D2 naming
    If ((StrReverse$(BotVars.Product) = "D2DV") Or (StrReverse$(BotVars.Product) = "D2XP")) Then
        If (BotVars.UseD2Naming = False) Then
            ' With reverseUsername() now being called from AddQ(), usernames
            ' in procedures called prior to AddQ() will no longer require
            ' prefixes; however, we want to ensure that a '*' was not already
            ' specified before conversion to allow older scripts and procedures
            ' to continue functioning correctly.  This check may be removed in
            ' future releases.
            If (Not Left$(Username, 1) = "*") Then
                Username = ("*" & Username)
            End If
        End If
    End If

    ' populate list
    Gateways(0, 0) = "USWest"
    Gateways(0, 1) = "Lordaeron"
    Gateways(1, 0) = "USEast"
    Gateways(1, 1) = "Azeroth"
    Gateways(2, 0) = "Asia"
    Gateways(2, 1) = "Kalimdor"
    Gateways(3, 0) = "Europe"
    Gateways(3, 1) = "Northrend"
    Gateways(4, 0) = "Beta"
    Gateways(4, 1) = "Westfall"
    
    ' store whether we are on WC3
    blnIsW3 = _
        ((StrReverse$(BotVars.Product) = "WAR3") Or _
         (StrReverse$(BotVars.Product) = "W3XP"))
    
    ' store how we will be converting namespaces
    intConvert = BotVars.GatewayConventions
    
    ' get my gateway
    MyGateway = BotVars.Gateway
    
    ' handle not having gateway yet
    If (LenB(MyGateway) = 0) Then
        ReverseConvertUsernameGateway = Username
        Exit Function
    End If
    
    ' get my namespace index
    For i = 0 To 4
        If (StrComp(Gateways(i, 0), MyGateway, vbTextCompare) = 0) Then
            MyGatewayIndex = 0
            Exit For
        ElseIf (StrComp(Gateways(i, 1), MyGateway, vbTextCompare) = 0) Then
            MyGatewayIndex = 1
            Exit For
        End If
    Next i
    
    ' is user on this namespace?
    Index = InStr(1, Username, "@" & MyGateway, vbTextCompare)
    
    ' store whether user is on this namespace
    ' (whether this @gateway was found)
    blnOnThis = (Index > 0)
    
    ' choose action
    Select Case intConvert
        Case 0: ' default, no conversions
            blnConvert = False
        Case 1: ' legacy, un-convert if we are on wc3
            blnConvert = (blnIsW3)
        Case 2: ' wc3, un-convert if we are not on wc3 (we are on legacy)
            blnConvert = (Not blnIsW3)
        Case 3: ' show all, un-convert if they are on this namespace
            blnConvert = (blnOnThis)
        Case Else: ' default
            blnConvert = False
    End Select
    
    If (blnConvert) Then
        ' we are converting
        If (blnOnThis) Then
            ' return username without this namespace
            ReverseConvertUsernameGateway = Left$(Username, Index - 1)
        Else
            ' return username with their namespace
            ReverseConvertUsernameGateway = Username & "@" & Gateways(i, IIf(MyGatewayIndex = 0, 1, 0))
        End If
    Else
        ' we are not converting, leave it alone
        ReverseConvertUsernameGateway = Username
    End If
End Function
