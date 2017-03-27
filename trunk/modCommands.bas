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
    
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim commands         As Collection
    Dim Command          As clsCommandObj
    Dim sCommand         As String
    
    ' replace message variables
    Message = Replace(Message, "%me", IIf(IsLocal, GetCurrentUsername, Username), 1, -1, vbTextCompare)
    
    ' Should the command system be bypassed entirely?
    If ((IsLocal) And (Left$(Message, 3) = "///")) Then
        frmChat.AddQ Mid$(Message, 3)
        Exit Function
    End If

    ' Get all of the commands in the message
    Set commands = clsCommandObj.IsCommand(Message, IIf(IsLocal, modGlobals.CurrentUsername, CleanUsername(Username)), _
            IsLocal, WasWhispered, Chr$(0))

    For Each Command In commands
        If (Command.HasAccess) Then
        
            ' Log the command
            sCommand = Command.Name
            If (LenB(Command.Args) > 0) Then
                sCommand = sCommand & ": " & Command.Args
            End If
            LogCommand IIf(Command.IsLocal, vbNullString, Username), sCommand
            
            ' Fire the command event
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
    
    'Unload memory
    Set Command = Nothing
    Set commands = Nothing
    
    Exit Function
    
' default (if all else fails) error handler to keep erroneous
' commands and/or input formats from killing me
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: #" & Err.Number & ": " & Err.Description & _
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
        Case "tagcheck":       Call modCommandsInfo.OnTagCheck(Command)
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
        Case "leaveclan":      Call modCommandsClan.OnLeaveClan(Command)
        Case "makechieftain":  Call modCommandsClan.OnMakeChieftain(Command)
        Case "motd":           Call modCommandsClan.OnMOTD(Command)
        Case "promote":        Call modCommandsClan.OnPromote(Command)
        Case "removemember":   Call modCommandsClan.OnRemoveMember(Command)
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
        Case "return":         Call modCommandsChat.OnReturn(Command)
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
        Case "pingban":        Call modCommandsOps.OnPingBan(Command)
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

' requires public
Public Function GetSafelist(ByVal Username As String) As Boolean
    Dim i As Long
    Dim gAcc As udtUserAccess
        
    gAcc = Database.GetUserAccess(Username)
        
    If (Not InStr(1, gAcc.Flags, "S", vbBinaryCompare) = 0) Then
        GetSafelist = True
    ElseIf (gAcc.Rank >= AutoModSafelistValue) Then
        GetSafelist = True
    End If
End Function

Public Function GetShitlist(ByVal Username As String) As String
    Dim gAcc As udtUserAccess
    Dim Ban  As Boolean
    
    gAcc = Database.GetUserAccess(Username)
    
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
    toCheck = Replace(toCheck, "[", ChrW(&HFF))
    toCheck = Replace(toCheck, "]", ChrW(&HF6))
    toCheck = Replace(toCheck, "~", ChrW(&HDC))
    toCheck = Replace(toCheck, "#", ChrW(&HE7))
    toCheck = Replace(toCheck, "-", ChrW(&HA3))
    toCheck = Replace(toCheck, "&", ChrW(&HA5))
    toCheck = Replace(toCheck, "@", ChrW(&HA4))
    toCheck = Replace(toCheck, "{", ChrW(&H192))
    toCheck = Replace(toCheck, "}", ChrW(&HE1))
    toCheck = Replace(toCheck, "^", ChrW(&HED))
    toCheck = Replace(toCheck, "`", ChrW(&HF3))
    toCheck = Replace(toCheck, "_", ChrW(&HFA))
    toCheck = Replace(toCheck, "+", ChrW(&HF1))
    toCheck = Replace(toCheck, "$", ChrW(&HF7))
    PrepareCheck = LCase$(toCheck)
End Function

' requires public
Public Function ReversePrepareCheck(ByVal toCheck As String) As String
    toCheck = Replace(toCheck, ChrW(&HFF), "[")
    toCheck = Replace(toCheck, ChrW(&HF6), "]")
    toCheck = Replace(toCheck, ChrW(&HDC), "~")
    toCheck = Replace(toCheck, ChrW(&HE7), "#")
    toCheck = Replace(toCheck, ChrW(&HA3), "-")
    toCheck = Replace(toCheck, ChrW(&HA5), "&")
    toCheck = Replace(toCheck, ChrW(&HA4), "@")
    toCheck = Replace(toCheck, ChrW(&H192), "{")
    toCheck = Replace(toCheck, ChrW(&HE1), "}")
    toCheck = Replace(toCheck, ChrW(&HED), "^")
    toCheck = Replace(toCheck, ChrW(&HF3), "`")
    toCheck = Replace(toCheck, ChrW(&HFA), "_")
    toCheck = Replace(toCheck, ChrW(&HF1), "+")
    toCheck = Replace(toCheck, ChrW(&HF7), "$")
    ReversePrepareCheck = LCase$(toCheck)
End Function


Private Function CheckUser(ByVal User As String, Optional ByVal allow_illegal As Boolean = False) As Boolean
    
    Dim i       As Integer
    Dim bln     As Boolean
    Dim illegal As Boolean
    Dim invalid As Boolean
    
    If (Left$(User, 1) = "*") Then
        User = Mid$(User, 2)
    End If
    
    User = Replace(User, Config.GatewayDelimiter & "USWest", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "USEast", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "Asia", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "Europe", vbNullString, 1)
    
    User = Replace(User, Config.GatewayDelimiter & "Lordaeron", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "Azeroth", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "Kalimdor", vbNullString, 1)
    User = Replace(User, Config.GatewayDelimiter & "Northrend", vbNullString, 1)
    
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
Public Function ConvertUsername(ByVal Username As String, Optional ByVal iConvention As Integer = -1) As String
    If (LenB(Username) = 0) Then
        ConvertUsername = Username
    Else
        ' handle namespace conversions (@gateways)
        ConvertUsername = ConvertUsernameGateway(Username, iConvention)
        
        ' handle D2 naming conventions
        ConvertUsername = ConvertUsernameD2(ConvertUsername, Username)
    End If
End Function

' this converts a username only on gateway conventions
Public Function ConvertUsernameGateway(ByVal Username As String, Optional ByVal iConvertType As Integer = -1) As String
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
        ((StrReverse$(BotVars.Product) = PRODUCT_WAR3) Or _
         (StrReverse$(BotVars.Product) = PRODUCT_W3XP))
    
    ' store how we will be converting namespaces
    If (iConvertType = -1) Then
        intConvert = BotVars.GatewayConventions
    Else
        intConvert = iConvertType
    End If
    
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
    Index = InStr(1, Username, Config.GatewayDelimiter & Gateways(i, IIf(MyGatewayIndex = 0, 1, 0)), vbTextCompare)
    
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
            ConvertUsernameGateway = Username & Config.GatewayDelimiter & BotVars.Gateway
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
            If ((StrReverse$(BotVars.Product) = PRODUCT_D2DV) Or (StrReverse$(BotVars.Product) = PRODUCT_D2XP)) Then
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
    If ((StrReverse$(BotVars.Product) = PRODUCT_D2DV) Or (StrReverse$(BotVars.Product) = PRODUCT_D2XP)) Then
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
        ((StrReverse$(BotVars.Product) = PRODUCT_WAR3) Or _
         (StrReverse$(BotVars.Product) = PRODUCT_W3XP))
    
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
    Index = InStr(1, Username, Config.GatewayDelimiter & MyGateway, vbTextCompare)
    
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
            ReverseConvertUsernameGateway = Username & Config.GatewayDelimiter & Gateways(i, IIf(MyGatewayIndex = 0, 1, 0))
        End If
    Else
        ' we are not converting, leave it alone
        ReverseConvertUsernameGateway = Username
    End If
End Function
