Attribute VB_Name = "modCommandCode"
' modCommandCode.bas
' ...

Option Explicit

' Winamp Constants
Private Const WA_PREVTRACK   As Long = 40044 ' ...
Private Const WA_NEXTTRACK   As Long = 40048 ' ...
Private Const WA_PLAY        As Long = 40045 ' ...
Private Const WA_PAUSE       As Long = 40046 ' ...
Private Const WA_STOP        As Long = 40047 ' ...
Private Const WA_FADEOUTSTOP As Long = 40147 ' ...

Public Flood    As String ' ...?
Public floodCap As Byte   ' ...?

' prepares commands for processing, and calls helper functions associated with
' processing
Public Function ProcessCommand(ByVal Username As String, ByVal Message As String, _
    Optional ByVal InBot As Boolean = False, Optional ByVal WhisperedIn As Boolean = False) As Boolean
    
    ' ...
    Dim ConsoleAccessResponse As udtGetAccessResponse
    
    Dim X()          As String  ' ...
    Dim i            As Integer ' ...
    Dim tmpMsg       As String  ' ...
    Dim cmdRet()     As String  ' ...
    Dim publicOutput As Boolean
    
    ' create single command data array element for safe bounds checking
    ReDim Preserve cmdRet(0)
    
    ' create console access response structure
    With ConsoleAccessResponse
        .Access = 1001
        .Flags = "A"
    End With

    ' store local copy of message
    tmpMsg = Message
    
    ' replace message variables
    tmpMsg = Replace(tmpMsg, "%me", IIf((InBot), CurrentUsername, Username), 1)
    
    If (InBot = False) Then
        ' check for commands using universal command identifier (?)
        If (StrComp(LCase$(tmpMsg), "?trigger", vbBinaryCompare) = 0) Then
            ' remove universal command identifier from message
            tmpMsg = Mid$(tmpMsg, 2)
            
        ElseIf ((Len(tmpMsg) >= Len(BotVars.Trigger)) And _
                (Left$(tmpMsg, Len(BotVars.Trigger)) = BotVars.Trigger)) Then
            
            ' remove command identifier from message
            tmpMsg = Mid$(tmpMsg, Len(BotVars.Trigger) + 1)
        
            ' check for command identifier and name combination
            ' (e.g., .Eric[nK] say hello)
            If (Len(tmpMsg) >= (Len(CurrentUsername) + 1)) Then
                If (StrComp(Left$(tmpMsg, Len(CurrentUsername) + 1), _
                    CurrentUsername & Space(1), vbTextCompare) = 0) Then
                        
                    ' remove username (and space) from message
                    tmpMsg = Mid$(tmpMsg, Len(CurrentUsername) + 2)
                End If
            End If
        Else
            ' return negative result indicating that message does not contain
            ' a valid command identifier
            ProcessCommand = False
            
            ' exit function
            Exit Function
        End If
    Else
        ' remove slash (/) from message
        tmpMsg = Mid$(tmpMsg, 2)
        
        ' check for second slash indicating
        ' public output
        If (Left$(tmpMsg, 1) = "/") Then
            ' enable public display of command
            publicOutput = True
        
            ' remove second slash (/) from message
            tmpMsg = Mid$(tmpMsg, 2)
        End If
    End If

    ' check for multiple commands
    If (InStr(1, tmpMsg, "; ", vbTextCompare) > 0) Then
        X = Split(tmpMsg, "; ")
        
        ' loop through commands
        For i = 0 To UBound(X)
            ' send command to main processor
            If (InBot = True) Then
                ProcessCommand = ExecuteCommand(Username, ConsoleAccessResponse, _
                    X(i), InBot, cmdRet())
            Else
                ProcessCommand = ExecuteCommand(Username, GetAccess(Username), X(i), _
                    InBot, cmdRet())
            End If
            
            ' display command response
            If (cmdRet(0) <> vbNullString) Then
                Dim j As Integer ' ...
            
                ' loop through command response
                For j = 0 To UBound(cmdRet)
                    If ((InBot) And (Not (publicOutput))) Then
                        Call AddChat(RTBColors.ConsoleText, cmdRet(i))
                    Else
                        Call AddQ(cmdRet(i), 1)
                    End If
                Next j
            End If
        Next i
    Else
        ' send command to main processor
        If (InBot = True) Then
            ProcessCommand = ExecuteCommand(Username, ConsoleAccessResponse, tmpMsg, _
                InBot, cmdRet())
        Else
            ProcessCommand = ExecuteCommand(Username, GetAccess(Username), tmpMsg, _
                InBot, cmdRet())
        End If
        
        ' display command response
        If (cmdRet(0) <> vbNullString) Then
            ' loop through command response
            For i = 0 To UBound(cmdRet)
                If ((InBot) And (Not (publicOutput))) Then
                    Call AddChat(RTBColors.ConsoleText, cmdRet(i))
                Else
                    Call AddQ(cmdRet(i), 1)
                End If
            Next i
        End If
    End If
End Function ' end function ProcessCommand

' command processing helper function
Public Function ExecuteCommand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal Message As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean

    Dim tmpMsg   As String  ' ...
    Dim cmdName  As String  ' stores command name
    Dim msgData  As String  ' stores unparsed command parameters
    Dim blnNoCmd As Boolean ' stores result of command switch (true = no command found)
    Dim i        As Integer ' ...
    
    ' create single command data array element for safe bounds checking
    ' and to help aide in a reduction of command function overhead
    ReDim Preserve cmdRet(0)
    
    ' store local copy of message
    tmpMsg = Message

    ' grab command name & message data
    If (InStr(1, tmpMsg, Space(1), vbBinaryCompare) <> 0) Then
        ' grab command name
        cmdName = Left$(tmpMsg, (InStr(1, tmpMsg, Space(1), _
            vbBinaryCompare) - 1))
        
        ' remove command name (and space) from message
        tmpMsg = Mid$(tmpMsg, Len(cmdName) + 2)
        
        ' grab message data
        msgData = tmpMsg
    Else
        ' grab command name
        cmdName = tmpMsg
    End If
    
    ' convert command name to lcase
    cmdName = LCase$(cmdName)

    If ((ValidateAccess(dbAccess, cmdName) = True) Or (InBot = True)) Then
        ' command switch
        Select Case (cmdName)
            Case "quit":                         Call OnQuit(Username, dbAccess, msgData, InBot, cmdRet())
            Case "locktext":                     Call OnLockText(Username, dbAccess, msgData, InBot, cmdRet())
            Case "allowmp3":                     Call OnAllowMp3(Username, dbAccess, msgData, InBot, cmdRet())
            Case "loadwinamp":                   Call OnLoadWinamp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "floodmode", "efp":             Call OnEfp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "home", "joinhome":             Call OnHome(Username, dbAccess, msgData, InBot, cmdRet())
            Case "clan", "c":                    Call OnClan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "peonban":                      Call OnPeonBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "invite":                       Call OnInvite(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setmotd":                      Call OnSetMotd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "where":                        Call OnWhere(Username, dbAccess, msgData, InBot, cmdRet())
            Case "qt", "quiettime":              Call OnQuietTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "roll":                         Call OnRoll(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sweepban", "cb":               Call OnSweepBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sweepignore", "cs":            Call OnSweepIgnore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setname":                      Call OnSetName(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setpass":                      Call OnSetPass(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setkey":                       Call OnSetKey(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setexpkey":                    Call OnSetExpKey(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setserver":                    Call OnSetServer(Username, dbAccess, msgData, InBot, cmdRet())
            Case "giveup", "op":                 Call OnGiveUp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "math", "eval":                 Call OnMath(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idlebans", "ib":               Call OnIdleBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "chpw":                         Call OnChPw(Username, dbAccess, msgData, InBot, cmdRet())
            Case "join":                         Call OnJoin(Username, dbAccess, msgData, InBot, cmdRet())
            Case "sethome":                      Call OnSetHome(Username, dbAccess, msgData, InBot, cmdRet())
            Case "resign":                       Call OnResign(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cbl", "clearbanlist":          Call OnClearBanList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "koy", "kickonyell":            Call OnKickOnYell(Username, dbAccess, msgData, InBot, cmdRet())
            Case "rejoin", "rj":                 Call OnRejoin(Username, dbAccess, msgData, InBot, cmdRet())
            Case "plugban":                      Call OnPlugBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "clist", "clientbans", "cbans": Call OnClientBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setvol":                       Call OnSetVol(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cadd":                         Call OnCAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cdel", "delclient":            Call OnCDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "banned":                       Call OnBanned(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ipbans":                       Call OnIPBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ipban":                        Call OnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unipban":                      Call OnUnIPBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "designate", "des":             Call OnDesignate(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shuffle":                      Call OnShuffle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "repeat":                       Call OnRepeat(Username, dbAccess, msgData, InBot, cmdRet())
            Case "next":                         Call OnNext(Username, dbAccess, msgData, InBot, cmdRet())
            Case "prev":                         Call OnPrev(Username, dbAccess, msgData, InBot, cmdRet())
            Case "protect":                      Call OnProtect(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whispercmds", "wc":            Call OnWhisperCmds(Username, dbAccess, msgData, InBot, cmdRet())
            Case "stop":                         Call OnStop(Username, dbAccess, msgData, InBot, cmdRet())
            Case "play":                         Call OnPlay(Username, dbAccess, msgData, InBot, cmdRet())
            Case "useitunes":                    Call OnUseiTunes(Username, dbAccess, msgData, InBot, cmdRet())
            Case "usewinamp":                    Call OnUseWinamp(Username, dbAccess, msgData, InBot, cmdRet())
            Case "pause":                        Call OnPause(Username, dbAccess, msgData, InBot, cmdRet())
            Case "fos":                          Call OnFos(Username, dbAccess, msgData, InBot, cmdRet())
            Case "rem", "del":                   Call OnRem(Username, dbAccess, msgData, InBot, cmdRet())
            Case "reconnect":                    Call OnReconnect(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unigpriv":                     Call OnUnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
            Case "igpriv":                       Call OnIgPriv(Username, dbAccess, msgData, InBot, cmdRet())
            Case "block":                        Call OnBlock(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idletime", "idlewait":         Call OnIdleTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idle":                         Call OnIdle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitdel":                      Call OnShitDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safedel":                      Call OnSafeDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagdel":                       Call OnTagDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setidle":                      Call OnSetIdle(Username, dbAccess, msgData, InBot, cmdRet())
            Case "idletype":                     Call OnIdleType(Username, dbAccess, msgData, InBot, cmdRet())
            Case "filter":                       Call OnFilter(Username, dbAccess, msgData, InBot, cmdRet())
            Case "trigger":                      Call OnTrigger(Username, dbAccess, msgData, InBot, cmdRet())
            Case "settrigger":                   Call OnSetTrigger(Username, dbAccess, msgData, InBot, cmdRet())
            Case "levelban":                     Call OnLevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "d2levelban":                   Call OnD2LevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "pon", "phrasebans on":         Call OnPhraseBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "poff", "phrasebans off":       Call OnPhraseBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cbans":                        Call OnCBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "pstatus", "phrasebans":        Call OnPhraseBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "mimic":                        Call OnMimic(Username, dbAccess, msgData, InBot, cmdRet())
            Case "nomimic":                      Call OnNoMimic(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setpmsg":                      Call OnSetPMsg(Username, dbAccess, msgData, InBot, cmdRet())
            Case "setcmdaccess":                 Call OnSetCmdAccess(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cmdadd", "addcmd":             Call OnCmdAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cmddel", "delcmd":             Call OnCmdDel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cmdlist":                      Call OnCmdList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "phrases", "plist":             Call OnPhrases(Username, dbAccess, msgData, InBot, cmdRet())
            Case "addphrase", "padd":            Call OnAddPhrase(Username, dbAccess, msgData, InBot, cmdRet())
            Case "delphrase", "pdel":            Call OnDelPhrase(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagban", "addtag", "tagadd":   Call OnTagBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "fadd":                         Call OnFAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "frem":                         Call OnFRem(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safelist":                     Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safeadd":                      Call OnSafeAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "exile":                        Call OnExile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unexile":                      Call OnUnExile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitlist", "sl":               Call OnShitList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "safelist":                     Call OnSafeList(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagbans":                      Call OnTagBans(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shitadd", "pban":              Call OnShitAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "dnd":                          Call OnDND(Username, dbAccess, msgData, InBot, cmdRet())
            Case "bancount":                     Call OnBanCount(Username, dbAccess, msgData, InBot, cmdRet())
            Case "tagcheck":                     Call OnTagCheck(Username, dbAccess, msgData, InBot, cmdRet())
            Case "slcheck", "shitcheck":         Call OnSLCheck(Username, dbAccess, msgData, InBot, cmdRet())
            Case "readfile":                     Call OnReadFile(Username, dbAccess, msgData, InBot, cmdRet())
            Case "levelban", "levelbans":        Call OnLevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "d2levelban", "d2levelbans":    Call OnD2LevelBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "greet":                        Call OnGreet(Username, dbAccess, msgData, InBot, cmdRet())
            Case "allseen":                      Call OnAllSeen(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ban":                          Call OnBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unban":                        Call OnUnBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "kick":                         Call OnKick(Username, dbAccess, msgData, InBot, cmdRet())
            Case "lastwhisper", "lw":            Call OnLastWhisper(Username, dbAccess, msgData, InBot, cmdRet())
            Case "say":                          Call OnSay(Username, dbAccess, msgData, InBot, cmdRet())
            Case "expand":                       Call OnExpand(Username, dbAccess, msgData, InBot, cmdRet())
            Case "detail", "dbd":                Call OnDetail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "info":                         Call OnInfo(Username, dbAccess, msgData, InBot, cmdRet())
            Case "shout":                        Call OnShout(Username, dbAccess, msgData, InBot, cmdRet())
            Case "voteban":                      Call OnVoteBan(Username, dbAccess, msgData, InBot, cmdRet())
            Case "votekick":                     Call OnVoteKick(Username, dbAccess, msgData, InBot, cmdRet())
            Case "vote":                         Call OnVote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cancel":                       Call OnCancel(Username, dbAccess, msgData, InBot, cmdRet())
            Case "back":                         Call OnBack(Username, dbAccess, msgData, InBot, cmdRet())
            Case "uptime":                       Call OnUptime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "away":                         Call OnAway(Username, dbAccess, msgData, InBot, cmdRet())
            Case "mp3":                          Call OnMP3(Username, dbAccess, msgData, InBot, cmdRet())
            Case "deldef":                       Call OnDelDef(Username, dbAccess, msgData, InBot, cmdRet())
            Case "define", "def":                Call OnDefine(Username, dbAccess, msgData, InBot, cmdRet())
            Case "newdef":                       Call OnNewDef(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ping":                         Call OnPing(Username, dbAccess, msgData, InBot, cmdRet())
            Case "addquote":                     Call OnAddQuote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "owner":                        Call OnOwner(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ignore", "ign":                Call OnIgnore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "quote":                        Call OnQuote(Username, dbAccess, msgData, InBot, cmdRet())
            Case "unignore":                     Call OnUnignore(Username, dbAccess, msgData, InBot, cmdRet())
            Case "cq", "scq":                    Call OnCQ(Username, dbAccess, msgData, InBot, cmdRet())
            Case "time":                         Call OnTime(Username, dbAccess, msgData, InBot, cmdRet())
            Case "getping", "pingme":            Call OnGetPing(Username, dbAccess, msgData, InBot, cmdRet())
            Case "checkmail":                    Call OnCheckMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "getmail":                      Call OnGetMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whoami":                       Call OnWhoAmI(Username, dbAccess, msgData, InBot, cmdRet())
            Case "add", "set":                   Call OnAdd(Username, dbAccess, msgData, InBot, cmdRet())
            Case "mmail":                        Call OnMMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "bmail", "mail":                Call OnMail(Username, dbAccess, msgData, InBot, cmdRet())
            Case "designated":                   Call OnDesignated(Username, dbAccess, msgData, InBot, cmdRet())
            Case "flip":                         Call OnFlip(Username, dbAccess, msgData, InBot, cmdRet())
            Case "ver", "about", "version":      Call OnAbout(Username, dbAccess, msgData, InBot, cmdRet())
            Case "server":                       Call OnServer(Username, dbAccess, msgData, InBot, cmdRet())
            Case "findr":                        Call OnFindR(Username, dbAccess, msgData, InBot, cmdRet())
            Case "find":                         Call OnFind(Username, dbAccess, msgData, InBot, cmdRet())
            Case "whois":                        Call OnWhoIs(Username, dbAccess, msgData, InBot, cmdRet())
            Case "findattr", "findflag":         Call OnFindAttr(Username, dbAccess, msgData, InBot, cmdRet())
            Case Else
                blnNoCmd = True
        End Select
    Else
        Exit Function
    End If
    
    ' append entry to command log
    Call LogCommand(Username, Message)
    
    ' was a command found? return.
    ExecuteCommand = (Not (blnNoCmd))
End Function

' handle quit command
Private Function OnQuit(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Call frmChat.Form_Unload(0)
End Function ' end function OnQuit

' handle locktext command
Private Function OnLockText(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Call frmChat.mnuLock_Click
End Function ' end function OnLockText

' handle allowmp3 command
Private Function OnAllowMp3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
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
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = LoadWinamp(ReadCFG("Other", "WinampPath"))
            
    If (Len(tmpBuf) < 1) Then
        Exit Function
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLoadWinamp

' handle efp command
Private Function OnEfp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (Left$(msgData, 2) = "on") Then
        ' enable efp
        Call frmChat.SetFloodbotMode(1)
        
        tmpBuf = "Emergency floodbot protection enabled."
     ElseIf (Left$(msgData, 6) = "status") Then
        If (bFlood) Then
            frmChat.AddChat RTBColors.TalkBotUsername, "Emergency floodbot protection is " & _
                "enabled. (No messages can be sent to battle.net.)"
        Else
            tmpBuf = "Emergency floodbot protection is disabled."
        End If
    ElseIf (Left$(msgData, 3) = "off") Then
        ' disable efp
        Call frmChat.SetFloodbotMode(0)
        
        tmpBuf = "Emergency floodbot protection disabled."
    End If
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnEfp

' handle home command
Private Function OnHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    AddQ "/join " & BotVars.HomeChannel, 1
End Function ' end function OnHome

' handle clan command
Private Function OnClan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    ' is bot a channel operator?
    If ((MyFlags) And (&H2)) Then
        Select Case (LCase$(msgData))
            Case "public", "pub"
                tmpBuf = "Clan channel is now public."
                
                ' set clan channel to public
                AddQ "/clan public", 1
            Case "private", "priv"
                tmpBuf = "Clan channel is now private."
                
                ' set clan channel to private
                AddQ "/clan private", 1
            Case Else
                tmpBuf = "/clan " & msgData
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
    
    Dim tmpBuf As String ' temporary output buffer
            
    Select Case (LCase$(msgData))
        Case "on"
            ' enable peon banning
            BotVars.BanPeons = 1
            
            ' write configuration entry
            WriteINI "Other", "PeonBans", "1"
            
            tmpBuf = "Peon banning activated."
        Case "off"
            ' disable peon banning
            BotVars.BanPeons = 0
            
            ' write configuration entry
            WriteINI "Other", "PeonBans", "0"
            
            tmpBuf = "Peon banning deactivated."
        Case "status"
            tmpBuf = "The bot is currently "
            
            If (BotVars.BanPeons = 0) Then
                tmpBuf = tmpBuf & "not banning peons."
            Else
                tmpBuf = tmpBuf & "banning peons."
            End If
    End Select
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPeonBan

' handle invite command
Private Function OnInvite(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (IsW3) Then
        If (Clan.MyRank >= 3) Then
            Call InviteToClan(msgData)
            
            tmpBuf = msgData & ": Clan invitation sent."
        Else
            tmpBuf = "The bot must hold Shaman or Chieftain rank to invite users."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnInvite

' handle setmotd command
Private Function OnSetMotd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (IsW3) Then
        If (Clan.MyRank >= 3) Then
            Call SetClanMOTD(msgData)
            
            tmpBuf = "Clan MOTD set."
        Else
            tmpBuf = "Shaman or Chieftain rank is required to set the MOTD."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetMotd

' handle where command
Private Function OnWhere(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = "I am currently in channel " & gChannel.Current & " (" & _
        colUsersInChannel.Count & " users present)"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhere

' handle quiettime command
Private Function OnQuietTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    Select Case LCase$(msgData)
        Case "on"
            ' enable quiettime
            BotVars.QuietTime = True
        
            ' write configuration entry
            WriteINI "Main", "QuietTime", "Y"
            
            tmpBuf = "Quiet-time enabled."
            
        Case "off"
            ' disable quiettime
            BotVars.QuietTime = False
            
            ' write configuration entry
            WriteINI "Main", "QuietTime", "N"
            
            tmpBuf = "Quiet-time disabled."
            
        Case "status"
            If (BotVars.QuietTime) Then
                tmpBuf = "Quiet-time is currently enabled."
            Else
                tmpBuf = "Quiet-time is currently disabled."
            End If
        
        Case Else
            tmpBuf = "Invalid arguments."
    End Select
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnQuietTime

' handle roll command
Private Function OnRoll(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
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
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnRoll

' handle sweepban command
Private Function OnSweepBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u As String
    Dim Y As String

    Caching = True
    Call Cache(vbNullString, 255, "ban ")
    AddQ "/who " & msgData, 1
End Function ' end function OnSweepBan

' handle sweepignore command
Private Function OnSweepIgnore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u As String
    Dim Y As String
    
    Caching = True
    Call Cache(vbNullString, 255, "squelch ")
    AddQ "/who " & msgData, 1
End Function ' end function OnSweepIgnore

' handle setname command
Private Function OnSetName(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' only allow use of setname command while on-line to prevent beta
    ' authorization bypassing
    If ((Not (g_Online = True)) Or (g_Connected = False)) Then
        Exit Function
    End If

    ' write configuration entry
    WriteINI "Main", "Username", msgData
    
    ' set username
    BotVars.Username = msgData
    
    tmpBuf = "New username set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetName

' handle setpass command
Private Function OnSetPass(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' write configuration entry
    WriteINI "Main", "Password", msgData
    
    ' set password
    BotVars.Password = msgData
    
    tmpBuf = "New password set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetPass

' handle math command
Private Function OnMath(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Math now has 3 levels.
    '50: No UI, and no CreateObject()
    '80: UI, no CreateObject()
    '100: No restrictions
    'Hdx - 09-25-07

    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer

    dbAccess = GetAccess(Username)
    
    If (dbAccess.Access >= GetAccessINIValue("math80", 80)) Then
        If (dbAccess.Access >= GetAccessINIValue("math100", 100)) Then
            frmChat.SCRestricted.AllowUI = True
            
            tmpBuf = frmChat.SCRestricted.Eval(msgData)
        Else
            If (InStr(LCase$(msgData), "createobject") > 0) Then
                tmpBuf = "Evaluation error."
            Else
                frmChat.SCRestricted.AllowUI = True
            End If
            
            tmpBuf = frmChat.SCRestricted.Eval(msgData)
        End If
    Else
        If (InStr(LCase$(msgData), "createobject") > 0) Then
            tmpBuf = "Evaluation error."
        Else
            frmChat.SCRestricted.AllowUI = False
            
            tmpBuf = frmChat.SCRestricted.Eval(msgData)
        End If
    End If
    
    While (Left$(tmpBuf, 1) = "/")
        tmpBuf = Mid$(tmpBuf, 2)
    Wend ' end loop
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMath

' handle setkey command
Private Function OnSetKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' clean data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)

    ' write configuration information
    WriteINI "Main", "CDKey", msgData
    
    ' set CD-Key
    BotVars.CDKey = msgData
    
    tmpBuf = "New cdkey set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetKey

' handle setexpkey command
Private Function OnSetExpKey(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' clean data
    msgData = Replace(msgData, "-", vbNullString)
    msgData = Replace(msgData, " ", vbNullString)
    
    ' write configuration entry
    WriteINI "Main", "LODKey", msgData
    
    ' set expansion CD-Key
    BotVars.LODKey = msgData
    
    tmpBuf = "New expansion CD-key set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetExpKey

' handle setserver command
Private Function OnSetServer(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    ' write configuration information
    WriteINI "Main", "Server", msgData
    
    ' set server
    BotVars.Server = msgData
    
    tmpBuf = "New server set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetServer

' handle giveup command
Private Function OnGiveUp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    If (CheckChannel(msgData) > 0) Then
        AddQ "/designate " & IIf(Dii, "*", vbNullString) & msgData
        AddQ "/resign"
    End If
End Function ' end function OnGiveUp

' handle idlebans command
Private Function OnIdleBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String ' ...
    Dim tmpBuf     As String ' temporary output buffer
    Dim subCmd     As String
    
    subCmd = LCase$(Mid$(msgData, 1, InStr(1, msgData, Space$(1), vbBinaryCompare)))
    
    If (Len(subCmd) > 0) Then
        Select Case (subCmd)
            Case "on"
                strArray() = Split(msgData, " ")
                
                BotVars.IB_On = BTRUE
                
                If (UBound(strArray) > 1) Then
                    If (StrictIsNumeric(strArray(2))) Then
                        BotVars.IB_Wait = strArray(2)
                    End If
                End If
                
                If (BotVars.IB_Wait > 0) Then
                    tmpBuf = "IdleBans activated, with a delay of " & BotVars.IB_Wait & "."
                    
                    WriteINI "Other", "IdleBans", "Y"
                    WriteINI "Other", "IdleBanDelay", BotVars.IB_Wait
                Else
                    BotVars.IB_Wait = 400
                    
                    tmpBuf = "IdleBans activated, using the default delay of 400."
                    
                    WriteINI "Other", "IdleBanDelay", "400"
                    WriteINI "Other", "IdleBans", "Y"
                End If
                
            Case "off"
                BotVars.IB_On = BFALSE
                
                tmpBuf = "IdleBans deactivated."
                
                WriteINI "Other", "IdleBans", "N"
            
            Case "wait", "delay"
                strArray() = Split(msgData, " ")
            
                If (StrictIsNumeric(strArray(1))) Then
                    BotVars.IB_Wait = CInt(strArray(1))
                    
                    tmpBuf = "IdleBan delay set to " & BotVars.IB_Wait & "."
                    
                    WriteINI "Other", "IdleBanDelay", CInt(strArray(1))
                Else
                    tmpBuf = "IdleBan delays require a numeric value."
                End If
                
            Case "kick"
                strArray() = Split(msgData, " ")
            
                If (UBound(strArray) > 1) Then
                    Select Case (LCase$(strArray(1)))
                        Case "on"
                            tmpBuf = "Idle users will now be kicked instead of banned."
                            
                            WriteINI "Other", "KickIdle", "Y"
                            
                            BotVars.IB_Kick = True
                            
                        Case "off"
                            tmpBuf = "Idle users will now be banned instead of kicked."
                            
                            WriteINI "Other", "KickIdle", "N"
                            
                            BotVars.IB_Kick = False
                            
                        Case Else
                            tmpBuf = "Unknown idle kick setting."
                    End Select
                Else
                    tmpBuf = "Not enough arguments were supplied."
                End If
                
            Case "status"
                If (BotVars.IB_On = BTRUE) Then
                    tmpBuf = IIf(BotVars.IB_Kick, "Kicking", "Banning") & _
                        " users who are idle for " & BotVars.IB_Wait & "+ seconds."
                Else
                    tmpBuf = "IdleBans are disabled."
                End If
                
            Case Else
                tmpBuf = "Invalid IdleBan command."
        End Select
    Else
        tmpBuf = "Invalid IdleBan command arguments."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdleBans

' handle chpw command
Private Function OnChPw(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String
    Dim tmpBuf     As String ' temporary output buffer
    
    strArray = Split(msgData, " ")
    
    If (UBound(strArray) > 0) Then
        Select Case (strArray(0))
            Case "on", "set"
                BotVars.ChannelPassword = strArray(2)
                
                If (BotVars.ChannelPasswordDelay < 1) Then
                    BotVars.ChannelPasswordDelay = 30
                    
                    tmpBuf = "Channel password protection enabled, delay set to " & _
                        BotVars.ChannelPasswordDelay & "."
                Else
                    tmpBuf = "Channel password protection enabled."
                End If
                
            Case "time", "delay", "wait"
                If (StrictIsNumeric(strArray(1))) Then
                    If (Val(strArray(1)) < 256) Then
                        BotVars.ChannelPasswordDelay = CByte(strArray(1))
                        
                        tmpBuf = "Channel password delay set to " & strArray(1) & "."
                    Else
                        tmpBuf = "Channel password delays cannot be more than 255 seconds."
                    End If
                Else
                    tmpBuf = "Time setting requires a numeric value."
                End If
                
            Case "off", "kill", "clear"
                BotVars.ChannelPassword = vbNullString
                
                BotVars.ChannelPasswordDelay = 0
                
                tmpBuf = "Channel password protection disabled."
                
            Case "info", "status"
                If ((BotVars.ChannelPassword = vbNullString) Or _
                    (BotVars.ChannelPasswordDelay = 0)) Then
                    
                    tmpBuf = "Channel password protection is disabled."
                Else
                    tmpBuf = "Channel password protection is enabled. Password [" & BotVars.ChannelPassword & "], Delay [" & _
                        BotVars.ChannelPasswordDelay & "]."
                End If
                
            Case Else
                tmpBuf = "Unknown channel password command."
        End Select
    Else
        tmpBuf = "Error setting channel password."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnChPw

' handle join command
Private Function OnJoin(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (LenB(msgData) > 0) Then
        AddQ "/join " & msgData
    Else
        tmpBuf = "Join what channel?"
    End If
End Function ' end function OnJoin

' handle sethome command
Private Function OnSetHome(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    WriteINI "Main", "HomeChan", msgData
    
    BotVars.HomeChannel = msgData
    
    tmpBuf = "Home channel set to [ " & msgData & " ]"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetHome

' handle resign command
Private Function OnResign(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    AddQ "/resign", 1
End Function ' end function OnResign

' handle clearbanlist
Private Function OnClearBanList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

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
            
        Case "off"
            BotVars.KickOnYell = 0
            
            tmpBuf = "Kick-on-yell disabled."
            
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
    
    AddQ "/join " & CurrentUsername & " Rejoin", 1
    AddQ "/join " & gChannel.Current, 1
End Function ' end function OnRejoin

' handle plugban command
Private Function OnPlugBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    Select Case (LCase$(msgData))
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
                            AddQ "/ban " & IIf(Dii, "*", "") & .Username & " PlugBan", 1
                        End If
                    End With
                Next i
            End If
            
        Case "off"
            If (BotVars.PlugBan) Then
                BotVars.PlugBan = False
                
                tmpBuf = "PlugBan deactivated."
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
    Dim tmpCount As Integer
    Dim BanCount As Integer
    Dim i        As Integer
    
    ReDim Preserve tmpBuf(0)
    
    tmpBuf(tmpCount) = "Clientbans: "

    For i = LBound(ClientBans()) To UBound(ClientBans())
        If (ClientBans(i) <> vbNullString) Then
            tmpBuf(tmpCount) = tmpBuf(tmpCount) & ", " & ClientBans(i)
            
            If (Len(tmpBuf(tmpCount)) > 90) Then
                ' increase array size
                ReDim Preserve tmpBuf(tmpCount + 1)
                
                tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1)) & _
                    " [more]"
                    
                ' increment counter
                tmpCount = (tmpCount + 1)
            End If
            
            ' increment counter
            BanCount = (BanCount + 1)
        End If
    Next i

    If (BanCount = 0) Then
        tmpBuf(tmpCount) = "There are currently no client bans."
    Else
        tmpBuf(tmpCount) = Replace(tmpBuf(tmpCount), " , ", Space(1))
    End If
    
    ' return message
    cmdRet() = tmpBuf()
End Function ' end function OnClientBans

' handle setvol command
Private Function OnSetVol(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long

    If (Not (BotVars.DisableMP3Commands)) Then
        If (StrictIsNumeric(msgData)) Then
            hWndWA = GetWinamphWnd()
            
            If (hWndWA = 0) Then
                tmpBuf = "Winamp is not loaded."
            End If
            
            If (CInt(msgData) > 100) Then
                msgData = 100
            End If
            
            Call SendMessage(hWndWA, WM_WA_IPC, 2.55 * CInt(msgData), 122)
            
            tmpBuf = "Volume set to " & msgData & "%."
        Else
            tmpBuf = "Invalid volume level (0-100)."
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetVol

' handle cadd command
Private Function OnCAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim cBans      As String
    Dim strArray() As String
    Dim i          As Integer

    If (Len(msgData) > 0) Then
        ' grab client bans from file
        cBans = UCase$(ReadCFG("Other", "ClientBans"))
        
        ' postfix new ban(s) to current listing
        cBans = cBans & Space(1) & UCase$(msgData)
        
        ' write client bans to file
        WriteINI "Other", "ClientBans", UCase$(cBans)
        
        ' write client bans to memory
        If (InStr(1, msgData, Space(1), vbBinaryCompare) = 0) Then
            ReDim Preserve ClientBans(0 To UBound(ClientBans) + 1)
                
            ClientBans(UBound(ClientBans)) = UCase$(msgData)
        Else
            strArray() = Split(msgData, " ")
            
            For i = LBound(strArray) To UBound(strArray)
                ReDim Preserve ClientBans(0 To UBound(ClientBans) + 1)
                
                ClientBans(UBound(ClientBans)) = UCase$(strArray(i))
            Next i
        End If
        
        tmpBuf = "Added clientban(s): " & UCase$(msgData)
    Else
        tmpBuf = "You must enter a client to ban."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCAdd

' handle cdel command
Private Function OnCDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim i      As Integer
    Dim cBans  As String
    
    For i = LBound(ClientBans) To UBound(ClientBans)
        cBans = cBans & UCase$(ClientBans(i)) & " "
    Next i
    
    If (InStr(1, cBans, msgData, vbBinaryCompare) <> 0) Then
        cBans = Replace(cBans, msgData, vbNullString)
        
        WriteINI "Other", "ClientBans", Replace(cBans, "  ", vbNullString)
        
        ClientBans() = Split(ReadCFG("Other", "ClientBans"), " ")
        
        If (UBound(ClientBans) = -1) Then
            ReDim ClientBans(0)
        End If
        
        tmpBuf = "Clientban """ & UCase$(msgData) & """ deleted."
    Else
        tmpBuf = "Client is not banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCDel

' handle banned command
Private Function OnBanned(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf() As String ' temporary output buffer
    Dim tmpCount As Integer
    Dim BanCount As Integer
    Dim i        As Integer
    
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

    If (Left$(msgData, 2)) = "on" Then
        BotVars.IPBans = True
        
        WriteINI "Other", "IPBans", "Y"
        
        tmpBuf = "IPBanning activated."
        
        If ((MyFlags = 2) Or (MyFlags = 18)) Then
            For i = 1 To colUsersInChannel.Count
                Select Case colUsersInChannel.Item(i).Flags
                    Case 20, 30, 32, 48
                        AddQ "/ban " & IIf(Dii, "*", "") & colUsersInChannel.Item(i).Username & _
                            " IPBanned.", 1
                End Select
            Next i
        End If
    ElseIf (Left$(msgData, 3) = "off") Then
        BotVars.IPBans = False
        
        WriteINI "Other", "IPBans", "N"
        
        tmpBuf = "IPBanning deactivated."
        
    ElseIf (Left$(msgData, 6) = "status") Then
        If (BotVars.IPBans) Then
            tmpBuf = "IPBanning is currently active."
        Else
            tmpBuf = "IPBanning is currently disabled."
        End If
    Else
        tmpBuf = "Unrecognized IPBan command. Use 'on', 'off' or 'status'."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIPBans

' handle ipban command
Private Function OnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim gAcc     As udtGetAccessResponse
    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer

    msgData = StripInvalidNameChars(msgData)
    dbAccess = GetAccess(Username)
    
    If (Len(msgData) > 0) Then
        If (InStr(1, msgData, "@") > 0) Then
            msgData = StripRealm(msgData)
        End If
        
        If (dbAccess.Access < 101) Then
            If ((GetSafelist(msgData)) Or (GetSafelist(msgData))) Then
                ' return message
                cmdRet(0) = "That user is safelisted."
                
                Exit Function
            End If
        End If
        
        gAcc = GetAccess(msgData)
        
        If ((gAcc.Access >= dbAccess.Access) Or _
            ((InStr(gAcc.Flags, "A") > 0) And (dbAccess.Access < 101))) Then

            tmpBuf = "You do not have enough access to do that."
        Else
            AddQ "/squelch " & IIf(Dii, "*", "") & msgData, 1
        
            tmpBuf = "User " & Chr(34) & msgData & Chr(34) & " IPBanned."
        End If
    Else
        ' return message
        tmpBuf = "You do not have enough access to do that."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIPBan

' handle unipban command
Private Function OnUnIPBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (Len(msgData) > 0) Then
        AddQ "/unsquelch " & IIf(Dii, "*", "") & msgData, 1
        AddQ "/unban " & IIf(Dii, "*", "") & msgData, 1
        
        tmpBuf = "User " & Chr(34) & msgData & Chr(34) & " Un-IPBanned."
    Else
        tmpBuf = "Un-IPBan who?"
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnIPBan

' handle designate command
Private Function OnDesignate(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        If (((MyFlags) And (&H2)) = &H2) Then
            'diablo 2 handling
            If (Dii = True) Then
                If (Not (Mid$(msgData, 1, 1) = "*")) Then
                    msgData = "*" & msgData
                End If
            End If
            
            AddQ "/designate " & msgData, 1
            
            tmpBuf = "I have designated [ " & msgData & " ]"
        Else
            tmpBuf = "The bot does not have ops."
        End If
    Else
        tmpBuf = "Designate who?"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDesignate

' handle shuffle command
Private Function OnShuffle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long
    
    If (Not (BotVars.DisableMP3Commands)) Then
        tmpBuf = "Winamp's Shuffle feature has been toggled."
        
        hWndWA = GetWinamphWnd()
        
        If (hWndWA = 0) Then
            tmpBuf = "Winamp is not loaded."
        Else
            Call SendMessage(hWndWA, WM_COMMAND, WA_TOGGLESHUFFLE, 0)
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShuffle

' handle repeat command
Private Function OnRepeat(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim hWndWA As Long
    Dim tmpBuf As String ' temporary output buffer
    
    If (Not (BotVars.DisableMP3Commands)) Then
        tmpBuf = "Winamp's Repeat feature has been toggled."
        
        hWndWA = GetWinamphWnd()
        
        If (hWndWA = 0) Then
            tmpBuf = "Winamp is not loaded."
        Else
            Call SendMessage(hWndWA, WM_COMMAND, WA_TOGGLEREPEAT, 0)
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

    If (Not (BotVars.DisableMP3Commands)) Then
        If (iTunesReady) Then
            iTunesNext
            
            tmpBuf = "Skipped forwards."
        Else
            hWndWA = GetWinamphWnd()
            
            If (hWndWA = 0) Then
               tmpBuf = "Winamp is not loaded."
            End If
        
            Call SendMessage(hWndWA, WM_COMMAND, WA_NEXTTRACK, 0)
            
            tmpBuf = "Skipped forwards."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnNext

' handle prev command
Private Function OnPrev(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long

    If (Not (BotVars.DisableMP3Commands)) Then
        If (iTunesReady) Then
            iTunesBack
            
            tmpBuf = "Skipped backwards."
        Else
            hWndWA = GetWinamphWnd()
            
            If (hWndWA = 0) Then
               tmpBuf = "Winamp is not loaded."
            End If
            
            Call SendMessage(hWndWA, WM_COMMAND, WA_PREVTRACK, 0)
            
            tmpBuf = "Skipped backwards."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPrev

' handle protect command
Private Function OnProtect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    Select Case (LCase$(msgData))
        Case "on"
            If ((MyFlags = 2) Or (MyFlags = 18)) Then
                Protect = True
                
                tmpBuf = "Lockdown activated by " & Username & "."
                
                WildCardBan "*", ProtectMsg, 1
                
                WriteINI "Main", "Protect", "Y"
            Else
                tmpBuf = "The bot does not have ops."
            End If
        
        Case "off"
            If (Protect) Then
                Protect = False
                
                tmpBuf = "Lockdown deactivated."
                
                WriteINI "Main", "Protect", "N"
            Else
                tmpBuf = "Protection was not enabled."
            End If
            
        Case "status"
            Select Case Protect
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
    
    Dim tmpBuf      As String ' temporary output buffer
    Dim WhisperCmds As Boolean
    
    If (BotVars.WhisperCmds) Then
        BotVars.WhisperCmds = False
        
        WhisperCmds = False
        
        WriteINI "Main", "WhisperBack", "N"
        
        tmpBuf = "Command responses will be displayed publicly."
    Else
        BotVars.WhisperCmds = True
        
        WhisperCmds = True
            
        WriteINI "Main", "WhisperBack", "Y"
        
        tmpBuf = "Command responses will be whispered back."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhisperCmds

' handle stop command
Private Function OnStop(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long
    
    If (Not (BotVars.DisableMP3Commands)) Then
        If (iTunesReady) Then
            iTunesStop
            
            tmpBuf = "iTunes playback stopped."
        Else
            hWndWA = GetWinamphWnd()
            
            If (hWndWA = 0) Then
               tmpBuf = "Winamp is not loaded."
            End If
            
            Call SendMessage(hWndWA, WM_COMMAND, WA_STOP, 0)
            
            tmpBuf = "Stopped play."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnStop

' handle play command
Private Function OnPlay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim hWndWA  As Long
    Dim Track   As Long
    Dim iWinamp As Long
    
    If (Len(msgData) > 0) Then
        If (Not (BotVars.DisableMP3Commands)) Then
            If (iTunesReady) Then
                iTunesPlayFile Mid$(msgData, 7)
                
                tmpBuf = "Attempted to play the specified filepath."
            Else
                hWndWA = GetWinamphWnd()
                
                If (hWndWA = 0) Then
                    tmpBuf = "Winamp is stopped, or isn't running."
                End If
                
                If (StrictIsNumeric(msgData)) Then
                    Track = CInt(msgData)
                    
                    Call SendMessage(hWndWA, WM_COMMAND, WA_STOP, 0)
                    Call SendMessage(hWndWA, WM_USER, Track - 1, 121)
                    Call SendMessage(hWndWA, WM_COMMAND, WA_PLAY, 0)
                    
                    tmpBuf = "Skipped to track " & Track & "."
                Else
                    WinampJumpToFile msgData
                End If
            End If
        End If
    Else
        If (Not (BotVars.DisableMP3Commands)) Then
            If (iTunesReady) Then
                iTunesPlay
                
                tmpBuf = "iTunes playback started."
            Else
                hWndWA = GetWinamphWnd()
        
                If (hWndWA = 0) Then
                   tmpBuf = "Winamp is not loaded."
                End If
        
                Call SendMessage(hWndWA, WM_COMMAND, WA_PLAY, 0)
        
                tmpBuf = "Skipped backwards."
        
                If (iWinamp = 0) Then
                    tmpBuf = "Play started."
                Else
                    tmpBuf = "Error sending your command to Winamp. Make sure it's running."
                End If
            End If
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPlay

' handle useitunes command
Private Function OnUseiTunes(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (iTunesReady) Then
        tmpBuf = "iTunes is already ready."
    Else
        If (InitITunes) Then
            tmpBuf = "iTunes is ready."
        Else
            tmpBuf = "Error launching iTunes."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUseiTunes

' handle usewinamp command
Private Function OnUseWinamp(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (iTunesReady) Then
        tmpBuf = "Returning to Winamp control."
        
        iTunesUnready
    Else
        tmpBuf = "iTunes was not ready."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUseWinamp

' handle pause command
Private Function OnPause(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim hWndWA As Long
    
    If (Not (BotVars.DisableMP3Commands)) Then
        If (iTunesReady) Then
            iTunesPause
            
            tmpBuf = "Pause toggled."
        Else
            hWndWA = GetWinamphWnd()
            
            If (hWndWA = 0) Then
               tmpBuf = "Winamp is not loaded."
            End If
            
            Call SendMessage(hWndWA, WM_COMMAND, WA_PAUSE, 0)
            
            tmpBuf = "Paused/resumed play."
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPause

' handle fos command
Private Function OnFos(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim hWndWA As Long
    Dim tmpBuf As String ' temporary output buffer

   If (Not (BotVars.DisableMP3Commands)) Then
        hWndWA = GetWinamphWnd()
        
        If (hWndWA = 0) Then
           tmpBuf = "Winamp is not loaded."
        End If
        
        Call SendMessage(hWndWA, WM_COMMAND, WA_FADEOUTSTOP, 0)
        
        tmpBuf = "Fade-out stop."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFos

' handle rem command
Private Function OnRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendz
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim u        As String
    Dim tmpBuf   As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)
    
    u = Right(msgData, (Len(msgData) - 5))
    
    If (GetAccess(u).Access >= dbAccess.Access) Then
        tmpBuf = "That user has higher or equal access."
    End If
    
    If InStr(1, GetAccess(u).Flags, "L") > 0 Then
        'If ((InStr(1, GetAccess(Username).Flags, "A") = 0) And _
        '   (GetAccess(Username).Access < 100) And (Not (InBot))) Then
        '
        '    tmpBuf = "That user is Locked."
        'End If
    End If
    
    tmpBuf = RemoveItem(u, "users")
    tmpBuf = Replace(tmpBuf, "%msgex%", "userlist entry")
    
    If (InStr(tmpBuf, "Successfully")) Then
        If (BotVars.LogDBActions) Then
            Call LogDBAction(RemEntry, Username, u, msgData)
        End If
    End If
    
    Call LoadDatabase
    
    ' return message
    cmdRet(0) = tmpBuf

sendz:
    tmpBuf = "Remove what user?"
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnRem

' handle reconnect command
Private Function OnReconnect(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    If (g_Online) Then
        BotVars.HomeChannel = gChannel.Current
        
        Call frmChat.DoDisconnect
        
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Reconnecting by command, please wait..."
        
        Pause 1
        
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "You must be online to reconnect. Try connecting first."
    End If
End Function ' end function OnReconnect

' handle unigpriv command
Private Function OnUnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    AddQ "/o unigpriv", 1
    
    tmpBuf = "Recieving text from non-friends."
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnIgPriv

' handle igpriv command
Private Function OnIgPriv(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    AddQ "/o igpriv", 1
    
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
        WriteINI "BlockList", "Total", "Total=0", "filters.ini"
        
        i = 0
    End If
    
    WriteINI "BlockList", "Filter" & (i + 1), u, "filters.ini"
    WriteINI "BlockList", "Total", i + 1, "filters.ini"
    
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
        WriteINI "Main", "IdleWait", 2 * Int(u)
        tmpBuf = "Idle wait time set to " & Int(u) & " minutes."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdleTime

' handle idle command
Private Function OnIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo IdleError
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
        
    u = Split(msgData, " ")(1)
    
    If LCase$(u) = "on" Then
        WriteINI "Main", "Idles", "Y"
        tmpBuf = "Idles activated."
    ElseIf LCase$(u) = "off" Then
        WriteINI "Main", "Idles", "N"
        tmpBuf = "Idles deactivated."
    ElseIf LCase$(u) = "kick" Then
        u = Split(msgData, " ")(2)
        
        If LCase$(u) = "on" Then
            BotVars.IB_Kick = True
            tmpBuf = "Idle kick is now enabled."
        ElseIf LCase$(u) = "off" Then
            BotVars.IB_Kick = False
            tmpBuf = "Idle kick disabled."
        Else
            tmpBuf = "Unknown idle kick command."
        End If
    Else
        GoTo IdleError
    End If
    
IdleError:
    tmpBuf = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIdle

' handle shitdel command
Private Function OnShitDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo IdleError23
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, Len(msgData) - 9)
    
IdleError23:

    If ((MyFlags = 2) Or (MyFlags = 18)) Then
        AddQ "/unban " & IIf(Dii, "*", "") & u, 1
    End If
    
    tmpBuf = RemoveItem(u, "autobans")
    tmpBuf = Replace(tmpBuf, "%msgex%", "shitlist")
    
    If (InStr(tmpBuf, "Successfully")) Then
        If (BotVars.LogDBActions) Then
            Call LogDBAction(RemEntry, Username, u, msgData)
        End If
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShitDel

' handle safedel command
Private Function OnSafeDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim b      As Boolean
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = msgData
        
    b = RemoveFromSafelist(u)
    
    If (b) Then
        tmpBuf = "That user has been removed from the safelist."
    Else
        tmpBuf = "That user is not safelisted, or there was an error removing them."
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSafeDel

' handle tagdel command
Private Function OnTagDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = msgData
    
    If (Len(u) > 0) Then
        tmpBuf = RemoveItem(u, "tagbans")
        tmpBuf = Replace(tmpBuf, "%msgex%", "tagban")
    Else
        tmpBuf = "Delete what tag?"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTagDel
        
' handle profile command
Private Function OnProfile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo Error614
    
    Dim u      As String
    Dim PPL    As Boolean
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, Len(msgData) - 9)
    PPL = True
    
    'If (BotVars.WhisperCmds Or WhisperedIn) And Not PublicOutput Then
    '    PPLRespondTo = Username
    'End If
    
    Call RequestProfile(u)

Error614:
    tmpBuf = "What profile would you like to look up?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnProfile

' handle setidle command
Private Function OnSetIdle(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo IdleError133
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, Len(msgData) - 9)
    
    If Left$(u, 1) = "/" Then
        u = " " & u
    End If
    
    WriteINI "Main", "IdleMsg", u
    
    tmpBuf = "Idle message set."
    
    
IdleError133:
    tmpBuf = "What do you want the idle message set to?"
        
' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetIdle

' handle idletype command
Private Function OnIdleType(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo IdleError2
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
        
    u = msgData
    
    If LCase$(u) = "msg" Or LCase$(u) = "message" Then
        WriteINI "Main", "IdleType", "msg"
        tmpBuf = "Idle type set to [ msg ]."
    ElseIf LCase$(u) = "quote" Or LCase$(u) = "quotes" Then
        WriteINI "Main", "IdleType", "quote"
        tmpBuf = "Idle type set to [ quote ]."
    ElseIf LCase$(u) = "uptime" Then
        WriteINI "Main", "IdleType", "uptime"
        tmpBuf = "Idle type set to [ uptime ]."
    ElseIf LCase$(u) = "mp3" Then
        WriteINI "Main", "IdleType", "mp3"
        tmpBuf = "Idle type set to [ mp3 ]."
    Else
        GoTo IdleError2
    End If
    
IdleError2:
    tmpBuf = "Error setting idle type. The types are [ message quote uptime mp3 ]."
    
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
        WriteINI "TextFilters", "Total", "Total=0", "filters.ini"
        
        i = 0
    End If
    
    WriteINI "TextFilters", "Filter" & (i + 1), u, "filters.ini"
    WriteINI "TextFilters", "Total", i + 1, "filters.ini"
    
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

    tmpBuf = "The bot's current trigger is " & Chr(34) & Space(1) & _
        BotVars.Trigger & Space(1) & Chr(34) & " (Alt + 0" & Asc(BotVars.Trigger) & ")"

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTrigger

' handle settrigger command
Private Function OnSetTrigger(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u          As String
    Dim tmpBuf     As String ' temporary output buffer
    Dim OldTrigger As String
    
    u = msgData
    
    If (Len(u) > 0) Then
        OldTrigger = u
        
        WriteINI "Main", "Trigger", u
    
        tmpBuf = "The new trigger is " & Chr(34) & OldTrigger & Chr(34) & "."
        
        BotVars.Trigger = u
    Else
        tmpBuf = "Change to what trigger?"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetTrigger

' handle levelban command
Private Function OnLevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo Error231
        
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
    
    If (StrictIsNumeric(Right(msgData, Len(msgData) - 10))) Then
        i = Val(Right(msgData, Len(msgData) - 10))
        
        If (i > 0) Then
            tmpBuf = "Banning Warcraft III users under level " & i & "."
            
            BotVars.BanUnderLevel = i
        Else
            tmpBuf = "Levelbans disabled."
            
            BotVars.BanUnderLevel = 0
        End If
    Else
        BotVars.BanUnderLevel = 0
        
        tmpBuf = "Levelbans disabled."
    End If
    
    WriteINI "Other", "BanUnderLevel", BotVars.BanUnderLevel
    
    'If (BotVars.BanUnderLevel = 0) Then
    '   tmpBuf = "Currently not banning Warcraft III users by level."
    'Else
    '   tmpBuf = "Currently banning Warcraft III users under level " & BotVars.BanUnderLevel & "."
    'End If
    
Error231:
    tmpBuf = "Error setting Levelban level."

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLevelBan

' handle d2levelban command
Private Function OnD2LevelBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo Error231
    
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer
        
    If (StrictIsNumeric(Right(msgData, Len(msgData) - 12))) Then
        i = Val(Right(msgData, Len(msgData) - 12))
        BotVars.BanD2UnderLevel = i
        
        If (i > 0) Then
            tmpBuf = "Banning Diablo II characters under level " & i & "."
            BotVars.BanD2UnderLevel = i
        Else
            tmpBuf = "Diablo II Levelbans disabled."
            BotVars.BanD2UnderLevel = 0
        End If
    Else
        tmpBuf = "Diablo II Levelbans disabled."
        BotVars.BanD2UnderLevel = 0
    End If
    
    WriteINI "Other", "BanD2UnderLevel", BotVars.BanD2UnderLevel

Error231:
    tmpBuf = "Error setting Levelban level."
        
    'If BotVars.BanD2UnderLevel = 0 Then
    '   tmpBuf = "Currently not banning Diablo II users by level."
    'Else
    '   tmpBuf = "Currently banning Diablo II users under level " & BotVars.BanD2UnderLevel & "."
    'End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnD2LevelBans

' handle phrasebans command
Private Function OnPhraseBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
  
    Dim tmpBuf As String ' temporary output buffer
  
    '### phrasebans on ###
 
    WriteINI "Other", "Phrasebans", "Y"
    
    Phrasebans = True
    
    tmpBuf = "Phrasebans activated."
    
    ' ### phrasebans off ###
    
    WriteINI "Other", "Phrasebans", "N"
    
    Phrasebans = False
    
    tmpBuf = "Phrasebans deactivated."
    
    ' ### phrasebans ###
    
    If (Phrasebans = True) Then
        tmpBuf = "Phrasebans are enabled."
    Else
        tmpBuf = "Phrasebans are disabled."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPhraseBans

' handle cbans command
Private Function OnCBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo cbansError
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    ' on/off/status
    u = Mid$(msgData, 8)
    
    Select Case u
        Case "on"
            tmpBuf = "ClientBans enabled."
            BotVars.ClientBans = True
            WriteINI "Other", "ClientBansOn", "Y"
            
        Case "off"
            tmpBuf = "ClientBans disabled."
            BotVars.ClientBans = False
            WriteINI "Other", "ClientBansOn", "N"
            
        Case "status"
            tmpBuf = "ClientBans are currently " & IIf(BotVars.ClientBans, "enabled.", "disabled.")
            
        Case Else
            GoTo cbansError
    End Select
    
cbansError:
    tmpBuf = "What do you want to do to your ClientBans?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCBans

' handle mimic command
Private Function OnMimic(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = msgData
    
    If (Len(u) > 0) Then
        Mimic = LCase$(u)
        tmpBuf = "Mimicking [ " & u & " ]"
    Else
        tmpBuf = "Mimic who?"
    End If
  
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMimic

' handle nomimic command
Private Function OnNoMimic(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    Mimic = vbNullString
    tmpBuf = "Mimic off."

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnNoMimic

' handle setpmsg command
Private Function OnSetPMsg(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = Right(msgData, Len(msgData) - 9)
    ProtectMsg = u
    WriteINI "Other", "ProtectMsg", u
    tmpBuf = "Channel protection message set."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetPMsg

' handle setcmdaccess command
Private Function OnSetCmdAccess(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim ccOut    As udtCustomCommandData
    'Dim dbAccess As udtGetAccessResponse
     
    Dim c        As Integer
    Dim iWinamp  As Long
    Dim tmpBuf   As String ' temporary output buffer
    Dim f        As Integer
    Dim Track    As Long
    Dim z        As String
    
    dbAccess = GetAccess(Username)
    c = 1
        
    If (Not (StrictIsNumeric(GetStringChunk(msgData, 3)))) Then
        tmpBuf = "You must specify a numeric value to change a command's required access."
    End If
    
    iWinamp = Val(GetStringChunk(msgData, 3))
    z = GetStringChunk(msgData, 2)
    
    If (iWinamp > 999) Then
        tmpBuf = "Your new required access must be between 0 and 999."
    End If
    
    Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
        While ((Track = 0) And (Not (EOF(f))))
            Get #f, c, ccOut
                        
            If (StrComp(Trim(ccOut.Query), z) = 0) Then
                If ((dbAccess.Access < 100) And (InStr(dbAccess.Flags, "A") = 0) And _
                    (ccOut.reqAccess >= dbAccess.Access) And (ccOut.reqAccess > iWinamp)) Then
            
                    tmpBuf = "You cannot decrease the required access on a command" & _
                        " without 100 access or the A flag."
                Else
                    ccOut.reqAccess = iWinamp
            
                    Put #f, c, ccOut
            
                    tmpBuf = "Command modified successfully."
               End If
            
               Track = 1
            End If
            
            c = (c + 1)
        Wend
    Close #f
    
    If (Track = 0) Then
        tmpBuf = "That command was not found."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSetCmdAccess

' handle cmdadd command
Private Function OnCmdAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo cmdAddError
    
    'Dim dbAccess   As udtGetAccessResponse
    Dim gAcc       As udtGetAccessResponse
    Dim ccOut      As udtCustomCommandData
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim strArray() As String
    Dim f          As Integer
    Dim c          As Integer
    
    dbAccess = GetAccess(Username)
    
    If ((InStr(1, msgData, "/add ", vbTextCompare) > 0) Or _
        (InStr(1, msgData, "/rem ", vbTextCompare) > 0) Or _
        (InStr(1, msgData, "/set ", vbTextCompare) > 0)) Then
            
            tmpBuf = "You cannot use '/add' or '/rem' in a custom command."
    End If
    
    If (InStr(1, msgData, "/quit", vbTextCompare) > 0) Then
        tmpBuf = "You cannot use '/quit' in a custom command."
    End If
    
    If ((InStr(1, msgData, "/shitadd", vbTextCompare) > 0) Or _
        (InStr(1, msgData, "/pban", vbTextCompare) > 0) Or _
        (InStr(1, msgData, "/shitlist", vbTextCompare) > 0)) Then
            
        If ((dbAccess.Access < 100) And (InStr(dbAccess.Flags, "A") = 0)) Then
            tmpBuf = "Shitlisting users through custom commands requires 100+ access."
        End If
    End If
        
    gAcc.Access = 1000
    gAcc.Flags = "A"
    
    'If (ProcessCC(Username, gAcc, "/" & Split(msgData, " ")(2), False, False, True)) Then
    '    tmpBuf = "A command by that name already exists."
    'End If
    
    gAcc.Access = 0
    gAcc.Flags = ""
            
    '0      1  2    3
    'cmdadd 10 fdsa actions
    strArray() = Split(msgData, " ", 4)
    
    If (Not (StrictIsNumeric(strArray(1)))) Then
        tmpBuf = "Command format error."
    End If
    
    If (UBound(strArray) > 2) Then
        If (Len(strArray(3)) < 1) Then
            tmpBuf = "Your command's actions cannot be blank."
        End If
    Else
        tmpBuf = "Your command's actions cannot be blank."
    End If
    
    ccOut.Query = strArray(2)
    ccOut.reqAccess = Int(strArray(1))
    
    ccOut.Action = strArray(3)
    
    Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
        c = LOF(f) \ LenB(ccOut)
        
        If ((LOF(f)) Mod (LenB(ccOut) <> 0)) Then
            c = c + 1
        End If
        
        If (c = 0) Then
            c = 1
        End If
        
        Put #f, c + 1, ccOut
    Close #f
    
    tmpBuf = "Command " & Chr(34) & strArray(2) & Chr(34) & " added."
    
cmdAddError:
    tmpBuf = "Error adding your command."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCmdAdd

' handle cmddel command
Private Function OnCmdDel(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim ccOut  As udtCustomCommandData
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    Dim c      As Integer
    Dim f      As Integer
    
    u = Right(msgData, Len(msgData) - 8)
        
    If Dir$((GetFilePath("commands.dat"))) = vbNullString Then
        tmpBuf = "No commands list exists."
    End If
    
    Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
        If (LOF(f) < 2) Then
            Exit Function
        End If
        
        c = (LOF(f) \ LenB(ccOut))
        
        If (LOF(f) Mod LenB(ccOut) <> 0) Then
            c = c + 1
        End If
        
        For c = 1 To c
            Get #f, c, ccOut
            
            If (ccOut.reqAccess > 1000) Then
                GoTo NextRecord2
            End If

            If (StrComp(LCase$(RTrim(ccOut.Query)), u, vbTextCompare) = 0) Then
                ccOut.reqAccess = 9999
                Put #f, c, ccOut
                tmpBuf = "Command " & Chr(34) & u & Chr(34) & " deleted."
            End If
            
NextRecord2:
        Next c
    Close #f
    
    If (LenB(tmpBuf) = 0) Then
        tmpBuf = "No such command exists."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCmdDel

' handle cmdlist command
Private Function OnCmdList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim ccOut    As udtCustomCommandData
    
    Dim f        As Integer
    Dim tmpBuf   As String ' temporary output buffer
    Dim n        As Integer
    Dim response As String
    Dim Track    As Long
    Dim c        As Integer
    
    f = FreeFile
    
    If (Dir$((GetFilePath("commands.dat"))) = vbNullString) Then
        tmpBuf = "No custom commands available."
    End If
    
    Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
    
        If (LOF(f) < 2) Then
            tmpBuf = "No custom commands available."
        End If
        
        n = (LOF(f) \ LenB(ccOut))
        
        If ((LOF(f)) Mod (LenB(ccOut) <> 0)) Then
            n = n + 1
        End If
        
        response = "Found commands: "
        
        For c = 1 To n
            Get #f, c, ccOut
            
            If ((ccOut.reqAccess < 1001) And _
                (Len(RTrim(ccOut.Query)) > 0)) Then
            
                response = response & ", " & RTrim(KillNull(ccOut.Query)) & _
                    " [" & ccOut.reqAccess & "]"
                
                If (Len(response) > 100) Then
                    response = Replace(response, ", ", " ")
                    
                    If (c < n) Then
                        response = response & " [more]"
                    End If
                    
                    'If WhisperCmds And Not InBot Then
                    '    If Not Dii Then AddQ "/w " & Username & Space(1) & response Else AddQ "/w *" & Username & response
                    'ElseIf InBot And Not PublicOutput Then
                    '    frmChat.AddChat RTBColors.ConsoleText, response
                    'Else
                    '    AddQ response
                    'End If
                    
                    response = "Found commands: "
                End If
                
                Track = 1
            
            End If
        Next c
        
    Close #f
    
    tmpBuf = Replace(response, ", ", " ") & "."
    
    If (StrComp(tmpBuf, "Found commands: .", vbBinaryCompare) = 0) Then
        If (Track = 0) Then
            tmpBuf = "No custom commands available."
        Else
            Exit Function
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCmdList

' handle phrases command
Private Function OnPhrases(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf   As String ' temporary output buffer
    Dim response As String
    Dim c        As Integer

    If UBound(Phrases) = 0 And Phrases(0) = vbNullString Then
        tmpBuf = "There are no phrasebans."
    End If
    
    response = "Phraseban(s): "
    
    For c = LBound(Phrases) To UBound(Phrases)
        If Phrases(c) <> " " And Phrases(c) <> vbNullString Then
            response = response & ", " & Phrases(c)
            If Len(response) > 89 Then
                response = Replace(response, ", ", " ") & " [more]"
                'If WhisperCmds And Not InBot Then
                '    AddQ "/w " & Username & Space(1) & response
                'ElseIf InBot = True Then
                '    frmChat.AddChat RTBColors.ConsoleText, response
                'Else
                '    AddQ response
                'End If
                response = "Phraseban(s): "
            End If
        End If
    Next c
    
    tmpBuf = Replace(response, ", ", " ")
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnPhrases

' handle addphrase command
Private Function OnAddPhrase(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer
    
    Dim c      As Integer
    Dim tmpBuf As String ' temporary output buffer
    Dim u      As String
    Dim i      As Integer
    
    If (InStr(1, msgData, "addphrase ", vbTextCompare) = 0) Then
        c = 6
    Else
        c = 11
    End If
    
    u = Right(msgData, Len(msgData) - c)
    
    For i = LBound(Phrases) To UBound(Phrases)
        If (StrComp(LCase$(u), LCase$(Phrases(i)), vbTextCompare) = 0) Then
            tmpBuf = "That phrase is already banned."
        End If
    Next i

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
    
    If (InStr(1, msgData, "delphrase ", vbTextCompare) = 0) Then
        c = 6
    Else
        c = 11
    End If
    
    u = Right(msgData, Len(msgData) - c)
    
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
        tmpBuf = "That phrase is not banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDelPhrase

' handle tagban command
Private Function OnTagBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f      As Integer

    Dim tmpBuf As String ' temporary output buffer
    Dim u      As String
    
    If (Len(msgData) > 0) Then
        If (Len(GetTagbans(u)) > 1) Then
            tmpBuf = "That tag is covered by an existing tagban."
        Else
            Dim saCmdRet() As String
            
            ' declare index zero of array
            ReDim Preserve saCmdRet(0)
        
            If ((InStr(1, u, "*", vbTextCompare) = 0) And (Len(u) > 4)) Then
                Call OnShitAdd(Username, dbAccess, msgData, InBot, saCmdRet())
            End If
            
            If (Dir$(GetFilePath("tagbans.txt")) = vbNullString) Then
                Open (GetFilePath("tagbans.txt")) For Output As #f
                Close #f
            End If
            
            Open (GetFilePath("tagbans.txt")) For Append As #f
                Print #f, u & vbCrLf
            Close #f
            
            If (InStr(u, " ") > 0) Then
                Call WildCardBan(u, Mid$(u, InStr(u, " ")), 1)
            Else
                Call WildCardBan(u, "Tagban: " & u, 1)
            End If
            
            tmpBuf = "Added tag " & Chr(34) & u & Chr(34) & " to the tagban list."
        End If
    Else
        tmpBuf = "What tag would you like to add?"
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTagBan

' handle fadd command
Private Function OnFAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        u = Right(msgData, (Len(msgData) - 6))
        AddQ "/f a " & u, 1
        tmpBuf = "Added user " & Chr(34) & u & Chr(34) & " to this account's friends list."
    Else
        tmpBuf = "Who do you want to add?"
    End If
        
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFAdd

' handle frem command
Private Function OnFRem(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    If (Len(msgData) > 0) Then
        u = Right(msgData, (Len(msgData) - 6))
        AddQ "/f r " & u, 1
        tmpBuf = "Removed user " & Chr(34) & u & Chr(34) & " from this account's friends list."
    Else
        tmpBuf = "Who do you want to remove?"
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFRem

' handle safelist command
Private Function OnSafeList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim i      As Integer

    If (colSafelist.Count = 0) Then
        tmpBuf = "There are no safelisted users or tags."
    Else
        tmpBuf = "Tags/users found: "

        For i = 1 To colSafelist.Count
            Debug.Print colSafelist.Item(i).Name
            
            tmpBuf = tmpBuf & ReversePrepareCheck(colSafelist.Item(i).Name)
            
            If (i < colSafelist.Count) Then
                tmpBuf = tmpBuf & ", "
            End If
            
            If (Len(tmpBuf) > 70) Then
                If (i < colSafelist.Count) Then
                    tmpBuf = tmpBuf & " [more]"
                End If
                
                'If WhisperCmds And Not InBot Then
                '    If Dii Then
                '        AddQ "/w *" & Username & Space(1) & tmpBuf
                '    Else
                '        AddQ "/w " & Username & Space(1) & tmpBuf
                '    End If
                'ElseIf InBot = True And Not PublicOutput Then
                '    frmChat.AddChat RTBColors.ConsoleText, tmpBuf
                'Else
                '    AddQ tmpBuf
                'End If
                
                tmpBuf = "Tags/users found: "
            End If
        Next i
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSafeList

' handle safeadd command
Private Function OnSafeAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim u      As String
    
    u = GetStringChunk(msgData, 2)
        
    tmpBuf = AddToSafelist(u, Username)
    
    If (LenB(tmpBuf) = 0) Then
        tmpBuf = "Added tag/user " & Chr(34) & u & Chr(34) & " to the safelist."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSafeAdd

' handle exile command
Private Function OnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim saCmdRet() As String
    Dim ibCmdRet() As String
    Dim u          As String
    Dim Y          As String
    
    ReDim Preserve saCmdRet(0)
    ReDim Preserve ibCmdRet(0)

    u = Mid$(msgData, 8)
    
    If InStr(1, u, " ") > 0 Then
        Y = Split(u, " ")(0)
    End If
    
    Call OnShitAdd(Username, dbAccess, msgData, InBot, saCmdRet())
    Call OnIPBan(Username, dbAccess, msgData, InBot, ibCmdRet())
End Function ' end function OnExile

' handle unexile command
Private Function OnUnExile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u          As String
    Dim sdCmdRet() As String
    Dim uiCmdRet() As String
    
    ' declare index zero of array
    ReDim Preserve sdCmdRet(0)
    ReDim Preserve uiCmdRet(0)

    u = Mid$(msgData, 10)
    
    If (InStr(1, u, " ") > 0) Then
        u = Split(u, " ")(0)
    End If
    
    Call OnShitDel(Username, dbAccess, msgData, InBot, sdCmdRet())
    Call OnUnignore(Username, dbAccess, msgData, InBot, uiCmdRet())
End Function ' end function OnUnExile

' handle shitlist command
Private Function OnShitList(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f          As Integer
    
    Dim strArray() As String
    Dim Y          As String
    Dim tmpBuf     As String ' temporary output buffer
    Dim i          As Integer
    Dim response   As String

    Y = GetFilePath("autobans.txt")
        
    If LenB(Dir$(Y)) = 0 Then
        tmpBuf = "No shitlist found."
    End If
    
    Open (Y) For Input As #f
    
        If (LOF(f) < 2) Then
            tmpBuf = "There are no shitlisted users."
        End If
        
        Do
            i = i + 1
            Line Input #f, response
            
            ReDim Preserve strArray(0 To i)
            
            If ((response <> vbNullString) And (Len(response) >= 2)) Then
                If (InStr(response, " ")) Then
                    strArray(i) = Mid$(response, 1, InStr(response, " ") - 1)
                Else
                    strArray(i) = response
                End If
            Else
                i = i - 1
            End If
        Loop While Not EOF(f)
        
    Close #f
    
    tmpBuf = "Tags/users found: "
    
    For i = (LBound(strArray) + 1) To UBound(strArray)
        tmpBuf = tmpBuf & strArray(i)
        
        If (i <> UBound(strArray)) Then
            tmpBuf = tmpBuf & ", "
        End If
        
        If Len(tmpBuf) > 70 Then
            If (i <> UBound(strArray)) Then
                tmpBuf = tmpBuf & " [more]"
            End If
            
            'If WhisperCmds And Not InBot Then
            '    If Dii Then AddQ "/w *" & Username & Space(1) & tmpBuf Else AddQ "/w " & Username & Space(1) & tmpBuf
            'ElseIf InBot = True And Not PublicOutput Then
            '    frmChat.AddChat RTBColors.ConsoleText, tmpBuf
            'Else
            '    AddQ tmpBuf
            'End If
            
            tmpBuf = "Tags/users found: "
        End If
    Next i
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShitList

' handle tagbans command
Private Function OnTagBans(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'If Dir$(GetFilePath("tagbans.txt")) = vbNullString Then
    '    strSend = "No tagbans list found."
    '    b = True
    '    GoTo Display
    'End If
    
    'Open (GetFilePath("tagbans.txt")) For Input As #f
    'If LOF(f) < 2 Then
    '    strSend = "No users are tagbanned."
    '    b = True
    '    GoTo Display
    'End If
    'Do
    '    i = i + 1
    '    Input #f, response
    '    ReDim Preserve strArray(0 To i)
    '    If response <> vbNullString And Len(response) >= 2 Then
    '        strArray(i) = response
    '    Else
    '        i = i - 1
    '    End If
    'Loop Until EOF(f)
    'strSend = "Tagbans found: "
    'For i = (LBound(strArray) + 1) To UBound(strArray)
    '    strSend = strSend & strArray(i) & ", "
    '    If Len(strSend) > 80 Then
    '        strSend = Left$(strSend, Len(strSend) - 2)
    '        strSend = strSend & " [more]"
    '        If WhisperCmds And Not InBot Then
    '            If Dii Then AddQ "/w *" & Username & Space(1) & strSend Else AddQ "/w " & Username & Space(1) & strSend
    '        ElseIf InBot = True And Not PublicOutput Then
    '            frmChat.AddChat RTBColors.ConsoleText, strSend
    '        Else
    '            AddQ strSend
    '        End If
    '        strSend = "Tagbans found: "
    '    End If
    'Next i
    'strSend = Left$(strSend, Len(strSend) - 2)
    'b = True
    'GoTo Display
End Function ' end function OnTagBans

' handle shitadd command
Private Function OnShitAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim f          As Integer
    
    Dim tmpBuf     As String ' temporary output buffer
    Dim i          As Integer
    Dim response   As String
    Dim strArray() As String

    If Dir$(GetFilePath("tagbans.txt")) = vbNullString Then
        tmpBuf = "No tagbans list found."
    End If
    
    Open (GetFilePath("tagbans.txt")) For Input As #f
    
        If (LOF(f) < 2) Then
            tmpBuf = "No users are tagbanned."
        End If
        
        Do
            i = i + 1
            
            Input #f, response
            
            ReDim Preserve strArray(0 To i)
            
            If ((response <> vbNullString) And (Len(response) >= 2)) Then
                strArray(i) = response
            Else
                i = i - 1
            End If
        Loop Until EOF(f)
        
    Close #f
    
    tmpBuf = "Tagbans found: "
    
    For i = (LBound(strArray) + 1) To UBound(strArray)
        tmpBuf = tmpBuf & strArray(i) & ", "
        
        If (Len(tmpBuf) > 80) Then
            tmpBuf = Left$(tmpBuf, Len(tmpBuf) - 2)
            tmpBuf = tmpBuf & " [more]"
            
            'If WhisperCmds And Not InBot Then
            '    If Dii Then AddQ "/w *" & Username & Space(1) & tmpBuf Else AddQ "/w " & Username & Space(1) & tmpBuf
            'ElseIf InBot = True And Not PublicOutput Then
            '    frmChat.AddChat RTBColors.ConsoleText, tmpBuf
            'Else
            '    AddQ tmpBuf
            'End If
            tmpBuf = "Tagbans found: "
        End If
    Next i
    
    tmpBuf = Left$(tmpBuf, Len(tmpBuf) - 2)
    
    ' return message
    cmdRet(0) = tmpBuf
        
End Function ' end function OnShitAdd

' handle dnd command
Private Function OnDND(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim DNDMsg As String
    
    If (Len(msgData) <= 5) Then
        AddQ "/dnd", 1
    End If
    
    DNDMsg = Right(msgData, (Len(msgData) - 5))
    
    AddQ "/dnd " & DNDMsg, 1
End Function ' end function OnDND

' handle bancount command
Private Function OnBanCount(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    If (BanCount = 0) Then
        tmpBuf = "No users have been banned since I joined this channel."
    Else
        tmpBuf = "Since I joined this channel, " & BanCount & " user(s) have been banned."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBanCount

' handle tagcheck command
Private Function OnTagCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Y      As String
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, Len(msgData) - 10)
    Y = GetTagbans(u)
    
    If (Len(Y) < 2) Then
        tmpBuf = "That user matches no tagbans."
    Else
        tmpBuf = "That user matches the following tagban(s): " & Y
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTagCheck

' handle slcheck command
Private Function OnSLCheck(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim gAcc   As udtGetAccessResponse
    
    Dim Y      As String
    Dim Track  As Long
    Dim tmpBuf As String ' temporary output buffer

    Y = GetStringChunk(msgData, 2)
            
    If (LenB(Y) > 0) Then
        tmpBuf = "That user "
        
        gAcc = GetAccess(Y)
        
        If (InStr(gAcc.Flags, "B") > 0) Then
            tmpBuf = tmpBuf & "has 'B' in their flags"
            Track = 1
        End If
        
        If (LenB(GetShitlist(Y))) Then
            If (Track = 1) Then
                tmpBuf = tmpBuf & " and "
            End If
            
            tmpBuf = tmpBuf & "is on the bot's shitlist"
            
            Track = 2
        End If
        
        If (Track > 0) Then
            tmpBuf = tmpBuf & "."
        Else
            tmpBuf = "That user is not shitlisted and does not have 'B' in their flags."
        End If
    Else
        tmpBuf = "Please specify a username to check."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSLCheck

' handle readfile command
Private Function OnReadFile(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim Y      As String
    Dim tmpBuf As String ' temporary output buffer
    Dim i      As Integer
    
    u = Trim$(Right(msgData, Len(msgData) - 10))
    
    While ((Mid$(u, 1, 1) = ".") And (Len(u) > 1))
        u = Mid$(u, 2)
    Wend
    
    If InStr(u, ".") > 0 Then
        Y = Left$(u, InStr(u, ".") - 1)
    Else
        Y = u
    End If
    
    ' Added 2/06 to fix an exploit brought by Fiend(KIP)
    If (InStr(u, "..") > 0) Then
        u = Replace(u, "..", "")
    End If
    
    'Debug.Print Y
    
    Select Case UCase$(Y)
        Case "CON", "PRN", "AUX", "CLOCK$", "NUL", "COM1", "COM2", "COM3", _
            "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", _
            "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            
            tmpBuf = "You cannot read that file."
    End Select
    
    If (InStr(1, u, ".ini", vbTextCompare) > 0) Then
        tmpBuf = ".INI files cannot be read."
    End If
    
    u = App.Path & "\" & u
    
    On Error GoTo readError
    
    If Dir$(u) = vbNullString Then
        tmpBuf = "File does not exist."
    End If
    
    i = FreeFile
    
    Open u For Input As #i
    
        If (LOF(i) < 2) Then
            tmpBuf = "File is empty."
            Close #i
        End If
        
        Do While Not EOF(i)
            Line Input #i, Y
            
            If (Len(Y) > 200) Then
                Y = Left$(Y, 200)
            End If
            
            If Len(Y) > 2 Then
                'FilteredSend Username, Y, WhisperCmds, InBot, PublicOutput
            End If
        Loop
    
    Close #i
    
readError:
    tmpBuf = "Error reading file."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnReadFile

' handle greet command
Private Function OnGreet(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String
    Dim tmpBuf     As String ' temporary output buffer
    
    strArray() = Split(msgData, " ", 3)
        
    If (UBound(strArray) > 0) Then
        Select Case LCase$(strArray(1))
            Case "on"
                BotVars.UseGreet = True
                tmpBuf = "Greet messages enabled."
                WriteINI "Other", "UseGreets", "Y"
            
            Case "off"
                BotVars.UseGreet = False
                tmpBuf = "Greet messages disabled."
                WriteINI "Other", "UseGreets", "N"
                
            Case "whisper"
                If UBound(strArray) > 1 Then
                    Select Case LCase$(strArray(2))
                        Case "on"
                            BotVars.WhisperGreet = True
                            tmpBuf = "Greet messages will now be whispered."
                            WriteINI "Other", "WhisperGreet", "Y"
                            
                        Case "off"
                            BotVars.WhisperGreet = False
                            tmpBuf = "Greet messages will no longer be whispered."
                            WriteINI "Other", "WhisperGreet", "N"
                            
                    End Select
                End If

            Case Else
                If InStr(1, msgData, "/squelch", vbTextCompare) > 0 Or _
                    InStr(1, msgData, "/ban ", vbTextCompare) > 0 Or _
                        InStr(1, msgData, "/ignore", vbTextCompare) > 0 Or _
                            InStr(1, msgData, "/des", vbTextCompare) > 0 Or _
                                InStr(1, msgData, "/re", vbTextCompare) > 0 Then
                                
                    tmpBuf = "One or more invalid terms are present. Greet message not set."
                Else
                    tmpBuf = "Greet message set."
                    BotVars.GreetMsg = Right(msgData, Len(msgData) - 7)
                    WriteINI "Other", "GreetMsg", Right(msgData, Len(msgData) - 7)
                End If
        
        End Select
    End If

    'If LenB(tmpBuf) > 0 Then
    'Else
    'End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGreet

' handle allseen command
Private Function OnAllSeen(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    Dim i      As Integer

    tmpBuf = "Last 15 users seen: "
        
    If (colLastSeen.Count = 0) Then
        tmpBuf = tmpBuf & "(list is empty)"
    Else
        For i = 1 To colLastSeen.Count
            tmpBuf = tmpBuf & colLastSeen.Item(i)
            
            If (Len(tmpBuf) > 90) Then
                tmpBuf = tmpBuf & " [more]"
                'FilteredSend Username, tmpBuf, WhisperCmds, InBot, PublicOutput
                tmpBuf = ""
            ElseIf (i < colLastSeen.Count) Then
                tmpBuf = tmpBuf & ", "
            End If
        Next i
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAllSeen

' handle ban command
Private Function OnBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim u        As String
    Dim tmpBuf   As String ' temporary output buffer
    Dim banMsg   As String
    Dim Y        As String
    Dim i        As Integer

    dbAccess = GetAccess(Username)

    If ((MyFlags <> 2) And (MyFlags <> 18)) Then
        'If (InBot) Then
        '    tmpBuf = "You are not a channel operator."
        'End If
    End If
        
    On Error GoTo sendit6:
    
    u = Right(msgData, Len(msgData) - 5)
    i = InStr(1, u, " ")
    
    If (i > 0) Then
        banMsg = Mid$(u, i + 1)
        u = Left$(u, i - 1)
    End If
    
    If (InStr(1, u, "*", vbTextCompare) > 0) Then
        WildCardBan u, banMsg, 1
    Else
        If (banMsg <> vbNullString) Then
            Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), dbAccess.Access)
        Else
            Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), dbAccess.Access)
        End If
    End If
    
    If (Len(Y) > 2) Then
        tmpBuf = Y
    Else
        
    End If
    
sendit6:
    
    tmpBuf = "Who do you want to ban?"
    
    ' return message
    cmdRet(0) = tmpBuf
        
End Function ' end function OnBan

' handle unban command
Private Function OnUnBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sende
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, (Len(msgData) - 7))
    
    If bFlood Then
        If floodCap < 45 Then
            floodCap = floodCap + 15
            bnetSend "/unban " & u
        End If
    End If
    
    If InStr(1, msgData, "*", vbTextCompare) <> 0 Then
        WildCardBan u, vbNullString, 2
    End If
    
    If Dii = True Then
        If Not (Mid$(u, 1, 1) = "*") Then u = "*" & u
    End If
    
    tmpBuf = "/unban " & u
    
    'If InBot Then
    '    AddQ tmpBuf, 1
    '    GoTo theEnd
    'End If
    
sende:
    tmpBuf = "Unban who?"
                
    ' return message
    cmdRet(0) = tmpBuf
                
End Function ' end function OnUnBan

' handle kick command
Private Function OnKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendit
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim u      As String
    Dim i      As Integer
    Dim banMsg As String
    Dim tmpBuf As String ' temporary output buffer
    Dim Y      As String
    
    dbAccess = GetAccess(Username)
    
    'If MyFlags <> 2 And MyFlags <> 18 Then
    '   If InBot Then
    '       tmpBuf = "You are not a channel operator."
    '   End If
    'End If
    
    u = Right(msgData, Len(msgData) - 6)
    i = InStr(1, u, " ", vbTextCompare)
    
    If (i > 0) Then
        banMsg = Mid$(u, i + 1)
        
        u = Left$(u, i - 1)
    End If
    
    If (InStr(1, u, "*", vbTextCompare) > 0) Then
        If (dbAccess.Access > 99) Then
            Call WildCardBan(u, banMsg, 0)
        Else
            Call WildCardBan(u, banMsg, 0)
        End If
    End If
    
    Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), _
        dbAccess.Access, 1)
    
    If (Len(Y) > 1) Then
        tmpBuf = Y
    Else
        GoTo sendit
    End If
    
    Exit Function
    
sendit:
    tmpBuf = "Kick what user?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnKick

' handle lastwhisper command
Private Function OnLastWhisper(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (LastWhisper <> vbNullString) Then
        tmpBuf = "The last whisper to this bot was from: " & LastWhisper
    Else
        tmpBuf = "The bot has not been whispered since it logged on."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnLastWhisper

' handle say command
Private Function OnSay(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer
    Dim tmpSend  As String ' ...
        
    dbAccess = GetAccess(Username)
    
    If (Len(msgData) > 0) Then
        If (dbAccess.Access >= GetAccessINIValue("say70", 70)) Then
            If (dbAccess.Access >= GetAccessINIValue("say90", 90)) Then
                tmpSend = msgData
            Else
                tmpSend = Replace(msgData, "/", "")
            End If
        Else
            tmpSend = Username & " says: " & msgData
        End If
    Else
        tmpBuf = "Say what?"
    End If
    
    If (Len(tmpSend) > 0) Then
        Call AddQ(tmpSend)
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnSay

' handle expand command
Private Function OnExpand(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = Right(msgData, Len(msgData) - 8)
    tmpBuf = Expand(u)
    
    If (Len(tmpBuf) > 220) Then
        tmpBuf = Mid$(tmpBuf, 1, 220)
    End If
    
    'If InBot And Not WhisperedIn Then
    '    AddQ tmpBuf
    '    GoTo theEnd
    'ElseIf WhisperedIn Then
    '    AddQ tmpBuf
    '    GoTo theEnd
    'Else
    'End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnExpand

' handle detail command
Private Function OnDetail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = GetDBDetail(Split(msgData, " ")(1))
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDetail

' handle info command
Private Function OnInfo(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u      As String
    Dim i      As Integer
    Dim tmpBuf As String ' temporary output buffer

    u = Right(msgData, Len(msgData) - 6)
    i = UsernameToIndex(u)
    
    If (i > 0) Then
        With colUsersInChannel.Item(i)
            tmpBuf = "User " & .Username & " is logged on using " & ProductCodeToFullName(.Product)
            
            If ((.Flags) And (&H2) = &H2) Then
                tmpBuf = tmpBuf & " with ops, and a ping time of " & .Ping & "ms."
            Else
                tmpBuf = tmpBuf & " with a ping time of " & .Ping & "ms."
            End If
            
            'If WhisperCmds And Not InBot Then
            '    AddQ "/w " & IIf(Dii, "*", "") & Username & Space(1) & tmpBuf
            'ElseIf InBot And Not PublicOutput Then
            '    frmChat.AddChat RTBColors.ConsoleText, tmpBuf
            'Else
            '    AddQ tmpBuf
            'End If
            
            tmpBuf = "He/she has been present in the channel for " & ConvertTime(.TimeInChannel(), 1) & "."
        End With
    Else
        tmpBuf = "No such user is present."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnInfo

' handle shout command
Private Function OnShout(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendy
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim u         As String
    Dim tmpBuf    As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)
    
    u = UCase$(Right(msgData, (Len(msgData) - 7)))

    If (dbAccess.Access > 69) Then
        If (dbAccess.Access > 89) Then
            tmpBuf = u
        Else
            tmpBuf = Replace(u, "/", vbNullString)
        End If
    Else
        tmpBuf = Username & " shouts: " & u
    End If
    
    AddQ tmpBuf
sendf2:
    tmpBuf = "Shout what?"
    
sendy:
    tmpBuf = "Say what?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnShout

' handle voteban command
Private Function OnVoteBan(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (VoteDuration < 0) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, Right(msgData, Len(msgData) - 9))
        
        VoteDuration = 30
        
        tmpBuf = "30-second VoteBan vote started. Type YES to ban " & Right(msgData, Len(msgData) - 9) & ", NO to acquit him/her."
    Else
        tmpBuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVoteBan

' handle votekick command
Private Function OnVoteKick(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)
    
    If (VoteDuration < 0) Then
        Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, Right(msgData, Len(msgData) - 10))
        
        VoteDuration = 30
        
        VoteInitiator = dbAccess
        
        tmpBuf = "30-second VoteKick vote started. Type YES to kick " & Right(msgData, Len(msgData) - 9) & ", NO to acquit him/her."
    Else
        tmpBuf = "A vote is currently in progress."
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVoteKick

' handle vote command
Private Function OnVote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)
    
    If (VoteDuration < 0) Then
        If ((StrictIsNumeric(Right(msgData, Len(msgData) - 6))) And _
            (Val(Mid$(msgData, 7)) <= 32000)) Then
            
            VoteDuration = Right(msgData, Len(msgData) - 6)
            
            VoteInitiator = dbAccess
            
            Call Voting(BVT_VOTE_START, BVT_VOTE_STD)
            
            tmpBuf = "Vote initiated. Type YES or NO to vote; your vote will be counted only once."
        Else
            tmpBuf = "Please enter a number of seconds for your vote to last."
        End If
    Else
        tmpBuf = "A vote is currently in progress."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnVote

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
        AddQ "/away", 1
        
        'If (Not (InBot)) Then
        '    tmpBuf = "/me is back from " & AwayMsg & "."
        '    AwayMsg = vbNullString
        '
        '    GoTo theEnd
        'End If
    Else
        If (Not (BotVars.DisableMP3Commands)) Then
            If (iTunesReady) Then
                iTunesBack
                
                tmpBuf = "Skipped backwards."
            Else
                On Error GoTo sendf3
                
                hWndWA = GetWinamphWnd()
                
                If (hWndWA = 0) Then
                   tmpBuf = "Winamp is not loaded."
                End If
                
                Call SendMessage(hWndWA, WM_COMMAND, WA_PREVTRACK, 0)
                
                tmpBuf = "Skipped backwards."

sendf3:
                tmpBuf = "Error."
            End If
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnBack

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
    
    If (LenB(AwayMsg) > 0) Then
        AddQ "/away", 1
        'If Not InBot Then AddQ "/me is back from (" & AwayMsg & ")"
        AwayMsg = ""
    Else
        AddQ "/away", 1
        'If Not InBot Then AddQ "/me is away (" & AwayMsg & ")"
        AwayMsg = "-"
    End If
    
    On Error GoTo sendcc
    
    If (Len(msgData) <= 5) Then
        AddQ "/away", 1
        'If Not InBot Then AddQ "/me is away."
        AwayMsg = "-"
    Else
        AwayMsg = Right(msgData, (Len(msgData) - 6))
        AddQ "/away " & AwayMsg, 1
        'If Not InBot Then AddQ "/me is away (" & AwayMsg & ")"
    End If
    
sendcc:
    tmpBuf = "/away"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAway

' handle MP3 command
Private Function OnMP3(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim WindowTitle As String
    Dim tmpBuf      As String ' temporary output buffer
    
    WindowTitle = GetCurrentSongTitle(True)

    If (WindowTitle = vbNullString) Then
        tmpBuf = "Winamp is not loaded."
    Else
        tmpBuf = "Current MP3: " & WindowTitle
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMP3

' handle deldef command
Private Function OnDelDef(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo error_deldef
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Mid$(msgData, 9)
    
    If (Len(u) > 0) Then
        WriteINI "Def", u, "%deleted%", "definitions.ini"
        
        tmpBuf = "That definition has been erased."
    Else
    
error_deldef:
        tmpBuf = "There was an error removing that definition."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDelDef

' handle define command
Private Function OnDefine(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendyz
    
    Dim response As String
    Dim u        As String
    Dim tmpBuf   As String ' temporary output buffer
    Dim Track    As Long
        
    If (Dir$(GetFilePath("definitions.ini")) = vbNullString) Then
        tmpBuf = "No definition list found. Please use " & _
            "'.newdef term|definition' to make one."
    End If
    
    If Left$(msgData, 8) = "define" Then
        Track = 8
    Else
        Track = 5
    End If
    
    u = LCase$(Trim(Mid$(msgData, Track)))
    
    response = ReadINI("Def", u, "definitions.ini")
    
    If ((response = vbNullString) Or _
        (StrComp(response, "%deleted%"))) = 0 Then
        
        tmpBuf = "No definition on file for " & u & "."
    Else
        tmpBuf = "[" & u & "]: " & response
    End If

sendyz:
    tmpBuf = "Define what?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnDefine

' handle newdef command
Private Function OnNewDef(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendi2
        
    Dim Track  As Long
    Dim tmpBuf As String ' temporary output buffer
    Dim u      As String
    Dim z      As String
    
    u = Right(msgData, Len(msgData) - 8)
    
    Track = InStr(1, u, "|", vbTextCompare)
    
    z = Right(u, Len(u) - Track)
    u = Left$(u, Len(u) - Len(z) - 1)
    
    If (z = "") Then
        tmpBuf = "You need to specify a definition."
    End If
    
    WriteINI "Def", u, z, "definitions.ini"
    
    tmpBuf = "Added a definition for """ & u & """."
    
sendi2:
    tmpBuf = "Error: Please format your definitions correctly. (.newdef term|definition)"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnNewDef

' handle ping command
Private Function OnPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendc
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, (Len(msgData) - 6))
    
    Dim UserPing As Long
    
    UserPing = GetPing(u)
    
    If (UserPing < -1) Then
        tmpBuf = "I can't see " & u & " in the channel."
    Else
        tmpBuf = u & "'s ping at login was " & UserPing & "ms."
    End If
    
sendc:
    tmpBuf = "Ping who?"
    
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
    
    If (Len(u) > 0) Then
        Y = Dir$(GetFilePath("quotes.txt"))
        
        If (LenB(Y) = 0) Then
            Open (Y) For Output As #f
                Print #f, u
            Close #f
        Else
            Open (Y) For Append As #f
                Print #f, u
            Close #f
        End If
            
        tmpBuf = "Quote added!"
    Else
        tmpBuf = "I need a quote to add."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAddQuote

' handle owner command
Private Function OnOwner(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer
    
    If (LenB(BotVars.BotOwner) > 0) Then
        tmpBuf = "This bot's owner is " & BotVars.BotOwner & "."
    Else
        tmpBuf = "No owner is set."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnOwner

' handle ignore command
Private Function OnIgnore(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    On Error GoTo sendit9
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim u        As String
    Dim tmpBuf   As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)
        
    If (Mid$(msgData, 5, 1) = "o") Then
        u = Right(msgData, Len(msgData) - 8)
    Else
        u = Mid$(msgData, 6)
    End If
    
    If ((GetAccess(u).Access >= dbAccess.Access) Or _
        (InStr(GetAccess(u).Flags, "A"))) Then
        
        tmpBuf = "That user has equal or higher access."
    Else
        AddQ "/ignore " & IIf(Dii, "*", "") & u, 1
        
        tmpBuf = "Ignoring messages from " & Chr(34) & u & Chr(34) & "."
    End If

sendit9:
    tmpBuf = "Unignore who?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnIgnore

' handle quote command
Private Function OnQuote(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = GetRandomQuote
    
    If (Len(tmpBuf) = 0) Then
        tmpBuf = "Error reading quotes, or no quote file exists."
    End If
    
    If (Len(tmpBuf) > 220) Then
        ' try one more time
        tmpBuf = GetRandomQuote
        
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
    
    On Error GoTo sendit7
    
    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer
    
    u = Right(msgData, Len(msgData) - 10)
    
    If (Dii = True) Then
        If (Not (Mid$(u, 1, 1) = "*")) Then
            u = "*" & u
        End If
    End If
    
    AddQ "/unignore " & u, 1
    tmpBuf = "Receiving messages from """ & u & """."
    
    
sendit7:
    tmpBuf = "Unignore who?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnUnignore

' handle cq command
Private Function OnCQ(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    While (colQueue.Count > 0)
        Call colQueue.Remove(1)
    Wend
    
    If (InStr(1, msgData, "s") > 0) Then
        ' silently clear queue by exiting
        ' function after clearing
        Exit Function
    Else
        tmpBuf = "Queue cleared."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnCQ

' handle time command
Private Function OnTime(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf As String ' temporary output buffer

    tmpBuf = "The current time on this computer is " & Time & " on " & Format(Date, "MM-dd-yyyy") & "."
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnTime

' handle getping command
Private Function OnGetPing(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim tmpBuf  As String ' temporary output buffer
    Dim latency As Long

'   If (InBot) Then
'       If (g_Online) Then
'           tmpBuf = "Your ping at login was " & GetPing(CurrentUsername) & "ms."
'       Else
'           tmpBuf = "You are not connected."
'       End If
'   Else
'       latency = GetPing(Username)
'
'       If (latency > -2) Then
'           tmpBuf = "Your ping at login was " & hWndWA & "ms."
'       End If
'   End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGetPing

' handle checkmail command
Private Function OnCheckMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Track  As Long
    Dim tmpBuf As String ' temporary output buffer
    
    Track = GetMailCount(CurrentUsername)
            
     If (Track > 0) Then
         tmpBuf = "You have " & Track & " new messages."
         
         If (Username = "(console)") Then
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
    
    Dim Msg As udtMail
    
    Dim tmpBuf As String ' temporary output buffer
            
    If (Username = "(console)") Then
        Username = CurrentUsername
    End If
    
    If (GetMailCount(Username) > 0) Then
        Call GetMailMessage(Username, Msg)
        
        If (Len(RTrim(Msg.To)) > 0) Then
            tmpBuf = "msgData from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message)
        End If
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnGetMail

' handle whoami command
Private Function OnWhoAmI(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'Dim dbAccess As udtGetAccessResponse
    
    Dim tmpBuf   As String ' temporary output buffer
    
    dbAccess = GetAccess(Username)

    If (Username = "(console)") Then
        tmpBuf = "You are the bot console."
    
        If (g_Online) Then
            AddQ "/whoami"
        End If
    ElseIf (dbAccess.Access = 1000) Then
        tmpBuf = "You are the bot owner, " & Username & "."
    Else
        tmpBuf = "You have "
    
        If (dbAccess.Access > 0) Then
            tmpBuf = tmpBuf & dbAccess.Access & " access"
    
            If (dbAccess.Flags <> vbNullString) Then
                tmpBuf = tmpBuf & " and "
            End If
       End If
    
        If (dbAccess.Flags <> vbNullString) Then
            tmpBuf = tmpBuf & "flags " & dbAccess.Flags
        End If
    
        If (StrComp(tmpBuf, "You have ") = 0) Then
            tmpBuf = "You have no access or flags, " & Username & "."
        Else
            tmpBuf = tmpBuf & ", " & Username & "."
        End If
    End If

    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnWhoAmI

' handle add command
Private Function OnAdd(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    'On Error GoTo AddError
    
    Dim gAcc       As udtGetAccessResponse
    'Dim dbAccess   As udtGetAccessResponse
    
    Dim Track      As Long
    Dim strArray() As String
    Dim strUser    As String
    Dim i          As Integer
    Dim thisChar   As String
    Dim c          As Integer
    Dim iWinamp    As Integer
    Dim tmpBuf     As String ' temporary output buffer
    Dim u          As String
    
    dbAccess = GetAccess(Username)
    
    Track = 0
    strArray() = Split(msgData, " ")
    
    If (UBound(strArray) > 1) Then
        While ((Left$(strArray(1), 1) = "/") And _
            (Len(strArray(1)) > 0))
            
            strArray(1) = Mid$(strArray(1), 2)
        Wend
        
        gAcc = GetAccess(strArray(1))
        
        If (InStr(1, gAcc.Flags, "L", vbTextCompare) > 0) Then
            If ((InStr(1, dbAccess.Flags, "A", vbBinaryCompare) = 0) And _
                (dbAccess.Access < 100)) Then
                
                tmpBuf = "You do not have permission to modify that user's access."
            End If
        End If
        
        '  0     1       2      3
        ' set username access flags
        
        If (IsW3) Then
            If (StrComp(Right$(strArray(1), Len(w3Realm) + 1), "@" & w3Realm, _
                vbBinaryCompare) = 0) Then
                
                'Debug.Print w3Realm
                'Debug.Print Right$(strArray(1), Len(w3Realm) + 1)
                'Debug.Print Left$(strArray(1), InStr(1, strArray(1), "@", vbBinaryCompare) - 1)
                
                strArray(1) = Left$(strArray(1), _
                    InStr(1, strArray(1), "@", vbBinaryCompare) - 1)
            End If
        End If
        
        tmpBuf = "Set " & strArray(1) & "'s"
        
        If (StrictIsNumeric(strArray(2))) Then
            If (dbAccess.Access <= gAcc.Access) Then
                c = 1
            End If
            
            If (Val(strArray(2)) >= dbAccess.Access) Then
                c = 1
            End If
            
            If (Val(strArray(2)) < 0) Then
                c = 1
            End If
        Else
            strUser = UCase$(strArray(2))
        End If
        
        If (UBound(strArray) > 2) Then
            If (StrictIsNumeric(strArray(3))) Then
                If (dbAccess.Access <= gAcc.Access) Then
                    c = 1
                End If
                
                If (Val(strArray(3)) >= dbAccess.Access) Then
                    c = 1
                End If
                
                If (Val(strArray(3)) < 0) Then
                    c = 1
                End If
            Else
                strUser = strUser & UCase$(strArray(3))
            End If
        End If
        
        If (c <> 1) Then
            If (InStr(1, strUser, "A", vbTextCompare) > 0) Then
                If (dbAccess.Access <= 100) Then
                    c = 1
                End If
            End If
            
            If (InStr(1, strUser, "B", vbTextCompare) > 0) Then
                If (dbAccess.Access < 70) Then
                    c = 1
                End If
                
                If (dbAccess.Access <= GetAccess(strArray(1)).Access) Then
                    c = 1
                End If
            End If
            
            If (InStr(1, strUser, "Z", vbTextCompare) > 0) Then
                If (dbAccess.Access < 70) Then
                    c = 1
                End If
            End If
            
            If (InStr(1, strUser, "D", vbTextCompare) > 0) Then
                If (dbAccess.Access < 100) Then
                    c = 1
                End If
            End If
            
            If (InStr(1, strUser, "L", vbTextCompare) > 0) Then
                If (dbAccess.Access < 100) Then
                    c = 1
                End If
            End If
            
            If (InStr(1, strUser, "S", vbTextCompare) > 0) Then
                If (dbAccess.Access < 70) Then
                    c = 1
                End If
            End If
        End If
        
        If (c = 1) Then
            tmpBuf = "You do not have enough access to do that."
        End If
        
        If ((InStr(1, msgData, "+", vbTextCompare) = 0) And _
            (InStr(1, msgData, "-", vbTextCompare) = 0)) Then
                gAcc.Flags = vbNullString
        End If
        
        If ((StrictIsNumeric(strArray(2))) And (Val(strArray(2)) < 1000)) Then
            gAcc.Access = strArray(2)
            
            tmpBuf = tmpBuf & " access to " & gAcc.Access
            
            c = 1
        Else
            strArray(2) = UCase$(strArray(2))
            
            If Left$(strArray(2), 1) = "+" Then
                For i = 2 To Len(strArray(2))
                    If (InStr(1, gAcc.Flags, Mid(strArray(2), i, 1), _
                        vbTextCompare) = 0) Then
                        
                        thisChar = Asc(Mid(strArray(2), i, 1))
                        
                        If ((thisChar >= 65) And (thisChar <= 90)) Then
                            gAcc.Flags = gAcc.Flags & Chr(thisChar)
                            
                            If (thisChar = Asc("B")) Then
                                Call Ban(strArray(1) & " AutoBan", (AutoModSafelistValue - 1))
                            ElseIf (thisChar = Asc("Z")) Then
                                Call WildCardBan(strArray(1), "Tagbanned", 1)
                            ElseIf (thisChar = Asc("S")) Then
                                Call AddToSafelist(strArray(1), Username)
                                
                                Track = 1
                            End If
                            
                        End If
                    End If
                Next i
            ElseIf (Left$(strArray(2), 1) = "-") Then
                For i = 2 To Len(strArray(2))
                    gAcc.Flags = Replace(gAcc.Flags, Mid(strArray(2), i, 1), vbNullString)
                    
                    thisChar = Asc(Mid$(strArray(2), i, 1))
                    
                    If (thisChar = Asc("B")) Then
                        For Track = LBound(gBans) To UBound(gBans)
                            If (StrComp(gBans(Track).Username, strArray(1), _
                                vbTextCompare) = 0) Then
                                
                                AddQ "/unban " & gBans(Track).Username
                            End If
                        Next Track
                    ElseIf (thisChar = Asc("S")) Then
                        Call RemoveFromSafelist(strArray(1))
                        
                        Track = 2
                    End If
                Next i
            Else
                For i = 1 To Len(strArray(2))
                    If (InStr(1, gAcc.Flags, Mid(strArray(2), i, 1), vbTextCompare) = 0) Then
                        thisChar = Asc(Mid$(strArray(2), i, 1))
                        
                        If ((thisChar >= 65) And (thisChar <= 90)) Then
                            If (InStr(strArray(2), "S")) Then
                                Call AddToSafelist(strArray(1), Username)
                                
                                Track = 1
                            End If
                        
                            gAcc.Flags = gAcc.Flags & Mid$(strArray(2), i, 1)
                            
                            If (thisChar = Asc("B")) Then
                                Call Ban(strArray(1) & " AutoBan", (AutoModSafelistValue - 1))
                            ElseIf (thisChar = Asc("Z")) Then
                                Call WildCardBan(strArray(1), "Tagbanned", 1)
                            End If
                        End If
                    End If
                Next i
            End If
            
            gAcc.Flags = Replace(gAcc.Flags, "S", "")
            
            tmpBuf = tmpBuf & " flags to " & gAcc.Flags
            
            c = 2
        End If
        
        If (UBound(strArray) > 2) Then
            If ((StrictIsNumeric(strArray(3))) And _
                (Val(strArray(3)) < 1000)) Then
                
                If (c <> 1) Then
                    gAcc.Access = strArray(3)
                    
                    If (c > 0) Then
                        tmpBuf = tmpBuf & " and access to " & gAcc.Access
                    Else
                        tmpBuf = tmpBuf & " access to " & gAcc.Access
                    End If
                End If
            Else
                strArray(3) = UCase$(strArray(3))
                
                If (c <> 2) Then
                    If (Left$(strArray(3), 1) = "+") Then
                        For i = 2 To Len(strArray(3))
                            If (InStr(1, gAcc.Flags, Mid(strArray(3), i, 1), _
                                vbTextCompare) = 0) Then
                                
                                thisChar = Asc(Mid$(strArray(3), i, 1))
                                
                                If ((thisChar >= 65) And (thisChar <= 90)) Then
                                    gAcc.Flags = gAcc.Flags & Chr(thisChar)
                                    
                                    If (thisChar = Asc("B")) Then
                                        Call Ban(strArray(1) & " AutoBan", (AutoModSafelistValue - 1))
                                    ElseIf (thisChar = Asc("Z")) Then
                                        Call WildCardBan(strArray(1), "Tagbanned", 1)
                                    ElseIf (thisChar = Asc("S")) Then
                                        Call AddToSafelist(strArray(1), Username)
                                        
                                        Track = 1
                                    End If
                                End If
                            End If
                        Next i
                    ElseIf (Left$(strArray(3), 1) = "-") Then
                        For i = 2 To Len(strArray(3))
                            gAcc.Flags = Replace(gAcc.Flags, Mid(strArray(3), i, 1), vbNullString)
                            
                            thisChar = Asc(Mid(strArray(3), i, 1))
                            
                            If (thisChar = Asc("B")) Then
                                For iWinamp = LBound(gBans) To UBound(gBans)
                                    If (StrComp(LCase$(gBans(iWinamp).Username), strArray(1), _
                                        vbTextCompare) = 0) Then
                                        
                                        AddQ "/unban " & gBans(iWinamp).Username
                                    End If
                                Next iWinamp
                            ElseIf (thisChar = Asc("S")) Then
                                Call RemoveFromSafelist(strArray(1))
                                
                                Track = 2
                            End If
                        Next i
                    Else
                        For i = 1 To Len(strArray(3))
                            If (InStr(1, gAcc.Flags, Mid(strArray(3), i, 1), vbTextCompare) = 0) Then
                                thisChar = Asc(Mid(strArray(3), i, 1))
                                
                                If ((thisChar >= 65) And (thisChar <= 90)) Then
                                    If InStr(strArray(3), "S") Then
                                        Call AddToSafelist(strArray(1), Username)
                                        
                                        Track = 1
                                    End If
                                    
                                    gAcc.Flags = gAcc.Flags & Mid(strArray(3), i, 1)
                                    
                                    If (thisChar = Asc("B")) Then
                                        Call Ban(strArray(1) & " AutoBan", (AutoModSafelistValue - 1))
                                    ElseIf (thisChar = Asc("Z")) Then
                                        Call WildCardBan(strArray(1), "Tagbanned", 1)
                                    End If
                                End If
                            End If
                        Next i
                    End If
                    
                    gAcc.Flags = Replace(gAcc.Flags, "S", "")
                    
                    If (c > 0) Then
                        If (Len(gAcc.Flags) = 0) Then
                            tmpBuf = tmpBuf & " and erased their flags"
                        Else
                            tmpBuf = tmpBuf & " and flags to " & gAcc.Flags
                        End If
                    Else
                        tmpBuf = tmpBuf & " flags to " & gAcc.Flags
                    End If
                End If
            End If
        End If
        
        u = GetFilePath("users.txt")
        
        For i = LBound(DB) To UBound(DB)
            With DB(i)
                If (StrComp(.Username, LCase$(strArray(1)), vbTextCompare) = 0) Then
                    .Access = gAcc.Access
                    .Flags = gAcc.Flags
                    .ModifiedBy = Username
                    .ModifiedOn = Now
                    
                    tmpBuf = tmpBuf & "."
                    
                    If ((InStr(1, tmpBuf, "to .", vbTextCompare) > 0) And _
                        (c = 0)) Then
                        
                        tmpBuf = "User " & strArray(1) & "'s flags erased."
                    End If
                    
                    If (BotVars.LogDBActions) Then
                        Call LogDBAction(ModEntry, Username, strArray(1), msgData)
                    End If
                    
                    Select Case (Track)
                        Case 1: tmpBuf = tmpBuf & " He/she is safelisted."
                        Case 2: tmpBuf = tmpBuf & " He/she is no longer safelisted."
                    End Select
        
                    Call WriteDatabase(u)
                End If
            End With
        Next i
        
        ReDim Preserve DB(UBound(DB) + 1)
        
        With DB(UBound(DB))
            .Username = LCase$(strArray(1))
            .Access = gAcc.Access
            .Flags = gAcc.Flags
            .ModifiedBy = Username
            .ModifiedOn = Now
            .AddedBy = Username
            .AddedOn = Now
        End With
        
        Call WriteDatabase(u)
                                
        tmpBuf = tmpBuf & "."
        
        '// c will control at this point whether or not flags have been erased
        If (InStr(1, tmpBuf, "to .", vbTextCompare) > 0) Then
            tmpBuf = "User " & strArray(1) & "'s flags erased."
        End If
        
        If (BotVars.LogDBActions) Then
            LogDBAction AddEntry, Username, strArray(1), msgData
        End If
        
        Select Case (Track)
            Case 1: tmpBuf = tmpBuf & " He/she is safelisted."
            Case 2: tmpBuf = tmpBuf & " He/she is no longer safelisted."
        End Select
        
        Call LoadDatabase
    Else
        tmpBuf = "Please specify access or flags."
    End If
    
AddError:
    tmpBuf = "Add error - Make sure you specified a username and access amount."
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnAdd

' handle mmail command
Private Function OnMMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Temp As udtMail
    
    Dim strArray() As String
    Dim tmpBuf As String ' temporary output buffer
    Dim c As Integer
    Dim f As Integer
    Dim Track As Long
    
    strArray = Split(msgData, " ", 2)
            
    If (UBound(strArray) > 0) Then
        tmpBuf = "Mass mailing "

        With Temp
            .From = Username
            .Message = strArray(1)
            
            If (StrictIsNumeric(strArray(0))) Then
                'number games
                Track = Val(strArray(0))
                
                For c = 0 To UBound(DB)
                    If (DB(c).Access = Track) Then
                        .To = DB(c).Username
                        
                        Call AddMail(Temp)
                    End If
                Next c
                
                tmpBuf = tmpBuf & "to users with access " & Track
            Else
                'word games
                strArray(0) = UCase$(strArray(0))
                
                For c = 0 To UBound(DB)
                    For f = 1 To Len(strArray(1))
                        If (InStr(DB(c).Flags, Mid$(strArray(0), f, 1)) > 0) Then
                            .To = DB(c).Username
                            
                            Call AddMail(Temp)
                            
                            Exit For
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

' handle mail command
Private Function OnMail(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim Temp       As udtMail

    Dim strArray() As String
    Dim tmpBuf     As String ' temporary output buffer
    
    strArray = Split(msgData, " ", 2)
    
    ' debug
    'For iWinamp = 0 To UBound(strArray)
    '    Debug.Print iWinamp & ": " & strArray(iWinamp)
    'Next
    
    If (UBound(strArray) > 0) Then
        Temp.From = Username
        Temp.To = strArray(0)
        Temp.Message = strArray(1)
        
        Call AddMail(Temp)
        
        tmpBuf = "Added mail for " & strArray(0) & "."
    Else
        tmpBuf = "Error processing mail."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnMail

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
    
    Dim i As Integer
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
    
    Dim tmpBuf As String ' temporary output buffer
    
    tmpBuf = "I am currently connected to " & BotVars.Server & "."
            
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnServer

' handle findr command
Private Function OnFindR(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim strArray() As String
    Dim tmpBuf     As String ' temporary output buffer
    Dim i          As Integer
    Dim c          As Integer
    Dim n          As Integer
    Dim Y          As String
    Dim b          As Boolean

    'Find in a Range added 4/12/06 thanks to a suggestion by rush4hire
    ' find 20 30
    ' c: upper bound
    ' n: lower bound
    ' y: previous message
    
    strArray() = Split(msgData, " ")
    
    tmpBuf = "User(s) found: "
    
    If (UBound(strArray) = 2) Then
        If (StrictIsNumeric(strArray(1)) And _
            (StrictIsNumeric(strArray(2)))) Then
            
            If ((Val(strArray(1)) < 1001) And _
                (Val(strArray(2)) < 1001)) Then
                
                c = Val(strArray(2))
                n = Val(strArray(1))
            Else
                tmpBuf = "You specified an invalid range for that command."
            End If
            
            For i = LBound(DB) To UBound(DB)
                If ((DB(i).Access >= n) And (DB(i).Access <= c)) Then
                    tmpBuf = tmpBuf & ", " & DB(i).Username & IIf(DB(i).Access > 0, "\" & _
                        DB(i).Access, vbNullString) & IIf(DB(i).Flags <> _
                        vbNullString, "\" & DB(i).Flags, vbNullString)
                    
                    If ((Len(tmpBuf) > 80) And (i <> UBound(DB))) Then
                        If (LenB(Y) > 0) Then
                            'FilteredSend Username, Y & " [more]", WhisperCmds, InBot, PublicOutput
                        End If
                        
                        Y = Replace(tmpBuf, ", ", " ")
                        Y = Replace(Y, ": , ", ": ")
                            
                        tmpBuf = "User(s) found: "
                    End If
                End If
            Next i
            
            If (LenB(Y) > 0) Then
                If (b) Then
                    'FilteredSend Username, Y, WhisperCmds, InBot, PublicOutput
                Else
                    'FilteredSend Username, Left$(Y, Len(Y) - 2), WhisperCmds, InBot, PublicOutput
                End If
            ElseIf (StrComp(tmpBuf, "User(s) found: ") = 0) Then
                tmpBuf = "No users were found in that range."
            Else
                tmpBuf = Replace(tmpBuf, ", ", " ")
                tmpBuf = Replace(tmpBuf, ": , ", ": ")
            End If
        Else
            tmpBuf = "You specified an invalid range for that command."
        End If
    Else
        tmpBuf = "You specified an invalid range for that command."
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFindR

' handle find command
Private Function OnFind(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim gAcc   As udtGetAccessResponse

    Dim u      As String
    Dim tmpBuf As String ' temporary output buffer

    u = GetFilePath("users.txt")
            
    If (Dir$(u) = vbNullString) Then
        tmpBuf = "No userlist available. Place a users.txt file" & _
            "in the bot's root directory."
    End If
    
    u = Right(msgData, (Len(msgData) - 6))
    
    If (StrictIsNumeric(u)) Then
        If (Len(u) < 4) Then
            'Call WildCardFind(u, 1, Username, InBot, Val(u), WhisperCmds, PublicOutput)
        End If
    End If
    
    If ((InStr(1, u, "*", vbTextCompare) <> 0) Or _
        (InStr(1, u, "?", vbTextCompare) <> 0)) Then
        
        'Call WildCardFind(u, 0, Username, InBot, , WhisperCmds, PublicOutput)
    End If
    
    gAcc = GetAccess(u)
    
    If (gAcc.Access > 0) Then
        If (gAcc.Flags <> vbNullString) Then
            tmpBuf = "Found user " & u & ", with access " & gAcc.Access & " and flags " & gAcc.Flags & "."
        Else
            tmpBuf = "Found user " & u & ", with access " & gAcc.Access & "."
        End If
    Else
        If (gAcc.Flags <> vbNullString) Then
            tmpBuf = "Found user " & u & ", with flags " & gAcc.Flags & "."
        Else
            tmpBuf = "User not found."
        End If
    End If
    
sendit2:
    tmpBuf = "Who do you want me to find?"
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFind

' handle whois command
Private Function OnWhoIs(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u As String

    u = Right(msgData, Len(msgData) - 7)
            
    'If InBot And Not PublicOutput Then
    '    AddQ "/whois " & u, 1
    'End If
    
    'Call Commands(dbAccess, Username, "find " & u, InBot, CC, WhisperedIn, PublicOutput)
End Function ' end function OnWhoIs

' handle findattr command
Private Function OnFindAttr(ByVal Username As String, ByRef dbAccess As udtGetAccessResponse, _
    ByVal msgData As String, ByVal InBot As Boolean, ByRef cmdRet() As String) As Boolean
    
    Dim u                As String
    Dim Track            As Long
    Dim response         As String
    Dim PreviousResponse As String
    Dim tmpBuf           As String ' temporary output buffer
    Dim i                As Integer

    u = UCase$(Mid(msgData, 11, 1))
            
    response = "Users found: "
    Track = 0
    Track = -1
    
    For i = LBound(DB) To UBound(DB)
        If (InStr(1, DB(i).Flags, u, vbTextCompare) > 0) Then
            Track = 1
            
            response = response & DB(i).Username & ", "
            
            If (Len(response) > 80) Then
                If (LenB(PreviousResponse) > 0) Then
                    'FilteredSend Username, PreviousResponse & " [more]", WhisperCmds, InBot, PublicOutput
                End If
                
                PreviousResponse = Left$(response, Len(response) - 2)
                
                response = "Users found: "
                
                Track = 0
            End If
        End If
    Next i
    
    If LenB(PreviousResponse) Then
        'FilteredSend Username, PreviousResponse & IIf(Track > 0, " [more]", ""), WhisperCmds, InBot, PublicOutput
        
        PreviousResponse = ""
    End If
                
    If (Track < 0) Then
        tmpBuf = "No user(s) match that flag."
    ElseIf (Track = 1) Then
        'FilteredSend Username, Left$(response, Len(response) - 2), WhisperCmds, InBot, PublicOutput
    End If
    
    ' return message
    cmdRet(0) = tmpBuf
End Function ' end function OnFindAttr

Public Function Cache(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
    Static s() As String
    Static sTyp As String
    Dim i As Integer
    
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

Public Function Expand(ByVal s As String) As String
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

Private Sub AddQ(ByVal s As String, Optional DND As Byte)
    Call frmChat.AddQ(s, DND)
End Sub

Public Sub WildCardBan(ByVal sMatch As String, ByVal smsgData As String, ByVal Banning As Byte) ', Optional ExtraMode As Byte)
    'Values for Banning byte:
    '0 = Kick
    '1 = Ban
    '2 = Unban
    Dim i As Integer, Typ As String, z As String
    Dim iSafe As Integer
    
    If smsgData = vbNullString Then smsgData = sMatch
    sMatch = PrepareCheck(sMatch)
    'frmchat.addchat rtbcolors.ConsoleText, "Fired."
    'frmchat.addchat rtbcolors.ConsoleText, "Initial smsgData: " & smsgData
    'frmchat.addchat rtbcolors.ConsoleText, "Initial sMatch: " & sMatch
    
    Select Case Banning
        Case 1: Typ = "ban "
        Case 2: Typ = "unban "
        Case Else: Typ = "kick "
    End Select
    
    If Dii Then Typ = Typ & "*"
    
    If colUsersInChannel.Count < 1 Then Exit Sub
    
    If Banning <> 2 Then
        ' Kicking or Banning
    
        For i = 1 To colUsersInChannel.Count
            
            With colUsersInChannel.Item(i)
                If Not .IsSelf() Then
                    z = PrepareCheck(.Username)
                    
                    If z Like sMatch Then
                        If GetAccess(.Username).Access <= 20 Then
    '                        If ExtraMode = 0 Then
                                If Not .Safelisted Then
                                    If LenB(.Username) > 0 And (.Flags <> 2 And .Flags <> 18) Then AddQ "/" & Typ & .Username & Space(1) & smsgData, 1
                                Else
                                    iSafe = iSafe + 1
                                End If
    '                        Else
    '                            If Not .Safelisted Then
    '                                If .Username <> vbNullString And (.Flags <> 2 Or .Flags <> 18) Then AddQ "/" & Typ & .Username & Space(1) & smsgData, 1
    '                            End If
    '                        End If
                        Else
                            iSafe = iSafe + 1
                        End If
                    End If
                End If
            End With
        Next i
        
        If iSafe > 0 Then
            If StrComp(smsgData, ProtectMsg, vbTextCompare) <> 0 Then
                AddQ "Encountered " & iSafe & " safelisted user(s)."
            End If
        End If
        
    Else '// unbanning
    
        For i = 0 To UBound(gBans)
            If sMatch = "*" Then
                AddQ "/" & Typ & gBans(i).UsernameActual, 1
            Else
                z = PrepareCheck(gBans(i).UsernameActual)
                
                If (z Like sMatch) Then
                    AddQ "/" & Typ & gBans(i).UsernameActual, 1
                End If
            End If
        Next i
        
    End If
End Sub

Public Sub WildCardFind(ByVal bMatch As String, ByVal Mode As Byte, ByVal user As String, ByVal InBot As Boolean, Optional iAccess As Integer, Optional ByVal WhisperCmds As Boolean, Optional ByVal publicOutput As Boolean)
    'MODE 0 = standard find
    'MODE 1 = access level find
    'Dim s As String
    Dim i As Integer
    Dim ReturnMsg As String, PrevMsg As String
    Dim Found As Boolean
    
    ReturnMsg = "User(s) found: "
            
    Select Case Mode
        Case 0
            
            bMatch = PrepareCheck(bMatch)
            
            For i = LBound(DB) To UBound(DB)
                If PrepareCheck(DB(i).Username) Like bMatch Then
                    
                    Found = True
                    ReturnMsg = ReturnMsg & ", " & DB(i).Username & IIf(DB(i).Access > 0, "\" & DB(i).Access, vbNullString) & IIf(DB(i).Flags <> vbNullString, "\" & DB(i).Flags, vbNullString)
                    
                    If Len(ReturnMsg) > 80 And i <> UBound(DB) Then
                        If LenB(PrevMsg) > 0 Then
                            FilteredSend user, PrevMsg & " [more]", WhisperCmds, InBot, publicOutput
                        End If
                        
                        PrevMsg = Replace(ReturnMsg, " , ", " ")
                        PrevMsg = Replace(PrevMsg, ": , ", ": ")
                            
                        ReturnMsg = "User(s) found: "
                    End If
                    
                End If
            Next i
            
            If LenB(PrevMsg) > 0 Then
                If Found Then
                    FilteredSend user, PrevMsg, WhisperCmds, InBot, publicOutput
                Else
                    FilteredSend user, Left$(PrevMsg, Len(PrevMsg) - 2), WhisperCmds, InBot, publicOutput
                End If
            End If
            
            
        Case 1
        
            For i = LBound(DB) To UBound(DB)
                If DB(i).Access = iAccess Then
                
                    ReturnMsg = ReturnMsg & ", " & DB(i).Username
                    
                    If Len(ReturnMsg) > 90 And i <> UBound(DB) Then
                    
                        ReturnMsg = ReturnMsg & " [more]"
                        ReturnMsg = Replace(ReturnMsg, " , ", " ")
                        ReturnMsg = Replace(ReturnMsg, ": , ", ": ")
                        
                        If WhisperCmds And Not InBot Then
                            If Dii Then AddQ "/w *" & user & Space(1) & ReturnMsg Else AddQ "/w *" & user & Space(1) & ReturnMsg
                        ElseIf InBot And Not publicOutput Then
                            frmChat.AddChat RTBColors.ConsoleText, ReturnMsg
                        Else
                            AddQ ReturnMsg
                        End If
                        
                        ReturnMsg = "User(s) found: "
                    End If
                    
                End If
                
            Next i
    End Select
    
        
    If StrComp(ReturnMsg, "User(s) found: ", vbTextCompare) = 0 Then
        ReturnMsg = "No such user(s) found."
        
        If WhisperCmds And Not InBot Then
            If Dii Then AddQ "/w *" & user & Space(1) & ReturnMsg Else AddQ "/w " & user & Space(1) & ReturnMsg
        ElseIf InBot And Not publicOutput Then
            frmChat.AddChat COLOR_TEAL, "No such user(s) found."
        Else
            AddQ ReturnMsg
        End If
        
        Exit Sub
    End If
    
    ReturnMsg = Replace(ReturnMsg, " , ", " ") & "¦"
    ReturnMsg = Replace(ReturnMsg, " , ", " ")
    ReturnMsg = Replace(ReturnMsg, ": , ", ": ")
    ReturnMsg = Replace(ReturnMsg, ", ¦", vbNullString)
    
    
    If InStr(1, ReturnMsg, "¦", vbTextCompare) > 0 Then ReturnMsg = Left$(ReturnMsg, Len(ReturnMsg) - 1)
    
    If BotVars.WhisperCmds And Not InBot Then
        If Dii Then AddQ "/w *" & user & Space(1) & ReturnMsg Else AddQ "/w " & user & Space(1) & ReturnMsg
    ElseIf InBot And Not publicOutput Then
        frmChat.AddChat RTBColors.ConsoleText, ReturnMsg
    Else
        AddQ ReturnMsg
    End If
End Sub

Public Function RemoveItem(ByVal rItem As String, File As String) As String
    Dim s() As String, f As Integer
    Dim Counter As Integer, strCompare As String
    Dim strAdd As String
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

'Public Function GetAryPos(ByVal Username As String) As Integer
'    Dim i As Integer
'    With gChannel
'        For i = LBound(.Username) To UBound(.Username)
'            If StrComp(LCase$(Username), LCase$(.Username(i)), vbTextCompare) = 0 Then
'                GetAryPos = i
'                Exit Function
'            End If
'        Next i
'    End With
'    GetAryPos = -1
'End Function

Public Function GetTagbans(ByVal Username As String) As String
    Dim f As Integer, strCompare As String, banMsg As String
    
    If Dir$(GetFilePath("tagbans.txt")) <> vbNullString Then
        f = FreeFile
        Open (GetFilePath("tagbans.txt")) For Input As #f
        If LOF(f) > 1 Then
            Username = PrepareCheck(Username)
            
            Do While Not EOF(f)
                Line Input #f, strCompare
                If InStr(1, strCompare, " ", vbBinaryCompare) > 0 Then
                
                    'Debug.Print "strc: " & strCompare
                    banMsg = Mid$(strCompare, InStr(strCompare, " ") + 1)
                    'Debug.Print "banm: " & banMsg
                    strCompare = Split(strCompare, " ")(0)
                    'Debug.Print "strc: " & strCompare
                    
                End If
                
                If Username Like PrepareCheck(strCompare) Then
                    GetTagbans = strCompare & IIf(Len(banMsg) > 0, Space(1) & banMsg, vbNullString)
                    Close #f
                    Exit Function
                End If
            Loop
        End If
        Close #f
    End If
End Function

Public Function GetSafelistMatches(ByVal Username As String) As String
    Dim f As Integer, strCompare As String, ret As String
    
    If Dir$(GetFilePath("safelist.txt")) <> vbNullString Then
        f = FreeFile
        Open (GetFilePath("safelist.txt")) For Input As #f
        If LOF(f) > 1 Then
            Username = PrepareCheck(Username)
            
            Do While Not EOF(f)
                Line Input #f, strCompare
                If InStr(1, strCompare, " ", vbBinaryCompare) > 0 Then
                    'Debug.Print "strc: " & strCompare
                    'banMsg = Mid$(strCompare, InStr(strCompare, " ") + 1)
                    'Debug.Print "banm: " & banMsg
                    strCompare = Split(strCompare, " ")(0)
                    'Debug.Print "strc: " & strCompare
                End If
                
                If Username Like PrepareCheck(strCompare) Then
                    ret = ret & strCompare & " "
                End If
            Loop
        End If
        Close #f
    End If
    
    GetSafelistMatches = ReversePrepareCheck(Trim(ret))
End Function

Public Function GetSafelist(ByVal Username As String) As Boolean
    Dim i As Long
    
    On Error Resume Next
    Username = PrepareCheck(Username)
    GetSafelist = False
    
    If Not bFlood Then
        
        For i = 1 To colSafelist.Count
            If Username Like colSafelist.Item(i).Name Then
                GetSafelist = True
                Exit Function
            End If
        Next i
    
    Else
        
        For i = 0 To (UBound(gFloodSafelist) - 1)
            If Username Like gFloodSafelist(i) Then
                GetSafelist = True
                Exit Function
            End If
        Next i

    End If
    
End Function

Public Function GetShitlist(ByVal Username As String) As String
    Dim strCompare As String, toCheck As String
    Dim Temp As String
    Dim f As Integer
    f = FreeFile
    
    Username = LCase$(Username)
    
    On Error Resume Next
    Temp = GetFilePath("autobans.txt")
    
    If Dir$(Temp) <> vbNullString Then Open Temp For Input As #f Else GoTo theEnd
    
    If LOF(f) < 2 Then GoTo theEnd
    Do
        Line Input #f, strCompare
        toCheck = LCase$(strCompare)
        If InStr(1, toCheck, " ", vbTextCompare) <> 0 Then
            toCheck = Left$(strCompare, InStr(1, strCompare, " ", vbTextCompare) - 1)
        End If
        
        If StrComp(toCheck, Username, vbTextCompare) = 0 Then
            If InStr(1, strCompare, " ", vbTextCompare) = 0 Then
                GetShitlist = Username & " Shitlisted"
            Else
                GetShitlist = Username & Space(1) & Right(strCompare, Len(strCompare) - InStr(1, strCompare, " ", vbTextCompare))
            End If
            Close #f
            Exit Function
        End If
    Loop Until EOF(f)
theEnd:
    Close #f
End Function

Public Function GetPing(ByVal Username As String) As Long
    Dim i As Integer
    
    i = UsernameToIndex(Username)
    
    If i > 0 Then
        GetPing = colUsersInChannel.Item(i).Ping
    Else
        GetPing = -3
    End If
End Function

'Public Function ProcessCC(ByVal Speaker As String, ByRef SpeakerAccess As udtGetAccessResponse, RawmsgData As String, ByVal WhisperedIn As Boolean, ByVal PublicOutput As Boolean, Optional ByVal ExistenceCheckOnly As Boolean = False) As Boolean
'    Dim Args() As String, Send() As String, Actions() As String
'    Dim ccIn As udtCustomCommandData
'    Dim n As Integer, HighestArgUsed As Integer, c As Integer, i As Integer, f As Integer
'    Dim Found As Boolean, FirstTime As Boolean
'    Dim Temp As String, ReplaceString As String
'
'    On Error GoTo ProcessCC_Error
'
'    f = FreeFile
'
'    ReDim Send(0)
'
'    If LenB(Dir$(GetFilePath("commands.dat"))) > 0 Then
'
'        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccIn)
'
'        If LOF(f) > 1 Then
'
'            ' /****** Step 1: Get data ******/
'
'            HighestArgUsed = -1
'
'            i = LOF(f) \ LenB(ccIn)
'            If LOF(f) Mod LenB(ccIn) <> 0 Then i = i + 1
'
'            For i = 1 To i
'
'                Get #f, i, ccIn
'
'                If ccIn.reqAccess > 1000 Then GoTo NextItem
'                If LenB(RTrim(ccIn.Action)) < 1 Then GoTo NextItem
'                If LenB(RTrim(ccIn.Query)) < 1 Then GoTo NextItem
'                If ccIn.reqAccess = 0 Then ccIn.reqAccess = -1 'zero-access command
'
'                If SpeakerAccess.Access >= ccIn.reqAccess Then
'                    Args() = Split(RawmsgData, " ")
'
'                    '   0    1   2  3  4     5
'                    ' .myCC say %1 is fat! %rest
'                    If StrComp(RTrim(LCase$(ccIn.Query)), LCase$(Mid$(Args(0), 2))) = 0 Then
'                        Found = True
'                        ProcessCC = True
'
'                        If ExistenceCheckOnly Then GoTo theEnd
'
'                        Actions = Split(RTrim(Replace(ccIn.Action, vbCrLf, "")), "& ")
'
'                        If UBound(Actions) = 0 Then
'
'                            ReDim Send(0)
'                            Send(0) = Actions(0)
'
'                        Else
'                            ReDim Send(UBound(Actions))
'
'                            For c = 0 To UBound(Actions)
'                                Send(c) = Actions(c)
'                            Next c
'                        End If
'
'
'                        FirstTime = True
'
'                        For n = 0 To UBound(Send)
'                            Send(n) = Replace(Send(n), "%0", Speaker)
'                            Send(n) = Replace(Send(n), "%bc", BanCount)
'
'                            If UBound(Args) > 0 Then
'                                If InStr(Send(n), "%") > 0 Then
'                                    ' has probable arguments
'                                    ' /****** Step 2: Do basic replacements ******/
'                                    HighestArgUsed = 0
'
'                                    ' This was a normal FOR loop then it stopped working
'                                    '  altogether. So, it's a while loop now..
'                                    c = UBound(Args)
'
'                                    While c >= 0
'                                        If InStr(Send(n), "%" & c) Then
'                                            Send(n) = Replace(Send(n), "%" & (c), Args(c))
'                                            If HighestArgUsed = 0 Then HighestArgUsed = c
'                                        End If
'
'                                        c = c - 1
'                                    Wend
'
'                                    ' Assemble the %rest string
'                                   If FirstTime Then
'                                        'Debug.Print "firsttime, highest: " & c
'
'                                        If HighestArgUsed > -1 And UBound(Args) > HighestArgUsed Then
'                                            For c = HighestArgUsed + 1 To UBound(Args)
'                                                ReplaceString = ReplaceString & Args(c) & IIf(c < UBound(Args), " ", "")
'                                            Next c
'                                        End If
'
'                                        FirstTime = False
'                                    End If
'
'                                    ' /****** Step 3: Do advanced replacements ******/
'                                    If LenB(ReplaceString) > 0 Then
'                                        Send(n) = Replace(Send(n), "%rest", ReplaceString)
'                                    End If
'                                End If
'                            End If
'                        Next n
'
'                        ' /****** Step 4: Send and/or process normal command ******/
'                        For c = 0 To UBound(Send)
'                            'Debug.Print Send(c)
'
'                            If Left$(Send(c), 1) = "/" Then
'                                Call ExecuteCommand(SpeakerAccess, Speaker, Send(c), False, 2, WhisperedIn, PublicOutput)
'                            Else
'                                If Left$(Send(c), 6) = "_call " Then
'                                    Temp = Split(Mid$(Send(c), 7), " ")(0)
'
'                                    On Error Resume Next
'                                    frmChat.SControl.Run Temp
'                                Else
'                                    If Left$(RawmsgData, 1) = "/" And Not PublicOutput Then
'                                        frmChat.AddChat RTBColors.ConsoleText, Send(c)
'                                    Else
'                                        ' // including other term replacements
'                                        Send(c) = DoReplacements(Send(c), Speaker, GetPing(Speaker))
'                                        AddQ Send(c)
'                                    End If
'                                End If
'                            End If
'                        Next c
'
'                    End If 'strcomp
'                End If 'speakeraccess
'NextItem:
'            Next i
'
'        End If 'lof
'
'    End If
'
'theEnd:
'    If Not Found And Not ExistenceCheckOnly Then
'        If StrComp(Left$(RawmsgData, 1), "/", vbTextCompare) = 0 Then _
'            Call Commands(SpeakerAccess, Speaker, RawmsgData, False, 1, WhisperedIn, PublicOutput)
'    ElseIf ExistenceCheckOnly Then
'        ProcessCC = Found
'    End If
'
'    Close #f
'    Exit Function
'Error:
'    AddQ "Invalid argument(s) or incorrect number of arguments."
'    Close #f
'
'    On Error GoTo 0
'    Exit Function
'
'ProcessCC_Error:
'
'    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessCC of Module modCommandCode"
'End Function

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


Public Sub DBRemove(ByVal s As String)
    
    Dim i As Integer
    Dim c As Integer
    Dim n As Integer
    Dim t() As udtDatabase
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

Public Sub LoadDatabase()
    ReDim DB(0)
    
    Dim s As String, X() As String
    Dim Path As String
    Dim i As Integer, f As Integer
    Dim gA As udtDatabase, Found As Boolean
    
    Path = GetFilePath("users.txt")
    
    On Error Resume Next
    
    If Dir$(Path) <> vbNullString Then
        f = FreeFile
        Open Path For Input As #f
            
        If LOF(f) > 1 Then
            Do
                
                Line Input #f, s
                
                If InStr(1, s, " ", vbTextCompare) > 0 Then
                    X() = Split(s, " ")
                    
                    If UBound(X) > 0 Then
                        ReDim Preserve DB(i)
                        With DB(i)
                            .Username = LCase$(X(0))
                            
                            If StrictIsNumeric(X(1)) Then
                                .Access = Val(X(1))
                            Else
                                If X(1) <> "%" Then
                                    .Flags = X(1)
                                    
                                    If InStr(X(1), "S") > 0 Then
                                        AddToSafelist .Username
                                        .Flags = Replace(.Flags, "S", "")
                                    End If
                                End If
                            End If
                            
                            If UBound(X) > 1 Then
                                If StrictIsNumeric(X(2)) Then
                                    .Access = Int(X(2))
                                Else
                                    If X(2) <> "%" Then
                                        .Flags = X(2)
                                        
                                        If InStr(X(2), "S") > 0 Then
                                            AddToSafelist .Username
                                            .Flags = Replace(.Flags, "S", "")
                                        End If
                                    End If
                                End If
                                
                                '  0        1       2       3       4       5       6
                                ' username access flags addedby addedon modifiedby modifiedon
                                If UBound(X) > 2 Then
                                    .AddedBy = X(3)
                                    
                                    If UBound(X) > 3 Then
                                        .AddedOn = CDate(Replace(X(4), "_", " "))
                                        
                                        If UBound(X) > 4 Then
                                            .ModifiedBy = X(5)
                                            
                                            If UBound(X) > 5 Then
                                                .ModifiedOn = CDate(Replace(X(6), "_", " "))
                                            End If
                                        End If
                                    End If
                                End If
                                
                            End If
                            
                            If .Access > 999 Then .Access = 999
                        End With

                        i = i + 1
                    End If
                End If
                
            Loop While Not EOF(f)
        End If
        
        Close #f
    End If
    
    ' 9/13/06: Add the bot owner 1000
    If LenB(BotVars.BotOwner) > 0 Then
        For i = 0 To UBound(DB)
            If StrComp(DB(i).Username, BotVars.BotOwner, vbTextCompare) = 0 Then
                DB(i).Access = 1000
                Found = True
                Exit For
            End If
        Next i
        
        If Not Found Then
            ReDim Preserve DB(UBound(DB) + 1)
            DB(UBound(DB)).Username = BotVars.BotOwner
            DB(UBound(DB)).Access = 1000
        End If
    End If
End Sub

Public Function ValidateAccess(ByRef Acc As udtGetAccessResponse, ByVal CWord As String) As Boolean

    Dim Temp As String, i As Integer, n As Integer
    
    If LenB(CWord) > 0 Then
        
        'CWord = Mid$(CWord, 2)
        
        If LenB(ReadINI("DisabledCommands", CWord, "access.ini")) > 0 Or ReadINI("DisabledCommands", "universal", "access.ini") = "Y" Then
            ValidateAccess = False
            Exit Function
        End If
        
        Temp = UCase$(ReadINI("Flags", CWord, "access.ini"))
        
        If Len(Temp) > 0 Then
            For i = 1 To Len(Temp)
                If InStr(1, Acc.Flags, Mid(Temp, i, 1)) > 0 Then
                    ValidateAccess = True
                    Exit Function
                End If
            Next i
        End If
        
        n = AccessNecessary(CWord, i)
        
        'Debug.Print "N: " & n
        'Debug.Print "I: " & i
        'Debug.Print "A: " & Acc.Access
        
        If Acc.Access = -1 And i = 1 Then Acc.Access = 0
        '// needs to be set to 0 so that the below statements work when
        '// an entry is present in the access.ini file.
        
        If Acc.Access >= n Then
            'Debug.Print "Init"
            If Acc.Access > 0 Then
                ValidateAccess = True
                Exit Function
            Else
                If i = 1 Then
                    If Not (StrComp(CWord, "ver", vbBinaryCompare) = 0) Then '// hardcoded fix for .ver exploit :(
                        ValidateAccess = True
                        Exit Function
                    Else
                        If Acc.Access > 0 Then ValidateAccess = True
                    End If
                End If
            End If
        End If
        
        If InStr(1, Acc.Flags, "A", vbBinaryCompare) > 0 Then
            ValidateAccess = True
            Exit Function
        End If
        
    End If
    
End Function

Public Function AccessNecessary(ByVal CW As String, Optional ByRef i As Integer) As Integer
    
    'Debug.Print CW & vbTab & "[" & ReadINI("Numeric", CW, "access.ini") & "]"
    If Len(ReadINI("Numeric", CW, "access.ini")) > 0 Then
        
        AccessNecessary = Val(ReadINI("Numeric", CW, "access.ini"))
        i = 1
        
    Else
        Select Case CW
            Case "getmail"
                AccessNecessary = 0
            Case "find", "whois", "about", "server", "add", "set", "whoami", "cq", "scq", "designated", "about", "ver", "version", "mail", "findattr", "findflag", "flip", "bmail", "roll", "findr"
                AccessNecessary = 20
            Case "time", "trigger", "getping", "pingme", "checkmail"
                AccessNecessary = 40
            Case "say", "shout", "ignore", "unignore", "addquote", "quote", "away", "back", "ping", "uptime", "mp3", "ign", "owner"
                AccessNecessary = 50
            Case "vote", "voteban", "votekick", "tally", "info", "expand", "math", "eval", "where", "safecheck", "cancel"
                AccessNecessary = 50
            Case "kick", "ban", "unban", "lastwhisper", "define", "newdef", "def", "fadd", "frem", "bancount", "allseen", "levelbans", "lw", "deldef"
                AccessNecessary = 60
            Case "d2levelbans", "tagcheck", "detail", "dbd", "slcheck", "shitcheck"
                AccessNecessary = 60
            Case "shitlist", "shitdel", "safeadd", "safedel", "safelist", "tagbans", "tagadd", "tagdel", "protect", "pban", "shitadd", "sl"
                AccessNecessary = 70
            Case "mimic", "nomimic", "cmdadd", "addcmd", "cmddel", "delcmd", "cmdlist", "plist", "setpmsg", "mmail", "addtag", "cbans", "setcmdaccess"
                AccessNecessary = 70
            Case "padd", "addphrase", "phrases", "delphrase", "pdel", "pon", "poff", "phrasebans", "pstatus", "ipban", "banned", "notify", "denotify"
                AccessNecessary = 70
            Case "reconnect", "des", "designate", "rejoin", "settrigger", "igpriv", "unigpriv", "rem", "del", "sethome", "idle", "rj", "allowmp3"
                AccessNecessary = 80
            Case "next", "play", "stop", "setvol", "fos", "pause", "shuffle", "repeat"
                If Not BotVars.DisableMP3Commands Then
                    AccessNecessary = 80
                Else
                    AccessNecessary = 1000
                End If
                
            Case "idletime", "idletype", "block", "filter", "whispercmds", "profile", "greet", "levelban", "d2levelban", "clist", "clientbans", "cbans", "cadd", "cdel", "koy", "plugban", "useitunes", "setidle", "usewinamp"
                AccessNecessary = 80
            Case "join", "home", "resign", "setname", "setpass", "setserver", "quiettime", "giveup", "readfile", "chpw", "ib", "cb", "cs", "clan", "sweepban", "sweepignore", "op", "setkey", "setexpkey", "idlebans"
                AccessNecessary = 90
            Case "c", "exile", "unexile"
                AccessNecessary = 90
            Case "clearbanlist", "cbl"
                AccessNecessary = 90
            Case "quit", "locktext", "efp", "floodmode", "loadwinamp", "setmotd", "invite", "peonban"
                AccessNecessary = 100
            Case Else
                AccessNecessary = 1000
        End Select
    End If
    
End Function

Public Function GetRandomQuote() As String
    Dim Rand As Integer, f As Integer
    Dim s As String
    Dim colQuotes As Collection
    
    Set colQuotes = New Collection
    
   On Error GoTo GetRandomQuote_Error

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

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRandomQuote of Module modCommandCode"
    Resume GetRandomQuote_Exit
End Function

' Writes database to disk
' Updated 9/13/06 for new features
Public Sub WriteDatabase(ByVal u As String)
    Dim f As Integer, i As Integer
    
   On Error GoTo WriteDatabase_Exit

    f = FreeFile
    
    Open u For Output As #f
                
    For i = LBound(DB) To UBound(DB)
        If (DB(i).Access > 0 Or Len(DB(i).Flags) > 0) Then
            Print #f, DB(i).Username;
            Print #f, " " & DB(i).Access;
            Print #f, " " & IIf(Len(DB(i).Flags) > 0, DB(i).Flags, "%");
            Print #f, " " & IIf(Len(DB(i).AddedBy) > 0, DB(i).AddedBy, "%");
            Print #f, " " & IIf(DB(i).AddedOn > 0, DateCleanup(DB(i).AddedOn), "%");
            Print #f, " " & IIf(Len(DB(i).ModifiedBy) > 0, DB(i).ModifiedBy, "%");
            Print #f, " " & IIf(DB(i).ModifiedOn > 0, DateCleanup(DB(i).ModifiedOn), "%");
            Print #f, vbCr
        End If
    Next i

WriteDatabase_Exit:
    Close #f
    
   Exit Sub

WriteDatabase_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDatabase of Module modCommandCode"
    Resume WriteDatabase_Exit
End Sub


Public Sub FilteredSend(ByVal Username As String, ByVal ToSend As String, ByVal WhisperCmds As Boolean, ByVal InBot As Boolean, ByVal publicOutput As Boolean)
    If InBot And Not publicOutput Then
        frmChat.AddChat RTBColors.ConsoleText, ToSend
    ElseIf WhisperCmds And Not publicOutput Then
        If Not Dii Then
            AddQ "/w " & Username & Space(1) & ToSend
        Else
            AddQ "/w *" & Username & Space(1) & ToSend
        End If
    Else
        AddQ ToSend
    End If
End Sub


Public Function GetDBDetail(ByVal Username As String) As String
    Dim sRetAdd As String, sRetMod As String
    Dim i As Integer
    
    For i = 0 To UBound(DB)
        With DB(i)
            If StrComp(Username, .Username, vbTextCompare) = 0 Then
                If .AddedBy <> "%" And LenB(.AddedBy) > 0 Then
                    sRetAdd = " was added by " & .AddedBy & " on " & .AddedOn & "."
                End If
                
                If .ModifiedBy <> "%" And LenB(.ModifiedBy) > 0 Then
                    If (.AddedOn <> .ModifiedOn) Or (.AddedBy <> .ModifiedBy) Then
                        sRetMod = " was last modified by " & .ModifiedBy & " on " & .ModifiedOn & "."
                    Else
                        sRetMod = " have not been modified since they were added."
                    End If
                End If
                
                If LenB(sRetAdd) > 0 Or LenB(sRetMod) > 0 Then
                    If LenB(sRetAdd) > 0 Then
                        GetDBDetail = Username & sRetAdd & " They" & sRetMod
                    Else
                        'no add, but we could have a modify
                        GetDBDetail = Username & sRetMod
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

Public Function DateCleanup(ByVal tDate As Date) As String
    Dim t As String
    
    t = Format(tDate, "dd-MM-yyyy_HH:MM:SS")
    
    DateCleanup = Replace(t, " ", "_")
End Function

Function GetAccessINIValue(ByVal sKey As String, Optional ByVal Default As Long) As Long
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
