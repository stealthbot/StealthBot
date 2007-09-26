Attribute VB_Name = "modCommandCode"
Option Explicit

Public Flood As String
Public floodCap As Byte

Private Const WA_PREVTRACK As Long = 40044
Private Const WA_NEXTTRACK As Long = 40048
Private Const WA_PLAY As Long = 40045
Private Const WA_PAUSE As Long = 40046
Private Const WA_STOP As Long = 40047
Private Const WA_FADEOUTSTOP As Long = 40147

Public Sub ProcessCommand(ByVal Message As String, ByVal Username As String, ByVal Whispered As Boolean)
    Dim X() As String
    Dim i As Integer
    
'    Debug.Print "ProcessCommand() for " & Username & ": " & Message
'    Debug.Print "Access information: " & GetAccess(Username).Access & "\" & GetAccess(Username).Flags
    
    
    If Len(Message) > 1 Then
        If StrComp(LCase(Left$(Message, Len(CurrentUsername) + 2)), BotVars.Trigger & LCase(CurrentUsername) & Space(1), vbTextCompare) = 0 Then
            Message = Right(Message, Len(Message) - (Len(CurrentUsername) + 2))
            Message = BotVars.Trigger & Message
        End If
        
        If InStr(1, Message, "%me", vbTextCompare) > 0 Then
            Message = Replace(Message, "%me", Username)
        End If
        
        If Left$(Message, 1) = BotVars.Trigger Or StrComp(Message, "?trigger", vbTextCompare) = 0 Then
            If InStr(2, Message, "; ", vbTextCompare) > 0 Then
                X = Split(Right(Message, Len(Message) - 1), "; ")
                
                For i = LBound(X) To UBound(X)
                    If Left$(X(i), 1) = BotVars.Trigger Then
                        Call Commands(GetAccess(Username), Username, X(i), False, , Whispered)
                    Else
                        Call Commands(GetAccess(Username), Username, BotVars.Trigger & X(i), False, , Whispered)
                    End If
                Next i
            Else
                Call Commands(GetAccess(Username), Username, Message, False, , Whispered)
            End If
        End If
    End If
End Sub

Public Function Commands(ByRef dbAccess As udtGetAccessResponse, ByVal Username As String, Message As String, InBot As Boolean, Optional CC As Byte = 0, Optional WhisperedIn As Boolean, Optional PublicOutput As Boolean) As String

    Dim OldTrigger As String * 1
    Dim strArray() As String
    
    Dim thisChar As Byte
    
    Dim b As Boolean, PassedAccess As Boolean, RealCommand As Boolean
    Dim TriggerChange As Boolean, WhisperCmds As Boolean
    
    Dim z As String
    Dim u As String, strUser As String, Y As String, cMsg As String, banMsg As String, strSend As String
    Dim WindowTitle As String, Response As String, PreviousResponse As String
    
    Dim iWinamp As Integer, c As Integer, i As Integer, f As Integer, n As Integer
    Dim hWndWA As Long, Track As Long
    
    Dim gAcc As udtGetAccessResponse
    Dim ccOut As udtCustomCommandData
    
    Dim Temp As udtMail
    
    'Debug.Print "Command: " & Message
    
    If InBot Then Username = "(console)"
    
    If PublicOutput Then
        Message = Mid$(Message, 2)
        
        If Left$(Message, 1) <> "/" Then
            Message = "/" & Message
        End If
    End If
    
    f = FreeFile
    OldTrigger = BotVars.Trigger
    cMsg = LCase(Message)
    
    If WhisperedIn Or BotVars.WhisperCmds Then
        WhisperCmds = True
    End If
    
    If InBot = True Then
        dbAccess.Access = 1000
        dbAccess.Flags = "A"
        BotVars.Trigger = "/"
    End If
    
    If CC <> 1 Then
        'Debug.Print "Start access: " & dbAccess.Access
        'Debug.Print "Given values: " & Username & vbTab & dbAccess.Flags & vbTab & Split(cMsg, " ")(0)
        If StrComp(Username, BotVars.BotOwner, vbTextCompare) = 0 Then
            dbAccess.Access = 1000
        End If
        
        PassedAccess = ValidateAccess(dbAccess, Split(cMsg, " ")(0))
        
        'Debug.Print "Authorized with: " & PassedAccess
        
        If Left$(cMsg, 1) <> BotVars.Trigger Then
            If StrComp(cMsg, "?trigger") <> 0 And Left$(cMsg, 1) <> "/" Then
                GoTo theEnd
            End If
        End If
        
        If dbAccess.Access = -1 Or Not PassedAccess Then GoTo AccessZero
        
        
        '  BEGIN CRAZY IF STATEMENT
            If cMsg = (BotVars.Trigger & "quit") Then
                Call frmChat.Form_Unload(0)
                GoTo theEnd
                
            ElseIf cMsg = BotVars.Trigger & "locktext" Then
                Call frmChat.mnuLock_Click
                GoTo theEnd
                
            ElseIf cMsg = BotVars.Trigger & "allowmp3" Then
                If BotVars.DisableMP3Commands Then
                    strSend = "Allowing MP3 commands."
                    BotVars.DisableMP3Commands = False
                Else
                    strSend = "MP3 commands are now disabled."
                    BotVars.DisableMP3Commands = True
                End If
                b = True
                GoTo Display
            
            ElseIf cMsg = BotVars.Trigger & "loadwinamp" Then
                
                strSend = LoadWinamp(ReadCFG("Other", "WinampPath"))
                If Len(strSend) < 1 Then GoTo theEnd
                b = True
                GoTo Display
            
            ElseIf Left$(cMsg, 11) = BotVars.Trigger & "floodmode " Or Left$(cMsg, 5) = BotVars.Trigger & "efp " Then
            
                If Right(cMsg, 2) = "on" Then
                    
                    If InBot = False Then
                        If Not WhisperCmds Then
                            AddQ "Emergency floodbot protection enabled."
                        ElseIf WhisperCmds Then
                            AddQ "/w " & Username & " Emergency floodbot protection enabled."
                        End If
                    End If
                    
                    Call frmChat.SetFloodbotMode(1)
                    GoTo theEnd
                    
                ElseIf Right(cMsg, 6) = "status" Then
                    If bFlood Then
                        frmChat.AddChat RTBColors.TalkBotUsername, "Emergency floodbot protection is enabled. (No messages can be sent to battle.net.)"
                        GoTo theEnd
                    Else
                        strSend = "Emergency floodbot protection is disabled."
                        b = True
                        GoTo Display
                    End If
                    
                ElseIf Right(cMsg, 3) = "off" Then
                    Call frmChat.SetFloodbotMode(0)
                    strSend = "Emergency floodbot protection disabled."
                    b = True
                    GoTo Display
                    
                End If
                GoTo theEnd
        
            ElseIf cMsg = (BotVars.Trigger & "home") Then
                AddQ "/join " & BotVars.HomeChannel, 1
                GoTo theEnd
                
            ElseIf Mid$(cMsg, 2, 5) = "clan " Or Mid$(cMsg, 2, 2) = "c " Then
            
                If Mid$(cMsg, 3, 1) = "l" Then
                    u = Mid$(cMsg, 7)
                Else
                    u = Mid$(cMsg, 4)
                End If
                
                If MyFlags And 2 Then
                
                    Select Case u
                        Case "public", "pub"
                            strSend = "Clan channel is now public."
                            AddQ "/clan public", 1
                        Case "private", "priv"
                            strSend = "Clan channel is now private."
                            AddQ "/clan private", 1
                        Case Else
                            If InBot Then
                                strSend = "/clan " & Right$(Message, Len(u))
                                b = True
                                InBot = False
                                GoTo Display
                            End If
                    End Select
                
                Else
                    
                    strSend = "The bot must have ops to change clan privacy status."
                
                End If
                
                If Len(strSend) > 0 Then
                    b = True
                    GoTo Display
                Else
                    GoTo theEnd
                End If
                
                
            ElseIf Mid$(cMsg, 2, 8) = "peonban " Then
                
                u = Mid$(cMsg, 10)
                
                Select Case u
                    Case "on"
                        BotVars.BanPeons = 1
                        strSend = "Peon banning activated."
                        WriteINI "Other", "PeonBans", "1"
                    Case "off"
                        BotVars.BanPeons = 0
                        strSend = "Peon banning deactivated."
                        WriteINI "Other", "PeonBans", "0"
                    Case "status"
                        strSend = "The bot is currently "
                        If BotVars.BanPeons = 0 Then
                            strSend = strSend & "not banning peons."
                        Else
                            strSend = strSend & "banning peons."
                        End If
                End Select
                
                If Len(strSend) > 0 Then
                    b = True
                    GoTo Display
                Else
                    GoTo theEnd
                End If
                
            ElseIf Left$(cMsg, 8) = BotVars.Trigger & "invite " Then
                
                If IsW3 Then
                    If Clan.MyRank >= 3 Then
                        InviteToClan Right(cMsg, Len(cMsg) - 8)
                        strSend = Right(cMsg, Len(cMsg) - 8) & ": Clan invitation sent."
                    Else
                        strSend = "The bot must hold Shaman or Chieftain rank to invite users."
                    End If
                    b = True
                    GoTo Display
                End If
                
            ElseIf Left$(cMsg, 9) = BotVars.Trigger & "setmotd " Then
            
                If IsW3 Then
                    If Clan.MyRank >= 3 Then
                        SetClanMOTD Right(Message, Len(Message) - 9)
                        strSend = "Clan MOTD set."
                    Else
                        strSend = "Shaman or Chieftain rank is required to set the MOTD."
                    End If
                    b = True
                    GoTo Display
                    
                Else: GoTo theEnd
                End If
                
            ElseIf cMsg = BotVars.Trigger & "where" Then
                strSend = "I am currently in channel " & gChannel.Current & " (" & colUsersInChannel.Count & " users present)"
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 4) = BotVars.Trigger & "qt " Or Left$(cMsg, 11) = BotVars.Trigger & "quiettime " Then
                If InStr(1, cMsg, "qt ", vbTextCompare) <> 0 Then
                    c = 4
                Else
                    c = 11
                End If
                
                u = Right(cMsg, Len(cMsg) - c)

                If u = "on" Then
                    WriteINI "Main", "QuietTime", "Y"
                    BotVars.QuietTime = True
                    strSend = "Quiet-time enabled."
                ElseIf u = "off" Then
                    WriteINI "Main", "QuietTime", "N"
                    BotVars.QuietTime = False
                    strSend = "Quiet-time disabled."
                ElseIf u = "status" Then
                    If BotVars.QuietTime Then
                        strSend = "Quiet-time is currently enabled."
                    Else
                        strSend = "Quiet-time is currently disabled."
                    End If
                Else
                    strSend = "Invalid arguments."
                End If
                b = True
                GoTo Display
                
            ElseIf cMsg = BotVars.Trigger & "roll" Then
                Randomize
                iWinamp = CLng(Rnd * 100)
                strSend = "Random number (0-100): " & iWinamp
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 6) = BotVars.Trigger & "roll " Then
                Randomize
                
                If StrictIsNumeric(Mid$(Message, 7)) Then
                    If Val(Mid$(Message, 7)) < 100000000 Then
                        Track = CLng(Rnd * CLng(Mid$(Message, 7)))
                        
                        strSend = "Random number (0-" & Mid$(Message, 7) & "): " & Track
                        b = True
                        GoTo Display
                    End If
                Else
                    GoTo theEnd
                End If
                            
            ElseIf Left$(cMsg, 4) = BotVars.Trigger & "cb " Or Left$(cMsg, 4) = BotVars.Trigger & "cs " Then
                u = Right(Message, Len(Message) - 4)
                
                If Mid(cMsg, 3, 1) = "s" Then
                    Y = "squelch "
                Else
                    Y = "ban "
                End If
                
                Caching = True
                Cache vbNullString, 255, Y
                AddQ "/who " & u, 1
                'frmChat.quLower.Interval = 1400
                
                GoTo theEnd
                
            ElseIf Left$(cMsg, 10) = BotVars.Trigger & "sweepban " Or Left$(cMsg, 13) = BotVars.Trigger & "sweepignore " Then
                u = Trim(Mid$(Message, InStr(1, Message, " ", vbTextCompare)))
                
                If Left$(cMsg, 10) = BotVars.Trigger & "sweepban " Then
                    Y = "ban "
                Else
                    Y = "squelch "
                End If
                
                Caching = True
                Cache vbNullString, 255, Y
                AddQ "/who " & u, 1
                
                GoTo theEnd
                
            ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "setname ") Then
                If InBot Then
                    If Not g_Online = True Or g_Connected = False Then
                        GoTo theEnd
                    End If
                End If
                
                WriteINI "Main", "Username", Right(Message, Len(Message) - 9)
                BotVars.Username = Right(Message, Len(Message) - 9)
                strSend = "New username set."
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "setpass ") Then
                WriteINI "Main", "Password", Right(Message, Len(Message) - 9)
                BotVars.Password = Right(Message, Len(Message) - 9)
                strSend = "New password set."
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 8) = BotVars.Trigger & "setkey " Then
                
                u = Replace(Mid$(Message, 9), "-", vbNullString)
                u = Replace(u, " ", vbNullString)
            
                WriteINI "Main", "CDKey", u
                BotVars.CDKey = u
                strSend = "New cdkey set."
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 11) = BotVars.Trigger & "setexpkey " Then
                
                u = Replace(Mid$(Message, 11), "-", vbNullString)
                u = Replace(u, " ", vbNullString)
            
                WriteINI "Main", "LODKey", u
                BotVars.LODKey = u
                strSend = "New expansion CD-key set."
                b = True
                GoTo Display
                
                
            ElseIf Left$(cMsg, 11) = (BotVars.Trigger & "setserver ") Then
                WriteINI "Main", "Server", Right(Message, Len(Message) - 11)
                BotVars.Server = Right(Message, Len(Message) - 11)
                strSend = "New server set."
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 8) = BotVars.Trigger & "giveup " Or Left$(cMsg, 4) = BotVars.Trigger & "op " Then
                u = Right(cMsg, Len(Message) - InStr(1, Message, " ", vbTextCompare))
                If CheckChannel(u) > 0 Then
                    AddQ "/designate " & IIf(Dii, "*", "") & u
                    AddQ "/resign"
                End If
                
                GoTo theEnd
                
            ElseIf Left$(cMsg, 6) = BotVars.Trigger & "math " Or Left$(cMsg, 6) = BotVars.Trigger & "eval " Then
                On Error GoTo evalError
                'Math now has 3 levels.
                '50: No UI, and no CreateObject()
                '80: UI, no CreateObject()
                '100: No restrictions
                'Hdx - 09-25-07
                u = Mid$(cMsg, 7)
                If dbAccess.Access >= GetAccessINIValue("math80", 80) Then
                    If dbAccess.Access >= GetAccessINIValue("math100", 100) Then
                        frmChat.SCRestricted.AllowUI = True
                        strSend = frmChat.SCRestricted.Eval(u)
                    Else
                        If (InStr(LCase(u), "createobject") > 0) Then GoTo evalError
                        frmChat.SCRestricted.AllowUI = True
                        strSend = frmChat.SCRestricted.Eval(u)
                    End If
                Else
                    If (InStr(LCase(u), "createobject") > 0) Then GoTo evalError
                    frmChat.SCRestricted.AllowUI = False
                    strSend = frmChat.SCRestricted.Eval(u)
                End If
                
                While Left$(strSend, 1) = "/"
                    strSend = Mid$(strSend, 2)
                Wend
                
                b = True
                GoTo Display
                
evalError:
                strSend = "Evaluation error."
                b = True
                GoTo Display
                    
            ElseIf Left$(cMsg, 10) = BotVars.Trigger & "idlebans " Or Left$(cMsg, 4) = BotVars.Trigger & "ib " Then
                strArray = Split(Message, " ")
                
                If UBound(strArray) > 0 Then
                    Select Case strArray(1)
                        Case "on"
                            BotVars.IB_On = BTRUE
                            If UBound(strArray) > 1 Then
                                If StrictIsNumeric(strArray(2)) Then BotVars.IB_Wait = strArray(2)
                            End If
                            
                            If BotVars.IB_Wait > 0 Then
                                strSend = "IdleBans activated, with a delay of " & BotVars.IB_Wait & "."
                                WriteINI "Other", "IdleBans", "Y"
                            Else
                                BotVars.IB_Wait = 400
                                strSend = "IdleBans activated, using the default delay of 400."
                                WriteINI "Other", "IdleBanDelay", "400"
                                WriteINI "Other", "IdleBans", "Y"
                            End If
                            
                        Case "off"
                            BotVars.IB_On = BFALSE
                            strSend = "IdleBans deactivated."
                            WriteINI "Other", "IdleBans", "N"
                        
                        Case "wait", "delay"
                            If StrictIsNumeric(strArray(2)) Then
                                BotVars.IB_Wait = CInt(strArray(2))
                                strSend = "IdleBan delay set to " & BotVars.IB_Wait & "."
                                WriteINI "Other", "IdleBanDelay", CInt(strArray(2))
                            Else
                                strSend = "IdleBan delays require a numeric value."
                            End If
                            
                        Case "kick"
                            If UBound(strArray) > 1 Then
                                Select Case strArray(2)
                                    Case "on"
                                        strSend = "Idle users will now be kicked instead of banned."
                                        WriteINI "Other", "KickIdle", "Y"
                                        BotVars.IB_Kick = True
                                        
                                    Case "off"
                                        strSend = "Idle users will now be banned instead of kicked."
                                        WriteINI "Other", "KickIdle", "N"
                                        BotVars.IB_Kick = False
                                        
                                    Case Else
                                        strSend = "Unknown idle kick setting."
                                        
                                End Select
                            Else
                                strSend = "Not enough arguments were supplied."
                            End If
                            
                            b = True
                            GoTo Display
                            
                        Case "status"
                            If BotVars.IB_On = BTRUE Then
                                strSend = IIf(BotVars.IB_Kick, "Kicking", "Banning") & " users who are idle for " & BotVars.IB_Wait & "+ seconds."
                            Else
                                strSend = "IdleBans are disabled."
                            End If
                            b = True
                            GoTo Display
                        
                        Case Else
                            strSend = "Invalid IdleBan command."
                            
                    End Select
                Else
                    strSend = "Invalid IdleBan command arguments."
                End If
                    
                b = True
                GoTo Display
                
            ElseIf Left$(cMsg, 6) = BotVars.Trigger & "chpw " Then
                
                On Error GoTo chpwError
                strArray = Split(Message, " ")
                
                If UBound(strArray) > 0 Then
                    Select Case strArray(1)
                        Case "on", "set"
                            BotVars.ChannelPassword = strArray(2)
                            
                            If BotVars.ChannelPasswordDelay < 1 Then
                                BotVars.ChannelPasswordDelay = 30
                                strSend = "Channel password protection enabled, delay set to " & BotVars.ChannelPasswordDelay & "."
                            Else
                                strSend = "Channel password protection enabled."
                            End If
                            
                        Case "time", "delay", "wait"
                            If StrictIsNumeric(strArray(2)) Then
                                If Val(strArray(2)) < 256 Then
                                    BotVars.ChannelPasswordDelay = CByte(strArray(2))
                                    strSend = "Channel password delay set to " & strArray(2) & "."
                                Else
                                    strSend = "Channel password delays cannot be more than 255 seconds."
                                End If
                            Else
                                strSend = "Time setting requires a numeric value."
                            End If
                        Case "off", "kill", "clear"
                            BotVars.ChannelPassword = vbNullString
                            BotVars.ChannelPasswordDelay = 0
                            strSend = "Channel password protection disabled."
                            
                        Case "info", "status"
                            If BotVars.ChannelPassword = vbNullString Or BotVars.ChannelPasswordDelay = 0 Then
                                strSend = "Channel password protection is disabled."
                            Else
                                strSend = "Channel password protection is enabled. Password [" & BotVars.ChannelPassword & "], Delay [" & BotVars.ChannelPasswordDelay & "]."
                            End If
                            
                        Case Else
                            strSend = "Unknown channel password command."
                                
                    End Select
                    
                Else
chpwError:
                    strSend = "Error setting channel password."
                End If
                
                b = True
                GoTo Display
                    
                ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "join ") Then
                    u = Right(Message, (Len(Message) - 6))
                    
                    If LenB(u) > 0 Then
                        AddQ "/join " & u
                        GoTo theEnd
                    Else
                        strSend = "Join what channel?"
                        b = True
                        GoTo Display
                    End If
                    
                ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "sethome ") Then
                    u = Right(Message, (Len(Message) - 9))
                    WriteINI "Main", "HomeChan", u
                    BotVars.HomeChannel = u
                    strSend = "Home channel set to [ " & u & " ]"
                    b = True
                    GoTo Display
                
                ElseIf cMsg = (BotVars.Trigger & "resign") Then
                    AddQ "/resign", 1
                    GoTo theEnd
                    
                ElseIf cMsg = BotVars.Trigger & "cbl" Or cMsg = BotVars.Trigger & "clearbanlist" Then
                    ReDim gBans(0)
                    strSend = "Banned user list cleared."
                    
                    b = True
                    GoTo Display
                    
                ElseIf Left$(cMsg, 5) = BotVars.Trigger & "koy " Or Left$(cMsg, 12) = BotVars.Trigger & "kickonyell " Then
                    If Right(cMsg, 2) = "on" Then
                        BotVars.KickOnYell = 1
                        strSend = "Kick-on-yell enabled."
                        
                    ElseIf Right(cMsg, 3) = "off" Then
                        BotVars.KickOnYell = 0
                        strSend = "Kick-on-yell disabled."
                        
                    ElseIf Right(cMsg, 6) = "status" Then
                        strSend = "Kick-on-yell is "
                        strSend = strSend & IIf(BotVars.KickOnYell = 1, "enabled", "disabled") & "."
                    
                    End If
                    
                    b = True
                    GoTo Display
            
                ElseIf cMsg = (BotVars.Trigger & "rejoin") Or cMsg = BotVars.Trigger & "rj" Then
                    AddQ "/join " & CurrentUsername & " Rejoin", 1
                    AddQ "/join " & gChannel.Current, 1
                    GoTo theEnd
                        
                ElseIf (cMsg = BotVars.Trigger & "home" Or cMsg = BotVars.Trigger & "joinhome") Then
                    AddQ "/join " & BotVars.HomeChannel, 1
                    GoTo theEnd
                        
                ElseIf Left$(cMsg, 9) = BotVars.Trigger & "plugban " Then
                    If Right(Message, 2) = "on" Then
                        If BotVars.PlugBan Then
                            strSend = "PlugBan is already activated."
                        Else
                            BotVars.PlugBan = True
                            strSend = "PlugBan activated."
                            
                            For i = 1 To colUsersInChannel.Count
                                With colUsersInChannel.Item(i)
                                    If .Flags = 16 And Not .Safelisted Then
                                        AddQ "/ban " & IIf(Dii, "*", "") & .Username & " PlugBan", 1
                                    End If
                                End With
                            Next i
                        End If
                        
        ElseIf Right(Message, 3) = "off" Then
            If BotVars.PlugBan Then
                BotVars.PlugBan = False
                strSend = "PlugBan deactivated."
            Else
                strSend = "PlugBan is already deactivated."
            End If
        
        ElseIf Right(Message, 6) = "status" Then
            If BotVars.PlugBan Then
                strSend = "PlugBan is activated."
            Else
                strSend = "PlugBan is deactivated."
            End If
        Else
            GoTo theEnd
        End If
        
        b = True
        GoTo Display
            
    ElseIf cMsg = (BotVars.Trigger & "clist") Or cMsg = (BotVars.Trigger & "clientbans") Or cMsg = BotVars.Trigger & "cbans" Then
    
        strSend = "Clientbans: "
        For i = LBound(ClientBans) To UBound(ClientBans)
            If ClientBans(i) <> vbNullString Then
                strSend = strSend & ", " & ClientBans(i)
                If Len(strSend) > 90 Then
                    strSend = Replace(strSend, " , ", " ") & " [more]"
                    If WhisperCmds And Not InBot Then
                        If Dii Then AddQ "/w *" & Username & Space(1) & strSend Else AddQ "/w " & Username & Space(1) & strSend
                    ElseIf InBot Then
                        frmChat.AddChat RTBColors.ConsoleText, strSend
                    Else
                        AddQ strSend
                    End If
                End If
            End If
        Next i

        If strSend = "Clientbans: " Then strSend = "There are currently no ClientBans."
        strSend = Replace(strSend, " , ", " ")
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 8) = BotVars.Trigger & "setvol " Then
        If Not BotVars.DisableMP3Commands Then
            On Error GoTo VolumeError
            u = Right(cMsg, Len(cMsg) - 8)
            If StrictIsNumeric(u) Then
                hWndWA = GetWinamphWnd()
                If hWndWA = 0 Then
                    strSend = "Winamp is not loaded."
                    b = True
                    GoTo Display
                End If
                If CInt(u) > 100 Then u = 100
                SendMessage hWndWA, WM_WA_IPC, 2.55 * CInt(u), 122
                strSend = "Volume set to " & u & "%."
            Else
VolumeError:
                strSend = "Invalid volume level (0-100)."
                b = True
                GoTo Display
            End If
            b = True
            GoTo Display
        End If
        
    ElseIf Left$(cMsg, 6) = BotVars.Trigger & "cadd " Then
        
        banMsg = UCase(ReadCFG("Other", "ClientBans"))
        u = UCase(Mid$(cMsg, 7))
        
        If Right(banMsg, 1) = " " Then
            banMsg = banMsg & u
        Else
            banMsg = banMsg & Space(1) & u
        End If
        WriteINI "Other", "ClientBans", UCase(banMsg)
        
        strArray() = Split(u, " ")
        For i = LBound(strArray) To UBound(strArray)
            ReDim Preserve ClientBans(0 To UBound(ClientBans) + 1)
            ClientBans(UBound(ClientBans)) = UCase(strArray(i))
        Next i
        
        strSend = "Added clientban(s): " & UCase(u)
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 6) = BotVars.Trigger & "cdel " Or Left$(cMsg, 11) = BotVars.Trigger & "delclient " Then
        If Left$(cMsg, 6) = BotVars.Trigger & "cdel " Then
            u = LCase(Right(cMsg, Len(cMsg) - 6))
        Else
            u = LCase(Right(cMsg, Len(cMsg) - 11))
        End If
        
        For c = LBound(ClientBans) To UBound(ClientBans)
            banMsg = banMsg & LCase(ClientBans(c)) & " "
        Next c
        banMsg = Replace(banMsg, u, vbNullString)
        WriteINI "Other", "ClientBans", Replace(banMsg, "  ", vbNullString)
        
        ClientBans() = Split(ReadCFG("Other", "ClientBans"), " ")
        If UBound(ClientBans) = -1 Then ReDim ClientBans(0)
        
        strSend = "Clientban """ & UCase(u) & """ deleted."
        b = True
        GoTo Display
        
        strSend = "Client is not banned."
        b = True
        GoTo Display
            
    ElseIf cMsg = BotVars.Trigger & "banned" Then
    
        If UBound(gBans) = 0 And gBans(0).Username = vbNullString Then
            strSend = "No users have been banned."
            b = True
            GoTo Display
        End If
        
        strSend = "Banned users: "
        
        For i = LBound(gBans) To UBound(gBans)
            If gBans(i).Username <> vbNullString Then
            
                strSend = strSend & ", " & gBans(i).Username
                
                If Len(strSend) > 90 And i <> UBound(gBans) Then
                    strSend = Replace(strSend, " , ", " ") & " [more]"
                    
                    FilteredSend Username, strSend, WhisperCmds, InBot, PublicOutput
                    
                    strSend = "Banned users: "
                End If
                
            End If
        Next i
        
        strSend = Replace(strSend, " , ", " ")
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "ipbans ") Then
    
        If Right(cMsg, 2) = "on" Then
            BotVars.IPBans = True
            WriteINI "Other", "IPBans", "Y"
            strSend = "IPBanning activated."
            
            If MyFlags = 2 Or MyFlags = 18 Then
                For i = 1 To colUsersInChannel.Count
                    Select Case colUsersInChannel.Item(i).Flags
                        Case 20, 30, 32, 48
                            AddQ "/ban " & IIf(Dii, "*", "") & colUsersInChannel.Item(i).Username & " IPBanned.", 1
                    End Select
                Next i
            End If
            
        ElseIf Right(cMsg, 3) = "off" Then
            BotVars.IPBans = False
            WriteINI "Other", "IPBans", "N"
            strSend = "IPBanning deactivated."
            
        ElseIf Right(cMsg, 6) = "status" Then
            If BotVars.IPBans Then
                strSend = "IPBanning is currently active."
            Else
                strSend = "IPBanning is currently disabled."
            End If
            
        Else
            strSend = "Unrecognized IPBan command. Use 'on', 'off' or 'status'."
        End If
        
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 7) = BotVars.Trigger & "ipban " Then
        
        u = Mid$(cMsg, 8)
        
        strUser = StripInvalidNameChars(u)
        
        'If InStr(1, CleanedUsername, "@") > 0 Then CleanedUsername = StripRealm(CleanedUsername)
        
        If dbAccess.Access < 101 Then
            If GetSafelist(strUser) Or GetSafelist(u) Then
                strSend = "That user is safelisted."
                b = True
                GoTo Display
            End If
        End If
        
        gAcc = GetAccess(u)
        
        If gAcc.Access >= dbAccess.Access Or (InStr(gAcc.Flags, "A") > 0 And dbAccess.Access < 101) Then
            strSend = "You do not have enough access to do that."
            b = True
            GoTo Display
        End If
        
        If Len(cMsg) = 7 Then
            strSend = "IPBan who?"
            b = True
            GoTo Display
        End If
        
        AddQ "/squelch " & IIf(Dii, "*", "") & Right(cMsg, Len(cMsg) - 7), 1
        strSend = "User " & Chr(34) & Right(Message, Len(Message) - 7) & Chr(34) & " IPBanned."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = BotVars.Trigger & "unipban " Then
        If Len(cMsg) = 9 Then
            strSend = "Un-IPBan who?"
            b = True
            GoTo Display
        End If
        
        AddQ "/unsquelch " & IIf(Dii, "*", "") & Right(cMsg, Len(cMsg) - 9), 1
        AddQ "/unban " & IIf(Dii, "*", "") & Right(cMsg, Len(cMsg) - 9), 1
        strSend = "User " & Chr(34) & Right(Message, Len(Message) - 9) & Chr(34) & " Un-IPBanned."
        b = True
        GoTo Display
                
    ElseIf Left$(cMsg, 11) = (BotVars.Trigger & "designate ") Or Left$(cMsg, 5) = (BotVars.Trigger & "des ") Then
        On Error GoTo sendr
        
        If (MyFlags And &H2) = &H2 Then
            If Left$(cMsg, 11) = BotVars.Trigger & "designate " Then
                u = Right(Message, (Len(Message) - 11))
            Else
                u = Mid$(Message, 6)
            End If
            
            'diablo 2 handling
            If Dii = True Then
                If Not (Mid$(u, 1, 1) = "*") Then u = "*" & u
            End If
            
            AddQ "/designate " & u, 1
            strSend = "I have designated [ " & u & " ]"
        Else
            strSend = "The bot does not have ops."
        End If
        
        b = True
        GoTo Display
sendr:
        strSend = "Designate who?"
        b = True
        GoTo Display
    
    ElseIf cMsg = (BotVars.Trigger & "shuffle") Then
        If Not BotVars.DisableMP3Commands Then
            strSend = "Winamp's Shuffle feature has been toggled."
            hWndWA = GetWinamphWnd()
            
            If hWndWA = 0 Then
                strSend = "Winamp is not loaded."
            Else
                SendMessage hWndWA, WM_COMMAND, WA_TOGGLESHUFFLE, 0
            End If
            
            b = True
            GoTo Display
        End If
    
    ElseIf cMsg = (BotVars.Trigger & "repeat") Then
        If Not BotVars.DisableMP3Commands Then
            strSend = "Winamp's Repeat feature has been toggled."
            hWndWA = GetWinamphWnd()
            
            If hWndWA = 0 Then
                strSend = "Winamp is not loaded."
            Else
                SendMessage hWndWA, WM_COMMAND, WA_TOGGLEREPEAT, 0
            End If
            
            b = True
            GoTo Display
        End If
            
    ElseIf cMsg = (BotVars.Trigger & "next") Then
        If Not BotVars.DisableMP3Commands Then
            If iTunesReady Then
                On Error GoTo iTunesErr
                
                iTunesNext
                
                strSend = "Skipped forwards."
                
                b = True
                GoTo Display
                
iTunesErr:
                strSend = "There was an error convincing iTunes to obey your command. Please restart iTunes and try again."
                b = True
                GoTo Display
            Else
                On Error GoTo sendx
                hWndWA = GetWinamphWnd()
                
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                
                SendMessage hWndWA, WM_COMMAND, WA_NEXTTRACK, 0
                strSend = "Skipped forwards."
                b = True
                GoTo Display
sendx:
                strSend = "Error."
                b = True
                GoTo Display
            End If
        End If
        
    ElseIf cMsg = (BotVars.Trigger & "prev") Then
        If Not BotVars.DisableMP3Commands Then
            If iTunesReady Then
                iTunesBack
                strSend = "Skipped backwards."
                b = True
                GoTo Display
            Else
                On Error GoTo sendj
                
                hWndWA = GetWinamphWnd()
                
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                
                SendMessage hWndWA, WM_COMMAND, WA_PREVTRACK, 0
                strSend = "Skipped backwards."
                b = True
                GoTo Display
sendj:
                strSend = "Error."
                b = True
                GoTo Display
            End If
        End If
                
    ElseIf cMsg = (BotVars.Trigger & "protect on") Then
            If MyFlags = 2 Or MyFlags = 18 Then
                Protect = True
                strSend = "Lockdown activated by " & Username & "."
                WildCardBan "*", ProtectMsg, 1
                WriteINI "Main", "Protect", "Y"
            Else
                strSend = "The bot does not have ops."
            End If
            b = True
            GoTo Display
            
    ElseIf cMsg = (BotVars.Trigger & "protect status") Then
            Select Case Protect
                Case True: strSend = "Lockdown is currently active."
                Case Else: strSend = "Lockdown is currently disabled."
            End Select
            
            b = True
            GoTo Display
       
    ElseIf cMsg = (BotVars.Trigger & "protect off") Then
        If Protect Then
            Protect = False
            strSend = "Lockdown deactivated."
            WriteINI "Main", "Protect", "N"
        Else
            strSend = "Protection was not enabled."
        End If
        b = True
        GoTo Display
            
    ElseIf cMsg = (BotVars.Trigger & "whispercmds") Or cMsg = (BotVars.Trigger & "wc") Then
        If BotVars.WhisperCmds Then
            BotVars.WhisperCmds = False
            WhisperCmds = False
            
            WriteINI "Main", "WhisperBack", "N"
            strSend = "Command responses will be displayed publicly."
        Else
            BotVars.WhisperCmds = True
            WhisperCmds = True
                
            WriteINI "Main", "WhisperBack", "Y"
            strSend = "Command responses will be whispered back."
        End If
        
        b = True
        GoTo Display
       
    ElseIf cMsg = (BotVars.Trigger & "stop") Then
        If Not BotVars.DisableMP3Commands Then
            If iTunesReady Then
                iTunesStop
                strSend = "iTunes playback stopped."
                
                b = True
                GoTo Display
            Else
                On Error GoTo sendxc
                
                hWndWA = GetWinamphWnd()
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                
                SendMessage hWndWA, WM_COMMAND, WA_STOP, 0
                strSend = "Stopped play."
                
                b = True
                GoTo Display
sendxc:
                strSend = "Error."
                b = True
                GoTo Display
            End If
        End If
                
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "play ") Then
        On Error GoTo WAPlayError
        
        If Not BotVars.DisableMP3Commands Then
            If iTunesReady Then
                
                iTunesPlayFile Mid$(Message, 7)
                
                strSend = "Attempted to play the specified filepath."
                b = True
                GoTo Display
                
            Else
                hWndWA = GetWinamphWnd()
                
                If hWndWA = 0 Then
                    strSend = "Winamp is stopped, or isn't running."
                    b = True
                    GoTo Display
                End If
                
                If StrictIsNumeric(Right(Message, Len(Message) - 6)) Then
                    Track = Right(Message, Len(Message) - 6)
                    SendMessage hWndWA, WM_COMMAND, WA_STOP, 0
                    SendMessage hWndWA, WM_USER, Track - 1, 121
                    SendMessage hWndWA, WM_COMMAND, WA_PLAY, 0
                    
                    strSend = "Skipped to track " & Track & "."
                    b = True
                    GoTo Display
                Else
                    WinampJumpToFile Right(Message, Len(Message) - 6)
                    GoTo theEnd
                End If
WAPlayError:
                strSend = "Error playing that track/song."
                b = True
                GoTo Display
            End If
        End If
        
    ElseIf cMsg = (BotVars.Trigger & "useitunes") Then
        
        If iTunesReady Then
            strSend = "iTunes is already ready."
        Else
            If InitITunes Then
                strSend = "iTunes is ready."
            Else
                strSend = "Error launching iTunes."
            End If
        End If
        
        b = True
        GoTo Display
        
    ElseIf cMsg = (BotVars.Trigger & "usewinamp") Then
        
        If iTunesReady Then
            strSend = "Returning to Winamp control."
            iTunesUnready
        Else
            strSend = "iTunes was not ready."
        End If
        
        b = True
        GoTo Display

    ElseIf cMsg = (BotVars.Trigger & "play") Then
    
        If Not BotVars.DisableMP3Commands Then
            On Error GoTo WAPlayError2
            
            If iTunesReady Then
                iTunesPlay
                strSend = "iTunes playback started."
            Else
                hWndWA = GetWinamphWnd()
                
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                
                SendMessage hWndWA, WM_COMMAND, WA_PLAY, 0
                
                strSend = "Skipped backwards."
                
                If iWinamp = 0 Then
                    strSend = "Play started."
                Else
                    strSend = "Error sending your command to Winamp. Make sure it's running."
                End If
            End If
            
            b = True
            GoTo Display
WAPlayError2:
            strSend = "Error."
            b = True
            GoTo Display
        End If
        
    ElseIf cMsg = (BotVars.Trigger & "pause") Then
    
        If Not BotVars.DisableMP3Commands Then
            If iTunesReady Then
                iTunesPause
                strSend = "Pause toggled."
                b = True
                GoTo Display
            Else
                On Error GoTo sendn
                
                hWndWA = GetWinamphWnd()
                
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                
                SendMessage hWndWA, WM_COMMAND, WA_PAUSE, 0
                strSend = "Paused/resumed play."
                b = True
                GoTo Display
sendn:
                strSend = "Error."
                b = True
                GoTo Display
            End If
        End If
        
    ElseIf cMsg = (BotVars.Trigger & "fos") Then
        If Not BotVars.DisableMP3Commands Then
            On Error GoTo sendm
                hWndWA = GetWinamphWnd()
                If hWndWA = 0 Then
                   strSend = "Winamp is not loaded."
                   b = True
                   GoTo Display
                End If
                SendMessage hWndWA, WM_COMMAND, WA_FADEOUTSTOP, 0
                strSend = "Fade-out stop."
                b = True
                GoTo Display
sendm:
                strSend = "Error."
                b = True
                GoTo Display
        End If
        
    ElseIf Left$(cMsg, 5) = (BotVars.Trigger & "rem ") Or Left$(cMsg, 5) = BotVars.Trigger & "del " Then
    'On Error GoTo sendz
        u = Right(Message, (Len(Message) - 5))
        If GetAccess(u).Access >= dbAccess.Access Then
            strSend = "That user has higher or equal access."
            b = True
            GoTo Display
        End If
        
        If InStr(1, GetAccess(u).Flags, "L") > 0 Then
            If InStr(1, GetAccess(Username).Flags, "A") = 0 And GetAccess(Username).Access < 100 And Not InBot Then
                strSend = "That user is Locked."
                b = True
                GoTo Display
            End If
        End If
        
        strSend = RemoveItem(u, "users")
        strSend = Replace(strSend, "%msgex%", "userlist entry")
        
        If InStr(strSend, "Successfully") Then
            If BotVars.LogDBActions Then
                LogDBAction RemEntry, Username, u, Message
            End If
        End If
        
        Call LoadDatabase
        
        b = True
        GoTo Display
sendz:
        strSend = "Remove what user?"
        b = True
        GoTo Display

    ElseIf cMsg = (BotVars.Trigger & "reconnect") Then
        If g_Online Then
            BotVars.HomeChannel = gChannel.Current
            
            Call frmChat.DoDisconnect
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Reconnecting by command, please wait..."
            
            Pause 1
            
            frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
            
            Call frmChat.DoConnect
        Else
            frmChat.AddChat RTBColors.ErrorMessageText, "You must be online to reconnect. Try connecting first."
        End If
        
        GoTo theEnd
        
    ElseIf cMsg = (BotVars.Trigger & "unigpriv") Then
        AddQ "/o unigpriv", 1
        strSend = "Recieving text from non-friends."
        b = True
        GoTo Display
        
    ElseIf cMsg = (BotVars.Trigger & "igpriv") Then
        AddQ "/o igpriv", 1
        strSend = "Ignoring text from non-friends."
        b = True
        GoTo Display

    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "block ") Then
        u = Right(Message, Len(Message) - 7)
        z = ReadINI("BlockList", "Total", "filters.ini")
        
        If StrictIsNumeric(z) Then
            i = z
        Else
            WriteINI "BlockList", "Total", "Total=0", "filters.ini"
            i = 0
        End If
        WriteINI "BlockList", "Filter" & (i + 1), u, "filters.ini"
        WriteINI "BlockList", "Total", i + 1, "filters.ini"
        strSend = "Added """ & u & """ to the username block list."
        b = True
        GoTo Display

    ElseIf Left$(cMsg, 10) = (BotVars.Trigger & "idletime ") Or Left$(cMsg, 10) = BotVars.Trigger & "idlewait " Then
        u = Right(Message, Len(Message) - 10)
        
        If Not StrictIsNumeric(u) Or Val(u) > 50000 Then
            strSend = "Error setting idle wait time."
        Else
            WriteINI "Main", "IdleWait", 2 * Int(u)
            strSend = "Idle wait time set to " & Int(u) & " minutes."
        End If
        
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "idle ") Then
        On Error GoTo IdleError
        
        u = Split(Message, " ")(1)
        
        If LCase(u) = "on" Then
            WriteINI "Main", "Idles", "Y"
            strSend = "Idles activated."
        ElseIf LCase(u) = "off" Then
            WriteINI "Main", "Idles", "N"
            strSend = "Idles deactivated."
        ElseIf LCase(u) = "kick" Then
            u = Split(Message, " ")(2)
            
            If LCase(u) = "on" Then
                BotVars.IB_Kick = True
                strSend = "Idle kick is now enabled."
            ElseIf LCase(u) = "off" Then
                BotVars.IB_Kick = False
                strSend = "Idle kick disabled."
            Else
                strSend = "Unknown idle kick command."
            End If
        Else
            GoTo IdleError
        End If
        
        b = True
        GoTo Display
IdleError:
        strSend = "Error setting idles. Make sure you used '.idle on' or '.idle off'."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "shitdel ") Then
        On Error GoTo IdleError23
        u = Right(Message, Len(Message) - 9)
IdleError23:
        If MyFlags = 2 Or MyFlags = 18 Then AddQ "/unban " & IIf(Dii, "*", "") & u, 1
        strSend = RemoveItem(u, "autobans")
        strSend = Replace(strSend, "%msgex%", "shitlist")
        
        If InStr(strSend, "Successfully") Then
            If BotVars.LogDBActions Then
                LogDBAction RemEntry, Username, u, Message
            End If
        End If
        
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "safedel ") Then
        u = GetStringChunk(cMsg, 2)
        
        b = RemoveFromSafelist(u)
        
        If b Then
            strSend = "That user has been removed from the safelist."
        Else
            b = True
            strSend = "That user is not safelisted, or there was an error removing them."
        End If
        
        GoTo Display
        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "tagdel ") Then
        On Error GoTo Error613
        u = Right(Message, Len(Message) - 8)
        strSend = RemoveItem(u, "tagbans")
        strSend = Replace(strSend, "%msgex%", "tagban")
        b = True
        GoTo Display
        
Error613:
        strSend = "Delete what tag?"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "profile ") Then
        On Error GoTo Error614
        
        u = Right(Message, Len(Message) - 9)
        PPL = True
        
        If (BotVars.WhisperCmds Or WhisperedIn) And Not PublicOutput Then
            PPLRespondTo = Username
        End If
        
        RequestProfile u
        
        GoTo theEnd
        
Error614:
        strSend = "What profile would you like to look up?"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "setidle ") Then
        On Error GoTo IdleError133
        
        u = Right(Message, Len(Message) - 9)
        
        If Left$(u, 1) = "/" Then
            u = " " & u
        End If
        
        WriteINI "Main", "IdleMsg", u
        
        strSend = "Idle message set."
        b = True
        GoTo Display
IdleError133:
        strSend = "What do you want the idle message set to?"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 10) = (BotVars.Trigger & "idletype ") Then
        On Error GoTo IdleError2
        
        u = Right(Message, Len(Message) - 10)
        
        If LCase(u) = "msg" Or LCase(u) = "message" Then
            WriteINI "Main", "IdleType", "msg"
            strSend = "Idle type set to [ msg ]."
        ElseIf LCase(u) = "quote" Or LCase(u) = "quotes" Then
            WriteINI "Main", "IdleType", "quote"
            strSend = "Idle type set to [ quote ]."
        ElseIf LCase(u) = "uptime" Then
            WriteINI "Main", "IdleType", "uptime"
            strSend = "Idle type set to [ uptime ]."
        ElseIf LCase(u) = "mp3" Then
            WriteINI "Main", "IdleType", "mp3"
            strSend = "Idle type set to [ mp3 ]."
        Else
            GoTo IdleError2
        End If
        
        b = True
        GoTo Display
IdleError2:
        strSend = "Error setting idle type. The types are [ message quote uptime mp3 ]."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "filter ") Then
        u = Right(Message, Len(Message) - 8)
        z = ReadINI("TextFilters", "Total", "filters.ini")
        
        If StrictIsNumeric(z) Then
            i = z
        Else
            WriteINI "TextFilters", "Total", "Total=0", "filters.ini"
            i = 0
        End If
        
        WriteINI "TextFilters", "Filter" & (i + 1), u, "filters.ini"
        WriteINI "TextFilters", "Total", i + 1, "filters.ini"
        
        ReDim Preserve gFilters(UBound(gFilters) + 1)
        gFilters(UBound(gFilters)) = u
        
        strSend = "Added " & Chr(34) & u & Chr(34) & " to the text message filter list."
        b = True
        GoTo Display
        
    ElseIf cMsg = "?trigger" And InBot = False Then
        strSend = "The bot's current trigger is " & Chr(34) & Space(1) & BotVars.Trigger & Space(1) & Chr(34) & " (Alt + 0" & Asc(BotVars.Trigger) & ")"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 12) = (BotVars.Trigger & "settrigger ") Then
        On Error GoTo sendg
        
        u = LCase(Mid$(Message, 13, 1))
        
        If StrComp(u, " ") = 0 Then
            strSend = "Sorry, you can't set your trigger to a blank space."
        Else
            OldTrigger = u
            WriteINI "Main", "Trigger", u
        
            strSend = "The new trigger is " & Chr(34) & OldTrigger & Chr(34) & "."
        End If
        
        b = True
        GoTo Display
sendg:
        strSend = "Change to what trigger?"
        b = True
        GoTo Display
                
    ElseIf Left$(cMsg, 10) = BotVars.Trigger & "levelban " Then
        On Error GoTo Error231
        
        If StrictIsNumeric(Right(cMsg, Len(cMsg) - 10)) Then
            i = Val(Right(cMsg, Len(cMsg) - 10))
            
            If i > 0 Then
                strSend = "Banning Warcraft III users under level " & i & "."
                BotVars.BanUnderLevel = i
            Else
                strSend = "Levelbans disabled."
                BotVars.BanUnderLevel = 0
            End If
        Else
            BotVars.BanUnderLevel = 0
            strSend = "Levelbans disabled."
        End If
        
        WriteINI "Other", "BanUnderLevel", BotVars.BanUnderLevel
        
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 12) = BotVars.Trigger & "d2levelban " Then
        On Error GoTo Error231
        
        If StrictIsNumeric(Right(cMsg, Len(cMsg) - 12)) Then
            i = Val(Right(cMsg, Len(cMsg) - 12))
            BotVars.BanD2UnderLevel = i
            
            If i > 0 Then
                strSend = "Banning Diablo II characters under level " & i & "."
                BotVars.BanD2UnderLevel = i
            Else
                strSend = "Diablo II Levelbans disabled."
                BotVars.BanD2UnderLevel = 0
            End If
        Else
            strSend = "Diablo II Levelbans disabled."
            BotVars.BanD2UnderLevel = 0
        End If
        
        WriteINI "Other", "BanD2UnderLevel", BotVars.BanD2UnderLevel
        
        b = True
        GoTo Display

        
Error231:       strSend = "Error setting Levelban level."
        b = True
        GoTo Display
        
    ElseIf cMsg = BotVars.Trigger & "pon" Or cMsg = BotVars.Trigger & "phrasebans on" Then
        WriteINI "Other", "Phrasebans", "Y"
        Phrasebans = True
        strSend = "Phrasebans activated."
        b = True
        GoTo Display
        
    ElseIf cMsg = BotVars.Trigger & "poff" Or cMsg = BotVars.Trigger & "phrasebans off" Then
        WriteINI "Other", "Phrasebans", "N"
        Phrasebans = False
        strSend = "Phrasebans deactivated."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "cbans ") Then
        
        On Error GoTo cbansError
        ' on/off/status
        u = Mid$(cMsg, 8)
        
        Select Case u
            Case "on"
                strSend = "ClientBans enabled."
                BotVars.ClientBans = True
                WriteINI "Other", "ClientBansOn", "Y"
                
            Case "off"
                strSend = "ClientBans disabled."
                BotVars.ClientBans = False
                WriteINI "Other", "ClientBansOn", "N"
                
            Case "status"
                strSend = "ClientBans are currently " & IIf(BotVars.ClientBans, "enabled.", "disabled.")
                
            Case Else
                GoTo cbansError
        End Select
        
        b = True
        GoTo Display
cbansError:
        strSend = "What do you want to do to your ClientBans?"
        b = True
        GoTo Display
    
    ElseIf cMsg = BotVars.Trigger & "pstatus" Or cMsg = BotVars.Trigger & "phrasebans" Then
        If Phrasebans = True Then
            strSend = "Phrasebans are enabled."
        Else
            strSend = "Phrasebans are disabled."
        End If
        b = True
        GoTo Display

    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "mimic ") Then
        On Error GoTo sendit2t
        u = Right(Message, (Len(Message) - 7))
        Mimic = LCase(u)
        strSend = "Mimicking [ " & u & " ]"
        b = True
        GoTo Display
       
sendit2t:
        strSend = "Mimic who?"
        b = True
        GoTo Display
        
    ElseIf cMsg = (BotVars.Trigger & "nomimic") Then
        Mimic = vbNullString
        strSend = "Mimic off."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = BotVars.Trigger & "setpmsg " Then
        u = Right(Message, Len(Message) - 9)
        ProtectMsg = u
        WriteINI "Other", "ProtectMsg", u
        strSend = "Channel protection message set."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 14) = BotVars.Trigger & "setcmdaccess " Then
        c = 1
        
        If Not StrictIsNumeric(GetStringChunk(cMsg, 3)) Then
            strSend = "You must specify a numeric value to change a command's required access."
            b = True
            GoTo Display
        End If
        
        iWinamp = Val(GetStringChunk(cMsg, 3))
        z = GetStringChunk(cMsg, 2)
        
        If iWinamp > 999 Then
            strSend = "Your new required access must be between 0 and 999."
            b = True
            GoTo Display
        End If
        
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
        
        While Track = 0 And Not EOF(f)
            Get #f, c, ccOut
                        
            If StrComp(Trim(ccOut.Query), z) = 0 Then
                If (dbAccess.Access < 100 And InStr(dbAccess.Flags, "A") = 0) And _
                    ccOut.reqAccess >= dbAccess.Access And _
                    ccOut.reqAccess > iWinamp Then
                    
                    strSend = "You cannot decrease the required access on a command without 100 access or the A flag."
                Else
                    ccOut.reqAccess = iWinamp
                  
                    Put #f, c, ccOut
                
                    strSend = "Command modified successfully."
                End If
                
                Track = 1
            End If
            
            c = c + 1
        Wend
        
        Close #f
        
        If Track = 0 Then
            strSend = "That command was not found."
        End If
        
        b = True
        GoTo Display
    
    
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "cmdadd ") Or Left$(cMsg, 8) = (BotVars.Trigger & "addcmd ") Then
    
        On Error GoTo cmdAddError
    
        If InStr(1, cMsg, "/add ", vbTextCompare) > 0 Or _
            InStr(1, cMsg, "/rem ", vbTextCompare) > 0 Or _
                InStr(1, cMsg, "/set ", vbTextCompare) > 0 Then
                    strSend = "You cannot use '/add' or '/rem' in a custom command."
                    b = True
                    GoTo Display
        End If
        
        If InStr(1, cMsg, "/quit", vbTextCompare) > 0 Then
            strSend = "You cannot use '/quit' in a custom command."
            b = True
            GoTo Display
        End If
        
        If InStr(1, cMsg, "/shitadd", vbTextCompare) > 0 Or _
            InStr(1, cMsg, "/pban", vbTextCompare) > 0 Or _
                InStr(1, cMsg, "/shitlist", vbTextCompare) > 0 Then
                
            If dbAccess.Access < 100 And InStr(dbAccess.Flags, "A") = 0 Then
                strSend = "Shitlisting users through custom commands requires 100+ access."
                b = True
                GoTo Display
            End If
        End If
        
        gAcc.Access = 1000
        gAcc.Flags = "A"
        
        If ProcessCC(Username, gAcc, "/" & Split(cMsg, " ")(2), False, False, True) Then
            strSend = "A command by that name already exists."
            b = True
            GoTo Display
        End If
        
        gAcc.Access = 0
        gAcc.Flags = ""
                
        '0      1  2    3
        'cmdadd 10 fdsa actions
        strArray() = Split(Message, " ", 4)
        
        If Not StrictIsNumeric(strArray(1)) Then
            strSend = "Command format error."
            b = True
            GoTo Display
        End If
        
        If UBound(strArray) > 2 Then
            If Len(strArray(3)) < 1 Then
                strSend = "Your command's actions cannot be blank."
                b = True
                GoTo Display
            End If
        Else
            strSend = "Your command's actions cannot be blank."
            b = True
            GoTo Display
        End If
        
        ccOut.Query = strArray(2)
        ccOut.reqAccess = Int(strArray(1))
        
        ccOut.Action = strArray(3)
        
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
            c = LOF(f) \ LenB(ccOut)
            
            If LOF(f) Mod LenB(ccOut) <> 0 Then c = c + 1
            If c = 0 Then c = 1
            
            Put #f, c + 1, ccOut
        Close #f
        
        strSend = "Command " & Chr(34) & strArray(2) & Chr(34) & " added."
        b = True
        GoTo Display
        
cmdAddError:
        strSend = "Error adding your command."
        b = True
        GoTo Display
        
        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "cmddel ") Or Left$(cMsg, 8) = (BotVars.Trigger & "delcmd ") Then
        
        u = Right(cMsg, Len(cMsg) - 8)
        
        If Dir$((GetFilePath("commands.dat"))) = vbNullString Then
            strSend = "No commands list exists."
            b = True
            GoTo Display
        End If
        
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
            If LOF(f) < 2 Then GoTo theEnd
            c = LOF(f) \ LenB(ccOut)
            If LOF(f) Mod LenB(ccOut) <> 0 Then c = c + 1
            
            For c = 1 To c
                Get #f, c, ccOut
                If ccOut.reqAccess > 1000 Then GoTo NextRecord2

                If StrComp(LCase(RTrim(ccOut.Query)), u, vbTextCompare) = 0 Then
                    ccOut.reqAccess = 9999
                    Put #f, c, ccOut
                    strSend = "Command " & Chr(34) & u & Chr(34) & " deleted."
                End If
NextRecord2:
            Next c
        Close #f
        
        If LenB(strSend) = 0 Then
            strSend = "No such command exists."
        End If
        
        b = True
        GoTo Display
        
    ElseIf cMsg = BotVars.Trigger & "cmdlist" Then
        
        If Dir$((GetFilePath("commands.dat"))) = vbNullString Then
            strSend = "No custom commands available."
            b = True
            GoTo Display
        End If
        
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccOut)
        
        If LOF(f) < 2 Then
            strSend = "No custom commands available."
            b = True
            GoTo Display
        End If
        
        n = LOF(f) \ LenB(ccOut)
        If LOF(f) Mod LenB(ccOut) <> 0 Then n = n + 1
        
        Response = "Found commands: "
        
        For c = 1 To n
            Get #f, c, ccOut
            If ccOut.reqAccess < 1001 And Len(RTrim(ccOut.Query)) > 0 Then
            
                Response = Response & ", " & RTrim(KillNull(ccOut.Query)) & " [" & ccOut.reqAccess & "]"
                
                If Len(Response) > 100 Then
                    Response = Replace(Response, " , ", " ")
                    
                    If c < n Then
                        Response = Response & " [more]"
                    End If
                    
                    If WhisperCmds And Not InBot Then
                        If Not Dii Then AddQ "/w " & Username & Space(1) & Response Else AddQ "/w *" & Username & Response
                    ElseIf InBot And Not PublicOutput Then
                        frmChat.AddChat RTBColors.ConsoleText, Response
                    Else
                        AddQ Response
                    End If
                    
                    Response = "Found commands: "
                End If
                
                Track = 1
            
            End If
        Next c
        
        strSend = Replace(Response, " , ", " ") & "."
        b = True
        
        If StrComp(strSend, "Found commands: .", vbBinaryCompare) = 0 Then
            If Track = 0 Then
                strSend = "No custom commands available."
            Else
                GoTo theEnd
            End If
        End If
        
        GoTo Display
        
    ElseIf cMsg = BotVars.Trigger & "phrases" Or cMsg = BotVars.Trigger & "plist" Then
        
        If UBound(Phrases) = 0 And Phrases(0) = vbNullString Then
            strSend = "There are no phrasebans."
            b = True
            GoTo Display
        End If
        Response = "Phraseban(s): "
        
        For c = LBound(Phrases) To UBound(Phrases)
            If Phrases(c) <> " " And Phrases(c) <> vbNullString Then
                Response = Response & ", " & Phrases(c)
                If Len(Response) > 89 Then
                    Response = Replace(Response, " , ", " ") & " [more]"
                    If WhisperCmds And Not InBot Then
                        AddQ "/w " & Username & Space(1) & Response
                    ElseIf InBot = True Then
                        frmChat.AddChat RTBColors.ConsoleText, Response
                    Else
                        AddQ Response
                    End If
                    Response = "Phraseban(s): "
                End If
            End If
        Next c
        
        strSend = Replace(Response, " , ", " ")
        b = True
        GoTo Display
    
    ElseIf Left$(cMsg, 11) = BotVars.Trigger & "addphrase " Or Left$(cMsg, 6) = BotVars.Trigger & "padd " Then
        
        If InStr(1, cMsg, BotVars.Trigger & "addphrase ", vbTextCompare) = 0 Then c = 6 Else c = 11
        u = Right(Message, Len(Message) - c)
        
        For i = LBound(Phrases) To UBound(Phrases)
            If StrComp(LCase(u), LCase(Phrases(i)), vbTextCompare) = 0 Then
                strSend = "That phrase is already banned."
                b = True
                GoTo Display
            End If
        Next i

        If Phrases(UBound(Phrases)) <> vbNullString Or Phrases(UBound(Phrases)) <> " " Then
            ReDim Preserve Phrases(0 To UBound(Phrases) + 1)
        End If
        Phrases(UBound(Phrases)) = u
        
        Open GetFilePath("phrasebans.txt") For Output As #f
        For c = LBound(Phrases) To UBound(Phrases)
            If Len(Phrases(c)) > 0 Then Print #f, Phrases(c)
        Next c
        
        strSend = "Phraseban " & Chr(34) & u & Chr(34) & " added."
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 11) = BotVars.Trigger & "delphrase " Or Left$(cMsg, 6) = BotVars.Trigger & "pdel " Then
        
        If InStr(1, cMsg, BotVars.Trigger & "delphrase ", vbTextCompare) = 0 Then c = 6 Else c = 11
        u = Right(Message, Len(Message) - c)
        
        Open GetFilePath("phrasebans.txt") For Output As #f
        Y = vbNullString
        For c = LBound(Phrases) To UBound(Phrases)
            If StrComp(Phrases(c), LCase(u), vbTextCompare) <> 0 Then
                Print #f, Phrases(c)
            Else
                Y = "x"
            End If
        Next c
        Close #f
        
        ReDim Phrases(0)
        Call frmChat.LoadArray(LOAD_PHRASES, Phrases())
        
        If Len(Y) > 0 Then
            strSend = "Phrase " & Chr(34) & u & Chr(34) & " deleted."
        Else
            strSend = "That phrase is not banned."
        End If
        
        b = True
        GoTo Display
        
'            ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "monitor ") Then
'                u = Right(Message, Len(Message) - 9)
'
'                If Not MonitorExists Then
'                    InitMonitor
'                End If
'
'                MonitorForm.txtAdd.text = u
'                Call MonitorForm.cmdAdd_Click
'
'                strSend = "Added " & Chr(34) & u & Chr(34) & " to the monitor."
'
'                b = True
'                GoTo Display
                
'            ElseIf Left$(cMsg, 8) = BotVars.Trigger & "notify " Then
'
'                If Not MonitorExists Then
'                    InitMonitor
'                End If
'
'                i = MonitorForm.SetStatusWatch(1, Right(cMsg, Len(cMsg) - 8))
'                If i = 0 Then
'                    strSend = "That user is not in the monitor."
'                ElseIf i = 2 Then
'                    strSend = "The monitor must be initialized before you can alter notification settings."
'                Else
'                    strSend = "Sign-on notifications for " & Right(Message, Len(Message) - 8) & " enabled."
'                End If
'
'                b = True
'                GoTo Display
'
'            ElseIf Left$(cMsg, 10) = BotVars.Trigger & "denotify " Then
'
'                If Not MonitorExists Then
'                    InitMonitor
'                End If
'
'                i = MonitorForm.SetStatusWatch(0, Mid$(cMsg, 11))
'
'                If i = 0 Then
'                    strSend = "That user is not in the monitor."
'                ElseIf i = 2 Then
'                    strSend = "The monitor must be initialized before you can alter notification settings."
'                Else
'                    strSend = "Sign-on notifications for " & Right(Message, Len(Message) - 11) & " disabled."
'                End If
'
'                b = True
'                GoTo Display
'
'            ElseIf Left$(cMsg, 11) = (BotVars.Trigger & "unmonitor ") Then
'                On Error Resume Next
'
'                If Not MonitorExists Then
'                    InitMonitor
'                End If
'
'                u = Right(Message, Len(Message) - 11)
'                MonitorForm.RemUser MonitorForm.lvMonitor.FindItem(u).Index
'
'                strSend = "Removed " & Chr(34) & u & Chr(34) & " from the monitor."
'                b = True
'                GoTo Display
'
'            ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "check ") Then
'                On Error GoTo Erroneouz
'                Dim X As ListItem
'                u = Right(Message, Len(Message) - 7)
'
'                If Not MonitorExists Then
'                    InitMonitor
'                End If
'
'                Set X = MonitorForm.lvMonitor.FindItem(u)
'                If X.Index = 0 Then
'                    strSend = "That user is not being monitored."
'                    b = True
'                    GoTo Display
'                End If
'
'                strSend = "User " & Chr(34) & u & Chr(34) & " is "
'
'                Select Case X.ListSubItems(1).text
'                    Case "Offline"
'                        strSend = strSend & "offline."
'                    Case "Online"
'                        strSend = strSend & "online."
'                    Case Else
'                        strSend = "Status not set."
'                End Select
'
'                Set X = Nothing
'
'                b = True
'                GoTo Display
'Erroneouz:
'                strSend = "That user is not being monitored."
'                b = True
'                GoTo Display
        
        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "tagban ") Then
        u = Right(Message, Len(Message) - 8)
        Call Commands(dbAccess, Username, BotVars.Trigger & "addtag " & u, InBot, CC, WhisperedIn, PublicOutput)
        GoTo theEnd

    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "addtag ") Then
        u = Right(Message, Len(Message) - 8)
        Call Commands(dbAccess, Username, BotVars.Trigger & "tagadd " & u, InBot, CC, WhisperedIn, PublicOutput)
        GoTo theEnd
    
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "tagadd ") Then
        On Error GoTo sendit3
        u = Right(Message, (Len(Message) - 8))
        
        If Len(GetTagbans(u)) > 1 Then
            strSend = "That tag is covered by an existing tagban."
            b = True
            GoTo Display
        End If
        
        If InStr(1, u, "*", vbTextCompare) = 0 And Len(u) > 4 Then
            Call Commands(dbAccess, Username, BotVars.Trigger & "shitadd " & u, InBot, CC, WhisperedIn, PublicOutput)
            GoTo theEnd
        End If
        
        If Dir$(GetFilePath("tagbans.txt")) = vbNullString Then
            Open (GetFilePath("tagbans.txt")) For Output As #f
            Close #f
        End If
        
        Open (GetFilePath("tagbans.txt")) For Append As #f
        Print #f, u & vbCrLf
        Close #f
        
        If InStr(u, " ") > 0 Then
            WildCardBan u, Mid$(u, InStr(u, " ")), 1
        Else
            WildCardBan u, "Tagban: " & u, 1
        End If
        
        b = True
        strSend = "Added tag " & Chr(34) & u & Chr(34) & " to the tagban list."
        GoTo Display:
sendit3:
        strSend = "What tag would you like to add?"
        b = True
        GoTo Display
                
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "fadd ") Then
        On Error GoTo sendb
        u = Right(Message, (Len(Message) - 6))
        AddQ "/f a " & u, 1
        strSend = "Added user " & Chr(34) & u & Chr(34) & " to this account's friends list."
        b = True
        GoTo Display:
sendb:
        b = True
        strSend = "Who do you want to add?"
        GoTo Display
        
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "frem ") Then
        On Error GoTo senda
        u = Right(Message, (Len(Message) - 6))
        AddQ "/f r " & u, 1
        strSend = "Removed user " & Chr(34) & u & Chr(34) & " from this account's friends list."
        b = True
        GoTo Display:
senda:
        strSend = "Who do you want to remove?"
        b = True
        GoTo Display
        
    ' new code written 1/17/06
'            ElseIf Left$(cMsg, 11) = BotVars.Trigger & "safecheck " Then
'                If GetSafelist(Mid$(cMsg, 12)) Then
'                    strSend = "That user is safelisted."
'                Else
'                    strSend = "That user is not safelisted."
'                End If
'                b = True
'                GoTo Display
        
    ElseIf Left$(cMsg, 10) = (BotVars.Trigger & "safelist ") Then
        u = Mid$(Message, 11)
        Call Commands(dbAccess, Username, BotVars.Trigger & "safeadd " & u, InBot, CC, WhisperedIn, PublicOutput)
        GoTo theEnd
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "safeadd ") Then
        u = GetStringChunk(cMsg, 2)
        
        strSend = AddToSafelist(u, Username)
        
        If LenB(strSend) = 0 Then
            strSend = "Added tag/user " & Chr(34) & u & Chr(34) & " to the safelist."
        End If
        
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "pban ") Then
        u = Right(Message, Len(Message) - 6)
        Call Commands(dbAccess, Username, BotVars.Trigger & "shitadd " & u, InBot, CC, WhisperedIn, PublicOutput)
        GoTo theEnd
        
    ElseIf Left$(cMsg, 4) = (BotVars.Trigger & "sl ") Then
        u = Mid$(Message, 5)
        Call Commands(dbAccess, Username, BotVars.Trigger & "shitadd " & u, InBot, CC, WhisperedIn, PublicOutput)
        GoTo theEnd
    
    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "exile ") Then
        u = Mid$(Message, 8)
        If InStr(1, u, " ") > 0 Then
            Y = Split(u, " ")(0)
        End If
        
        Call Commands(dbAccess, Username, BotVars.Trigger & "shitadd " & u, InBot, CC, WhisperedIn, PublicOutput)
        Call Commands(dbAccess, Username, BotVars.Trigger & "ipban " & IIf(LenB(Y), Y, u), InBot, CC, WhisperedIn, PublicOutput)
        
        GoTo theEnd
        
    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "unexile ") Then
        u = Mid$(Message, 10)
        If InStr(1, u, " ") > 0 Then
            u = Split(u, " ")(0)
        End If
        
        Call Commands(dbAccess, Username, BotVars.Trigger & "shitdel " & u, InBot, CC, WhisperedIn, PublicOutput)
        Call Commands(dbAccess, Username, BotVars.Trigger & "unignore " & u, InBot, CC, WhisperedIn, PublicOutput)
        
        GoTo theEnd
        
    ElseIf cMsg = (BotVars.Trigger & "shitlist") Then
        Y = GetFilePath("autobans.txt")
        
        If LenB(Dir$(Y)) = 0 Then
            strSend = "No shitlist found."
            b = True
            GoTo Display
        End If
        
        Open (Y) For Input As #f
        
        If LOF(f) < 2 Then
            strSend = "There are no shitlisted users."
            b = True
            GoTo Display
        End If
        
        Do
            i = i + 1
            Line Input #f, Response
            
            ReDim Preserve strArray(0 To i)
            
            If Response <> vbNullString And Len(Response) >= 2 Then
                If InStr(Response, " ") Then
                    strArray(i) = Mid$(Response, 1, InStr(Response, " ") - 1)
                Else
                    strArray(i) = Response
                End If
            Else
                i = i - 1
            End If
        Loop While Not EOF(f)
        
        strSend = "Tags/users found: "
        
        For i = (LBound(strArray) + 1) To UBound(strArray)
            strSend = strSend & strArray(i)
            
            If i <> UBound(strArray) Then strSend = strSend & ", "
            
            If Len(strSend) > 70 Then
                If i <> UBound(strArray) Then strSend = strSend & " [more]"
                
                If WhisperCmds And Not InBot Then
                    If Dii Then AddQ "/w *" & Username & Space(1) & strSend Else AddQ "/w " & Username & Space(1) & strSend
                ElseIf InBot = True And Not PublicOutput Then
                    frmChat.AddChat RTBColors.ConsoleText, strSend
                Else
                    AddQ strSend
                End If
                
                strSend = "Tags/users found: "
            End If
        Next i
        
        b = True
        GoTo Display
        
        
                
    ElseIf cMsg = (BotVars.Trigger & "safelist") Then
        If colSafelist.Count = 0 Then
            strSend = "There are no safelisted users or tags."
        Else
            strSend = "Tags/users found: "
            b = False
            
            For i = 1 To colSafelist.Count
                Debug.Print colSafelist.Item(i).Name
                
                strSend = strSend & ReversePrepareCheck(colSafelist.Item(i).Name)
                
                If i < colSafelist.Count Then
                    strSend = strSend & ", "
                End If
                
                b = True
                
                If Len(strSend) > 70 Then
                    If i < colSafelist.Count Then strSend = strSend & " [more]"
                    
                    If WhisperCmds And Not InBot Then
                        If Dii Then AddQ "/w *" & Username & Space(1) & strSend Else AddQ "/w " & Username & Space(1) & strSend
                    ElseIf InBot = True And Not PublicOutput Then
                        frmChat.AddChat RTBColors.ConsoleText, strSend
                    Else
                        AddQ strSend
                    End If
                    
                    b = False
                    strSend = "Tags/users found: "
                End If
            Next i
            
            If Not b Then
                GoTo theEnd
            End If
        
        End If
        
        b = True
        GoTo Display

                
    ElseIf cMsg = (BotVars.Trigger & "tagbans") Then
        If Dir$(GetFilePath("tagbans.txt")) = vbNullString Then
            strSend = "No tagbans list found."
            b = True
            GoTo Display
        End If
        
        Open (GetFilePath("tagbans.txt")) For Input As #f
        If LOF(f) < 2 Then
            strSend = "No users are tagbanned."
            b = True
            GoTo Display
        End If
        Do
            i = i + 1
            Input #f, Response
            ReDim Preserve strArray(0 To i)
            If Response <> vbNullString And Len(Response) >= 2 Then
                strArray(i) = Response
            Else
                i = i - 1
            End If
        Loop Until EOF(f)
        strSend = "Tagbans found: "
        For i = (LBound(strArray) + 1) To UBound(strArray)
            strSend = strSend & strArray(i) & ", "
            If Len(strSend) > 80 Then
                strSend = Left$(strSend, Len(strSend) - 2)
                strSend = strSend & " [more]"
                If WhisperCmds And Not InBot Then
                    If Dii Then AddQ "/w *" & Username & Space(1) & strSend Else AddQ "/w " & Username & Space(1) & strSend
                ElseIf InBot = True And Not PublicOutput Then
                    frmChat.AddChat RTBColors.ConsoleText, strSend
                Else
                    AddQ strSend
                End If
                strSend = "Tagbans found: "
            End If
        Next i
        strSend = Left$(strSend, Len(strSend) - 2)
        b = True
        GoTo Display

    ElseIf Left$(cMsg, 9) = (BotVars.Trigger & "shitadd ") Then
            
        On Error GoTo sendit4
            u = Right(Message, (Len(Message) - 10))
            
            If LenB(GetShitlist(u)) > 0 Then
                strSend = "That user is already shitlisted."
                b = True
                GoTo Display
            End If
            
            If InStr(1, Split(cMsg, " ")(1), "*", vbTextCompare) > 0 Then
                Call Commands(dbAccess, Username, BotVars.Trigger & "addtag " & u, InBot, CC, WhisperedIn, PublicOutput)
                GoTo theEnd
            End If
            
            gAcc = GetAccess(Split(cMsg, " ")(1))
            
            If dbAccess.Access <= gAcc.Access Then
                strSend = "You do not have access to do that."
            End If
            
            If dbAccess.Access < 100 And InStr(gAcc.Access, "A") > 0 Then
                strSend = "You do not have access to do that."
            End If
            
            If LenB(strSend) > 0 Then
                b = True
                GoTo Display
            End If
            
            Y = GetFilePath("autobans.txt")
            
            If Dir$(Y) = vbNullString Then
             Open (Y) For Output As #f
             Close #f
            End If
            
            Open (Y) For Append As #f
            Print #f, u
            Close #f
            
            If InStr(1, u, " ", vbTextCompare) = 0 Then
                If MyFlags = 2 Or MyFlags = 18 Then AddQ "/ban " & IIf(Dii, "*", "") & u & " Shitlisted.", 1
                strSend = "Added " & u & " to the shitlist."
            Else
                If MyFlags = 2 Or MyFlags = 18 Then AddQ "/ban " & IIf(Dii, "*", "") & u, 1
                strSend = "Added " & Left$(u, InStr(1, u, " ", vbTextCompare) - 1) & " to the shitlist."
            End If
            
            b = True
            GoTo Display:
sendit4:
            strSend = "Who would you like to add to the shitlist?"
            b = True
            GoTo Display
            
    ElseIf Left$(cMsg, 5) = (BotVars.Trigger & "dnd ") Then
            Dim DNDMsg As String
            If Len(Message) <= 5 Then
                AddQ "/dnd", 1
                GoTo theEnd:
            End If
            DNDMsg = Right(Message, (Len(Message) - 5))
            AddQ "/dnd " & DNDMsg, 1
            GoTo theEnd
                
    ElseIf Left$(cMsg, 4) = BotVars.Trigger & "dnd" Then
        AddQ "/dnd", 1
        GoTo theEnd

    ElseIf cMsg = (BotVars.Trigger & "bancount") Then
            If BanCount = 0 Then
                strSend = "No users have been banned since I joined this channel."
            Else
                strSend = "Since I joined this channel, " & BanCount & " user(s) have been banned."
            End If
            b = True
            GoTo Display
                
    ElseIf Left$(cMsg, 10) = BotVars.Trigger & "tagcheck " Then
    
            u = Right(cMsg, Len(cMsg) - 10)
            Y = GetTagbans(u)
            
            If Len(Y) < 2 Then
                strSend = "That user matches no tagbans."
            Else
                strSend = "That user matches the following tagban(s): " & Y
            End If
            b = True
            GoTo Display
            
    'shitcheck / slcheck added 2007-06-10
    ElseIf Left$(cMsg, 9) = BotVars.Trigger & "slcheck " Or Left$(cMsg, 11) = BotVars.Trigger & "shitcheck " Then
    
            Y = GetStringChunk(cMsg, 2)
            
            If LenB(Y) > 0 Then
                strSend = "That user "
                
                gAcc = GetAccess(Y)
                
                If InStr(gAcc.Flags, "B") > 0 Then
                    strSend = strSend & "has 'B' in their flags"
                    Track = 1
                End If
                
                If LenB(GetShitlist(Y)) Then
                    If Track = 1 Then
                        strSend = strSend & " and "
                    End If
                    
                    strSend = strSend & "is on the bot's shitlist"
                    Track = 2
                End If
                
                If Track > 0 Then
                    strSend = strSend & "."
                Else
                    strSend = "That user is not shitlisted and does not have 'B' in their flags."
                End If
            Else
                strSend = "Please specify a username to check."
            End If
            
            b = True
            GoTo Display

            
    ElseIf Left$(cMsg, 11) = BotVars.Trigger & "safecheck " Then
    
            u = Right(cMsg, Len(cMsg) - 11)
            Y = GetSafelistMatches(u)
            'Debug.Print Y
            
            If Len(Y) < 2 Then
                strSend = "That user matches no safelist entries."
            Else
                strSend = "That user matches the following safelist entr"
                
                strArray() = Split(Y, " ")
                
                If UBound(strArray) > 0 Then
                    strSend = strSend & "ies: "
                
                    i = 0
                    Track = 0
                    While i <= UBound(strArray)
                        While Track < 10 And Track <= UBound(strArray)
                            If Len(strArray(i)) > 0 Then
                                strSend = strSend & strArray(i)
                                
                                If Track < 9 Then
                                    If i <> UBound(strArray) Then
                                        strSend = strSend & ", "
                                    End If
                                Else
                                    strSend = strSend & " [more]"
                                End If
                            End If
                            
                            Track = Track + 1
                            i = i + 1
                        Wend
                        
                        Track = 0
                    Wend
                Else
                    strSend = strSend & "y: " & Y
                End If
            End If
            
            b = True
            GoTo Display
            
    ElseIf Left$(cMsg, 10) = BotVars.Trigger & "readfile " Then
            u = Trim$(Right(Message, Len(Message) - 10))
            
            While (Mid$(u, 1, 1) = ".") And (Len(u) > 1)
                u = Mid$(u, 2)
            Wend
        
            If InStr(u, ".") > 0 Then
                Y = Left$(u, InStr(u, ".") - 1)
            Else
                Y = u
            End If
            
            ' Added 2/06 to fix an exploit brought by Fiend(KIP)
            If InStr(u, "..") > 0 Then
                u = Replace(u, "..", "")
            End If
            
            'Debug.Print Y
            
            Select Case UCase(Y)
                Case "CON", "PRN", "AUX", "CLOCK$", "NUL", "COM1", "COM2", "COM3", _
                        "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", _
                            "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
                    
                    strSend = "You cannot read that file."
                    b = True
                    GoTo Display
            End Select
            
            If InStr(1, u, ".ini", vbTextCompare) > 0 Then
                strSend = ".INI files cannot be read."
                b = True
                GoTo Display
            End If
            
            u = App.Path & "\" & u
            
            On Error GoTo readError
            
            If Dir$(u) = vbNullString Then
                strSend = "File does not exist."
                b = True
                GoTo Display
            End If
            
            i = FreeFile
            Open u For Input As #i
            
            If LOF(i) < 2 Then
                strSend = "File is empty."
                Close #i
                b = True
                GoTo Display
            End If
            
            Do While Not EOF(i)
                Line Input #i, Y
                
                If Len(Y) > 200 Then Y = Left$(Y, 200)
                
                If Len(Y) > 2 Then
                    FilteredSend Username, Y, WhisperCmds, InBot, PublicOutput
                End If
            Loop
            
            Close #i
            
            GoTo theEnd
            
readError:
            strSend = "Error reading file."
            b = True
            GoTo Display

            
    ElseIf cMsg = (BotVars.Trigger & "levelban") Or cMsg = BotVars.Trigger & "levelbans" Then
            If BotVars.BanUnderLevel = 0 Then
                strSend = "Currently not banning Warcraft III users by level."
            Else
                strSend = "Currently banning Warcraft III users under level " & BotVars.BanUnderLevel & "."
            End If
            b = True
            GoTo Display
            
    ElseIf cMsg = (BotVars.Trigger & "d2levelban") Or cMsg = BotVars.Trigger & "d2levelbans" Then
            If BotVars.BanD2UnderLevel = 0 Then
                strSend = "Currently not banning Diablo II users by level."
            Else
                strSend = "Currently banning Diablo II users under level " & BotVars.BanD2UnderLevel & "."
            End If
            b = True
            GoTo Display

    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "greet ") Then
    
        strArray() = Split(cMsg, " ", 3)
        
        If UBound(strArray) > 0 Then
            Select Case LCase(strArray(1))
            
                Case "on"
                    BotVars.UseGreet = True
                    strSend = "Greet messages enabled."
                    WriteINI "Other", "UseGreets", "Y"
                
                Case "off"
                    BotVars.UseGreet = False
                    strSend = "Greet messages disabled."
                    WriteINI "Other", "UseGreets", "N"
                    
                Case "whisper"
                    If UBound(strArray) > 1 Then
                        Select Case LCase(strArray(2))
                            Case "on"
                                BotVars.WhisperGreet = True
                                strSend = "Greet messages will now be whispered."
                                WriteINI "Other", "WhisperGreet", "Y"
                                
                            Case "off"
                                BotVars.WhisperGreet = False
                                strSend = "Greet messages will no longer be whispered."
                                WriteINI "Other", "WhisperGreet", "N"
                                
                        End Select
                    End If
    
                Case Else
                    If InStr(1, cMsg, "/squelch", vbTextCompare) > 0 Or _
                        InStr(1, cMsg, "/ban ", vbTextCompare) > 0 Or _
                            InStr(1, cMsg, "/ignore", vbTextCompare) > 0 Or _
                                InStr(1, cMsg, "/des", vbTextCompare) > 0 Or _
                                    InStr(1, cMsg, "/re", vbTextCompare) > 0 Then
                                    
                        strSend = "One or more invalid terms are present. Greet message not set."
                    Else
                        strSend = "Greet message set."
                        BotVars.GreetMsg = Right(Message, Len(Message) - 7)
                        WriteINI "Other", "GreetMsg", Right(Message, Len(Message) - 7)
                    End If
            
            End Select
        End If
    
        If LenB(strSend) > 0 Then
            b = True
            GoTo Display
        Else
            GoTo theEnd
        End If
            
    ElseIf cMsg = (BotVars.Trigger & "allseen") Then
        strSend = "Last 15 users seen: "
        
        If colLastSeen.Count = 0 Then
            strSend = strSend & "(list is empty)"
        Else
            For i = 1 To colLastSeen.Count
                strSend = strSend & colLastSeen.Item(i)
                
                If Len(strSend) > 90 Then
                    strSend = strSend & " [more]"
                    FilteredSend Username, strSend, WhisperCmds, InBot, PublicOutput
                    strSend = ""
                ElseIf i < colLastSeen.Count Then
                    strSend = strSend & ", "
                End If
                
            Next i
        End If
        
        b = True
        GoTo Display
        
        
            
    ElseIf Left$(cMsg, 5) = (BotVars.Trigger & "ban ") Then
        If MyFlags <> 2 And MyFlags <> 18 Then
            If InBot Then
                strSend = "You are not a channel operator."
                b = True
                GoTo Display
            End If
            
            GoTo theEnd
        End If
        
        On Error GoTo sendit6:
        
        u = Right(Message, Len(Message) - 5)
        i = InStr(1, u, " ")
        If i > 0 Then
            banMsg = Mid$(u, i + 1)
            u = Left$(u, i - 1)
        End If
        
        If InStr(1, u, "*", vbTextCompare) > 0 Then
            WildCardBan u, banMsg, 1
            GoTo theEnd
        Else
            If banMsg <> vbNullString Then
                Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), dbAccess.Access)
            Else
                Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), dbAccess.Access)
            End If
        End If
        
        If Len(Y) > 2 Then
            strSend = Y
            b = True
            GoTo Display
        Else
            GoTo theEnd
        End If
sendit6:
        b = True
        strSend = "Who do you want to ban?"
        GoTo Display
                      
    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "unban ") Then
                On Error GoTo sende
                u = Right(Message, (Len(Message) - 7))
                
                If bFlood Then
                    If floodCap < 45 Then
                        floodCap = floodCap + 15
                        bnetSend "/unban " & u
                        GoTo theEnd
                    End If
                End If
                
                If InStr(1, cMsg, "*", vbTextCompare) <> 0 Then
                    WildCardBan u, vbNullString, 2
                    GoTo theEnd
                End If
                
                If Dii = True Then
                    If Not (Mid$(u, 1, 1) = "*") Then u = "*" & u
                End If
                
                strSend = "/unban " & u
                If InBot Then
                    AddQ strSend, 1
                    GoTo theEnd
                End If
                b = True
                GoTo Display
sende:
                b = True
                strSend = "Unban who?"
                GoTo Display
                    
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "kick ") Then
        If MyFlags <> 2 And MyFlags <> 18 Then
            If InBot Then
                strSend = "You are not a channel operator."
                b = True
                GoTo Display
            End If
            
            GoTo theEnd
        End If
    
        On Error GoTo sendit
        
        u = Right(Message, Len(Message) - 6)
        i = InStr(1, u, " ", vbTextCompare)
        
        
        If i > 0 Then
            banMsg = Mid$(u, i + 1)
            u = Left$(u, i - 1)
        End If
        
        If InStr(1, u, "*", vbTextCompare) > 0 Then
'                            If dbAccess.Access > 99 Then
'                                WildCardBan u, banMsg, 0, 1
'                            Else
                WildCardBan u, banMsg, 0
'                            End If
            GoTo theEnd
        End If
        
        Y = Ban(u & IIf(Len(banMsg) > 0, " " & banMsg, vbNullString), dbAccess.Access, 1)
        
        If Len(Y) > 1 Then
            strSend = Y
            b = True
            GoTo Display
        Else
            GoTo theEnd
        End If
sendit:
        strSend = "Kick what user?"
        b = True
        GoTo Display
                    
    ElseIf cMsg = (BotVars.Trigger & "lastwhisper") Or cMsg = (BotVars.Trigger & "lw") Then
        If LastWhisper <> vbNullString Then
            strSend = "The last whisper to this bot was from: " & LastWhisper
            b = True
            GoTo Display
        Else
            strSend = "The bot has not been whispered since it logged on."
            b = True
            GoTo Display
        End If

    ElseIf Left$(cMsg, 5) = (BotVars.Trigger & "say ") Then
        On Error GoTo sendy
        
        u = Mid$(Message, 6)
        
        If dbAccess.Access >= GetAccessINIValue("say70", 70) Then
            If dbAccess.Access >= GetAccessINIValue("say90", 90) Then
                strSend = u
            Else
                strSend = Replace(u, "/", "")
            End If
        Else
            strSend = Username & " says: " & u
        End If
        
        If InBot Or WhisperedIn Then
            AddQ strSend
            GoTo theEnd
            
        Else
            b = True
            GoTo Display
            
        End If
sendy:
        strSend = "Say what?"
        b = True
        GoTo Display
        
        
'            ElseIf cMsg = BotVars.Trigger & "online" Then
'
'                If Not MonitorExists Then InitMonitor
'
'                If MonitorForm.lvMonitor.ListItems.Count = 0 Then
'                    strSend = "No users are monitored."
'                    b = True
'                    GoTo Display
'                End If
'
'                strSend = "User(s) online:"
'
'                With MonitorForm.lvMonitor
'                    For i = 1 To .ListItems.Count
'                        If .ListItems(i).ListSubItems(1).text = "Online" Then
'
'                            strSend = strSend & ", " & .ListItems(i).text
'                            c = c + 1
'                            If Len(strSend) > 75 Then
'                                strSend = strSend & " [more]"
'                                strSend = Replace(strSend, ":,", ":")
'
'                                if BotVars.whispercmds And Not InBot Then
'                                    AddQ "/w " & IIf(Dii, "*", "") & Username & Space(1) & strSend
'                                ElseIf InBot Then
'                                    frmChat.AddChat RTBColors.ConsoleText, strSend
'                                Else
'                                    AddQ strSend
'                                End If
'
'                                strSend = "User(s) online:"
'                            End If
'
'                        End If
'                    Next i
'                End With
'
'                If StrComp(strSend, "User(s) online:", vbTextCompare) = 0 And c = 0 Then strSend = "No monitored users are online."
'
'                strSend = Replace(strSend, ":,", ":")
'                strSend = strSend & " [" & c & " of " & i & "]"
'
'                b = True
'                GoTo Display
                
    ElseIf Left$(cMsg, 8) = BotVars.Trigger & "expand " Then
        u = Right(Message, Len(Message) - 8)
        strSend = Expand(u)
        
        If Len(strSend) > 220 Then strSend = Mid$(strSend, 1, 220)
        
        If InBot And Not WhisperedIn Then
            AddQ strSend
            GoTo theEnd
            
        ElseIf WhisperedIn Then
            AddQ strSend
            GoTo theEnd
            
        Else
            b = True
            GoTo Display
        End If

    ElseIf Left$(cMsg, 8) = BotVars.Trigger & "detail " Or Left$(cMsg, 5) = BotVars.Trigger & "dbd " Then
        strSend = GetDBDetail(Split(cMsg, " ")(1))
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 6) = BotVars.Trigger & "info " Then
        u = Right(Message, Len(Message) - 6)
        i = UsernameToIndex(u)
        
        If i > 0 Then
            With colUsersInChannel.Item(i)
                strSend = "User " & .Username & " is logged on using " & ProductCodeToFullName(.Product)
                
                If .Flags And &H2 = &H2 Then
                    strSend = strSend & " with ops, and a ping time of " & .Ping & "ms."
                Else
                    strSend = strSend & " with a ping time of " & .Ping & "ms."
                End If
                
                If WhisperCmds And Not InBot Then
                    AddQ "/w " & IIf(Dii, "*", "") & Username & Space(1) & strSend
                ElseIf InBot And Not PublicOutput Then
                    frmChat.AddChat RTBColors.ConsoleText, strSend
                Else
                    AddQ strSend
                End If
                
                strSend = "He/she has been present in the channel for " & ConvertTime(.TimeInChannel(), 1) & "."
                
            End With
        Else
            strSend = "No such user is present."
        End If
        
        b = True
        GoTo Display
                
    ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "shout ") Then
        On Error GoTo sendy
        u = UCase(Right(Message, (Len(Message) - 7)))

        If dbAccess.Access > 69 Then
            If dbAccess.Access > 89 Then
                strSend = u
            Else
                strSend = Replace(u, "/", vbNullString)
            End If
        Else
            strSend = Username & " shouts: " & u
        End If
        
        AddQ strSend
        GoTo theEnd
sendf2:
        strSend = "Shout what?"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 9) = BotVars.Trigger & "voteban " Then
        If VoteDuration < 0 Then
            Call Voting(BVT_VOTE_START, BVT_VOTE_BAN, Right(cMsg, Len(cMsg) - 9))
            VoteDuration = 30
            strSend = "30-second VoteBan vote started. Type YES to ban " & Right(cMsg, Len(cMsg) - 9) & ", NO to acquit him/her."
        Else
            strSend = "A vote is currently in progress."
        End If
        
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 10) = BotVars.Trigger & "votekick " Then
        If VoteDuration < 0 Then
            Call Voting(BVT_VOTE_START, BVT_VOTE_KICK, Right(cMsg, Len(cMsg) - 10))
            VoteDuration = 30
            VoteInitiator = dbAccess
            strSend = "30-second VoteKick vote started. Type YES to kick " & Right(cMsg, Len(cMsg) - 9) & ", NO to acquit him/her."
        Else
            strSend = "A vote is currently in progress."
        End If
        
        b = True
        GoTo Display
            
    ElseIf Left$(cMsg, 6) = BotVars.Trigger & "vote " Then
        If VoteDuration < 0 Then
            If StrictIsNumeric(Right(cMsg, Len(cMsg) - 6)) And Val(Mid$(cMsg, 7)) <= 32000 Then
                VoteDuration = Right(cMsg, Len(cMsg) - 6)
                VoteInitiator = dbAccess
                Call Voting(BVT_VOTE_START, BVT_VOTE_STD)
                strSend = "Vote initiated. Type YES or NO to vote; your vote will be counted only once."
            Else
                strSend = "Please enter a number of seconds for your vote to last."
            End If
            
            b = True
            GoTo Display
        Else
            strSend = "A vote is currently in progress."
            b = True
            GoTo Display
        End If
            
    ElseIf cMsg = BotVars.Trigger & "tally" Then
        If VoteDuration > 0 Then
            strSend = Voting(BVT_VOTE_TALLY)
        Else
            strSend = "No vote is currently in progress."
        End If
        b = True
        GoTo Display
            
    ElseIf cMsg = BotVars.Trigger & "cancel" Then
        If VoteDuration > 0 Then
            strSend = Voting(BVT_VOTE_END, BVT_VOTE_CANCEL)
        Else
            strSend = "No vote in progress."
        End If
        
        b = True
        GoTo Display
            
    ElseIf cMsg = (BotVars.Trigger & "back") Then
        If AwayMsg <> vbNullString Then
            AddQ "/away", 1
            
            If Not InBot Then
                strSend = "/me is back from " & AwayMsg & "."
                AwayMsg = vbNullString
                b = True
                GoTo theEnd
            End If
        Else
            If Not BotVars.DisableMP3Commands Then
                If iTunesReady Then
                    iTunesBack
                    strSend = "Skipped backwards."
                    b = True
                    GoTo Display
                Else
                    On Error GoTo sendf3
                    hWndWA = GetWinamphWnd()
                    If hWndWA = 0 Then
                       strSend = "Winamp is not loaded."
                       b = True
                       GoTo Display
                    End If
                    SendMessage hWndWA, WM_COMMAND, WA_PREVTRACK, 0
                    strSend = "Skipped backwards."
                    b = True
                    GoTo Display
sendf3:
                    strSend = "Error."
                    b = True
                    GoTo Display
                End If
            End If
        End If

    ElseIf cMsg = (BotVars.Trigger & "uptime") Then
        strSend = "System uptime " & ConvertTime(GetUptimeMS) & ", connection uptime " & ConvertTime(uTicks) & "."
        b = True
        GoTo Display

    ElseIf cMsg = (BotVars.Trigger & "away") Then
        If LenB(AwayMsg) > 0 Then
            AddQ "/away", 1
            If Not InBot Then AddQ "/me is back from (" & AwayMsg & ")"
            AwayMsg = ""
        Else
            AddQ "/away", 1
            If Not InBot Then AddQ "/me is away (" & AwayMsg & ")"
            AwayMsg = "-"
        End If
        
        GoTo theEnd
        
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "away ") Then
        On Error GoTo sendcc
        
        If Len(Message) <= 5 Then
            AddQ "/away", 1
            If Not InBot Then AddQ "/me is away."
            AwayMsg = "-"
        Else
            AwayMsg = Right(Message, (Len(Message) - 6))
            AddQ "/away " & AwayMsg, 1
            If Not InBot Then AddQ "/me is away (" & AwayMsg & ")"
        End If
        
        GoTo theEnd
        
sendcc:
        strSend = "/away"
        b = True
        GoTo Display
                
                
    ElseIf cMsg = (BotVars.Trigger & "mp3") Then
    On Error GoTo sendp
        WindowTitle = GetCurrentSongTitle(True)

        If WindowTitle = vbNullString Then
            strSend = "Winamp is not loaded."
            b = True
            GoTo Display
        End If
        strSend = "Current MP3: " & WindowTitle
        b = True
        GoTo Display
sendp:
        strSend = "Error. Winamp may be stopped or not loaded."
        b = True
        GoTo Display
    
    ElseIf Left$(cMsg, 8) = BotVars.Trigger & "deldef " Then
        On Error GoTo error_deldef
        u = Mid$(cMsg, 9)
        
        If Len(u) > 0 Then
            WriteINI "Def", u, "%deleted%", "definitions.ini"
            strSend = "That definition has been erased."
        Else
error_deldef:
            strSend = "There was an error removing that definition."
        End If
        
        b = True
        GoTo Display

    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "define ") Or Left$(cMsg, 5) = (BotVars.Trigger & "def ") Then
        On Error GoTo sendyz
        
        If Dir$(GetFilePath("definitions.ini")) = vbNullString Then
            strSend = "No definition list found. Please use '.newdef term|definition' to make one."
            b = True
            GoTo Display
        End If
        
        If Left$(cMsg, 8) = (BotVars.Trigger & "define ") Then
            Track = 8
        Else
            Track = 5
        End If
        
        u = LCase(Trim(Mid$(Message, Track)))
        
        Response = ReadINI("Def", u, "definitions.ini")
        
        If Response = vbNullString Or StrComp(Response, "%deleted%") = 0 Then
            strSend = "No definition on file for " & u & "."
        Else
            strSend = "[" & u & "]: " & Response
        End If
        
        b = True
        GoTo Display
        
sendyz:
        strSend = "Define what?"
        b = True
        GoTo Display
        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "newdef ") Then
        On Error GoTo sendi2
        
        u = Right(Message, Len(Message) - 8)
        
        Track = InStr(1, u, "|", vbTextCompare)
        
        z = Right(u, Len(u) - Track)
        u = Left$(u, Len(u) - Len(z) - 1)
        
        If z = "" Then
            strSend = "You need to specify a definition."
            b = True
            GoTo Display
        End If
        
        WriteINI "Def", u, z, "definitions.ini"
        
        strSend = "Added a definition for """ & u & """."
        b = True
        GoTo Display
        
sendi2:
        strSend = "Error: Please format your definitions correctly. (.newdef term|definition)"
        b = True
        GoTo Display
                        
    ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "ping ") Then
        On Error GoTo sendc
        u = Right(Message, (Len(Message) - 6))
        
        Dim UserPing As Long
        UserPing = GetPing(u)
        
        If UserPing < -1 Then
            strSend = "I can't see " & u & " in the channel."
        Else
            strSend = u & "'s ping at login was " & UserPing & "ms."
        End If
        
        b = True
        GoTo Display
sendc:
        b = True
        strSend = "Ping who?"
        GoTo Display
        
    ElseIf Left$(cMsg, 10) = (BotVars.Trigger & "addquote ") Then
        On Error GoTo sendit0
        
        u = Right(Message, (Len(Message) - 10))
        
        Y = Dir$(GetFilePath("quotes.txt"))
        
        If LenB(Y) = 0 Then
            Open (GetFilePath("quotes.txt")) For Output As #f
            Close #f
        End If
        
        Open (GetFilePath("quotes.txt")) For Append As #f
            Print #f, u
        Close #f
        
        strSend = "Quote added!"
        b = True
        GoTo Display
sendit0:
        b = True
        strSend = "I need a quote to add."
        GoTo Display
        
    ElseIf cMsg = BotVars.Trigger & "owner" Then
        If LenB(BotVars.BotOwner) > 0 Then
            strSend = "This bot's owner is " & BotVars.BotOwner & "."
        Else
            strSend = "No owner is set."
        End If
        
        b = True
        GoTo Display
                        
    ElseIf Left$(cMsg, 8) = (BotVars.Trigger & "ignore ") Or Left$(cMsg, 5) = BotVars.Trigger & "ign " Then
        On Error GoTo sendit9
        
        If Mid$(Message, 5, 1) = "o" Then
            u = Right(Message, Len(Message) - 8)
        Else
            u = Mid$(Message, 6)
        End If
        
        If GetAccess(u).Access >= dbAccess.Access Or InStr(GetAccess(u).Flags, "A") Then
            strSend = "That user has equal or higher access."
        Else
            AddQ "/ignore " & IIf(Dii, "*", "") & u, 1
            strSend = "Ignoring messages from " & Chr(34) & u & Chr(34) & "."
        End If
        
        b = True
        GoTo Display
sendit9:
        strSend = "Unignore who?"
        b = True
        GoTo Display
                        
        ElseIf cMsg = (BotVars.Trigger & "quote") Then
                
            strSend = GetRandomQuote
            If Len(strSend) = 0 Then strSend = "Error reading quotes, or no quote file exists."
            
            If Len(strSend) > 220 Then
                ' try one more time
                strSend = GetRandomQuote
                
                If Len(strSend) > 220 Then
                    'too long? too bad. truncate
                    strSend = Left$(strSend, 220)
                End If
            End If
            
            b = True
            GoTo Display
                    
            
        ElseIf Left$(cMsg, 10) = (BotVars.Trigger & "unignore ") Then
            On Error GoTo sendit7
            u = Right(Message, Len(Message) - 10)
            
            If Dii = True Then
                If Not (Mid$(u, 1, 1) = "*") Then u = "*" & u
            End If
            
            AddQ "/unignore " & u, 1
            strSend = "Receiving messages from """ & u & """."
            b = True
            GoTo Display
sendit7:
            strSend = "Unignore who?"
            b = True
            GoTo Display
                    
        ElseIf cMsg = BotVars.Trigger & "cq" Or cMsg = BotVars.Trigger & "scq" Then
            While colQueue.Count > 0
                colQueue.Remove 1
            Wend
            
            If InStr(1, cMsg, "s") > 0 Then
                GoTo theEnd
            Else
                strSend = "Queue cleared."
            End If
            b = True
            GoTo Display
            
        ElseIf cMsg = (BotVars.Trigger & "time") Then
            strSend = "The current time on this computer is " & Time & " on " & Format(Date, "MM-dd-yyyy") & "."
            b = True
            GoTo Display
            
        ElseIf cMsg = BotVars.Trigger & "getping" Or cMsg = BotVars.Trigger & "pingme" Then
            
            If InBot Then
                If g_Online Then
                    strSend = "Your ping at login was " & GetPing(CurrentUsername) & "ms."
                Else
                    strSend = "You are not connected."
                End If
            Else
                hWndWA = GetPing(Username)
                
                If hWndWA > -2 Then
                    strSend = "Your ping at login was " & hWndWA & "ms."
                Else
                    GoTo theEnd
                End If
            End If
            
            b = True
            GoTo Display
            
        ElseIf cMsg = BotVars.Trigger & "checkmail" Then
            Track = GetMailCount(CurrentUsername)
            
            If Track > 0 Then
                strSend = "You have " & Track & " new messages."
                
                If InBot Then
                    strSend = strSend & " Type /getmail to retrieve them."
                Else
                    strSend = strSend & " Type !inbox to retrieve them."
                End If
            Else
                strSend = "You have no mail."
            End If
            
            b = True
            GoTo Display
            
        ElseIf cMsg = BotVars.Trigger & "getmail" Then
            Dim Msg As udtMail
            
            If InBot Then Username = CurrentUsername
            
            If GetMailCount(Username) > 0 Then
                Call GetMailMessage(Username, Msg)
                
                If Len(RTrim(Msg.To)) > 0 Then
                    strSend = "Message from " & RTrim(Msg.From) & ": " & RTrim(Msg.Message)
                    b = True
                    GoTo Display
                End If
            End If
            
            GoTo theEnd
        
        ElseIf cMsg = BotVars.Trigger & "whoami" Then
        
            If InBot Then
                strSend = "You are the bot console."
                
                If (g_Online) Then
                    AddQ "/whoami"
                End If
                
            ElseIf dbAccess.Access = 1000 Then
                strSend = "You are the bot owner, " & Username & "."
                
            Else
                strSend = "You have "
                
                If dbAccess.Access > 0 Then
                    strSend = strSend & dbAccess.Access & " access"
                    
                    If dbAccess.Flags <> vbNullString Then
                        strSend = strSend & " and "
                    End If
                End If
                
                If dbAccess.Flags <> vbNullString Then
                    strSend = strSend & "flags " & dbAccess.Flags
                End If
                
                'strSend = strSend & "."
                
                If StrComp(strSend, "You have ") = 0 Then
                    strSend = "You have no access or flags, " & Username & "."
                Else
                    strSend = strSend & ", " & Username & "."
                End If
            End If
            
            b = True
            GoTo Display
                    
        'ElseIf left$(cMsg, 5) = BotVars.Trigger & "dns " Then
            'strSend = "Resolved to: " & frmChat.Access.NSLookup(Right(cMsg, Len(cMsg) - 5))
            'If InStr(1, strSend, "-1", vbTextCompare) > 0 Then
            '    strSend = "DNS lookup failed."
            'End If
            'b = True
            'GoTo Display
            
        ElseIf Left$(cMsg, 5) = (BotVars.Trigger & "add ") Or Left$(cMsg, 5) = (BotVars.Trigger & "set ") Then
        
            'On Error GoTo AddError
            Track = 0
            strArray() = Split(Message, " ")
            
            If UBound(strArray) > 1 Then
                While Left$(strArray(1), 1) = "/" And Len(strArray(1)) > 0
                    strArray(1) = Mid$(strArray(1), 2)
                Wend
                
                gAcc = GetAccess(strArray(1))
                
                If InStr(1, gAcc.Flags, "L", vbTextCompare) > 0 Then
                    If InStr(1, dbAccess.Flags, "A", vbBinaryCompare) = 0 And dbAccess.Access < 100 Then
                        strSend = "You do not have permission to modify that user's access."
                        b = True
                        GoTo Display
                    End If
                End If
                
                '/*
                ' * 0    1        2      3
                ' * .set username access flags
                ' */
                
                If IsW3 Then
                    If StrComp(Right$(strArray(1), Len(w3Realm) + 1), "@" & w3Realm, vbBinaryCompare) = 0 Then
                        
                        'Debug.Print w3Realm
                        'Debug.Print Right$(strArray(1), Len(w3Realm) + 1)
                        'Debug.Print Left$(strArray(1), InStr(1, strArray(1), "@", vbBinaryCompare) - 1)
                        strArray(1) = Left$(strArray(1), InStr(1, strArray(1), "@", vbBinaryCompare) - 1)
                        
                    End If
                End If
                
                strSend = "Set " & strArray(1) & "'s"
                
                If StrictIsNumeric(strArray(2)) Then
                    If dbAccess.Access <= gAcc.Access Then c = 1
                    If Val(strArray(2)) >= dbAccess.Access Then c = 1
                    If Val(strArray(2)) < 0 Then c = 1
                Else
                    strUser = UCase(strArray(2))
                End If
                
                If UBound(strArray) > 2 Then
                    If StrictIsNumeric(strArray(3)) Then
                        If dbAccess.Access <= gAcc.Access Then c = 1
                        If Val(strArray(3)) >= dbAccess.Access Then c = 1
                        If Val(strArray(3)) < 0 Then c = 1
                    Else
                        strUser = strUser & UCase(strArray(3))
                    End If
                End If
                
                If c <> 1 Then
                    If InStr(1, strUser, "A", vbTextCompare) > 0 Then
                        If dbAccess.Access <= 100 Then c = 1
                    End If
                    
                    If InStr(1, strUser, "B", vbTextCompare) > 0 Then
                        If dbAccess.Access < 70 Then c = 1
                        If dbAccess.Access <= GetAccess(strArray(1)).Access Then c = 1
                    End If
                    
                    If InStr(1, strUser, "Z", vbTextCompare) > 0 Then
                        If dbAccess.Access < 70 Then c = 1
                    End If
                    
                    If InStr(1, strUser, "D", vbTextCompare) > 0 Then
                        If dbAccess.Access < 100 Then c = 1
                    End If
                    
                    If InStr(1, strUser, "L", vbTextCompare) > 0 Then
                        If dbAccess.Access < 100 Then c = 1
                    End If
                    
                    If InStr(1, strUser, "S", vbTextCompare) > 0 Then
                        If dbAccess.Access < 70 Then c = 1
                    End If
                End If
                
                If c = 1 Then
                    strSend = "You do not have enough access to do that."
                    b = True
                    GoTo Display
                End If
                
                If InStr(1, cMsg, "+", vbTextCompare) = 0 And _
                    InStr(1, cMsg, "-", vbTextCompare) = 0 Then
                        gAcc.Flags = vbNullString
                End If
                
                If StrictIsNumeric(strArray(2)) And Val(strArray(2)) < 1000 Then
                    gAcc.Access = strArray(2)
                    strSend = strSend & " access to " & gAcc.Access
                    c = 1
                Else
                    strArray(2) = UCase(strArray(2))
                    
                    If Left$(strArray(2), 1) = "+" Then
                        For i = 2 To Len(strArray(2))
                            If InStr(1, gAcc.Flags, Mid(strArray(2), i, 1), vbTextCompare) = 0 Then
                                thisChar = Asc(Mid(strArray(2), i, 1))
                                
                                If thisChar >= 65 And thisChar <= 90 Then
                                
                                    gAcc.Flags = gAcc.Flags & Chr(thisChar)
                                    
                                    If thisChar = Asc("B") Then
                                        Ban strArray(1) & " AutoBan", (AutoModSafelistValue - 1)
                                    ElseIf thisChar = Asc("Z") Then
                                        WildCardBan strArray(1), "Tagbanned", 1
                                    ElseIf thisChar = Asc("S") Then
                                        AddToSafelist strArray(1), Username
                                        Track = 1
                                    End If
                                    
                                End If
                            End If
                        Next i
                    ElseIf Left$(strArray(2), 1) = "-" Then
                        For i = 2 To Len(strArray(2))
                            gAcc.Flags = Replace(gAcc.Flags, Mid(strArray(2), i, 1), vbNullString)
                            
                            thisChar = Asc(Mid$(strArray(2), i, 1))
                            
                            If thisChar = Asc("B") Then
                                For Track = LBound(gBans) To UBound(gBans)
                                    If StrComp(gBans(Track).Username, strArray(1), vbTextCompare) = 0 Then
                                        AddQ "/unban " & gBans(Track).Username
                                    End If
                                Next Track
                            ElseIf thisChar = Asc("S") Then
                                RemoveFromSafelist strArray(1)
                                Track = 2
                            End If
                        Next i
                    Else
                        For i = 1 To Len(strArray(2))
                            If InStr(1, gAcc.Flags, Mid(strArray(2), i, 1), vbTextCompare) = 0 Then
                                thisChar = Asc(Mid$(strArray(2), i, 1))
                                
                                If thisChar >= 65 And thisChar <= 90 Then
                                    
                                    If InStr(strArray(2), "S") Then
                                        AddToSafelist strArray(1), Username
                                        Track = 1
                                    End If
                                
                                    gAcc.Flags = gAcc.Flags & Mid$(strArray(2), i, 1)
                                    
                                    If thisChar = Asc("B") Then
                                        Ban strArray(1) & " AutoBan", (AutoModSafelistValue - 1)
                                    ElseIf thisChar = Asc("Z") Then
                                        WildCardBan strArray(1), "Tagbanned", 1
                                    End If
                                End If
                            End If
                        Next i
                    End If
                    
                    gAcc.Flags = Replace(gAcc.Flags, "S", "")
                    strSend = strSend & " flags to " & gAcc.Flags
                    c = 2
                End If
                
                If UBound(strArray) > 2 Then
                    If StrictIsNumeric(strArray(3)) And Val(strArray(3)) < 1000 Then
                        If c <> 1 Then
                            gAcc.Access = strArray(3)
                            If c > 0 Then
                                strSend = strSend & " and access to " & gAcc.Access
                            Else
                                strSend = strSend & " access to " & gAcc.Access
                            End If
                        End If
                    Else
                        strArray(3) = UCase(strArray(3))
                        If c <> 2 Then
                            
                            If Left$(strArray(3), 1) = "+" Then
                                For i = 2 To Len(strArray(3))
                                
                                    If InStr(1, gAcc.Flags, Mid(strArray(3), i, 1), vbTextCompare) = 0 Then
                                        thisChar = Asc(Mid$(strArray(3), i, 1))
                                        
                                        If thisChar >= 65 And thisChar <= 90 Then
                                            gAcc.Flags = gAcc.Flags & Chr(thisChar)
                                            
                                            If thisChar = Asc("B") Then
                                                Ban strArray(1) & " AutoBan", (AutoModSafelistValue - 1)
                                            ElseIf thisChar = Asc("Z") Then
                                                WildCardBan strArray(1), "Tagbanned", 1
                                            ElseIf thisChar = Asc("S") Then
                                                AddToSafelist strArray(1), Username
                                                Track = 1
                                            End If
                                        End If
                                    End If
                                Next i
                            ElseIf Left$(strArray(3), 1) = "-" Then
                                For i = 2 To Len(strArray(3))
                                    gAcc.Flags = Replace(gAcc.Flags, Mid(strArray(3), i, 1), vbNullString)
                                    
                                    thisChar = Asc(Mid(strArray(3), i, 1))
                                    
                                    If thisChar = Asc("B") Then
                                        For iWinamp = LBound(gBans) To UBound(gBans)
                                            If StrComp(LCase(gBans(iWinamp).Username), strArray(1), vbTextCompare) = 0 Then
                                                AddQ "/unban " & gBans(iWinamp).Username
                                            End If
                                        Next iWinamp
                                    ElseIf thisChar = Asc("S") Then
                                        RemoveFromSafelist strArray(1)
                                        Track = 2
                                    End If
                                    
                                Next i
                            Else
                                For i = 1 To Len(strArray(3))
                                    If InStr(1, gAcc.Flags, Mid(strArray(3), i, 1), vbTextCompare) = 0 Then
                                        thisChar = Asc(Mid(strArray(3), i, 1))
                                        
                                        If thisChar >= 65 And thisChar <= 90 Then
                                            
                                            If InStr(strArray(3), "S") Then
                                                AddToSafelist strArray(1), Username
                                                Track = 1
                                            End If
                                            
                                            gAcc.Flags = gAcc.Flags & Mid(strArray(3), i, 1)
                                            
                                            If thisChar = Asc("B") Then
                                                Ban strArray(1) & " AutoBan", (AutoModSafelistValue - 1)
                                            ElseIf thisChar = Asc("Z") Then
                                                WildCardBan strArray(1), "Tagbanned", 1
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                            
                            gAcc.Flags = Replace(gAcc.Flags, "S", "")
                            
                            If c > 0 Then
                                
                                If Len(gAcc.Flags) = 0 Then
                                    strSend = strSend & " and erased their flags"
                                Else
                                    strSend = strSend & " and flags to " & gAcc.Flags
                                End If
                            Else
                                strSend = strSend & " flags to " & gAcc.Flags
                            End If
                        End If
                    End If
                End If
                
                u = GetFilePath("users.txt")
                
                For i = LBound(DB) To UBound(DB)
                    With DB(i)
                        If StrComp(.Username, LCase(strArray(1)), vbTextCompare) = 0 Then
                            .Access = gAcc.Access
                            .Flags = gAcc.Flags
                            .ModifiedBy = Username
                            .ModifiedOn = Now
                            
                            strSend = strSend & "."
                            If InStr(1, strSend, "to .", vbTextCompare) > 0 And c = 0 Then
                                strSend = "User " & strArray(1) & "'s flags erased."
                            End If
                            
                            If BotVars.LogDBActions Then
                                LogDBAction ModEntry, Username, strArray(1), Message
                            End If
                            
                            Select Case Track
                                Case 1: strSend = strSend & " He/she is safelisted."
                                Case 2: strSend = strSend & " He/she is no longer safelisted."
                            End Select
                
                            WriteDatabase (u)
                            
                            b = True
                            GoTo Display
                        End If
                    End With
                Next i
                
                ReDim Preserve DB(UBound(DB) + 1)
                
                With DB(UBound(DB))
                    .Username = LCase(strArray(1))
                    .Access = gAcc.Access
                    .Flags = gAcc.Flags
                    .ModifiedBy = Username
                    .ModifiedOn = Now
                    .AddedBy = Username
                    .AddedOn = Now
                End With
                
                WriteDatabase (u)
                                        
                strSend = strSend & "."
                
                '// c will control at this point whether or not flags have been erased
                If InStr(1, strSend, "to .", vbTextCompare) > 0 Then
                    strSend = "User " & strArray(1) & "'s flags erased."
                End If
                
                If BotVars.LogDBActions Then
                    LogDBAction AddEntry, Username, strArray(1), Message
                End If
                
                Select Case Track
                    Case 1: strSend = strSend & " He/she is safelisted."
                    Case 2: strSend = strSend & " He/she is no longer safelisted."
                End Select
                
                Call LoadDatabase
                
            Else
                
                strSend = "Please specify access or flags."
                
            End If
            
            
            b = True
            GoTo Display
            
AddError:
            strSend = "Add error - Make sure you specified a username and access amount."
            b = True
            GoTo Display
            
        ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "mmail ") Then
        
            strArray = Split(Message, " ", 3)
            
            If UBound(strArray) = 2 Then
                strSend = "Mass mailing "

                With Temp
                    .From = Username
                    .Message = strArray(2)
                    
                    If StrictIsNumeric(strArray(1)) Then
                        'number games
                        Track = Val(strArray(1))
                        
                        
                        For c = 0 To UBound(DB)
                            If DB(c).Access = Track Then
                                .To = DB(c).Username
                                Call AddMail(Temp)
                            End If
                        Next c
                        
                        strSend = strSend & "to users with access " & Track
                    Else
                        'word games
                        strArray(1) = UCase(strArray(1))
                        
                        For c = 0 To UBound(DB)
                            For f = 1 To Len(strArray(1))
                                If InStr(DB(c).Flags, Mid$(strArray(1), f, 1)) > 0 Then
                                    .To = DB(c).Username
                                    Call AddMail(Temp)
                                    Exit For
                                End If
                            Next f
                        Next c
                        
                        strSend = strSend & "to users with any of the flags " & strArray(1)
                    End If
                End With
                
                strSend = strSend & " complete."
            Else
                strSend = "Format: .mmail <flag(s)> <message> OR .mmail <access> <message>"
            End If
            
            b = True
            GoTo Display
                
        ElseIf (Left$(cMsg, 7) = (BotVars.Trigger & "bmail ")) Or (Left$(cMsg, 6) = BotVars.Trigger & "mail " And Not InBot) Then
            On Error GoTo MailError
            
            strArray = Split(Message, " ", 3)
            
            'For iWinamp = 0 To UBound(strArray)
            '    Debug.Print iWinamp & ": " & strArray(iWinamp)
            'Next
            
            If UBound(strArray) > 1 Then
                Temp.From = Username
                Temp.To = strArray(1)
                Temp.Message = strArray(2)
                
                Call AddMail(Temp)
                
                strSend = "Added mail for " & strArray(1) & "."
            Else
                strSend = "Error processing mail."
            End If
            
            b = True
            GoTo Display
            
MailError:
            
            strSend = "Error processing mail."
            b = True
            GoTo Display
            
        ElseIf Left$(cMsg, 6) = BotVars.Trigger & "mail " Then
            AddQ "/mail " & Mid$(cMsg, 7)
            
            GoTo theEnd
            
                    
        ElseIf cMsg = BotVars.Trigger & "designated" Then
        
            If MyFlags <> 2 And MyFlags <> 18 Then
                strSend = "The bot does not currently have ops."
            ElseIf gChannel.Designated = vbNullString Then
                strSend = "No users have been designated."
            Else
                strSend = "I have designated """ & gChannel.Designated & """."
            End If
            b = True
            GoTo Display
                
        ElseIf cMsg = BotVars.Trigger & "flip" Then
        
            Randomize
            i = Rnd * 200 + 1
            
            If i <= 100 Then
                strSend = "Tails."
            Else
                strSend = "Heads."
            End If
            b = True
            GoTo Display

        ElseIf cMsg = BotVars.Trigger & "ver" Or cMsg = BotVars.Trigger & "about" Or cMsg = BotVars.Trigger & "version" Then
            strSend = ".: " & CVERSION & " by Stealth."
            
            If InStr(1, strSend, CVERSION, vbTextCompare) = 0 Then
                MsgBox frmChat.GetHexProtectionMessage, vbCritical + vbOKOnly
                Call frmChat.Form_Unload(0)
            End If
            
            b = True
            GoTo Display
            
        ElseIf cMsg = (BotVars.Trigger & "server") Then
            strSend = "I am currently connected to " & BotVars.Server & "."
            b = True
            GoTo Display
            
        ElseIf Left$(cMsg, 7) = BotVars.Trigger & "findr " Then
            'Find in a Range added 4/12/06 thanks to a suggestion by rush4hire
            ' find 20 30
            ' c: upper bound
            ' n: lower bound
            ' y: previous message
            
            strArray() = Split(cMsg, " ")
            strSend = "User(s) found: "
            
            If UBound(strArray) = 2 Then
                ' OK: run the command
                If StrictIsNumeric(strArray(1)) And StrictIsNumeric(strArray(2)) Then
                    If Val(strArray(1)) < 1001 And Val(strArray(2)) < 1001 Then
                        c = Val(strArray(2))
                        n = Val(strArray(1))
                    Else
                        strSend = "You specified an invalid range for that command."
                        b = True
                        GoTo Display
                    End If
                    
                    For i = LBound(DB) To UBound(DB)
                        If DB(i).Access >= n And DB(i).Access <= c Then
                            
                            b = True
                            strSend = strSend & ", " & DB(i).Username & IIf(DB(i).Access > 0, "\" & DB(i).Access, vbNullString) & IIf(DB(i).Flags <> vbNullString, "\" & DB(i).Flags, vbNullString)
                            
                            If Len(strSend) > 80 And i <> UBound(DB) Then
                                If LenB(Y) > 0 Then
                                    FilteredSend Username, Y & " [more]", WhisperCmds, InBot, PublicOutput
                                End If
                                
                                Y = Replace(strSend, " , ", " ")
                                Y = Replace(Y, ": , ", ": ")
                                    
                                strSend = "User(s) found: "
                            End If
                            
                        End If
                    Next i
                    
                    If LenB(Y) > 0 Then
                        If b Then
                            FilteredSend Username, Y, WhisperCmds, InBot, PublicOutput
                        Else
                            FilteredSend Username, Left$(Y, Len(Y) - 2), WhisperCmds, InBot, PublicOutput
                        End If
                    ElseIf StrComp(strSend, "User(s) found: ") = 0 Then
                        strSend = "No users were found in that range."
                    Else
                        strSend = Replace(strSend, " , ", " ")
                        strSend = Replace(strSend, ": , ", ": ")
                    End If
                    
                    b = True
                    GoTo Display
                Else
                    strSend = "You specified an invalid range for that command."
                    b = True
                    GoTo Display
                End If
                
            Else
                ' BAD: Syntax error
                strSend = "You specified an invalid range for that command."
                b = True
                GoTo Display
            End If
            
                    
        ElseIf Left$(cMsg, 6) = (BotVars.Trigger & "find ") Then
        
            u = GetFilePath("users.txt")
            
            If Dir$(u) = vbNullString Then
                strSend = "No userlist available. Place a users.txt file in the bot's root directory."
                b = True
                GoTo Display
            End If
            
            u = Right(Message, (Len(Message) - 6))
            If StrictIsNumeric(u) Then
                If Len(u) < 4 Then
                    Call WildCardFind(u, 1, Username, InBot, Val(u), WhisperCmds, PublicOutput)
                    GoTo theEnd
                End If
            End If
            
            If InStr(1, u, "*", vbTextCompare) <> 0 Or InStr(1, u, "?", vbTextCompare) <> 0 Then
                Call WildCardFind(u, 0, Username, InBot, , WhisperCmds, PublicOutput)
                GoTo theEnd
            End If
            
            gAcc = GetAccess(u)
            If gAcc.Access > 0 Then
                If gAcc.Flags <> vbNullString Then
                    strSend = "Found user " & u & ", with access " & gAcc.Access & " and flags " & gAcc.Flags & "."
                Else
                    strSend = "Found user " & u & ", with access " & gAcc.Access & "."
                End If
            Else
                If gAcc.Flags <> vbNullString Then
                    strSend = "Found user " & u & ", with flags " & gAcc.Flags & "."
                Else
                    strSend = "User not found."
                End If
            End If
            
            b = True
            GoTo Display
sendit2:
            strSend = "Who do you want me to find?"
            b = True
            GoTo Display
            
        ElseIf Left$(cMsg, 7) = (BotVars.Trigger & "whois ") Then
            u = Right(Message, Len(Message) - 7)
            
            If InBot And Not PublicOutput Then
                AddQ "/whois " & u, 1
            End If
            
            Call Commands(dbAccess, Username, BotVars.Trigger & "find " & u, InBot, CC, WhisperedIn, PublicOutput)
            GoTo theEnd
                    
                    
        ElseIf Left$(cMsg, 10) = BotVars.Trigger & "findattr " Or Left$(cMsg, 10) = BotVars.Trigger & "findflag " Then
            
            u = UCase(Mid(cMsg, 11, 1))
            
            Response = "Users found: "
            Track = 0
            Track = -1
            
            For i = LBound(DB) To UBound(DB)
                If InStr(1, DB(i).Flags, u, vbTextCompare) > 0 Then
                    Track = 1
                    Response = Response & DB(i).Username & ", "
                    
                    If Len(Response) > 80 Then
                        If LenB(PreviousResponse) > 0 Then
                            FilteredSend Username, PreviousResponse & " [more]", WhisperCmds, InBot, PublicOutput
                        End If
                        
                        PreviousResponse = Left$(Response, Len(Response) - 2)
                        
                        Response = "Users found: "
                        Track = 0
                    End If
                End If
            Next i
            
            If LenB(PreviousResponse) Then
                FilteredSend Username, PreviousResponse & IIf(Track > 0, " [more]", ""), WhisperCmds, InBot, PublicOutput
                PreviousResponse = ""
            End If
                        
            If Track < 0 Then
                strSend = "No user(s) match that flag."
                b = True
                GoTo Display
            ElseIf Track = 1 Then
                FilteredSend Username, Left$(Response, Len(Response) - 2), WhisperCmds, InBot, PublicOutput
            End If
            
            GoTo theEnd
            
        Else
            RealCommand = False
        End If ' huge IF statement
        
        If dbAccess.Access > -1 Then
AccessZero:
            If CC = 0 Then          '// it could be a CC, evaluate
                Call ProcessCC(Username, dbAccess, Message, WhisperedIn, PublicOutput)
                GoTo theEnd
            ElseIf CC = 2 Then
                'Debug.Print Message
                AddQ Message
            End If
        End If
        
    ElseIf CC = 1 Then
        
        If Not WhisperedIn And dbAccess.Access > -1 Then
            AddQ Message
        End If

    End If

Display:
    Commands = Message

    If b Then
        RealCommand = True
        
        If Len(strSend) > 0 Then
            If InBot = True And Not PublicOutput Then
                frmChat.AddChat RTBColors.ConsoleText, strSend
            Else
                If WhisperCmds And Not PublicOutput Then
                    If Left$(cMsg, 5) <> (BotVars.Trigger & "say ") Then
                        If Left$(strSend, 1) <> "/" Then
                            If Dii Then strSend = "/w *" & Username & Space(1) & strSend Else strSend = "/w " & Username & Space(1) & strSend
                        End If
                    End If
                End If
                AddQ strSend
            End If
        End If
    Else
        If InBot = True Then AddQ Message
    End If
    
theEnd:
    Debug.Print "RC: " & RealCommand
    
    If RealCommand And Not CC = 2 Then
        'Debug.Print "command logged"
        If StrComp(Username, CurrentUsername, vbBinaryCompare) = 0 Then
            Username = "(bot console)"
        End If
        
        LogCommand Username, cMsg
        
        Y = Mid$(Split(cMsg, " ")(0), 2)
    End If
    
    If Not TriggerChange Then BotVars.Trigger = OldTrigger
    
    Close #f
End Function

Public Function Cache(ByVal Inpt As String, ByVal Mode As Byte, Optional ByRef Typ As String) As String
    Static s() As String
    Static sTyp As String
    Dim i As Integer
    
    'Debug.Print "cache input: " & Inpt
    
    If InStr(1, LCase(Inpt), "in channel ", vbTextCompare) = 0 Then
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

Public Sub WildCardBan(ByVal sMatch As String, ByVal sMessage As String, ByVal Banning As Byte) ', Optional ExtraMode As Byte)
    'Values for Banning byte:
    '0 = Kick
    '1 = Ban
    '2 = Unban
    Dim i As Integer, Typ As String, z As String
    Dim iSafe As Integer
    
    If sMessage = vbNullString Then sMessage = sMatch
    sMatch = PrepareCheck(sMatch)
    'frmchat.addchat rtbcolors.ConsoleText, "Fired."
    'frmchat.addchat rtbcolors.ConsoleText, "Initial sMessage: " & sMessage
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
                                    If LenB(.Username) > 0 And (.Flags <> 2 And .Flags <> 18) Then AddQ "/" & Typ & .Username & Space(1) & sMessage, 1
                                Else
                                    iSafe = iSafe + 1
                                End If
    '                        Else
    '                            If Not .Safelisted Then
    '                                If .Username <> vbNullString And (.Flags <> 2 Or .Flags <> 18) Then AddQ "/" & Typ & .Username & Space(1) & sMessage, 1
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
            If StrComp(sMessage, ProtectMsg, vbTextCompare) <> 0 Then
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

Public Sub WildCardFind(ByVal bMatch As String, ByVal Mode As Byte, ByVal User As String, ByVal InBot As Boolean, Optional iAccess As Integer, Optional ByVal WhisperCmds As Boolean, Optional ByVal PublicOutput As Boolean)
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
                            FilteredSend User, PrevMsg & " [more]", WhisperCmds, InBot, PublicOutput
                        End If
                        
                        PrevMsg = Replace(ReturnMsg, " , ", " ")
                        PrevMsg = Replace(PrevMsg, ": , ", ": ")
                            
                        ReturnMsg = "User(s) found: "
                    End If
                    
                End If
            Next i
            
            If LenB(PrevMsg) > 0 Then
                If Found Then
                    FilteredSend User, PrevMsg, WhisperCmds, InBot, PublicOutput
                Else
                    FilteredSend User, Left$(PrevMsg, Len(PrevMsg) - 2), WhisperCmds, InBot, PublicOutput
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
                            If Dii Then AddQ "/w *" & User & Space(1) & ReturnMsg Else AddQ "/w *" & User & Space(1) & ReturnMsg
                        ElseIf InBot And Not PublicOutput Then
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
            If Dii Then AddQ "/w *" & User & Space(1) & ReturnMsg Else AddQ "/w " & User & Space(1) & ReturnMsg
        ElseIf InBot And Not PublicOutput Then
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
        If Dii Then AddQ "/w *" & User & Space(1) & ReturnMsg Else AddQ "/w " & User & Space(1) & ReturnMsg
    ElseIf InBot And Not PublicOutput Then
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
            
            If StrComp(LCase(rItem), LCase(strCompare), vbTextCompare) = 0 Then GoTo Successful
        End If
    Next Counter
    
    RemoveItem = "No such user found."
    GoTo theEnd
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
'            If StrComp(LCase(Username), LCase(.Username(i)), vbTextCompare) = 0 Then
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
    Dim f As Integer, strCompare As String, Ret As String
    
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
                    Ret = Ret & strCompare & " "
                End If
            Loop
        End If
        Close #f
    End If
    
    GetSafelistMatches = ReversePrepareCheck(Trim(Ret))
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
    
    Username = LCase(Username)
    
    On Error Resume Next
    Temp = GetFilePath("autobans.txt")
    
    If Dir$(Temp) <> vbNullString Then Open Temp For Input As #f Else GoTo theEnd
    
    If LOF(f) < 2 Then GoTo theEnd
    Do
        Line Input #f, strCompare
        toCheck = LCase(strCompare)
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

Public Function ProcessCC(ByVal Speaker As String, ByRef SpeakerAccess As udtGetAccessResponse, RawMessage As String, ByVal WhisperedIn As Boolean, ByVal PublicOutput As Boolean, Optional ByVal ExistenceCheckOnly As Boolean = False) As Boolean
    Dim Args() As String, Send() As String, Actions() As String
    Dim ccIn As udtCustomCommandData
    Dim n As Integer, HighestArgUsed As Integer, c As Integer, i As Integer, f As Integer
    Dim Found As Boolean, FirstTime As Boolean
    Dim Temp As String, ReplaceString As String
    
    On Error GoTo ProcessCC_Error

    f = FreeFile
    
    ReDim Send(0)
    
    If LenB(Dir$(GetFilePath("commands.dat"))) > 0 Then
    
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccIn)
        
        If LOF(f) > 1 Then
        
            ' /****** Step 1: Get data ******/
        
            HighestArgUsed = -1
            
            i = LOF(f) \ LenB(ccIn)
            If LOF(f) Mod LenB(ccIn) <> 0 Then i = i + 1
            
            For i = 1 To i
            
                Get #f, i, ccIn
                
                If ccIn.reqAccess > 1000 Then GoTo NextItem
                If LenB(RTrim(ccIn.Action)) < 1 Then GoTo NextItem
                If LenB(RTrim(ccIn.Query)) < 1 Then GoTo NextItem
                If ccIn.reqAccess = 0 Then ccIn.reqAccess = -1 'zero-access command
                
                If SpeakerAccess.Access >= ccIn.reqAccess Then
                    Args() = Split(RawMessage, " ")
                    
                    '   0    1   2  3  4     5
                    ' .myCC say %1 is fat! %rest
                    If StrComp(RTrim(LCase(ccIn.Query)), LCase(Mid$(Args(0), 2))) = 0 Then
                        Found = True
                        ProcessCC = True
                        
                        If ExistenceCheckOnly Then GoTo theEnd
                    
                        Actions = Split(RTrim(Replace(ccIn.Action, vbCrLf, "")), "& ")
                    
                        If UBound(Actions) = 0 Then
                        
                            ReDim Send(0)
                            Send(0) = Actions(0)
                    
                        Else
                            ReDim Send(UBound(Actions))
                            
                            For c = 0 To UBound(Actions)
                                Send(c) = Actions(c)
                            Next c
                        End If
                
                    
                        FirstTime = True
                        
                        For n = 0 To UBound(Send)
                            Send(n) = Replace(Send(n), "%0", Speaker)
                            Send(n) = Replace(Send(n), "%bc", BanCount)
                            
                            If UBound(Args) > 0 Then
                                If InStr(Send(n), "%") > 0 Then
                                    ' has probable arguments
                                    ' /****** Step 2: Do basic replacements ******/
                                    HighestArgUsed = 0
                                    
                                    ' This was a normal FOR loop then it stopped working
                                    '  altogether. So, it's a while loop now..
                                    c = UBound(Args)
                                    
                                    While c >= 0
                                        If InStr(Send(n), "%" & c) Then
                                            Send(n) = Replace(Send(n), "%" & (c), Args(c))
                                            If HighestArgUsed = 0 Then HighestArgUsed = c
                                        End If
                                        
                                        c = c - 1
                                    Wend
                                    
                                    ' Assemble the %rest string
                                    If FirstTime Then
                                        'Debug.Print "firsttime, highest: " & c
                                    
                                        If HighestArgUsed > -1 And UBound(Args) > HighestArgUsed Then
                                            For c = HighestArgUsed + 1 To UBound(Args)
                                                ReplaceString = ReplaceString & Args(c) & IIf(c < UBound(Args), " ", "")
                                            Next c
                                        End If
                                        
                                        FirstTime = False
                                    End If
                                    
                                    ' /****** Step 3: Do advanced replacements ******/
                                    If LenB(ReplaceString) > 0 Then
                                        Send(n) = Replace(Send(n), "%rest", ReplaceString)
                                    End If
                                End If
                            End If
                        Next n
                        
                        ' /****** Step 4: Send and/or process normal command ******/
                        For c = 0 To UBound(Send)
                            'Debug.Print Send(c)
                            
                            If Left$(Send(c), 1) = "/" Then
                                Call Commands(SpeakerAccess, Speaker, Send(c), False, 2, WhisperedIn, PublicOutput)
                            Else
                                If Left$(Send(c), 6) = "_call " Then
                                    Temp = Split(Mid$(Send(c), 7), " ")(0)
                                    
                                    On Error Resume Next
                                    frmChat.SControl.Run Temp
                                Else
                                    If Left$(RawMessage, 1) = "/" And Not PublicOutput Then
                                        frmChat.AddChat RTBColors.ConsoleText, Send(c)
                                    Else
                                        ' // including other term replacements
                                        Send(c) = DoReplacements(Send(c), Speaker, GetPing(Speaker))
                                        AddQ Send(c)
                                    End If
                                End If
                            End If
                        Next c
                        
                    End If 'strcomp
                End If 'speakeraccess
NextItem:
            Next i
            
        End If 'lof
    
    End If
    
theEnd:
    If Not Found And Not ExistenceCheckOnly Then
        If StrComp(Left$(RawMessage, 1), "/", vbTextCompare) = 0 Then _
            Call Commands(SpeakerAccess, Speaker, RawMessage, False, 1, WhisperedIn, PublicOutput)
    ElseIf ExistenceCheckOnly Then
        ProcessCC = Found
    End If
    
    Close #f
    Exit Function
Error:
    AddQ "Invalid argument(s) or incorrect number of arguments."
    Close #f

    On Error GoTo 0
    Exit Function

ProcessCC_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessCC of Module modCommandCode"
End Function

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
    PrepareCheck = LCase(toCheck)
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
    ReversePrepareCheck = LCase(toCheck)
End Function


Public Sub DBRemove(ByVal s As String)
    
    Dim i As Integer
    Dim c As Integer
    Dim n As Integer
    Dim t() As udtDatabase
    Dim Temp As String
    s = LCase(s)
    
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
                            .Username = LCase(X(0))
                            
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
        
        CWord = Mid$(CWord, 2)
        
        If LenB(ReadINI("DisabledCommands", CWord, "access.ini")) > 0 Or ReadINI("DisabledCommands", "universal", "access.ini") = "Y" Then
            ValidateAccess = False
            Exit Function
        End If
        
        Temp = UCase(ReadINI("Flags", CWord, "access.ini"))
        
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


Public Sub FilteredSend(ByVal Username As String, ByVal ToSend As String, ByVal WhisperCmds As Boolean, ByVal InBot As Boolean, ByVal PublicOutput As Boolean)
    If InBot And Not PublicOutput Then
        frmChat.AddChat RTBColors.ConsoleText, ToSend
    ElseIf WhisperCmds And Not PublicOutput Then
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
