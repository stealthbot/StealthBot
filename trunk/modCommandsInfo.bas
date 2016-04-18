Attribute VB_Name = "modCommandsInfo"
Option Explicit
'This module will hold all of the 'Info' Commands
'Commands that return information, but have really no functionality

Public Sub OnAbout(Command As clsCommandObj)
    Command.Respond ".: " & CVERSION & " :."
End Sub

Public Sub OnAccountInfo(Command As clsCommandObj)
    If ((Not Command.IsLocal) Or Command.PublicOutput) Then
        PPL = True
        If ((BotVars.WhisperCmds Or Command.WasWhispered) And (Not Command.IsLocal)) Then
            PPLRespondTo = Command.Username
        End If
    End If
    RequestSystemKeys
End Sub

Public Sub OnBanCount(Command As clsCommandObj)
    If (g_Channel.BanCount = 0) Then
        Command.Respond "No users have been banned since I joined this channel."
    Else
        Command.Respond StringFormat("Since I joined this channel, {0} user{1} have been banned.", g_Channel.BanCount, IIf(g_Channel.BanCount > 1, "s", vbNullString))
    End If
End Sub

Public Sub OnBanListCount(Command As clsCommandObj)
    If (g_Channel.Banlist.Count = 0) Then
        Command.Respond "There are no users on the internal ban list."
    Else
        Command.Respond StringFormat("There {0} currently {1} user{2} on the internal ban list.", _
            IIf(g_Channel.Banlist.Count > 1, "are", "is"), g_Channel.Banlist.Count, _
            IIf(g_Channel.Banlist.Count > 1, "s", vbNullString))
    End If
End Sub

Public Sub OnBanned(Command As clsCommandObj)
    Dim sResult  As String
    Dim i        As Integer
    Dim j        As Integer
    Dim BanCount As Integer
    
    If (g_Channel.Banlist.Count = 0) Then
        Command.Respond "There are presently no users on the bot's internal banlist."
    Else
        sResult = "User(s) banned: "
        For i = 1 To g_Channel.Banlist.Count
            If (Not g_Channel.Banlist(i).IsDuplicateBan) Then
                For j = 1 To g_Channel.Banlist.Count
                    If (StrComp(g_Channel.Banlist(j).DisplayName, g_Channel.Banlist(i).DisplayName, vbTextCompare) = 0) Then
                        BanCount = (BanCount + 1)
                    End If
                Next j
                sResult = StringFormat("{0}, {1}", sResult, g_Channel.Banlist(i).DisplayName)
                
                If (BanCount > 1) Then sResult = StringFormat("{0} ({1})", sResult, BanCount)
                      
                If ((Len(sResult) > 90) And (Not i = g_Channel.Banlist.Count)) Then
                    Command.Respond Replace(sResult, " , ", Space$(1))
                    sResult = "Users(s) banned: "
                End If
            End If
            BanCount = 0
        Next i
        If (LenB(sResult) > LenB("Users(s) banned: , ")) Then 'We don't want to send an empty line
            Command.Respond Replace(sResult, " , ", Space$(1))
        End If
    End If
End Sub

Public Sub OnClientBans(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Command.IsValid) Then
        Call SearchDatabase(bufResponse(), , , , "GAME", , , "B")
        
        For Each strResponse In bufResponse
            Command.Respond CStr(strResponse)
        Next
    End If
End Sub

Public Sub OnDetail(Command As clsCommandObj)
    If (Command.IsValid) Then
        Dim sRetAdd As String, sRetMod As String
        Dim i As Integer
        
        For i = 0 To UBound(DB)
            With DB(i)
                If (StrComp(Command.Argument("Username"), .Username, vbTextCompare) = 0) Then
                    If ((Not .AddedBy = "%") And (LenB(.AddedBy) > 0)) Then
                        sRetAdd = StringFormat("{0} was added by {1} on {2}.", .Username, .AddedBy, .AddedOn)
                    End If
                    
                    If ((Not .ModifiedBy = "%") And (LenB(.ModifiedBy) > 0)) Then
                        If ((Not .AddedOn = .ModifiedOn) Or (Not StrComp(.AddedBy, .ModifiedBy, vbTextCompare) = 0)) Then
                            sRetMod = StringFormat(" The entry was last modified by {0} on {1}.", .ModifiedBy, .ModifiedOn)
                        Else
                            sRetMod = " The entry has not been modified since it was added."
                        End If
                    End If
                    
                    If ((LenB(sRetAdd) > 0) Or (LenB(sRetMod) > 0)) Then
                        If (LenB(sRetAdd) > 0) Then
                            Command.Respond sRetAdd & sRetMod
                        Else
                            Command.Respond Trim$(sRetMod)
                        End If
                    Else
                        Command.Respond "No detailed information is available for that user."
                    End If
                    
                    Exit Sub
                End If
            End With
        Next i
        
        Command.Respond "That user was not found in the database."
    End If
End Sub

Public Sub OnFind(Command As clsCommandObj)
    Dim dbAccess      As udtGetAccessResponse
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Not Command.IsValid) Then Exit Sub
    
    If (LenB(Dir$(GetFilePath(FILE_USERDB))) = 0) Then
       Command.Respond "No userlist available. Place a users.txt file in the bot's root directory."
       Exit Sub
    End If
    
    ReDim Preserve bufResponse(0)
    
    If (StrictIsNumeric(Command.Argument("Username/Rank"))) Then
        If (LenB(Command.Argument("UpperRank")) = 0) Then
            Call SearchDatabase(bufResponse(), , , , , Val(Command.Argument("Username/Rank")))
        Else
            Dim LowerRank As Integer
            Dim UpperRank As Integer
            
            LowerRank = Val(Command.Argument("Username/Rank"))
            UpperRank = Command.Argument("UpperRank")
            
            If (UpperRank = LowerRank) Then
                Call SearchDatabase(bufResponse(), , , , , LowerRank)
            ElseIf (UpperRank > LowerRank) Then
                Call SearchDatabase(bufResponse(), , , , , LowerRank, UpperRank)
            Else
                Call SearchDatabase(bufResponse(), , , , , UpperRank, LowerRank)
            End If
        End If
    Else
        Call SearchDatabase(bufResponse(), , PrepareCheck(Command.Argument("Username/Rank")))
    End If
    
    For Each strResponse In bufResponse
        Command.Respond CStr(strResponse)
    Next
End Sub

Public Sub OnFindAttr(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Command.IsValid) Then
        Call SearchDatabase(bufResponse(), , , , , , , Command.Argument("Attributes"))
        For Each strResponse In bufResponse
            Command.Respond CStr(strResponse)
        Next
    Else
        Command.Respond "You must specify flag(s) to search for."
    End If
End Sub

Public Sub OnFindGrp(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Command.IsValid) Then
        Call SearchDatabase(bufResponse(), , , Command.Argument("Group"))
        For Each strResponse In bufResponse
            Command.Respond CStr(strResponse)
        Next
    Else
        Command.Respond "You must specify a group to find."
    End If
End Sub

Public Sub OnHelp(Command As clsCommandObj)
    Dim strCommand As String
    Dim strScript  As String
    Dim docs       As clsCommandDocObj
    
    strCommand = IIf(Command.IsValid, Command.Argument("Command"), "help")
    strScript = IIf(LenB(Command.Argument("ScriptOwner")) > 0, Command.Argument("ScriptOwner"), Chr$(0))
    
    Set docs = OpenCommand(strCommand, strScript)
    If (LenB(docs.Name) = 0) Then
        Command.Respond "Sorry, but no related documentation could be found."
    Else
        If (docs.aliases.Count > 1) Then
            Command.Respond StringFormat("[{0} (Aliases: {4})]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString(Command.IsLocal), docs.RequirementsStringShort, docs.AliasString)
        ElseIf (docs.aliases.Count = 1) Then
            Command.Respond StringFormat("[{0} (Alias: {4})]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString(Command.IsLocal), docs.RequirementsStringShort, docs.AliasString)
        Else
            Command.Respond StringFormat("[{0}]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString(Command.IsLocal), docs.RequirementsStringShort)
        End If
    End If
    Set docs = Nothing
    
End Sub

Public Sub OnHelpAttr(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    
    If (Command.IsValid) Then
        tmpbuf = GetAllCommandsFor(Command.docs, , Command.Argument("Flags"))
        If (LenB(tmpbuf) > 0) Then
            Command.Respond "Commands available to specified flag(s): " & tmpbuf
        Else
            Command.Respond "No commands are available to the given flag(s)."
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnHelpAttr()."
End Sub

Public Sub OnHelpRank(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    
    If (Command.IsValid) Then
        If (CInt(Command.Argument("Rank")) > -1) Then
            tmpbuf = GetAllCommandsFor(Command.docs, Command.Argument("Rank"))
            If (LenB(tmpbuf) > 0) Then
                Command.Respond "Commands available to specified rank: " & tmpbuf
            Else
                Command.Respond "No commands are available to the given rank."
            End If
        Else
            Command.Respond "The specified rank must be greater or equal to zero."
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnHelpRank()."
End Sub

Public Sub OnInfo(Command As clsCommandObj)
    Dim UserIndex As Integer
    
    If (Command.IsValid) Then
        UserIndex = g_Channel.GetUserIndex(Command.Argument("Username"))
        
        If (UserIndex > 0) Then
            With g_Channel.Users(UserIndex)
                Command.Respond StringFormat("User {0} is logged on using {1} with {2}a ping time of {3}ms.", _
                    .DisplayName, ProductCodeToFullName(.Game), _
                    IIf(.IsOperator, "ops, and ", vbNullString), .Ping)
            
                Command.Respond StringFormat("He/she has been present in the channel for {0}.", ConvertTime(.TimeInChannel(), 1))
            End With
        Else
            Command.Respond "No such user is present."
        End If
    End If
End Sub

Public Sub OnInitPerf(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim ModName As String
    Dim Name    As String
    Dim i       As Integer
    Dim strRet  As String
    Dim Script  As Module
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled via the override."
        Exit Sub
    End If
    
    If (LenB(Command.Argument("Script")) > 0) Then
        Set Script = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Script Is Nothing) Then
            Command.Respond "Could not find the script specified."
        Else
            If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 1) Then
                Command.Respond StringFormat("The Script {0} loaded in {1}ms.", _
                    GetScriptName(Script.Name), _
                    GetScriptDictionary(Script)("InitPerf"))
            Else
                Command.Respond "That script is currently disabled."
            End If
        End If
    Else
        If (frmChat.SControl.Modules.Count > 1) Then
            If (Command.IsLocal And Not Command.PublicOutput) Then
                Command.Respond "Script initialization performance:"
            Else
                strRet = "Script initialization performance:"
            End If
            For i = 2 To frmChat.SControl.Modules.Count
                Set Script = frmChat.SControl.Modules(i)
                
                If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 0) Then
                    If (Command.IsLocal And Not Command.PublicOutput) Then
                        Command.Respond StringFormat(" '{0}' loaded in {1}ms.", _
                            GetScriptName(Script.Name), _
                            GetScriptDictionary(Script)("InitPerf"))
                    Else
                        strRet = StringFormat("{0} '{1}' {2}ms{3}", _
                            strRet, _
                            GetScriptName(Script.Name), _
                            GetScriptDictionary(Script)("InitPerf"), _
                            IIf(i = frmChat.SControl.Modules.Count, vbNullString, ","))
                    End If
                End If
            Next i
            
            If (LenB(strRet)) Then Command.Respond strRet
        Else
            Command.Respond "There are no scripts currently loaded."
        End If
    End If
        
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnInitPerf()."
End Sub

Public Sub OnLastSeen(Command As clsCommandObj)
    Dim retVal As String
    Dim i      As Integer

    If (colLastSeen.Count = 0) Then
        retVal = "I have not seen anyone yet."
    Else
        retVal = "Last 15 users seen: "
        For i = 1 To colLastSeen.Count
            retVal = StringFormat("{0}{1}{2}", _
                retVal, colLastSeen.Item(i), _
                IIf(i = colLastSeen.Count, vbNullString, ", "))
            If (i = 15) Then Exit For
        Next i
    End If
    Command.Respond retVal
End Sub

Public Sub OnLastWhisper(Command As clsCommandObj)
    If (LenB(LastWhisper) > 0) Then
        Command.Respond StringFormat("The last whisper to this bot was from {0} at {1} on {2}.", _
            LastWhisper, _
            FormatDateTime(LastWhisperFromTime, vbLongTime), _
            FormatDateTime(LastWhisperFromTime, vbLongDate))
    Else
        Command.Respond "The bot has not been whispered since it logged on."
    End If
End Sub

Public Sub OnLocalIp(Command As clsCommandObj)
    Command.Respond StringFormat("{0} local IPv4 IP address is: {1}", IIf(Command.IsLocal, "Your", "My"), frmChat.sckBNet.LocalIP)
End Sub

Public Sub OnOwner(Command As clsCommandObj)
    If (LenB(BotVars.BotOwner)) Then
        Command.Respond "This bot's owner is " & BotVars.BotOwner & "."
    Else
        Command.Respond "There is no owner currently set."
    End If
End Sub

Public Sub OnPhrases(Command As clsCommandObj)
    Dim sSubCommand As String
    Dim iCount As Integer
    
    If (Command.IsValid) Then
        sSubCommand = Command.Argument("subcommand")
        
        If (LenB(sSubCommand) = 0) Or (LCase$(sSubCommand) = "list") Then
            If ListRespond(Command, Phrases, "Banned phrases{%p}: ", IIf(LenB(sSubCommand) = 0, 2, -1)) Then
                Exit Sub
            End If
        End If
        
        iCount = GetListCount(Phrases)
        If iCount = 0 Then
            Command.Respond "There are no banned phrases."
        Else
            Command.Respond StringFormat("There are {0} banned phrases. Use the '{1} list' command to show them.", iCount, Command.Name)
        End If
    Else
        Command.Respond StringFormat("Invalid command. Correct usage: {0} [list/count]", Command.Name)
    End If
End Sub


Public Sub OnPing(Command As clsCommandObj)
    Dim Latency As Long
    If (Command.IsValid) Then
        Latency = GetPing(Command.Argument("Username"))
        If (Latency >= -1) Then
            Command.Respond StringFormat("{0}'s ping at login was {1}ms.", Command.Argument("Username"), Latency)
        Else
            Command.Respond StringFormat("I can not see {0} in the channel.", Command.Argument("Username"))
        End If
    Else
        Command.Respond "Please specify a user to ping."
    End If
End Sub

Public Sub OnPingMe(Command As clsCommandObj)
    Dim Latency As Long
    If (Command.IsLocal) Then
        If (g_Online) Then
            Command.Respond StringFormat("Your ping at login was {0}ms.", GetPing(GetCurrentUsername))
        Else
            Command.Respond "Error: You are not logged on."
        End If
    Else
        Latency = GetPing(Command.Username)
        If (Latency >= -1) Then
            Command.Respond StringFormat("Your ping at login was {0}ms.", Latency)
        Else
            Command.Respond "I can not see you in the channel."
        End If
    End If
End Sub

Public Sub OnProfile(Command As clsCommandObj)
    If (Command.IsValid) Then
        If ((Not Command.IsLocal) Or (Command.PublicOutput)) Then
            PPL = True
            If ((BotVars.WhisperCmds Or Command.WasWhispered) And (Not Command.IsLocal)) Then
                PPLRespondTo = Command.Username
            Else
                PPLRespondTo = vbNullString
            End If
            Call RequestProfile(Command.Argument("Username"))
        Else
            frmProfile.PrepareForProfile Command.Argument("Username"), False
            Call RequestProfile(Command.Argument("Username"))
        End If
    End If
End Sub

Public Sub OnSafeCheck(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (GetSafelist(Command.Argument("Username"))) Then
            Command.Respond StringFormat("{0} is on the bot's safelist.", Command.Argument("Username"))
        Else
            Command.Respond StringFormat("{0} is not on the bot's safelist.", Command.Argument("Username"))
        End If
    End If
End Sub

' handle safelist command
Public Sub OnSafeList(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim i             As Long
    
    Call SearchDatabase(bufResponse(), , , , , , , "S")
    
    For i = 0 To UBound(bufResponse)
        Command.Respond bufResponse(i)
    Next i
End Sub

Public Sub OnScriptDetail(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim Script As Module
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled via the override."
        Exit Sub
    End If
    
    If (Command.IsValid) Then
        Set Script = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Script Is Nothing) Then
            Command.Respond "Error: Could not find specified script."
        Else
            Dim ScriptInfo  As Dictionary
            Dim Version     As String
            Dim VerTotal    As Integer
            Dim Author      As String
            Dim Description As String
            
            Set ScriptInfo = GetScriptDictionary(Script)
            
            Version = StringFormat("{0}.{1}{2}", Val(ScriptInfo("Major")), Val(ScriptInfo("Minor")), _
                IIf(Val(ScriptInfo("Revision")) > 0, " Revision " & Val(ScriptInfo("Revision")), vbNullString))
                
            VerTotal = Val(ScriptInfo("Major")) + Val(ScriptInfo("Minor")) + Val(ScriptInfo("Revision"))
                     
            Author = ScriptInfo("Author")
            Description = ScriptInfo("Description")
            
            If ((LenB(Author) = 0) And (VerTotal = 0)) Then
                Command.Respond StringFormat("There is no additional information for the '{0}' script.", _
                    GetScriptName(Script.Name))
            Else
                Command.Respond StringFormat("{0}{1}{2}{3}", _
                    GetScriptName(Script.Name), _
                    IIf(VerTotal > 0, " v" & Version, vbNullString), _
                    IIf(LenB(Author) > 0, " by " & Author, vbNullString), _
                    IIf(LenB(Description) > 0, ": " & Description, vbNullString))
            End If
        End If
    End If
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnScriptDetail()."
End Sub

Public Sub OnScripts(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim retVal  As String
    Dim i       As Integer
    Dim Enabled As Boolean
    Dim Name    As String
    Dim Count   As Integer
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled via the override."
        Exit Sub
    End If
    
    If (frmChat.SControl.Modules.Count > 1) Then
        For i = 2 To frmChat.SControl.Modules.Count
            Name = modScripting.GetScriptName(CStr(i))
            Enabled = Not (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 0)
                
            retVal = StringFormat("{0}{1}{2}{3}{4}", _
                retVal, _
                IIf(Enabled, vbNullString, "("), _
                Name, _
                IIf(Enabled, vbNullString, ")"), _
                IIf(i = frmChat.SControl.Modules.Count, vbNullString, ", "))
                
            Count = (Count + 1)
        Next i
        
        Command.Respond StringFormat("Loaded Scripts ({0}): {1}", Count, retVal)
    Else
        Command.Respond "There are no scripts currently loaded."
    End If
    
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnScripts()."
End Sub

Public Sub OnServer(Command As clsCommandObj)
    Dim RemoteHost   As String
    Dim RemoteHostIP As String
    
    RemoteHost = frmChat.sckBNet.RemoteHost
    RemoteHostIP = frmChat.sckBNet.RemoteHostIP
    
    If (StrComp(RemoteHost, RemoteHostIP, vbBinaryCompare) = 0) Then
        Command.Respond "I am currently connected to " & RemoteHostIP & "."
    Else
        Command.Respond "I am currently connected to " & RemoteHost & " (" & RemoteHostIP & ")."
    End If
End Sub

Public Sub OnShitCheck(Command As clsCommandObj)
    Dim dbAccess As udtGetAccessResponse
    Dim compare  As VbCompareMethod
    If (Command.IsValid) Then
        dbAccess = GetCumulativeAccess(Command.Argument("Username"))
        compare = IIf(BotVars.CaseSensitiveFlags, vbBinaryCompare, vbTextCompare)
        If (Not InStr(1, dbAccess.Flags, "B", compare) = 0) Then
            If (Not InStr(1, dbAccess.Flags, "S", compare) = 0) Then
                Command.Respond Command.Argument("Username") & _
                    "{0} is on the bot's shitlist; also on the safelist and will not be banned."
            Else
                Command.Respond Command.Argument("Username") & " is on the bot's shitlist."
            End If
        Else
            Command.Respond Command.Argument("Username") & " is not on the bot's shitlist."
        End If
    End If
End Sub

Public Sub OnShitList(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim i             As Integer
    
    Call SearchDatabase(bufResponse(), , "!*[*]*", , , , , "B")
    
    For i = 0 To UBound(bufResponse)
        Command.Respond bufResponse(i)
    Next i
End Sub

Public Sub OnTagBans(Command As clsCommandObj)
    Dim bufResponse() As String
    Dim i             As Integer
    
    Call SearchDatabase(bufResponse(), , "*[*]*", , , , , "B")
    
    For i = 0 To UBound(bufResponse)
        Command.Respond bufResponse(i)
    Next i
End Sub

Public Sub OnTime(Command As clsCommandObj)
    Command.Respond "The current time on this computer is " & Time & " on " & Format(Date, "MM-dd-yyyy") & "."
End Sub

Public Sub OnTrigger(Command As clsCommandObj)
    If (LenB(BotVars.TriggerLong) = 1) Then
        Command.Respond StringFormat("The bot's current trigger is {0} {1} {0} (Alt +0{2})", _
            Chr$(34), BotVars.TriggerLong, Asc(BotVars.TriggerLong))
    Else
        Command.Respond StringFormat("The bot's current trigger is {0} {1} {0} (Length: {2})", _
          Chr$(34), BotVars.TriggerLong, Len(BotVars.TriggerLong))
    End If
End Sub

Public Sub OnUptime(Command As clsCommandObj)
    Command.Respond StringFormat("System uptime {0}, connection uptime {1}.", ConvertTime(GetUptimeMS), ConvertTime(uTicks))
End Sub

Public Sub OnWhere(Command As clsCommandObj)
    If (Command.IsLocal) Then
        Call frmChat.AddQ("/where " & Command.Args, PRIORITY.COMMAND_RESPONSE_MESSAGE, "(console)")
    End If

    Command.Respond StringFormat("I am currently in channel {0} ({1} users present)", g_Channel.Name, g_Channel.Users.Count)
End Sub

Public Sub OnWhoAmI(Command As clsCommandObj)
    Dim dbAccess As udtGetAccessResponse

    If (Command.IsLocal) Then
        Command.Respond "You are the bot console."
        
        If (g_Online) Then
            Call frmChat.AddQ("/whoami", PRIORITY.CONSOLE_MESSAGE)
        End If
    Else
        dbAccess = GetCumulativeAccess(Command.Username)
        If (dbAccess.Rank = 1000) Then
            Command.Respond StringFormat("You are the bot owner, {0}.", Command.Username)
        Else
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond StringFormat("{0} holds rank {1} and flags {2}.", Command.Username, dbAccess.Rank, dbAccess.Flags)
                Else
                    Command.Respond StringFormat("{0} holds rank {1}.", Command.Username, dbAccess.Rank)
                End If
            Else
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond StringFormat("{0} has flags {1}.", Command.Username, dbAccess.Flags)
                Else
                    Command.Respond StringFormat("{0} has no rank or flags.", Command.Username)
                End If
            End If
        End If
    End If
End Sub

Public Sub OnWhoIs(Command As clsCommandObj)
    Dim dbAccess As udtGetAccessResponse
    
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            Call frmChat.AddQ("/whois " & Command.Argument("Username"), PRIORITY.CONSOLE_MESSAGE)
        End If

        dbAccess = GetCumulativeAccess(Command.Argument("Username"))
        
        If (LenB(dbAccess.Username) > 0) Then
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.Username & " holds rank " & dbAccess.Rank & " and flags " & dbAccess.Flags & "."
                Else
                    Command.Respond dbAccess.Username & " holds rank " & dbAccess.Rank & "."
                End If
            Else
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.Username & " has flags " & dbAccess.Flags & "."
                End If
            End If
        Else
            Command.Respond "There was no such user found."
        End If
    End If
End Sub

Private Function GetAllCommandsFor(ByRef commandDoc As clsCommandDocObj, Optional Rank As Integer = -1, Optional Flags As String = vbNullString) As String
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    Dim i           As Integer
    Dim xmlDoc      As DOMDocument60
    Dim commands    As IXMLDOMNodeList
    Dim xpath       As String
    Dim lastCommand As String
    Dim thisCommand As String
    Dim xFunction   As String
    Dim AZ          As String
    Dim Flag        As String
    AZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    'If (LenB(Dir$(GetFilePath(FILE_COMMANDS))) = 0) Then
    '    Command.Respond "Error: The XML database could not be found in the working directory."
    '    Exit Function
    'End If

    If (LenB(Flags) > 0) Then
        If (BotVars.CaseSensitiveFlags) Then
            xFunction = "text()='{0}'"
        Else
            xFunction = StringFormat("translate(text(), '{0}', '{1}')='{2}'", UCase$(AZ), LCase$(AZ), "{0}")
        End If
        
        For i = 1 To Len(Flags)
            Flag = IIf(Mid$(Flags, i, 1) = "\", "\\", Mid$(Flags, i, 1))
            If (Not BotVars.CaseSensitiveFlags) Then Flag = LCase$(Flag)
            If (Flag = "'") Then Flag = "&apos;"
            
            xpath = StringFormat("{0}{1}{2}", _
                xpath, _
                StringFormat(xFunction, Flag), _
                IIf(i = Len(Flags), vbNullString, " or "))
        Next i
        xpath = StringFormat("./command/access/flags/flag[{0}]", xpath)
    Else
        xpath = StringFormat("./command/access/rank[number() <= {0}]", Rank)
    End If
    
    Set xmlDoc = commandDoc.XMLDocument
    
    Set commands = xmlDoc.documentElement.selectNodes(xpath)

    If (commands.length > 0) Then
        For i = 0 To commands.length - 1
            If (LenB(Flags) > 0) Then
                thisCommand = commands(i).parentNode.parentNode.parentNode.Attributes.getNamedItem("name").Text
            Else
                thisCommand = commands(i).parentNode.parentNode.Attributes.getNamedItem("name").Text
            End If
            
            If (StrComp(thisCommand, lastCommand, vbTextCompare) <> 0) Then
                tmpbuf = StringFormat("{0}{1}, ", tmpbuf, thisCommand)
            End If
        
            lastCommand = thisCommand
        Next i
        
        tmpbuf = Left$(tmpbuf, Len(tmpbuf) - 2)
    End If
    GetAllCommandsFor = tmpbuf
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.GetAllCommandsFor()."
End Function

Public Function GetPing(ByVal Username As String) As Long
    Dim i As Integer
    
    i = g_Channel.GetUserIndex(Username)
    
    If i > 0 Then
        GetPing = g_Channel.Users(i).Ping
    Else
        GetPing = -3
    End If
End Function

Private Sub SearchDatabase(ByRef arrReturn() As String, Optional Username As String = vbNullString, _
    Optional ByVal match As String = vbNullString, Optional Group As String = vbNullString, _
        Optional dbType As String = vbNullString, Optional lowerBound As Integer = -1, _
            Optional upperBound As Integer = -1, Optional Flags As String = vbNullString)
    
    On Error GoTo ERROR_HANDLER
    
    Dim i        As Integer
    Dim found    As Integer
    Dim tmpbuf   As String
    ReDim arrReturn(0)
    
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
        For i = LBound(DB) To UBound(DB)
            Dim res        As Boolean
            Dim blnChecked As Boolean
        
            If (LenB(DB(i).Username) > 0) Then
                If (LenB(match) > 0) Then
                    If (Left$(match, 1) = "!") Then
                        res = (Not (LCase$(PrepareCheck(DB(i).Username)) Like (LCase$(Mid$(match, 2)))))
                    Else
                        res = (LCase$(PrepareCheck(DB(i).Username)) Like (LCase$(match)))
                    End If
                    blnChecked = True
                End If
                
                If (LenB(Group) > 0) Then
                    If (StrComp(DB(i).Groups, Group, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If

                If (LenB(dbType) > 0) Then
                    If (StrComp(DB(i).Type, dbType, vbTextCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If ((lowerBound >= 0) And (upperBound >= 0)) Then
                    If ((DB(i).Rank >= lowerBound) And (DB(i).Rank <= upperBound)) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                ElseIf (lowerBound >= 0) Then
                    If (DB(i).Rank = lowerBound) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If (LenB(Flags) > 0) Then
                    Dim j As Integer
                
                    For j = 1 To Len(Flags)
                        If (InStr(1, DB(i).Flags, Mid$(Flags, j, 1), vbBinaryCompare) = 0) Then
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
                    tmpbuf = tmpbuf & DBUserToString(DB(i).Username, DB(i).Type)
                    tmpbuf = StringFormat("{0}{1}{2}, ", tmpbuf, _
                        IIf(DB(i).Rank > 0, "\" & DB(i).Rank, vbNullString), _
                        IIf(LenB(DB(i).Flags) > 0, "\" & DB(i).Flags, vbNullString))
                    found = (found + 1)
                End If
            End If
            
            res = False
            blnChecked = False
        Next i

        If (found = 0) Then
            arrReturn(0) = "No such user(s) found."
        Else
            Call SplitByLen(Mid$(tmpbuf, 1, Len(tmpbuf) - Len(", ")), 180, arrReturn(), "User(s) found: ", , ", ")
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.description & " in modCommandCode.SearchDatabase()."
End Sub

' Gets the number of non-null items in a list. Optionally passes the list formatted as a string.
Public Function GetListCount(ByRef aList() As String, Optional ByRef sList As String) As Integer
    Dim i       As Integer
    Dim iCount  As Integer
    sList = vbNullString
    
    iCount = 0
    For i = LBound(aList) To UBound(aList)
        If (LenB(Trim$(aList(i))) > 0) Then
            sList = StringFormat("{0}{1}, ", sList, aList(i))
            iCount = iCount + 1
        End If
    Next i
    sList = Left$(sList, Len(sList) - 2)
    GetListCount = iCount
End Function

' Outputs the specified list in response to the given command.
'   If more than iFailOnSize messages will be output (and the value is > -1), or there are no items, the function will fail.
'   If the prefix contains {%p} that will be replaced with the formatted position of the output
'     (x/y, where x is the number of the current message and y is the total number of messages)
Public Function ListRespond(ByRef oCommand As clsCommandObj, ByRef aList() As String, _
                            Optional ByVal sPrefix As String = vbNullString, Optional ByVal iFailOnSize As Integer = -1) As Boolean
    Dim sBuffer     As String
    Dim aBuffer()   As String
    Dim iCount      As Integer
    Dim i           As Integer
    
    ListRespond = False
    
    iCount = GetListCount(aList, sBuffer)
    If (iCount > 0) Then
        If (oCommand.IsLocal) Then
            oCommand.Respond StringFormat("{0}{1}", Replace(sPrefix, "{%p}", vbNullString), sBuffer)
            ListRespond = True
        Else
            Call SplitByLen(sBuffer, 200, aBuffer, sPrefix)
            If iFailOnSize = -1 Or (UBound(aBuffer) + 1) < iFailOnSize Then
                For i = LBound(aBuffer) To UBound(aBuffer)
                    If UBound(aBuffer) > 0 Then
                        oCommand.Respond Replace(aBuffer(i), "{%p}", StringFormat(" ({0}/{1})", (i + 1), UBound(aBuffer) + 1))
                    Else
                        oCommand.Respond Replace(aBuffer(i), "{%p}", vbNullString)
                    End If
                Next i
                ListRespond = True
            End If
        End If
    End If
End Function

