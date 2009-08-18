Attribute VB_Name = "modCommandsInfo"
Option Explicit
'This module will hold all of the 'Info' Commands
'Commands that return information, but have really no functionality

Public Function OnAbout(Command As clsCommandObj)
    Command.Respond ".: " & CVERSION & " :."
End Function

Public Function OnAccountInfo(Command As clsCommandObj) As Boolean
    If ((Not Command.IsLocal) Or Command.PublicOutput) Then
        PPL = True
        If ((BotVars.WhisperCmds Or Command.WasWhispered) And (Not Command.IsLocal)) Then
            PPLRespondTo = Command.Username
        End If
    End If
    RequestSystemKeys
End Function

' handle allseen command
Public Function OnAllSeen(Command As clsCommandObj) As Boolean
    Dim retVal As String
    Dim I      As Integer

    If (colLastSeen.Count = 0) Then
        retVal = "I have not seen anyone yet."
    Else
        retVal = "Last 15 users seen: "
        For I = 1 To colLastSeen.Count
            retVal = StringFormatA("{0}{1}{2}", _
                retVal, colLastSeen.Item(I), _
                IIf(I = colLastSeen.Count, vbNullString, ", "))
            If (I = 15) Then Exit For
        Next I
    End If
    Command.Respond retVal
End Function

Public Function OnDetail(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Dim sRetAdd As String, sRetMod As String
        Dim I As Integer
        
        For I = 0 To UBound(DB)
            With DB(I)
                If (StrComp(Command.Argument("Username"), .Username, vbTextCompare) = 0) Then
                    If ((Not .AddedBy = "%") And (LenB(.AddedBy) > 0)) Then
                        sRetAdd = StringFormatA("{0} was added by {1} on {2}.", .Username, .AddedBy, .AddedOn)
                    End If
                    
                    If ((Not .ModifiedBy = "%") And (LenB(.ModifiedBy) > 0)) Then
                        If ((Not .AddedOn = .ModifiedOn) Or (Not StrComp(.AddedBy, .ModifiedBy, vbTextCompare) = 0)) Then
                            sRetMod = StringFormatA(" The entry was last modified by {0} on {1}.", .ModifiedBy, .ModifiedOn)
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
                    
                    Exit Function
                End If
            End With
        Next I
        
        Command.Respond "That user was not found in the database."
    End If
End Function

Public Function OnFind(Command As clsCommandObj) As Boolean
    Dim dbAccess      As udtGetAccessResponse
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Not Command.IsValid) Then Exit Function
    
    If (Dir$(GetFilePath("users.txt")) = vbNullString) Then
       Command.Respond "No userlist available. Place a users.txt file in the bot's root directory."
       Exit Function
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
End Function

Public Function OnFindAttr(Command As clsCommandObj) As Boolean
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Command.IsValid) Then
        Call SearchDatabase(bufResponse(), , , , , , , Command.Argument("Attributes"))
    End If
    For Each strResponse In bufResponse
        Command.Respond CStr(strResponse)
    Next
End Function

Public Function OnFindGrp(Command As clsCommandObj) As Boolean
    Dim bufResponse() As String
    Dim strResponse   As Variant
    
    If (Command.IsValid) Then
        Call SearchDatabase(bufResponse(), , , Command.Argument("Group"))
    End If
    For Each strResponse In bufResponse
        Command.Respond CStr(strResponse)
    Next
End Function

Public Function OnHelp(Command As clsCommandObj)
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
            Command.Respond StringFormatA("[{0} (Aliases: {4})]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString, docs.RequirementsStringShort, docs.AliasString)
        If (docs.aliases.Count = 1) Then
            Command.Respond StringFormatA("[{0} (Alias: {4})]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString, docs.RequirementsStringShort, docs.AliasString)
        Else
            Command.Respond StringFormatA("[{0}]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString, docs.RequirementsStringShort)
        End If
    End If
    Set docs = Nothing
    
End Function

Public Function OnHelpAttr(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    
    If (Command.IsValid) Then
        tmpbuf = GetAllCommandsFor(, Command.Argument("Flags"))
        If (LenB(tmpbuf) > 0) Then
            Command.Respond "Commands available to specified flag(s): " & tmpbuf
        Else
            Command.Respond "No commands are available to the given flag(s)."
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnHelpAttr()."
End Function

Public Function OnHelpRank(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    
    If (Command.IsValid) Then
        tmpbuf = GetAllCommandsFor(Command.Argument("Rank"))
        If (LenB(tmpbuf) > 0) Then
            Command.Respond "Commands available to specified rank: " & tmpbuf
        Else
            Command.Respond "No commands are available to the given rank."
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnHelpRank()."
End Function

Public Function OnInfo(Command As clsCommandObj) As Boolean
    Dim UserIndex As Integer
    
    If (Command.IsValid) Then
        UserIndex = g_Channel.GetUserIndex(Command.Argument("Username"))
        
        If (UserIndex > 0) Then
            With g_Channel.Users(UserIndex)
                Command.Respond StringFormatA("User {0} is logged on using {1} with {2}a ping time of {3}ms.", _
                    .displaayname, ProductCodeToFullName(.game), _
                    IIf(.IsOperator, "ops, and ", vbNullString), .Ping)
            
                Command.Respond StringFormatA("He/she has been present in the channel for {0}.", ConvertTime(.TimeInChannel(), 1))
            End With
        Else
            Command.Respond "No such user is present."
        End If
    End If
End Function

Public Function OnInitPerf(Command As clsCommandObj) As Boolean
On Error GoTo ERROR_HANDLER
    
    Dim ModName As String
    Dim Name    As String
    Dim I       As Integer
    Dim strRet  As String
    Dim Script  As Module
    
    If (LenB(Command.Argument("Script")) > 0) Then
        Set Script = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Script Is Nothing) Then
            Command.Respond "Could not find the script specified."
        Else
            If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 1) Then
                Command.Respond StringFormatA("The Script {0} loaded in {1}ms.", _
                    Script.CodeObject.Script("Name"), _
                    Script.CodeObject.Script("InitPerf"))
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
            For I = 2 To frmChat.SControl.Modules.Count
                Set Script = frmChat.SControl.Modules(I)
                
                If (Not StrComp(Script.CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 0) Then
                    If (Command.IsLocal And Not Command.PublicOutput) Then
                        Command.Respond StringFormatA(" '{0}' loaded in {1}ms.", _
                            Script.CodeObject.Script("Name"), _
                            Script.CodeObject.Script("InitPerf"))
                    Else
                        strRet = StringFormatA("{0} '{1}' {2}ms{3}", _
                            strRet, _
                            Script.CodeObject.Script("Name"), _
                            Script.CodeObject.Script("InitPerf"), _
                            IIf(I = frmChat.SControl.Modules.Count, vbNullString, ","))
                    End If
                End If
            Next I
            
            If (LenB(strRet)) Then Command.Respond strRet
        Else
            Command.Respond "There are no scripts currently loaded."
        End If
    End If
        
    Exit Function
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnInitPerf()."
End Function

Public Function OnLastWhisper(Command As clsCommandObj) As Boolean
    If (LenB(LastWhisper) > 0) Then
        Command.Respond StringFormatA("The last whisper to this bot was from {0} at {1} on {2}.", _
            LastWhisper, _
            FormatDateTime(LastWhisperFromTime, vbLongTime), _
            FormatDateTime(LastWhisperFromTime, vbLongDate))
    Else
        Command.Respond "The bot has not been whispered since it logged on."
    End If
End Function

Public Function OnLocalIp(Command As clsCommandObj) As Boolean
    Command.Respond StringFormatA("{0} local IPv4 IP address is: {1}", IIf(Command.IsLocal, "Your", "My"), frmChat.sckBNet.LocalIP)
End Function

Public Function OnOwner(Command As clsCommandObj) As Boolean
    If (LenB(BotVars.BotOwner)) Then
        Command.Respond "This bot's owner is " & BotVars.BotOwner & "."
    Else
        Command.Respond "There is no owner currently set."
    End If
End Function

Public Function OnPing(Command As clsCommandObj) As Boolean
    Dim Latency As Long
    If (Command.IsValid) Then
        Latency = GetPing(Command.Argument("Username"))
        If (Latency >= -1) Then
            Command.Respond Command.Argument("Username") & "'s ping at login was " & Latency & "ms."
        Else
            Command.Respond "I can not see " & Command.Argument("Username") & " in the channel."
        End If
    Else
        Command.Respond "Please specify a user to ping."
    End If
End Function

Public Function OnPingMe(Command As clsCommandObj) As Boolean
    Dim Latency As Long
    If (Command.IsLocal) Then
        If (g_Online) Then
            Command.Respond "Your ping at login was " & GetPing(GetCurrentUsername) & "ms."
        Else
            Command.Respond "Error: You are not logged on."
        End If
    Else
        Latency = GetPing(Command.Username)
        If (Latency >= -1) Then
            Command.Respond "Your ping at login was " & Latency & "ms."
        Else
            Command.Respond "I can not see you in the channel."
        End If
    End If
End Function

Public Function OnProfile(Command As clsCommandObj)
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
            Call RequestProfile(Command.Argument("Username"))
            frmProfile.PrepareForProfile Command.Argument("Username"), False
        End If
    End If
End Function

Public Function OnScriptDetail(Command As clsCommandObj) As Boolean
On Error GoTo ERROR_HANDLER
    
    Dim Script As Module
    
    If (Command.IsValid) Then
        Set Script = modScripting.GetModuleByName(Command.Argument("Script"))
        If (Script Is Nothing) Then
            Command.Respond "Error: Could not find specified script."
        Else
            Dim Version  As String
            Dim VerTotal As Integer
            Dim Author   As String
            
            Version = StringFormatA("{0}.{1} Revision {2}", _
                Script.CodeObject.Script("Major"), _
                Script.CodeObject.Script("Minor"), _
                Script.CodeObject.Script("Revision"))
                
            VerTotal = Val(Script.CodeObject.Script("Major")) _
                     + Val(Script.CodeObject.Script("Minor")) _
                     + Val(Script.CodeObject.Script("Revision"))
                     
            Author = Script.CodeObject.Script("Author")
            
            If ((LenB(Author) = 0) And (VerTotal = 0)) Then
                Command.Respond StringFormatA("There is no additional information for the '{0}' script.", _
                    Script.CodeObject.Script("Name"))
            Else
                Command.Respond StringFormatA("{0}{1}{2}", _
                    Script.CodeObject.Script("Name"), _
                    IIf(VerTotal > 0, " v" & Version, vbNullString), _
                    IIf(LenB(Author) > 0, " by " & Author, vbNullString))
            End If
        End If
    End If
    Exit Function
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnScriptDetail()."
End Function

Public Function OnScripts(Command As clsCommandObj) As Boolean
On Error GoTo ERROR_HANDLER
    
    Dim retVal  As String
    Dim I       As Integer
    Dim Enabled As Boolean
    Dim Name    As String
    Dim Count   As Integer
    
    If (frmChat.SControl.Modules.Count > 1) Then
        For I = 2 To frmChat.SControl.Modules.Count
            Name = modScripting.GetScriptName(CStr(I))
            Enabled = Not (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), "False", vbTextCompare) = 0)
                
            retVal = StringFormatA("{0}{1}{2}{3}{4}", _
                retVal, _
                IIf(Enabled, vbNullString, "("), _
                Name, _
                IIf(Enabled, vbNullString, ")"), _
                IIf(I = frmChat.SControl.Modules.Count, vbNullString, ", "))
                
            Count = (Count + 1)
        Next I
        
        Command.Respond StringFormatA("Loaded Scripts ({0}): {1}", Count, retVal)
    Else
        Command.Respond "There are no scripts currently loaded."
    End If
    
    Exit Function
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.OnScripts()."
End Function


Public Function OnServer(Command As clsCommandObj)
    Dim RemoteHost   As String
    Dim RemoteHostIP As String
    
    RemoteHost = frmChat.sckBNet.RemoteHost
    RemoteHostIP = frmChat.sckBNet.RemoteHostIP
    
    If (StrComp(RemoteHost, RemoteHostIP, vbBinaryCompare) = 0) Then
        Command.Respond "I am currently connected to " & RemoteHostIP & "."
    Else
        Command.Respond "I am currently connected to " & RemoteHost & " (" & RemoteHostIP & ")."
    End If
End Function

Public Function OnTime(Command As clsCommandObj) As Boolean
    Command.Respond "The current time on this computer is " & Time & " on " & Format(Date, "MM-dd-yyyy") & "."
End Function

Public Function OnTrigger(Command As clsCommandObj) As Boolean
    If (LenB(BotVars.TriggerLong) = 1) Then
        Command.Respond StringFormatA("The bot's current trigger is {0} {1} {0} (Alt +0{2})", _
            Chr$(34), BotVars.TriggerLong, Asc(BotVars.TriggerLong))
    Else
        Command.Respond StringFormatA("The bot's current trigger is {0} {1} {0} (Length: {2})", _
          Chr$(34), BotVars.TriggerLong, Len(BotVars.TriggerLong))
    End If
End Function

Public Function OnUptime(Command As clsCommandObj) As Boolean
    Command.Respond StringFormatA("System uptime {0}, connection uptime {1}.", ConvertTime(GetUptimeMS), ConvertTime(uTicks))
End Function

Public Function OnWhere(Command As clsCommandObj)
    If (Command.IsLocal) Then
        Call frmChat.AddQ("/where " & Command.Args, PRIORITY.COMMAND_RESPONSE_MESSAGE, "(console)")
    End If

    Command.Respond StringFormatA("I am currently in channel {0} ({1} users present)", g_Channel.Name, g_Channel.Users.Count)
End Function

Public Function OnWhoAmI(Command As clsCommandObj) As Boolean
    Dim dbAccess As udtGetAccessResponse

    If (Command.IsLocal) Then
        Command.Respond "You are the bot console."
        
        If (g_Online) Then
            Call frmChat.AddQ("/whoami", PRIORITY.CONSOLE_MESSAGE)
        End If
    Else
        dbAccess = GetCumulativeAccess(Command.Username)
        If (dbAccess.Rank = 1000) Then
            Command.Respond "You are the bot owner, " & Command.Username & "."
        Else
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.Username & " holds rank " & dbAccess.Rank & _
                        " and flags " & dbAccess.Flags & "."
                Else
                    Command.Respond dbAccess.Username & " holds rank " & dbAccess.Rank & "."
                End If
            Else
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.Username & " has flags " & dbAccess.Flags & "."
                End If
            End If
        End If
    End If
End Function

Public Function OnWhoIs(Command As clsCommandObj)
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
End Function


Private Function GetAllCommandsFor(Optional Rank As Integer = -1, Optional Flags As String = vbNullString) As String
On Error GoTo ERROR_HANDLER
    
    Dim tmpbuf      As String
    Dim I           As Integer
    Dim xmldoc      As New DOMDocument60
    Dim Commands    As IXMLDOMNodeList
    Dim xpath       As String
    Dim lastCommand As String
    Dim thisCommand As String
    
    
    If (Dir$(GetFilePath("commands.xml")) = vbNullString) Then
        Command.Respond "Error: The XML database could not be found in the working directory."
        Exit Function
    End If

    If (LenB(Flags) > 0) Then

        For I = 1 To Len(Flags)
            xpath = StringFormatA("{0}'{1}'{2}", _
                xpath, _
                IIf(Mid$(Flags, I, 1) = "\", "\\", Mid$(Flags, I, 1)), _
                IIf(I = Len(Flags), vbNullString, " or "))
        Next I
        If (BotVars.CaseSensitiveFlags) Then
            xpath = StringFormatA("./command/access/flags/flag[text()={0}]", xpath)
        Else
            xpath = StringFormatA("./command/access/flags/flag[" & _
            "translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')={0}]", _
            LCase$(xpath))
        End If
    Else
        xpath = StringFormatA("./command/access/rank[number() <= {0}]", Rank)
    End If
    
    xmldoc.Load GetFilePath("commands.xml")
    
    Set Commands = xmldoc.documentElement.selectNodes(xpath)

    If (Commands.length > 0) Then
        For I = 0 To Commands.length - 1
            If (LenB(Flags) > 0) Then
                thisCommand = Commands(I).parentNode.parentNode.parentNode.Attributes.getNamedItem("name").Text
            Else
                thisCommand = Commands(I).parentNode.parentNode.Attributes.getNamedItem("name").Text
            End If
            
            If (StrComp(thisCommand, lastCommand, vbTextCompare) <> 0) Then
                tmpbuf = StringFormatA("{0}{1}, ", tmpbuf, thisCommand)
            End If
        
            lastCommand = thisCommand
        Next I
        
        tmpbuf = Left$(tmpbuf, Len(tmpbuf) - 2)
    End If
    GetAllCommandsFor = tmpbuf
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modCommandsInfo.GetAllCommandsFor()."
End Function
