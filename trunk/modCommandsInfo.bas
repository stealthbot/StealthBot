Attribute VB_Name = "modCommandsInfo"
Option Explicit
'This module will hold all of the 'Info' Commands
'Commands that return information, but have really no functionality

Public Sub OnAbout(Command As clsCommandObj)
    Command.Respond ".: " & CVERSION & " :."
End Sub

Public Sub OnAccountInfo(Command As clsCommandObj)
    RequestSystemKeys reqUserCommand, Command
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
    If Not Command.IsValid Then
        Command.Respond "You must specify the name of an entry."
        Exit Sub
    End If
    
    Dim sRetAdd As String, sRetMod As String
    Dim i As Integer
        
    Dim oEntry As clsDBEntryObj
    Set oEntry = Database.GetEntry(Command.Argument("Username"))
        
    If oEntry Is Nothing Then
        Command.Respond "That entry was not found in the database."
        Exit Sub
    Else
        With oEntry
            ' Was it created by someone or something?
            If Len(.CreatedBy) > 0 Then
                sRetAdd = StringFormat("{0} was added by {1} on {2}.", .Name, .CreatedBy, .CreatedOn)
            End If
                
            ' Has it been modified?
            If Len(.ModifiedBy) > 0 Then
                If ((Not .CreatedOn = .ModifiedOn) Or (Not StrComp(.CreatedBy, .ModifiedBy, vbTextCompare) = 0)) Then
                    sRetMod = StringFormat("The entry was last modified by {0} on {1}.", .ModifiedBy, .ModifiedOn)
                Else
                    sRetMod = "The entry has not been modified since it was added."
                End If
            End If
                
            ' Make the response
            If ((Len(sRetAdd) > 0) Or (Len(sRetMod) > 0)) Then
                If Len(sRetAdd) > 0 Then
                    Command.Respond sRetAdd & Space(1) & sRetMod
                Else
                    Command.Respond sRetMod
                End If
            Else
                Command.Respond "No detailed information is available for that user."
            End If
        End With
    End If
End Sub

Public Sub OnFind(Command As clsCommandObj)
    Dim dbAccess      As udtUserAccess
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
            Dim iLowerRank As Integer
            Dim iUpperRank As Integer
            
            iLowerRank = Val(Command.Argument("Username/Rank"))
            iUpperRank = Command.Argument("UpperRank")
            
            If (iUpperRank = iLowerRank) Then
                Call SearchDatabase(bufResponse(), , , , , iLowerRank)
            ElseIf (iUpperRank > iLowerRank) Then
                Call SearchDatabase(bufResponse(), , , , , iLowerRank, iUpperRank)
            Else
                Call SearchDatabase(bufResponse(), , , , , iUpperRank, iLowerRank)
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
    Dim StateStr   As String
    Dim AliasStr1  As String
    Dim AliasStr2  As String
    Dim docs       As clsCommandDocObj
    
    strCommand = IIf(Command.IsValid, Command.Argument("Command"), "help")
    strScript = IIf(LenB(Command.Argument("ScriptOwner")) > 0, Command.Argument("ScriptOwner"), Chr$(0))
    
    Set docs = OpenCommand(strCommand, strScript, False)
    If (LenB(docs.Name) = 0) Then
        Command.Respond "Sorry, but no related documentation could be found."
    Else
        If (docs.Aliases.Count > 1) Then
            AliasStr1 = " (aliases: "
            AliasStr2 = ")"
        ElseIf (docs.Aliases.Count = 1) Then
            AliasStr1 = " (alias: "
            AliasStr2 = ")"
        End If
        If (Not docs.IsEnabled) Then StateStr = " (disabled)"

        Command.Respond StringFormat("[{0}{1}{2}{3}]{4}: {5} [syntax: {6}] {7}", _
                docs.Name, AliasStr1, docs.AliasString, AliasStr2, StateStr, _
                docs.Description, docs.SyntaxString(Command.IsLocal), docs.RequirementsStringShort)
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
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnHelpAttr()."
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
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnHelpRank()."
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
            
                Command.Respond StringFormat("He/she has been present in the channel for {0}.", ConvertTimeInterval(.TimeInChannel(), True))
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
        Command.Respond "Error: Scripts are globally disabled."
        Exit Sub
    End If
    
    Name = Command.Argument("Script")
    If (LenB(Name) > 0) Then
        Set Script = modScripting.GetModuleByName(Name, True)
        If (Script Is Nothing) Then
            Command.Respond StringFormat("Error: Could not find the script ""{0}"".", Name)
        Else
            Name = modScripting.GetScriptName(Script.Name)
            If (modScripting.IsScriptModuleEnabled(Script) = False) Then
                Command.Respond StringFormat("Error: Script ""{0}"" is currently disabled or failed to load.", Name)
            Else
                Command.Respond StringFormat("The script ""{0}"" loaded in {1}ms.", _
                        Name, Format$(GetScriptDictionary(Script)("InitPerf"), "#,##0"))
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
                Name = modScripting.GetScriptName(CStr(i))
                If (modScripting.IsScriptModuleEnabled(Script)) Then
                    If (Command.IsLocal And Not Command.PublicOutput) Then
                        Command.Respond StringFormat("    ""{0}"": {1}ms.", _
                            Name, Format$(GetScriptDictionary(Script)("InitPerf"), "#,##0"))
                    Else
                        strRet = StringFormat("{0} ""{1}"": {2}ms{3}", _
                            strRet, Name, _
                            Format$(GetScriptDictionary(Script)("InitPerf"), "#,##0"), _
                            IIf(i = frmChat.SControl.Modules.Count, ".", "; "))
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
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnInitPerf()."
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
        Command.Respond "This bot's owner is " & Config.BotOwner & "."
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
            Command.Respond StringFormat("{0}'s ping at logon was {1}ms.", Command.Argument("Username"), Latency)
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
            Command.Respond StringFormat("Your ping at logon was {0}ms.", GetPing(GetCurrentUsername))
        Else
            Command.Respond "Error: You are not logged on."
        End If
    Else
        Latency = GetPing(Command.Username)
        If (Latency >= -1) Then
            Command.Respond StringFormat("Your ping at logon was {0}ms.", Latency)
        Else
            Command.Respond "I can not see you in the channel."
        End If
    End If
End Sub

Public Sub OnProfile(Command As clsCommandObj)
    If (Command.IsValid) Then
        If ((Not Command.IsLocal) Or (Command.PublicOutput)) Then
            Call RequestProfile(Command.Argument("Username"), reqUserCommand, Command)
        Else
            frmProfile.PrepareForProfile Command.Argument("Username"), False
            Call RequestProfile(Command.Argument("Username"), reqUserInterface, Command)
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
    
    Dim Module As Module
    Dim Name   As String
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled."
        Exit Sub
    End If
    
    If (Command.IsValid) Then
        Name = Command.Argument("Script")
        Set Module = modScripting.GetModuleByName(Name, True)
        If (Module Is Nothing) Then
            Command.Respond StringFormat("Error: Could not find the script ""{0}"".", Name)
        Else
            Dim ScriptInfo  As Dictionary
            Dim Version     As String
            Dim VerTotal    As Double
            Dim Author      As String
            Dim Description As String
            Dim StateStr    As String
            Dim VerStr      As String
            Dim AuthorStr   As String
            Dim DescrStr    As String
            
            Set ScriptInfo = GetScriptDictionary(Module)

            Name = modScripting.GetScriptName(Module.Name)
            Version = StringFormat("{0}.{1}{2}", Val(ScriptInfo("Major")), Val(ScriptInfo("Minor")), _
                IIf(Val(ScriptInfo("Revision")) > 0, " Revision " & Val(ScriptInfo("Revision")), vbNullString))
            VerTotal = Val(ScriptInfo("Major")) + Val(ScriptInfo("Minor")) + Val(ScriptInfo("Revision"))
            Author = ScriptInfo("Author")
            Description = ScriptInfo("Description")

            If modScripting.IsScriptModuleEnabled(Module, True) = False Then
                StateStr = " (disabled)"
            ElseIf ScriptInfo("LoadError") Then
                StateStr = " (parsing error)"
            End If
            If VerTotal > 0 Then VerStr = " v" & Version
            If LenB(Author) > 0 Then AuthorStr = " by " & Author
            If LenB(Description) > 0 Then DescrStr = ": " & Description
            If Right$(DescrStr, 1) <> "." Then DescrStr = DescrStr & "."
            
            If ((LenB(Author) = 0) And (VerTotal = 0) And (LenB(Description) = 0)) Then
                Command.Respond StringFormat("There is no additional information for the script ""{0}""{1}.", Name, StateStr)
            Else
                Command.Respond StringFormat("""{0}""{1}{2}{3}{4}", Name, StateStr, VerStr, AuthorStr, DescrStr)
            End If
        End If
    End If
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnScriptDetail()."
End Sub

Public Sub OnScripts(Command As clsCommandObj)
On Error GoTo ERROR_HANDLER
    
    Dim retVal  As String
    Dim i       As Integer
    Dim Enabled As Boolean
    Dim Part    As String
    Dim Comma   As String
    Dim Name    As String
    Dim EnCount As Integer
    Dim Count   As Integer
    
    If modScripting.GetScriptSystemDisabled() Then
        Command.Respond "Error: Scripts are globally disabled."
        Exit Sub
    End If
    
    If (frmChat.SControl.Modules.Count > 1) Then
        Comma = ", "
        For i = 2 To frmChat.SControl.Modules.Count
            Name = modScripting.GetScriptName(CStr(i))
            Enabled = modScripting.IsScriptModuleEnabled(frmChat.SControl.Modules(i))
            Part = "{0}{1}{2}"
            If Not Enabled Then Part = "{0}{3}{1}{4}{2}"
            If (i = frmChat.SControl.Modules.Count) Then Comma = vbNullString

            retVal = StringFormat(Part, retVal, Name, Comma, "(", ")")
            Count = (Count + 1)
            If Enabled Then EnCount = EnCount + 1
        Next i
        
        Command.Respond StringFormat("Scripts ({0}/{1}): {2}", EnCount, Count, retVal)
    Else
        Command.Respond "There are no scripts currently loaded."
    End If
    
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.OnScripts()."
End Sub

Public Sub OnServer(Command As clsCommandObj)
    Dim RemoteHost   As String
    Dim RemoteHostIP As String
    Dim ProxyVia     As String
    Dim ProxyHost    As String
    Dim ProxyHostIP  As String
    Dim s1           As String
    Dim s2           As String
    Dim ProxyPort    As String
    
    If frmChat.sckBNet.State <> sckClosed Then
        If ProxyConnInfo(stBNCS).IsUsingProxy Then
            RemoteHost = ProxyConnInfo(stBNCS).RemoteHost
            RemoteHostIP = ProxyConnInfo(stBNCS).RemoteHostIP
            
            ProxyHost = frmChat.sckBNet.RemoteHost
            ProxyHostIP = frmChat.sckBNet.RemoteHostIP
            ProxyPort = CStr(frmChat.sckBNet.RemotePort)
            
            s1 = " ("
            s2 = ")"
            If (StrComp(ProxyHost, ProxyHostIP, vbBinaryCompare) = 0) Then
                ProxyHost = vbNullString
                s1 = vbNullString
                s2 = vbNullString
            End If
            ProxyVia = StringFormat(" via {0}{1}{2}:{3}{4}", ProxyHost, s1, ProxyHostIP, ProxyPort, s2)
        Else
            RemoteHost = frmChat.sckBNet.RemoteHost
            RemoteHostIP = frmChat.sckBNet.RemoteHostIP
            
            ProxyVia = vbNullString
        End If
        
        s1 = " ("
        s2 = ")"
        If (StrComp(RemoteHost, RemoteHostIP, vbBinaryCompare) = 0) Then
            RemoteHost = vbNullString
            s1 = vbNullString
            s2 = vbNullString
        End If
       
        Command.Respond StringFormat("I am currently connected to {0}{1}{2}{3}{4}.", RemoteHost, s1, RemoteHostIP, s2, ProxyVia)
    Else
        Command.Respond "I am not connected to Battle.net."
    End If
End Sub

Public Sub OnShitCheck(Command As clsCommandObj)
    Dim dbAccess As udtUserAccess
    Dim sName    As String
    
    sName = Command.Argument("Username")
    
    If (Command.IsValid) Then
        dbAccess = Database.GetUserAccess(sName)
        If Len(dbAccess.Username) > 0 Then sName = dbAccess.Username
        
        If InStr(1, dbAccess.Flags, "B", vbBinaryCompare) > 0 Then
            If InStr(1, dbAccess.Flags, "S", vbBinaryCompare) > 0 Then
                Command.Respond sName & " is on the bot's shitlist, but is also on the safelist and will not be banned."
            Else
                Command.Respond sName & " is on the bot's shitlist."
            End If
        Else
            Command.Respond sName & " is not on the bot's shitlist."
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

Public Sub OnTagCheck(Command As clsCommandObj)
    If Not (Command.IsValid) Then
        Command.Respond "You must specify a username to check."
        Exit Sub
    End If
    
    Dim i       As Integer
    Dim sName   As String
    Dim cGroups As Collection
    Dim oUser   As clsUserObj
    Dim oAccess As udtUserAccess
    
    sName = Command.Argument("username")
    Set cGroups = Database.GetDynamicGroups()
    
    ' Is the user in the channel?
    Set oUser = g_Channel.GetUser(sName, , False)
    If Len(oUser.Name) > 0 Then
        sName = oUser.Name
    End If
    
    ' Check each dynamic group entry.
    For i = 1 To cGroups.Count
        With cGroups.Item(i)
            ' Is this "group" a user?
            If StrComp(.EntryType, DB_TYPE_USER, vbBinaryCompare) = 0 Then
                ' Is there a wildcard in the name?
                If ((InStr(1, .Name, "*", vbBinaryCompare) > 0) Or (InStr(1, .Name, "?", vbBinaryCompare) > 0)) Then
                    ' Does it match the name we're looking for?
                    If LCase(PrepareCheck(sName)) Like LCase(PrepareCheck(.Name)) Then
                        ' Get the entry's access and check if its banned.
                        oAccess = Database.GetEntryAccess(cGroups.Item(i))
                        If InStr(1, oAccess.Flags, "B", vbBinaryCompare) > 0 Then
                            Command.Respond StringFormat("The user ""{0}"" is tagbanned under the entry: {1}", sName, oAccess.Username)
                            Exit Sub
                        End If
                    End If
                End If
                ' Is it a clan?
            ElseIf StrComp(.EntryType, DB_TYPE_CLAN, vbBinaryCompare) = 0 Then
                ' Do we have any additional info on this user?
                If Len(oUser.Name) > 0 Then
                    If Len(oUser.Clan) > 0 Then
                        If StrComp(oUser.Clan, .Name, vbTextCompare) = 0 Then
                            ' Get the entry's access and check if its banned.
                            oAccess = Database.GetEntryAccess(cGroups.Item(i))
                            If InStr(1, oAccess.Flags, "B", vbBinaryCompare) > 0 Then
                                Command.Respond StringFormat("The user ""{0}"" is tagbanned under the entry: {1}", sName, oAccess.Username)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    Command.Respond "That user does not match any tagbans."
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
    Command.Respond StringFormat("The current time on this computer is {0} on {1} ({2}).", Time, Format(Date, "MM-dd-yyyy"), GetTimeZoneName())
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
    If g_Online Then
        Command.Respond StringFormat("System uptime {0}; connection uptime {1}.", ConvertTimeInterval(modDateTime.GetTickCountMS()), ConvertTimeInterval(modBNCS.GetConnectionUptime()))
    Else
        Command.Respond StringFormat("System uptime {0}.", ConvertTimeInterval(modDateTime.GetTickCountMS()))
    End If
End Sub

Public Sub OnWhere(Command As clsCommandObj)
    If (Command.IsLocal) Then
        Call frmChat.AddQ("/where " & Command.Args, enuPriority.COMMAND_RESPONSE_MESSAGE, "(console)")
    End If

    Command.Respond StringFormat("I am currently in channel {0} ({1} users present)", g_Channel.Name, g_Channel.Users.Count)
End Sub

Public Sub OnWhoAmI(Command As clsCommandObj)
    Dim dbAccess As udtUserAccess

    If (Command.IsLocal) Then
        Command.Respond "You are the bot console."
        
        If (g_Online) Then
            Call frmChat.AddQ("/whoami", enuPriority.CONSOLE_MESSAGE)
        End If
    Else
        dbAccess = Database.GetUserAccess(Command.Username)
        
        If ((Len(Config.BotOwner) > 0) And (StrComp(dbAccess.Username, Config.BotOwner, vbTextCompare) = 0)) Then
            ' Special case case for bot owner.
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
    Dim dbAccess    As udtUserAccess
    Dim sAccessName As String
    Dim sType       As String
    
    If (Command.IsValid) Then
        If (Command.IsLocal And g_Online) Then
            Call frmChat.AddQ("/whois " & Command.Argument("Username"), enuPriority.CONSOLE_MESSAGE)
        End If

        ' Check if we know who this is.
        If Database.HasAccess(Command.Argument("Username"), sType) Then
            
            dbAccess = Database.GetAccess(Command.Argument("Username"), sType)
            sAccessName = Database.GetAccessNameString(dbAccess)
            
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond sAccessName & " holds rank " & dbAccess.Rank & " and flags " & dbAccess.Flags & "."
                Else
                    Command.Respond sAccessName & " holds rank " & dbAccess.Rank & "."
                End If
            Else
                If (Len(dbAccess.Flags) > 0) Then
                    Command.Respond sAccessName & " has flags " & dbAccess.Flags & "."
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
    Dim Flag        As String
    Dim CommandNode As IXMLDOMNode
    Dim NameNode    As IXMLDOMNode
    Dim OwnerNode   As IXMLDOMNode
    Dim Owner       As String
    
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
        xpath = StringFormat("./command[not(@enabled) or @enabled='1']/access/flags/flag[{0}]", xpath)
    Else
        xpath = StringFormat("./command[not(@enabled) or @enabled='1']/access/rank[number() <= {0}]", Rank)
    End If
    
    Set xmlDoc = commandDoc.XMLDocument
    
    Set commands = xmlDoc.documentElement.selectNodes(xpath)

    If (commands.Length > 0) Then
        For i = 0 To commands.Length - 1
            If (LenB(Flags) > 0) Then
                Set CommandNode = commands(i).parentNode.parentNode.parentNode
            Else
                Set CommandNode = commands(i).parentNode.parentNode
            End If
            Set NameNode = CommandNode.Attributes.getNamedItem("name")
            If (Not NameNode Is Nothing) Then
                thisCommand = NameNode.Text

                If (StrComp(thisCommand, lastCommand, vbTextCompare) <> 0) Then
                    Owner = vbNullString
                    Set OwnerNode = CommandNode.Attributes.getNamedItem("owner")
                    If (Not OwnerNode Is Nothing) Then
                        Owner = OwnerNode.Text
                    End If
                    If LenB(Owner) > 0 Then
                        If modScripting.IsScriptEnabled(Owner) Then
                            tmpbuf = StringFormat("{0}{1}, ", tmpbuf, thisCommand)
                        End If
                    Else
                        tmpbuf = StringFormat("{0}{1}, ", tmpbuf, thisCommand)
                    End If
                End If

                lastCommand = thisCommand
            End If
        Next i

        If (LenB(tmpbuf) > 0) Then
            tmpbuf = Left$(tmpbuf, Len(tmpbuf) - 2)
        End If
    End If
    GetAllCommandsFor = tmpbuf
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandsInfo.GetAllCommandsFor()."
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

Private Sub SearchDatabase(ByRef arrReturn() As String, Optional sName As String = vbNullString, _
    Optional ByVal sMatch As String = vbNullString, Optional sGroup As String = vbNullString, _
        Optional dbType As String = vbNullString, Optional iLowerBound As Integer = -1, _
            Optional iUpperBound As Integer = -1, Optional sFlags As String = vbNullString)
    
    On Error GoTo ERROR_HANDLER
    
    Dim i       As Integer
    Dim iFound  As Integer
    Dim sTemp   As String
    Dim sTypeID As String
    
    Dim dbAccess As udtUserAccess
    
    ReDim arrReturn(0)
    
    If (LenB(sName) > 0) Then
        dbAccess = Database.GetAccess(sName, dbType)
        sTypeID = IIf(StrComp(dbType, DB_TYPE_USER, vbBinaryCompare) = 0, vbNullString, " (" & dbType & ")")
        
        If (dbAccess.Rank > 0) Then
            sTemp = "Found user " & dbAccess.Username & sTypeID & ", who holds rank " & dbAccess.Rank & _
                IIf(Len(dbAccess.Flags) > 0, " and flags " & dbAccess.Flags, vbNullString) & "."
        ElseIf (LenB(dbAccess.Flags) > 0) Then
            sTemp = "Found user " & dbAccess.Username & sTypeID & ", with flags " & dbAccess.Flags & "."
        Else
            sTemp = "No such user(s) found."
        End If
    Else
        For i = 1 To Database.Entries.Count
            Dim res        As Boolean
            Dim blnChecked As Boolean
        
            With Database.Entries.Item(i)
                If (LenB(sMatch) > 0) Then
                    If (Left$(sMatch, 1) = "!") Then
                        res = (Not (LCase$(PrepareCheck(.Name)) Like (LCase$(Mid$(sMatch, 2)))))
                    Else
                        res = (LCase$(PrepareCheck(.Name)) Like (LCase$(sMatch)))
                    End If
                    blnChecked = True
                End If
                
                If (LenB(sGroup) > 0) Then
                    If .IsInGroup(sGroup) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If

                If (LenB(dbType) > 0) Then
                    If (StrComp(.EntryType, dbType, vbBinaryCompare) = 0) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If ((iLowerBound >= 0) And (iUpperBound >= 0)) Then
                    If ((.Rank >= iLowerBound) And (.Rank <= iUpperBound)) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                ElseIf (iLowerBound >= 0) Then
                    If (.Rank = iLowerBound) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If (LenB(sFlags) > 0) Then
                    If .HasAnyFlag(sFlags) Then
                        res = IIf(blnChecked, res, True)
                    Else
                        res = False
                    End If
                    blnChecked = True
                End If
                
                If (res = True) Then
                    sTemp = sTemp & .ToString()
                    sTemp = StringFormat("{0}{1}{2}, ", sTemp, _
                        IIf(.Rank > 0, "\" & .Rank, vbNullString), _
                        IIf(LenB(.Flags) > 0, "\" & .Flags, vbNullString))
                    iFound = (iFound + 1)
                End If
            End With
            
            res = False
            blnChecked = False
        Next i

        If (iFound = 0) Then
            arrReturn(0) = "No such user(s) found."
        Else
            Call SplitByLen(Mid$(sTemp, 1, Len(sTemp) - Len(", ")), 180, arrReturn(), "User(s) found: ", , ", ")
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modCommandCode.SearchDatabase()."
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
    
    ' Trim the extra separator off the end of the list.
    If iCount > 0 And Len(sList) > 1 Then
        sList = Left$(sList, Len(sList) - 2)
    End If
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

