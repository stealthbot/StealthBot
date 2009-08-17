Attribute VB_Name = "modCommandsInfo"
Option Explicit
'This module will hold all of the 'Info' Commands
'Commands that return information, but have really no functionality

Public Function OnAbout(Command As clsCommandObj)
    Command.Respond ".: " & CVERSION & " :."
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
    If (docs Is Nothing) Then
        Command.Respond "Sorry, but no related documentation could be found."
    Else
        If (docs.aliases.Count > 0) Then
            Command.Respond StringFormatA("[{0} (Alises: {4})]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString, docs.RequirementsString, docs.AliasString)
        Else
            Command.Respond StringFormatA("[{0}]: {1} (Syntax: {2}). {3}", _
            docs.Name, docs.description, docs.SyntaxString, docs.RequirementsString)
        End If
    End If
    Set docs = Nothing
    
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
        Latency = GetPing(Command.UserName)
        If (Latency >= -1) Then
            Command.Respond "Your ping at login was " & Latency & "ms."
        Else
            Command.Respond "I can not see you in the channel."
        End If
    End If
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

Public Function OnWhoAmI(Command As clsCommandObj) As Boolean
    Dim dbAccess As udtGetAccessResponse

    If (Command.IsLocal) Then
        Command.Respond "You are the bot console."
        
        If (g_Online) Then
            Call frmChat.AddQ("/whoami", PRIORITY.CONSOLE_MESSAGE)
        End If
    Else
        dbAccess = GetCumulativeAccess(Command.UserName)
        If (dbAccess.Rank = 1000) Then
            Command.Respond "You are the bot owner, " & Command.UserName & "."
        Else
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.UserName & " holds rank " & dbAccess.Rank & _
                        " and flags " & dbAccess.Flags & "."
                Else
                    Command.Respond dbAccess.UserName & " holds rank " & dbAccess.Rank & "."
                End If
            Else
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.UserName & " has flags " & dbAccess.Flags & "."
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
        
        If (LenB(dbAccess.UserName) > 0) Then
            If (dbAccess.Rank > 0) Then
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.UserName & " holds rank " & dbAccess.Rank & " and flags " & dbAccess.Flags & "."
                Else
                    Command.Respond dbAccess.UserName & " holds rank " & dbAccess.Rank & "."
                End If
            Else
                If (LenB(dbAccess.Flags) > 0) Then
                    Command.Respond dbAccess.UserName & " has flags " & dbAccess.Flags & "."
                End If
            End If
        Else
            Command.Respond "There was no such user found."
        End If
    End If
End Function


