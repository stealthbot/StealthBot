Attribute VB_Name = "modCommandsMisc"
Option Explicit
'This modules holds all other command code that i couldnt think of which catergory it fell info :P

Public Sub OnAddQuote(Command As clsCommandObj)
    If (g_Quotes Is Nothing) Then
        Set g_Quotes = New clsQuotesObj
    End If
    If (Command.IsValid) Then
        g_Quotes.Add Command.Argument("Quote")
        Command.Respond "Quote added!"
    Else
        Command.Respond "You must provide a quote to add."
    End If
End Sub

Public Sub OnBMail(Command As clsCommandObj)
    Dim temp       As udtMail
    Dim strArray() As String
    
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            If (LenB(Command.Username) = 0) Then Command.Username = BotVars.Username
        End If
        With temp
            .To = Command.Argument("Recipient")
            .From = Command.Username
            .Message = Command.Argument("Message")
        End With
        Command.Respond StringFormat("Added mail for {0}.", Trim(temp.To))
        Call AddMail(temp)
    Else
        Command.Respond "Error: You must supply a recipient and a message."
    End If
End Sub

Public Sub OnCancel(Command As clsCommandObj)
    If (VoteDuration > 0) Then
        Command.Respond Voting(BVT_VOTE_END, BVT_VOTE_CANCEL)
    Else
        Command.Respond "No vote in progress."
    End If
End Sub

Public Sub OnCheckMail(Command As clsCommandObj)
    Dim Count As Integer
    
    Count = GetMailCount(IIf(Command.IsLocal, GetCurrentUsername, Command.Username))
    If (Count > 0) Then
        Command.Respond StringFormat("You have {0} new message{1}. Type {2}inbox to retrieve {3}.", _
            Count, IIf(Count > 1, "s", vbNullString), IIf(Command.IsLocal, "/", BotVars.Trigger), _
            IIf(Count > 1, "them", "it"))
    Else
        Command.Respond "You have no mail."
    End If
End Sub

Public Sub OnExec(Command As clsCommandObj)
    On Error GoTo ERROR_HANDLER
    Dim ErrType As String

    If (Command.IsValid) Then frmChat.SControl.ExecuteStatement Command.Argument("Code")
    
    Exit Sub
    
ERROR_HANDLER:
    
    With frmChat.SControl
        ErrType = "runtime"
        
        If InStr(1, .Error.source, "compilation", vbBinaryCompare) > 0 Then ErrType = "parsing"
        
        Command.Respond StringFormat("Execution {0} error #{1}: {2}", ErrType, .Error.Number, .Error.description)
        
        .Error.Clear
    End With
End Sub

Public Sub OnFlip(Command As clsCommandObj)
    Command.Respond IIf(((Rnd * 1000) Mod 2) = 0, "Heads.", "Tails.")
End Sub

Public Sub OnGreet(Command As clsCommandObj)
    If (Command.IsValid) Then
        Select Case LCase$(Command.Argument("SubCommand"))
            Case "on":
                BotVars.UseGreet = True
                Command.Respond "Greet messages enabled."
                
            Case "off":
                BotVars.UseGreet = False
                Command.Respond "Greet messages disabled."
            
            Case "whisper":
                Select Case (LCase$(Command.Argument("Value")))
                    Case "on":
                        BotVars.WhisperGreet = True
                        Command.Respond "Greet messages will now be whispered."
                        
                    Case "off":
                        BotVars.WhisperGreet = False
                        Command.Respond "Gree messages will no longer be whispered."
                End Select
                
            Case "status":
                If (BotVars.UseGreet) Then
                    If (BotVars.WhisperGreet) Then
                        Command.Respond "Greet messages are currently enabled, and whispered."
                    Else
                        Command.Respond "Greet messages are currently enabled, and public."
                    End If
                Else
                    Command.Respond "Greet messages are currently disabled."
                End If
            
            Case "set":
                If (LenB(Command.Argument("Value")) > 0) Then
                    BotVars.GreetMsg = Command.Argument("Value")
                    Command.Respond "Greet message set." '1732 - 10/15/2009 - Hdx - Greet Set command will respond now.
                Else
                    Command.Respond "You must supply a greet message."
                End If
        End Select
        
        If LCase$(Command.Argument("SubCommand")) <> "status" Then
            Config.UseGreetMessage = BotVars.UseGreet
            Config.WhisperGreet = BotVars.WhisperGreet
            Config.GreetMessage = BotVars.GreetMsg
            Call Config.Save
        End If
    End If
End Sub

Public Sub OnIdle(Command As clsCommandObj)
    If (Command.IsValid) Then
        Select Case LCase$(Command.Argument("Enable"))
            Case "on", "true":
                Config.IdlesEnabled = True
                Command.Respond "Idles activated."
            
            Case "off", "false":
                Config.IdlesEnabled = False
                Command.Respond "Idles deactivated."
        End Select
        
        Call Config.Save
    End If
End Sub

Public Sub OnIdleTime(Command As clsCommandObj)
    Dim delay As Integer
    If (Command.IsValid) Then
        delay = Val(Command.Argument("Delay"))
        Config.IdleDelay = (delay * 2)
        Call Config.Save
        
        Command.Respond StringFormat("Idle wait time set to {0} minute{1}.", delay, IIf(delay > 1, "s", vbNullString))
    Else
        Command.Respond "You must supply a delay when setting the idle time."
    End If
End Sub

Public Sub OnIdleType(Command As clsCommandObj)
    If (Command.IsValid) Then
        Select Case (LCase$(Command.Argument("Type")))
            Case "msg", "message":
                Config.IdleType = "msg"
                Command.Respond "Idle type set to [ msg ]"
                
            Case "quote", "quotes":
                Config.IdleType = "quote"
                Command.Respond "Idle type set to [ quote ]"
                
            Case "uptime":
                Config.IdleType = "uptime"
                Command.Respond "Idle type set to [ uptime ]"
                
            Case "mp3":
                Config.IdleType = "mp3"
                Command.Respond "Idle type set to [ MP3 ]"
            
            Case Else:
                Command.Respond "Unknown Idle type, Type must be: Msg, Quote, Uptime, or MP3"
        End Select
        
        Call Config.Save
    Else
        Command.Respond "You must specify an idle type."
    End If
End Sub

Public Sub OnInbox(Command As clsCommandObj)
    Dim Msg      As udtMail
    Dim mcount   As Integer
    Dim Index    As Integer
    Dim dbAccess As udtGetAccessResponse
    
    If (Command.IsLocal) Then
        Command.Username = IIf(g_Online, GetCurrentUsername, BotVars.Username)
        dbAccess.Rank = 201
        dbAccess.Flags = "A"
    Else
        dbAccess = GetCumulativeAccess(Command.Username)
    End If
    
    If (GetMailCount(Command.Username) > 0) Then
        Do
            GetMailMessage Command.Username, Msg
            If (Len(RTrim$(Msg.To)) > 0) Then
                Command.Respond StringFormat("Message from: {0}: {1}", Trim$(Msg.From), Trim$(Msg.Message))
            End If
        Loop While (GetMailCount(Command.Username) > 0)
    Else
        If (dbAccess.Rank > 0) Then
            Command.Respond "You do not currently have any messages in your inbox."
        End If
    End If
    
    If (Not Command.IsLocal) Then
        Command.WasWhispered = True
    End If
End Sub

Public Sub OnMath(Command As clsCommandObj)
    ' This command will execute a specified mathematical statement using the
    ' restricted script control, SCRestricted, on frmChat.  The execution
    ' of any code through direct user-interaction can become quite error-prone
    ' and, as such, this command requires its own error handler.  The input
    ' of this command must also be properly sanitized to ensure that no
    ' harmful statements are inadvertently allowed to launch on the user's
    ' machine.
    
    On Error GoTo ERROR_HANDLER
    
    Dim sStatement As String
    Dim sResult    As String
    
    If (Command.IsValid) Then
        sStatement = Command.Argument("Expression")
        
        If (InStr(1, sStatement, "CreateObject", vbTextCompare)) Then
            Command.Respond "Evaluation error, CreateObject is restricted."
        Else
            With frmChat.SCRestricted
                .AllowUI = False
                .UseSafeSubset = True
            End With
            sStatement = Replace(sStatement, Chr$(34), vbNullString)
            sResult = frmChat.SCRestricted.Eval(sStatement)
            
            If (LenB(sResult) > 0) Then
                Command.Respond StringFormat("The statement {0}{1}{0} evaluates to: {2}.", Chr$(34), sStatement, sResult)
            Else
                Command.Respond "Evaluation error."
            End If
        End If
    End If
    Exit Sub
    
ERROR_HANDLER:
    Command.Respond "Evaluation error."
End Sub

Public Sub OnMMail(Command As clsCommandObj)
    Dim temp     As udtMail
    Dim Rank     As Long
    Dim Flags    As String
    Dim i        As Integer
    Dim x        As Integer
    Dim dbAccess As udtGetAccessResponse
    
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            If (LenB(Command.Username) = 0) Then Command.Username = BotVars.Username
        End If
        
        temp.From = Command.Username
        temp.Message = Command.Argument("Message")
        
        If (StrictIsNumeric(Command.Argument("Criteria"))) Then
            Rank = Val(Command.Argument("Criteria"))
            
            For i = 0 To UBound(DB)
                If (StrComp(DB(i).Type, "USER", vbTextCompare) = 0) Then
                    dbAccess = GetCumulativeAccess(DB(i).Username)
                    If (dbAccess.Rank = Rank) Then
                        temp.To = DB(i).Username
                        Call AddMail(temp)
                    End If
                End If
            Next i
            Command.Respond StringFormat("Mass mailing to users with rank {0} complete.", Rank)
        Else
            Flags = Command.Argument("Criteria")
            For i = 0 To UBound(DB)
                If (StrComp(DB(i).Type, "USER", vbTextCompare) = 0) Then
                    dbAccess = GetCumulativeAccess(DB(i).Username)
                    For x = 1 To Len(Flags)
                        If (InStr(1, dbAccess.Flags, Mid$(Flags, x, 1), IIf(BotVars.CaseSensitiveFlags, vbBinaryCompare, vbTextCompare)) > 0) Then
                            temp.To = DB(i).Username
                            Call AddMail(temp)
                            Exit For
                        End If
                    Next x
                End If
            Next i
            Command.Respond StringFormat("Mass mailing to users with any of the flags {0} complete.", Flags)
        End If
    Else
        Command.Respond StringFormat("Format: {0}mmail <flag(s)> <message> OR {0}mmail <access> <message>", IIf(Command.IsLocal, "/", BotVars.Trigger))
    End If
End Sub

Public Sub OnQuote(Command As clsCommandObj)
    Dim tmpQuote As String
    tmpQuote = "Quote: " & g_Quotes.GetRandomQuote
    
    If (Len(tmpQuote) < 8) Then
        Command.Respond "Error reading your quotes, or no quote file exists."
    Else
        Command.Respond tmpQuote
    End If
End Sub

Public Sub OnReadFile(Command As clsCommandObj)
    On Error GoTo ERROR_HANDLER
    Dim sFilePath   As String
    Dim iFile       As Long
    Dim sLine       As String
    Dim iLineNumber As Integer
    
    If (Command.IsValid) Then
        sFilePath = Command.Argument("File")
        If ((Not Mid$(sFilePath, 2, 2) = ":\") And (Not Mid$(sFilePath, 2, 2) = ":/")) Then
            sFilePath = StringFormat("{0}\{1}", CurDir$(), sFilePath)
        End If
        
        If (LenB(Dir$(sFilePath)) > 0) Then
            Command.Respond StringFormat("Contents of file {0}:", Replace$(sFilePath, StringFormat("{0}\", CurDir$()), vbNullString, , 1, vbTextCompare))
            
            iFile = FreeFile
            Open sFilePath For Input As #iFile
                Do While (Not EOF(iFile))
                    Line Input #iFile, sLine
                    If (LenB(sLine) > 0) Then
                        iLineNumber = iLineNumber + 1
                        Command.Respond StringFormat("Line {0}: {1}", iLineNumber, sLine)
                    End If
                Loop
            Close #iFile
            
            Command.Respond "End of File."
        Else
            Command.Respond "Error: The specified file could not be found."
        End If
    End If
    
    Exit Sub
 
ERROR_HANDLER:
    Command.ClearResponse
    Command.Respond "There was an error reading the specified file."
End Sub

Public Sub OnRoll(Command As clsCommandObj)
    Dim maxValue As Long
    Dim Number   As Long

    If (LenB(Command.Argument("Value")) > 0) Then
        maxValue = Abs(Val(Command.Argument("Value")))
    Else
        maxValue = 100
    End If
    
    Randomize
    Number = CLng(Rnd * maxValue)
    
    Command.Respond StringFormat("Random number (0-{0}): {1}", maxValue, Number)
End Sub

Public Sub OnSetIdle(Command As clsCommandObj)
    If (Command.IsValid) Then
        Config.IdleMessage = Command.Argument("Message")
        Call Config.Save
        
        Command.Respond "Idle message set."
    End If
End Sub

Public Sub OnTally(Command As clsCommandObj)
    If (VoteDuration > 0) Then
        Command.Respond Voting(BVT_VOTE_TALLY)
    Else
        Command.Respond "No vote is currently in progress."
    End If
End Sub

Public Sub OnVote(Command As clsCommandObj)
    Dim Duration As Integer
    If (Command.IsValid) Then
        If (VoteDuration = -1) Then
            Duration = Val(Command.Argument("Duration"))
            If ((Duration > 0) And (Duration < 32000)) Then
                VoteDuration = Duration
                If (Command.IsLocal) Then
                    With VoteInitiator
                        .Rank = 201
                        .Flags = "A"
                        .Username = "(Console)"
                    End With
                Else
                    VoteInitiator = GetCumulativeAccess(Command.Username)
                End If
                Command.Respond Voting(BVT_VOTE_START, BVT_VOTE_STD, vbNullString)
            Else
                Command.Respond "Vote durations must be between 1 and 32000"
            End If
        Else
            Command.Respond "A vote is currently in progress."
        End If
    Else
        Command.Respond "Please enter a number of seconds for your vote to last."
    End If
End Sub


