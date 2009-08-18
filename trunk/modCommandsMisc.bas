Attribute VB_Name = "modCommandsMisc"
Option Explicit
'This modules holds all other command code that i couldnt think of which catergory it fell info :P

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
        Command.Respond StringFormatA("Added mail for {0}.", temp.To)
        Call AddMail(temp)
    Else
        Command.Respond "Error: You must supply a recipient and a message."
    End If
End Sub

Public Sub OnCheckMail(Command As clsCommandObj)
    Dim Count As Integer
    
    Count = GetMailCount(IIf(Command.IsLocal, GetCurrentUsername, Command.Username))
    If (Count > 0) Then
        Command.Respond StringFormatA("You have {0} new message{1}. Type {2}inbox to retrieve {3}.", _
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
        
        Command.Respond StringFormatA("Execution {0} error #{1}: {2}", ErrType, .Error.Number, .Error.description)
        
        .Error.Clear
    End With
End Sub

Public Sub OnFlip(Command As clsCommandObj)
    Command.Respond IIf(((Rnd * 1000) Mod 2) = 0, "Heads.", "Tails.")
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
                Command.Respond StringFormatA("The statement {0}{1}{0} evaluates to: {2}.", Chr$(34), sStatement, sResult)
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
    Dim I        As Integer
    Dim X        As Integer
    Dim dbAccess As udtGetAccessResponse
    
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            If (LenB(Command.Username) = 0) Then Command.Username = BotVars.Username
        End If
        
        temp.From = Command.Username
        temp.Message = Command.Argument("Message")
        
        If (StrictIsNumeric(Command.Argument("Criteria"))) Then
            Rank = Val(Command.Argument("Criteria"))
            
            For I = 0 To UBound(DB)
                If (StrComp(DB(I).Type, "USER", vbTextCompare) = 0) Then
                    dbAccess = GetCumulativeAccess(DB(I).Username)
                    If (dbAccess.Rank = Rank) Then
                        temp.To = DB(I).Username
                        Call AddMail(temp)
                    End If
                End If
            Next I
            Command.Respond StringFormatA("Mass mailing to users with rank {0} complete.", Rank)
        Else
            Flags = Command.Argument("Criteria")
            For I = 0 To UBound(DB)
                If (StrComp(DB(I).Type, "USER", vbTextCompare) = 0) Then
                    dbAccess = GetCumulativeAccess(DB(I).Username)
                    For X = 1 To Len(Flags)
                        If (InStr(1, dbAccess.Flags, Mid$(Flags, X, 1), IIf(BotVars.CaseSensitiveFlags, vbBinaryCompare, vbTextCompare)) > 0) Then
                            temp.To = DB(I).Username
                            Call AddMail(temp)
                            Exit For
                        End If
                    Next X
                End If
            Next I
            Command.Respond StringFormatA("Mass mailing to users with any of the flags {0} complete.", Flags)
        End If
    Else
        Command.Respond StringFormatA("Format: {0}mmail <flag(s)> <message> OR {0}mmail <access> <message>", IIf(Command.IsLocal, "/", BotVars.Trigger))
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
            sFilePath = App.Path & "\" & sFilePath
        End If
        
        If (LenB(Dir$(sFilePath)) > 0) Then
            Command.Respond StringFormatA("Contents of file {0}:", Replace$(sFilePath, App.Path & "\", vbNullString, , 1, vbTextCompare))
            
            iFile = FreeFile
            Open sFilePath For Input As #iFile
                Do While (Not EOF(iFile))
                    Line Input #iFile, sLine
                    If (LenB(sLine) > 0) Then
                        iLineNumber = iLineNumber + 1
                        Command.Respond StringFormatA("Line {0}: {1}", iLineNumber, sLine)
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
    
    Command.Respond StringFormatA("Random number (0-{0}): {1}", maxValue, Number)
End Sub


