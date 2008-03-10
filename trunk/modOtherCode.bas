Attribute VB_Name = "modOtherCode"
Option Explicit

Public Type COMMAND_DATA
    Name         As String
    Params       As String
    local        As Boolean
    PublicOutput As Boolean
End Type

'Read/WriteIni code thanks to ickis
Public Sub WriteINI(ByVal wiSection$, ByVal wiKey As String, ByVal wiValue As String, Optional ByVal wiFile As String = "x")
    If (StrComp(wiFile, "x", vbBinaryCompare) = 0) Then
        wiFile = GetConfigFilePath()
    Else
        If (InStr(1, wiFile$, "\", vbBinaryCompare) = 0) Then
            wiFile = GetFilePath(wiFile)
        End If
    End If
    
    WritePrivateProfileString wiSection, wiKey, _
        wiValue, wiFile
End Sub

Public Function ReadCFG$(ByVal riSection$, ByVal riKey$)
    Dim sRiBuffer As String
    Dim sRiValue  As String
    Dim sRiLong   As String
    Dim riFile    As String
    
    riFile = GetConfigFilePath()
    
    If (Dir(riFile) <> vbNullString) Then
        sRiBuffer = String(255, vbNull)
        
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr$(1), _
            sRiBuffer, 255, riFile)
        
        If (Left$(sRiBuffer, 1) <> Chr$(1)) Then
            sRiValue = Left$(sRiBuffer, sRiLong)
            
            ReadCFG = sRiValue
        End If
    Else
        ReadCFG = ""
    End If
End Function

Public Function ReadINI$(ByVal riSection$, ByVal riKey$, ByVal riFile$)
    Dim sRiBuffer As String
    Dim sRiValue  As String
    Dim sRiLong   As String
    
    If (InStr(1, riFile$, "\", vbBinaryCompare) = 0) Then
        riFile$ = GetFilePath(riFile)
    End If
    
    If (Dir(riFile$) <> vbNullString) Then
        sRiBuffer = String(255, vbNull)
        
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr$(1), _
            sRiBuffer, 255, riFile)
        
        If (Left$(sRiBuffer, 1) <> Chr(1)) Then
            sRiValue = Left$(sRiBuffer, sRiLong)
            ReadINI = sRiValue
        End If
    Else
        ReadINI = ""
    End If
End Function

Public Function GetTimeStamp() As String
    Select Case (BotVars.TSSetting)
        Case 0
            GetTimeStamp = _
                " [" & Format(Time, "HH:MM:SS AM/PM") & "] "
            
        Case 1
            GetTimeStamp = _
                " [" & Format(Time, "HH:MM:SS") & "] "
        
        Case 2
            GetTimeStamp = _
                " [" & Format(Time, "HH:MM:SS") & "." & _
                    Right$("000" & GetCurrentMS, 3) & "] "
        
        Case Else
            GetTimeStamp = vbNullString
    End Select
End Function

'// Converts a millisecond or second time value to humanspeak.. modified to support BNet's Time
'// Logged report.
'// Updated 2/11/05 to support timeGetSystemTime() unsigned longs, which in VB are doubles after conversion
Public Function ConvertTime(ByVal dblMS As Double, Optional seconds As Byte) As String
    Dim dblSeconds As Double
    Dim dblDays    As Double
    Dim dblHours   As Double
    Dim dblMins    As Double
    Dim strSeconds As String
    Dim strDays    As String
    
    If (seconds = 0) Then
        dblSeconds = (dblMS / 1000)
    Else
        dblSeconds = dblMS
    End If
    
    dblDays = Int(dblSeconds / 86400)
    dblSeconds = dblSeconds Mod 86400
    dblHours = Int(dblSeconds / 3600)
    dblSeconds = dblSeconds Mod 3600
    dblMins = Int(dblSeconds / 60)
    dblSeconds = (dblSeconds Mod 60)
    
    If (dblSeconds <> 1) Then
        strSeconds = "s"
    End If
    
    If (dblDays <> 1) Then
        strDays = "s"
    End If
    
    ConvertTime = dblDays & " day" & strDays & ", " & dblHours & " hours, " & _
        dblMins & " minutes and " & dblSeconds & " second" & strSeconds
End Function

Public Function GetVerByte(Product As String, Optional ByVal UseHardcode As Integer) As Long
    Dim Key As String ' ...
    
    Key = GetProductKey(Product)
    
    If ((ReadCFG("Override", Key & "VerByte") = vbNullString) Or _
        (UseHardcode = 1)) Then
        
        Select Case StrReverse(Product)
            Case "W2BN": GetVerByte = &H4F
            Case "STAR": GetVerByte = &HCF
            Case "SEXP": GetVerByte = &HCF
            Case "D2DV": GetVerByte = &HB
            Case "D2XP": GetVerByte = &HB
            Case "W3XP": GetVerByte = &H14
            Case "WAR3": GetVerByte = &H14
        End Select
    Else
        GetVerByte = _
            CLng(val("&H" & ReadCFG("Override", Key & "VerByte")))
    End If
    
End Function

Public Function GetGamePath(ByVal Client As String) As String
    Dim Key As String
    
    Key = GetProductKey(Client)
    
    If (LenB(ReadCFG("Override", Key & "Hashes")) > 0) Then
        GetGamePath = ReadCFG("Override", Key & "Hashes")
        
        If (Right$(GetGamePath, 1) <> "\") Then
            GetGamePath = GetGamePath & "\"
        End If
    Else
        Select Case (StrReverse$(UCase$(Client)))
            Case "W2BN": GetGamePath = App.Path & "\W2BN\"
            Case "STAR": GetGamePath = App.Path & "\STAR\"
            Case "SEXP": GetGamePath = App.Path & "\STAR\"
            Case "D2DV": GetGamePath = App.Path & "\D2DV\"
            Case "D2XP": GetGamePath = App.Path & "\D2XP\"
            Case "W3XP": GetGamePath = App.Path & "\WAR3\"
            Case "WAR3": GetGamePath = App.Path & "\WAR3\"
            
            Case Else
                frmChat.AddChat RTBColors.ErrorMessageText, _
                    "Warning: Invalid game client in GetGamePath()"
        End Select
    End If
End Function

Function MKL(Value As Long) As String
    Dim result As String * 4
    
    Call CopyMemory(ByVal result, Value, 4)
    
    MKL = result
End Function

Function MKI(Value As Integer) As String
    Dim result As String * 2
    
    Call CopyMemory(ByVal result, Value, 2)
    
    MKI = result
End Function

Public Function CheckPath(ByVal sPath As String) As Long
    If (LenB(Dir$(sPath)) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[HASHES] " & _
            Mid$(sPath, InStrRev(sPath, "\") + 1) & " is missing."
            
        CheckPath = 1
    End If
End Function

Public Function Ban(ByVal Inpt As String, SpeakerAccess As Integer, Optional Kick As Integer) As String
    Static LastBan      As String
    
    Dim Username        As String
    Dim CleanedUsername As String
    Dim i               As Integer
    
    If (LenB(Inpt) > 0) Then
        If (Kick > 2) Then
            LastBan = vbNullString
            
            Exit Function
        End If
        
        'If (g_Queue.Count) Then
        '    For i = 1 To g_Queue.Count
        '        With g_Queue.Peek
        '            If (Left$(.Message, 5) = "/ban ") Then
        '                If (StrComp(LastBan & Space(1), LCase$(Mid$(.Message, 6, Len(LastBan) + 1)), _
        '                    vbTextCompare) = 0) Then
        '
        '                    Exit Function
        '                End If
        '            End If
        '        End With
        '    Next i
        'End If
        
        If ((MyFlags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
            If (InStr(1, Inpt, Space$(1), vbBinaryCompare) <> 0) Then
                Username = LCase$(Left$(Inpt, InStr(1, Inpt, Space(1), _
                    vbBinaryCompare) - 1))
            Else
                Username = LCase$(Inpt)
            End If
            
            If (LenB(Username) > 0) Then
                LastBan = LCase$(Username)
                
                CleanedUsername = StripInvalidNameChars(Username)

                If (SpeakerAccess < 999) Then
                    If ((GetSafelist(CleanedUsername)) Or (GetSafelist(Username))) Then
                        Ban = "That user is safelisted."
                        
                        Exit Function
                    End If
                End If
                
                If (GetAccess(Username).access >= SpeakerAccess) Then
                    Ban = "You do not have enough access to do that."
                    
                    Exit Function
                End If
                
                If (Kick = 0) Then
                    Call frmChat.AddQ("/ban " & Inpt)
                Else
                    Call frmChat.AddQ("/kick " & Inpt)
                End If
            End If
        Else
            Ban = "The bot does not have ops."
        End If
    End If
End Function

' This function created in response to http://www.stealthbot.net/forum/index.php?showtopic=20550
Public Function StripInvalidNameChars(ByVal Username As String) As String
    Dim Allowed(14) As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim thisChar    As Integer
    Dim NewUsername As String
    Dim ThisCharOK  As Boolean
    
    If (LenB(Username) > 0) Then
        NewUsername = Username
        
        Allowed(0) = Asc("`")
        Allowed(1) = Asc("[")
        Allowed(2) = Asc("]")
        Allowed(3) = Asc("{")
        Allowed(4) = Asc("}")
        Allowed(5) = Asc("_")
        Allowed(6) = Asc("-")
        Allowed(7) = Asc("@")
        Allowed(8) = Asc("^")
        Allowed(9) = Asc(".")
        Allowed(10) = Asc("+")
        Allowed(11) = Asc("=")
        Allowed(12) = Asc("~")
        Allowed(13) = Asc("|")
        Allowed(14) = Asc("*")
        
        For i = 1 To Len(Username)
            thisChar = Asc(Mid$(Username, i, 1))
            
            ThisCharOK = False
            
            If (Not (IsAlpha(thisChar))) Then
                If (Not (IsNumber(thisChar))) Then
                    For j = 0 To UBound(Allowed)
                        If (thisChar = Allowed(j)) Then
                            ThisCharOK = True
                        End If
                    Next j
                    
                    If (Not (ThisCharOK)) Then
                        NewUsername = Replace(NewUsername, Chr(thisChar), _
                            vbNullString)
                    End If
                End If
            End If
        Next i
        
        StripInvalidNameChars = NewUsername
    End If
End Function

Public Function StripRealm(ByVal Username As String) As String
    If (InStr(1, Username, "@", vbBinaryCompare) > 0) Then
        ' ...
        Username = Replace(Username, "@USWest", vbNullString, 1)
        Username = Replace(Username, "@USEast", vbNullString, 1)
        Username = Replace(Username, "@Asia", vbNullString, 1)
        Username = Replace(Username, "@Euruope", vbNullString, 1)
        
        ' ...
        Username = Replace(Username, "@Lordaeron", vbNullString, 1)
        Username = Replace(Username, "@Azeroth", vbNullString, 1)
        Username = Replace(Username, "@Kalimdor", vbNullString, 1)
        Username = Replace(Username, "@Northrend", vbNullString, 1)
    End If
    
    StripRealm = Username
End Function

Public Sub bnetSend(ByVal Message As String, Optional ByVal tag As String = vbNullString)
    If (frmChat.sckBNet.State = 7) Then
        With PBuffer
            If (frmChat.mnuUTF8.Checked = False) Then
                .InsertNTString Message, UTF8
            Else
                .InsertNTString Message
            End If

            .SendPacket &HE
        End With
        
        If (tag = "request_receipt") Then
            ' ...
            g_request_receipt = True
        
            With PBuffer
                .SendPacket &H65
            End With
        End If
    End If

    ' ...
    If (bFlood = False) Then
        On Error Resume Next
        
        frmChat.SControl.Run "Event_MessageSent", Message, tag
    End If
End Sub

Public Sub APISend(ByRef s As String) '// faster API-based sending for EFP

    Dim i As Long
    
    i = Len(s) + 5
    
    Call Send(frmChat.sckBNet.SocketHandle, "ÿ" & "" & Chr(i) & _
        Chr(0) & s & Chr(0), i, 0)
End Sub

Public Function Voting(ByVal Mode1 As Byte, Optional Mode2 As Byte, Optional Username As String) As String
    Static Voted()  As String
    Static VotesYes As Integer
    Static VotesNo  As Integer
    Static VoteMode As Byte
    Static Target   As String
        
    Dim i           As Integer
    
    Select Case (Mode1)
        Case BVT_VOTE_ADD
            For i = LBound(Voted) To UBound(Voted)
                If (StrComp(Voted(i), LCase$(Username), _
                    vbTextCompare) = 0) Then
                    
                    Exit Function
                End If
            Next i
            
            Select Case (Mode2)
                Case BVT_VOTE_ADDYES
                    VotesYes = VotesYes + 1
                Case BVT_VOTE_ADDNO
                    VotesNo = VotesNo + 1
            End Select
            
            Voted(UBound(Voted)) = _
                LCase$(Username)
            
            ReDim Preserve Voted(0 To UBound(Voted) + 1)
        
        Case BVT_VOTE_START
            VotesYes = 0
            VotesNo = 0
            
            ReDim Voted(0)
            
            VoteMode = Mode2
            Target = Username
            Voting = "Vote started. Type YES or NO to vote. Your vote will " & _
                "only be counted once."
        
        Case BVT_VOTE_END
            If (Mode2 = BVT_VOTE_CANCEL) Then
                Voting = "Vote cancelled. Final results: [" & VotesYes & "] YES, " & _
                    "[" & VotesNo & "] NO. "
            Else
                Select Case (VoteMode)
                    Case BVT_VOTE_STD
                        Voting = "Final results: [" & VotesYes & "] YES, [" & VotesNo & "] NO. "
                        
                        If VotesYes > VotesNo Then
                            Voting = Voting & "YES wins, with " & _
                                Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
                        ElseIf (VotesYes < VotesNo) Then
                            Voting = Voting & "NO wins, with " & _
                                Format(VotesNo / (VotesYes + VotesNo), "percent") & " of the vote."
                        Else
                            Voting = Voting & "The vote was a draw."
                        End If
                        
                    Case BVT_VOTE_BAN
                        If (VotesYes > VotesNo) Then
                            Voting = Ban(Target & " Banned by vote", VoteInitiator.access)
                        Else
                            Voting = "Ban vote failed."
                        End If
                        
                    Case BVT_VOTE_KICK
                        If (VotesYes > VotesNo) Then
                            Voting = Ban(Target & " Kicked by vote", VoteInitiator.access, 1)
                        Else
                            Voting = "Kick vote failed."
                        End If
                End Select
            End If
            
            VoteDuration = -1
            VotesYes = 0
            VotesNo = 0
            VoteMode = 0
            Target = vbNullString
            
            ReDim Voted(0)
        
        Case BVT_VOTE_TALLY
            Voting = "Current results: [" & VotesYes & "] YES, [" & VotesNo & "] NO; " & _
                VoteDuration & " seconds remain."
            
            If (VotesYes > VotesNo) Then
                Voting = Voting & " YES leads, with " & _
                    Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
            ElseIf (VotesYes < VotesNo) Then
                Voting = Voting & _
                    " NO leads, with " & Format(VotesNo / (VotesYes + VotesNo), "percent") & _
                        " of the vote."
            Else
                Voting = Voting & " The vote is a draw."
            End If
    End Select
End Function

Public Function GetAccess(ByVal Username As String, Optional dbType As String = _
    vbNullString) As udtGetAccessResponse
    
    Dim i   As Integer ' ...
    Dim bln As Boolean ' ...

    'If (Left$(Username, 1) = "*") Then
    '    Username = Mid$(Username, 2)
    'End If

    For i = LBound(DB) To UBound(DB)
        If (StrComp(DB(i).Username, Username, vbTextCompare) = 0) Then
            If (Len(dbType)) Then
                If (StrComp(DB(i).Type, dbType, vbBinaryCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
                
            If (bln = True) Then
                With GetAccess
                    .Username = DB(i).Username
                    .access = DB(i).access
                    .flags = DB(i).flags
                    .AddedBy = DB(i).AddedBy
                    .AddedOn = DB(i).AddedOn
                    .ModifiedBy = DB(i).ModifiedBy
                    .ModifiedOn = DB(i).ModifiedOn
                    .Type = DB(i).Type
                    .Groups = DB(i).Groups
                    .BanMessage = DB(i).BanMessage
                End With
                
                Exit Function
            End If
        End If
        
        bln = False
    Next i

    GetAccess.access = -1
End Function

Public Function GetCumulativeAccess(ByVal Username As String, Optional dbType As String = _
    vbNullString) As udtGetAccessResponse
    
    On Error GoTo ERROR_HANDLER

    Dim gAcc    As udtGetAccessResponse ' ...
    
    Dim i       As Integer ' ...
    Dim k       As Integer ' ...
    Dim j       As Integer ' ...
    Dim found   As Boolean ' ...
    Dim dbIndex As Integer ' ...
    Dim dbCount As Integer ' ...
    Dim splt()  As String  ' ...
    Dim bln     As Boolean ' ...
    
    ' default index to negative one to
    ' indicate that no matching users have
    ' been found
    dbIndex = -1

    ' ...
    If (DB(LBound(DB)).Username <> vbNullString) Then
        ' ...
        For i = LBound(DB) To UBound(DB)
            ' ...
            If (StrComp(Username, DB(i).Username, vbTextCompare) = 0) Then
                With GetCumulativeAccess
                    .Username = DB(i).Username & _
                        IIf(((DB(i).Type <> "%") And (StrComp(DB(i).Type, "USER", vbTextCompare) <> 0)), _
                            " (" & LCase$(DB(i).Type) & ")", vbNullString)
                    .access = DB(i).access
                    .flags = DB(i).flags
                    .AddedBy = DB(i).AddedBy
                    .AddedOn = DB(i).AddedOn
                    .ModifiedBy = DB(i).ModifiedBy
                    .ModifiedOn = DB(i).ModifiedOn
                    .Type = IIf(((DB(i).Type <> "%") And (DB(i).Type <> vbNullString)), _
                        DB(i).Type, "USER")
                    .Groups = DB(i).Groups
                    .BanMessage = DB(i).BanMessage
                End With
                
                If ((Len(DB(i).Groups) > 0) And (DB(i).Groups <> "%")) Then
                    ' ...
                    If (InStr(1, DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                        ' ...
                        splt() = Split(DB(i).Groups, ",")
                    Else
                        ' ...
                        ReDim Preserve splt(0)
                        
                        ' ...
                        splt(0) = DB(i).Groups
                    End If
                    
                    ' ...
                    For j = 0 To UBound(splt)
                        ' ...
                        gAcc = GetCumulativeGroupAccess(splt(j))
                    
                        ' ...
                        If (GetCumulativeAccess.access < gAcc.access) Then
                            ' ...
                            GetCumulativeAccess.access = gAcc.access
                            
                            ' ...
                            bln = True
                        End If
                        
                        ' ...
                        For k = 1 To Len(gAcc.flags)
                            ' ...
                            If (InStr(1, GetCumulativeAccess.flags, Mid$(gAcc.flags, k, 1), _
                                vbBinaryCompare) = 0) Then
                                
                                ' ...
                                GetCumulativeAccess.flags = GetCumulativeAccess.flags & _
                                    Mid$(gAcc.flags, k, 1)
                                    
                                ' ...
                                bln = True
                            End If
                        Next k
                        
                        ' ...
                        If ((GetCumulativeAccess.BanMessage = vbNullString) Or _
                            (GetCumulativeAccess.BanMessage = "%")) Then
                            
                            ' ...
                            GetCumulativeAccess.BanMessage = gAcc.BanMessage
                            
                            ' ...
                            bln = True
                        End If
                        
                        ' ...
                        If (bln) Then
                            ' ...
                            If (dbCount = 0) Then
                                GetCumulativeAccess.Username = GetCumulativeAccess.Username & _
                                    IIf((i + 1), Space(1), vbNullString) & "["
                            End If
                                    
                            ' ...
                            GetCumulativeAccess.Username = GetCumulativeAccess.Username & gAcc.Username & _
                                IIf(((gAcc.Type <> "%") And (StrComp(gAcc.Type, "USER", vbTextCompare) <> 0)), _
                                    " (" & LCase$(gAcc.Type) & ")", vbNullString) & ", "
                                
                            ' ...
                            dbCount = (dbCount + 1)
                        End If
                        
                        ' ...
                        bln = False
                    Next j
                End If
                
                dbIndex = i
    
                Exit For
            End If
        Next i
    
        ' ...
        If (InStr(1, GetCumulativeAccess.flags, "I", vbBinaryCompare) = 0) Then
            ' ...
            If ((InStr(1, Username, "*", vbBinaryCompare) = 0) And _
                (InStr(1, Username, "?", vbBinaryCompare) = 0) And _
                (GetCumulativeAccess.Type <> "GAME") And _
                (GetCumulativeAccess.Type <> "CLAN") And _
                (GetCumulativeAccess.Type <> "GROUP")) Then
                
                ' ...
                For i = LBound(DB) To UBound(DB)
                    Dim doCheck As Boolean ' ...
                    
                    If (i <> dbIndex) Then
                        ' default type to user
                        DB(i).Type = IIf(((DB(i).Type <> "%") And (DB(i).Type <> vbNullString)), _
                            DB(i).Type, "USER")
                    
                        If (StrComp(DB(i).Type, "USER", vbTextCompare) = 0) Then
                            ' ...
                            If ((LCase$(PrepareCheck(Username))) Like _
                                (LCase$(PrepareCheck(DB(i).Username)))) Then
                                
                                ' ...
                                doCheck = True
                            End If
                        ElseIf (StrComp(DB(i).Type, "GAME", vbTextCompare) = 0) Then
                            ' ...
                            For j = 1 To colUsersInChannel.Count
                                If (StrComp(Username, colUsersInChannel.Item(j).Username, _
                                        vbTextCompare) = 0) Then
                                     
                                    If (StrComp(DB(i).Username, colUsersInChannel.Item(j).Product, _
                                        vbTextCompare) = 0) Then
                                
                                        ' ...
                                        doCheck = True
                                    End If
                                    
                                    Exit For
                                End If
                            Next j
                        ElseIf (StrComp(DB(i).Type, "CLAN", vbTextCompare) = 0) Then
                            ' ...
                            For j = 1 To colUsersInChannel.Count
                                If (StrComp(Username, colUsersInChannel.Item(j).Username, _
                                        vbTextCompare) = 0) Then
                                     
                                    If (StrComp(DB(i).Username, colUsersInChannel.Item(j).Clan, _
                                        vbTextCompare) = 0) Then
                                
                                        ' ...
                                        doCheck = True
                                    End If
                                    
                                    Exit For
                                End If
                            Next j
                        End If
                        
                        ' ...
                        If (doCheck = True) Then
                            Dim tmp As udtDatabase ' ...
                            
                            ' ...
                            tmp = DB(i)
            
                            ' ...
                            If ((Len(tmp.Groups) > 0) And (tmp.Groups <> "%")) Then
                                ' ...
                                If (InStr(1, tmp.Groups, ",", vbBinaryCompare) <> 0) Then
                                    ' ...
                                    splt() = Split(tmp.Groups, ",")
                                Else
                                    ' ...
                                    ReDim Preserve splt(0)
                                    
                                    ' ...
                                    splt(0) = tmp.Groups
                                End If
                                
                                ' ...
                                For j = 0 To UBound(splt)
                                    ' ...
                                    gAcc = GetCumulativeGroupAccess(splt(j))
                                
                                    ' ...
                                    If (tmp.access < gAcc.access) Then
                                        tmp.access = gAcc.access
                                    End If
                                    
                                    ' ...
                                    For k = 1 To Len(gAcc.flags)
                                        ' ...
                                        If (InStr(1, tmp.flags, Mid$(gAcc.flags, k, 1), _
                                            vbBinaryCompare) = 0) Then
                                            
                                            ' ...
                                            tmp.flags = tmp.flags & _
                                                Mid$(gAcc.flags, k, 1)
                                        End If
                                    Next k
                                    
                                    ' ...
                                    If ((tmp.BanMessage = vbNullString) Or _
                                        (tmp.BanMessage = "%")) Then
                                        
                                        ' ...
                                        tmp.BanMessage = gAcc.BanMessage
                                    End If
                                Next j
                            End If
    
                            ' ...
                            If (GetCumulativeAccess.access < tmp.access) Then
                                ' ...
                                GetCumulativeAccess.access = tmp.access
                                
                                ' ...
                                bln = True
                            End If
                            
                            ' ...
                            For j = 1 To Len(tmp.flags)
                                ' ...
                                If (InStr(1, GetCumulativeAccess.flags, Mid$(tmp.flags, j, 1), _
                                    vbBinaryCompare) = 0) Then
                                    
                                    ' ...
                                    GetCumulativeAccess.flags = GetCumulativeAccess.flags & _
                                        Mid$(tmp.flags, j, 1)
                                    
                                    ' ...
                                    bln = True
                                End If
                            Next j
                            
                            ' ...
                            If ((GetCumulativeAccess.BanMessage = vbNullString) Or _
                                (GetCumulativeAccess.BanMessage = "%")) Then
                                
                                ' ...
                                GetCumulativeAccess.BanMessage = tmp.BanMessage
                                
                                ' ...
                                bln = True
                            End If
       
                            If (bln) Then
                                ' ...
                                If (dbCount = 0) Then
                                    ' ...
                                    GetCumulativeAccess.Username = GetCumulativeAccess.Username & _
                                        IIf((dbIndex + 1), Space(1), vbNullString) & "["
                                End If
                            
                                ' ...
                                GetCumulativeAccess.Username = GetCumulativeAccess.Username & tmp.Username & _
                                    IIf(((tmp.Type <> "%") And (StrComp(tmp.Type, "USER", vbTextCompare) <> 0)), _
                                        " (" & LCase$(tmp.Type) & ")", vbNullString) & ", "
                                    
                                ' ...
                                dbCount = (dbCount + 1)
                            End If
                        End If
                    End If
                    
                    ' ...
                    bln = False
                    doCheck = False
                Next i
            End If
        End If
        
        If (dbCount = 0) Then
            If (dbIndex = -1) Then
                With GetCumulativeAccess
                    .Username = vbNullString
                    .access = -1
                    .flags = vbNullString
                End With
            End If
        Else
            ' ...
            GetCumulativeAccess.Username = Left$(GetCumulativeAccess.Username, _
                Len(GetCumulativeAccess.Username) - 2) & "]"
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in " & _
        "GetCumulativeAccess().")

    Exit Function
End Function

' ...
Private Function GetCumulativeGroupAccess(ByVal Group As String) As udtGetAccessResponse
    Dim gAcc   As udtGetAccessResponse ' ...
    Dim splt() As String               ' ...
    
    ' ...
    gAcc = GetAccess(Group, "GROUP")
    
    ' ...
    If ((Len(gAcc.Groups) > 0) And (gAcc.Groups <> "%")) Then
        Dim recAcc As udtGetAccessResponse ' ...
    
        ' ...
        If (InStr(1, gAcc.Groups, ",", vbBinaryCompare) <> 0) Then
            Dim i As Integer ' ...
            Dim j As Integer ' ...
        
            ' ...
            splt() = Split(gAcc.Groups, ",")
            
            ' ...
            For i = 0 To UBound(splt)
                ' ...
                recAcc = GetCumulativeGroupAccess(splt(i))
                    
                ' ...
                If (gAcc.access < recAcc.access) Then
                    gAcc.access = recAcc.access
                End If
                
                ' ...
                For j = 1 To Len(recAcc.flags)
                    ' ...
                    If (InStr(1, gAcc.flags, Mid$(recAcc.flags, j, 1), _
                        vbBinaryCompare) = 0) Then
                        
                        ' ...
                        gAcc.flags = gAcc.flags & _
                            Mid$(recAcc.flags, j, 1)
                    End If
                Next j
                
                ' ...
                If ((gAcc.BanMessage = vbNullString) Or _
                    (gAcc.BanMessage = "%")) Then
                    
                    ' ...
                    gAcc.BanMessage = recAcc.BanMessage
                End If
            Next i
        Else
            ' ...
            recAcc = GetCumulativeGroupAccess(gAcc.Groups)
        
            ' ...
            If (gAcc.access < recAcc.access) Then
                gAcc.access = recAcc.access
            End If
            
            ' ...
            For j = 1 To Len(recAcc.flags)
                ' ...
                If (InStr(1, gAcc.flags, Mid$(recAcc.flags, j, 1), _
                    vbBinaryCompare) = 0) Then
                    
                    ' ...
                    gAcc.flags = gAcc.flags & _
                        Mid$(recAcc.flags, j, 1)
                End If
            Next j
            
            ' ...
            If ((gAcc.BanMessage = vbNullString) Or _
                (gAcc.BanMessage = "%")) Then
                
                ' ...
                gAcc.BanMessage = recAcc.BanMessage
            End If
        End If
    End If
    
    ' ...
    GetCumulativeGroupAccess = gAcc
End Function

' ...
Public Function CheckGroup(ByVal Group As String, ByVal Check As String) As Boolean
    Dim gAcc   As udtGetAccessResponse ' ...
    Dim splt() As String               ' ...
    
    ' ...
    gAcc = GetAccess(Group, "GROUP")
    
    ' ...
    If ((Len(gAcc.Groups) > 0) And (gAcc.Groups <> "%")) Then
        Dim recAcc As Boolean ' ...
    
        ' ...
        If (InStr(1, gAcc.Groups, ",", vbBinaryCompare) <> 0) Then
            Dim i As Integer ' ...
            Dim j As Integer ' ...
        
            ' ...
            splt() = Split(gAcc.Groups, ",")
            
            ' ...
            For i = 0 To UBound(splt)
                If (StrComp(splt(i), Check, vbTextCompare) = 0) Then
                    CheckGroup = True
                    
                    Exit Function
                Else
                    ' ...
                    recAcc = CheckGroup(splt(i), Check)
                
                    If (recAcc) Then
                        CheckGroup = True
                    
                        Exit Function
                    End If
                End If
            Next i
        Else
            If (StrComp(gAcc.Groups, Check, vbTextCompare) = 0) Then
                CheckGroup = True
                
                Exit Function
            Else
                ' ...
                recAcc = CheckGroup(gAcc.Groups, Check)
            
                If (recAcc) Then
                    CheckGroup = True
                
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' ...
    CheckGroup = False
End Function

Public Sub RequestSystemKeys()
    AwaitingSystemKeys = 1
    
    With PBuffer
        .InsertDWord &H1
        .InsertDWord &H4
        .InsertDWord GetTickCount()
        .InsertNTString CurrentUsername
            
        .InsertNTString "System\Account Created"
        .InsertNTString "System\Last Logon"
        .InsertNTString "System\Last Logoff"
        .InsertNTString "System\Time Logged"
        
        .SendPacket &H26
    End With
End Sub

'// parses a system time and returns in the format:
'//     mm/dd/yy, hh:mm:ss
'//
Public Function SystemTimeToString(ByRef sT As SYSTEMTIME) As String
    Dim buf As String

    With sT
        buf = buf & .wMonth & "/"
        buf = buf & .wDay & "/"
        buf = buf & .wYear & ", "
        buf = buf & IIf(.wHour > 9, .wHour, "0" & .wHour) & ":"
        buf = buf & IIf(.wMinute > 9, .wMinute, "0" & .wMinute) & ":"
        buf = buf & IIf(.wSecond > 9, .wSecond, "0" & .wSecond)
    End With
    
    SystemTimeToString = buf
End Function

Public Function GetCurrentMS() As String
    Dim sT As SYSTEMTIME
    GetLocalTime sT
    
    GetCurrentMS = Right$("000" & sT.wMilliseconds, 3)
End Function

Public Function ZeroOffset(ByVal lInpt As Long, ByVal lDigits As Long) As String
    Dim sOut As String
    
    sOut = Hex(lInpt)
    ZeroOffset = Right$(String(lDigits, "0") & sOut, lDigits)
End Function

Public Function ZeroOffsetEx(ByVal lInpt As Long, ByVal lDigits As Long) As String
    ZeroOffsetEx = Right$(String(lDigits, "0") & lInpt, lDigits)
End Function

Public Function GetSmallIcon(ByVal sProduct As String, ByVal flags As Long) As Long
    Dim i As Long
    
    If ((flags And USER_BLIZZREP) = USER_BLIZZREP) Then 'Flags = 1: blizzard rep
        i = ICBLIZZ
    ElseIf ((flags And USER_SYSOP) = USER_SYSOP) Then 'Flags = 8: battle.net sysop
        i = ICSYSOP
    ElseIf (flags And USER_CHANNELOP&) = USER_CHANNELOP& Then 'op
        i = ICGAVEL
    ElseIf (flags And USER_SQUELCHED) = USER_SQUELCHED Then 'squelched
        i = ICSQUELCH
    ElseIf (g_ThisIconCode <> -1) Then
        If (sProduct = "W3XP") Then
            i = g_ThisIconCode + ICON_START_W3XP + _
                IIf(g_ThisIconCode + ICON_START_W3XP = ICSCSW, 1, 0)
        Else
            i = g_ThisIconCode + ICON_START_WAR3
        End If
    Else
        Select Case (UCase$(sProduct))
            Case Is = "STAR": i = ICSTAR
            Case Is = "SEXP": i = ICSEXP
            Case Is = "D2DV": i = ICD2DV
            Case Is = "D2XP": i = ICD2XP
            Case Is = "W2BN": i = ICW2BN
            Case Is = "CHAT": i = ICCHAT
            Case Is = "DRTL": i = ICDIABLO
            Case Is = "DSHR": i = ICDIABLOSW
            Case Is = "JSTR": i = ICJSTR
            Case Is = "SSHR": i = ICSCSW
                
            '*** Special icons for WCG added 6/24/07 ***
            Case Is = "WCRF": i = IC_WCRF
            Case Is = "WCPL": i = IC_WCPL
            Case Is = "WCGO": i = IC_WCGO
            Case Is = "WCSI": i = IC_WCSI
            Case Is = "WCBR": i = IC_WCBR
            Case Is = "WCPG": i = IC_WCPG
                
            '*** Special icons for PGTour ***
            Case Is = "__A+": i = IC_PGT_A + 1
            Case Is = "___A": i = IC_PGT_A
            Case Is = "__A-": i = IC_PGT_A - 1
            Case Is = "__B+": i = IC_PGT_B + 1
            Case Is = "___B": i = IC_PGT_B
            Case Is = "__B-": i = IC_PGT_B - 1
            Case Is = "__C+": i = IC_PGT_C + 1
            Case Is = "___C": i = IC_PGT_C
            Case Is = "__C-": i = IC_PGT_C - 1
            Case Is = "__D+": i = IC_PGT_D + 1
            Case Is = "___D": i = IC_PGT_D
            Case Is = "__D-": i = IC_PGT_D - 1
                
            Case Else: i = ICUNKNOWN
        End Select
    End If
    
    GetSmallIcon = i
End Function

Public Sub AddName(ByVal Username As String, ByVal Product As String, ByVal flags As Long, ByVal Ping As Long, Optional Clan As String, Optional ForcePosition As Integer)
    Dim i          As Integer
    Dim LagIcon    As Integer
    Dim isPriority As Integer
    Dim IsSelf     As Boolean
    
    If (StrComp(Username, CurrentUsername, vbTextCompare) = 0) Then
        MyFlags = flags
        
        SharedScriptSupport.BotFlags = _
            MyFlags
        
        IsSelf = True
    End If
    
    If (checkChannel(Username) > 0) Then
        Exit Sub
    End If
    
    Select Case (Ping)
        Case 0 To 199
            LagIcon = LAG_1
        Case 200 To 299
            LagIcon = LAG_2
        Case 300 To 399
            LagIcon = LAG_3
        Case 400 To 499
            LagIcon = LAG_4
        Case 500 To 599
            LagIcon = LAG_5
        Case Is >= 600 Or -1
            LagIcon = LAG_6
        Case Else
            LagIcon = ICUNKNOWN
    End Select
    
    If ((flags And USER_NOUDP) = USER_NOUDP) Then
        LagIcon = LAG_PLUG
    End If
    
    isPriority = (frmChat.lvChannel.ListItems.Count + 1)
    
    i = GetSmallIcon(Product, flags)
    
    'Special Cases
    'If i = ICSQUELCH Then
    '    'Debug.Print "Returned a SQUELCH icon"
    '    If ForcePosition > 0 Then isPriority = ForcePosition
    '
    If (((flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or _
        ((flags And USER_CHANNELOP&) = USER_CHANNELOP&)) Then
        
        If (ForcePosition = 0) Then
            isPriority = 1
        Else
            isPriority = ForcePosition
        End If
    Else
        If (ForcePosition > 0) Then
            isPriority = ForcePosition
        End If
    End If
    
    If (i > frmChat.imlIcons.ListImages.Count) Then
        i = frmChat.imlIcons.ListImages.Count
    End If
        
    With frmChat.lvChannel.ListItems
        .Add isPriority, , Username, , i
        
        .Item(isPriority).ListSubItems.Add , , , LagIcon
        
        If (Not (BotVars.NoColoring)) Then
            .Item(isPriority).ForeColor = _
                GetNameColor(flags, 0, IsSelf)
        End If
        
        g_ThisIconCode = -1
    End With
End Sub


Public Function CheckBlock(ByVal Username As String) As Boolean
    Dim s As String
    Dim i As Integer
    
    If (Dir$(GetFilePath("filters.ini")) <> vbNullString) Then
        s = ReadINI("BlockList", "Total", "filters.ini")
        
        If (StrictIsNumeric(s)) Then
            i = s
        Else
            Exit Function
        End If
        
        Username = PrepareCheck(Username)
        
        For i = 0 To i
            s = ReadINI("BlockList", "Filter" & i, "filters.ini")
            
            If (Username Like PrepareCheck(s)) Then
                CheckBlock = True
                
                Exit Function
            End If
        Next i
    End If
End Function

Public Function CheckMsg(ByVal Msg As String, Optional ByVal Username As String, _
    Optional ByVal Ping As Long) As Boolean
    
    Dim i As Integer ' ...
    
    Msg = LCase$(Msg)
    
    For i = 0 To UBound(gFilters)
        If (Len(gFilters(i)) > 1) Then
            If (InStr(1, gFilters(i), "%", vbBinaryCompare) > 0) Then
                If (InStr(1, Msg, LCase$(DoReplacements(gFilters(i), _
                    Username, Ping))) > 0) Then
                    
                    CheckMsg = True
                    
                    Exit Function
                End If
            Else
                If (InStr(1, Msg, LCase$(gFilters(i)), vbBinaryCompare) > 0) Then
                    CheckMsg = True
                    
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Sub UpdateProfile()
    Dim s As String
    
    s = MediaPlayer.TrackName
    
    If (s = vbNullString) Then
        Exit Sub
    End If
    
    SetProfile "", ":[ ProfileAmp ]:" & vbCrLf & "WinAmp is currently playing: " & _
        vbCrLf & s & vbCrLf & "Last updated " & Time & ", " & Format(Date, "d-MM-yyyy") & _
            vbCrLf & CVERSION & " - http://www.stealthbot.net"
End Sub

Public Function FlashWindow() As Boolean
    Dim pfwi As FLASHWINFO
    
    'Me.WindowState = vbMinimized
    
    With pfwi
        .cbSize = 20
        .dwFlags = FLASHW_ALL Or 12
        .dwTimeout = 0
        .hWnd = frmChat.hWnd
        .uCount = 0
    End With
    
    FlashWindow = FlashWindowEx(pfwi)
End Function

Public Sub ReadyINet()
    frmChat.INet.Cancel
End Sub

Public Function GetNewsURL() As String
    GetNewsURL = "http://www.stealthbot.net/getver3.php?vc=" & VERCODE
End Function

Public Function HTMLToRGBColor(ByVal s As String) As Long
    HTMLToRGBColor = RGB(val("&H" & Mid$(s, 1, 2)), val("&H" & Mid$(s, 3, 2)), _
        val("&H" & Mid$(s, 5, 2)))
End Function

Public Function StrictIsNumeric(ByVal sCheck As String) As Boolean
    Dim i As Long
    
    StrictIsNumeric = True
    
    If (Len(sCheck) > 0) Then
        For i = 1 To Len(sCheck)
            If (Not ((Asc(Mid$(sCheck, i, 1)) >= 48) And _
                     (Asc(Mid$(sCheck, i, 1)) <= 57))) Then
                
                StrictIsNumeric = False
                
                Exit Function
            End If
        Next i
    Else
        StrictIsNumeric = False
    End If
End Function


'---------------------------------------------------------------------------------------
' These two bad boys are cdkey writing/loading for frmSettings and frmManageKeys
'---------------------------------------------------------------------------------------
'
Public Sub LoadCDKeys(ByRef cboCDKey As ComboBox)
    Dim Count As Integer
    Dim sKey  As String
    
    Count = val(ReadCFG("StoredKeys", "Count"))
    
    If (Count) Then
        For Count = 1 To Count
            sKey = ReadCFG("StoredKeys", "Key" & Count)
            
            If (Len(sKey) > 0) Then
                cboCDKey.AddItem sKey
            End If
        Next Count
    End If
End Sub

Public Sub WriteCDKeys(ByRef cboCDKey As ComboBox)
    Dim i As Integer
    
    Call WriteINI("StoredKeys", "Count", cboCDKey.ListCount + 1)
    
    For i = 0 To cboCDKey.ListCount
        If (Len(cboCDKey.List(i)) > 0) Then
            WriteINI "StoredKeys", "Key" & (i + 1), cboCDKey.List(i)
        End If
    Next i
End Sub

Public Sub GetCountryData(ByRef CountryAbbrev As String, ByRef CountryName As String)
    Const LOCALE_USER_DEFAULT = &H400
    Const LOCALE_SABBREVCTRYNAME As Long = &H7 'abbreviated country name
    Const LOCALE_SENGCOUNTRY As Long = &H1002  'English name of country
    
    Dim sBuf As String
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, sBuf, Len(sBuf))
    CountryAbbrev = KillNull(sBuf)
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, sBuf, Len(sBuf))
    CountryName = KillNull(sBuf)
End Sub

Public Function GetTimeZoneBias() As Long
    Dim TZinfo As TIME_ZONE_INFORMATION
    Dim lngL   As Long
    
    lngL = GetTimeZoneInformation(TZinfo)

    Select Case (lngL)
        Case TIME_ZONE_ID_STANDARD
            GetTimeZoneBias = (TZinfo.Bias + TZinfo.StandardBias)
        Case TIME_ZONE_ID_DAYLIGHT
            GetTimeZoneBias = (TZinfo.Bias + TZinfo.DaylightBias)
        Case TIME_ZONE_ID_UNKNOWN
            Debug.Print "Bias unknown. Setting to standard."
            GetTimeZoneBias = TZinfo.Bias
        Case Else
            Debug.Print "Error in GetTimeZoneBias. Defaulting to 0x480."
            
            GetTimeZoneBias = &H480
    End Select
End Function


Public Sub SetNagelStatus(ByVal lSocketHandle As Long, ByVal bEnabled As Boolean)
    If (lSocketHandle > 0) Then
        If (bEnabled) Then
            Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_OFF, NAGLE_OPTLEN)
        Else
            Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_ON, NAGLE_OPTLEN)
        End If
    End If
End Sub

Public Sub EnableSO_KEEPALIVE(ByVal lSocketHandle As Long)
    Call SetSockOpt(lSocketHandle, IPPROTO_TCP, SO_KEEPALIVE, True, 4) 'thanks Eric
End Sub

Function MonitorExists() As Boolean
    MonitorExists = (Not (MonitorForm Is Nothing))
End Function

Sub InitMonitor()
    If (MonitorExists) Then
        frmChat.DeconstructMonitor
    End If
    
    Set MonitorForm = New frmMonitor
    
    Call frmMonitor.Hide
End Sub

Public Function ProductCodeToFullName(ByVal pCode As String) As String
    Select Case (pCode)
        Case "SEXP": ProductCodeToFullName = "Starcraft: Brood War"
        Case "STAR": ProductCodeToFullName = "Starcraft Original"
        Case "WAR3": ProductCodeToFullName = "Warcraft III"
        Case "W2BN": ProductCodeToFullName = "Warcraft II BNE"
        Case "D2DV": ProductCodeToFullName = "Diablo II"
        Case "D2XP": ProductCodeToFullName = "Diablo II: Lord of Destruction"
        Case "SSHR": ProductCodeToFullName = "Starcraft Shareware"
        Case "JSTR": ProductCodeToFullName = "Starcraft Japanese"
        Case "DRTL": ProductCodeToFullName = "Diablo Retail"
        Case "W3XP": ProductCodeToFullName = "Warcraft III: The Frozen Throne"
        Case "CHAT": ProductCodeToFullName = "a telnet connection"
        Case "DSHR": ProductCodeToFullName = "Diablo Shareware"
        Case "SSHR": ProductCodeToFullName = "Starcraft Shareware"
        Case Else:   ProductCodeToFullName = "an unknown or non-standard product"
    End Select
End Function

' Assumes that sIn has length >=1
Public Function PercentActualUppercase(ByVal sIn As String) As Double
    Dim UppercaseChars As Integer
    Dim i              As Integer
    
    sIn = Replace$(sIn, Space(1), vbNullString)
    
    If (Len(sIn) > 0) Then
        For i = 1 To Len(sIn)
            If (IsAlpha(Asc(Mid$(sIn, i, 1)))) Then
                If (IsUppercase(Asc(Mid$(sIn, i, 1)))) Then
                    UppercaseChars = (UppercaseChars + 1)
                End If
            End If
        Next i
    
        PercentActualUppercase = _
            CDbl(100 * (UppercaseChars / Len(sIn)))
    End If
End Function

Public Function MyUCase(ByVal sIn As String) As String
    Dim i           As Integer
    Dim CurrentByte As Byte

    If (LenB(sIn) > 0) Then
        For i = 1 To Len(sIn)
            CurrentByte = Asc(Mid$(sIn, i, 1))
            
            If (IsAlpha(CurrentByte)) Then
                If (Not (IsUppercase(CurrentByte))) Then
                    Mid$(sIn, i, 1) = _
                        Chr(CurrentByte - 32)
                End If
            End If
        Next i
    End If

    MyUCase = sIn
End Function

Public Function IsAlpha(ByVal bCharValue As Byte) As Boolean
    IsAlpha = ((bCharValue >= 65 And bCharValue <= 90) Or _
        (bCharValue >= 97 And bCharValue <= 122))
End Function

Public Function IsNumber(ByVal bCharValue As Byte) As Boolean
    IsNumber = ((bCharValue >= 48 And bCharValue <= 57))
End Function

Public Function IsUppercase(ByVal bCharValue As Byte) As Boolean
    IsUppercase = (bCharValue >= 65 And bCharValue <= 90)
End Function

Public Function VBHexToHTMLHex(ByVal sIn As String) As String
    sIn = Left$((sIn) & "000000", 6)
    
    VBHexToHTMLHex = Mid$(sIn, 5, 2) & Mid$(sIn, 3, 2) & _
        Mid$(sIn, 1, 2)
End Function

Public Sub GetW3LadderProfile(ByVal sPlayer As String, ByVal eType As enuWebProfileTypes)
    If (LenB(sPlayer) > 0) Then
        ShellExecute frmChat.hWnd, "Open", "http://www.battle.net/war3/ladder/" & _
            IIf(eType = W3XP, "w3xp", "war3") & "-player-profile.aspx?Gateway=" & _
                GetW3Realm(sPlayer) & "&PlayerName=" & NameWithoutRealm(sPlayer), 0&, 0&, 0&
    End If
End Sub

Public Sub DoLastSeen(ByVal Username As String)
    Dim i     As Integer
    Dim found As Boolean
    
    If (colLastSeen.Count > 0) Then
        For i = 1 To colLastSeen.Count
            If (StrComp(colLastSeen.Item(i), Username, _
                vbTextCompare) = 0) Then
                
                found = True
                
                Exit For
            End If
        Next i
    End If
    
    If (Not (found)) Then
        colLastSeen.Add Username
        
        If (colLastSeen.Count > 15) Then
            Call colLastSeen.Remove(1)
        End If
    End If
End Sub

Public Sub SetTitle(ByVal sTitle As String)
    frmChat.Caption = "[" & sTitle & "]" & " - " & CVERSION
End Sub

Public Function NameWithoutRealm(ByVal Username As String, Optional ByVal Strict As Byte = 0) As String
    If ((IsW3) And (Strict = 0)) Then
        NameWithoutRealm = Username
    Else
        If (InStr(1, Username, "@", vbBinaryCompare) > 0) Then
            NameWithoutRealm = Left$(Username, InStr(1, Username, "@") - 1)
        Else
            NameWithoutRealm = Username
        End If
    End If
End Function

Public Function GetW3Realm(Optional ByVal Username As String) As String
    If (LenB(Username) = 0) Then
        GetW3Realm = BotVars.Gateway
    Else
        If (InStr(1, Username, "@", vbBinaryCompare) > 0) Then
            GetW3Realm = Mid$(Username, InStr(1, Username, "@", _
                vbBinaryCompare) + 1)
        Else
            GetW3Realm = BotVars.Gateway
        End If
    End If
End Function

Public Function GetConfigFilePath() As String
    Static FilePath As String
    
    If (LenB(FilePath) = 0) Then
        If ((LenB(ConfigOverride) > 0)) Then
            FilePath = ConfigOverride
        Else
            FilePath = GetProfilePath()
            
            FilePath = FilePath & _
                IIf(Right$(FilePath, 1) = "\", "", "\") & "config.ini"
        End If
    End If
    
    If (InStr(1, FilePath, "\", vbBinaryCompare) = 0) Then
        FilePath = App.Path & "\" & FilePath
    End If
    
    GetConfigFilePath = FilePath
End Function

Public Function GetFilePath(ByVal filename As String) As String
    Dim s As String
    
    If (InStr(filename, "\") = 0) Then
        GetFilePath = GetProfilePath() & "\" & filename
        
        s = ReadCFG("FilePaths", filename)
        
        If (LenB(s) > 0) Then
            If (LenB(Dir$(s))) Then
                GetFilePath = s
            End If
        End If
    Else
        GetFilePath = filename
    End If
End Function

'Public Function OKToDoAutocompletion(ByRef sText As String, ByVal KeyAscii As Integer) As Boolean
'    If BotVars.NoAutocompletion Then
'        OKToDoAutocompletion = False
'    Else
'        If (InStr(sText, " ") = 0) And KeyAscii <> 32 Then       ' one word only
'            OKToDoAutocompletion = True
'        Else                                                            ' > 1 words
'            If StrComp(Left$(sText, 1), "/") = 0 And KeyAscii <> 32 Then ' left character is a /
'
'                If InStr(InStr(sText, " ") + 1, sText, " ") = 0 Then    ' only two words
'                    If StrComp(Left$(sText, 3), "/m ") = 0 Or _
'                        StrComp(Left$(sText, 3), "/w ") = 0 Then
'
'                            OKToDoAutocompletion = True
'                    Else
'                        OKToDoAutocompletion = False
'                    End If
'                Else                                                    ' more than two words
'                    OKToDoAutocompletion = False
'                End If
'            Else
'                OKToDoAutocompletion = False
'            End If
'        End If
'    End If
'End Function

' PROFILE FILE FORMAT
' profilename | profile path
' assumes colProfiles is instantiated!
'Public Sub LoadProfileList(ByRef cbo As ComboBox)
'    Dim f As Integer
'    Dim sInput As String
'
'    If LenB(ReadIfNI("Main", "ConfigVersion")) Then
'        If LenB(Dir$(App.Path & "\profilelist.txt")) > 0 Then
'            f = FreeFile
'
'            Open App.Path & "\profilelist.txt" For Input As #f
'
'                If LOF(f) > 0 Then
'                    Do
'                        Line Input #f, sInput
'                        colProfiles.Add sInput
'                    Loop While Not EOF(f)
'                End If
'
'            Close #f
'        End If
'    End If
'End Sub


' ProfileIndex param should only be used when changing profiles as a SET
'   - colProfiles MUST be instantiated in order to call with a ProfileIndex > 0!
Public Function GetProfilePath(Optional ByVal ProfileIndex As Integer) As String
    Static LastPath As String
    
    Dim s As String

'    If ProfileIndex > 0 Then
'        If ProfileIndex <= colProfiles.Count Then
'            s = colProfiles.Item(ProfileIndex)
'            'check for path\
'            If Right$(s, 1) <> "\" Then
'                s = s & "\"
'            End If
'
'            GetProfilePath = s
'        Else
'            If LenB(LastPath) > 0 Then
'                GetProfilePath = LastPath
'            Else
'                GetProfilePath = App.Path & "\"
'            End If
'        End If
'    Else
        If LenB(LastPath) > 0 Then
            GetProfilePath = LastPath
        Else
            GetProfilePath = App.Path & "\"
        End If
'    End If
    
    LastPath = GetProfilePath
End Function

Public Sub OpenReadme()
    ShellExecute frmChat.hWnd, "Open", "http://www.stealthbot.net/readme/", 0, 0, 0
    frmChat.AddChat RTBColors.SuccessText, "You are being taken to the StealthBot Online Readme."
End Sub

Public Function GetReadmePath() As String
    GetReadmePath = GetProfilePath() & "readme.chm"
End Function

'Checks the queue for duplicate bans
Public Sub RemoveBanFromQueue(ByVal sUser As String)
    Dim tmp As String ' ...

    ' ...
    tmp = "/ban " & sUser
        
    ' ...
    Call g_Queue.RemoveLines(tmp)
    
    ' ...
    If ((StrReverse$(BotVars.Product) = "WAR3") Or _
        (StrReverse$(BotVars.Product) = "W3XP")) Then
        
        Dim strGateway As String ' ...
        
        ' ...
        Select Case (BotVars.Gateway)
            Case "Lordaeron": strGateway = "@USWest"
            Case "Azeroth":   strGateway = "@USEast"
            Case "Kalimdor":  strGateway = "@Asia"
            Case "Northrend": strGateway = "@Europe"
        End Select
        
        ' ...
        Call g_Queue.RemoveLines(tmp & strGateway)
    End If
End Sub

Public Function AllowedToTalk(ByVal sUser As String, ByVal Msg As String) As Boolean
    Dim i As Integer
    
    ' default to true
    AllowedToTalk = True
    
    'For each condition where the user is NOT allowed to talk, set to false
    
    'i = UsernameToIndex(sUser)
    '
    'If i > 0 Then
    '    Dim CurrentGTC As Long
    '
    '    CurrentGTC = GetTickCount()
    '
    '    With colUsersInChannel.Item(i)
    '        If ((CurrentGTC - .JoinTime) < BotVars.AutofilterMS) Then
    '            AllowedToTalk = False
    '        End If
    '    End With
    'End If
    
    ' ...
    If (Filters) Then
        ' ...
        If ((CheckBlock(sUser)) Or (CheckMsg(Msg, sUser, -5))) Then
            AllowedToTalk = False
        End If
    End If
End Function


' Used by the Individual Whisper Window system to determine whether a message should be
'  forwarded to an IWW
Public Function IrrelevantWhisper(ByVal sIn As String, ByVal sUser As String) As Boolean
    IrrelevantWhisper = False
    
    If InStr(sIn, "ß~ß") Then
        IrrelevantWhisper = True
        Exit Function
    End If
    
    sUser = NameWithoutRealm(sUser, 1)
    
    'Debug.Print "strComp(" & Left$(sIn, 12 + Len(sUser)) & ", Your friend " & sUser & ")"
    
    If StrComp(Left$(sIn, 12 + Len(sUser)), "Your friend " & sUser) = 0 Then
        IrrelevantWhisper = True
        Exit Function
    End If
End Function

Public Sub UpdateSafelistedStatus(ByVal sUser As String, ByVal bStatus As Boolean)
    Dim i As Integer
    
    i = UsernameToIndex(sUser)
    
    If i > 0 Then
        colUsersInChannel.Item(i).Safelisted = bStatus
    End If
End Sub

Public Sub AddBannedUser(ByVal sUser As String, ByVal cOperator As String)
    Const MAX_BAN_COUNT As Integer = 80

    Dim i      As Integer ' ...
    Dim bCount As Integer ' ...
    
    ' check for duplicate entry in banlist
    For i = 0 To UBound(gBans)
        If (StrComp(gBans(i).Username, StripRealm(sUser), vbTextCompare) = 0) Then
            Exit Sub
        End If
    Next i
    
    ' count bans for channel operator
    For i = 0 To UBound(gBans)
        If (StrComp(gBans(i).cOperator, cOperator, vbTextCompare) = 0) Then
            bCount = (bCount + 1)
        End If
    Next i
    
    ' if ban count for operator greater than operator
    ' max, begin removing oldest bans.
    If (bCount >= MAX_BAN_COUNT) Then
        For i = 1 To (MAX_BAN_COUNT - 1)
            gBans(i - 1) = gBans(i)
        Next i
        
        With gBans(MAX_BAN_COUNT - 1)
            .Username = StripRealm(sUser)
            .UsernameActual = sUser
            .cOperator = cOperator
        End With
    Else
        With gBans(UBound(gBans))
            .Username = StripRealm(sUser)
            .UsernameActual = sUser
            .cOperator = cOperator
        End With
        
        ReDim Preserve gBans(0 To UBound(gBans) + 1)
    End If
End Sub

' collapse array on top of the removed user
Public Sub UnbanBannedUser(ByVal sUser As String, ByVal cOperator As String)
    Dim i          As Integer
    Dim c          As Integer
    Dim NumRemoved As Integer
    Dim iterations As Long
    Dim uBnd       As Integer
    
    sUser = StripRealm(sUser)
    
    uBnd = UBound(gBans)
    
    While (i <= (uBnd - NumRemoved))
        If (StrComp(sUser, gBans(i).Username, vbTextCompare) = 0) Then
            If (i <> UBound(gBans)) Then
                For c = i To UBound(gBans)
                    gBans(i) = gBans(i + 1)
                Next c
            End If
            
            ' UBound(gBans) - 1 when UBound(gBans) = 0
            ' causes an RTE.  Thanks PyroManiac and
            ' Phix (2008-10-1).
            If (UBound(gBans)) Then
                ReDim Preserve gBans(UBound(gBans) - 1)
            Else
                ReDim gBans(0)
            End If
            
            NumRemoved = (NumRemoved + 1)
        Else
            i = (i + 1)
        End If
        
        iterations = (iterations + 1)
        
        If (iterations > 9000) Then
            If (MDebug("debug")) Then
                frmChat.AddChat RTBColors.ErrorMessageText, "Warning: Loop size limit exceeded " & _
                    "in UnbanBannedUser()!"
                frmChat.AddChat RTBColors.ErrorMessageText, "The banned-user list has been reset.. " & _
                    "hope it works!"
            End If
            
            ReDim gBans(0)
            
            Exit Sub
        End If
    Wend
End Sub

Public Function IsBanned(ByVal sUser As String) As Boolean
    Dim i As Integer

    If (InStr(1, sUser, "#", vbBinaryCompare)) Then
        sUser = Left$(sUser, InStr(1, sUser, _
            "#", vbBinaryCompare) - 1)
        
        Debug.Print sUser
    End If
    
    For i = 0 To UBound(gBans)
        If (StrComp(sUser, gBans(i).UsernameActual, _
            vbTextCompare) = 0) Then
            
            IsBanned = True
            
            Exit Function
        End If
    Next i
End Function

Public Function IsValidIPAddress(ByVal sIn As String) As Boolean
    Dim s() As String
    Dim i   As Integer
    
    IsValidIPAddress = True
    
    If (InStr(1, sIn, ".", vbBinaryCompare)) Then
        s() = Split(sIn, ".")
        
        If (UBound(s) = 3) Then
            For i = 0 To 3
                If (Not (StrictIsNumeric(s(i)))) Then
                    IsValidIPAddress = False
                End If
            Next i
        Else
            IsValidIPAddress = False
        End If
    Else
        IsValidIPAddress = False
    End If
End Function

Public Function GetNameColor(ByVal flags As Long, ByVal IdleTime As Long, ByVal IsSelf As Boolean) As Long
    '/* Self */
    If (IsSelf) Then
        'Debug.Print "Assigned color IsSelf"
        GetNameColor = vbWhite
        
        Exit Function
    End If
    
    '/* Squelched */
    If ((flags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
        'Debug.Print "Assigned color SQUELCH"
        GetNameColor = &H99
        
        Exit Function
    End If
    
    '/* Blizzard */
    If (((flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or _
        ((flags And USER_SYSOP&) = USER_SYSOP&)) Then
       
        GetNameColor = COLOR_BLUE
        
        Exit Function
    End If
    
    '/* Operator */
    If ((flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        'Debug.Print "Assigned color OP"
        GetNameColor = &HDDDDDD
        Exit Function
    End If
    
    '/* Idle */
    If (IdleTime > BotVars.SecondsToIdle) Then
        'Debug.Print "Assigned color IDLE"
        GetNameColor = &HBBBBBB
        Exit Function
    End If
    
    '/* Default */
    'Debug.Print "Assigned color NORMAL"
    GetNameColor = COLOR_TEAL
End Function

Public Function FlagDescription(ByVal flags As Long) As String
    Dim s0ut          As String
    Dim multipleFlags As Boolean
        
    If ((flags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
        s0ut = "Squelched"
        
        multipleFlags = True
    End If
    
    If ((flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        If (multipleFlags) Then
            s0ut = s0ut & ", channel op"
        Else
            s0ut = "Channel op"
        End If
        
        multipleFlags = True
    End If
    
    If (((flags And USER_BLIZZREP) = USER_BLIZZREP) Or _
        ((flags And USER_SYSOP) = USER_SYSOP)) Then
       
        If (multipleFlags) Then
            s0ut = s0ut & _
                ", Blizzard representative"
        Else
            s0ut = "Blizzard representative"
        End If
        
        multipleFlags = True
    End If
    
    If ((flags And USER_NOUDP&) = USER_NOUDP&) Then
        If (multipleFlags) Then
            s0ut = s0ut & ", UDP plug"
        Else
            s0ut = "UDP plug"
        End If
        
        multipleFlags = True
    End If
    
    If (LenB(s0ut) = 0) Then
        If (flags = &H0) Then
            s0ut = "Normal"
        Else
            s0ut = "Altered"
        End If
    End If
    
    FlagDescription = s0ut & " [0x" & Right$("00000000" & Hex(flags), 8) & "]"
End Function

'Returns TRUE if the specified argument was a command line switch,
' such as -debug
Public Function MDebug(ByVal sArg As String) As Boolean
    MDebug = InStr(1, CommandLine, "-" & sArg, vbTextCompare) > 0
End Function

'Returns system uptime in milliseconds
Public Function GetUptimeMS() As Double
    Dim mmt   As MMTIME
    Dim lSize As Long
    
    lSize = LenB(mmt)
    
    Call timeGetSystemTime(mmt, lSize)

    GetUptimeMS = LongToUnsigned(mmt.units)
End Function


Public Function UsernameToIndex(ByVal sUsername As String) As Long
    Dim user        As clsUserInfo
    Dim FirstLetter As String * 1
    Dim i           As Integer
    
    FirstLetter = Mid$(sUsername, 1, 1)
    
    If (colUsersInChannel.Count > 0) Then
        For i = 1 To colUsersInChannel.Count
            Set user = colUsersInChannel.Item(i)
            
            With user
                If (StrComp(Mid$(.Username, 1, 1), FirstLetter, vbTextCompare) = 0) Then
                    If (StrComp(sUsername, .Username, vbTextCompare) = 0) Then
                        UsernameToIndex = i
                        
                        Exit Function
                    End If
                End If
            End With
        Next i
        
    End If
    
    UsernameToIndex = 0
End Function


Public Function checkChannel(ByVal NameToFind As String) As Integer
    Dim lvItem As ListItem

    Set lvItem = frmChat.lvChannel.FindItem(NameToFind)

    If (lvItem Is Nothing) Then
        checkChannel = 0
    Else
        checkChannel = lvItem.index
    End If
End Function


Public Sub CheckPhrase(ByRef Username As String, ByRef Msg As String, ByVal mType As Byte)
    Dim i As Integer
    
    If UBound(Catch) = 0 Then
        If Catch(0) = vbNullString Then Exit Sub
    End If
    
    For i = LBound(Catch) To UBound(Catch)
        If (Catch(i) <> vbNullString) Then
            If (InStr(1, LCase(Msg), Catch(i), vbTextCompare) <> 0) Then
                Call CaughtPhrase(Username, Msg, Catch(i), mType)
                
                Exit Sub
            End If
        End If
    Next i
End Sub


Public Sub CaughtPhrase(ByVal Username As String, ByVal Msg As String, ByVal Phrase As String, ByVal mType As Byte)
    Dim i As Integer
    Dim s As String
    
    i = FreeFile
    
    If (LenB(ReadCFG("Other", "FlashOnCatchPhrases")) > 0) Then
        Call FlashWindow
    End If
    
    Select Case (mType)
        Case CPTALK:    s = "TALK"
        Case CPEMOTE:   s = "EMOTE"
        Case CPWHISPER: s = "WHISPER"
    End Select
    
    If (Dir$(GetProfilePath() & "\caughtphrases.htm") = vbNullString) Then
        Open GetProfilePath() & "\caughtphrases.htm" For Output As #i
            Print #i, "<html>"
        Close #i
    End If
    
    Open GetProfilePath() & "\caughtphrases.htm" For Append As #i
        If (LOF(i) > 10000000) Then
            Close #i
            
            Call Kill(GetProfilePath() & "\caughtphrases.htm")
            
            Open GetProfilePath() & "\caughtphrases.htm" For Output As #i
        End If
        
        Msg = Replace(Msg, "<", "&lt;", 1)
        Msg = Replace(Msg, ">", "&gt;", 1)
        
        Print #i, "<B>" & Format(Date, "MM-dd-yyyy") & " - " & Time & _
            " - " & s & Space(1) & Username & ": </B>" & _
                Replace(Msg, Phrase, "<i>" & Phrase & "</i>", 1) & _
                    "<br>"
    Close #i
End Sub


Public Function DoReplacements(ByVal s As String, Optional Username As String, _
    Optional Ping As Long) As String

    Dim gAcc As udtGetAccessResponse
    
    gAcc = GetCumulativeAccess(Username)

    s = Replace(s, "%0", Username, 1)
    s = Replace(s, "%1", CurrentUsername, 1)
    s = Replace(s, "%c", gChannel.Current, 1)
    s = Replace(s, "%bc", BanCount, 1)
    
    If (Ping > -2) Then
        s = Replace(s, "%p", Ping, 1)
    End If
    
    s = Replace(s, "%v", CVERSION, 1)
    s = Replace(s, "%a", IIf(gAcc.access >= 0, gAcc.access, "0"), 1)
    s = Replace(s, "%f", gAcc.flags, 1)
    s = Replace(s, "%t", Time$, 1)
    s = Replace(s, "%d", Date, 1)
    s = Replace(s, "%m", GetMailCount(Username), 1)
    
    DoReplacements = s
End Function

' Updated 4/10/06 to support millisecond pauses
'  If using milliseconds pause for at least 100ms
Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True, Optional ByVal milliseconds As Boolean = False)
    Dim i As Integer
    
    If (AllowEvents) Then
        For i = 0 To (fSeconds * (IIf(milliseconds, 1, 1000))) \ 100
            'Debug.Print "sleeping 100ms"
            Call Sleep(100)
            
            DoEvents
        Next i
    Else
        Call Sleep(fSeconds * (IIf(milliseconds, 1, 1000)))
    End If
End Sub

Public Sub LogDBAction(ByVal ActionType As enuDBActions, ByVal Caller As String, ByVal Target As String, ByVal Instruction As String)
    Dim sPath  As String
    Dim Action As String
    Dim f      As Integer
    
    f = FreeFile
    sPath = GetProfilePath() & "\Logs\database.txt"
    
    If (Len(Caller) < 2) Then
        Caller = "bot console"
    End If
    
    If (LenB(Dir$(sPath)) = 0) Then
        Open sPath For Output As #f
    Else
        Open sPath For Append As #f
        
        If ((LOF(f) > BotVars.MaxLogFileSize) And (BotVars.MaxLogFileSize > 0)) Then
            Close #f
            
            Call Kill(sPath)
            
            Open sPath For Output As #f
                Print #f, "Logfile cleared automatically on " & _
                    Format(Now, "HH:MM:SS MM/DD/YY") & "."
        End If
    End If
    
    Select Case (ActionType)
        Case AddEntry: Action = "adds"
        Case RemEntry: Action = "removes"
        Case ModEntry: Action = "modifies"
    End Select
    
    Action = "[" & Format(Now, "HH:MM:SS MM/DD/YY") & "] " & _
        Caller & " " & Action & Space(1) & Target & ": " & Instruction
    
    Print #f, Action
    
    Close #f
End Sub

Public Sub LogCommand(ByVal Caller As String, ByVal CString As String)
    On Error GoTo LogCommand_Error
    
    Dim sPath  As String
    Dim Action As String
    Dim f      As Integer

    If (LenB(CString) > 0) Then
        f = FreeFile
        
        sPath = GetProfilePath() & "\Logs\commands.txt"
        
        If (LenB(Caller) = 0) Then
            Caller = "bot console"
        End If
        
        If (LenB(Dir$(sPath)) = 0) Then
            Open sPath For Output As #f
        Else
            Open sPath For Append As #f
            
            If ((LOF(f) > BotVars.MaxLogFileSize) And (BotVars.MaxLogFileSize > 0)) Then
                Close #f
                
                Call Kill(sPath)
                
                Open sPath For Output As #f
                    Print #f, "Logfile cleared automatically on " & _
                        Format(Now, "HH:MM:SS MM/DD/YY") & "."
            End If
        End If
        
        Action = "[" & Format(Now, "HH:MM:SS MM/DD/YY") & _
            "][" & Caller & "]-> " & CString
        
        Print #f, Action
        
        Close #f
    End If

    Exit Sub

LogCommand_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in " & _
        "Procedure; LogCommand; of; Module; modOtherCode; "
    
    Exit Sub
End Sub

'Pos must be >0
' Returns a single chunk of a string as if that string were Split() and that chunk
' extracted
' 1-based
Public Function GetStringChunk(ByVal str As String, ByVal pos As Integer)
    Dim c           As Integer
    Dim i           As Integer
    Dim TargetSpace As Integer
    
    'one two three
    '   1   2
    
    c = 0
    i = 1
    pos = pos
    
    ' The string must have at least (pos-1) spaces to be valid
    While ((c < pos) And (i > 0))
        TargetSpace = i
        
        i = (InStr(i + 1, str, Space(1), vbBinaryCompare))
        
        c = (c + 1)
    Wend
    
    If (c >= pos) Then
        c = InStr(TargetSpace + 1, str, " ") ' check for another space (more afterwards)
        
        If (c > 0) Then
            GetStringChunk = Mid$(str, TargetSpace, c - (TargetSpace))
        Else
            GetStringChunk = Mid$(str, TargetSpace)
        End If
    Else
        GetStringChunk = ""
    End If
    
    GetStringChunk = Trim(GetStringChunk)
End Function

'Public Sub SpamCheck(ByVal User As String, ByVal Msg As String)
'    Static Top As Integer
'    Dim i As Integer
'
'    If Len(Msg) > 8 Then
'        i = InStr(User, "#")
'
'        If i > 0 Then
'            user = left$(
'
'        Last4Messages(Top) = LCase(Msg)
'        Last4Speakers(Top) = LCase(User)
'        Top = Top + 1
'
'        If Top > 3 Then Top = 0
'    End If
'
'End Sub

Function GetProductKey(Optional ByVal Product As String) As String
    If (LenB(Product) = 0) Then
        Product = StrReverse$(BotVars.Product)
    End If
    
    Select Case Product
        Case "W2BN", "NB2W": GetProductKey = "W2"
        Case "STAR", "RATS": GetProductKey = "SC"
        Case "SEXP", "PXES": GetProductKey = "SC"
        Case "D2DV", "PX2D": GetProductKey = "D2"
        Case "D2XP", "VD2D": GetProductKey = "D2"
        Case "WAR3", "3RAW": GetProductKey = "W3"
        Case "W3XP", "PX3W": GetProductKey = "W3"
    End Select
End Function

Public Function InsertDummyQueueEntry()
    ' %%%%%blankqueuemessage%%%%%
    frmChat.AddQ "%%%%%blankqueuemessage%%%%%"
End Function

' This procedure splits a message by a specified length, with optional line postfixes
' and split delimiters.
Public Function SplitByLen(StringSplit As String, SplitLength As Long, ByRef StringRet() As String, _
    Optional LinePostfix As String = " [more]", Optional OversizeDelimiter As String = " ")
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    Dim lineCount As Long    ' stores line number
    Dim pos       As Long    ' stores position of delimiter
    Dim strTmp    As String  ' stores working copy of StringSplit
    Dim Length    As Long    ' stores length after postfix
    Dim bln       As Boolean ' stores result of delimiter split
    
    ' initialize our array
    ReDim StringRet(0)
    
    ' default our first index
    StringRet(0) = vbNullString
    
    ' do loop until our string is empty
    Do While (StringSplit <> vbNullString)
        ' resize array so that it can store
        ' the next line
        ReDim Preserve StringRet(lineCount)
        
        ' store working copy of string
        strTmp = StringSplit
        
        ' does our string already equal to or fall
        ' below the specified length?
        If (Len(strTmp) <= SplitLength) Then
            ' ...
            'If (Right$(strTmp, _
            '    Len(OversizeDelimiter)) = OversizeDelimiter) Then
            '
            '    ' ...
            '    strTmp = Left$(strTmp, Len(strTmp) - _
            '        Len(OversizeDelimiter))
            'End If
        
            ' assign our string to the current line
            StringRet(lineCount) = strTmp
        Else
            ' Our string is over the size limit, so we're
            ' going to postfix it.  Because of this, we're
            ' going to have to calculate the length after
            ' the post fix has been accounted for.
            Length = (SplitLength - Len(LinePostfix))
        
            ' if we're going to be splitting the oversized
            ' message at a specified character, we need to
            ' determine the position of the character in the
            ' string
            If (OversizeDelimiter <> vbNullString) Then
                ' grab position of delimiter character that is
                ' the closest to our specified length
                pos = InStrRev(StringSplit, OversizeDelimiter, _
                    Length, vbBinaryCompare)
            End If
            
            ' if the delimiter we were looking for was found,
            ' and the position was greater than or equal to
            ' half of the message (this check prevents breaks
            ' in unecessary locations), split the message
            ' accordingly.
            If ((pos) And (pos >= Round(Length / 2))) Then
                ' truncate message
                strTmp = Mid$(strTmp, 1, pos - 1)
                
                ' indicate that an additional
                ' character will require removal
                ' from official copy
                bln = True
            Else
                ' truncate message
                strTmp = Mid$(strTmp, 1, Length)
            End If
            
            ' store truncated message in line
            StringRet(lineCount) = strTmp & _
                LinePostfix
        End If
        
        ' remove line from official string
        StringSplit = Mid$(StringSplit, _
            (Len(strTmp) + 1))
        
        ' if we need to remove an additional
        ' character, lets do so now.
        If (bln) Then
            StringSplit = Mid$(StringSplit, 2)
        End If
            
        ' increment line counter
        lineCount = (lineCount + 1)
    Loop
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, Err.Description & " in SplitByLen().")
    
    Exit Function
End Function

' Thanks strtok()!
Public Function IsCommand(Optional ByVal str As String = vbNullString, Optional DontCheckTrigger As Boolean = _
    False) As clsCommandObj
    
    ' ...
    Const CMD_DELIMITER As String = "; "

    Static Message   As String  ' ...
    Static CropLen   As Integer ' ...

    Dim index        As Integer ' ...
    Dim bln          As Boolean ' ...
    Dim tmp          As String  ' ...
    Dim console      As Boolean ' ...
    Dim PublicOutput As Boolean ' ...
    
    ' ...
    Set IsCommand = New clsCommandObj
    
    ' ...
    If (str <> vbNullString) Then
        ' ...
        Message = str
        
        ' ...
        CropLen = 0
    Else
        ' ...
        If (Len(Message) <= CropLen) Then
            With IsCommand
                .Name = vbNullString
                .Args = vbNullString
            End With
        
            Exit Function
        End If
    End If

    ' ...
    If (Left$(Message, 1) = "/") Then
        ' ...
        console = True
        
        ' ...
        If (Left$(Message, 2) = "//") Then
            PublicOutput = True
            
            If (CropLen = 0) Then
                CropLen = Len("//")
            End If
        Else
            PublicOutput = False
            
            If (CropLen = 0) Then
                CropLen = Len("/")
            End If
        End If
    Else
        ' ...
        console = False
        
        ' ...
        PublicOutput = False
    End If
    
    ' ...
    If (CropLen) Then
        tmp = Mid$(Message, CropLen + 1)
    Else
        tmp = Message
    End If
    
    ' ...
    If ((console = False) And (DontCheckTrigger = False)) Then
        ' ...
        If (Left$(Message, Len(BotVars.TriggerLong)) = BotVars.TriggerLong) Then
            ' ...
            CropLen = (CropLen + Len(BotVars.TriggerLong))
        
            ' ...
            If (StrComp(Left$(Message, Len(CurrentUsername) + 1), CurrentUsername & Space(1), _
                vbTextCompare) = 0) Then
                
                CropLen = (CropLen + Len(CurrentUsername))
            End If
            
            bln = True
        Else
            If (Left$(tmp, 1) = "?") Then
                If (StrComp(tmp, "trigger", vbTextCompare) = 0) Then
                    CropLen = (CropLen + Len("?"))
                
                    bln = True
                End If
            Else
                ' ...
                If (StrComp(Left$(tmp, Len(CurrentUsername)), CurrentUsername, _
                        vbTextCompare) = 0) Then
    
                    If (Mid$(tmp, Len(CurrentUsername) + 1, 2) = ": ") Then
                        CropLen = (CropLen + (Len(CurrentUsername) + Len(": ")))
                            
                        bln = True
                    ElseIf (Mid$(tmp, Len(CurrentUsername) + 1, 2) = ", ") Then
                        CropLen = (CropLen + (Len(CurrentUsername) + Len(", ")))
    
                        bln = True
                    End If
                End If
            End If
        End If
        
        ' ...
        If (bln = True) Then
            tmp = Mid$(tmp, CropLen + 1)
        End If
    End If
    
    ' check our message for a command delimiter
    index = InStr(Len(BotVars.TriggerLong) + 1, tmp, CMD_DELIMITER, _
        vbBinaryCompare)
    
    ' using a delimiter can be undesirable at times, so
    ' we require a way of bypassing such a feature, and
    ' that way is to entirely disable internal support!
    If ((index) And (console = False)) Then
        ' ...
        tmp = Mid$(tmp, 1, index - 1)
        
        ' ...
        CropLen = (CropLen + (Len(tmp) + Len(CMD_DELIMITER)))
    Else
        ' ...
        CropLen = Len(Message)
    End If
    
    ' ...
    If ((console) Or (bln)) Then
        ' ...
        index = InStr(1, tmp, Space$(1), vbBinaryCompare)
        
        ' ...
        If (index) Then
            With IsCommand
                .Name = Mid$(tmp, 1, index - 1)
                .Args = Mid$(tmp, index + 1)
            End With
        Else
            IsCommand.Name = tmp
        End If

        With IsCommand
            .IsLocal = console
            .PublicOutput = PublicOutput
        End With

        ' ...
        If (IsCommand.Name <> vbNullString) Then
            ' ...
            If (IsCommand.docs.Name = vbNullString) Then
                ' ...
                IsCommand.Name = convertAlias(IsCommand.Name)
                
                ' ...
                If (IsCommand.docs.Name = vbNullString) Then
                    ' ...
                    Set IsCommand = IsCommand(vbNullString)
                End If
                
                ' ...
                Exit Function
            End If
        End If

        ' ...
        Exit Function
    End If
    
    With IsCommand
        .Name = vbNullString
        .Args = vbNullString
    End With
End Function

' ...
Public Function convertAlias(ByVal cmdName As String) As String
    ' ...
    On Error GoTo ERROR_HANDLER

    ' ...
    If (Len(cmdName) > 0) Then
        Dim Commands As MSXML2.DOMDocument
        Dim Command  As MSXML2.IXMLDOMNode
        
        ' ...
        Set Commands = New MSXML2.DOMDocument
        
        ' ...
        If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
            Call frmChat.AddChat(RTBColors.ConsoleText, "Error: The XML database could not be found in the " & _
                "working directory.")
                
            Exit Function
        End If
        
        ' ...
        Call Commands.Load(App.Path & "\commands.xml")
        
        ' ...
        For Each Command In Commands.documentElement.childNodes
            Dim aliases As MSXML2.IXMLDOMNodeList
            Dim alias   As MSXML2.IXMLDOMNode
            
            ' ...
            Set aliases = Command.selectNodes("alias")

            ' ...
            For Each alias In aliases
                ' ...
                If (StrComp(alias.text, cmdName, vbTextCompare) = 0) Then
                    ' ...
                    convertAlias = Command.Attributes.getNamedItem("name").text
                    
                    Exit Function
                End If
            Next
        Next
    End If
    
    ' ...
    convertAlias = cmdName
    
    Exit Function
    
' ...
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ConsoleText, "Error: XML Database Processor has encountered an error " & _
        "during alias lookup.")
        
    ' ...
    convertAlias = False

    Exit Function
End Function

Public Sub DisplayRichText(ByRef rtb As RichTextBox, ByRef saElements() As Variant)
    On Error Resume Next
    
    Dim s              As String
    Dim l              As Long
    Dim lngVerticalPos As Long
    Dim Diff           As Long
    Dim i              As Integer
    Dim intRange       As Integer
    Dim f              As Integer
    Dim blUnlock       As Boolean
    Dim LogThis        As Boolean
    
    If (BotVars.LockChat = False) Then
        f = FreeFile
    
        If (IsWin2000Plus()) Then
            Call GetScrollRange(rtb.hWnd, SB_VERT, 0, intRange)
            
            lngVerticalPos = SendMessage(rtb.hWnd, EM_GETTHUMB, 0&, 0&)
            
            Diff = ((lngVerticalPos + _
                (rtb.Height / Screen.TwipsPerPixelY)) - intRange)
            
            ' In testing it appears that if the value I calcuate as Diff is negative,
            ' the scrollbar is not at the bottom.
            If (Diff < 0) Then
                rtb.Visible = False
            
                blUnlock = True
            End If
        End If
        
        LogThis = (BotVars.Logging <= 1)
        
        If ((BotVars.MaxBacklogSize) And _
            (rtbChatLength >= BotVars.MaxBacklogSize)) Then
            
            With rtb
                .Visible = False
                .SelStart = 0
                .SelLength = InStr(1, .text, vbLf, vbBinaryCompare)
                
                rtbChatLength = (rtbChatLength - .SelLength)
                
                .SelText = ""
                .Visible = True
            End With
        End If
        
        s = GetTimeStamp()
        
        With rtb
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = RTBColors.TimeStamps
            If .SelBold = True Then: .SelBold = False
            If .SelItalic = True Then: .SelItalic = False
            If .SelUnderline = True Then: .SelUnderline = False
            .SelText = s
            .SelStart = Len(.text)
        End With
        
        If (LogThis) Then
            Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #f
            
            If ((BotVars.MaxLogFileSize) And _
                (LOF(f) >= BotVars.MaxLogFileSize)) Then
                
                LogThis = False
                
                Close #f
            Else
                Print #f, s;
            End If
        End If

        For i = LBound(saElements) To UBound(saElements) Step 2
            If (InStr(1, saElements(i + 1), Chr(0), vbBinaryCompare) > 0) Then
                Call KillNull(saElements(i + 1))
            End If
            
            If (Len(saElements(i + 1)) > 0) Then
                l = InStr(1, saElements(i + 1), "{\rtf", vbTextCompare)
                
                While (l > 0)
                    Mid$(saElements(i + 1), l + 1, 1) = "/"
                    
                    l = InStr(1, saElements(i + 1), "{\rtf", vbTextCompare)
                Wend
            
                With rtb
                    .SelStart = Len(.text)
                    
                    ' store position of selection
                    l = .SelStart
                    
                    .SelLength = 0
                    .SelColor = saElements(i)
                    .SelText = saElements(i + 1) & _
                        Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                    
                    rtbChatLength = (rtbChatLength + _
                                     Len(s) + _
                                     Len(saElements(i + 1)) + _
                                     Len(Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))))
                    
                    .SelStart = Len(.text)
                End With
                
                'With TextBox1
                '    .SelStart = Len(.text)
                '
                '    ' store position of selection
                '    l = .SelStart
                '
                '    .SelLength = 0
                '    '.SelColor = saElements(i)
                '    .SelText = saElements(i + 1) & _
                '        Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                '
                '    rtbChatLength = (rtbChatLength + _
                '                     Len(s) + _
                '                     Len(saElements(i + 1)) + _
                '                     Len(Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))))
                '
                '    .SelStart = Len(.text)
                'End With
                
                ' Fixed 11/21/06 to properly log timestamps
                If (LogThis) Then
                    Print #f, saElements(i + 1) & _
                        Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)));
                End If
            End If
        Next i
        
        Call ColorModify(rtb, l)

        If (blUnlock) Then
            rtb.Visible = True
            
            Call SendMessage(rtb.hWnd, WM_VSCROLL, _
                SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&)
        End If
        
        If (LogThis) Then
            Close #f
        End If
    End If
    
    'Dim hm As String
    
    
    'Dim blah As Object
    
    'set blah = CreateObject(
    
    'hm = &HC0C9
    
    'With rtbChat
    '    .SelStart = Len(.text)
    '    .SelFontName = "Arial Unicode MS"
    'End With

    'SendMessage frmChat.rtbChat.hWnd, WM_SETTEXT, 0, ByVal StrPtr(hm)
    
    'rtbChat.Refresh
End Sub
