Attribute VB_Name = "modOtherCode"
Option Explicit

'Read/WriteIni code thanks to ickis
Public Sub WriteINI(ByVal wiSection$, ByVal wiKey As String, ByVal wiValue As String, Optional ByVal wiFile As String = "x")
    If StrComp(wiFile, "x") = 0 Then
        wiFile = GetConfigFilePath()
    Else
        If InStr(wiFile$, "\") = 0 Then
            wiFile = GetFilePath(wiFile)
        End If
    End If
    
    WritePrivateProfileString wiSection, wiKey, wiValue, wiFile
End Sub

Public Function ReadCFG$(ByVal riSection$, ByVal riKey$)
    Dim sRiBuffer As String
    Dim sRiValue As String
    Dim sRiLong As String
    Dim riFile As String
    
    riFile = GetConfigFilePath()
    
    If Dir(riFile) <> vbNullString Then
        sRiBuffer = String(255, vbNull)
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, riFile)
        If Left$(sRiBuffer, 1) <> Chr(1) Then
            sRiValue = Left$(sRiBuffer, sRiLong)
            ReadCFG = sRiValue
        End If
    Else
        ReadCFG = ""
    End If
End Function

Public Function ReadINI$(ByVal riSection$, ByVal riKey$, ByVal riFile$)
    Dim sRiBuffer$
    Dim sRiValue$
    Dim sRiLong$
    
    If InStr(riFile$, "\") = 0 Then
        riFile$ = GetFilePath(riFile)
    End If
    
    If Dir(riFile$) <> vbNullString Then
        sRiBuffer = String(255, vbNull)
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, riFile)
        If Left$(sRiBuffer, 1) <> Chr(1) Then
            sRiValue = Left$(sRiBuffer, sRiLong)
            ReadINI = sRiValue
        End If
    Else
        ReadINI = ""
    End If
End Function

Public Function GetTimestamp() As String
    Select Case BotVars.TSSetting
        Case 0: GetTimestamp = " [" & Time & "] "
        Case 1: GetTimestamp = " [" & Format(Time, "HH:MM:SS") & "] "
        Case 2: GetTimestamp = " [" & Format(Time, "HH:MM:SS") & "." & GetCurrentMS & "] "
        Case 3: GetTimestamp = ""
    End Select
End Function

'// Converts a millisecond or second time value to humanspeak.. modified to support BNet's Time
'// Logged report.
'// Updated 2/11/05 to support timeGetSystemTime() unsigned longs, which in VB are doubles after conversion
Public Function ConvertTime(ByVal dblMS As Double, Optional Seconds As Byte) As String
    Dim dblSeconds As Double, dblDays As Double, dblHours As Double, dblMins As Double
    Dim strSeconds As String, strDays As String
    
    If Seconds = 0 Then
        dblSeconds = dblMS / 1000
    Else
        dblSeconds = dblMS
    End If
    
    dblDays = Int(dblSeconds / 86400)
    dblSeconds = dblSeconds Mod 86400
    dblHours = Int(dblSeconds / 3600)
    dblSeconds = dblSeconds Mod 3600
    dblMins = Int(dblSeconds / 60)
    dblSeconds = dblSeconds Mod 60
    
    If dblSeconds <> 1 Then strSeconds = "s"
    If dblDays <> 1 Then strDays = "s"
    
    ConvertTime = dblDays & " day" & strDays & ", " & dblHours & " hours, " & dblMins & " minutes and " & dblSeconds & " second" & strSeconds
End Function

Public Function GetVerByte(Product As String, Optional ByVal UseHardcode As Integer) As Long

    Dim Key As String
    
    Key = GetProductKey(Product)
    
    If ReadCFG("Override", Key & "VerByte") = vbNullString Or UseHardcode = 1 Then
        Select Case StrReverse(Product)
            Case "W2BN"
                GetVerByte = &H4F
            Case "STAR", "SEXP"
                GetVerByte = &HCF   ' Patch: 1.14 7/31/06
            Case "D2DV", "D2XP"
                GetVerByte = &HB
            Case "W3XP", "WAR3"
                GetVerByte = &H14
        End Select
    Else
        GetVerByte = CLng(Val("&H" & ReadCFG("Override", Key & "VerByte")))
    End If
    
End Function

Public Function GetGamePath(ByVal Client As String) As String
    Dim Key As String
    
    Key = GetProductKey(Client)
    
    If LenB(ReadCFG("Override", Key & "Hashes")) > 0 Then
        GetGamePath = ReadCFG("Override", Key & "Hashes")
    Else
        Select Case StrReverse(UCase(Client))
            Case "W2BN"
                GetGamePath = App.Path & "\W2BN\"
            Case "STAR", "SEXP"
                GetGamePath = App.Path & "\STAR\"
            Case "D2DV"
                GetGamePath = App.Path & "\D2DV\"
            Case "D2XP"
                GetGamePath = App.Path & "\D2XP\"
            Case "W3XP", "WAR3"
                GetGamePath = App.Path & "\WAR3\"
            Case Else
                frmChat.AddChat RTBColors.ErrorMessageText, "Warning: Invalid game client in GetGamePath()"
        End Select
    End If
End Function

Function MKL(Value As Long) As String
    Dim result As String * 4
    CopyMemory ByVal result, Value, 4
    MKL = result
End Function

Function MKI(Value As Integer) As String
    Dim result As String * 2
    CopyMemory ByVal result, Value, 2
    MKI = result
End Function

Public Function CheckPath(ByVal sPath As String) As Long
    If LenB(Dir$(sPath)) = 0 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[HASHES] " & Mid$(sPath, InStrRev(sPath, "\") + 1) & " is missing."
        CheckPath = 1
    End If
End Function

Public Function Ban(ByVal Inpt As String, SpeakerAccess As Integer, Optional Kick As Integer) As String
    Dim Username As String, CleanedUsername As String
    Static LastBan As String
    Dim i As Integer
    
    If LenB(Inpt) > 0 Then
        If Kick > 2 Then
            LastBan = vbNullString
            Exit Function
        End If
        
        If colQueue.Count > 0 Then
            For i = 1 To colQueue.Count
                With colQueue.Item(i)
                    If Left$(.Message, 5) = "/ban " Then
                        If StrComp(LastBan & " ", LCase$(Mid$(.Message, 6, Len(LastBan) + 1)), vbTextCompare) = 0 Then
                            Exit Function
                        End If
                    End If
                End With
            Next i
        End If
        
        If MyFlags = 2 Or MyFlags = 18 Then
            
            If InStr(1, Inpt, " ", vbTextCompare) > 0 Then
                Username = LCase(Left$(Inpt, InStr(1, Inpt, " ", vbTextCompare) - 1))
            Else
                Username = LCase(Inpt)
            End If
            
            If LenB(Username) > 0 Then
                LastBan = LCase(Username)
                CleanedUsername = StripInvalidNameChars(Username)
                
                'If InStr(1, CleanedUsername, "@") > 0 Then CleanedUsername = StripRealm(CleanedUsername)
                
                If SpeakerAccess < 999 Then
                    If GetSafelist(CleanedUsername) Or GetSafelist(Username) Then
                        Ban = "That user is safelisted."
                        Exit Function
                    End If
                End If
                
                If GetAccess(Username).Access >= SpeakerAccess Then
                    Ban = "You do not have enough access to do that."
                    Exit Function
                End If
                
                If Dii Then
                    If Kick = 0 Then
                        frmChat.AddQ "/ban *" & Inpt, 1
                    Else
                        frmChat.AddQ "/kick *" & Inpt, 1
                    End If
                Else
                    If Kick = 0 Then
                        frmChat.AddQ "/ban " & Inpt, 1
                    Else
                        frmChat.AddQ "/kick " & Inpt, 1
                    End If
                End If
            End If
        Else
            Ban = "The bot does not have ops."
        End If
    End If
End Function

' This function created in response to http://www.stealthbot.net/forum/index.php?showtopic=20550
Public Function StripInvalidNameChars(ByVal Username As String) As String
    Dim Allowed(13) As Integer
    Dim i As Integer, j As Integer, thisChar As Integer
    Dim NewUsername As String
    Dim ThisCharOK As Boolean
    
    If LenB(Username) > 0 Then
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
        
        For i = 1 To Len(Username)
            thisChar = Asc(Mid$(Username, i, 1))
            ThisCharOK = False
            
            If Not IsAlpha(thisChar) Then
                If Not IsNumber(thisChar) Then
                    For j = 0 To UBound(Allowed)
                        If thisChar = Allowed(j) Then
                            ThisCharOK = True
                        End If
                    Next j
                    
                    If Not ThisCharOK Then
                        NewUsername = Replace(NewUsername, Chr(thisChar), "")
                    End If
                End If
            End If
        Next i
        
        StripInvalidNameChars = NewUsername
    End If
End Function

Public Function StripRealm(ByVal Username As String) As String
    If InStr(1, Username, "@") > 0 Then
        Username = Left$(Username, InStr(Username, "@") - 1)
    End If
    
    StripRealm = Username
End Function

Public Sub bnetSend(ByVal Message As String)
    If frmChat.sckBNet.State = 7 Then
        With PBuffer
'            If frmChat.mnuUTF8.Checked Then
'                .InsertNTString UTF8Encode(Message)
'            Else
                .InsertNTString Message
'            End If
            
            .SendPacket &HE
        End With
    End If
    
    If Not bFlood Then
        On Error Resume Next
        frmChat.SControl.Run "Event_MessageSent", Message
    End If
End Sub

Public Sub APISend(ByRef s As String) '// faster API-based sending for EFP

    Dim i As Long
    
    i = Len(s) + 5
    
    Call Send(frmChat.sckBNet.SocketHandle, "ÿ" & "" & Chr(i) & Chr(0) & s & Chr(0), i, 0)
    
End Sub

Public Function Voting(ByVal Mode1 As Byte, Optional Mode2 As Byte, Optional Username As String) As String
    Static Voted() As String
    Static VotesYes As Integer
    Static VotesNo As Integer
    Static VoteMode As Byte
    Static Target As String
        
    Dim i As Integer
    
    Select Case Mode1
        Case BVT_VOTE_ADD
            
            For i = LBound(Voted) To UBound(Voted)
                If StrComp(Voted(i), LCase(Username), vbTextCompare) = 0 Then
                    Exit Function
                End If
            Next i
            
            Select Case Mode2
                Case BVT_VOTE_ADDYES
                    VotesYes = VotesYes + 1
                Case BVT_VOTE_ADDNO
                    VotesNo = VotesNo + 1
            End Select
            
            Voted(UBound(Voted)) = LCase(Username)
            ReDim Preserve Voted(0 To UBound(Voted) + 1)
        
        Case BVT_VOTE_START
        
            VotesYes = 0
            VotesNo = 0
            ReDim Voted(0)
            VoteMode = Mode2
            Target = Username
            Voting = "Vote started. Type YES or NO to vote. Your vote will only be counted once."
        
        Case BVT_VOTE_END
            If Mode2 = BVT_VOTE_CANCEL Then
                Voting = "Vote cancelled. Final results: [" & VotesYes & "] YES, [" & VotesNo & "] NO. "
            Else
            
                Select Case VoteMode
                    Case BVT_VOTE_STD
                    
                        Voting = "Final results: [" & VotesYes & "] YES, [" & VotesNo & "] NO. "
                        If VotesYes > VotesNo Then
                            Voting = Voting & "YES wins, with " & Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
                        ElseIf VotesYes < VotesNo Then
                            Voting = Voting & "NO wins, with " & Format(VotesNo / (VotesYes + VotesNo), "percent") & " of the vote."
                        Else
                            Voting = Voting & "The vote was a draw."
                        End If
                        
                    Case BVT_VOTE_BAN
                        If VotesYes > VotesNo Then
                            Voting = Ban(Target & " Banned by vote", VoteInitiator.Access)
                        Else
                            Voting = "Ban vote failed."
                        End If
                        
                    Case BVT_VOTE_KICK
                        If VotesYes > VotesNo Then
                            Voting = Ban(Target & " Kicked by vote", VoteInitiator.Access, 1)
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
        
            Voting = "Current results: [" & VotesYes & "] YES, [" & VotesNo & "] NO; " & VoteDuration & " seconds remain."
            If VotesYes > VotesNo Then
                Voting = Voting & " YES leads, with " & Format(VotesYes / (VotesYes + VotesNo), "percent") & " of the vote."
            ElseIf VotesYes < VotesNo Then
                Voting = Voting & " NO leads, with " & Format(VotesNo / (VotesYes + VotesNo), "percent") & " of the vote."
            Else
                Voting = Voting & " The vote is a draw."
            End If
                    
            
    End Select
End Function

Public Function GetAccess(ByVal Username As String) As udtGetAccessResponse
    Dim i As Integer

    Username = Username
    
    If (Left$(Username, 1) = "*") Then
        Username = Mid$(Username, 2)
    End If

    For i = LBound(DB) To UBound(DB)
        If (StrComp(DB(i).Username, Username, vbTextCompare) = 0) Then
            GetAccess.Username = DB(i).Username
            GetAccess.Access = DB(i).Access
            GetAccess.flags = DB(i).flags
            GetAccess.AddedBy = DB(i).AddedBy
            GetAccess.AddedOn = DB(i).AddedOn
            GetAccess.ModifiedBy = DB(i).ModifiedBy
            GetAccess.ModifiedOn = DB(i).ModifiedOn
            
            Exit Function
        End If
    Next i

    GetAccess.Access = -1
End Function

Public Sub RequestSystemKeys()
    
    AwaitingSystemKeys = 1
    
    With PBuffer
        .InsertDWORD &H1
        .InsertDWORD &H4
        .InsertDWORD GetTickCount()
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
        
    ElseIf (flags And USER_SQUELCHED) = USER_SQUELCHED Then 'squelched
        i = ICSQUELCH
        
    ElseIf (flags And USER_CHANNELOP) = USER_CHANNELOP Then 'op
        i = ICGAVEL
        
    ElseIf g_ThisIconCode <> -1 Then
        If sProduct = "W3XP" Then
            i = g_ThisIconCode + ICON_START_W3XP + IIf(g_ThisIconCode + ICON_START_W3XP = ICSCSW, 1, 0)
        Else
            i = g_ThisIconCode + ICON_START_WAR3
        End If
        
    Else
        Select Case sProduct
            Case Is = "STAR"
                i = ICSTAR
            Case Is = "SEXP"
                i = ICSEXP
            Case Is = "D2DV"
                i = ICD2DV
            Case Is = "D2XP"
                i = ICD2XP
            Case Is = "W2BN"
                i = ICW2BN
            Case Is = "CHAT"
                i = ICCHAT
            Case Is = "DRTL"
                i = ICDIABLO
            Case Is = "DSHR"
                i = ICDIABLOSW
            Case Is = "JSTR"
                i = ICJSTR
            Case Is = "SSHR"
                i = ICSCSW
                
            '*** Special icons for WCG added 6/24/07 ***
            Case Is = "WCRF"
                i = IC_WCRF
            Case Is = "WCPL"
                i = IC_WCPL
            Case Is = "WCGO"
                i = IC_WCGO
            Case Is = "WCSI"
                i = IC_WCSI
            Case Is = "WCBR"
                i = IC_WCBR
            Case Is = "WCPG"
                i = IC_WCPG
                
            '*** Special icons for PGTour ***
            Case Is = "__A+"
                i = IC_PGT_A + 1
            Case Is = "___A"
                i = IC_PGT_A
            Case Is = "__A-"
                i = IC_PGT_A - 1
                
            Case Is = "__B+"
                i = IC_PGT_B + 1
            Case Is = "___B"
                i = IC_PGT_B
            Case Is = "__B-"
                i = IC_PGT_B - 1
                
            Case Is = "__C+"
                i = IC_PGT_C + 1
            Case Is = "___C"
                i = IC_PGT_C
            Case Is = "__C-"
                i = IC_PGT_C - 1
                
            Case Is = "__D+"
                i = IC_PGT_D + 1
            Case Is = "___D"
                i = IC_PGT_D
            Case Is = "__D-"
                i = IC_PGT_D - 1
                
            Case Else
                i = ICUNKNOWN
        End Select
        
    End If
    
    If (flags = 2) Or (flags = 18) Then
        i = ICGAVEL
    End If
    
    GetSmallIcon = i
End Function

Public Sub AddName(ByVal Username As String, ByVal Product As String, ByVal flags As Long, ByVal Ping As Long, Optional Clan As String, Optional ForcePosition As Integer)
    Dim i As Integer, LagIcon As Integer, isPriority As Integer
    Dim IsSelf As Boolean
    
    If StrComp(Username, CurrentUsername, vbTextCompare) = 0 Then
        MyFlags = flags
        SharedScriptSupport.BotFlags = MyFlags
        IsSelf = True
    End If
    
    If CheckChannel(Username) > 0 Then Exit Sub
    
    Select Case Ping
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
    
    If (flags And USER_NOUDP) = USER_NOUDP Then LagIcon = LAG_PLUG
    
    isPriority = frmChat.lvChannel.ListItems.Count + 1
    
    i = GetSmallIcon(Product, flags)
    
    'Special Cases
    If i = ICSQUELCH Then
        'Debug.Print "Returned a SQUELCH icon"
        If ForcePosition > 0 Then isPriority = ForcePosition
        
    ElseIf i = ICBLIZZ Or i = ICGAVEL Then
        If ForcePosition = 0 Then isPriority = 1 Else isPriority = ForcePosition
        
    Else
        If ForcePosition > 0 Then isPriority = ForcePosition
        
    End If
    
    If i > frmChat.imlIcons.ListImages.Count Then i = frmChat.imlIcons.ListImages.Count
        
    With frmChat.lvChannel.ListItems
        .Add isPriority, , Username, , i
        .Item(isPriority).ListSubItems.Add , , , LagIcon
        
        If Not BotVars.NoColoring Then
            .Item(isPriority).ForeColor = GetNameColor(flags, 0, IsSelf)
        End If
        
        g_ThisIconCode = -1
    End With
End Sub


Public Function CheckBlock(ByVal Username As String) As Boolean
    Dim s As String, i As Integer
    
    If Dir$(GetFilePath("filters.ini")) <> vbNullString Then
        s = ReadINI("BlockList", "Total", "filters.ini")
        If StrictIsNumeric(s) Then
            i = s
        Else
            Exit Function
        End If
        
        Username = PrepareCheck(Username)
        
        For i = 0 To i
            s = ReadINI("BlockList", "Filter" & i, "filters.ini")
            If Username Like PrepareCheck(s) Then
                CheckBlock = True
                Exit Function
            End If
        Next i
    End If
End Function

Public Function CheckMsg(ByVal Msg As String, Optional ByVal Username As String, Optional ByVal Ping As Long) As Boolean
    Dim i As Integer
    Msg = LCase(Msg)
    
    For i = 0 To UBound(gFilters)
        If Len(gFilters(i)) > 1 Then
            If InStr(gFilters(i), "%") > 0 Then
                If InStr(1, Msg, LCase(DoReplacements(gFilters(i), Username, Ping))) > 0 Then
                    CheckMsg = True
                    Exit Function
                End If
            Else
                If InStr(1, Msg, LCase(gFilters(i))) > 0 Then
                    CheckMsg = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Sub UpdateProfile()
    Dim s As String
    s = GetCurrentSongTitle(True)
    If s = vbNullString Then Exit Sub
    
    SetProfile "", ":[ ProfileAmp ]:" & vbCrLf & "WinAmp is currently playing: " & _
        vbCrLf & s & vbCrLf & "Last updated " & Time & ", " & Format(Date, "d-MM-yyyy") & vbCrLf & CVERSION & " - http://www.stealthbot.net"
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
    GetNewsURL = Chr(Asc("h")) & Chr(Asc("t")) & Chr(Asc("t")) & Chr(Asc("p")) & Chr(Asc(":")) & Chr(Asc("/")) & Chr(Asc("/")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc(".")) & Chr(Asc("s")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("b")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc("/")) & Chr(Asc("g")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc("v")) & Chr(Asc("e")) & Chr(Asc("r")) & Chr(Asc("3")) & Chr(Asc(".")) & Chr(Asc("p")) & Chr(Asc("h")) & Chr(Asc("p")) & Chr(Asc("?")) & Chr(Asc("v")) & Chr(Asc("c")) & Chr(Asc("=")) & VERCODE
End Function

Public Function HTMLToRGBColor(ByVal s As String) As Long
    HTMLToRGBColor = RGB(Val("&H" & Mid$(s, 1, 2)), Val("&H" & Mid$(s, 3, 2)), Val("&H" & Mid$(s, 5, 2)))
End Function

Public Function StrictIsNumeric(ByVal sCheck As String) As Boolean
    Dim i As Long
    
    StrictIsNumeric = True
    
    If Len(sCheck) > 0 Then
        For i = 1 To Len(sCheck)
            If Not ((Asc(Mid$(sCheck, i, 1)) >= 48) And (Asc(Mid$(sCheck, i, 1)) <= 57)) Then
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
    Count = Val(ReadCFG("StoredKeys", "Count"))
    
    Dim sKey As String
    
    If Count > 0 Then
        For Count = 1 To Count
            sKey = ReadCFG("StoredKeys", "Key" & Count)
            If Len(sKey) > 0 Then cboCDKey.AddItem sKey
        Next Count
    End If
    
'    cboCDKey.ListIndex = 0
End Sub

Public Sub WriteCDKeys(ByRef cboCDKey As ComboBox)
    Dim i As Integer
    
    WriteINI "StoredKeys", "Count", cboCDKey.ListCount + 1
    
    For i = 0 To cboCDKey.ListCount
        If Len(cboCDKey.List(i)) > 0 Then WriteINI "StoredKeys", "Key" & (i + 1), cboCDKey.List(i)
    Next i
End Sub

Public Sub GetCountryData(ByRef CountryAbbrev As String, ByRef CountryName As String)
    Dim sBuf As String
    Const LOCALE_USER_DEFAULT = &H400
    Const LOCALE_SABBREVCTRYNAME As Long = &H7 'abbreviated country name
    Const LOCALE_SENGCOUNTRY As Long = &H1002  'English name of country
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, sBuf, Len(sBuf))
    CountryAbbrev = KillNull(sBuf)
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, sBuf, Len(sBuf))
    CountryName = KillNull(sBuf)
End Sub

Public Function GetTimeZoneBias() As Long
    Dim TZinfo As TIME_ZONE_INFORMATION
    Dim lngL As Long
    
    lngL = GetTimeZoneInformation(TZinfo)

    Select Case lngL
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
    If lSocketHandle > 0 Then
        If bEnabled Then
            Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_ON, NAGLE_OPTLEN)
        Else
            Call SetSockOpt(lSocketHandle, IPPROTO_TCP, TCP_NODELAY, NAGLE_OFF, NAGLE_OPTLEN)
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
    If MonitorExists Then frmChat.DeconstructMonitor
    Set MonitorForm = New frmMonitor
End Sub

Public Function ProductCodeToFullName(ByVal pCode As String) As String
    Select Case pCode
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
        Case Else: ProductCodeToFullName = "an unknown or non-standard product"
    End Select
End Function

' Assumes that sIn has length >=1
Public Function PercentActualUppercase(ByVal sIn As String) As Double
    Dim UppercaseChars As Integer, i As Integer
    
    sIn = Replace(sIn, " ", "")
    
    For i = 1 To Len(sIn)
        If IsAlpha(Asc(Mid$(sIn, i, 1))) Then
            If IsUppercase(Asc(Mid$(sIn, i, 1))) Then
                UppercaseChars = UppercaseChars + 1
            End If
        End If
    Next i
    
    PercentActualUppercase = CDbl(100 * (UppercaseChars / Len(sIn)))
End Function

Public Function MyUCase(ByVal sIn As String) As String
    Dim i As Integer
    Dim CurrentByte As Byte

    If LenB(sIn) > 0 Then
        For i = 1 To Len(sIn)
            CurrentByte = Asc(Mid$(sIn, i, 1))
            If IsAlpha(CurrentByte) Then
                If Not IsUppercase(CurrentByte) Then
                    Mid$(sIn, i, 1) = Chr(CurrentByte - 32)
                End If
            End If
        Next i
    End If

    MyUCase = sIn
End Function

Public Function IsAlpha(ByVal bCharValue As Byte) As Boolean
    IsAlpha = ((bCharValue >= 65 And bCharValue <= 90) Or (bCharValue >= 97 And bCharValue <= 122))
End Function

Public Function IsNumber(ByVal bCharValue As Byte) As Boolean
    IsNumber = ((bCharValue >= 48 And bCharValue <= 57))
End Function

Public Function IsUppercase(ByVal bCharValue As Byte) As Boolean
    IsUppercase = (bCharValue >= 65 And bCharValue <= 90)
End Function

Public Function VBHexToHTMLHex(ByVal sIn As String) As String
    sIn = Left$((sIn) & "000000", 6)
    
    VBHexToHTMLHex = Mid$(sIn, 5, 2) & Mid$(sIn, 3, 2) & Mid$(sIn, 1, 2)
End Function


Public Sub GetW3LadderProfile(ByVal sPlayer As String, ByVal eType As enuWebProfileTypes)
    If LenB(sPlayer) > 0 Then
        ShellExecute frmChat.hWnd, "Open", "http://www.battle.net/war3/ladder/" & IIf(eType = W3XP, "w3xp", "war3") & "-player-profile.aspx?Gateway=" & GetW3Realm(sPlayer) & "&PlayerName=" & NameWithoutRealm(sPlayer), 0&, 0&, 0&
    End If
End Sub

Public Sub DoLastSeen(ByVal Username As String)
    Dim i As Integer
    Dim found As Boolean
    
    If colLastSeen.Count > 0 Then
        For i = 1 To colLastSeen.Count
            If StrComp(colLastSeen.Item(i), Username, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next i
    End If
    
    If Not found Then
        colLastSeen.Add Username
        If colLastSeen.Count > 15 Then
            colLastSeen.Remove 1
        End If
    End If
End Sub

Public Sub SetTitle(ByVal sTitle As String)
    frmChat.Caption = "[" & sTitle & "]" & " - " & CVERSION
End Sub

Public Function NameWithoutRealm(ByVal Username As String, Optional ByVal Strict As Byte = 0) As String
    If IsW3 And Strict = 0 Then
        NameWithoutRealm = Username
    Else
        If InStr(1, Username, "@") > 0 Then
            NameWithoutRealm = Left$(Username, InStr(1, Username, "@") - 1)
        Else
            NameWithoutRealm = Username
        End If
    End If
End Function

Public Function GetW3Realm(Optional ByVal Username As String) As String
    If LenB(Username) = 0 Then
        GetW3Realm = w3Realm
    Else
        If InStr(1, Username, "@") > 0 Then
            GetW3Realm = Mid$(Username, InStr(1, Username, "@") + 1)
        Else
            GetW3Realm = w3Realm
        End If
    End If
End Function

Public Function GetConfigFilePath() As String
    Static FilePath As String
    
    If LenB(FilePath) = 0 Then
        If (LenB(ConfigOverride) > 0) Then
            FilePath = ConfigOverride
        Else
            FilePath = GetProfilePath()
            FilePath = FilePath & IIf(Right$(FilePath, 1) = "\", "", "\") & "config.ini"
        End If
    End If
    
    If InStr(FilePath, "\") = 0 Then
        FilePath = App.Path & "\" & FilePath
    End If
    
    GetConfigFilePath = FilePath
End Function

Public Function GetFilePath(ByVal Filename As String) As String
    Dim s As String
    
    If InStr(Filename, "\") = 0 Then
        GetFilePath = GetProfilePath() & "\" & Filename
        
        s = ReadCFG("FilePaths", Filename)
        
        If LenB(s) > 0 Then
            If LenB(Dir$(s)) Then
                GetFilePath = s
            End If
        End If
    Else
        GetFilePath = Filename
    End If
End Function


Public Function OKToDoAutocompletion(ByRef sText As String, ByVal KeyAscii As Integer) As Boolean
    If BotVars.NoAutocompletion Then
        OKToDoAutocompletion = False
    Else
        If (InStr(sText, " ") = 0) And KeyAscii <> 32 Then       ' one word only
            OKToDoAutocompletion = True
        Else                                                            ' > 1 words
            If StrComp(Left$(sText, 1), "/") = 0 And KeyAscii <> 32 Then ' left character is a /
                
                If InStr(InStr(sText, " ") + 1, sText, " ") = 0 Then    ' only two words
                    If StrComp(Left$(sText, 3), "/m ") = 0 Or _
                        StrComp(Left$(sText, 3), "/w ") = 0 Then
                        
                            OKToDoAutocompletion = True
                    Else
                        OKToDoAutocompletion = False
                    End If
                Else                                                    ' more than two words
                    OKToDoAutocompletion = False
                End If
            Else
                OKToDoAutocompletion = False
            End If
        End If
    End If
End Function

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
    Dim s As String
    Static LastPath As String

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
    Dim i As Integer, CompareLen As Integer, QueueUBound As Integer, NumRemoved As Integer
    
    Dim Iterations As Integer

    sUser = "/ban " & LCase(sUser) & " "
    CompareLen = Len(sUser)
    QueueUBound = colQueue.Count
    
    i = 1
    
    While QueueUBound > 0 And i < (QueueUBound - NumRemoved)
        With colQueue.Item(i)
            If Len(.Message) >= CompareLen Then
            
                If StrComp(Left$(LCase(.Message), CompareLen), sUser, vbBinaryCompare) = 0 Then
                    On Error GoTo ArrayIsLocked
                
                    If i < colQueue.Count Then
'                        For c = i + 1 To UBound(Queue)
'                            Queue(c - 1) = Queue(c)
'                        Next c
                        colQueue.Remove i
                    End If
                    
                    'ReDim Preserve Queue(UBound(Queue) - 1)
                    
                    NumRemoved = NumRemoved + 1
                    
ArrayIsLocked:
                    Debug.Print "Error " & Err.Number & ": " & Err.Description
                    Debug.Print "Array was locked in RemoveBanFromQueue() with user " & sUser & "!"
                    Resume Next
                End If
                
            End If
            
            Iterations = Iterations + 1
            
            If Iterations > 10000 Then
                If MDebug("debug") Then
                    frmChat.AddChat vbRed, "Warning: Loop size limit exceeded in RemoveBanFromQueue()!"
                End If
                
                Exit Sub
            End If
            
        End With
        
        i = i + 1
    Wend
End Sub


Public Function AllowedToTalk(ByVal sUser As String, ByVal Msg As String) As Boolean
    Dim i As Integer
    Dim CurrentGTC As Long
    
    AllowedToTalk = True    ' Default to true

    i = UsernameToIndex(sUser)
    CurrentGTC = GetTickCount()
    
    'For each condition where the user is NOT allowed to talk, set to false
    
    If i > 0 Then
        With colUsersInChannel.Item(i)
            If CurrentGTC - .JoinTime < BotVars.AutofilterMS Then
                AllowedToTalk = False
            End If
        End With
    End If
    
    If Filters Then
        If CheckBlock(sUser) Or CheckMsg(Msg, sUser, -5) Then
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

Public Sub AddBannedUser(ByVal sUser As String)
    Dim i As Integer
    
    sUser = LCase(sUser)
    
    If UBound(gBans) > 10000 Then
        ReDim gBans(0)
    End If

    For i = 0 To UBound(gBans)
        If StrComp(gBans(i).UsernameActual, sUser) = 0 Then
            Exit Sub
        End If
    Next i
    
    ReDim Preserve gBans(0 To UBound(gBans) + 1)
    
    gBans(UBound(gBans)).UsernameActual = sUser
    gBans(UBound(gBans)).Username = StripRealm(sUser)
End Sub

Public Sub UnbanBannedUser(ByVal sUser As String)
    ' collapse array on top of the removed user
    Dim i As Integer, c As Integer, NumRemoved As Integer, Iterations As Long
    Dim uBnd As Integer
    
    sUser = LCase(StripRealm(sUser))
    uBnd = UBound(gBans)
    
    While i <= (uBnd - NumRemoved)
        If StrComp(sUser, gBans(i).Username) = 0 Then
            If i <> UBound(gBans) Then
                For c = i To UBound(gBans)
                    gBans(i) = gBans(i + 1)
                Next c
            End If
            
            ReDim Preserve gBans(UBound(gBans) - 1)
            
            NumRemoved = NumRemoved + 1
        Else
            i = i + 1
        End If
        
        Iterations = Iterations + 1
        
        If Iterations > 9000 Then
            If MDebug("debug") Then
                frmChat.AddChat vbRed, "Warning: Loop size limit exceeded in UnbanBannedUser()!"
                frmChat.AddChat vbRed, "The banned-user list has been reset.. hope it works!"
            End If
            
            ReDim gBans(0)
            
            Exit Sub
        End If
    Wend
End Sub

Public Function IsBanned(ByVal sUser As String) As Boolean
    Dim i As Integer
    
    sUser = LCase(sUser)
    
    If InStr(sUser, "#") Then
        sUser = Left$(sUser, InStr(sUser, "#") - 1)
        Debug.Print sUser
    End If
    
    For i = 0 To UBound(gBans)
        If StrComp(sUser, gBans(i).UsernameActual) = 0 Then
            IsBanned = True
            Exit Function
        End If
    Next i
End Function

Public Function IsValidIPAddress(ByVal sIn As String) As Boolean
    Dim s() As String
    Dim i As Integer
    
    IsValidIPAddress = True
    
    If InStr(sIn, ".") Then
    
        s() = Split(sIn, ".")
        
        If UBound(s) = 3 Then
            For i = 0 To 3
                If Not StrictIsNumeric(s(i)) Then
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
    If IsSelf Then
        'Debug.Print "Assigned color IsSelf"
        GetNameColor = vbWhite
        Exit Function
    End If
    
    '/* Blizzard */
    If (flags And &H1) = &H1 Then
        GetNameColor = COLOR_BLUE
        Exit Function
    End If
    
    '/* Operator */
    If (flags And &H2) = &H2 Then
        'Debug.Print "Assigned color OP"
        GetNameColor = &HDDDDDD
        Exit Function
    End If
    
    '/* Squelched */
    If (flags And &H20) = &H20 Then
        'Debug.Print "Assigned color SQUELCH"
        GetNameColor = &H99
        Exit Function
    End If
    
    '/* Idle */
    If IdleTime > BotVars.SecondsToIdle Then
        'Debug.Print "Assigned color IDLE"
        GetNameColor = &HBBBBBB
        Exit Function
    End If
    
    '/* Default */
    'Debug.Print "Assigned color NORMAL"
    GetNameColor = COLOR_TEAL
End Function

Public Function FlagDescription(ByVal flags As Long) As String
    Dim s0ut As String
    Dim multipleFlags As Boolean
        
    If (flags And &H20) = &H20 Then
        s0ut = "Squelched"
        multipleFlags = True
    End If
    
    If (flags And &H2) = &H2 Then
        If multipleFlags Then
            s0ut = s0ut & ", channel op"
        Else
            s0ut = "Channel op"
        End If
        
        multipleFlags = True
    End If
    
    If ((flags And USER_BLIZZREP) = USER_BLIZZREP) Or ((flags And USER_SYSOP) = USER_SYSOP) Then
        If multipleFlags Then
            s0ut = s0ut & ", Blizzard representative"
        Else
            s0ut = "Blizzard representative"
        End If
        
        multipleFlags = True
    End If
    
    If (flags And &H10) = &H10 Then
        If multipleFlags Then
            s0ut = s0ut & ", UDP plug"
        Else
            s0ut = "UDP plug"
        End If
        
        multipleFlags = True
    End If
    
    If LenB(s0ut) = 0 Then
        If flags = 0 Then
            s0ut = "Normal"
        Else
            s0ut = "Altered"
        End If
    End If
    
    FlagDescription = s0ut & " [" & flags & "]"
End Function

'Returns TRUE if the specified argument was a command line switch,
' such as -debug
Public Function MDebug(ByVal sArg As String) As Boolean
    MDebug = InStr(1, CommandLine, "-" & sArg, vbTextCompare) > 0
End Function

'Returns system uptime in milliseconds
Public Function GetUptimeMS() As Double
    Dim mmt As MMTIME
    Dim lSize As Long
    
    lSize = LenB(mmt)
    
    Call timeGetSystemTime(mmt, lSize)

    GetUptimeMS = LongToUnsigned(mmt.units)
End Function


Public Function UsernameToIndex(ByVal sUsername As String) As Long
    Dim user As clsUserInfo
    Dim FirstLetter As String * 1
    Dim i As Integer
    
    FirstLetter = Mid$(sUsername, 1, 1)
    
    If colUsersInChannel.Count > 0 Then
    
        For i = 1 To colUsersInChannel.Count
            Set user = colUsersInChannel.Item(i)
            
            With user
                If StrComp(Mid$(.Username, 1, 1), FirstLetter, vbTextCompare) = 0 Then
                    If StrComp(sUsername, .Username, vbTextCompare) = 0 Then
                        UsernameToIndex = i
                        Exit Function
                    End If
                End If
            End With
        Next i
        
    End If
    
    UsernameToIndex = 0
End Function


Public Function CheckChannel(ByVal NameToFind As String) As Integer
    Dim itmFound As ListItem
    
    Set itmFound = frmChat.lvChannel.FindItem(NameToFind)
    
    If itmFound Is Nothing Then
        CheckChannel = 0
    Else
        CheckChannel = itmFound.Index
    End If
End Function


Public Sub CheckPhrase(ByRef Username As String, ByRef Msg As String, ByVal mType As Byte)
    Dim i As Integer
    
    If UBound(Catch) = 0 Then
        If Catch(0) = vbNullString Then Exit Sub
    End If
    
    For i = LBound(Catch) To UBound(Catch)
        If Catch(i) <> vbNullString Then
            If InStr(1, LCase(Msg), Catch(i), vbTextCompare) <> 0 Then
                Call CaughtPhrase(Username, Msg, Catch(i), mType)
                Exit Sub
            End If
        End If
    Next i
End Sub


Public Sub CaughtPhrase(ByVal Username As String, ByVal Msg As String, ByVal Phrase As String, ByVal mType As Byte)
    Dim i As Integer, s As String
    i = FreeFile
    
    If LenB(ReadCFG("Other", "FlashOnCatchPhrases")) > 0 Then FlashWindow
    
    Select Case mType
        Case CPTALK: s = "TALK"
        Case CPEMOTE: s = "EMOTE"
        Case CPWHISPER: s = "WHISPER"
    End Select
    
    If Dir$(GetProfilePath() & "\caughtphrases.htm") = vbNullString Then
        Open GetProfilePath() & "\caughtphrases.htm" For Output As #i
        Print #i, "<html>"
        Close #i
    End If
    
    Open GetProfilePath() & "\caughtphrases.htm" For Append As #i
        If LOF(i) > 10000000 Then
            Close #i
            Kill GetProfilePath() & "\caughtphrases.htm"
            Open GetProfilePath() & "\caughtphrases.htm" For Output As #i
        End If
        
        Msg = Replace(Msg, "<", "&lt;")
        Msg = Replace(Msg, ">", "&gt;")
        
        Print #i, "<B>" & Format(Date, "MM-dd-yyyy") & " - " & Time & " - " & s & Space(1) & Username & ": </B>" & Replace(Msg, Phrase, "<i>" & Phrase & "</i>") & "<br>"
    Close #i
End Sub


Public Function DoReplacements(ByVal s As String, Optional Username As String, Optional Ping As Long) As String
    Dim gAcc As udtGetAccessResponse
    
    gAcc = GetAccess(Username)

    s = Replace(s, "%0", Username)
    s = Replace(s, "%1", CurrentUsername)
    s = Replace(s, "%c", gChannel.Current)
    s = Replace(s, "%bc", BanCount)
    
    If (Ping > -2) Then
        s = Replace(s, "%p", Ping)
    End If
    
    s = Replace(s, "%v", CVERSION)
    s = Replace(s, "%a", IIf(gAcc.Access >= 0, gAcc.Access, "0"))
    s = Replace(s, "%f", gAcc.flags)
    s = Replace(s, "%t", Time$)
    s = Replace(s, "%d", Date)
    s = Replace(s, "%m", GetMailCount(Username))
    
    DoReplacements = s
End Function

' Updated 4/10/06 to support millisecond pauses
'  If using milliseconds pause for at least 100ms
Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True, Optional ByVal Milliseconds As Boolean = False)
    Dim i As Integer
    
    If AllowEvents Then
        For i = 0 To (fSeconds * (IIf(Milliseconds, 1, 1000))) \ 100
            'Debug.Print "sleeping 100ms"
            Sleep 100
            DoEvents
        Next i
    Else
        Sleep fSeconds * (IIf(Milliseconds, 1, 1000))
    End If
End Sub


Public Sub LogDBAction(ByVal ActionType As enuDBActions, ByVal Caller As String, ByVal Target As String, ByVal Instruction As String)
    Dim sPath As String
    Dim Action As String
    Dim f As Integer
    
    f = FreeFile
    sPath = GetProfilePath() & "\Logs\database.txt"
    
    If Len(Caller) < 2 Then
        Caller = "bot console"
    End If
    
    If LenB(Dir$(sPath)) = 0 Then
        Open sPath For Output As #f
    Else
        Open sPath For Append As #f
        
        If LOF(f) > BotVars.MaxLogfileSize And BotVars.MaxLogfileSize > 0 Then
            Close #f
            Kill sPath
            Open sPath For Output As #f
            Print #f, "Logfile cleared automatically on " & Format(Now, "HH:MM:SS MM/DD/YY") & "."
        End If
    End If
    
    Select Case ActionType
        Case AddEntry
            Action = "adds"

        Case RemEntry
            Action = "removes"

        Case ModEntry
            Action = "modifies"
    End Select
    
    Action = "[" & Format(Now, "HH:MM:SS MM/DD/YY") & "] " & Caller & " " & Action & Space(1) & Target & ": " & Instruction
    
    Print #f, Action
    Close #f
End Sub

Public Sub LogCommand(ByVal Caller As String, ByVal CString As String)
    Dim sPath As String
    Dim Action As String
    Dim f As Integer
    
    On Error GoTo LogCommand_Error

    If LenB(CString) > 0 Then
        f = FreeFile
        sPath = GetProfilePath() & "\Logs\commands.txt"
        
        If LenB(Caller) = 0 Then
            Caller = "bot console"
        End If
        
        If LenB(Dir$(sPath)) = 0 Then
            Open sPath For Output As #f
        Else
            Open sPath For Append As #f
            
            If LOF(f) > BotVars.MaxLogfileSize And BotVars.MaxLogfileSize > 0 Then
                Close #f
                Kill sPath
                Open sPath For Output As #f
                Print #f, "Logfile cleared automatically on " & Format(Now, "HH:MM:SS MM/DD/YY") & "."
            End If
        End If
        
        Action = "[" & Format(Now, "HH:MM:SS MM/DD/YY") & "][" & Caller & "]-> " & CString
        
        Print #f, Action
        Close #f
    End If

    On Error GoTo 0
    Exit Sub

LogCommand_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure LogCommand of Module modOtherCode"
End Sub

'Pos must be >0
' Returns a single chunk of a string as if that string were Split() and that chunk
' extracted
' 1-based
Public Function GetStringChunk(ByVal str As String, ByVal Pos As Integer)
    'one two three
    '   1   2
    Dim c As Integer, i As Integer, TargetSpace As Integer
    c = 0
    i = 1
    Pos = Pos
    
    ' The string must have at least (pos-1) spaces to be valid
    While (c < Pos) And (i > 0)
        TargetSpace = i
        i = InStr(i + 1, str, " ")
        c = c + 1
    Wend
    
    If c >= Pos Then
        c = InStr(TargetSpace + 1, str, " ") ' check for another space (more afterwards)
        
        If c > 0 Then
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
    If LenB(Product) = 0 Then
        Product = StrReverse(BotVars.Product)
    End If
    
    Select Case Product
        Case "W2BN", "NB2W"
            GetProductKey = "W2"
        Case "STAR", "SEXP", "PXES", "RATS"
            GetProductKey = "SC"
        Case "D2DV", "D2XP", "PX2D", "VD2D"
            GetProductKey = "D2"
        Case "W3XP", "WAR3", "3RAW", "PX3W"
            GetProductKey = "W3"
    End Select
End Function

Public Function InsertDummyQueueEntry()
    frmChat.AddQ "%%%%%blankqueuemessage%%%%%"
End Function
