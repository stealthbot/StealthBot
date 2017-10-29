Attribute VB_Name = "modOtherCode"
Option Explicit
Private Const OBJECT_NAME As String = "modOtherCode"

Public Const AZ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const Num09 As String = "0123456789"

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

Public Function ReadCfg$(ByVal riSection$, ByVal riKey$)
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
            
            ReadCfg = sRiValue
        End If
    Else
        ReadCfg = ""
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

'// http://www.vbforums.com/showpost.php?p=2909252
Public Sub InsertionSort(ByRef pvarArray As Variant)
    Dim i       As Long
    Dim j       As Long
    Dim iMin    As Long
    Dim iMax    As Long
    Dim varSwap As Variant

    iMin = LBound(pvarArray) + 1
    iMax = UBound(pvarArray)
    For i = iMin To iMax
        varSwap = pvarArray(i)
        For j = i To iMin Step -1
            If varSwap < pvarArray(j - 1) Then pvarArray(j) = pvarArray(j - 1) Else Exit For
        Next j
        pvarArray(j) = varSwap
    Next i
End Sub


Public Function GetTimeStamp(Optional DateTime As Date) As String
    If (DateDiff("s", DateTime, "00:00:00 12/30/1899") = 0) Then
        DateTime = Now
    End If

    Select Case (BotVars.TSSetting)
        Case 0
            GetTimeStamp = _
                " [" & Format(DateTime, "HH:MM:SS AM/PM") & "] "
            
        Case 1
            GetTimeStamp = _
                " [" & Format(DateTime, "HH:MM:SS") & "] "
        
        Case 2
            GetTimeStamp = _
                " [" & Format(DateTime, "HH:MM:SS") & "." & Right$("000" & GetCurrentMS, 3) & "] "
        
        Case 3
            GetTimeStamp = vbNullString
        
        Case Else
            GetTimeStamp = _
                " [" & Format(DateTime, "HH:MM:SS AM/PM") & "] "
    End Select
End Function

Public Function GetVerByte(Product As String, Optional ByVal UseHardcode As Integer) As Long
    Dim Key As String
    
    Key = GetProductKey(Product)
    
    If ((Config.GetVersionByte(Key) = -1) Or _
        (UseHardcode = 1)) Then
        
        GetVerByte = GetProductInfo(Product).VersionByte
    Else
        GetVerByte = Config.GetVersionByte(Key)
    End If
    
End Function

Public Function GetGamePath(ByVal Client As String) As String
    ' [override] XXHashes= functionality replaced by checkrevision.ini -> [CRev_XX] Path=
    ' removed ~Ribose/2010-08-12
    Dim CRevINIPath As String
    Dim Key As String
    Dim Path As String
    Dim sep1 As String
    Dim sep2 As String
    
    ' Moved CheckRevision.ini to profile directory instead of install directory. -Pyro, 2016-03-21
    'CRevINIPath = GetFilePath(FILE_CREV_INI, StringFormat("{0}\", App.Path))
    CRevINIPath = GetFilePath(FILE_CREV_INI)
    
    Key = GetProductKey(Client)
    Path = ReadINI$(StringFormat("CRev_{0}", Key), "Path", CRevINIPath)
    sep1 = vbNullString
    sep2 = vbNullString
    
    If (InStr(1, Path, ":\") = 0) Then
        If (Left$(Path, 1) <> "\") Then sep1 = "\"
        If (Right$(Path, 1) <> "\") Then sep2 = "\"
        Path = StringFormat("{0}{1}{2}{3}", App.Path, sep1, Path, sep2)
    End If
    
    GetGamePath = Path
End Function

Function MKL(Value As Long) As String
    Dim Result As String * 4
    
    Call CopyMemory(ByVal Result, Value, 4)
    
    MKL = Result
End Function

Function MKI(Value As Integer) As String
    Dim Result As String * 2
    
    Call CopyMemory(ByVal Result, Value, 2)
    
    MKI = Result
End Function

Public Function CheckPath(ByVal sPath As String) As Long
    If (LenB(Dir$(sPath)) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[HASHES] " & _
            Mid$(sPath, InStrRev(sPath, "\") + 1) & " is missing."
            
        CheckPath = 1
    End If
End Function

Public Function Ban(ByVal Inpt As String, SpeakerAccess As Integer, Optional Kick As Integer) As String
    On Error GoTo ERROR_HANDLER

    Static LastBan      As String
    
    Dim Username        As String
    Dim CleanedUsername As String
    Dim i               As Integer
    Dim pos             As Integer
    
    If (LenB(Inpt) > 0) Then
        If (Kick > 2) Then
            LastBan = vbNullString
            
            Exit Function
        End If
        
        If (g_Channel.Self.IsOperator) Then
            If (InStr(1, Inpt, Space$(1), vbBinaryCompare) <> 0) Then
                Username = LCase$(Left$(Inpt, InStr(1, Inpt, Space(1), _
                    vbBinaryCompare) - 1))
            Else
                Username = LCase$(Inpt)
            End If
            
            If (LenB(Username) > 0) Then
                LastBan = LCase$(Username)
                
                'CleanedUsername = StripRealm(CleanedUsername)
                CleanedUsername = StripInvalidNameChars(Username)

                If (SpeakerAccess < 200) Then
                    If ((GetSafelist(CleanedUsername)) Or (GetSafelist(Username))) Then
                        Ban = "Error: That user is safelisted."
                        
                        Exit Function
                    End If
                End If
                
                If (Database.GetUserAccess(Username).Rank >= SpeakerAccess) Then
                    Ban = "Error: You do not have sufficient access to do that."
                    
                    Exit Function
                End If
                
                pos = g_Channel.GetUserIndex(Username)
                
                If (pos > 0) Then
                    If (g_Channel.Users(pos).IsOperator) Then
                        Ban = "Error: You cannot ban a channel operator."
                    
                        Exit Function
                    End If
                End If
                
                If (Kick = 0) Then
                    Call frmChat.AddQ("/ban " & Inpt)
                Else
                    Call frmChat.AddQ("/kick " & Inpt)
                End If
            End If
        Else
            Ban = "Error: The bot does not have ops."
        End If
    End If
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Ban()."

    Exit Function
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
                        NewUsername = Replace(NewUsername, Chr(thisChar), vbNullString)
                    End If
                End If
            End If
        Next i
        
        StripInvalidNameChars = NewUsername
    End If
End Function

'// Utility Function for joining strings
'// EXAMPLE
'// StringFormat("This is an {1} of its {0}.", Array("use", "example")) '// OUTPUT: This is an example of its use.
'// 08/29/2008 JSM - Created
Public Function StringFormatA(Source As String, params() As Variant) As String
    
    On Error GoTo ERROR_HANDLER:

    Dim retVal As String, i As Integer
    retVal = Source
    For i = LBound(params) To UBound(params)
        retVal = Replace(retVal, "{" & i & "}", CStr(params(i)))
    Next
    StringFormatA = retVal
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in StringFormatA().")

    StringFormatA = vbNullString
    
End Function


'// Utility Function for joining strings
'// EXAMPLE
'// StringFormat("This is an {1} of its {0}.", "use", "example") '// OUTPUT: This is an example of its use.
'// 08/14/2009 JSM - Created
Public Function StringFormat(Source As String, ParamArray params() As Variant)

    On Error GoTo ERROR_HANDLER:

    Dim retVal As String, i As Integer
    retVal = Source
    For i = LBound(params) To UBound(params)
        retVal = Replace(retVal, "{" & i & "}", CStr(params(i)))
    Next
    StringFormat = retVal
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in StringFormat().")


    StringFormat = vbNullString
    
End Function

Public Function StripAccountNumber(ByVal Username As String) As String
    Dim numpos As Integer
    Dim atpos As Integer
    
    numpos = InStr(1, Username, "#", vbBinaryCompare)
    If numpos > 0 Then
        atpos = InStr(numpos, Username, Config.GatewayDelimiter, vbBinaryCompare)
        If atpos > 0 Then
            StripAccountNumber = Left$(Username, numpos - 1) & Mid$(Username, atpos)
        Else
            StripAccountNumber = Left$(Username, numpos - 1)
        End If
    Else
        StripAccountNumber = Username
    End If
End Function

Public Function StripRealm(ByVal Username As String) As String
    If (InStr(1, Username, Config.GatewayDelimiter, vbBinaryCompare) > 0) Then
        Username = Replace(Username, Config.GatewayDelimiter & "USWest", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "USEast", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Asia", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Euruope", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Beta", vbNullString, 1)
        
        Username = Replace(Username, Config.GatewayDelimiter & "Lordaeron", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Azeroth", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Kalimdor", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Northrend", vbNullString, 1)
        Username = Replace(Username, Config.GatewayDelimiter & "Westfall", vbNullString, 1)
        
        Username = Replace(Username, Config.GatewayDelimiter & "Blizzard", vbNullString, 1)
    End If
    
    StripRealm = Username
End Function

Public Sub bnetSend(ByVal Message As String, Optional ByVal Tag As String = vbNullString, Optional ByVal _
    ID As Double = 0)
    Dim pBuf As clsDataBuffer

    On Error GoTo ERROR_HANDLER

    If (frmChat.sckBNet.State = sckConnected) Then
        Set pBuf = New clsDataBuffer
        With pBuf
            If (frmChat.mnuUTF8.Checked) Then
                .InsertNTString Message, UTF8
            Else
                .InsertNTString Message
            End If

            .SendPacket SID_CHATCOMMAND
            .Clear

            If (Tag = "request_receipt") Then
                g_request_receipt = True

                .SendPacket SID_FRIENDSLIST
            End If
        End With
        Set pBuf = Nothing
    End If
    
    If (Left$(Message, 1) <> "/") Then
        If (g_Channel.IsSilent) Then
            frmChat.AddChat RTBColors.Carats, "<", RTBColors.TalkBotUsername, GetCurrentUsername, _
                RTBColors.Carats, "> ", RTBColors.WhisperText, Message
        Else
            frmChat.AddChat RTBColors.Carats, "<", RTBColors.TalkBotUsername, GetCurrentUsername, _
                RTBColors.Carats, "> ", RTBColors.TalkNormalText, Message
        End If
    End If

    On Error Resume Next
        
    RunInAll "Event_MessageSent", ID, Message, Tag
    
    Exit Sub

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.Description & " in bnetSend().")

    Exit Sub
    
End Sub

Public Function Voting(ByVal Mode1 As Byte, Optional Mode2 As Byte, Optional Username As String) As String
On Error GoTo ERROR_HANDLER:
    Static Voted()  As String
    Static VotesYes As Integer
    Static VotesNo  As Integer
    Static VoteMode As Byte
    Static Target   As String
        
    Dim i             As Integer
    
    Select Case (Mode1)
        Case BVT_VOTE_ADD
            For i = LBound(Voted) To UBound(Voted)
                If (StrComp(Voted(i), LCase$(Username), vbTextCompare) = 0) Then
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
                            Voting = Ban(Target & " Banned by vote", VoteInitiator.Rank)
                        Else
                            Voting = "Ban vote failed."
                        End If
                        
                    Case BVT_VOTE_KICK
                        If (VotesYes > VotesNo) Then
                            Voting = Ban(Target & " Kicked by vote", VoteInitiator.Rank, 1)
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
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Voting()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Sub RequestSystemKeys(Optional ByVal eType As enuServerRequestHandlerType = reqInternal, Optional oCommand As clsCommandObj)
    Dim aKeys(3) As String
    aKeys(0) = "System\Account Created"
    aKeys(1) = "System\Last Logon"
    aKeys(2) = "System\Last Logoff"
    aKeys(3) = "System\Time Logged"

    RequestUserData BotVars.Username, aKeys, eType, oCommand
End Sub

'// parses a system time and returns in the format:
'//     mm/dd/yy, hh:mm:ss
'//
Public Function SystemTimeToString(ByRef st As SYSTEMTIME) As String
    Dim buf As String

    With st
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
    Dim st As SYSTEMTIME
    GetLocalTime st
    
    GetCurrentMS = Right$("000" & st.wMilliseconds, 3)
End Function

Public Function ZeroOffset(ByVal lInpt As Long, ByVal lDigits As Long) As String
    Dim sOut As String
    
    sOut = Hex(lInpt)
    ZeroOffset = Right$(String(lDigits, "0") & sOut, lDigits)
End Function

Public Function ZeroOffsetEx(ByVal lInpt As Long, ByVal lDigits As Long) As String
    ZeroOffsetEx = Right$(String(lDigits, "0") & lInpt, lDigits)
End Function


Public Function CheckBlock(ByVal Username As String) As Boolean
    Dim s As String
    Dim i As Integer
    
    For i = 0 To UBound(g_Blocklist)
        If Len(g_Blocklist(i)) > 0 Then
            If (PrepareCheck(Username) Like PrepareCheck(g_Blocklist(i))) Then
                CheckBlock = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function CheckMsg(ByVal Msg As String, Optional ByVal Username As String, Optional ByVal Ping As _
        Long) As Boolean
    
    Dim i As Integer
    
    Msg = PrepareCheck(Msg)
    
    For i = 0 To UBound(gFilters)
        If (Len(gFilters(i)) > 0) Then
            If (InStr(1, gFilters(i), "%", vbBinaryCompare) > 0) Then
                Msg = PrepareCheck(DoReplacements(gFilters(i), Username, Ping))
            End If
            
            If (Msg Like "*" & gFilters(i) & "*") Then
                CheckMsg = True
                   
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub UpdateProfile()
    Dim s As String
    
    s = MediaPlayer.TrackName
    
    If (s = vbNullString) Then
        SetProfile "", ":[ ProfileAmp ]:" & vbCrLf & MediaPlayer.Name & " is not currently playing " & _
                vbCrLf & "Last updated " & Time & ", " & Format(Date, "d-MM-yyyy") & vbCrLf & _
                    CVERSION & " - http://www.stealthbot.net"
    Else
        SetProfile "", ":[ ProfileAmp ]:" & vbCrLf & MediaPlayer.Name & " is currently playing: " & _
                vbCrLf & s & vbCrLf & "Last updated " & Time & ", " & Format(Date, "d-MM-yyyy") & _
                    vbCrLf & CVERSION & " - http://www.stealthbot.net"
    End If
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

Public Sub ReadyInet()
    frmChat.Inet.Cancel
End Sub

Public Function HTMLToRGBColor(ByVal s As String) As Long
    HTMLToRGBColor = RGB(Val("&H" & Mid$(s, 1, 2)), Val("&H" & Mid$(s, 3, 2)), _
        Val("&H" & Mid$(s, 5, 2)))
End Function

Public Function StrictIsNumeric(ByVal sCheck As String, Optional AllowNegatives As Boolean = False) As Boolean
    Dim i As Long
    
    StrictIsNumeric = True
    
    If (Len(sCheck) > 0) Then
        For i = 1 To Len(sCheck)
            If ((Asc(Mid$(sCheck, i, 1)) = 45)) Then
                If ((Not AllowNegatives) Or (i > 1)) Then
                    StrictIsNumeric = False
                    Exit Function
                End If
            ElseIf (Not ((Asc(Mid$(sCheck, i, 1)) >= 48) And _
                     (Asc(Mid$(sCheck, i, 1)) <= 57))) Then
                
                StrictIsNumeric = False
                
                Exit Function
            End If
        Next i
    Else
        StrictIsNumeric = False
    End If
End Function

Public Sub GetCountryData(ByRef CountryAbbrev As String, ByRef CountryName As String, ByRef sCountryCode As String)
    Const LOCALE_USER_DEFAULT    As Long = &H400
    Const LOCALE_ICOUNTRY        As Long = &H5     'Country Code
    Const LOCALE_SABBREVCTRYNAME As Long = &H7     'abbreviated country name
    Const LOCALE_SENGCOUNTRY     As Long = &H1002  'English name of country
    
    Dim sBuf As String
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, sBuf, Len(sBuf))
    CountryAbbrev = KillNull(sBuf)
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, sBuf, Len(sBuf))
    CountryName = KillNull(sBuf)
    
    sBuf = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICOUNTRY, sBuf, Len(sBuf))
    sCountryCode = KillNull(sBuf)
End Sub

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

Public Function ProductCodeToFullName(ByVal pCode As String) As String
    ProductCodeToFullName = GetProductInfo(pCode).FullName
End Function

' Assumes that sIn has Length >=1
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

'//10-15-2009 - Hdx - Updated url to new address
Public Sub GetW3LadderProfile(ByVal sPlayer As String, ByVal eType As enuWebProfileTypes)
    Const W3LadderURLFormat As String = "http://{0}.battle.net/war3/ladder/{1}-player-profile.aspx?Gateway={2}&PlayerName={3}"
    Dim W3LadderURL As String
    Dim W3WebProfileType As String
    Dim W3Realm As String
    Dim W3Domain As String
    W3Domain = "classic"
    
    If (LenB(sPlayer) > 0) Then
        W3WebProfileType = IIf(eType = W3XP, PRODUCT_W3XP, PRODUCT_WAR3)
        W3Realm = GetW3Realm(sPlayer)
        If W3Realm = "Kalimdor" Then W3Domain = "asialadders"
        W3LadderURL = StringFormat(W3LadderURLFormat, W3Domain, LCase$(W3WebProfileType), W3Realm, NameWithoutRealm(sPlayer, 1))
        
        ShellOpenURL W3LadderURL, sPlayer & "'s " & UCase$(W3WebProfileType) & " ladder profile"
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
    If ((frmChat.ClanHandler.IsW3) And (Strict = 0)) Then
        NameWithoutRealm = Username
    Else
        If (InStr(1, Username, Config.GatewayDelimiter, vbBinaryCompare) > 0) Then
            NameWithoutRealm = Left$(Username, InStr(1, Username, Config.GatewayDelimiter) - 1)
        Else
            NameWithoutRealm = Username
        End If
    End If
End Function

Public Function GetCurrentUsername() As String

    GetCurrentUsername = ConvertUsername(CurrentUsername)

End Function

Public Function GetW3Realm(Optional ByVal Username As String) As String
    If (LenB(Username) = 0) Then
        GetW3Realm = BotVars.Gateway
    Else
        If (InStr(1, Username, Config.GatewayDelimiter, vbBinaryCompare) > 0) Then
            GetW3Realm = Mid$(Username, InStr(1, Username, Config.GatewayDelimiter, _
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
            FilePath = StringFormat("{0}Config.ini", GetProfilePath())
        End If
    End If
    
    If (InStr(1, FilePath, "\", vbBinaryCompare) = 0) Then
        FilePath = StringFormat("{0}\{1}", CurDir$(), FilePath)
    End If
    
    GetConfigFilePath = FilePath
End Function

Public Function GetFilePath(ByVal FileName As String, Optional DefaultPath As String = vbNullString) As String
    Dim s As String
    
    If (InStr(FileName, "\") = 0) Then
        If (LenB(DefaultPath) = 0) Then
            GetFilePath = StringFormat("{0}{1}", GetProfilePath(), FileName)
        Else
            GetFilePath = StringFormat("{0}{1}", DefaultPath, FileName)
        End If
        
        s = Config.GetFilePath(FileName)
        
        If (LenB(s) > 0) Then
            If (LenB(Dir$(s))) Then
                GetFilePath = s
            End If
        End If
    Else
        GetFilePath = FileName
    End If
End Function

Public Function GetFolderPath(ByVal sFolderName As String) As String
On Error GoTo ERROR_HANDLER:
    Dim sPath As String
    sPath = GetFilePath(sFolderName)
    If (Not Right$(sPath, 1) = "\") Then sPath = sPath & "\"
    GetFolderPath = sPath
    
    Exit Function
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.{3}()", Err.Number, Err.Description, "modOtherCode", "GetFolderPath")
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
            GetProfilePath = StringFormat("{0}\", CurDir$())
        End If
'    End If
    
    LastPath = GetProfilePath
End Function

Public Sub OpenReadme()
    ShellOpenURL "http://www.stealthbot.net/wiki/Main_Page", "the StealthBot Wiki"
End Sub

Sub ShellOpenURL(ByVal FullURL As String, Optional ByVal Description As String = vbNullString, Optional ByVal DisplayMessage As Boolean = True, Optional ByVal Verb As String = "open")
    ShellExecute frmChat.hWnd, Verb, FullURL, vbNullString, vbNullString, vbNormalFocus
    
    If DisplayMessage Then
        If LenB(Description) > 0 Then Description = Description & " at "
        frmChat.AddChat RTBColors.ConsoleText, "Opening " & Description & "[ " & FullURL & " ]..."
    End If
End Sub

'Checks the queue for duplicate bans
Public Sub RemoveBanFromQueue(ByVal sUser As String)
    Dim tmp As String

    tmp = "/ban " & sUser
        
    g_Queue.RemoveLines tmp & "*"

    If ((StrReverse$(BotVars.Product) = PRODUCT_WAR3) Or _
        (StrReverse$(BotVars.Product) = PRODUCT_W3XP)) Then
        
        Dim strGateway As String
        
        Select Case (BotVars.Gateway)
            Case "Lordaeron": strGateway = Config.GatewayDelimiter & "USWest"
            Case "Azeroth":   strGateway = Config.GatewayDelimiter & "USEast"
            Case "Kalimdor":  strGateway = Config.GatewayDelimiter & "Asia"
            Case "Northrend": strGateway = Config.GatewayDelimiter & "Europe"
        End Select
        
        If (InStr(1, tmp, strGateway, vbTextCompare) = 0) Then
            g_Queue.RemoveLines tmp & strGateway & "*"
        End If
    End If
    
    'frmChat.AddChat vbRed, tmp & "*" & " : " & tmp & strGateway & "*"
End Sub

Public Function AllowedToTalk(ByVal sUser As String, ByVal Msg As String) As Boolean
    Dim i As Integer
    
    ' default to true
    AllowedToTalk = True
    
    If (Filters) Then
        If ((CheckBlock(sUser)) Or (CheckMsg(Msg, sUser, -5))) Then
            AllowedToTalk = False
        End If
    End If
End Function


' Used by the Individual Whisper Window system to determine whether a message should be
'  forwarded to an IWW
Public Function IrrelevantWhisper(ByVal sIn As String, ByVal sUser As String) As Boolean
    IrrelevantWhisper = False
    
    If InStr(sIn, Chr(223) & Chr(126) & Chr(223)) Then
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
    
    'i = UsernameToIndex(sUser)
    
    If i > 0 Then
        'colUsersInChannel.Item(i).Safelisted = bStatus
    End If
End Sub

Public Sub AddBanlistUser(ByVal sUser As String, ByVal cOperator As String)
    Const MAX_BAN_COUNT As Integer = 80

    Dim i      As Integer
    Dim bCount As Integer
    
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
Public Sub UnbanBanlistUser(ByVal sUser As String, ByVal cOperator As String)
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
                frmChat.AddChat RTBColors.ErrorMessageText, "Warning! Loop size limit exceeded " & _
                    "in UnbanBanlistUser()!"
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

'Returns TRUE if the specified argument was a command line switch,
' such as -debug
Public Function MDebug(ByVal sArg As String) As Boolean
    MDebug = InStr(1, CommandLine, StringFormat("-{0} ", sArg), vbTextCompare) > 0
End Function

Public Function SetCommandLine(sCommandLine As String)
On Error GoTo ERROR_HANDLER:
    Dim sTemp    As String
    Dim sSetting As String
    Dim sValue   As String
    Dim sRet     As String
    CommandLine = vbNullString
    sTemp = sCommandLine
    
    Do While Left$(Trim$(sTemp), 1) = "-"
        sTemp = Trim$(sTemp)
        sSetting = Split(Mid$(sTemp, 2) & Space$(1), Space$(1))(0)
        sTemp = Mid$(sTemp, Len(sSetting) + 3)
        Select Case LCase$(sSetting)
            Case "ppath":
                If (Left$(sTemp, 1) = Chr$(34)) Then
                    If (InStr(2, sTemp, Chr$(34), vbTextCompare) > 0) Then
                        sValue = Mid$(sTemp, 2, InStr(2, sTemp, Chr$(34), vbTextCompare) - 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 4)
                    Else
                        sValue = Mid$(Split(sTemp & " -", " -")(0), 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 3)
                    End If
                Else
                    sValue = Split(sTemp & " -", " -")(0)
                    sTemp = Mid$(sTemp, Len(sValue) + 2)
                End If
                
                If (LenB(sValue) > 0) Then
                    If Dir$(sValue) <> vbNullString Then
                        ChDir sValue
                        CommandLine = StringFormat("{0}-ppath {1}{2}{1} ", CommandLine, Chr$(34), sValue)
                    End If
                End If
        
            Case "cpath":
                If (Left$(sTemp, 1) = Chr$(34)) Then
                    If (InStr(2, sTemp, Chr$(34), vbTextCompare) > 0) Then
                        sValue = Mid$(sTemp, 2, InStr(2, sTemp, Chr$(34), vbTextCompare) - 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 4)
                    Else
                        sValue = Mid$(Split(sTemp & " -", " -")(0), 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 3)
                    End If
                Else
                    sValue = Split(sTemp & " -", " -")(0)
                    sTemp = Mid$(sTemp, Len(sValue) + 2)
                End If
                
                If (LenB(sValue) > 0) Then
                    ConfigOverride = sValue
                    If (LenB(GetConfigFilePath()) = 0) Then
                        ConfigOverride = vbNullString
                    Else
                        CommandLine = StringFormat("{0}-cpath {1}{2}{1} ", CommandLine, Chr$(34), sValue)
                    End If
                End If
                
            Case "addpath":
                If (Left$(sTemp, 1) = Chr$(34)) Then
                    If (InStr(2, sTemp, Chr$(34), vbTextCompare) > 0) Then
                        sValue = Mid$(sTemp, 2, InStr(2, sTemp, Chr$(34), vbTextCompare) - 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 4)
                    Else
                        sValue = Mid$(Split(sTemp & " -", " -")(0), 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 3)
                    End If
                Else
                    sValue = Split(sTemp & " -", " -")(0)
                    sTemp = Mid$(sTemp, Len(sValue) + 2)
                End If
                
                AddEnvPath sValue
                
                CommandLine = StringFormat("{0}-addpath {1}{2}{1} ", CommandLine, Chr$(34), sValue)
                
            Case "launcherver":
                If (Len(sTemp) >= 8) Then
                    sValue = Left$(sTemp, 8)
                    lLauncherVersion = CLng(StringFormat("&H{0}", sValue))
                    sTemp = Mid$(sTemp, 9)
                    
                    CommandLine = StringFormat("{0}-launcherver {1} ", CommandLine, sValue)
                End If
                
            Case "launchererror":
                If (Left$(sTemp, 1) = Chr$(34)) Then
                    If (InStr(2, sTemp, Chr$(34), vbTextCompare) > 0) Then
                        sValue = Mid$(sTemp, 2, InStr(2, sTemp, Chr$(34), vbTextCompare) - 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 4)
                    Else
                        sValue = Mid$(Split(sTemp & " -", " -")(0), 2)
                        sTemp = Mid$(sTemp, Len(sValue) + 3)
                    End If
                Else
                    sValue = Split(sTemp & " -", " -")(0)
                    sTemp = Mid$(sTemp, Len(sValue) + 2)
                End If
                sRet = StringFormat("{0}{1}YThe StealthBot Profile Launcher had an error!|", Chr(193), sRet)
                If (LenB(sValue) > 0) Then
                    sRet = StringFormat("{0}{1}YOpen {2} for more information|", sRet, Chr(195), sValue)
                End If
                
                CommandLine = StringFormat("{0}-launchererror {1}{2}{1} ", CommandLine, Chr$(34), sValue)
                
            Case Else:
                CommandLine = StringFormat("{0}-{1} ", CommandLine, sSetting)
        End Select
    Loop
    
    If MDebug("debug") Then
        sRet = StringFormat("{0} * Program executed in debug mode; unhandled packet information will be displayed.|", sRet)
    End If
    
    SetCommandLine = Split(sRet, "|")
    
    Exit Function
ERROR_HANDLER:
    sRet = StringFormat("Error #{0}: {1} in modOtherCode.SetCommandLine()|CommandLine: {2}", Err.Number, Err.Description, sCommandLine)
    SetCommandLine = Split(sRet, "|")
    Err.Clear
End Function

Private Function AddEnvPath(sPath As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sTemp As String
    Dim lRet  As Long
    AddEnvPath = False
    
    sTemp = String$(1024, Chr$(0))
    lRet = GetEnvironmentVariable("PATH", sTemp, Len(sTemp))
        
    If (Not lRet = 0) Then
        If (InStr(1, sTemp, sPath, vbTextCompare) = 0) Then
            sTemp = Left$(sTemp, lRet)
            lRet = SetEnvironmentVariable("PATH", StringFormat("{0};{1}", sTemp, sPath))
            AddEnvPath = (lRet = 0)
            If (MDebug("debug")) Then
                frmChat.AddChat RTBColors.ConsoleText, "AddEnvPath failed: Set"
                frmChat.AddChat RTBColors.ConsoleText, StringFormat("PATH: {0}", sTemp)
                frmChat.AddChat RTBColors.ConsoleText, StringFormat("ADD:  {0}", sPath)
            End If
        End If
    Else
        If (MDebug("debug")) Then
            frmChat.AddChat RTBColors.ConsoleText, "AddEnvPath failed: Get"
            frmChat.AddChat RTBColors.ConsoleText, StringFormat("Ret: {0}", lRet)
        End If
    End If

    Exit Function
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.{3}()", Err.Number, Err.Description, "modOtherCode", "AddEnvPath")
    Err.Clear
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
    
    If Config.FlashOnCatchPhrases Then
        Call FlashWindow
    End If
    
    Select Case (mType)
        Case CPTALK:    s = "TALK"
        Case CPEMOTE:   s = "EMOTE"
        Case CPWHISPER: s = "WHISPER"
    End Select
    
    If (Dir$(GetFilePath(FILE_CAUGHT_PHRASES)) = vbNullString) Then
        Open GetFilePath(FILE_CAUGHT_PHRASES) For Output As #i
            Print #i, "<html>"
        Close #i
    End If
    
    Open GetFilePath(FILE_CAUGHT_PHRASES) For Append As #i
        If (LOF(i) > 10000000) Then
            Close #i
            
            Call Kill(GetFilePath(FILE_CAUGHT_PHRASES))
            
            Open GetFilePath(FILE_CAUGHT_PHRASES) For Output As #i
        End If
        
        Msg = Replace(Msg, "<", "&lt;", 1)
        Msg = Replace(Msg, ">", "&gt;", 1)
        
        Print #i, "<B>" & Format(Date, "MM-dd-yyyy") & " - " & Time & _
            " - " & s & Space(1) & Username & ": </B>" & _
                Replace(Msg, Phrase, "<i>" & Phrase & "</i>", 1) & _
                    "<br>"
    Close #i
End Sub


Public Function DoReplacements(ByVal s As String, Optional ByVal Username As String, _
    Optional ByVal Ping As Long) As String

    Dim gAcc As udtUserAccess
    Dim Rank As Integer

    s = Replace(s, "%ver", CVERSION)
    s = Replace(s, "%v", CVERSION)
    s = Replace(s, "%botup", ConvertTimeInterval(modBNCS.GetConnectionUptime()))
    s = Replace(s, "%cpuup", ConvertTimeInterval(modDateTime.GetTickCountMS()))
    s = Replace(s, "%chan", g_Channel.Name)
    s = Replace(s, "%c", g_Channel.Name)
    s = Replace(s, "%me", GetCurrentUsername)
    s = Replace(s, "%1", GetCurrentUsername)
    s = Replace(s, "%rnd", GetRandomUserInChannel)
    s = Replace(s, "%bc", BanCount)
    s = Replace(s, "%mp3", Replace(MediaPlayer.TrackName, "&", "+"))
    s = Replace(s, "%quote", g_Quotes.GetRandomQuote)
    s = Replace(s, "%t", FormatDateTime(Now, vbLongTime))
    s = Replace(s, "%d", FormatDateTime(Now, vbShortDate))

    If Not IsMissing(Username) Then
        gAcc = Database.GetUserAccess(Username)
        Rank = 0
        If gAcc.Rank > 0 Then Rank = gAcc.Rank

        s = Replace(s, "%user", Username)
        s = Replace(s, "%u", Username)
        s = Replace(s, "%0", Username)
        If Not IsMissing(Ping) And Ping > -2 Then
            s = Replace(s, "%p", Ping)
        End If
        s = Replace(s, "%a", Rank)
        s = Replace(s, "%rank", Rank)
        s = Replace(s, "%r", Rank)
        s = Replace(s, "%flags", gAcc.Flags)
        s = Replace(s, "%flag", gAcc.Flags)
        s = Replace(s, "%f", gAcc.Flags)
        s = Replace(s, "%mail", GetMailCount(Username))
        s = Replace(s, "%m", GetMailCount(Username))
    End If

    DoReplacements = s
End Function

Public Function GetRandomUserInChannel() As String
    Dim i As Integer
    
    If (g_Channel.Users.Count > 0) Then
        Randomize
        
        i = Int(g_Channel.Users.Count * Rnd + 1)

        GetRandomUserInChannel = g_Channel.Users(i).DisplayName
    End If
End Function

Public Function ListFileLoad(ByVal sPath As String, Optional ByVal MaxItems As Integer = -1) As Collection
    On Error GoTo ERROR_HANDLER
    
    Dim f As Integer
    Dim i As Integer
    Dim s As String
    Dim List As Collection
    
    Set List = New Collection
    
    If (LenB(Dir$(sPath)) > 0) Then
        f = FreeFile
        i = 0
        
        Open sPath For Input As #f
            If (LOF(f) > 0) Then
                Do
                    Line Input #f, s
                    
                    If LenB(s) > 0 Then
                        List.Add s
                        i = i + 1
                    End If
                    
                Loop Until EOF(f) Or (MaxItems >= 0 And i >= MaxItems)
            End If
        Close #f
    End If
    
    Set ListFileLoad = List
    
    Exit Function
    
ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileLoad()", _
        Err.Number, Err.Description, OBJECT_NAME)
End Function

Public Sub ListFileAppendItem(ByVal sPath As String, ByVal Item As String)
    On Error GoTo ERROR_HANDLER

    Dim f As Integer
    
    f = FreeFile
    
    If (LenB(Dir$(sPath)) > 0) Then
        Open sPath For Append As #f
            Print #f, Item
        Close #f
    Else
        Open sPath For Output As #f
            Print #f, Item
        Close #f
    End If
    
    Exit Sub
    
ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileAppendItem()", _
        Err.Number, Err.Description, OBJECT_NAME)
End Sub

Public Sub ListFileSave(ByVal sPath As String, ByVal List As Collection)
    On Error GoTo ERROR_HANDLER

    Dim f As Integer
    Dim i As Long
    
    f = FreeFile
    
    Open sPath For Output As #f
        ' print each quote
        For i = 1 To List.Count
            Print #f, List.Item(i)
        Next i
    Close #f
    
    Exit Sub
    
ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.ListFileSave()", _
        Err.Number, Err.Description, OBJECT_NAME)
End Sub

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

Public Sub LogDbAction(ByVal ActionType As enuDBActions, ByVal Caller As String, ByVal Target As String, _
    ByVal TargetType As String, Optional ByVal Rank As Integer, Optional ByVal Flags As String, _
        Optional ByVal Group As String)
    
    'Dim sPath  As String
    'Dim Action As String
    'Dim f      As Integer
    Dim str As String
    
    If ((LenB(Caller) = 0) Or (StrComp(Caller, "(console)", vbTextCompare) = 0)) Then
        Caller = "console"
    End If
    
    Select Case (ActionType)
        Case AddEntry
            str = Caller & " adds " & Target
        Case ModEntry
            str = Caller & " modifies " & Target
        Case RemEntry
            str = Caller & " removes " & Target
    End Select
    
    If (StrComp(TargetType, "user", vbTextCompare) <> 0) Then
        str = str & " (" & LCase$(TargetType) & ")"
    End If
    
    If (Rank > 0) Then
        str = str & " " & Rank
    End If
    
    If (Flags <> vbNullString) Then
        str = str & " " & Flags
    End If
    
    If (Group <> vbNullString) Then
        str = str & ", groups: " & Group
    End If
    
    g_Logger.WriteDatabase str
    
    'f = FreeFile
    'sPath = GetProfilePath() & "\Logs\database.txt"
    
    'If (LenB(Dir$(sPath)) = 0) Then
    '    Open sPath For Output As #f
    'Else
    '    Open sPath For Append As #f
    '
    '    If ((LOF(f) > BotVars.MaxLogFileSize) And (BotVars.MaxLogFileSize > 0)) Then
    '        Close #f
    '
    '        Call Kill(sPath)
    '
    '        Open sPath For Output As #f
    '            Print #f, "Logfile cleared automatically on " & _
    '                Format(Now, "HH:MM:SS MM/DD/YY") & "."
    '    End If
    'End If
    
    'Select Case (ActionType)
    '    Case AddEntry
    '        Action = Caller & _
    '            " adds " & Target & " " & Instruction
    '
    '    Case ModEntry
    '        Action = Caller & _
    '            " modifies " & Target & " " & Instruction
    '
    '    Case RemEntry
    '        Action = Caller & " removes " & Target
    'End Select
    
    'Action = _
    '    Caller & " " & Action & Space(1) & Target & ": " & Instruction
        
    'g_Logger.WriteDatabase Action
    
    'Print #f, Action
    
    'Close #f
End Sub

Public Sub LogCommand(ByVal Caller As String, ByVal CString As String)
    On Error GoTo LogCommand_Error
    
    Dim sPath  As String
    Dim Action As String
    Dim f      As Integer

    If (LenB(CString) > 0) Then
        If (LenB(Caller) = 0) Then
            Caller = "%console%"
        End If
    
        g_Logger.WriteCommand Caller & " -> " & CString
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

Function GetProductKey(Optional ByVal Product As String) As String
    If (LenB(Product) = 0) Then
        Product = StrReverse$(BotVars.Product)
    End If
    
    GetProductKey = GetProductInfo(Product).ShortCode
    
    If (LenB(ReadCfg$("Override", StringFormat("{0}ProdKey", Product))) > 0) Then
        GetProductKey = ReadCfg$("Override", StringFormat("{0}ProdKey", Product))
    End If
End Function

Public Function InsertDummyQueueEntry()
    ' %%%%%blankqueuemessage%%%%%
    frmChat.AddQ "%%%%%blankqueuemessage%%%%%"
End Function

' This procedure splits a message by a specified Length, with optional line LinePostfixes
' and split delimiters.
' No longer puts a delim before [more] if it wasn't split by the delim -Ribose/2009-08-30
Public Function SplitByLen(ByVal StringSplit As String, ByVal SplitLength As Long, ByRef StringRet() As String, Optional ByVal LinePrefix As String = _
    vbNullString, Optional ByVal LinePostfix As String, Optional ByVal OversizeDelimiter As String = " ") As Long
    
    On Error GoTo ERROR_HANDLER

    ' maximum size of battle.net messages
    Const BNET_MSG_LENGTH = 223
    
    Dim lineCount As Long    ' stores line number
    Dim pos       As Long    ' stores position of delimiter
    Dim strTmp    As String  ' stores working copy of StringSplit
    Dim Length    As Long    ' stores Length after LinePostfix
    Dim bln       As Boolean ' stores result of delimiter split
    Dim s         As String  ' stores temp string for settings
    
    ' check for custom line postfix
    s = Config.MultiLinePostfix
    If LenB(s) > 0 Then
        If Left$(s, 1) = "{" And Right$(s, 1) = "}" Then
            LinePostfix = Mid$(s, 2, Len(s) - 2)
        Else
            LinePostfix = s
        End If
    Else
        LinePostfix = "[more]"
    End If
    
    ' initialize our array
    ReDim StringRet(0)
    
    ' default our first index
    StringRet(0) = vbNullString
    
    If (SplitLength = 0) Then
        SplitLength = BNET_MSG_LENGTH
    End If
    
    If (Len(LinePrefix) >= SplitLength) Then
        Exit Function
    End If
    
    If (Len(LinePostfix) >= SplitLength) Then
        Exit Function
    End If
    
    ' do loop until our string is empty
    Do While (StringSplit <> vbNullString)
        ' resize array so that it can store
        ' the next line
        ReDim Preserve StringRet(lineCount)
        
        ' store working copy of string
        strTmp = LinePrefix & StringSplit
        
        ' does our string already equal to or fall
        ' below the specified Length?
        If (Len(strTmp) <= SplitLength) Then
            ' assign our string to the current line
            StringRet(lineCount) = strTmp
        Else
            ' Our string is over the size limit, so we're
            ' going to postfix it.  Because of this, we're
            ' going to have to calculate the Length after
            ' the postfix has been accounted for.
            Length = (SplitLength - Len(LinePostfix))
        
            ' if we're going to be splitting the oversized
            ' message at a specified character, we need to
            ' determine the position of the character in the
            ' string
            If (OversizeDelimiter <> vbNullString) Then
                ' grab position of delimiter character that is the closest to our
                ' specified Length
                pos = InStrRev(strTmp, OversizeDelimiter, Length, vbTextCompare)
            End If
            
            ' if the delimiter we were looking for was found,
            ' and the position was greater than or equal to
            ' half of the message (this check prevents breaks
            ' in unecessary locations), split the message
            ' accordingly.
            If ((pos) And (pos >= Round(Length / 2))) Then
                ' truncate message
                strTmp = Mid$(strTmp, 1, pos)
                
                ' indicate that an additional
                ' character will require removal
                ' from official copy
                'bln = (Not KeepDelim)
            Else
                ' truncate message
                strTmp = Mid$(strTmp, 1, Length)
            End If
            
            ' store truncated message in line
            StringRet(lineCount) = strTmp & LinePostfix
        End If
        
        ' remove line from official string
        StringSplit = Mid$(StringSplit, (Len(strTmp) - Len(LinePrefix)) + 1)
        
        ' if we need to remove an additional
        ' character, lets do so now.
        'If (bln) Then
        '    StringSplit = Mid$(StringSplit, Len(OversizeDelimiter) + 1)
        'End If
            
        ' increment line counter
        lineCount = (lineCount + 1)
    Loop
    
    SplitByLen = lineCount
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: " & Err.Description & " in SplitByLen()."
    
    Exit Function
End Function

Public Function UsernameRegex(ByVal Username As String, ByVal sPattern As String) As Boolean
    Dim prepName As String
    Dim prepPatt As String

    prepName = Replace(Username, "[", "{")
    prepName = Replace(prepName, "]", "}")
    prepName = LCase$(prepName)
    
    prepPatt = Replace(sPattern, "\[", "{")
    prepPatt = Replace(prepPatt, "\]", "}")
    prepPatt = LCase$(prepPatt)

    UsernameRegex = (prepName Like prepPatt)
End Function


' Fixed font issue when an element was only 1 character long -Pyro (9/28/08)
' Fixed issue with displaying null text.

' I changed the location of where fontStr was being declared. I can't figure out why it makes any difference, but _
    when I have it declared above I get memory errors, my IDE crashes, I runtime erro 380 and _
  - Error (#-2147417848): Method 'SelFontName' of object 'IRichText' failed in DisplayRichText().
'I believe it's something to do with the subclassing overwriting the memory, but it only occurs when run from the IDE. - FrOzeN

Public Sub DisplayRichText(ByRef rtb As RichTextBox, ByRef saElements() As Variant)
    On Error GoTo ERROR_HANDLER
   
    Dim arr()          As Variant
    Dim s              As String
    Dim lngVerticalPos As Long
    Dim Diff           As Long
    Dim i              As Long
    Dim intRange       As Long
    Dim blUnlock       As Boolean
    Dim LogThis        As Boolean
    Dim Length         As Long
    Dim NewLength      As Long
    Dim RemoveLength   As Long
    Dim EscapePos      As Long
    Dim Index          As Long
    Dim str            As String
    Dim arrCount       As Long
    Dim SelStart       As Long
    Dim SelLength      As Long
    Dim blnHasFocus    As Boolean
    Dim blnAtEnd       As Boolean
    
    Static RichTextErrorCounter As Integer

    ' *****************************************
    '              SANITY CHECKS
    ' *****************************************

    ' empty array
    If LBound(saElements) > UBound(saElements) Then
        Exit Sub
    End If

    ' if first element is numeric
    If (IsNumeric(saElements(0))) Then
        Index = 2

        ' convert AddChat Color, Text, Color, Text, ...
        '      to AddChat DefaultFont, Color, Text, DefaultFont, Color, Text, ...
        For i = LBound(saElements) To UBound(saElements) Step 2
            ReDim Preserve arr(0 To Index) As Variant

            arr(Index) = saElements(i + 1)
            arr(Index - 1) = saElements(i)
            arr(Index - 2) = rtb.Font.Name

            Index = Index + 3
        Next i

        saElements() = arr()
    End If

    ' verify arguments
    For i = LBound(saElements) To UBound(saElements) Step 3
        ' element count not a multiple of 3
        If ((i + 2) > UBound(saElements)) Then
            Exit Sub
        End If

        ' color is not positive Integer or Long value
        If (IsNumeric(saElements(i + 1)) = False) Then
            Exit Sub
        End If

        ' convert negative Integer values to Long (for example &H99CC is negative, convert to &H99CC&)
        If (saElements(i + 1) < 0) Then
            saElements(i + 1) = CLng(saElements(i + 1) + &H10000)
        Else
            saElements(i + 1) = CLng(saElements(i + 1))
        End If

        ' out of color range
        If (saElements(i + 1) < 0 Or saElements(i + 1) > &HFFFFFF) Then
            Exit Sub
        End If

        ' store combined length of input
        NewLength = NewLength + Len(KillNull(saElements(i + 2)))
    Next i

    ' input must have non-zero length
    If (NewLength = 0) Then
        Exit Sub
    End If

    If ((BotVars.LockChat = False) Or (rtb <> frmChat.rtbChat)) Then
        ' store rtb carat and whether rtb has focus
        GetTextSelection rtb, SelStart, SelLength
        blnHasFocus = (rtb.Parent.ActiveControl Is rtb And rtb.Parent.WindowState <> vbMinimized)
        ' whether it's at the end or within one vbCrLf of the end
        blnAtEnd = (SelStart >= GetRTBLength(rtb) - 2)

        ' is the RTB at the bottom?
        lngVerticalPos = IsScrolling(rtb)
    
        If (lngVerticalPos) Then
            rtb.Visible = False

            ' below causes smooth scrolling, but also screen flickers :(
            'LockWindowUpdate rtb.hWnd

            blUnlock = True
        End If

        ' how to log this event
        If (rtb = frmChat.rtbChat) Or (rtb = frmChat.rtbWhispers) Then
            LogThis = (BotVars.Logging > 0)
        End If

        ' remove from backlog if overflow
        Length = GetRTBLength(rtb)
        If ((BotVars.MaxBacklogSize) And (Length >= BotVars.MaxBacklogSize)) Then
            If (blUnlock = False) Then
                rtb.Visible = False

                ' below causes smooth scrolling, but also screen flickers :(
                'LockWindowUpdate rtb.hWnd
            End If

            With rtb
                If Length > BotVars.MaxBacklogSize Then
                    RemoveLength = InStr(Length - BotVars.MaxBacklogSize, GetRTBText(rtb), vbLf, vbBinaryCompare)
                Else
                    RemoveLength = InStr(1, Length, vbLf, vbBinaryCompare)
                End If
                SetTextSelection rtb, 0, RemoveLength
                ' remove line from stored selection
                SelStart = SelStart - RemoveLength
                SelLength = SelLength - RemoveLength
                ' if selection included part of what was removed, add negative start point
                ' to length to get difference in length, and set start selection at 0
                If SelStart < 0 Then SelStart = 0
                If SelLength < 0 Then SelLength = 0
                RTBSetSelectedText rtb, vbNullString
            End With

            If (blUnlock = False) Then
                rtb.Visible = True

                ' below causes smooth scrolling, but also screen flickers :(
                'LockWindowUpdate &H0
            End If
        End If

        ' place timestamp
        s = GetTimeStamp()
        If LenB(s) > 0 Then
            With rtb
                SetTextSelection rtb, -1, -1
                .SelFontName = rtb.Font.Name
                .SelFontSize = rtb.Font.Size
                .SelBold = False
                .SelItalic = False
                .SelUnderline = False
                .SelStrikeThru = False
                .SelColor = RTBColors.TimeStamps
                RTBSetSelectedText rtb, s
            End With
        End If

        ' place each element
        For i = LBound(saElements) To UBound(saElements) Step 3
            If ((StrictIsNumeric(saElements(i + 1))) And (Len(saElements(i + 2)) > 0)) Then
                s = KillNull(saElements(i + 2))

                With rtb
                    SetTextSelection rtb, -1, -1
                    .SelFontName = saElements(i)
                    .SelColor = saElements(i + 1)
                    RTBSetSelectedText rtb, s
                    str = str & s
                End With
            End If
        Next i

        With rtb
            SetTextSelection rtb, -1, -1
            RTBSetSelectedText rtb, vbCrLf
        End With

        If (LogThis) Then
            If (rtb = frmChat.rtbChat) Then
                g_Logger.WriteChat str
            ElseIf (rtb = frmChat.rtbWhispers) Then
                g_Logger.WriteWhisper str
            End If
        End If

        ColorModify rtb, GetRTBLength(rtb) - NewLength

        If (blUnlock) Then
            SendMessage rtb.hWnd, WM_VSCROLL, _
                SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&
                
            rtb.Visible = True
                
            ' below causes smooth scrolling, but also screen flickers :(
            'LockWindowUpdate &H0
        End If
        
        With rtb
            ' if has focus
            If blnHasFocus Then
                ' restore carat location and selection if not previously at end
                If Not blnAtEnd Then
                    SetTextSelection rtb, SelStart, SelLength
                End If
            End If
        End With
    End If

    RichTextErrorCounter = 0

    Exit Sub
    
ERROR_HANDLER:

    RichTextErrorCounter = RichTextErrorCounter + 1
    If RichTextErrorCounter > 2 Then
        RichTextErrorCounter = 0
        Exit Sub
    End If
    
    If (Err.Number = 13 Or Err.Number = 91) Then
        Exit Sub
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DisplayRichText()."
    
    Exit Sub
    
End Sub

Public Function IsScrolling(ByRef rtb As RichTextBox) As Long

    Dim lngVerticalPos As Long
    Dim difference     As Long
    Dim range          As Integer

    If (g_OSVersion.IsWin2000Plus()) Then

        GetScrollRange rtb.hWnd, SB_VERT, 0, range
        
        lngVerticalPos = SendMessage(rtb.hWnd, EM_GETTHUMB, 0&, 0&)
        
        If ((lngVerticalPos = 0) And (range > 0)) Then
            lngVerticalPos = 1
        End If

        difference = ((lngVerticalPos + (rtb.Height / Screen.TwipsPerPixelY)) - _
            range)

        ' In testing it appears that if the value I calcuate as Diff is negative,
        ' the scrollbar is not at the bottom.
        If (difference < 0) Then
            IsScrolling = lngVerticalPos
        End If
        
    End If

End Function

Public Sub CloseAllConnections(Optional ShowMessage As Boolean = True)
    If (frmChat.sckBNLS.State <> sckClosed) Then: frmChat.sckBNLS.Close
    If (frmChat.sckBNet.State <> sckClosed) Then: frmChat.sckBNet.Close
    If (frmChat.sckMCP.State <> sckClosed) Then: frmChat.sckMCP.Close
    
    If (ShowMessage) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "All connections closed."
    End If
    
    BNLSAuthorized = False
    
    SetTitle "Disconnected"
    
    frmChat.UpdateTrayTooltip
    
    g_Online = False
    
    RunInAll "Event_ServerError", "All connections closed."
End Sub

Public Sub BuildProductInfo()
    ' 4-digit code, short code, full name, home channel, number of keys, BNLS ID, logon system, version byte
    ProductList(0) = CreateProductInfo("UNKW", vbNullString, "Unknown Product", "Unknown", 0, &H0, &H0, &H0)
    ProductList(1) = CreateProductInfo(PRODUCT_STAR, "SC", "StarCraft", "StarCraft", 0, &H1, BNCS_NLS, &HD6)
    ProductList(2) = CreateProductInfo(PRODUCT_SEXP, "SC", "StarCraft Broodwar", "Brood War", 0, &H2, BNCS_NLS, &HD6)
    ProductList(3) = CreateProductInfo(PRODUCT_W2BN, "W2", "WarCraft II: Battle.net Edition", "WarCraft II", 1, &H3, BNCS_OLS, &H4F)
    ProductList(4) = CreateProductInfo(PRODUCT_D2DV, "D2", "Diablo II", "Diablo II", 1, &H4, BNCS_NLS, &HE)
    ProductList(5) = CreateProductInfo(PRODUCT_D2XP, "D2X", "Diablo II: Lord of Destruction", "Diablo II", 2, &H5, BNCS_NLS, &HE)
    ProductList(6) = CreateProductInfo(PRODUCT_WAR3, "W3", "WarCraft III: Reign of Chaos", "W3", 1, &H7, BNCS_NLS, &H1C)
    ProductList(7) = CreateProductInfo(PRODUCT_W3XP, "W3", "WarCraft III: The Frozen Throne", "W3", 2, &H8, BNCS_NLS, &H1C)
    ProductList(8) = CreateProductInfo(PRODUCT_DSHR, "DS", "Diablo Shareware", "Diablo Shareware", 0, &HA, BNCS_OLS, &H2A)
    ProductList(9) = CreateProductInfo(PRODUCT_DRTL, "D1", "Diablo", "Diablo Retail", 0, &H9, BNCS_OLS, &H2A)
    ProductList(10) = CreateProductInfo(PRODUCT_SSHR, "SS", "StarCraft Shareware", "StarCraft", 0, &HB, BNCS_LLS, &HA9)
    ProductList(11) = CreateProductInfo(PRODUCT_JSTR, "JS", "Japanese StarCraft", "StarCraft", 1, &H6, BNCS_LLS, &HA9)
    ProductList(12) = CreateProductInfo(PRODUCT_CHAT, "CHAT", "Telnet Chat", "Telnet", 0, &H0, &H0, &H0)
End Sub

Private Function CreateProductInfo(ByVal sCode As String, ByVal sShort As String, ByVal sLongName As String, ByVal sChannelName As String, ByVal iKeys As Integer, ByVal iBnlsId As Long, ByVal iLogonSystem As Long, ByVal iVerByte As Long) As udtProductInfo
    Dim pi As udtProductInfo
    pi.Code = UCase$(sCode)
    pi.ShortCode = UCase$(sShort)
    pi.FullName = sLongName
    pi.ChannelName = sChannelName
    pi.KeyCount = iKeys
    pi.BNLS_ID = iBnlsId
    pi.LogonSystem = iLogonSystem
    pi.VersionByte = iVerByte
    
    CreateProductInfo = pi
End Function

Public Function GetProductInfo(ByVal sProductCode As String) As udtProductInfo
    Dim pi As udtProductInfo
    Dim Index As Integer
    sProductCode = UCase$(sProductCode)

    For Index = 0 To UBound(ProductList)
        pi = ProductList(Index)
        
        If StrComp(pi.Code, sProductCode, vbBinaryCompare) = 0 Or _
            StrComp(pi.Code, StrReverse(sProductCode), vbBinaryCompare) = 0 Or _
            StrComp(pi.ShortCode, sProductCode, vbBinaryCompare) = 0 Then
            
            GetProductInfo = pi
            Exit Function
        End If
    Next
    GetProductInfo = ProductList(0)
End Function

'Returns the number of monitors active on the computer.
Public Function GetMonitorCount() As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, GetMonitorCount
End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hDCMonitor As Long, ByVal lprcMonitor As Long, dwData As Long) As Long
    dwData = dwData + 1
    MonitorEnumProc = 1
End Function
