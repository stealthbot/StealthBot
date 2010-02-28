Attribute VB_Name = "modParsing"
Option Explicit

Public Const COLOR_BLUE2 = 12092001

Public Sub SendHeader()
    frmChat.sckBNet.SendData ChrW(1)
End Sub

Public Sub BNCSParsePacket(ByVal PacketData As String)
    On Error GoTo ERROR_HANDLER

    Dim pD          As clsDataBuffer ' Packet debuffer object
    Dim PacketLen   As Long              ' Length of the packet minus the header
    Dim PacketID    As Byte              ' Battle.net packet ID
    Dim s           As String            ' Temporary string
    Dim L           As Long              ' Temporary long
    Dim s2          As String            ' Temporary string
    Dim s3          As String            ' Temporary string
    Dim ClanTag     As String            ' User clan tag
    Dim Product     As String            ' User product
    Dim w3icon      As String            ' Warcraft III icon code
    Dim B           As Boolean           ' Temporary bool
    Dim sArr()      As String            ' Temp String array
    Dim veto        As Boolean
    
    
    '--------------
    '| Initialize |
    '--------------
    Set pD = New clsDataBuffer
    PacketLen = Len(PacketData) - 4
    
    '###########################################################################
    
    If PacketLen >= 0 Then
        ' Start packet debuffer
        pD.Data = Mid$(PacketData, 5)
        ' Get packet ID
        PacketID = Asc(Mid$(PacketData, 2, 1))
        
        If MDebug("all") Then
            frmChat.AddChat COLOR_BLUE, "BNET RECV 0x" & ZeroOffset(PacketID, 2)
        End If
        
        ' ...
        CachePacket StoC, stBNCS, PacketID, Len(PacketData), PacketData
        
        ' Added 2007-06-08 for a packet logging menu feature to aid tech support
        WritePacketData stBNCS, StoC, PacketID, PacketLen, PacketData
                
        ' ...-
        If (RunInAll("Event_PacketReceived", "BNCS", PacketID, Len(PacketData), PacketData)) Then
            Exit Sub
        End If
        
        'This will be taken out when Warden is moved to a script like I want.
        If (modWarden.WardenData(WardenInstance, PacketData, False)) Then
          Exit Sub
        End If
        
        '--------------
        '| Parse      |
        '--------------
        
        Select Case PacketID
                
            '###########################################################################
            Case &H26 'SID_READUSERDATA
                ProfileParse PacketData
            
            '###########################################################################
            Case Is >= &H65 'Friends List or Clan-related packet
                ' Hand the packet off to the appropriate handler
                If PacketID >= &H70 Then
                    ' added in response to the clan channel takeover exploit
                    ' discovered 11/7/05
                    If IsW3 Then
                        frmChat.ParseClanPacket PacketID, IIf(Len(PacketData) > 4, Mid$(PacketData, 5), vbNullString)
                    End If
                Else
                    If (g_request_receipt) Then
                        g_request_receipt = False
                        
                        If (Caching) Then
                            frmChat.cacheTimer_Timer
                        End If
                        
                        Exit Sub
                    End If
                
                    frmChat.ParseFriendsPacket PacketID, Mid$(PacketData, 5)
                End If
            
            '###########################################################################
            Case Else
                Call modBNCS.BNCSRecvPacket(PacketData)
            
        End Select
    End If
    
    Set pD = Nothing
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in BNCSParsePacket()."
    
    Exit Sub
End Sub

Public Function StrToHex(ByVal String1 As String, Optional ByVal NoSpaces As Boolean = False) As String
    Dim strTemp As String, strReturn As String, i As Long
    
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        
        strReturn = strReturn & IIf(NoSpaces, "", Space(1)) & strTemp
    Next i
        
    StrToHex = strReturn
End Function

Public Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    'on error resume next
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function

Public Function KillNull(ByVal Text As String) As String
    Dim i As Integer
    i = InStr(1, Text, Chr(0))
    If (i = 0) Then
        KillNull = Text
        Exit Function
    End If
    KillNull = Left$(Text, i - 1)
End Function

Public Function GetHexValue(ByVal v As Long) As String

    v = v And &HF
    
    If v < 10 Then
    
        GetHexValue = Chr$(v + &H30)
        
    Else
    
        GetHexValue = Chr$(v + &H37)
        
    End If
    
End Function

Public Function GetNumValue(ByVal c As String) As Long
'on error resume next
    c = UCase(c)
    
    If StrictIsNumeric(c) Then
    
        GetNumValue = Asc(c) - &H30
        
    Else
    
        GetNumValue = Asc(c) - &H37
        
    End If
    
End Function

Public Sub NullTruncString(ByRef Text As String)
'on error resume next
    Dim i As Integer
    
    i = InStr(Text, Chr(0))
    If i = 0 Then Exit Sub
    
    Text = Left$(Text, i - 1)
End Sub

Public Sub FullJoin(Channel As String, Optional ByVal i As Long = -1)
    If i >= 0 Then
        PBuffer.InsertDWord CLng(i)
    Else
        PBuffer.InsertDWord &H2
    End If
    PBuffer.InsertNTString Channel
    PBuffer.SendPacket &HC
End Sub

Public Function HexToStr(ByVal Hex1 As String) As String
'on error resume next
    Dim strReturn As String, i As Long
    If Len(Hex1) Mod 2 <> 0 Then Exit Function
    For i = 1 To Len(Hex1) Step 2
    strReturn = strReturn & Chr(Val("&H" & Mid(Hex1, i, 2)))
    Next i
    HexToStr = strReturn
End Function

Public Sub RejoinChannel(Channel As String)
    'on error resume next
    PBuffer.SendPacket &H10
    PBuffer.InsertDWord &H2
    PBuffer.InsertNTString Channel
    PBuffer.SendPacket &HC
End Sub

Public Sub RequestProfile(strUser As String)
    'on error resume next
    With PBuffer
        .InsertDWord 1
        .InsertDWord 4
        .InsertDWord GetTickCount()
        .InsertNTString CleanUsername(ReverseConvertUsernameGateway(strUser))
        .InsertNTString "Profile\Age"
        .InsertNTString "Profile\Sex"
        .InsertNTString "Profile\Location"
        .InsertNTString "Profile\Description"
        .SendPacket &H26
    End With
End Sub

Public Sub RequestSpecificKey(ByVal sUsername As String, ByVal sKey As String)
    With PBuffer
        .InsertDWord 1
        .InsertDWord 1
        .InsertDWord GetTickCount()
        .InsertNTString ReverseConvertUsernameGateway(sUsername)
        .InsertNTString sKey
        .SendPacket &H26
    End With
End Sub

Public Sub SetProfile(ByVal Location As String, ByVal description As String, Optional ByVal Sex As String = vbNullString)
    'Dim i As Byte
    Const MAX_DESCR As Long = 510
    Const MAX_SEX As Long = 200
    Const MAX_LOC As Long = 200
    
    '// Sanity checks
    If LenB(description) = 0 Then
        description = Space(1)
    ElseIf Len(description) > MAX_DESCR Then
        description = Left$(description, MAX_DESCR)
    End If
    
    If LenB(Sex) = 0 Then
        Sex = Space(1)
    ElseIf Len(Sex) > MAX_SEX Then
        Sex = Left$(Sex, MAX_SEX)
    End If
    
    If LenB(Location) = 0 Then
        Location = Space(1)
    ElseIf Len(Location) > MAX_LOC Then
        Location = Left$(Location, MAX_LOC)
    End If
    
    
    With PBuffer
        .InsertDWord &H1                    '// #accounts
        .InsertDWord 3                      '// #keys
        
        .InsertNTString CurrentUsername     '// account to update
                                            '// keys
        .InsertNTString "Profile\Location"
        .InsertNTString "Profile\Description"
        .InsertNTString "Profile\Sex"
                                            '// Values()
        .InsertNTString Location
        .InsertNTString description
        .InsertNTString Sex
        
        .SendPacket &H27
    End With
End Sub

'// Extended version of this function for scripting use
'//  Will not ERASE if a field is left blank
'// 2007-06-07: SEX value is ignored because Blizzard removed that
'//     field from profiles
'// 2009-07-14: corrected a problem in this method, thanks Jack (t=42494) -andy
'//     method was erasing profile data
Public Sub SetProfileEx(ByVal Location As String, ByVal description As String)
    'Dim i As Byte
    Const MAX_DESCR As Long = 510
    Const MAX_SEX As Long = 200
    Const MAX_LOC As Long = 200
    
    Dim nKeys As Integer, i As Integer
    Dim pKeys(1 To 3) As String
    Dim pData(1 To 3) As String
    
    If (LenB(Location) > 0) Then
        If (Len(Location) > MAX_LOC) Then
            Location = Left$(Location, MAX_LOC)
        End If
        
        nKeys = nKeys + 1
        pKeys(nKeys) = "Profile\Location"
        pData(nKeys) = Location
    End If
    
    '// Sanity checks
    If (LenB(description) > 0) Then
        If (Len(description) > MAX_DESCR) Then
            description = Left$(description, MAX_DESCR)
        End If
        
        nKeys = nKeys + 1
        pKeys(nKeys) = "Profile\Description"
        pData(nKeys) = description
    End If
        
    If nKeys > 0 Then
        With PBuffer
            .InsertDWord &H1                    '// #accounts
            .InsertDWord nKeys                  '// #keys
            .InsertNTString CurrentUsername     '// account to update
                                                '// keys
            For i = 1 To nKeys
                .InsertNTString pKeys(i)
            Next i
           
            '// Values()
            For i = 1 To nKeys
                .InsertNTString pData(i)
            Next i
            
            .SendPacket &H27
        End With
    End If
End Sub

Public Sub sPrintF(ByRef source As String, ByVal nText As String, _
    Optional ByVal a As Variant, _
    Optional ByVal B As Variant, _
    Optional ByVal c As Variant, _
    Optional ByVal d As Variant, _
    Optional ByVal e As Variant, _
    Optional ByVal f As Variant, _
    Optional ByVal g As Variant, _
    Optional ByVal H As Variant)
    
    nText = Replace(nText, "%S", "%s")
    
    Dim i As Byte
    i = 0
    
    Do While (InStr(1, nText, "%s") <> 0)
        Select Case i
            Case 0
                If IsEmpty(a) Then GoTo theEnd
                nText = Replace(nText, "%s", a, 1, 1)
            Case 1
                If IsEmpty(B) Then GoTo theEnd
                nText = Replace(nText, "%s", B, 1, 1)
            Case 2
                If IsEmpty(c) Then GoTo theEnd
                nText = Replace(nText, "%s", c, 1, 1)
            Case 3
                If IsEmpty(d) Then GoTo theEnd
                nText = Replace(nText, "%s", d, 1, 1)
            Case 4
                If IsEmpty(e) Then GoTo theEnd
                nText = Replace(nText, "%s", e, 1, 1)
            Case 5
                If IsEmpty(f) Then GoTo theEnd
                nText = Replace(nText, "%s", f, 1, 1)
            Case 6
                If IsEmpty(g) Then GoTo theEnd
                nText = Replace(nText, "%s", g, 1, 1)
            Case 7
                If IsEmpty(H) Then GoTo theEnd
                nText = Replace(nText, "%s", H, 1, 1)
        End Select
        i = i + 1
    Loop
theEnd:
    source = source & nText
End Sub

Public Function ParseStatstring(ByVal Statstring As String, ByRef outbuf As String, ByRef sClan As String) As String
    Dim Values() As String
    Dim temp() As String
    Dim cType As String
    Dim WCG As Boolean
   On Error GoTo ParseStatString_Error

    'Dim Icon As String
    
    ' FRCW = Ref
    ' LPCW = Player
    
    'Debug.Print "Received statstring: " & Statstring
    If LenB(Statstring) > 0 Then
        
        g_ThisIconCode = -1
    
        Select Case Left$(Statstring, 4)
            Case "3RAW", "PX3W"
                If Len(Statstring) > 4 Then
                    temp() = Split(Statstring, " ")
                    
                    ReDim Values(3)
                    
                    If StrComp(Right$(temp(1), 2), "CW") = 0 Then
                        WCG = True
                    Else
                        Values(1) = Mid$(Statstring, 6, 1)
                        Values(2) = Mid$(Statstring, 7, 1)
                    End If
                    
                    Values(0) = temp(2)
                    
                    If UBound(temp) > 2 Then
                        Values(3) = StrReverse(temp(3))
                    End If
                    
                    g_ThisIconCode = GetRaceAndIcon(Values(1), Values(2), Left$(Statstring, 4), IIf(WCG, temp(1), ""))
                    
                    sClan = IIf(UBound(Values) > 2, Values(3), "")
                    
                    If Left$(Statstring, 4) = "3RAW" Then
                        Call sPrintF(outbuf, "Warcraft III: Reign of Chaos (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                    Else
                        Call sPrintF(outbuf, "Warcraft III: The Frozen Throne (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                    End If
                Else
                    If Left$(Statstring, 4) = "3RAW" Then
                        Call StrCpy(outbuf, "Warcraft III: Reign of Chaos.")
                        g_ThisIconCode = -56
                    Else
                        Call StrCpy(outbuf, "Warcraft III: The Frozen Throne.")
                        g_ThisIconCode = -10
                    End If
                End If
                
            Case "RHSS"
                Call StrCpy(outbuf, "Starcraft Shareware.")
                
            Case "RATS"
                Values() = Split(Mid$(Statstring, 6), " ")
                If UBound(Values) <> 8 Then
                    Call sPrintF(outbuf, "a Starcraft %sbot", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Else
                    If Values(0) > 0 Then
                        Call sPrintF(outbuf, "Starcraft%s (%s wins, with a rating of %s on the ladder)", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                    Else
                        Call sPrintF(outbuf, "Starcraft%s (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
                    End If
                End If
                
            Case "PXES"
                Values() = Split(Mid(Statstring, 6), " ")
                If UBound(Values) <> 8 Then
                    Call sPrintF(outbuf, "a Starcraft Brood War bot.", vbNullString)
                    
                    If UBound(Values) > 2 Then
                        outbuf = outbuf & "(spawn) "
                    End If
                Else
                    If Values(0) > 0 Then
                        Call sPrintF(outbuf, "Starcraft Brood War%s (%s wins, with a rating of %s on the ladder)", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                    Else
                        Call sPrintF(outbuf, "Starcraft Brood War%s (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
                    End If
                End If
                
            Case "RTSJ"
                Values() = Split(Mid(Statstring, 6), " ")
                If UBound(Values) <> 8 Then
                    Call sPrintF(outbuf, "a Starcraft Japanese %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Else
                    If Values(0) > 0 Then
                        Call sPrintF(outbuf, "Starcraft Japanese%s (%s wins, with a rating of %s on the ladder)", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                    Else
                        Call sPrintF(outbuf, "Starcraft Japanese%s (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
                    End If
                End If
                
            Case "NB2W"
                Values() = Split(Mid$(Statstring, 6), " ")
                
                If UBound(Values) <> 8 Then
                    Call sPrintF(outbuf, "a Warcraft II %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Else
                    If Values(0) > 0 Then
                        Call sPrintF(outbuf, "Warcraft II%s (%s wins, with a rating of %s on the ladder)", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                    Else
                        Call sPrintF(outbuf, "Warcraft II%s (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
                    End If
                End If
                
            Case "RHSD"
                Values() = Split(Mid$(Statstring, 6), " ")
                If UBound(Values) <> 8 Then
                    Call StrCpy(outbuf, "A Diablo shareware bot.")
                Else
                    Select Case Values(1)
                        Case 0: cType = "warrior"
                        Case 1: cType = "rogue"
                        Case 2: cType = "sorceror"
                    End Select
                    Call sPrintF(outbuf, "Diablo shareware (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold)", Values(0), cType, Values(2), Values(3), Values(4), Values(5), Values(6), Values(7))
                End If
                
            Case "LTRD"
                Values() = Split(Mid$(Statstring, 6), " ")
                
                If UBound(Values) <> 8 Then
                    Call StrCpy(outbuf, "A Diablo bot.")
                Else
                    Select Case Values(1)
                        Case 0: cType = "warrior"
                        Case 1: cType = "rogue"
                        Case 2: cType = "sorceror"
                    End Select
                    Call sPrintF(outbuf, "Diablo (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold)", Values(0), cType, Values(2), Values(3), Values(4), Values(5), Values(6), Values(7))
                End If
                
            Case "PX2D"
                Call StrCpy(outbuf, ParseD2Stats(Statstring))
                
            Case "VD2D"
                Call StrCpy(outbuf, ParseD2Stats(Statstring))
                
            Case "TAHC"
                Call StrCpy(outbuf, "a Chat bot.")
                
            Case Else
                Call StrCpy(outbuf, "an unknown client.")
                
        End Select
        
        ParseStatstring = StrReverse(Left$(Statstring, 4))
        
    End If

ParseStatString_Exit:
    Exit Function

ParseStatString_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure ParseStatString of Module modParsing"
    outbuf = "- Error parsing statstring. [" & Replace(Statstring, Chr(0), "") & "]"
    
    Resume ParseStatString_Exit
End Function

' This code cleaned 3/4/2005
Public Function ParseD2Stats(ByVal Stats As String)
    Dim Female As Boolean, Expansion As Boolean
    Dim sLen As Byte, Version As Byte, CharClass As Byte, Hardcore As Byte, CharLevel As Byte
    Dim StatBuf As String, P() As String, Server As String, Name As String
    
    Dim D2Classes(0 To 7) As String
        D2Classes(0) = "amazon"
        D2Classes(1) = "sorceress"
        D2Classes(2) = "necromancer"
        D2Classes(3) = "paladin"
        D2Classes(4) = "barbarian"
        D2Classes(5) = "druid"
        D2Classes(6) = "assassin"
        D2Classes(7) = "unknown class"
    
    If Len(Stats) > 4 Then
        sLen = GetServer(Stats, Server)
        sLen = GetCharacterName(Stats, sLen, Name)
        Call MakeArray(Mid$(Stats, sLen), P())
    End If
    
    If Left$(Stats, 4) = "VD2D" Then
        Call StrCpy(StatBuf, "Diablo II (")
    Else
        Call StrCpy(StatBuf, "Diablo II Lord of Destruction (")
    End If
    
    If (Len(Stats) = 4) Then
        Call StrCpy(StatBuf, "Open Character).")
    Else
        Version = Asc(P(0)) - &H80
        
        CharClass = Asc(P(13)) - 1
        If (CharClass < 0) Or (CharClass > 6) Then
            CharClass = 7
        End If
        
        If (CharClass = 0) Or (CharClass = 1) Or (CharClass = 6) Then
            Female = True
        Else
            Female = False
        End If
        
        CharLevel = Asc(P(25))
        Hardcore = Asc(P(26)) And 4
    
        If Left$(Stats, 4) = "PX2D" Then
            If (Asc(P(26)) And &H20) Then
                Select Case RShift((Asc(P(27)) And &H18), 3)
                    Case 1
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Destroyer ")
                        Else
                            Call StrCpy(StatBuf, "Slayer ")
                        End If
                    Case 2, 3
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Conquerer ")
                        Else
                            Call StrCpy(StatBuf, "Champion ")
                        End If
                    Case 4
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Guardian ")
                        Else
                            If Not Female Then
                                Call StrCpy(StatBuf, "Patriarch ")
                            Else
                                Call StrCpy(StatBuf, "Matriarch ")
                            End If
                        End If
                End Select
                
                Expansion = True
            End If
        End If
        
        If Not Expansion Then
            Select Case RShift((Asc(P(27)) And &H18), 3)
                Case 1
                    If Female = False Then
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Count ")
                        Else
                            Call StrCpy(StatBuf, "Sir ")
                        End If
                    Else
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Countess ")
                        Else
                            Call StrCpy(StatBuf, "Dame ")
                        End If
                    End If
                Case 2
                    If Female = False Then
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Duke ")
                        Else
                            Call StrCpy(StatBuf, "Lord ")
                        End If
                    Else
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Duchess ")
                        Else
                            Call StrCpy(StatBuf, "Lady ")
                        End If
                    End If
                Case 3
                    If Female = False Then
                        If Hardcore Then
                            Call StrCpy(StatBuf, "King ")
                        Else
                            Call StrCpy(StatBuf, "Baron ")
                        End If
                    Else
                        If Hardcore Then
                            Call StrCpy(StatBuf, "Queen ")
                        Else
                            Call StrCpy(StatBuf, "Baroness ")
                        End If
                    End If
            End Select
        End If
        
        Call sPrintF(StatBuf, "%s, a ", Name)
        
        If Hardcore Then
            If (Asc(P(26)) And &H8) Then
                Call StrCpy(StatBuf, "dead ")
            End If
            Call StrCpy(StatBuf, "hardcore ")
        End If
        
        If (Asc(P(26)) And &H40 = &H40) Then
            'frmChat.AddChat vbRed, Asc(P(26))
            If ((Asc(P(26)) And &H8) = &H0) Then
                Call StrCpy(StatBuf, "ladder ")
            End If
        End If
        
        Call sPrintF(StatBuf, "level %s ", CharLevel)
        
        Call sPrintF(StatBuf, "%s on realm %s).", D2Classes(CharClass), Server)
    End If
    
    ParseD2Stats = StatBuf
End Function

Public Function GetServer(ByVal Statstring As String, ByRef Server As String) As Byte
    'returns the begining of the character name
    Server = Mid$(Statstring, 5, InStr(5, Statstring, ",") - 5)
    GetServer = InStr(5, Statstring, ",") + 1
End Function

Public Function GetCharacterName(ByVal Statstring As String, ByVal Start As Byte, ByRef cName As String) As Byte
    cName = Mid$(Statstring, Start, InStr(Start, Statstring, ",") - Start)
    GetCharacterName = InStr(Start, Statstring, ",") + 1
End Function

Public Sub StrCpy(ByRef source As String, ByVal nText As String)
    'on error resume next
    source = source & nText
End Sub

Public Sub MakeArray(ByVal Text As String, ByRef nArray() As String)
    Dim i As Long
    ReDim nArray(0)
    For i = 0 To Len(Text)
        nArray(i) = Mid$(Text, i + 1, 1)
        If i <> Len(Text) Then
            ReDim Preserve nArray(0 To UBound(nArray) + 1)
        End If
    Next i
End Sub

Public Function GetRaceAndIcon(ByRef Icon As String, ByRef Race As String, ByVal Product As String, Optional ByRef WCGCode As String) As Integer
    Dim i As Integer, IMLPos As Integer
    Dim PerTier As Integer
        
    If Product = "3RAW" Then
        PerTier = 5
    Else
        PerTier = 6
    End If
    
    'Debug.Print "Icon: " & Icon & "; Race: " & Race
    
    Select Case Race
        Case "H"
            IMLPos = 1
            i = 0
            Race = "Human"
        Case "N"
            IMLPos = 1 + (PerTier * 1)
            i = 10
            Race = "Night Elves"
        Case "U"
            IMLPos = 1 + (PerTier * 2)
            i = 20
            Race = "Undead"
        Case "O"
            IMLPos = 1 + (PerTier * 3)
            i = 30
            Race = "Orcs"
        Case "R"
            IMLPos = 1 + (PerTier * 4)
            i = 40
            Race = "Random"
        Case "T", "D"
            IMLPos = 1 + (PerTier * 5)
            i = 50
            Race = "Tournament"
        
        Case Else
            IMLPos = 1 + (PerTier * 5)
            i = 50
            Race = "unknown"
            
    End Select
    
    If StrictIsNumeric(Icon) Then
        i = i + CInt(Icon)
        IMLPos = IMLPos + (CInt(Icon) - 1)
    End If
    
    If LenB(WCGCode) > 0 Then
        ' This is a WCG PLAYER or OTHER PERSON
        Select Case StrReverse(WCGCode)
        '*** Special icons for WCG added 6/24/07 ***
            Case Is = "WCRF"
                Icon = "WCG Referee"
                IMLPos = IC_WCRF
            Case Is = "WCPL"
                Icon = "WCG Player"
                IMLPos = IC_WCPL
            Case Is = "WCGO"
                Icon = "WCG Gold Medalist"
                IMLPos = IC_WCGO
            Case Is = "WCSI"
                Icon = "WCG Silver Medalist"
                IMLPos = IC_WCSI
            Case Is = "WCBR"
                Icon = "WCG Bronze Medalist"
                IMLPos = IC_WCBR
            Case Is = "WCPG"
                Icon = "WCG Professional Gamer"
                IMLPos = IC_WCPG
            Case Else
                Icon = "unknown"
                IMLPos = ICUNKNOWN ' 37 (38)
        End Select
    Else
        If Product = "3RAW" Then
            Select Case i
                'Peon Icon
                Case 1, 11, 21, 31, 41
                    Icon = "peon"
                'Human Icons
                Case 2: Icon = "footman"
                Case 3: Icon = "knight"
                Case 4: Icon = "Archmage"
                Case 5: Icon = "Medivh"
                'Night Elf Icons
                Case 12: Icon = "archer"
                Case 13: Icon = "druid of the claw"
                Case 14: Icon = "Priestess of the Moon"
                Case 15: Icon = "Furion"
                'Undead Icons
                Case 22: Icon = "ghoul"
                Case 23: Icon = "abomination"
                Case 24: Icon = "Lich"
                Case 25: Icon = "Tichondrius"
                'Orc Icons
                Case 32: Icon = "grunt"
                Case 33: Icon = "tauren"
                Case 34: Icon = "Far Seer"
                Case 35: Icon = "Thrall"
                'Random Icons
                Case 42: Icon = "dragon whelp"
                Case 43: Icon = "blue dragon"
                Case 44: Icon = "red dragon"
                Case 45: Icon = "Deathwing"
                'else
                Case Else
                    Icon = "unknown"
                    IMLPos = ICUNKNOWN ' 26
            End Select
        Else
            Select Case i
                'Peon Icon
                Case 1, 11, 21, 31, 41, 51
                    Icon = "peon"
                'Human Icons
                Case 2: Icon = "rifleman"
                Case 3: Icon = "sorceress"
                Case 4: Icon = "spellbreaker"
                Case 5: Icon = "Blood Mage"
                Case 6: Icon = "Jaina Proudmore"
                'Night Elf Icons
                Case 12: Icon = "huntress"
                Case 13: Icon = "druid of the talon"
                Case 14: Icon = "dryad"
                Case 15: Icon = "Keeper of the Grove"
                Case 16: Icon = "Maiev"
                'Undead Icons
                Case 22: Icon = "crypt fiend"
                Case 23: Icon = "banshee"
                Case 24: Icon = "destroyer"
                Case 25: Icon = "Crypt Lord"
                Case 26: Icon = "Sylvanas"
                'Orc Icons
                Case 32: Icon = "headhunter"
                Case 33: Icon = "shaman"
                Case 34: Icon = "Spirit Walker"
                Case 35: Icon = "Shadow Hunter"
                Case 36: Icon = "Rexxar"
                'Random Icons
                Case 42: Icon = "myrmidon"
                Case 43: Icon = "siren"
                Case 44: Icon = "dragon turtle"
                Case 45: Icon = "sea witch"
                Case 46: Icon = "Illidan"
                'Tournament Icons
                Case 52: Icon = "Felguard"
                Case 53: Icon = "infernal"
                Case 54: Icon = "doomguard"
                Case 55: Icon = "pit lord"
                Case 56: Icon = "Archimonde"
                'Everything else
                Case Else
                    Icon = "unknown"
                    IMLPos = ICUNKNOWN ' 37 (38)
                    
            End Select
        End If
    End If
    
    If IMLPos > frmChat.imlIcons.ListImages.Count Then
        IMLPos = ICUNKNOWN
    End If
    'Debug.Print "Icon: " & Icon & "; Race: " & Race & "; IMLPos: " & IMLPos
    
    GetRaceAndIcon = IMLPos
    
End Function

Public Function Conv(ByVal RawString As String) As Long
    Dim lReturn As Long
    
    If Len(RawString) = 4 Then
        Call CopyMemory(lReturn, ByVal RawString, 4)
    Else
        Debug.Print "---------- WARNING: Invalid string Length in Conv()!"
        Debug.Print "---------- Length: " & Len(RawString)
        Debug.Print DebugOutput(RawString)
    End If
    
    Conv = lReturn
End Function


'// COLORMODIFY - where L is passed as the start position of the text to be checked
Public Sub ColorModify(ByRef rtb As RichTextBox, ByRef L As Long)
    Dim i As Long
    Dim s As String
    Dim temp As Long
    
    If L = 0 Then L = 1
    
    temp = L
    
    With rtb
        If InStr(temp, .Text, "ÿc", vbTextCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "ÿc", vbTextCompare)
                
                If StrictIsNumeric(Mid$(.Text, i + 2, 1)) Then
                    s = GetColorVal(Mid$(.Text, i + 2, 1))
                    .selStart = i - 1
                    .selLength = 3
                    .SelText = vbNullString
                    .selStart = i - 1
                    .selLength = Len(.Text) - i
                    .SelColor = s
                Else
                    Select Case Mid$(.Text, i + 2, 1)
                        Case "i"
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            If .SelItalic = True Then
                                .SelItalic = False
                            Else
                                .SelItalic = True
                            End If
                            
                        Case "b", "."       'BOLD
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            If .SelBold = True Then
                                .SelBold = False
                            Else
                                .SelBold = True
                            End If
                            
                        Case "u", "."       'underline
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            If .SelUnderline = True Then
                                .SelUnderline = False
                            Else
                                .SelUnderline = True
                            End If
                            
                        Case ";"
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            .SelColor = HTMLToRGBColor("8D00CE")    'Purple
                            
                        Case ":"
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            .SelColor = 186408      '// Lighter green
                            
                        Case "<"
                            .selStart = i - 1
                            .selLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .selLength = Len(.Text) - 1
                            .SelColor = HTMLToRGBColor("00A200")    'Dark green
                        'Case Else: Debug.Print s
                    End Select
                End If
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "ÿc", vbTextCompare) > 0
            .Visible = True
        End If
        
        '// Check for SC color codes
        temp = L
        
        If InStr(temp, .Text, "Á", vbBinaryCompare) > 0 Then
            Do
                i = InStr(temp, .Text, "Á", vbBinaryCompare)
                s = GetScriptColorString(Mid$(.Text, i + 1, 1))
                
                If Len(s) > 0 Then
                    .Visible = False
                    .selStart = i - 1
                    .selLength = 2
                    .SelText = vbNullString
                    .selStart = i - 1
                    .selLength = Len(.Text) - 1
                    .SelColor = s
                    .Visible = True
                End If
                
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "Á", vbBinaryCompare) > 0
        End If
    End With
End Sub

Public Function GetScriptColorString(ByVal scCC As String) As String
    Select Case Asc(scCC)
        Case Asc("Q"): GetScriptColorString = RGB(93, 93, 93)       'Grey
        Case Asc("R"): GetScriptColorString = RGB(30, 224, 54)      'Green
        Case Asc("Z"), Asc("X"), Asc("S"): GetScriptColorString = RGB(160, 169, 116)    'Yellow
        Case Asc("Y"), Asc("["), Asc("Y"): GetScriptColorString = RGB(231, 38, 82)      'Red
        Case Asc("V"), Asc("@"): GetScriptColorString = RGB(98, 77, 232)      'Blue
        Case Asc("W"), Asc("P"): GetScriptColorString = vbWhite               'White
        Case Asc("T"), Asc("U"), Asc("V"): GetScriptColorString = HTMLToRGBColor("00CCCC") 'cyan/teal
        Case Else: GetScriptColorString = vbNullString
    End Select
End Function

Public Function GetColorVal(ByVal d2CC As String) As String
    Select Case CInt(d2CC)
        Case 1: GetColorVal = HTMLToRGBColor("CE3E3E")  'Red
        Case 2: GetColorVal = HTMLToRGBColor("00CE00")  'Green
        Case 3: GetColorVal = HTMLToRGBColor("44409C")  'Blue
        Case 4: GetColorVal = HTMLToRGBColor("A19160")  'Gold
        Case 5: GetColorVal = HTMLToRGBColor("555555")  'Grey
        Case 6: GetColorVal = HTMLToRGBColor("080808")  'Black
        Case 7: GetColorVal = HTMLToRGBColor("A89D65")  'Gold
        Case 8: GetColorVal = HTMLToRGBColor("CE8800")  'Gold-Orange
        Case 9: GetColorVal = HTMLToRGBColor("CECE51")  'Light Yellow
        Case 0: GetColorVal = HTMLToRGBColor("FFFFFF")  'White
    End Select
End Function


'Originally from DPChat by Zorm - cleaned up and adapted to my needs
Public Sub ProfileParse(Data As String)
    On Error Resume Next
    Dim x As Integer
    Dim ProfileEnd As String
    Dim SplitProfile() As String
    
    ProfileEnd = Mid(Data, 17, Len(Data))
    SplitProfile = Split(ProfileEnd, Chr(&H0))
    
    If AwaitingSystemKeys = 1 Then
        
        Event_KeyReturn "System\Account Created", SplitProfile(0)
        Event_KeyReturn "System\Last Logon", SplitProfile(1)
        Event_KeyReturn "System\Last Logoff", SplitProfile(2)
        Event_KeyReturn "System\Time Logged", SplitProfile(3)
        AwaitingSystemKeys = 0
        
    Else
    
        Event_KeyReturn "Profile\Age", SplitProfile(0)
        Event_KeyReturn "Profile\Sex", SplitProfile(1)
        Event_KeyReturn "Profile\Location", SplitProfile(2)
        Event_KeyReturn "Profile\Description", SplitProfile(3)
        
    End If
End Sub

