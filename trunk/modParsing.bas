Attribute VB_Name = "modParsing"
Option Explicit

Public Const COLOR_BLUE2 = 12092001

Public Sub SendHeader()
    frmChat.sckBNet.SendData ChrW(1)
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
    PBuffer.SendPacket SID_JOINCHANNEL
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
    PBuffer.SendPacket SID_LEAVECHAT
    PBuffer.InsertDWord &H2
    PBuffer.InsertNTString Channel
    PBuffer.SendPacket SID_JOINCHANNEL
End Sub

Public Sub RequestProfile(strUser As String, ByVal eType As enuUserDataRequestType, Optional ByRef oCommand As clsCommandObj)
    Dim aKeys(3) As String
    
    aKeys(0) = "Profile\Age"
    aKeys(1) = "Profile\Sex"
    aKeys(2) = "Profile\Location"
    aKeys(3) = "Profile\Description"
    
    Call RequestUserData(strUser, aKeys, eType, oCommand)
End Sub

Public Sub RequestUserData(ByVal sUsername As String, ByRef aKeys() As String, Optional ByVal eType As enuUserDataRequestType, Optional ByRef oCommand As clsCommandObj)
    Dim oRequest As udtUserDataRequest
    Dim i As Integer
    Dim bFoundSlot As Boolean

    With oRequest
        ' Attach handling info
        .RequestType = eType
        Set .Command = oCommand
        .ResponseReceived = False
    
        ' Add request data
        .Account = sUsername
        .keys = aKeys
    End With
    
    ' Find an open slot in the request list
    bFoundSlot = False
    If UBound(UserDataRequests) > 0 Then
        For i = 1 To UBound(UserDataRequests)
            If UserDataRequests(i).ResponseReceived Then
                bFoundSlot = True
            
                oRequest.RequestID = i
                UserDataRequests(i) = oRequest
                Exit For
            End If
        Next
    End If
    
    ' If no slot was found, add the request to the end
    If Not bFoundSlot Then
        oRequest.RequestID = UBound(UserDataRequests) + 1
        
        ReDim Preserve UserDataRequests(oRequest.RequestID)
        UserDataRequests(oRequest.RequestID) = oRequest
    End If

    ' Build the packet
    With PBuffer
        .InsertDWord 1
        .InsertDWord UBound(oRequest.keys) + 1
        .InsertDWord oRequest.RequestID
        
        .InsertNTString CleanUsername(ReverseConvertUsernameGateway(oRequest.Account))
        
        For i = 0 To UBound(aKeys)
            .InsertNTString oRequest.keys(i)
        Next

        .SendPacket SID_READUSERDATA
    End With
End Sub

Public Sub SetProfile(ByVal Location As String, ByVal Description As String, Optional ByVal Sex As String = vbNullString)
    'Dim i As Byte
    Const MAX_DESCR As Long = 510
    Const MAX_SEX As Long = 200
    Const MAX_LOC As Long = 200
    
    '// Sanity checks
    If Len(Description) > MAX_DESCR Then
        Description = Left$(Description, MAX_DESCR)
    End If
    
    If Len(Sex) > MAX_SEX Then
        Sex = Left$(Sex, MAX_SEX)
    End If
    
    If Len(Location) > MAX_LOC Then
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
        .InsertNTString Description
        .InsertNTString Sex
        
        .SendPacket SID_WRITEUSERDATA
    End With
End Sub

'// Extended version of this function for scripting use
'//  Will not ERASE if a field is left blank
'// 2007-06-07: SEX value is ignored because Blizzard removed that
'//     field from profiles
'// 2009-07-14: corrected a problem in this method, thanks Jack (t=42494) -andy
'//     method was erasing profile data
Public Sub SetProfileEx(ByVal Location As String, ByVal Description As String)
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
    If (LenB(Description) > 0) Then
        If (Len(Description) > MAX_DESCR) Then
            Description = Left$(Description, MAX_DESCR)
        End If
        
        nKeys = nKeys + 1
        pKeys(nKeys) = "Profile\Description"
        pData(nKeys) = Description
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
            
            .SendPacket SID_WRITEUSERDATA
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
                    
                    sClan = IIf(UBound(Values) > 2, Values(3), "")
                    
                    If Left$(Statstring, 4) = "3RAW" Then
                        Call sPrintF(outbuf, "Warcraft III: Reign of Chaos (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                    Else
                        Call sPrintF(outbuf, "Warcraft III: The Frozen Throne (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                    End If
                Else
                    If Left$(Statstring, 4) = "3RAW" Then
                        Call StrCpy(outbuf, "Warcraft III: Reign of Chaos.")
                    Else
                        Call StrCpy(outbuf, "Warcraft III: The Frozen Throne.")
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

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ParseStatString of Module modParsing"
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

Public Function Conv(ByVal RawString As String) As Long
    Dim lReturn As Long
    
    If Len(RawString) = 4 Then
        Call CopyMemory(lReturn, ByVal RawString, 4)
    Else
        Debug.Print "---------- WARNING! Invalid string Length in Conv()!"
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
    Dim SelStart As Long
    Dim SelLength As Long
    
    If L = 0 Then L = 1
    
    temp = L
    
    With rtb
        ' store previous selstart and len
        SelStart = .SelStart
        SelLength = .SelLength
        
        If InStr(temp, .Text, "ÿc", vbTextCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "ÿc", vbTextCompare)
                
                If StrictIsNumeric(Mid$(.Text, i + 2, 1)) Then
                    s = GetColorVal(Mid$(.Text, i + 2, 1))
                    .SelStart = i - 1
                    .SelLength = 3
                    .SelText = vbNullString
                    .SelStart = i - 1
                    .SelLength = Len(.Text) + 1 - i
                    .SelColor = s
                Else
                    Select Case Mid$(.Text, i + 2, 1)
                        Case "i"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelItalic = True Then
                                .SelItalic = False
                            Else
                                .SelItalic = True
                            End If
                            
                        Case "b", "."       'BOLD
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelBold = True Then
                                .SelBold = False
                            Else
                                .SelBold = True
                            End If
                            
                        Case "u", "."       'underline
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelUnderline = True Then
                                .SelUnderline = False
                            Else
                                .SelUnderline = True
                            End If
                            
                        Case ";"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = HTMLToRGBColor("8D00CE")    'Purple
                            
                        Case ":"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = 186408      '// Lighter green
                            
                        Case "<"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
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
            .Visible = False
            Do
                i = InStr(temp, .Text, "Á", vbBinaryCompare)
                s = GetScriptColorString(Mid$(.Text, i + 1, 1))
                
                If Len(s) > 0 Then
                    .Visible = False
                    .SelStart = i - 1
                    .SelLength = 2
                    .SelText = vbNullString
                    .SelStart = i - 1
                    .SelLength = Len(.Text) + 1 - 1
                    .SelColor = s
                    .Visible = True
                End If
                
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "Á", vbBinaryCompare) > 0
            .Visible = True
        End If
        
        ' restore previous selstart and len
        .SelStart = SelStart
        .SelLength = SelLength
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


