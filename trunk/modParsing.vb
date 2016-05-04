Option Strict Off
Option Explicit On
Module modParsing
	
	Public Const COLOR_BLUE2 As Integer = 12092001
	
	Public Sub SendHeader()
		frmChat.sckBNet.SendData(ChrW(1))
	End Sub

    Public Sub BNCSParsePacket(ByVal PacketData As String)
        BNCSParsePacket(System.Text.Encoding.Default.GetBytes(PacketData))
    End Sub

    Public Sub BNCSParsePacket(ByVal PacketData() As Byte)
        On Error GoTo ERROR_HANDLER

        Dim pD As clsDataBuffer ' Packet debuffer object
        Dim PacketLen As Integer ' Length of the packet minus the header
        Dim PacketID As Byte ' Battle.net packet ID
        Dim s As String ' Temporary string
        Dim L As Integer ' Temporary long
        Dim s2 As String ' Temporary string
        Dim s3 As String ' Temporary string
        Dim ClanTag As String ' User clan tag
        Dim Product As String ' User product
        Dim w3icon As String ' Warcraft III icon code
        Dim B As Boolean ' Temporary bool
        Dim sArr() As String ' Temp String array
        Dim veto As Boolean


        '--------------
        '| Initialize |
        '--------------
        pD = New clsDataBuffer
        PacketLen = Len(PacketData) - 4

        '###########################################################################

        If PacketLen >= 0 Then
            ' Start packet debuffer
            Buffer.BlockCopy(PacketData, 4, pD.Data, 0, PacketLen)

            ' Get packet ID
            PacketID = PacketData(1)

            If MDebug("all") Then
                frmChat.AddChat(COLOR_BLUE, "BNET RECV 0x" & ZeroOffset(PacketID, 2))
            End If

            CachePacket(modEnum.enuPL_DirectionTypes.StoC, modEnum.enuPL_ServerTypes.stBNCS, PacketID, Len(PacketData), PacketData)

            ' Added 2007-06-08 for a packet logging menu feature to aid tech support
            WritePacketData(modEnum.enuPL_ServerTypes.stBNCS, modEnum.enuPL_DirectionTypes.StoC, PacketID, PacketLen, PacketData)

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
                    ProfileParse(System.Text.Encoding.Default.GetString(PacketData))

                    '###########################################################################
                Case Is >= &H65 'Friends List or Clan-related packet
                    ' Hand the packet off to the appropriate handler
                    If PacketID >= &H70 Then
                        ' added in response to the clan channel takeover exploit
                        ' discovered 11/7/05
                        If IsW3() Then
                            frmChat.ParseClanPacket(PacketID, IIf(PacketLen > 0, pD.Data, vbNullString))
                        End If
                    Else
                        If (g_request_receipt) Then
                            g_request_receipt = False

                            If (Caching) Then
                                frmChat.cacheTimer_Tick(Nothing, New System.EventArgs())
                            End If

                            Exit Sub
                        End If

                        frmChat.ParseFriendsPacket(PacketID, pD.Data)
                    End If

                    '###########################################################################
                Case Else
                    Call modBNCS.BNCSRecvPacket(PacketData)

            End Select
        End If

        'UPGRADE_NOTE: Object pD may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pD = Nothing

        Exit Sub

ERROR_HANDLER:
        frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in BNCSParsePacket().")

        Exit Sub
    End Sub
	
	Public Function StrToHex(ByVal String1 As String, Optional ByVal NoSpaces As Boolean = False) As String
		Dim strTemp, strReturn As String
		Dim i As Integer
		
		For i = 1 To Len(String1)
			strTemp = Hex(Asc(Mid(String1, i, 1)))
			If Len(strTemp) = 1 Then strTemp = "0" & strTemp
			
			strReturn = strReturn & IIf(NoSpaces, "", Space(1)) & strTemp
		Next i
		
		StrToHex = strReturn
	End Function
	
	Public Function RShift(ByVal pnValue As Integer, ByVal pnShift As Integer) As Double
		'on error resume next
		RShift = CDbl(pnValue \ (2 ^ pnShift))
	End Function
	
	Public Function KillNull(ByVal Text As String) As String
		Dim i As Short
		i = InStr(1, Text, Chr(0))
		If (i = 0) Then
			KillNull = Text
			Exit Function
		End If
		KillNull = Left(Text, i - 1)
	End Function
	
	Public Function GetHexValue(ByVal v As Integer) As String
		
		v = v And &HF
		
		If v < 10 Then
			
			GetHexValue = Chr(v + &H30)
			
		Else
			
			GetHexValue = Chr(v + &H37)
			
		End If
		
	End Function
	
	Public Function GetNumValue(ByVal c As String) As Integer
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
		Dim i As Short
		
		i = InStr(Text, Chr(0))
		If i = 0 Then Exit Sub
		
		Text = Left(Text, i - 1)
	End Sub
	
	Public Sub FullJoin(ByRef Channel As String, Optional ByVal i As Integer = -1)
		If i >= 0 Then
			PBuffer.InsertDWord(CInt(i))
		Else
			PBuffer.InsertDWord(&H2)
		End If
		PBuffer.InsertNTString(Channel)
		PBuffer.SendPacket(SID_JOINCHANNEL)
	End Sub
	
	Public Function HexToStr(ByVal Hex1 As String) As String
		'on error resume next
		Dim strReturn As String
		Dim i As Integer
		If Len(Hex1) Mod 2 <> 0 Then Exit Function
		For i = 1 To Len(Hex1) Step 2
			strReturn = strReturn & Chr(Val("&H" & Mid(Hex1, i, 2)))
		Next i
		HexToStr = strReturn
	End Function
	
	Public Sub RejoinChannel(ByRef Channel As String)
		'on error resume next
		PBuffer.SendPacket(SID_LEAVECHAT)
		PBuffer.InsertDWord(&H2)
		PBuffer.InsertNTString(Channel)
		PBuffer.SendPacket(SID_JOINCHANNEL)
	End Sub
	
	Public Sub RequestProfile(ByRef strUser As String)
		'on error resume next
		With PBuffer
			.InsertDWord(1)
			.InsertDWord(4)
			.InsertDWord(GetTickCount())
			.InsertNTString(CleanUsername(ReverseConvertUsernameGateway(strUser)))
			.InsertNTString("Profile\Age")
			.InsertNTString("Profile\Sex")
			.InsertNTString("Profile\Location")
			.InsertNTString("Profile\Description")
			.SendPacket(SID_READUSERDATA)
		End With
	End Sub
	
	Public Sub RequestSpecificKey(ByVal sUsername As String, ByVal sKey As String)
		With PBuffer
			.InsertDWord(1)
			.InsertDWord(1)
			.InsertDWord(GetTickCount())
			.InsertNTString(ReverseConvertUsernameGateway(sUsername))
			.InsertNTString(sKey)
			.SendPacket(SID_READUSERDATA)
		End With
	End Sub
	
	Public Sub SetProfile(ByVal Location As String, ByVal Description As String, Optional ByVal Sex As String = vbNullString)
		'Dim i As Byte
		Const MAX_DESCR As Integer = 510
		Const MAX_SEX As Integer = 200
		Const MAX_LOC As Integer = 200
		
		'// Sanity checks
		If Len(Description) > MAX_DESCR Then
			Description = Left(Description, MAX_DESCR)
		End If
		
		If Len(Sex) > MAX_SEX Then
			Sex = Left(Sex, MAX_SEX)
		End If
		
		If Len(Location) > MAX_LOC Then
			Location = Left(Location, MAX_LOC)
		End If
		
		
		With PBuffer
			.InsertDWord(&H1) '// #accounts
			.InsertDWord(3) '// #keys
			
			.InsertNTString(CurrentUsername) '// account to update
			'// keys
			.InsertNTString("Profile\Location")
			.InsertNTString("Profile\Description")
			.InsertNTString("Profile\Sex")
			'// Values()
			.InsertNTString(Location)
			.InsertNTString(Description)
			.InsertNTString(Sex)
			
			.SendPacket(SID_WRITEUSERDATA)
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
		Const MAX_DESCR As Integer = 510
		Const MAX_SEX As Integer = 200
		Const MAX_LOC As Integer = 200
		
		Dim nKeys, i As Short
		'UPGRADE_WARNING: Lower bound of array pKeys was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim pKeys(3) As String
		'UPGRADE_WARNING: Lower bound of array pData was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim pData(3) As String
		
        If (Len(Location) > 0) Then
            If (Len(Location) > MAX_LOC) Then
                Location = Left(Location, MAX_LOC)
            End If

            nKeys = nKeys + 1
            pKeys(nKeys) = "Profile\Location"
            pData(nKeys) = Location
        End If
		
		'// Sanity checks
        If (Len(Description) > 0) Then
            If (Len(Description) > MAX_DESCR) Then
                Description = Left(Description, MAX_DESCR)
            End If

            nKeys = nKeys + 1
            pKeys(nKeys) = "Profile\Description"
            pData(nKeys) = Description
        End If
		
		If nKeys > 0 Then
			With PBuffer
				.InsertDWord(&H1) '// #accounts
				.InsertDWord(nKeys) '// #keys
				.InsertNTString(CurrentUsername) '// account to update
				'// keys
				For i = 1 To nKeys
					.InsertNTString(pKeys(i))
				Next i
				
				'// Values()
				For i = 1 To nKeys
					.InsertNTString(pData(i))
				Next i
				
				.SendPacket(SID_WRITEUSERDATA)
			End With
		End If
	End Sub
	
	Public Sub sPrintF(ByRef source As String, ByVal nText As String, Optional ByVal a As Object = Nothing, Optional ByVal B As Object = Nothing, Optional ByVal c As Object = Nothing, Optional ByVal d As Object = Nothing, Optional ByVal e As Object = Nothing, Optional ByVal f As Object = Nothing, Optional ByVal g As Object = Nothing, Optional ByVal H As Object = Nothing)
		
		nText = Replace(nText, "%S", "%s")
		
		Dim i As Byte
		i = 0
		
		Do While (InStr(1, nText, "%s") <> 0)
			Select Case i
				Case 0
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(a) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", a, 1, 1)
				Case 1
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(B) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", B, 1, 1)
				Case 2
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(c) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object c. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", c, 1, 1)
				Case 3
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(d) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object d. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", d, 1, 1)
				Case 4
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(e) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object e. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", e, 1, 1)
				Case 5
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(f) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object f. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", f, 1, 1)
				Case 6
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(g) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nText = Replace(nText, "%s", g, 1, 1)
				Case 7
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsNothing(H) Then GoTo theEnd
					'UPGRADE_WARNING: Couldn't resolve default property of object H. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_NOTE: cType was upgraded to cType_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim cType_Renamed As String
		Dim WCG As Boolean
		On Error GoTo ParseStatString_Error
		
		'Dim Icon As String
		
		' FRCW = Ref
		' LPCW = Player
		
		'Debug.Print "Received statstring: " & Statstring
        If Len(Statstring) > 0 Then

            Select Case Left(Statstring, 4)
                Case "3RAW", "PX3W"
                    If Len(Statstring) > 4 Then
                        temp = Split(Statstring, " ")

                        ReDim Values(3)

                        If StrComp(Right(temp(1), 2), "CW") = 0 Then
                            WCG = True
                        Else
                            Values(1) = Mid(Statstring, 6, 1)
                            Values(2) = Mid(Statstring, 7, 1)
                        End If

                        Values(0) = temp(2)

                        If UBound(temp) > 2 Then
                            Values(3) = StrReverse(temp(3))
                        End If

                        sClan = IIf(UBound(Values) > 2, Values(3), "")

                        If Left(Statstring, 4) = "3RAW" Then
                            Call sPrintF(outbuf, "Warcraft III: Reign of Chaos (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                        Else
                            Call sPrintF(outbuf, "Warcraft III: The Frozen Throne (Level: %s, icon tier %s, %s icon" & IIf(UBound(temp) > 2, ", in Clan " & sClan, vbNullString) & ")", Values(0), Values(2), Values(1))
                        End If
                    Else
                        If Left(Statstring, 4) = "3RAW" Then
                            Call StrCpy(outbuf, "Warcraft III: Reign of Chaos.")
                        Else
                            Call StrCpy(outbuf, "Warcraft III: The Frozen Throne.")
                        End If
                    End If

                Case "RHSS"
                    Call StrCpy(outbuf, "Starcraft Shareware.")

                Case "RATS"
                    Values = Split(Mid(Statstring, 6), " ")
                    If UBound(Values) <> 8 Then
                        Call sPrintF(outbuf, "a Starcraft %sbot", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString))
                    Else
                        If CDbl(Values(0)) > 0 Then
                            Call sPrintF(outbuf, "Starcraft%s (%s wins, with a rating of %s on the ladder)", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                        Else
                            Call sPrintF(outbuf, "Starcraft%s (%s wins).", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2))
                        End If
                    End If

                Case "PXES"
                    Values = Split(Mid(Statstring, 6), " ")
                    If UBound(Values) <> 8 Then
                        Call sPrintF(outbuf, "a Starcraft Brood War bot.", vbNullString)

                        If UBound(Values) > 2 Then
                            outbuf = outbuf & "(spawn) "
                        End If
                    Else
                        If CDbl(Values(0)) > 0 Then
                            Call sPrintF(outbuf, "Starcraft Brood War%s (%s wins, with a rating of %s on the ladder)", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                        Else
                            Call sPrintF(outbuf, "Starcraft Brood War%s (%s wins).", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2))
                        End If
                    End If

                Case "RTSJ"
                    Values = Split(Mid(Statstring, 6), " ")
                    If UBound(Values) <> 8 Then
                        Call sPrintF(outbuf, "a Starcraft Japanese %sbot.", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString))
                    Else
                        If CDbl(Values(0)) > 0 Then
                            Call sPrintF(outbuf, "Starcraft Japanese%s (%s wins, with a rating of %s on the ladder)", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                        Else
                            Call sPrintF(outbuf, "Starcraft Japanese%s (%s wins).", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2))
                        End If
                    End If

                Case "NB2W"
                    Values = Split(Mid(Statstring, 6), " ")

                    If UBound(Values) <> 8 Then
                        Call sPrintF(outbuf, "a Warcraft II %sbot.", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString))
                    Else
                        If CDbl(Values(0)) > 0 Then
                            Call sPrintF(outbuf, "Warcraft II%s (%s wins, with a rating of %s on the ladder)", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
                        Else
                            Call sPrintF(outbuf, "Warcraft II%s (%s wins).", IIf((CDbl(Values(3)) = 1), " (spawn) ", vbNullString), Values(2))
                        End If
                    End If

                Case "RHSD"
                    Values = Split(Mid(Statstring, 6), " ")
                    If UBound(Values) <> 8 Then
                        Call StrCpy(outbuf, "A Diablo shareware bot.")
                    Else
                        Select Case Values(1)
                            Case CStr(0) : cType_Renamed = "warrior"
                            Case CStr(1) : cType_Renamed = "rogue"
                            Case CStr(2) : cType_Renamed = "sorceror"
                        End Select
                        Call sPrintF(outbuf, "Diablo shareware (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold)", Values(0), cType_Renamed, Values(2), Values(3), Values(4), Values(5), Values(6), Values(7))
                    End If

                Case "LTRD"
                    Values = Split(Mid(Statstring, 6), " ")

                    If UBound(Values) <> 8 Then
                        Call StrCpy(outbuf, "A Diablo bot.")
                    Else
                        Select Case Values(1)
                            Case CStr(0) : cType_Renamed = "warrior"
                            Case CStr(1) : cType_Renamed = "rogue"
                            Case CStr(2) : cType_Renamed = "sorceror"
                        End Select
                        Call sPrintF(outbuf, "Diablo (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold)", Values(0), cType_Renamed, Values(2), Values(3), Values(4), Values(5), Values(6), Values(7))
                    End If

                Case "PX2D"
                    'UPGRADE_WARNING: Couldn't resolve default property of object ParseD2Stats(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call StrCpy(outbuf, ParseD2Stats(Statstring))

                Case "VD2D"
                    'UPGRADE_WARNING: Couldn't resolve default property of object ParseD2Stats(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call StrCpy(outbuf, ParseD2Stats(Statstring))

                Case "TAHC"
                    Call StrCpy(outbuf, "a Chat bot.")

                Case Else
                    Call StrCpy(outbuf, "an unknown client.")

            End Select

            ParseStatstring = StrReverse(Left(Statstring, 4))

        End If
		
ParseStatString_Exit: 
		Exit Function
		
ParseStatString_Error: 
		
		Debug.Print("Error " & Err.Number & " (" & Err.Description & ") in procedure ParseStatString of Module modParsing")
		outbuf = "- Error parsing statstring. [" & Replace(Statstring, Chr(0), "") & "]"
		
		Resume ParseStatString_Exit
	End Function
	
	' This code cleaned 3/4/2005
	Public Function ParseD2Stats(ByVal Stats As String) As Object
		Dim Female, Expansion As Boolean
		Dim Hardcore, Version, sLen, CharClass, CharLevel As Byte
		Dim Server, StatBuf, Name As String
		Dim P() As String
		
		Dim D2Classes(7) As String
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
			Call MakeArray(Mid(Stats, sLen), P)
		End If
		
		If Left(Stats, 4) = "VD2D" Then
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
			
			If Left(Stats, 4) = "PX2D" Then
				If (Asc(P(26)) And &H20) Then
					Select Case RShift(Asc(P(27)) And &H18, 3)
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
				Select Case RShift(Asc(P(27)) And &H18, 3)
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseD2Stats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseD2Stats = StatBuf
	End Function
	
	Public Function GetServer(ByVal Statstring As String, ByRef Server As String) As Byte
		'returns the begining of the character name
		Server = Mid(Statstring, 5, InStr(5, Statstring, ",") - 5)
		GetServer = InStr(5, Statstring, ",") + 1
	End Function
	
	Public Function GetCharacterName(ByVal Statstring As String, ByVal Start As Byte, ByRef cName As String) As Byte
		cName = Mid(Statstring, Start, InStr(Start, Statstring, ",") - Start)
		GetCharacterName = InStr(Start, Statstring, ",") + 1
	End Function
	
	Public Sub StrCpy(ByRef source As String, ByVal nText As String)
		'on error resume next
		source = source & nText
	End Sub
	
	Public Sub MakeArray(ByVal Text As String, ByRef nArray() As String)
		Dim i As Integer
		ReDim nArray(0)
		For i = 0 To Len(Text)
			nArray(i) = Mid(Text, i + 1, 1)
			If i <> Len(Text) Then
				ReDim Preserve nArray(UBound(nArray) + 1)
			End If
		Next i
	End Sub
	
	Public Function Conv(ByVal RawString As String) As Integer
		Dim lReturn As Integer
		
		If Len(RawString) = 4 Then
			Call CopyMemory(lReturn, RawString, 4)
		Else
			Debug.Print("---------- WARNING: Invalid string Length in Conv()!")
			Debug.Print("---------- Length: " & Len(RawString))
			Debug.Print(DebugOutput(RawString))
		End If
		
		Conv = lReturn
	End Function
	
	
	'// COLORMODIFY - where L is passed as the start position of the text to be checked
	Public Sub ColorModify(ByRef rtb As System.Windows.Forms.RichTextBox, ByRef L As Integer)
		Dim i As Integer
		Dim s As String
		Dim temp As Integer
		Dim SelStart As Integer
		Dim SelLength As Integer
		
		If L = 0 Then L = 1
		
		temp = L
		
		With rtb
			' store previous selstart and len
			SelStart = .SelectionStart
			SelLength = .SelectionLength
			
			If InStr(temp, .Text, "ÿc", CompareMethod.Text) > 0 Then
				.Visible = False
				Do 
					i = InStr(temp, .Text, "ÿc", CompareMethod.Text)
					
					If StrictIsNumeric(Mid(.Text, i + 2, 1)) Then
						s = GetColorVal(Mid(.Text, i + 2, 1))
						.SelectionStart = i - 1
						.SelectionLength = 3
						.SelectedText = vbNullString
						.SelectionStart = i - 1
						.SelectionLength = Len(.Text) + 1 - i
						.SelectionColor = System.Drawing.ColorTranslator.FromOle(CInt(s))
					Else
						Select Case Mid(.Text, i + 2, 1)
							Case "i"
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								If .SelectionFont.Italic = True Then
									.SelectionFont = VB6.FontChangeItalic(.SelectionFont, False)
								Else
									.SelectionFont = VB6.FontChangeItalic(.SelectionFont, True)
								End If
								
							Case "b", "." 'BOLD
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								If .SelectionFont.Bold = True Then
									.Font = VB6.FontChangeBold(.SelectionFont, False)
								Else
									.Font = VB6.FontChangeBold(.SelectionFont, True)
								End If
								
							Case "u", "." 'underline
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								If .SelectionFont.Underline = True Then
									.SelectionFont = VB6.FontChangeUnderline(.SelectionFont, False)
								Else
									.SelectionFont = VB6.FontChangeUnderline(.SelectionFont, True)
								End If
								
							Case ";"
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								.SelectionColor = System.Drawing.ColorTranslator.FromOle(HTMLToRGBColor("8D00CE")) 'Purple
								
							Case ":"
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								.SelectionColor = System.Drawing.ColorTranslator.FromOle(186408) '// Lighter green
								
							Case "<"
								.SelectionStart = i - 1
								.SelectionLength = 3
								.SelectedText = vbNullString
								.SelectionStart = i - 1
								.SelectionLength = Len(.Text) + 1 - 1
								.SelectionColor = System.Drawing.ColorTranslator.FromOle(HTMLToRGBColor("00A200")) 'Dark green
								'Case Else: Debug.Print s
						End Select
					End If
					temp = temp + 1
					
				Loop While InStr(temp, .Text, "ÿc", CompareMethod.Text) > 0
				.Visible = True
			End If
			
			'// Check for SC color codes
			temp = L
			
			If InStr(temp, .Text, "Á", CompareMethod.Binary) > 0 Then
				.Visible = False
				Do 
					i = InStr(temp, .Text, "Á", CompareMethod.Binary)
					s = GetScriptColorString(Mid(.Text, i + 1, 1))
					
					If Len(s) > 0 Then
						.Visible = False
						.SelectionStart = i - 1
						.SelectionLength = 2
						.SelectedText = vbNullString
						.SelectionStart = i - 1
						.SelectionLength = Len(.Text) + 1 - 1
						.SelectionColor = System.Drawing.ColorTranslator.FromOle(CInt(s))
						.Visible = True
					End If
					
					temp = temp + 1
					
				Loop While InStr(temp, .Text, "Á", CompareMethod.Binary) > 0
				.Visible = True
			End If
			
			' restore previous selstart and len
			.SelectionStart = SelStart
			.SelectionLength = SelLength
		End With
	End Sub
	
	Public Function GetScriptColorString(ByVal scCC As String) As String
		Select Case Asc(scCC)
			Case Asc("Q") : GetScriptColorString = CStr(RGB(93, 93, 93)) 'Grey
			Case Asc("R") : GetScriptColorString = CStr(RGB(30, 224, 54)) 'Green
			Case Asc("Z"), Asc("X"), Asc("S") : GetScriptColorString = CStr(RGB(160, 169, 116)) 'Yellow
			Case Asc("Y"), Asc("["), Asc("Y") : GetScriptColorString = CStr(RGB(231, 38, 82)) 'Red
			Case Asc("V"), Asc("@") : GetScriptColorString = CStr(RGB(98, 77, 232)) 'Blue
			Case Asc("W"), Asc("P") : GetScriptColorString = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White).ToString 'White
			Case Asc("T"), Asc("U"), Asc("V") : GetScriptColorString = CStr(HTMLToRGBColor("00CCCC")) 'cyan/teal
			Case Else : GetScriptColorString = vbNullString
		End Select
	End Function
	
	Public Function GetColorVal(ByVal d2CC As String) As String
		Select Case CShort(d2CC)
			Case 1 : GetColorVal = CStr(HTMLToRGBColor("CE3E3E")) 'Red
			Case 2 : GetColorVal = CStr(HTMLToRGBColor("00CE00")) 'Green
			Case 3 : GetColorVal = CStr(HTMLToRGBColor("44409C")) 'Blue
			Case 4 : GetColorVal = CStr(HTMLToRGBColor("A19160")) 'Gold
			Case 5 : GetColorVal = CStr(HTMLToRGBColor("555555")) 'Grey
			Case 6 : GetColorVal = CStr(HTMLToRGBColor("080808")) 'Black
			Case 7 : GetColorVal = CStr(HTMLToRGBColor("A89D65")) 'Gold
			Case 8 : GetColorVal = CStr(HTMLToRGBColor("CE8800")) 'Gold-Orange
			Case 9 : GetColorVal = CStr(HTMLToRGBColor("CECE51")) 'Light Yellow
			Case 0 : GetColorVal = CStr(HTMLToRGBColor("FFFFFF")) 'White
		End Select
	End Function
	
	
	'Originally from DPChat by Zorm - cleaned up and adapted to my needs
	Public Sub ProfileParse(ByRef Data As String)
		On Error Resume Next
		Dim x As Short
		Dim ProfileEnd As String
		Dim SplitProfile() As String
		
		ProfileEnd = Mid(Data, 17, Len(Data))
		SplitProfile = Split(ProfileEnd, Chr(&H0))
		
		If (UBound(SplitProfile) = 4) Then
			
			If AwaitingSystemKeys = 1 Then
				
				Event_KeyReturn("System\Account Created", SplitProfile(0))
				Event_KeyReturn("System\Last Logon", SplitProfile(1))
				Event_KeyReturn("System\Last Logoff", SplitProfile(2))
				Event_KeyReturn("System\Time Logged", SplitProfile(3))
				AwaitingSystemKeys = 0
				
			Else
				
				Event_KeyReturn("Profile\Age", SplitProfile(0))
				Event_KeyReturn("Profile\Sex", SplitProfile(1))
				Event_KeyReturn("Profile\Location", SplitProfile(2))
				Event_KeyReturn("Profile\Description", SplitProfile(3))
				
			End If
			
		Else
			
			' for SSC.RequestProfileKey()
			Event_KeyReturn(SpecificProfileKey, SplitProfile(0))
			
		End If
	End Sub
End Module