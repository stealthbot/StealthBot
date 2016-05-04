Option Strict Off
Option Explicit On
Friend Class clsClanPacketHandler
	
	'clsClanPacketHandler - project StealthBot - authored by Stealth (stealth@stealthbot.net)
	
	'Special thanks:
	'-  Ethereal packetlogger was used in my own research
	'-  thanks to Arta[vL] and BNetDocs (http://bnetdocs.valhallalegends.com) for additional assistance
	
	Const SID_FINDCLANCANDIDATES As Integer = &H70
	'// Const SID_INVITEMULTIPLEUSERS& = &H71 -- not supported
	Const SID_DISBANDCLAN As Integer = &H73
	Const SID_CLANINFO As Integer = &H75
	Const SID_REMOVEDFROMCLAN As Integer = &H76
	Const SID_CLANREQUEST As Integer = &H77
	Const SID_REMOVEMEMBER As Integer = &H78
	Const SID_CLANINVITE As Integer = &H79
	Const SID_CLANINVITE2 As Integer = &H72
	Const SID_CLANMEMBERLIST As Integer = &H7D
	Const SID_CLANMEMBERUPDATE As Integer = &H7F
	Const SID_MEMBERLEFT As Integer = &H7E
	Const SID_CHANGERANK As Integer = &H7A
	Const SID_NEWRANKRECEIVED As Integer = &H81
	Const SID_CLANUSERINFO As Integer = &H82 '// arbitrary name
	Const SID_CLANMOTD As Integer = &H7C
	
	
	Public Event CandidateList(ByVal Status As Byte, ByRef Users As System.Array)
	Public Event DisbandClanReply(ByVal Success As Boolean)
	Public Event ClanInfo(ByVal ClanTag As String, ByVal RawClanTag As String, ByVal Rank As Byte)
	Public Event InviteUserReply(ByVal Status As Byte)
	Public Event ClanInvitation(ByVal Token As String, ByVal ClanTag As String, ByVal RawClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal NewClan As Boolean)
	Public Event ClanMemberUpdate(ByVal Username As String, ByVal Rank As Byte, ByVal IsOnline As Byte, ByVal Location As String)
	Public Event ClanMemberList(ByRef Members As System.Array)
	Public Event UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)
	Public Event DemoteUserReply(ByVal Success As Boolean)
	Public Event PromoteUserReply(ByVal Success As Boolean)
	Public Event RemoveUserReply(ByVal Result As Byte)
	Public Event MyRankChange(ByVal NewRank As Byte)
	Public Event MemberLeaves(ByVal Member As String)
	Public Event RemovedFromClan(ByVal Status As Byte)
	Public Event ClanMOTD(ByVal cookie As Integer, ByVal Message As String)
	
	'10-18-07 - Hdx - Changed to use clsPacketDebuffer
    Public Sub ParseClanPacket(ByVal PacketID As Byte, ByVal Data() As Byte)
        On Error GoTo ERROR_HANDLER

        Dim ary() As String
        'Dim dwTemp As Long
        Dim ClanTag As New VB6.FixedLengthString(4)
        Dim sTemp As String ', Token As String
        Dim sTemp2 As String
        Dim iTemp As Short
        Dim bTemp As Byte
        Dim bRank As Byte
        Dim bStatus As Byte

        Dim inBuf As New clsDataBuffer
        inBuf.Data = Data

        Dim cookie As Integer
        Dim Message As String
        Select Case PacketID

            Case SID_FINDCLANCANDIDATES 'Clan candidates
                inBuf.GetDWORD() 'Cookie
                bStatus = inBuf.GetByte
                bRank = inBuf.GetByte
                If (bRank > 0) Then
                    ReDim ary(bRank - 1)
                    For iTemp = 1 To bRank
                        ary(iTemp - 1) = inBuf.GetString
                    Next iTemp
                Else
                    ReDim ary(0)
                    ary(0) = vbNullString
                End If
                RaiseEvent CandidateList(bStatus, ary)


            Case SID_CLANMEMBERUPDATE 'Clan Info Update
                sTemp = inBuf.GetString
                bRank = inBuf.GetByte
                bStatus = inBuf.GetByte
                sTemp2 = inBuf.GetString
                RaiseEvent ClanMemberUpdate(sTemp, bRank, bStatus, sTemp2)


            Case SID_CLANINFO 'Clan Info
                inBuf.GetByte() 'Unknown (0)
                ClanTag.Value = inBuf.GetRaw(4)
                bRank = inBuf.GetByte
                RaiseEvent ClanInfo(KillNull(StrReverse(ClanTag.Value)), ClanTag.Value, bRank)

            Case SID_CHANGERANK 'Action Response

                'demote: 1; promote: 3

                iTemp = inBuf.GetDWORD
                bStatus = inBuf.GetByte

                Select Case bStatus
                    Case 0 'success
                        Select Case iTemp
                            Case 1 : RaiseEvent DemoteUserReply(True)
                            Case 3 : RaiseEvent PromoteUserReply(True)
                        End Select

                    Case 2, 7, 8 'too soon
                        Select Case iTemp
                            Case 1 : RaiseEvent DemoteUserReply(False)
                            Case 3 : RaiseEvent PromoteUserReply(False)
                        End Select

                End Select

            Case SID_CLANMEMBERLIST 'Clan listing
                inBuf.GetDWORD()
                bTemp = inBuf.GetByte
                ReDim ary(bTemp * 4 - 1)
                For iTemp = 0 To bTemp - 1
                    ary(iTemp * 4) = inBuf.GetString()
                    ary(iTemp * 4 + 1) = CStr(inBuf.GetByte())
                    ary(iTemp * 4 + 2) = CStr(inBuf.GetByte())
                    ary(iTemp * 4 + 3) = inBuf.GetString()
                Next iTemp
                RaiseEvent ClanMemberList(ary)


            Case SID_REMOVEMEMBER
                inBuf.GetDWORD()
                RaiseEvent RemoveUserReply(inBuf.GetByte)

            Case SID_MEMBERLEFT : RaiseEvent MemberLeaves(inBuf.GetString)
            Case SID_REMOVEDFROMCLAN : RaiseEvent RemovedFromClan(inBuf.GetByte)

            Case SID_DISBANDCLAN
                inBuf.GetDWORD() 'cookie
                RaiseEvent DisbandClanReply((inBuf.GetByte = 0))


            Case SID_CLANINVITE, SID_CLANINVITE2

                sTemp = inBuf.GetRaw(4)
                ClanTag.Value = inBuf.GetRaw(4)
                ReDim ary(2)
                ary(0) = inBuf.GetString
                ary(1) = inBuf.GetString

                RaiseEvent ClanInvitation(sTemp, StrReverse(ClanTag.Value), ClanTag.Value, ary(0), ary(1), (PacketID = SID_CLANINVITE2))

            Case SID_CLANREQUEST
                inBuf.GetDWORD()
                RaiseEvent InviteUserReply(inBuf.GetByte)

            Case SID_NEWRANKRECEIVED
                '            (BYTE) - Old rank
                '            (BYTE) - New rank
                '            (STRING) - User who changed your rank
                inBuf.GetByte()
                RaiseEvent MyRankChange(inBuf.GetByte)

            Case SID_CLANUSERINFO
                'frmChat.AddChat vbRed, "!"

            Case SID_CLANMOTD

                cookie = inBuf.GetDWORD
                inBuf.GetDWORD()
                Message = inBuf.GetString

                RaiseEvent ClanMOTD(cookie, Message)

            Case Else

                RaiseEvent UnknownClanEvent(PacketID, DebugOutput(Data))

        End Select
        inBuf.Clear()
        'UPGRADE_NOTE: Object inBuf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        inBuf = Nothing
        Exit Sub

ERROR_HANDLER:
        frmChat.AddChat(RTBColors.ErrorMessageText, "Error: " & Err.Description & " in ParseClanPacket().")

        Exit Sub

        'Debug.Print "ParseClanPacket Error: " & Err.Number & ": " & Err.description
        'Debug.Print DebugOutput(Data)
        'Err.Clear
    End Sub
End Class