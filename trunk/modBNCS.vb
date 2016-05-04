Option Strict Off
Option Explicit On
Module modBNCS
	Private Const OBJECT_NAME As String = "modBNCS"
	
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SystemTimeToFileTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Integer
	
	' Packet IDs
	Public Const SID_NULL As Byte = &H0
	Public Const SID_CLIENTID As Byte = &H5
	Public Const SID_STARTVERSIONING As Byte = &H6
	Public Const SID_REPORTVERSION As Byte = &H7
	Public Const SID_ENTERCHAT As Byte = &HA
	Public Const SID_GETCHANNELLIST As Byte = &HB
	Public Const SID_JOINCHANNEL As Byte = &HC
	Public Const SID_CHATCOMMAND As Byte = &HE
	Public Const SID_CHATEVENT As Byte = &HF
	Public Const SID_LEAVECHAT As Byte = &H10
	Public Const SID_LOCALEINFO As Byte = &H12
	Public Const SID_UDPPINGRESPONSE As Byte = &H14
	Public Const SID_MESSAGEBOX As Byte = &H19
	Public Const SID_LOGONCHALLENGEEX As Byte = &H1D
	Public Const SID_CLIENTID2 As Byte = &H1E
	Public Const SID_NOTIFYJOIN As Byte = &H22
	Public Const SID_PING As Byte = &H25
	Public Const SID_READUSERDATA As Byte = &H26
	Public Const SID_WRITEUSERDATA As Byte = &H27
	Public Const SID_LOGONCHALLENGE As Byte = &H28
	Public Const SID_GETICONDATA As Byte = &H2D
	Public Const SID_CDKEY As Byte = &H30
	Public Const SID_CHANGEPASSWORD As Byte = &H31
	Public Const SID_PROFILE As Byte = &H35
	Public Const SID_CDKEY2 As Byte = &H36
	Public Const SID_LOGONRESPONSE2 As Byte = &H3A
	Public Const SID_CREATEACCOUNT2 As Byte = &H3D
	Public Const SID_LOGONREALMEX As Byte = &H3E
	Public Const SID_QUERYREALMS2 As Byte = &H40
	Public Const SID_WARCRAFTGENERAL As Byte = &H44
	Public Const SID_EXTRAWORK As Byte = &H4C
	Public Const SID_AUTH_INFO As Byte = &H50
	Public Const SID_AUTH_CHECK As Byte = &H51
	Public Const SID_AUTH_ACCOUNTCREATE As Byte = &H52
	Public Const SID_AUTH_ACCOUNTLOGON As Byte = &H53
	Public Const SID_AUTH_ACCOUNTLOGONPROOF As Byte = &H54
	Public Const SID_AUTH_ACCOUNTCHANGE As Byte = &H55
	Public Const SID_AUTH_ACCOUNTCHANGEPROOF As Byte = &H56
	Public Const SID_SETEMAIL As Byte = &H59
	Public Const SID_RESETPASSWORD As Byte = &H5A
	Public Const SID_CHANGEEMAIL As Byte = &H5B
	Public Const SID_FRIENDSLIST As Byte = &H65
	Public Const SID_FRIENDSUPDATE As Byte = &H66
	Public Const SID_FRIENDSADD As Byte = &H67
	Public Const SID_FRIENDSREMOVE As Byte = &H68
	Public Const SID_FRIENDSPOSITION As Byte = &H69
	Public Const SID_CLANFINDCANDIDATES As Byte = &H70
	Public Const SID_CLANINVITEMULTIPLE As Byte = &H71
	Public Const SID_CLANCREATIONINVITATION As Byte = &H72
	Public Const SID_CLANDISBAND As Byte = &H73
	Public Const SID_CLANMAKECHIEFTAIN As Byte = &H74
	Public Const SID_CLANINFO As Byte = &H75
	Public Const SID_CLANQUITNOTIFY As Byte = &H76
	Public Const SID_CLANINVITATION As Byte = &H77
	Public Const SID_CLANREMOVEMEMBER As Byte = &H78
	Public Const SID_CLANINVITATIONRESPONSE As Byte = &H79
	Public Const SID_CLANRANKCHANGE As Byte = &H7A
	Public Const SID_CLANSETMOTD As Byte = &H7B
	Public Const SID_CLANMOTD As Byte = &H7C
	Public Const SID_CLANMEMBERLIST As Byte = &H7D
	Public Const SID_CLANMEMBERREMOVED As Byte = &H7E
	Public Const SID_CLANMEMBERSTATUSCHANGE As Byte = &H7F
	Public Const SID_CLANMEMBERRANKCHANGE As Byte = &H81
	Public Const SID_CLANMEMBERINFORMATION As Byte = &H82
	
	
	' SID_CHATEVENT EVENT IDs
	Public Const ID_USER As Integer = &H1
	Public Const ID_JOIN As Integer = &H2
	Public Const ID_LEAVE As Integer = &H3
	Public Const ID_WHISPER As Integer = &H4
	Public Const ID_TALK As Integer = &H5
	Public Const ID_BROADCAST As Integer = &H6
	Public Const ID_CHANNEL As Integer = &H7
	Public Const ID_USERFLAGS As Integer = &H9
	Public Const ID_WHISPERSENT As Integer = &HA
	Public Const ID_CHANNELFULL As Integer = &HD
	Public Const ID_CHANNELDOESNOTEXIST As Integer = &HE
	Public Const ID_CHANNELRESTRICTED As Integer = &HF
	Public Const ID_INFO As Integer = &H12
	Public Const ID_ERROR As Integer = &H13
	Public Const ID_EMOTE As Integer = &H17
	' Additional event constants for logging
	Public Const ID_CONNECTED As Integer = &H18
	Public Const ID_DISCONNECTED As Integer = &H19
	
	' battle.net user flag constants
	Public Const USER_BLIZZREP As Integer = &H1
	Public Const USER_CHANNELOP As Integer = &H2
	Public Const USER_SPEAKER As Integer = &H4
	Public Const USER_SYSOP As Integer = &H8
	Public Const USER_NOUDP As Integer = &H10
	Public Const USER_BEEPENABLED As Integer = &H100
	Public Const USER_KBKOFFICIAL As Integer = &H1000
	Public Const USER_JAILED As Integer = &H100000
	Public Const USER_SQUELCHED As Integer = &H20
	Public Const USER_PGLPLAYER As Integer = &H200
	Public Const USER_GFOFFICIAL As Integer = &H100000
	Public Const USER_GFPLAYER As Integer = &H200000
	Public Const USER_GUEST As Integer = &H40
	Public Const USER_PGLOFFICIAL As Integer = &H400
	Public Const USER_KBKPLAYER As Integer = &H800
	
	
	Public Const BNCS_NLS As Integer = 1 'New:    SID_AUTH_*
	Public Const BNCS_OLS As Integer = 2 'Old:    SID_CLIENTID2
	Public Const BNCS_LLS As Integer = 3 'Legacy: SID_CLIENTID
	
	Public Const PLATFORM_INTEL As Integer = &H49583836 'IX86
	Public Const PLATFORM_POWERPC As Integer = &H504D4143 'PMAC
	Public Const PLATFORM_OSX As Integer = &H584D4143 'XMAC
	
	
	Public ds As New clsDataStorage 'Need to rename this -.-
	
    Public Function BNCSRecvPacket(ByVal Data() As Byte) As Boolean
        On Error GoTo ERROR_HANDLER
        Static pBuff As New clsDataBuffer

        Dim PacketID As Byte

        BNCSRecvPacket = True
        With pBuff
            .Clear()
            .Data = Data
            .GetByte()
            PacketID = .GetByte
            .GetWord()
        End With

        Select Case PacketID
            Case SID_NULL 'Don't Throw Unknown Error                  '0x00
            Case SID_CLIENTID 'Don't Throw Unknown Error                  '0x05
            Case SID_STARTVERSIONING : Call RECV_SID_STARTVERSIONING(pBuff) '0x06
            Case SID_REPORTVERSION : Call RECV_SID_REPORTVERSION(pBuff) '0x07
            Case SID_ENTERCHAT : Call RECV_SID_ENTERCHAT(pBuff) '0x0A
            Case SID_GETCHANNELLIST : Call RECV_SID_GETCHANNELLIST(pBuff) '0x0B
            Case SID_CHATEVENT : Call RECV_SID_CHATEVENT(pBuff) '0x0F
            Case SID_MESSAGEBOX : Call RECV_SID_MESSAGEBOX(pBuff) '0x19
            Case SID_LOGONCHALLENGEEX : Call RECV_SID_LOGONCHALLENGEEX(pBuff) '0x1D
            Case SID_PING : Call RECV_SID_PING(pBuff) '0x25
            Case SID_LOGONCHALLENGE : Call RECV_SID_LOGONCHALLENGE(pBuff) '0x28
            Case SID_GETICONDATA 'Don't Throw Unknown Error                  '0x2D
            Case SID_CDKEY : Call RECV_SID_CDKEY(pBuff) '0x30
            Case SID_CDKEY2 : Call RECV_SID_CDKEY2(pBuff) '0x36
            Case SID_LOGONRESPONSE2 : Call RECV_SID_LOGONRESPONSE2(pBuff) '0x3A
            Case SID_CREATEACCOUNT2 : Call RECV_SID_CREATEACCOUNT2(pBuff) '0x3D
            Case SID_LOGONREALMEX : Call RECV_SID_LOGONREALMEX(pBuff) '0x3C
            Case SID_QUERYREALMS2 : Call RECV_SID_QUERYREALMS2(pBuff) '0x40
            Case SID_EXTRAWORK 'Don't Throw Unknown Error                  '0x4C
            Case SID_AUTH_INFO : Call RECV_SID_AUTH_INFO(pBuff) '0x50
            Case SID_AUTH_CHECK : Call RECV_SID_AUTH_CHECK(pBuff) '0x51
            Case SID_AUTH_ACCOUNTCREATE : Call RECV_SID_AUTH_ACCOUNTCREATE(pBuff) '0x52
            Case SID_AUTH_ACCOUNTLOGON : Call RECV_SID_AUTH_ACCOUNTLOGON(pBuff) '0x53
            Case SID_AUTH_ACCOUNTLOGONPROOF : Call RECV_SID_AUTH_ACCOUNTLOGONPROOF(pBuff) '0x54
            Case SID_SETEMAIL : Call RECV_SID_SETEMAIL(pBuff) '0x59

            Case Else
                BNCSRecvPacket = False
                If (MDebug("debug") And (MDebug("all") Or MDebug("unknown"))) Then
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unhandled packet 0x{0}", ZeroOffset(CInt(PacketID), 2)))
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Packet data: {0}{1}", vbNewLine, DebugOutput(Data)))
                End If

        End Select

        Exit Function
ERROR_HANDLER:
        Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.BNCSRecvPacket()", Err.Number, Err.Description, OBJECT_NAME))
    End Function
	
	'*********************************
	' SID_CLIENTID (0x05) C->S
	'*********************************
	' (DWORD) Registration Version
	' (DWORD) Registration Authority
	' (DWORD) Account Number
	' (DWORD) Registration Token
	' (STRING) LAN computer name
	' (STRING) LAN username
	'*********************************
	'For legacy login system (JSTR, SSHR).
	'*********************************
	Public Sub SEND_SID_CLIENTID()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertNTString(GetComputerLanName)
			.InsertNTString(GetComputerUsername)
			.SendPacket(SID_CLIENTID)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CLIENTID()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_STARTVERSIONING (0x06) S->C
	'*******************************
	' (FILETIME) MPQ Filetime
	' (STRING) MPQ Filename
	' (STRING) ValueString
	'*******************************
	Public Sub RECV_SID_STARTVERSIONING(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		With pBuff
			ds.CRevFileTime = .GetRaw(8)
			ds.CRevFileName = .GetString
			ds.CRevSeed = .GetString
		End With
		
		Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Checking version...")
		If (MDebug("all") Or MDebug("crev")) Then
			frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Name: {0}", ds.CRevFileName))
			frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Time: {0}", ds.CRevFileTime))
			If (InStr(1, ds.CRevFileName, "lockdown", CompareMethod.Text) > 0) Then
				frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Seed: {0}", StrToHex(ds.CRevSeed)))
			Else
				frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Seed: {0}", ds.CRevSeed))
			End If
		End If
		
		' If the server does not recognize the version byte we sent it, it will send back an empty seed string.
        If (Len(ds.CRevSeed) = 0) Then
            Call HandleEmptyCRevSeed()

            Exit Sub
        End If
		
		If (BotVars.BNLS) Then
			Call modBNLS.SEND_BNLS_VERSIONCHECKEX2((ds.CRevFileTimeRaw), (ds.CRevFileName), (ds.CRevSeed))
		Else
			Call SEND_SID_REPORTVERSION()
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_STARTVERSIONING (0x06) C->S
	'*******************************
	' (DWORD) Platform ID
	' (DWORD) Product ID
	' (DWORD) Version Byte
	' (DWORD) Unknown (0)
	'*******************************
	Public Sub SEND_SID_STARTVERSIONING(Optional ByRef lVerByte As Integer = 0)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		
		With pBuff
			.InsertDWord(GetDWORDOverride(Config.PlatformID)) 'Platform ID
			.InsertDWord(GetDWORD((BotVars.Product))) 'Product ID
			.InsertDWord(IIf(lVerByte = 0, GetVerByte((BotVars.Product)), lVerByte)) 'VersionByte
			.InsertDWord(0) 'Unknown
			.SendPacket(SID_STARTVERSIONING)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	'*******************************
	'SID_REPORTVERSION (0x07) S->C
	'*******************************
	' (DWORD) Result
	' (STRING) Patch path
	'*******************************
	Private Sub RECV_SID_REPORTVERSION(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lResult As Integer
		Dim sInfo As String
		Dim bSuccess As Boolean
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		bSuccess = False
		
		Select Case lResult
			Case 0 : Call Event_VersionCheck(1, sInfo) 'Failed Version Check
			Case 1 : Call Event_VersionCheck(1, sInfo) 'Old Game Version
			Case 2 'Success
				bSuccess = True
				'Call Event_VersionCheck(0, sInfo)
			Case 3 : Call Event_VersionCheck(1, sInfo) '"Reinstall Required", Invalid version
			Case Else
				Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unknown SID_REPORTVERSION Response: 0x" & ZeroOffset(lResult, 8))
		End Select
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNet.CtlState = 7 And bSuccess) Then
			If (GetCDKeyCount > 0) Then
				'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending CDKey information...")
				Select Case GetLogonSystem()
					Case BNCS_OLS : Call SEND_SID_CDKEY2()
					Case BNCS_LLS : Call SEND_SID_CDKEY()
					Case Else
						frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem()))
						frmChat.AddChat(RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error.")
						frmChat.DoDisconnect()
				End Select
			Else
				Call Event_VersionCheck(0, sInfo) ' display success here
				Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending login information...")
				frmChat.tmrAccountLock.Enabled = True
				SEND_SID_LOGONRESPONSE2()
			End If
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_REPORTVERSION()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_REPORTVERSION (0x07) C->S
	'*******************************
	' (DWORD) Platform ID
	' (DWORD) Product ID
	' (DWORD) Version Byte
	' (DWORD) EXE Version
	' (DWORD) EXE Hash
	' (STRING) EXE Information
	'*******************************
	Public Sub SEND_SID_REPORTVERSION(Optional ByRef lVerByte As Integer = 0)
		On Error GoTo ERROR_HANDLER
		
		If (Not BotVars.BNLS) Then
			If (Not CompileCheckrevision()) Then
				frmChat.DoDisconnect()
				Exit Sub
			End If
		End If
		
        If (ds.CRevChecksum = 0 Or ds.CRevVersion = 0 Or Len(ds.CRevResult) = 0) Then
            frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Check Revision Failed, sanity failed")
            frmChat.DoDisconnect()
            Exit Sub
        End If
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertDWord(GetDWORDOverride(Config.PlatformID)) 'Platform ID
			.InsertDWord(GetDWORD((BotVars.Product))) 'Product ID
			.InsertDWord(IIf(lVerByte = 0, GetVerByte((BotVars.Product)), lVerByte)) 'VersionByte
			.InsertDWord(ds.CRevVersion) 'Exe Version
			.InsertDWord(ds.CRevChecksum) 'Checksum
			.InsertNTString((ds.CRevResult)) 'Result
			.SendPacket(SID_REPORTVERSION)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_ENTERCHAT (0x0A) S->C
	'*******************************
	' (String) Unique name
	' (String) Statstring
	' (String) Account name
	'*******************************
	Private Sub RECV_SID_ENTERCHAT(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Call Event_LoggedOnAs(pBuff.GetString, pBuff.GetString, pBuff.GetString)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_ENTERCHAT (0x0A) C->S
	'*******************************
	' (STRING) Username *
	' (STRING) Statstring **
	'*******************************
	Private Sub SEND_SID_ENTERCHAT()
		On Error GoTo ERROR_HANDLER
		Dim pBuff As New clsDataBuffer
		pBuff.InsertNTString((BotVars.Username))
		pBuff.InsertNTString((Config.CustomStatstring))
		pBuff.SendPacket(SID_ENTERCHAT)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_LEAVECHAT (0x10) C->S
	'*******************************
	' [blank]
	'*******************************
	Public Sub SEND_SID_LEAVECHAT()
		On Error GoTo ERROR_HANDLER
		Dim pBuff As New clsDataBuffer
		pBuff.SendPacket(SID_LEAVECHAT)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_GETCHANNELLIST (0x0B) S->C
	'*******************************
	' (String[]) Channels
	'*******************************
	Private Sub RECV_SID_GETCHANNELLIST(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim sChannels() As String
		Dim sChannel As String
		
		sChannels = Split(vbNullString)
		
		Do 
			sChannel = pBuff.GetString()
			
            If Len(sChannel) > 0 Then
                ReDim Preserve sChannels(UBound(sChannels) + 1)
                sChannels(UBound(sChannels)) = Trim(sChannel)
            End If
        Loop While Len(sChannel) > 0
		
		Call Event_ChannelList(sChannels)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_GETCHANNELLIST()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_GETCHANNELLIST (0x0B) S->C
	'*******************************
	' (DWORD) Product ID
	'*******************************
	Private Sub SEND_SID_GETCHANNELLIST()
		On Error GoTo ERROR_HANDLER
		Dim pBuff As New clsDataBuffer
		pBuff.InsertDWord(GetDWORD((BotVars.Product)))
		pBuff.SendPacket(SID_GETCHANNELLIST)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_GETCHANNELLIST()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CHATCOMMAND (0x0E) S->C
	'*******************************
	' (STRING) Text
	'*******************************
	Public Sub SEND_SID_CHATCOMMAND(ByRef sText As String)
		On Error GoTo ERROR_HANDLER
		
        If (Len(sText) = 0) Then Exit Sub
		
		Dim pBuff As New clsDataBuffer
		pBuff.InsertNTString(sText)
		pBuff.SendPacket(SID_CHATCOMMAND)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CHATCOMMAND()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	
	'*******************************
	'SID_CHATEVENT (0x0F) S->C
	'*******************************
	' (DWORD) Event ID
	' (DWORD) User's Flags
	' (DWORD) Ping
	' (DWORD) IP Address (Defunct)
	' (DWORD) Account number (Defunct)
	' (DWORD) Registration Authority (Defunct)
	' (STRING) Username
	' (STRING) Text
	'*******************************
	Private Sub RECV_SID_CHATEVENT(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim EventID As Integer
		Dim lFlags As Integer
		Dim lPing As Integer
		Dim sUsername As String
		Dim sText As String
		
		Dim sProduct As String
		Dim sParsed As String
		Dim sClanTag As String
		Dim sW3Icon As String
		
		EventID = pBuff.GetDWORD
		lFlags = pBuff.GetDWORD
		lPing = pBuff.GetDWORD
		pBuff.GetDWORD() 'IP Address
		pBuff.GetDWORD() 'Account Number
		pBuff.GetDWORD() 'Reg Auth
		sUsername = pBuff.GetString
		sText = pBuff.GetString
		
		
		
		Select Case EventID
			Case ID_JOIN : Call Event_UserJoins(sUsername, lFlags, sText, lPing)
			Case ID_LEAVE : Call Event_UserLeaves(sUsername, lFlags)
			Case ID_USER : Call Event_UserInChannel(sUsername, lFlags, sText, lPing)
			Case ID_WHISPER : Call Event_WhisperFromUser(sUsername, lFlags, sText, lPing)
			Case ID_TALK : Call Event_UserTalk(sUsername, lFlags, sText, lPing)
			Case ID_BROADCAST
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Event_ServerInfo(sUsername, StringFormat("BROADCAST from {0}: {1}", sUsername, sText))
			Case ID_CHANNEL : Call Event_JoinedChannel(sText, lFlags)
			Case ID_USERFLAGS : Call Event_FlagsUpdate(sUsername, lFlags, sText, lPing)
			Case ID_WHISPERSENT : Call Event_WhisperToUser(sUsername, lFlags, sText, lPing)
			Case ID_CHANNELFULL, ID_CHANNELDOESNOTEXIST, ID_CHANNELRESTRICTED : Call Event_ChannelJoinError(EventID, sText)
			Case ID_INFO : Call Event_ServerInfo(sUsername, sText)
			Case ID_ERROR : Call Event_ServerError(sText)
			Case ID_EMOTE : Call Event_UserEmote(sUsername, lFlags, sText)
			Case Else
				If MDebug("debug") Then
					Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Unhandled SID_CHATEVENT Event: 0x{0}", ZeroOffset(EventID, 8)))
					Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Packet data: {0}{1}", vbNewLine, DebugOutput(pBuff.Data)))
				End If
		End Select
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CHATEVENT()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*********************************
	' SID_LOCALEINFO (0x12) C->S
	'*********************************
	' (FILETIME) System time
	' (FILETIME) Local time
	' (DWORD) Timezone bias
	' (DWORD) SystemDefaultLCID
	' (DWORD) UserDefaultLCID
	' (DWORD) UserDefaultLangID
	' (STRING) Abbreviated language name
	' (STRING) Country Code
	' (STRING) Abbreviated country name
	' (STRING) Country Name
	'*********************************
	Public Sub SEND_SID_LOCALEINFO()
		On Error GoTo ERROR_HANDLER
		Const LOCALE_SABBREVLANGNAME As Integer = &H3
		Const LOCALE_USER_DEFAULT As Integer = &H400
		Dim LanguageAbr As String
		Dim CountryCode As String
		Dim CountryAbr As String
		Dim CountryName As String
		Dim lRet As String
		
		Dim st As SYSTEMTIME
		Dim ft As FILETIME
		
		Dim pBuff As New clsDataBuffer
		
		LanguageAbr = New String(Chr(0), 256)
		Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVLANGNAME, LanguageAbr, Len(LanguageAbr))
		LanguageAbr = KillNull(LanguageAbr)
		
		Call GetCountryData(CountryAbr, CountryName, CountryCode)
		If (Len(LanguageAbr) = 0) Then LanguageAbr = "ENU"
		If (Len(CountryCode) = 0) Then CountryCode = "1"
		If (Not Len(CountryAbr) = 3) Then CountryAbr = "USA"
        If (Len(CountryName) = 0) Then CountryName = "United States"
		
		With pBuff
			Call GetSystemTime(st)
			Call SystemTimeToFileTime(st, ft)
			.InsertDWord(ft.dwLowDateTime) 'SystemTime
			.InsertDWord(ft.dwHighDateTime) 'SystemTime
			
			Call GetLocalTime(st)
			Call SystemTimeToFileTime(st, ft)
			.InsertDWord(ft.dwLowDateTime) 'LocalTime
			.InsertDWord(ft.dwHighDateTime) 'LocalTime
			
			.InsertDWord(GetTimeZoneBias) 'Time Zone Bias
			If Config.ForceDefaultLocaleID Then
				.InsertDWord(1033) 'SystemDefaultLCID
				.InsertDWord(1033) 'UserDefaultLCID
				.InsertDWord(1033) 'UserDefaultLangID
			Else
				.InsertDWord(CInt(GetSystemDefaultLCID)) 'SystemDefaultLCID
				.InsertDWord(CInt(GetUserDefaultLCID)) 'UserDefaultLCID
				.InsertDWord(CInt(GetUserDefaultLangID)) 'UserDefaultLangID
			End If
			
			.InsertNTString(LanguageAbr) 'Language Abbrev
			.InsertNTString(CountryCode) 'Country Code
			.InsertNTString(CountryAbr) 'Country Abbrev
			.InsertNTString(CountryName) 'Country Name
			
			.SendPacket(SID_LOCALEINFO)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOCALEINFO()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_UDPPINGRESPONSE (0x14) C->S
	'*******************************
	' (DWORD) UDP value
	'*******************************
	Private Sub SEND_SID_UDPPINGRESPONSE()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		
		pBuff.InsertDWord(GetDWORDOverride(Config.UDPString, &H626E6574)) 'default: bnet
		pBuff.SendPacket(SID_UDPPINGRESPONSE)
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_UDPPINGRESPONSE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_MESSAGEBOX (0x19) S->C
	'*******************************
	' (DWORD) Style
	' (String) Text
	' (String) Caption
	'*******************************
	Private Sub RECV_SID_MESSAGEBOX(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Call Event_MessageBox(pBuff.GetDWORD, pBuff.GetString, pBuff.GetString)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_MESSAGEBOX()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'********************************
	'SID_LOGONCHALLENGEEX (0x1D) S->C
	'********************************
	' (DWORD) UDP Token
	' (DWORD) Server Token
	'********************************
	Private Sub RECV_SID_LOGONCHALLENGEEX(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		ds.UDPValue = pBuff.GetDWORD
		ds.ServerToken = pBuff.GetDWORD
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONCHALLENGEEX()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*********************************
	' SID_CLIENTID2 (0x1E) C->S
	'*********************************
	' (DWORD) Server Version
	' For server version 1:
	'  (DWORD) Registration Version
	'  (DWORD) Registration Authority
	' For server version 0:
	'  (DWORD) Registration Authority
	'  (DWORD) Registration Version
	' (DWORD) Account Number
	' (DWORD) Registration Token
	' (STRING) LAN computer name
	' (STRING) LAN username
	'*********************************
	'This is eww, I don't like hard coding,
	'but to get this crap I would need to
	'use Storm.dll, which we don't want to
	'distribute with the bot.
	'*********************************
	Public Sub SEND_SID_CLIENTID2()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertDWord(1)
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertDWord(0)
			.InsertNTString(GetComputerLanName)
			.InsertNTString(GetComputerUsername)
			.SendPacket(SID_CLIENTID2)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CLIENTID2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_PING (0x25) S->C
	'*******************************
	' (DWORD) Ping value
	'*******************************
	Private Sub RECV_SID_PING(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		If (BotVars.Spoof = 0 Or g_Online) Then
			Call SEND_SID_PING(pBuff.GetDWORD)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_PING()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_PING (0x25) C->S
	'*******************************
	' (DWORD) Ping value
	'*******************************
	Private Sub SEND_SID_PING(ByRef lPingValue As Integer)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		
		SetNagelStatus(frmChat.sckBNet.SocketHandle, False)
		
		pBuff.InsertDWord(lPingValue)
		pBuff.SendPacket(SID_PING)
		
		SetNagelStatus(frmChat.sckBNet.SocketHandle, True)
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_PING()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'********************************
	'SID_LOGONCHALLENGE (0x28) S->C
	'********************************
	' (DWORD) Server Token
	'********************************
	Private Sub RECV_SID_LOGONCHALLENGE(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		ds.ServerToken = pBuff.GetDWORD
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONCHALLENGE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CDKEY (0x30) S->C
	'*******************************
	' (DWORD) Result
	' (STRING) Key owner
	'*******************************
	Private Sub RECV_SID_CDKEY(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lResult As Integer
		Dim sInfo As String
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		
		Select Case lResult
			Case 1
				Call Event_VersionCheck(0, sInfo) ' display success here
				'frmChat.AddChat RTBColors.SuccessText, "[BNCS] Your CDKey was accepted!"
				frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending login information...")
				frmChat.tmrAccountLock.Enabled = True
				SEND_SID_LOGONRESPONSE2()
				Exit Sub
			Case 2 : Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
			Case 3 : Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
			Case 4 : Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
			Case 5 : Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
			Case Else : frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown SID_CDKEY Response 0x{0}: {1}", ZeroOffset(lResult, 8), sInfo))
		End Select
		'CloseAllConnections
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CDKEY()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CDKEY (0x30) C->S
	'*******************************
	' (DWORD) Spawn (0/1)
	' (STRING) Key
	' (STRING) Key owner
	'*******************************
	Public Sub SEND_SID_CDKEY()
		On Error GoTo ERROR_HANDLER
		Dim oKey As New clsKeyDecoder
		Dim pBuff As New clsDataBuffer
		
		oKey.Initialize(BotVars.CDKey)
		If Not oKey.IsValid Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "Your CD-Key is invalid.")
			frmChat.DoDisconnect()
			Exit Sub
		End If
		
		With pBuff
			.InsertBool(CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
			.InsertNTString((BotVars.CDKey))
			
            If (Len(Config.CDKeyOwnerName) > 0) Then
                .InsertNTString((Config.CDKeyOwnerName))
            Else
                .InsertNTString((BotVars.Username))
            End If
			.SendPacket(SID_CDKEY)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		'UPGRADE_NOTE: Object oKey may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oKey = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CDKEY()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CDKEY2 (0x36) S->C
	'*******************************
	' (DWORD) Result
	' (STRING) Key owner
	'*******************************
	Private Sub RECV_SID_CDKEY2(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lResult As Integer
		Dim sInfo As String
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		
		Select Case lResult
			Case 1
				Call Event_VersionCheck(0, sInfo) ' display success here
				'frmChat.AddChat RTBColors.SuccessText, "[BNCS] Your CDKey was accepted!"
				frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending login information...")
				frmChat.tmrAccountLock.Enabled = True
				SEND_SID_LOGONRESPONSE2()
				Exit Sub
			Case 2 : Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
			Case 3 : Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
			Case 4 : Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
			Case 5 : Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
			Case Else : frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown SID_CDKEY2 Response 0x{0}: {1}", ZeroOffset(lResult, 8), sInfo))
		End Select
		'CloseAllConnections
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CDKEY2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CDKEY2 (0x36) C->S
	'*******************************
	' (DWORD) Spawn (0/1)
	' (DWORD) Key Length
	' (DWORD) CDKey Product
	' (DWORD) CDKey Value1
	' (DWORD) Server Token
	' (DWORD) Client Token
	' (DWORD) [5] Hashed Data
	' (STRING) Key owner
	'*******************************
	Public Sub SEND_SID_CDKEY2()
		On Error GoTo ERROR_HANDLER
		Dim oKey As New clsKeyDecoder
		Dim pBuff As New clsDataBuffer
		
		oKey.Initialize(BotVars.CDKey)
		If Not oKey.IsValid Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "Your CD-Key is invalid.")
			frmChat.DoDisconnect()
			Exit Sub
		End If
		
		If Not oKey.CalculateHash(ds.ClientToken, ds.ServerToken, BNCS_OLS) Then Exit Sub
		
		With pBuff
			.InsertBool(CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
			.InsertDWord(oKey.KeyLength)
			.InsertDWord(oKey.ProductValue)
			.InsertDWord(oKey.PublicValue)
			.InsertDWord(ds.ServerToken)
			.InsertDWord(ds.ClientToken)
            .InsertByteArr(oKey.Hash)
			
            If (Len(Config.CDKeyOwnerName) > 0) Then
                .InsertNTString((Config.CDKeyOwnerName))
            Else
                .InsertNTString((BotVars.Username))
            End If
			.SendPacket(SID_CDKEY2)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		'UPGRADE_NOTE: Object oKey may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oKey = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CDKEY2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	
	'*******************************
	'SID_LOGONRESPONSE2 (0x3A) S->C
	'*******************************
	' (DWORD) Result
	' (STRING) Reason
	'*******************************
	Private Sub RECV_SID_LOGONRESPONSE2(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lResult As Integer
		Dim sInfo As String
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		
		Select Case lResult
			Case &H0 'Logon Successful
				Call Event_LogonEvent(2, sInfo)
				
				If (Not ds.WaitingForEmail) Then
					If (Dii And BotVars.UseRealm) Then
						'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Asking Battle.net for a list of Realm servers...")
						Call DoQueryRealms()
					Else
						SendEnterChatSequence()
					End If
				Else
					DoRegisterEmail()
				End If
				
			Case &H1 'Nonexistent account.
				Call Event_LogonEvent(0, sInfo)
				Call Event_LogonEvent(3, sInfo)
				SEND_SID_CREATEACCOUNT2()
				
			Case &H2 'Invalid password.
				Call Event_LogonEvent(1, sInfo)
				Call frmChat.DoDisconnect()
				
			Case &H6 'Account has been closed (includes a reason)
				Call Event_LogonEvent(5, sInfo)
				Call frmChat.DoDisconnect()
				
			Case Else
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown response to SID_LOGONRESPONSE2: 0x{0}", ZeroOffset(lResult, 8)))
		End Select
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONRESPONSE2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_LOGONRESPONSE2 (0x3A) C->S
	'*******************************
	' (DWORD) Client Token
	' (DWORD) Server Token
	' (DWORD) [5] Password Hash
	' (STRING) Username
	'*******************************
	Public Sub SEND_SID_LOGONRESPONSE2()
		On Error GoTo ERROR_HANDLER
		Dim sHash As String
		Dim pBuff As New clsDataBuffer
		
		If Not Config.UseLowerCasePassword Then
			sHash = doubleHashPassword((BotVars.Password), ds.ClientToken, ds.ServerToken)
		Else
			sHash = doubleHashPassword(LCase(BotVars.Password), ds.ClientToken, ds.ServerToken)
		End If
		
		With pBuff
			.InsertDWord(ds.ClientToken)
			.InsertDWord(ds.ServerToken)
			.InsertNonNTString(sHash)
			.InsertNTString((BotVars.Username))
			.SendPacket(SID_LOGONRESPONSE2)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOGONRESPONSE2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_CREATEACCOUNT2 (0x3D) S->C
	'*******************************
	' (DWORD) Status
	' (STRING) Account name suggestion
	'*******************************
	Private Sub RECV_SID_CREATEACCOUNT2(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lResult As Integer
		Dim sInfo As String
		Dim sOut As String
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		
		Select Case lResult
			Case 0
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Account created successfully!")
				modBNCS.SEND_SID_LOGONRESPONSE2()
				Exit Sub
				
			Case 2 : sOut = "Your desired account name contains invalid characters."
			Case 3 : sOut = "Your desired account name contains a banned word."
			Case 4 : sOut = "Your desired account name already exists."
			Case 6 : sOut = "Your desired account name does not contain enough alphanumeric characters."
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOut = StringFormat("Unknown response to SID_CREATEACCOUNT2. Result code: 0x{0}", ZeroOffset(lResult, 8))
		End Select
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] There was an error in trying to create a new account.")
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] {0}", sOut))
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CREATEACCOUNT2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**************************************
	'SID_CREATEACCOUNT2 (0x3D) C->S
	'**************************************
	' (DWORD) [5] Password hash
	' (STRING) Username
	'**************************************
	Private Sub SEND_SID_CREATEACCOUNT2()
		On Error GoTo ERROR_HANDLER
		
		Dim sHash As String
		If Not Config.UseLowerCasePassword Then
			sHash = hashPassword((BotVars.Password))
		Else
			sHash = hashPassword(LCase(BotVars.Password))
		End If
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertNonNTString(sHash)
			.InsertNTString((BotVars.Username))
			.SendPacket(SID_CREATEACCOUNT2)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CREATEACCOUNT2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_LOGONREALMEX (0x3E) S->C
	'*******************************
	' (DWORD) MCP Cookie
	' (DWORD) MCP Status
	' (DWORD) [2] MCP Chunk 1
	' (DWORD) IP
	' (DWORD) Port
	' (DWORD) [12] MCP Chunk 2
	' (STRING) Battle.net unique name
	'*******************************
	Private Sub RECV_SID_LOGONREALMEX(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lError As Integer
		Dim sMCPData As String
		Dim sTitle As String
		Dim sIP As String
		Dim lPort As Integer
		Dim sUniq As String
		Dim x As Short
		
		If (Len(pBuff.GetRaw( , True)) > 8) Then
			sMCPData = pBuff.GetRaw(16)
			
			For x = 1 To 4
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sIP = StringFormat("{0}{1}{2}", sIP, pBuff.GetByte, IIf(x = 4, vbNullString, "."))
			Next x
			lPort = ntohs(pBuff.GetDWORD)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMCPData = StringFormat("{0}{1}", sMCPData, pBuff.GetRaw(48))
			sUniq = pBuff.GetString
			
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If (Not frmChat.sckMCP.CtlState = 0) Then frmChat.sckMCP.Close()
			
			If Not ds.MCPHandler Is Nothing Then
				Call ds.MCPHandler.SetStartupData(sMCPData, sUniq, sIP, lPort)
				sTitle = ds.MCPHandler.RealmServerTitle(ds.MCPHandler.RealmServerConnected)
				
				frmChat.AddChat(RTBColors.InformationText, StringFormat("[REALM] Connecting to the Diablo II Realm {0} at {1}:{2}...", sTitle, sIP, lPort))
				frmChat.sckMCP.Connect(sIP, lPort)
			End If
		Else
			pBuff.GetDWORD()
			lError = pBuff.GetDWORD
			
			Select Case lError
				Case &H80000001 : frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] The Diablo II Realm is currently unavailable. Please try again later.")
				Case &H80000002 : frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] Diablo II Realm logon has failed. Please try again later.")
				Case Else : frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[REALM] Login to the Diablo II Realm has failed for an unknown reason (0x{0}). Please try again later.", ZeroOffset(lError, 8)))
			End Select
			
			If Not ds.MCPHandler Is Nothing Then
				If ds.MCPHandler.FormActive Then
					frmRealm.UnloadRealmError()
				End If
			End If
			
			SendEnterChatSequence()
			frmChat.mnuRealmSwitch.Enabled = True
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONREALMEX()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_LOGONREALMEX (0x3E) C->S
	'*******************************
	' (DWORD) Client Token
	' (DWORD) [5] Hashed realm password
	' (STRING) Realm title
	'*******************************
	Public Sub SEND_SID_LOGONREALMEX(ByRef sRealmTitle As String, ByRef sRealmServerPassword As String)
		On Error GoTo ERROR_HANDLER
		
        If (Len(sRealmTitle) = 0) Then Exit Sub
		
		Dim pBuff As New clsDataBuffer
		pBuff.InsertDWord(ds.ClientToken)
		pBuff.InsertNonNTString(doubleHashPassword(sRealmServerPassword, ds.ClientToken, ds.ServerToken))
		pBuff.InsertNTString(sRealmTitle)
		pBuff.SendPacket(SID_LOGONREALMEX)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Logging on to the Diablo II Realm...")
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOGONREALMEX()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_QUERYREALMS2 (0x40) S->C
	'*******************************
	' (DWORD) Unknown
	' (DWORD) Count
	' For Each Realm:
	'   (DWORD) Unknown
	'   (STRING) Realm title
	'   (STRING) Realm description
	'*******************************
	Private Sub RECV_SID_QUERYREALMS2(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lCount As Integer
		Dim i As Short
		Dim List() As Object
		Dim Server(1) As String
		
		pBuff.GetDWORD() 'Unknown
		lCount = pBuff.GetDWORD
		
		If (MDebug("debug") And (MDebug("all") Or MDebug("info"))) Then
			frmChat.AddChat(RTBColors.InformationText, "Received Realm List:")
		End If
		
		If lCount > 0 Then
			ReDim List(lCount - 1)
			
			For i = 0 To lCount - 1
				pBuff.GetDWORD() 'Unknown
				
				Server(0) = pBuff.GetString
				Server(1) = pBuff.GetString
				'UPGRADE_WARNING: Couldn't resolve default property of object List(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				List(i) = VB6.CopyArray(Server)
				
				If (MDebug("debug") And (MDebug("all") Or MDebug("info"))) Then
					frmChat.AddChat(RTBColors.InformationText, StringFormat("  {0}: {1}", Server(0), Server(1)))
				End If
			Next i
			
			If Not ds.MCPHandler Is Nothing Then
				Call ds.MCPHandler.SetRealmServerInfo(List)
			End If
		End If
		
		'Call frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Received Realm list.")
		
		Call ds.MCPHandler.HandleQueryRealmServersResponse()
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.Recv_SID_QUERYREALMS2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_QUERYREALMS2 (0x40) C->S
	'*******************************
	' [Blank]
	'*******************************
	Public Sub SEND_SID_QUERYREALMS2()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		pBuff.SendPacket(SID_QUERYREALMS2)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_QUERYREALMS2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_AUTH_INFO (0x50) S->C
	'*******************************
	' (DWORD) Logon Type
	' (DWORD) Server Token
	' (DWORD) UDPValue
	' (FILETIME) CRev Archive File Time
	' (STRING) CRev Archive File Name
	' (STRING) CRev Seed Values
	' WAR3/W3XP Only:
	'   (VOID) 128-byte Server signature
	'*******************************
	Private Sub RECV_SID_AUTH_INFO(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		ds.LogonType = pBuff.GetDWORD
		ds.ServerToken = pBuff.GetDWORD
		ds.UDPValue = pBuff.GetDWORD
		ds.CRevFileTime = pBuff.GetRaw(8)
		ds.CRevFileName = pBuff.GetString
		ds.CRevSeed = pBuff.GetString
		ds.ServerSig = pBuff.GetRaw(128)
		
		Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Checking version...")
		
		If (MDebug("all") Or MDebug("crev")) Then
			frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Name: {0}", ds.CRevFileName))
			frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Time: {0}", ds.CRevFileTime))
			If (InStr(1, ds.CRevFileName, "lockdown", CompareMethod.Text) > 0) Then
				frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Seed: {0}", StrToHex(ds.CRevSeed)))
			Else
				frmChat.AddChat(RTBColors.InformationText, StringFormat("CRev Seed: {0}", ds.CRevSeed))
			End If
		End If
		
		' If the server does not recognize the version byte we sent it, it will send back an empty seed string.
        If (Len(ds.CRevSeed) = 0) Then
            Call HandleEmptyCRevSeed()

            Exit Sub
        End If
		
		If (Len(ds.ServerSig) = 128) Then
			If (ds.NLS.VerifyServerSignature(frmChat.sckBNet.RemoteHostIP, ds.ServerSig)) Then
				frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Server signature validated!")
			Else
				If (Not BotVars.UseProxy) Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Warning, Server signature is invalid, this may not be a valid server.")
				End If
			End If
		ElseIf (GetProductKey = "W3") Then 
			frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Warning, Server signature is missing, this may not be a valid server.")
		End If
		
		If (BotVars.BNLS) Then
			modBNLS.SEND_BNLS_VERSIONCHECKEX2((ds.CRevFileTimeRaw), (ds.CRevFileName), (ds.CRevSeed))
		Else
			modBNCS.SEND_SID_AUTH_CHECK()
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_INFO()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_AUTH_INFO (0x50) C->S
	'*******************************
	' (DWORD) Protocol ID (0)
	' (DWORD) Platform ID
	' (DWORD) Product ID
	' (DWORD) Version Byte
	' (DWORD) Product language
	' (DWORD) Local IP for NAT compatibility*
	' (DWORD) Time zone bias*
	' (DWORD) Locale ID*
	' (DWORD) Language ID*
	' (STRING) Country abreviation
	' (STRING) Country
	'*******************************
	Public Sub SEND_SID_AUTH_INFO(Optional ByRef lVerByte As Integer = 0)
		On Error GoTo ERROR_HANDLER
		
		Dim LocalIP As Integer
		Dim CountryAbr As String
		Dim CountryName As String
		
		Dim pBuff As New clsDataBuffer
		
		LocalIP = aton((frmChat.sckBNet.LocalIP))
		
		Call GetCountryData(CountryAbr, CountryName, CStr(VariantType.Null))
		If (Not Len(CountryAbr) = 3) Then CountryAbr = "USA"
        If (Len(CountryName) = 0) Then CountryName = "United States"
		
		With pBuff
			
			.InsertDWord(Config.ProtocolID) 'ProtocolID
			.InsertDWord(GetDWORDOverride(Config.PlatformID, PLATFORM_INTEL)) 'Platform ID
			.InsertDWord(GetDWORD((BotVars.Product))) 'Product ID
			.InsertDWord(IIf(lVerByte = 0, GetVerByte((BotVars.Product)), lVerByte)) 'VersionByte
			.InsertDWord(GetDWORDOverride(Config.ProductLanguage)) 'Product Language
			.InsertDWord(LocalIP) 'Local IP
			.InsertDWord(GetTimeZoneBias) 'Time Zone Bias
			If Config.ForceDefaultLocaleID Then
				.InsertDWord(1033) 'LocalID
				.InsertDWord(1033) 'LangID
			Else
				.InsertDWord(CInt(GetUserDefaultLCID)) 'LocalID
				.InsertDWord(CInt(GetUserDefaultLangID)) 'LangID
			End If
			.InsertNTString(CountryAbr) 'Country abreviation
			.InsertNTString(CountryName) 'Country Name
			.SendPacket(SID_AUTH_INFO)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		If (BotVars.Spoof = 1) Then
			Call SEND_SID_PING(pBuff.GetDWORD)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_INFO()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_AUTH_CHECK (0x51) S->C
	'*******************************
	' (DWORD) Result
	' (STRING) Additional Information
	'*******************************
	Private Sub RECV_SID_AUTH_CHECK(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lResult As Integer
		Dim sInfo As String
		Dim bSuccess As Boolean
		
		lResult = pBuff.GetDWORD
		sInfo = pBuff.GetString
		bSuccess = False
		
		Select Case lResult
			Case &H0
				bSuccess = True
				Call Event_VersionCheck(0, sInfo)
				
			Case &H100, &H101 : Call Event_VersionCheck(1, sInfo) 'Outdated/Invalid Version
			Case &H200 : Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
			Case &H201 : Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
			Case &H202 : Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
			Case &H203 : Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
			Case &H210 : Call Event_VersionCheck(7, sInfo) 'Invalid Exp CDKey
			Case &H211 : Call Event_VersionCheck(8, sInfo) 'Exp CDKey is In Use
			Case &H212 : Call Event_VersionCheck(9, sInfo) 'Exp CDKey is Banned
			Case &H213 : Call Event_VersionCheck(10, sInfo) 'Exp CDKey is for the wrong product
			Case Else
				If Config.IgnoreVersionCheck Then bSuccess = True
				Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unknown 0x51 Response: 0x" & ZeroOffset(lResult, 8))
		End Select
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNet.CtlState = 7 And (Not ds.WaitingForEmail) And bSuccess) Then
			Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending login information...")
			frmChat.tmrAccountLock.Enabled = True
			
			If (ds.LogonType = 2) Then
				ds.NLS.Initialize(BotVars.Username, BotVars.Password)
				modBNCS.SEND_SID_AUTH_ACCOUNTLOGON()
			Else
				modBNCS.SEND_SID_LOGONRESPONSE2()
			End If
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_CHECK()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_AUTH_CHECK (0x51) C->S
	'*******************************
	' (DWORD) Client Token
	' (DWORD) EXE Version
	' (DWORD) EXE Hash
	' (DWORD) Number of CD-keys in this packet
	' (BOOLEAN) Spawn CD-key
	' For Each Key:
	'   (DWORD) Key Length
	'   (DWORD) CD-key's product value
	'   (DWORD) CD-key's public value
	'   (DWORD) Unknown (0)
	'   (DWORD) [5] Hashed Key Data
	' (STRING) Exe Information
	' (STRING) CD-Key owner name
	'*******************************
	Public Sub SEND_SID_AUTH_CHECK()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		Dim i As Integer
		Dim keys As Integer
		Dim sKey As String
		Dim oKey As New clsKeyDecoder
		
		If (Not BotVars.BNLS) Then
			If (Not CompileCheckrevision()) Then
				frmChat.DoDisconnect()
				Exit Sub
			End If
		End If
		
        If (ds.CRevChecksum = 0 Or ds.CRevVersion = 0 Or Len(ds.CRevResult) = 0) Then
            frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Check Revision Failed, sanity failed")
            frmChat.DoDisconnect()
            Exit Sub
        End If
		
		keys = GetCDKeyCount
		
		With pBuff
			.InsertDWord(ds.ClientToken) 'Client Token
			.InsertDWord(ds.CRevVersion) 'CRev Version
			.InsertDWord(ds.CRevChecksum) 'CRev Checksum
			.InsertDWord(keys) 'CDKey Count
			.InsertBool(CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
			
			For i = 1 To keys
				If (i = 1) Then
					sKey = BotVars.CDKey
				ElseIf (i = 2) Then 
					sKey = BotVars.ExpKey
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sKey = ReadCfg("Main", StringFormat("CDKey{0}", i))
				End If
				
				'Initialize the key decoder and validate the key.
				oKey.Initialize(sKey)
				If Not oKey.IsValid Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "Your CD-Key is invalid.")
					frmChat.DoDisconnect()
					Exit Sub
				End If
				
				'Calculate the hash
				If Not oKey.CalculateHash(ds.ClientToken, ds.ServerToken, BNCS_NLS) Then Exit Sub
				
				.InsertDWord(oKey.KeyLength)
				.InsertDWord(oKey.ProductValue)
				.InsertDWord(oKey.PublicValue)
				.InsertDWord(0)
                .InsertByteArr(oKey.Hash)
			Next i
			
			.InsertNTString((ds.CRevResult))
            If (Len(Config.CDKeyOwnerName) > 0) Then
                .InsertNTString((Config.CDKeyOwnerName))
            Else
                .InsertNTString((BotVars.Username))
            End If
			
			.SendPacket(SID_AUTH_CHECK)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		'UPGRADE_NOTE: Object oKey may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oKey = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_CHECK()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**********************************
	'SID_AUTH_ACCOUNTCREATE (0x52) S->C
	'**********************************
	' (DWORD) Status
	'**********************************
	Private Sub RECV_SID_AUTH_ACCOUNTCREATE(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lResult As Integer
		
		lResult = pBuff.GetDWORD
		
		Select Case lResult
			Case &H0
				Call Event_LogonEvent(4)
				
				'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
				If frmChat.sckBNet.CtlState = 7 Then
					Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending login information...")
					SEND_SID_AUTH_ACCOUNTLOGON()
					Exit Sub
				End If
				
			Case &H4 : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name already exists.")
			Case &H7 : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name is too short/blank.")
			Case &H8 : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name contains an illegal character.")
			Case &H9 : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name contains an illegal word.")
			Case &HA : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name contains too few alphanumeric characters.")
			Case &HB : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name contains adjacent punctuation characters.")
			Case &HC : frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Account creation failed because your name contains too many punctuation characters.")
			Case Else
				Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Account creation failed for an unknown reason: 0x{0}", ZeroOffset(lResult, 8)))
		End Select
		
		Call frmChat.DoDisconnect()
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTCREATE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**********************************
	'SID_AUTH_ACCOUNTCREATE (0x52) C->S
	'**********************************
	' (BYTE[32]) Salt (s)
	' (BYTE[32]) Verifier (v)
	' (STRING) Username
	'**********************************
	Private Sub SEND_SID_AUTH_ACCOUNTCREATE()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		ds.NLS.Initialize(BotVars.Username, BotVars.Password)
		With pBuff
			.InsertNonNTString(ds.NLS.SrpSalt)
			.InsertNonNTString(ds.NLS.Srpv)
			.InsertNTString(ds.NLS.Username)
			.SendPacket(SID_AUTH_ACCOUNTCREATE)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTCREATE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**********************************
	'SID_AUTH_ACCOUNTLOGON (0x53) S->C
	'**********************************
	' (DWORD) Status
	' (BYTE[32]) Salt (s)
	' (BYTE[32]) Server Key (B)
	'**********************************
	Private Sub RECV_SID_AUTH_ACCOUNTLOGON(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lResult As Integer
		Dim s As String
		Dim B As String
		
		lResult = pBuff.GetDWORD
		ds.NLS.SrpSalt = pBuff.GetRaw(32)
		ds.NLS.SrpB = pBuff.GetRaw(32)
		
		Select Case lResult
			Case &H0 'Accepted, requires proof.
				SEND_SID_AUTH_ACCOUNTLOGONPROOF()
				
			Case &H1 'Account doesn't exist.
				Call Event_LogonEvent(0)
				Call Event_LogonEvent(3)
				Call SEND_SID_AUTH_ACCOUNTCREATE()
				
			Case &H5 'Account requires upgrade, Not possible anymore
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Your account needs to be upgraded. This is no longer possible on Battle.net. Choose a different account.")
				
			Case Else
				Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown response to SID_AUTH_ACCOUNTLOGON: 0x{0}", ZeroOffset(lResult, 8)))
				frmChat.DoDisconnect()
		End Select
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTLOGON()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**********************************
	'SID_AUTH_ACCOUNTLOGON (0x53) C->S
	'**********************************
	' (BYTE[32]) Client Key ('A')
	' (STRING) Username
	'**********************************
	Private Sub SEND_SID_AUTH_ACCOUNTLOGON()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		pBuff.InsertNonNTString(ds.NLS.SrpA())
		pBuff.InsertNTString(ds.NLS.Username)
		pBuff.SendPacket(SID_AUTH_ACCOUNTLOGON)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTLOGON()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**************************************
	'SID_AUTH_ACCOUNTLOGONPROOF (0x54) S->C
	'**************************************
	' (DWORD) Status
	' (BYTE[20]) Server Password Proof (M2)
	' (STRING) Additional information
	'**************************************
	Private Sub RECV_SID_AUTH_ACCOUNTLOGONPROOF(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lResult As Integer
		Dim M2 As String
		Dim sInfo As String
		
		lResult = pBuff.GetDWORD
		M2 = pBuff.GetRaw(20)
		sInfo = pBuff.GetString
		
		Select Case lResult
			Case &H0 'Logon successful.
				Call Event_LogonEvent(2)
				If (Not ds.NLS.SrpVerifyM2(M2)) Then
					frmChat.AddChat(RTBColors.InformationText, "[BNCS] Warning, The server sent an invalid password proof, it may be a fake server.")
				End If
				SendEnterChatSequence()
				
			Case &H2 'Invalid password
				Call Event_LogonEvent(1)
				Call frmChat.DoDisconnect()
				
			Case &HE : DoRegisterEmail() 'Email registration requried
			Case &HF 'Custom message
				Call Event_LogonEvent(5, sInfo)
				Call frmChat.DoDisconnect()
				
			Case Else
				Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown response to SID_AUTH_ACCOUNTLOGONPROOF: 0x{0}", ZeroOffset(lResult, 8)))
				Call frmChat.DoDisconnect()
				
		End Select
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTLOGONPROOF()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**************************************
	'SID_AUTH_ACCOUNTLOGONPROOF (0x54) C->S
	'**************************************
	' (BYTE[20]) Client Password Proof (M1)
	'**************************************
	Private Sub SEND_SID_AUTH_ACCOUNTLOGONPROOF()
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertNonNTString(ds.NLS.SrpM1)
			.SendPacket(SID_AUTH_ACCOUNTLOGONPROOF)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTLOGONPROOF()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'SID_SETEMAIL (0x59) S->C
	'*******************************
	' [Blank]
	'*******************************
	Private Sub RECV_SID_SETEMAIL(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		' do not call into EmailReg here,
		' let receiving a successful account logon response call into it!
		' XXX: is there any case in which SID_SETEMAIL will be received after a successful account logon (instead of before)?
		ds.WaitingForEmail = True
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_SID_SETEMAIL()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'**************************************
	'SID_SETEMAIL (0x59) C->S
	'**************************************
	' (STRING) Email Address
	'**************************************
	Public Sub SEND_SID_SETEMAIL(ByRef sEMailAddress As String)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		With pBuff
			.InsertNTString(sEMailAddress)
			.SendPacket(SID_SETEMAIL)
		End With
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_SID_SETEMAIL()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'=======================================================================================================
	'This function will open the form to prompt the user for their email, or if the overrides are set, automatically register an email.
	Private Sub DoRegisterEmail()
		On Error GoTo ERROR_HANDLER
		
		Dim EMailValue As String
		Dim EMailAction As String
		
		EMailAction = Config.RegisterEmailAction
		EMailValue = Config.RegisterEmailDefault
		
		Call frmEMailReg.DoRegisterEmail(EMailAction, EMailValue)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.DoRegisterEmail()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'=======================================================================================================
	'This function will attempt to complete the CRev request that Bnet has sent to us.
	'Returns True if successful.
	Private Function CompileCheckrevision() As Boolean
		On Error GoTo ERROR_HANDLER
		Dim lVersion As Integer
		Dim lChecksum As Integer
		Dim sResult As String
		Dim sHeader As String
		Dim sFile As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sHeader = StringFormat("CRev_{0}", GetProductKey)
		If (Warden_CheckRevision((ds.CRevFileName), (ds.CRevFileTimeRaw), (ds.CRevSeed), sHeader, lVersion, lChecksum, sResult)) Then
			ds.CRevChecksum = lChecksum
			ds.CRevResult = sResult
			ds.CRevVersion = lVersion
			CompileCheckrevision = True
		Else
			Call frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Local Hashing Failed")
			CompileCheckrevision = False
		End If
		Exit Function
ERROR_HANDLER: 
		CompileCheckrevision = False
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.CompileCheckrevision()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Function GetCDKeyCount(Optional ByRef sProduct As String = vbNullString) As Integer
		On Error GoTo ERROR_HANDLER
		
		Dim sOverride As String
		Dim lRet As Integer
		
        If (Len(sProduct) = 0) Then sProduct = BotVars.Product
		
		lRet = GetProductInfo(sProduct).KeyCount
		
		GetCDKeyCount = lRet
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.GetCDKeyCount()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Function GetLogonSystem(Optional ByRef sProduct As String = vbNullString) As Integer
		On Error GoTo ERROR_HANDLER
		
		Dim sOverride As String
		Dim tLng As Integer
		Dim lRet As Integer
		
		' Temporary short-circuit:
		'  Return BNCS_NLS because no other login sequences are supported
		'  -andy
		'GetLogonSystem = BNCS_NLS
		'Exit Function
		
        If (Len(sProduct) = 0) Then sProduct = BotVars.Product
		
		lRet = GetProductInfo(sProduct).LogonSystem
		
		tLng = Config.GetLogonSystem(GetProductKey(sProduct))
		If tLng > -1 Then
			Select Case tLng
				Case BNCS_NLS : lRet = BNCS_NLS
				Case BNCS_LLS : lRet = BNCS_LLS
				Case BNCS_OLS : lRet = BNCS_OLS
			End Select
		End If
		
		GetLogonSystem = lRet
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.GetLogonSystem()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	'Converts the normalized (forward: IX86, STAR, etc) string representation of a DWORD into it's numeric equivalent.
	Private Function GetDWORDOverride(ByVal sDwordString As String, Optional ByVal lDefault As Integer = 0) As Integer
		On Error GoTo ERROR_HANDLER
		
		Dim lRet As Integer
		lRet = lDefault
		
        If ((Len(sDwordString) > 0) And (Len(sDwordString) < 5)) Then
            lRet = GetDWORD(StrReverse(sDwordString))
        End If
		
		GetDWORDOverride = lRet
		Exit Function
ERROR_HANDLER: 
		GetDWORDOverride = lRet
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.GetDWORDOverride()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Private Function GetDWORD(ByRef sData As String) As Integer
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sData = Left(StringFormat("{0}{1}", sData, New String(Chr(0), 4)), 4)
		CopyMemory(GetDWORD, sData, 4)
		
		Exit Function
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.GetDWORD()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	Public Sub SendEnterChatSequence()
		On Error GoTo ERROR_HANDLER
		If ds.EnteredChatFirstTime Then
			Call DoChannelJoinHome(False)
		Else
			ds.EnteredChatFirstTime = True
			
			If ((Not BotVars.Product = "VD2D") And (Not BotVars.Product = "PX2D") And (Not BotVars.Product = "PX3W") And (Not BotVars.Product = "3RAW")) Then
				
				If (Not BotVars.UseUDP) Then
					SEND_SID_UDPPINGRESPONSE()
					'We dont use ICONDATA .SendPacket SID_GETICONDATA
				End If
			End If
			
			SEND_SID_ENTERCHAT()
			SEND_SID_GETCHANNELLIST()
			
			BotVars.Gateway = Config.PredefinedGateway
            If (Len(BotVars.Gateway) = 0) Then
                If ((Not BotVars.Product = "VD2D") And (Not BotVars.Product = "PX2D") And (Not BotVars.Product = "PX3W") And (Not BotVars.Product = "3RAW")) Then
                    ' join nowhere to force non-W3-non-D2 to enter chat environment
                    ' so they can use /whoami (see Event_ChannelJoinError for where this completes)
                    Call FullJoin(BotVars.Product & BotVars.Username & Config.HomeChannel, 0)
                Else
                    SEND_SID_CHATCOMMAND("/whoami")
                End If
            Else
                'PvPGN: Straight home
                Call DoChannelJoinHome()
            End If
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SendEnterChatSequence()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' do channel join home (first time?)
	Public Sub DoChannelJoinHome(Optional ByRef FirstTime As Boolean = True)
		On Error GoTo ERROR_HANDLER
		
        If (FirstTime And Config.DefaultChannelJoin) Or (Len(Config.HomeChannel) = 0 And Len(BotVars.LastChannel) = 0) Then
            ' product home override or home/last channel are both empty
            Call DoChannelJoinProductHome()
        End If
		
        If (Len(BotVars.LastChannel) > 0) Then
            ' go to "last channel" (for /reconnect and re-entering chat)
            Call FullJoin((BotVars.LastChannel), 2)
        ElseIf (Len(Config.HomeChannel) > 0) Then
            ' go home
            Call FullJoin((Config.HomeChannel), 2)
        End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinHome()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	
	Public Sub DoChannelJoinProductHome()
		On Error GoTo ERROR_HANDLER
		
		' empty homechannel or
		' config override to force joinhome
		If BotVars.Product = "PX2D" Or BotVars.Product = "VD2D" Then
			Call FullJoin((BotVars.Product), 5)
		Else
			Call FullJoin((BotVars.Product), 1)
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinProductHome()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Sub DoQueryRealms()
		On Error GoTo ERROR_HANDLER
		
		ds.MCPHandler = New clsMCPHandler
		
		SEND_SID_QUERYREALMS2()
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinHome()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	Public Function CanSpawn(ByVal sProduct As String, ByVal iKeyLength As Short) As Boolean
		sProduct = GetProductInfo(sProduct).Code
		
		Select Case sProduct
			Case PRODUCT_STAR, PRODUCT_JSTR, PRODUCT_W2BN
				CanSpawn = CBool(iKeyLength <> 26)
				Exit Function
		End Select
		CanSpawn = False
	End Function
	
	Public Sub HandleEmptyCRevSeed()
		frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] CheckRevision seed was returned empty! This is usually due to an unrecognized verison byte.")
		If (BotVars.BNLS) Then
			frmChat.HandleBnlsError("[BNCS] The BNLS server you are using may be misconfigured.")
		Else
			frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] You can reset your version bytes to the latest by going to Bot -> Update Version Bytes")
			frmChat.DoDisconnect()
		End If
	End Sub
End Module