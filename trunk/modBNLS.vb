Option Strict Off
Option Explicit On
Module modBNLS
	Private Const OBJECT_NAME As String = "modBNLS"
	
	Private Const BNLS_CHOOSENLSREVISION As Byte = &HD
	Private Const BNLS_AUTHORIZE As Byte = &HE
	Private Const BNLS_AUTHORIZEPROOF As Byte = &HF
	Private Const BNLS_REQUESTVERSIONBYTE As Byte = &H10
	Private Const BNLS_VERSIONCHECKEX2 As Byte = &H1A
	
	Public BNLSAuthorized As Boolean
	
    Public Function BNLSRecvPacket(ByVal Data() As Byte) As Boolean
        On Error GoTo ERROR_HANDLER
        Static pBuff As New clsDataBuffer

        Dim PacketID As Byte

        BNLSRecvPacket = True
        With pBuff
            .Clear()
            .Data = Data
            .GetWord()
            PacketID = .GetByte
        End With

        If (MDebug("all")) Then
            frmChat.AddChat(COLOR_BLUE, StringFormat("BNLS RECV 0x{0}", ZeroOffset(PacketID, 2)))
        End If

        CachePacket(modEnum.enuPL_DirectionTypes.StoC, modEnum.enuPL_ServerTypes.stBNLS, PacketID, Len(Data), Data)

        ' Added 2007-06-08 for a packet logging menu feature to aid tech support
        WritePacketData(modEnum.enuErrorSources.BNLS, modEnum.enuPL_DirectionTypes.StoC, PacketID, Len(Data), Data)

        If (RunInAll("Event_PacketReceived", "BNLS", PacketID, Len(Data), Data)) Then
            Exit Function
        End If

        Select Case PacketID

            Case BNLS_AUTHORIZE : Call RECV_BNLS_AUTHORIZE(pBuff)
            Case BNLS_AUTHORIZEPROOF : Call RECV_BNLS_AUTHORIZEPROOF(pBuff)
            Case BNLS_REQUESTVERSIONBYTE : Call RECV_BNLS_REQUESTVERSIONBYTE(pBuff)
            Case BNLS_VERSIONCHECKEX2 : Call RECV_BNLS_VERSIONCHECKEX2(pBuff)

            Case Else
                BNLSRecvPacket = False
                If (MDebug("debug") And (MDebug("all") Or MDebug("unknown"))) Then
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNLS] Unhandled packet 0x{0}", ZeroOffset(CInt(PacketID), 2)))
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNLS] Packet data: {0}{1}", vbNewLine, DebugOutput(Data)))
                End If

        End Select

        Exit Function
ERROR_HANDLER:
        Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.BNLSRecvPacket()", Err.Number, Err.Description, OBJECT_NAME))
    End Function
	
	'*******************************
	'BNLS_AUTHORIZE (0x0E) S->C
	'*******************************
	' (DWORD) Server Token
	'*******************************
	Private Sub RECV_BNLS_AUTHORIZE(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Call SEND_BNLS_AUTHORIZEPROOF(pBuff.GetDWORD)
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_AUTHORIZE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'BNLS_AUTHORIZE (0x0E) C->S
	'*******************************
	' (String) Bot ID
	'*******************************
	Public Sub SEND_BNLS_AUTHORIZE(Optional ByRef sBotID As String = vbNullString)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
        pBuff.InsertNTString(IIf(Len(sBotID) = 0, "stealth", sBotID))
		pBuff.vLSendPacket(BNLS_AUTHORIZE)
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_AUTHORIZE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'BNLS_AUTHORIZEPROOF (0x0F) S->C
	'*******************************
	' (DWORD) Server Token
	'*******************************
	Private Sub RECV_BNLS_AUTHORIZEPROOF(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		BNLSAuthorized = True
		Call frmChat.Event_BNLSAuthEvent(True)
		Call frmChat.Event_BNetConnecting()
		frmChat.sckBNet.Connect()
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_AUTHORIZEPROOF()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'*******************************
	'BNLS_AUTHORIZEPROOF (0x0F) C->S
	'*******************************
	' (DWORD) Password Checksum
	'*******************************
	Private Sub SEND_BNLS_AUTHORIZEPROOF(ByRef lServerToken As Integer, Optional ByRef sPassword As String = vbNullString)
		On Error GoTo ERROR_HANDLER
		
		Dim cCRC As New clsCRC32
		Dim lChecksum As Integer
		Dim pBuff As New clsDataBuffer
		
        lChecksum = cCRC.CRC32(StringFormat("{0}{1}", IIf(Len(sPassword) = 0, "gn1ftx14oc", sPassword), ZeroOffset(lServerToken, 8)))
		
		
		pBuff.InsertDWord(lChecksum)
		pBuff.vLSendPacket(BNLS_AUTHORIZEPROOF)
		
		'UPGRADE_NOTE: Object cCRC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cCRC = Nothing
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_AUTHORIZEPROOF()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'************************************
	'BNLS_REQUESTVERSIONBYTE (0x10) S->C
	'************************************
	' (DWORD) Product ID (0 if failure)
	' (DWORD) Version Byte (Not included if failure)
	'************************************
	Private Sub RECV_BNLS_REQUESTVERSIONBYTE(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		
		Dim lVerByte As Integer
		
		If (Not pBuff.GetDWORD = 0) Then
			lVerByte = pBuff.GetDWORD
			Config.SetVersionByte(GetProductKey(), lVerByte) 'Save BNLS's Version Byte
			Call Config.Save()
			
			Select Case modBNCS.GetLogonSystem()
				Case modBNCS.BNCS_NLS : Call modBNCS.SEND_SID_AUTH_INFO(lVerByte)
				Case modBNCS.BNCS_OLS
					modBNCS.SEND_SID_CLIENTID2()
					modBNCS.SEND_SID_LOCALEINFO()
					modBNCS.SEND_SID_STARTVERSIONING(lVerByte)
				Case modBNCS.BNCS_LLS
					modBNCS.SEND_SID_CLIENTID()
					modBNCS.SEND_SID_STARTVERSIONING(lVerByte)
				Case Else
					frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem()))
					frmChat.AddChat(RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error.")
					frmChat.DoDisconnect()
			End Select
		Else
			frmChat.AddChat(RTBColors.ErrorMessageText, "[BNLS] Version byte request failed!")
			frmChat.DoDisconnect()
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_REQUESTVERSIONBYTE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'************************************
	'BNLS_REQUESTVERSIONBYTE (0x10) C->S
	'************************************
	' (DWORD) ProductID
	'************************************
	Public Sub SEND_BNLS_REQUESTVERSIONBYTE(Optional ByRef sProduct As String = vbNullString)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		
        pBuff.InsertDWord(GetBNLSProductID(IIf(Len(sProduct) = 0, BotVars.Product, sProduct)))
		pBuff.vLSendPacket(BNLS_REQUESTVERSIONBYTE)
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_REQUESTVERSIONBYTE()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'************************************
	'BNLS_VERSIONCHECKEX2 (0x1A) S->C
	'************************************
	' (DWORD) Success*
	' (DWORD) Version.
	' (DWORD) Checksum.
	' (STRING) Version check stat string.
	' (DWORD) Cookie.
	' (DWORD) Version Byte
	'************************************
	Private Sub RECV_BNLS_VERSIONCHECKEX2(ByRef pBuff As clsDataBuffer)
		On Error GoTo ERROR_HANDLER
		Dim lVersionByte As Integer
		
		If (pBuff.GetDWORD = 1) Then
			ds.CRevVersion = pBuff.GetDWORD
			ds.CRevChecksum = pBuff.GetDWORD
			ds.CRevResult = pBuff.GetString
			pBuff.GetDWORD()
			lVersionByte = pBuff.GetDWORD
			
			Select Case modBNCS.GetLogonSystem()
				Case modBNCS.BNCS_NLS : Call modBNCS.SEND_SID_AUTH_CHECK()
				Case modBNCS.BNCS_OLS : Call modBNCS.SEND_SID_REPORTVERSION(lVersionByte)
				Case modBNCS.BNCS_LLS : Call modBNCS.SEND_SID_REPORTVERSION(lVersionByte)
				Case Else
					frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem()))
					frmChat.AddChat(RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error.")
					frmChat.DoDisconnect()
			End Select
		Else
			frmChat.Event_BNLSDataError(2)
			frmChat.DoDisconnect()
		End If
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_VERSIONCHECKEX2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'************************************
	'BNLS_VERSIONCHECKEX2 (0x1A) C->S
	'************************************
	' (DWORD) Product ID.*
	' (DWORD) Flags.**
	' (DWORD) Cookie.
	' (FILETIME) CRev Archive File Time
	' (STRING) CRev Archive File Name
	' (STRING) CRev Seed Values
	'************************************
	Public Sub SEND_BNLS_VERSIONCHECKEX2(ByRef sCRevFileTime As String, ByRef sCRevFileName As String, ByRef sCRevSeeds As String, Optional ByRef sProduct As String = vbNullString, Optional ByRef lFlags As Integer = 0, Optional ByRef lCookie As Integer = 1)
		On Error GoTo ERROR_HANDLER
		
		Dim pBuff As New clsDataBuffer
		
		With pBuff
			.InsertDWord(GetBNLSProductID(sProduct))
			.InsertDWord(lFlags)
			.InsertDWord(lCookie)
			.InsertNonNTString(sCRevFileTime)
			.InsertNTString(sCRevFileName)
			.InsertNTString(sCRevSeeds)
			.vLSendPacket(BNLS_VERSIONCHECKEX2)
		End With
		
		'UPGRADE_NOTE: Object pBuff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBuff = Nothing
		
		Exit Sub
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_VERSIONCHECKEX2()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	'===================================================================================================
	Public Function GetBNLSProductID(Optional ByVal sProdID As String = vbNullString) As Integer
        If (Len(sProdID) = 0) Then sProdID = BotVars.Product
		
		GetBNLSProductID = GetProductInfo(sProdID).BNLS_ID
	End Function
End Module