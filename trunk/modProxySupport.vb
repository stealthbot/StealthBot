Option Strict Off
Option Explicit On
Module modProxySupport
	
	'modProxySupport - project StealthBot
	'February 2004 - by Stealth [stealth at stealthbot dot net]
	
	Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Short) As Short
	Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Integer
	
	Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Integer
	Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inaddr As Integer) As Integer
    Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Integer
    Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Integer
	Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Integer
	Public Declare Function WSACleanup Lib "wsock32.dll" () As Integer
	
	Public Structure HOSTENT
		Dim h_name As Integer
		Dim h_aliases As Integer
		Dim h_addrtype As Short
		Dim h_length As Short
		Dim h_addr_list As Integer
	End Structure
	
	Public Const PROXY_LOGIN_FAILED As Integer = &H401
	Public Const PROXY_LOGGING_IN As Integer = &H402
	Public Const PROXY_IS_NOT_PUBLIC As Integer = &H403
	Public Const PROXY_LOGIN_SUCCESS As Integer = &H404
	Public Const PROXY_ATTEMPTING_LOGIN As Integer = &H405
	
	' Thanks to macyui for the original implementation of this subroutine.
	'   It has been heavily modified to conform with my needs.
	
	Public Sub LogonToProxy(ByRef ds As AxMSWinsockLib.AxWinsock, ByVal sServerIP As String, ByVal lPort As Integer, ByVal Socks5 As Boolean)
		Dim INet(3) As Byte
		Dim lServer As Integer
		Dim buf As String
		Dim lPort1 As Object
		Dim lPort2 As Integer
		Dim HostInfo As HOSTENT
		Dim ptrIP As Integer
		
		If Socks5 Then
			buf = Chr(5) & Chr(1)
		Else
			buf = Chr(4) & Chr(1) 'connect request
		End If
		
		'Do we have an IP address or a hostname?
		If Not IsValidIPAddress(sServerIP) Then
			sServerIP = ResolveHost(sServerIP)
			
			If sServerIP = vbNullString Then
				Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "[PROXY] Unable to resolve hostname. Error code: 0x" & Right(New String("0", 8) & Hex(WSAGetLastError()), 8))
				Call frmChat.DoDisconnect()
				Exit Sub
			End If
		End If
		
		'Make the IP into a long.
		lServer = inet_addr(sServerIP)
		
		'Copy ip long into a struct with 4 byte members
		CopyMemory(INet(0), lServer, 4)
		
		If Socks5 Then
			'Method (0x00 No authentication)
			buf = buf & Chr(0)
			'protocol version VER
			buf = buf & Chr(5)
			'command CMD 0x01 CONNECT
			buf = buf & Chr(1)
			'reserved
			buf = buf & Chr(0)
			'IPv4 address type (0x01)
			buf = buf & Chr(1)
			
			'dest address
			buf = buf & Chr(INet(0)) & Chr(INet(1)) & Chr(INet(2)) & Chr(INet(3))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lPort1 = CInt(lPort) \ 256
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lPort2 = CInt(lPort) - htons(lPort1)
			
			'dest port
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			buf = buf & Chr(lPort1) & Chr(lPort2)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lPort1 = CInt(lPort) \ 256
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lPort2 = CInt(lPort) - htons(lPort1)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object lPort1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			buf = buf & Chr(lPort1) & Chr(lPort2)
			
			buf = buf & Chr(INet(0)) & Chr(INet(1)) & Chr(INet(2)) & Chr(INet(3))
			
			buf = buf & Chr(0)
		End If
		
		ds.SendData(buf)
	End Sub
	
	Public Sub ParseProxyPacket(ByVal Data As String)
		Dim s0ut As String
		Dim lColor As Integer
		Dim Message As Byte
		Static ReceivedMethodReply As Boolean
		
		Select Case Asc(Mid(Data, 1, 1))
			Case &H0, &H4, &H5
				Message = Asc(Mid(Data, 2, 1))
				
				If Not BotVars.ProxyIsSocks5 Then
					Select Case Message
						Case 90
							s0ut = "[PROXY] SOCKS4 request granted."
							lColor = RTBColors.SuccessText
							UpdateProxyStatus(modEnum.enuProxyStatus.psOnline)
							frmChat.InitBNetConnection()
						Case 91
							s0ut = "[PROXY] Request rejected or failed."
							UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
							lColor = RTBColors.ErrorMessageText
						Case 92
							s0ut = "[PROXY] Request rejected: the SOCKS server cannot connect to identd on the client."
							UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
							lColor = RTBColors.ErrorMessageText
						Case 93
							s0ut = "[PROXY] Request rejected: client program and identd report different userids."
							UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
							lColor = RTBColors.ErrorMessageText
						Case Else
							s0ut = "[PROXY] Unknown/unhandled proxy message. ID: " & Message
							lColor = RTBColors.InformationText
							UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
							
					End Select
					
					If Message >= 91 Then
						'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						If frmChat.sckBNet.CtlState <> 0 Then
							frmChat.sckBNet.Close()
						End If
					End If
				Else
					'o  REP    Reply field:
					s0ut = "[PROXY] "
					
					Select Case Message
						Case 0 '00' succeeded
							If ReceivedMethodReply Then
								ReceivedMethodReply = False
								s0ut = s0ut & "SOCKS5 logon succeeded."
								lColor = RTBColors.SuccessText
								
								UpdateProxyStatus(modEnum.enuProxyStatus.psOnline)
								
								frmChat.InitBNetConnection()
							Else
								ReceivedMethodReply = True
							End If
							
						Case 1 '01' general SOCKS server failure
							s0ut = s0ut & "General server failure."
							lColor = RTBColors.ErrorMessageText
						Case 2 '02' connection not allowed by ruleset
							s0ut = s0ut & "Connection not allowed by server ruleset."
							lColor = RTBColors.ErrorMessageText
						Case 3 '03' Network unreachable
							s0ut = s0ut & "Destination network unreachable."
							lColor = RTBColors.ErrorMessageText
						Case 4 '04' Host unreachable
							s0ut = s0ut & "Destination host unreachable."
							lColor = RTBColors.ErrorMessageText
						Case 5 '05' Connection refused
							s0ut = s0ut & "Connection refused."
							lColor = RTBColors.ErrorMessageText
						Case 6 '06' TTL expired
							s0ut = s0ut & "Server-side TTL expiration."
							lColor = RTBColors.ErrorMessageText
						Case 7 '07' Command not supported
							s0ut = s0ut & "Command not supported."
							lColor = RTBColors.ErrorMessageText
						Case 8 '08' Address type not supported
							s0ut = s0ut & "Address type not supported."
							lColor = RTBColors.ErrorMessageText
						Case Else '09' to X'FF' unassigned
							s0ut = s0ut & "Unknown response code."
							lColor = RTBColors.ErrorMessageText
					End Select
					
					If Message > 0 Then UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
					
					s0ut = s0ut & " (" & Message & ")"
				End If
				
			Case Else
				s0ut = "[PROXY] Unknown/unhandled proxy message. Message ID: " & Asc(Mid(Data, 2, 1)) & ", VN Byte: " & Asc(Mid(Data, 1, 1))
				lColor = RTBColors.InformationText
				UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
				
		End Select
		
		If lColor > 0 Then frmChat.AddChat(lColor, s0ut)
	End Sub
	
	Sub UpdateProxyStatus(ByVal NewStatus As modEnum.enuProxyStatus, Optional ByVal AddtlInfo As Integer = 0)
		BotVars.ProxyStatus = NewStatus
		
		Dim sOut As String
		Dim lColor As Integer
		
		Select Case AddtlInfo
			Case PROXY_LOGIN_FAILED
				lColor = RTBColors.ErrorMessageText
				sOut = "[PROXY] Login failed."
				
			Case PROXY_LOGGING_IN
				lColor = RTBColors.SuccessText
				sOut = "[PROXY] Logging in..."
				
			Case PROXY_IS_NOT_PUBLIC
				lColor = RTBColors.ErrorMessageText
				sOut = "[PROXY] This proxy does not allow anonymous logons."
				
			Case PROXY_LOGIN_SUCCESS
				lColor = RTBColors.SuccessText
				sOut = "[PROXY] Logon successful."
				
			Case PROXY_ATTEMPTING_LOGIN
				lColor = RTBColors.InformationText
				sOut = "[PROXY] Attempting to log onto proxy..."
				
		End Select
		
        If Len(sOut) > 0 Then frmChat.AddChat(lColor, sOut)
	End Sub
End Module