Option Strict Off
Option Explicit On
Friend Class clsDataBuffer
	' clsDataBuffer.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_buf() As Byte
	Private m_bufpos As Short
	Private m_bufsize As Short
	Private m_cripple As Boolean
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' clear buffer contents
		Clear()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' clear buffer contents
		Clear()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub setCripple()
		m_cripple = True
	End Sub
	
	Public Function getCripple() As Boolean
		getCripple = m_cripple
	End Function
	
	
	Public Property Data() As String
		Get
			DataControl = New String(Chr(0), m_bufsize)
			CopyMemory(DataControl, m_buf(0), m_bufsize)
		End Get
		Set(ByVal Value As String)
			ReDim m_buf(Len(Value))
			CopyMemory(m_buf(0), Value, Len(Value))
			m_bufsize = Len(Value)
		End Set
	End Property ' end function GetString
	
	Public ReadOnly Property Length() As Integer
		Get
			Length = m_bufsize
		End Get
	End Property
	
	
	Public Property Position() As Integer
		Get
			Position = m_bufpos
		End Get
		Set(ByVal Value As Integer)
			m_bufpos = Value
		End Set
	End Property
	
	Public Function InsertByte(ByVal Data As Byte) As Object
		' resize buffer
		ReDim Preserve m_buf((m_bufsize + 1))
		
		' copy data to buffer
		m_buf(m_bufsize) = Data
		
		' store buffer Length
		m_bufsize = (m_bufsize + 1)
	End Function
	
	Public Function InsertByteArr(ByRef Data() As Byte) As Object
		' resize buffer
		ReDim Preserve m_buf((m_bufsize + (UBound(Data) + 1)))
		
		' copy data to buffer
		CopyMemory(m_buf(m_bufsize), Data(0), (UBound(Data) + 1))
		
		' store buffer Length
		m_bufsize = (m_bufsize + (UBound(Data) + 1))
	End Function
	
	Public Sub GetByteArr(ByRef Data() As Byte, Optional ByVal Length As Short = -1, Optional ByRef Peek As Boolean = False)
		If (Length = -1) Then
			Length = m_bufsize
		End If
		
		If ((m_bufpos + Length) > m_bufsize) Then
			Exit Sub
		End If
		
		' resize buffer
		ReDim Data(Length)
		
		' copy data to buffer
		CopyMemory(Data(0), m_buf(m_bufpos), Length)
		
		' store buffer Length
		If (Not Peek) Then m_bufpos = (m_bufpos + Length)
	End Sub
	
	Public Function GetByte(Optional ByRef Peek As Boolean = False) As Byte
		If ((m_bufpos + 1) > m_bufsize) Then
			Exit Function
		End If
		
		GetByte = m_buf(m_bufpos)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 1)
	End Function
	
	Public Function InsertWord(ByVal Data As Short) As Object
		' resize buffer
		ReDim Preserve m_buf((m_bufsize + 2))
		
		' copy data to buffer
		CopyMemory(m_buf(m_bufsize), Data, 2)
		
		' store buffer Length
		m_bufsize = (m_bufsize + 2)
	End Function
	
	Public Function GetWord(Optional ByRef Peek As Boolean = False) As Short
		If ((m_bufpos + 2) > m_bufsize) Then
			Exit Function
		End If
		
		' copy data to buffer
		CopyMemory(GetWord, m_buf(m_bufpos), 2)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 2)
	End Function
	
	Public Function InsertDWord(ByVal Data As Integer) As Object
		' resize data buffer
		ReDim Preserve m_buf((m_bufsize + 4))
		
		' copy data to buffer
		CopyMemory(m_buf(m_bufsize), Data, 4)
		
		' store buffer Length
		m_bufsize = (m_bufsize + 4)
	End Function
	
	Public Function GetDWORD(Optional ByRef Peek As Boolean = False) As Integer
		If ((m_bufpos + 4) > m_bufsize) Then
			Exit Function
		End If
		
		' copy data to buffer
		CopyMemory(GetDWORD, m_buf(m_bufpos), 4)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 4)
	End Function
	
	Public Function InsertBool(ByVal Data As Boolean) As Object
		InsertDWord(System.Math.Abs(CInt(Data)))
	End Function
	
	Public Function GetBool(Optional ByRef Peek As Boolean = False) As Boolean
		GetBool = Not CBool(GetDWORD(Peek) = 0)
	End Function
	
	Public Function GetFileTime(Optional ByRef Peek As Boolean = False) As Date
		Dim ft As FILETIME
		
		If ((m_bufpos + 8) > m_bufsize) Then
			Exit Function
		End If
		
		' copy data to buffer
		'UPGRADE_WARNING: Couldn't resolve default property of object ft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(ft, m_buf(m_bufpos), 8)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 8)
		
		GetFileTime = FileTimeToDate(ft)
	End Function
	
	Public Function InsertNonNTString(ByVal Data As String) As Object
		' resize buffer
		ReDim Preserve m_buf((m_bufsize + Len(Data)))
		
		' copy data to buffer
		CopyMemory(m_buf(m_bufsize), Data, Len(Data))
		
		' store buffer Length
		m_bufsize = (m_bufsize + Len(Data))
	End Function
	
	Public Function InsertNTString(ByRef Data As String, Optional ByVal Encoding As modPacketBuffer.STRINGENCODING = STRINGENCODING.ANSI_Renamed) As Object
		
		Dim arrStr() As Byte
		
		Select Case (Encoding)
			Case STRINGENCODING.ANSI_Renamed
				'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
				arrStr = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Data, vbFromUnicode))
				
			Case modPacketBuffer.STRINGENCODING.UTF8
				arrStr = modUTF8.UTF8Encode(Data)
				
			Case modPacketBuffer.STRINGENCODING.UTF16
				'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
				arrStr = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Data, vbUnicode))
		End Select
		
		' resize buffer and include terminating null character
		ReDim Preserve m_buf((m_bufsize + (UBound(arrStr) + 2)))
		
		' copy data to buffer
		If (Data <> vbNullString) Then
			CopyMemory(m_buf(m_bufsize), arrStr(0), (UBound(arrStr) + 1))
		End If
		
		' store buffer Length
		m_bufsize = (m_bufsize + (UBound(arrStr) + 2))
	End Function
	
	Public Function GetString(Optional ByVal Encoding As modPacketBuffer.STRINGENCODING = STRINGENCODING.ANSI_Renamed, Optional ByRef Peek As Boolean = False) As String
		Dim i As Short
		
		For i = m_bufpos To m_bufsize
			If (m_buf(i) = &H0) Then
				Exit For
			End If
		Next i
		
		If (i < m_bufsize) Then
			GetString = New String(Chr(0), i - m_bufpos)
			
			' copy data to buffer
			CopyMemory(GetString, m_buf(m_bufpos), i - m_bufpos)
			
			If (Not Peek) Then m_bufpos = i + 1
		End If
	End Function
	
	Public Function GetRaw(Optional ByVal Length As Short = -1, Optional ByRef Peek As Boolean = False) As String
		If (Length = -1) Then
			Length = m_bufsize - m_bufpos
		End If
		
		If ((m_bufpos + Length) > m_bufsize) Then
			Exit Function
		End If
		
		GetRaw = New String(Chr(0), Length)
		
		' copy data to buffer
		CopyMemory(GetRaw, m_buf(m_bufpos), Length)
		
		If (Not Peek) Then m_bufpos = m_bufpos + Length
	End Function
	
	Public Function Clear() As Object
		' resize buffer
		ReDim m_buf(0)
		
		' clear first index
		m_buf(0) = 0
		
		' reset buffer Length
		m_bufsize = 0
		
		' reset buffer position
		m_bufpos = 0
	End Function
	
	Public Function DebugOutput() As String
		DebugOutput = modWar3Clan.DebugOutput(Data)
	End Function
	
	Public Function SendPacketMCP(Optional ByRef PacketID As Byte = 0) As Object
		Dim buf() As Byte
		Dim strbuf As String
		Dim veto As Boolean
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckMCP.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
			Clear()
			Exit Function
		End If
		
		' resize temporary data buffer
		ReDim buf(m_bufsize + 2)
		
		' copy packet data Length to temporary buffer
		CopyMemory(buf(0), m_bufsize + 3, 2)
		
		buf(2) = PacketID ' packet identification number
		
		' copy data from buffer to temporary buffer
		If (m_bufsize) Then
			CopyMemory(buf(3), m_buf(0), m_bufsize)
		End If
		
		strbuf = New String(vbNullChar, m_bufsize + 3)
		
		CopyMemory(strbuf, buf(0), m_bufsize + 3)
		
		If (Not RunInAll("Event_PacketSent", "MCP", PacketID, m_bufsize + 3, strbuf)) Then
			If (MDebug("all")) Then
				frmChat.AddChat(COLOR_BLUE, "MCP SEND 0x" & ZeroOffset(PacketID, 2))
			End If
			
			Send(frmChat.sckMCP.SocketHandle, strbuf, m_bufsize + 3, 0)
			
			CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stMCP, PacketID, m_bufsize + 3, strbuf)
			
			WritePacketData(modEnum.enuPL_ServerTypes.stMCP, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 3, strbuf)
		End If
		
		' clear buffer contents
		Clear()
	End Function
	
	Public Function SendPacket(Optional ByRef PacketID As Byte = 0) As Object
		Dim L As Integer
		Dim buf() As Byte
		Dim i As Short
		Dim strbuf As String
		Dim veto As Boolean
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNet.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
			Clear()
			Exit Function
		End If
		
		If (getCripple) Then
			Select Case (PacketID)
				Case SID_CHATCOMMAND, SID_JOINCHANNEL
					Clear()
					Exit Function
			End Select
		End If
		
		' These two packets cause the client to leave chat, and do not have any responses.
		'  (SID_NOTIFYJOIN is not valid unless it's at least 10 bytes long)
		If (PacketID = SID_LEAVECHAT Or (PacketID = SID_NOTIFYJOIN And m_bufsize >= 10)) Then
			' there's no response to this one!
			Call Event_LeftChatEnvironment()
		End If
		
		' resize temporary data buffer
		ReDim buf(m_bufsize + 3)
		
		buf(0) = &HFF ' header
		buf(1) = PacketID ' packet identification number
		
		' copy packet data Length to temporary buffer
		CopyMemory(buf(2), m_bufsize + 4, 2)
		
		' copy data from buffer to temporary buffer
		If (m_bufsize) Then
			CopyMemory(buf(4), m_buf(0), m_bufsize)
		End If
		
		strbuf = New String(vbNullChar, m_bufsize + 4)
		
		CopyMemory(strbuf, buf(0), m_bufsize + 4)
		
		If (Not RunInAll("Event_PacketSent", "BNCS", PacketID, m_bufsize + 4, strbuf)) Then
			
			If (MDebug("all")) Then
				frmChat.AddChat(COLOR_BLUE, "BNET SEND 0x" & ZeroOffset(PacketID, 2))
			End If
			
			Send(frmChat.sckBNet.SocketHandle, strbuf, m_bufsize + 4, 0)
			
			CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stBNCS, PacketID, m_bufsize + 4, strbuf)
			
			WritePacketData(modEnum.enuPL_ServerTypes.stBNCS, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 4, strbuf)
			
			'Send Warden Everything thats Sent to Bnet
			Call modWarden.WardenData(WardenInstance, strbuf, True)
		End If
		' clear buffer contents
		Clear()
	End Function
	
	Public Function vLSendPacket(Optional ByRef PacketID As Byte = 0) As Object
		Dim buf() As Byte
		Dim strbuf As String
		Dim veto As Boolean
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (frmChat.sckBNLS.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
			Clear()
			Exit Function
		End If
		
		' resize temporary data buffer
		ReDim buf(m_bufsize + 2)
		
		' copy packet data Length to temporary buffer
		CopyMemory(buf(0), m_bufsize + 3, 2)
		
		buf(2) = PacketID ' packet identification number
		
		' copy data from buffer to temporary buffer
		If (m_bufsize) Then
			CopyMemory(buf(3), m_buf(0), m_bufsize)
		End If
		
		strbuf = New String(vbNullChar, m_bufsize + 3)
		
		CopyMemory(strbuf, buf(0), m_bufsize + 3)
		
		If (Not RunInAll("Event_PacketSent", "BNLS", PacketID, m_bufsize + 3, strbuf)) Then
			If (MDebug("all")) Then
				frmChat.AddChat(COLOR_BLUE, "BNLS SEND 0x" & ZeroOffset(PacketID, 2))
			End If
			
			Send(frmChat.sckBNLS.SocketHandle, strbuf, m_bufsize + 3, 0)
			
			CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stBNLS, PacketID, m_bufsize + 3, strbuf)
			
			WritePacketData(modEnum.enuPL_ServerTypes.stBNLS, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 3, strbuf)
		End If
		
		' clear buffer contents
		Clear()
	End Function
End Class