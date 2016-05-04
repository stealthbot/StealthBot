Option Strict Off
Option Explicit On
Friend Class clsDataBuffer
	' clsDataBuffer.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_buf() As Byte
	Private m_bufpos As Short
	Private m_bufsize As Short
	Private m_cripple As Boolean
	
	Public Sub New()
		MyBase.New()

        ' clear buffer contents
        Clear()
	End Sub

	Protected Overrides Sub Finalize()
        ' clear buffer contents
        Clear()

		MyBase.Finalize()
	End Sub
	
	Public Sub setCripple()
		m_cripple = True
	End Sub
	
	Public Function getCripple() As Boolean
		getCripple = m_cripple
	End Function
	
	
    Public Property Data() As Byte()
        Get
            ReDim Data(m_bufsize)
            m_buf.CopyTo(Data, 0)
        End Get

        Set(ByVal Value As Byte())
            ReDim m_buf(Value.Length)
            Value.CopyTo(m_buf, 0)
            m_bufsize = Value.Length
        End Set
    End Property
	
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
	
    Public Sub InsertByte(ByVal Data As Byte)
        ' resize buffer
        ReDim Preserve m_buf((m_bufsize + 1))

        ' copy data to buffer
        m_buf(m_bufsize) = Data

        ' store buffer Length
        m_bufsize = (m_bufsize + 1)
    End Sub
	
    Public Sub InsertByteArr(ByRef Data() As Byte)
        ' resize buffer
        ReDim Preserve m_buf((m_bufsize + (UBound(Data) + 1)))

        ' copy data to buffer
        Data.CopyTo(m_buf, m_bufsize)

        ' store buffer Length
        m_bufsize = (m_bufsize + (UBound(Data) + 1))
    End Sub
	
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
        Buffer.BlockCopy(m_buf, m_bufpos, Data, 0, Length)
		
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
	
    Public Sub InsertWord(ByVal Data As Short)

        Me.InsertByteArr(BitConverter.GetBytes(Data))

    End Sub
	
	Public Function GetWord(Optional ByRef Peek As Boolean = False) As Short
		If ((m_bufpos + 2) > m_bufsize) Then
			Exit Function
		End If

        GetWord = BitConverter.ToInt16(m_buf, m_bufpos)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 2)
	End Function
	
    Public Sub InsertDWord(ByVal Data As Integer)

        Me.InsertByteArr(BitConverter.GetBytes(Data))

    End Sub
	
	Public Function GetDWORD(Optional ByRef Peek As Boolean = False) As Integer
		If ((m_bufpos + 4) > m_bufsize) Then
			Exit Function
		End If
		
        GetDWORD = BitConverter.ToInt32(m_buf, m_bufpos)
		
		If (Not Peek) Then m_bufpos = (m_bufpos + 4)
	End Function
	
    Public Sub InsertBool(ByVal Data As Boolean)
        InsertDWord(Math.Abs(CInt(Data)))
    End Sub
	
	Public Function GetBool(Optional ByRef Peek As Boolean = False) As Boolean
		GetBool = Not CBool(GetDWORD(Peek) = 0)
	End Function
	
	Public Function GetFileTime(Optional ByRef Peek As Boolean = False) As Date
        Dim raw() As Byte
        Dim ft As FILETIME
		
		If ((m_bufpos + 8) > m_bufsize) Then
			Exit Function
		End If
		
        Me.GetByteArr(raw, 8, Peek)
        ft.dwLowDateTime = BitConverter.ToInt32(raw, 0)
        ft.dwHighDateTime = BitConverter.ToInt32(raw, 4)
		
		GetFileTime = FileTimeToDate(ft)
	End Function
	
    Public Sub InsertNonNTString(ByVal Data As String)

        Me.InsertByteArr(System.Text.Encoding.Default.GetBytes(Data))

    End Sub

    Public Sub InsertNTString(ByRef Data As String)

        Me.InsertNTString(Data, System.Text.Encoding.Default)

    End Sub

    Public Sub InsertNTString(ByRef Data As String, ByVal Encoding As System.Text.Encoding)

        If (Data.Length > 0) Then
            Me.InsertByteArr(Encoding.GetBytes(Data))
        End If

        Me.InsertByte(0)    ' null terminator

    End Sub

    Public Function GetString()
        GetString = Me.GetString(System.Text.Encoding.Default)
    End Function

    Public Function GetString(ByVal Encoding As System.Text.Encoding, Optional ByRef Peek As Boolean = False) As String
        Dim i As Short
        Dim raw() As Byte

        ' Find the null terminator
        For i = m_bufpos To m_bufsize
            If (m_buf(i) = &H0) Then
                Exit For
            End If
        Next i

        If (i < m_bufsize) Then
            Me.GetByteArr(raw, i - m_bufpos, Peek)

            GetString = Encoding.GetString(raw)
        End If
    End Function
	
	Public Function GetRaw(Optional ByVal Length As Short = -1, Optional ByRef Peek As Boolean = False) As String
        Dim raw() As Byte

        If (Length = -1) Then
            Length = m_bufsize - m_bufpos
        End If
		
		If ((m_bufpos + Length) > m_bufsize) Then
			Exit Function
        End If

        Me.GetByteArr(raw, Length, Peek)
        GetRaw = System.Text.Encoding.Default.GetString(raw)

	End Function
	
    Public Sub Clear()
        ' resize buffer
        ReDim m_buf(0)

        ' clear first index
        m_buf(0) = 0

        ' reset buffer Length
        m_bufsize = 0

        ' reset buffer position
        m_bufpos = 0
    End Sub
	
	Public Function DebugOutput() As String
        DebugOutput = modWar3Clan.DebugOutput(System.Text.Encoding.Default.GetString(Data))
	End Function
	
    Public Sub SendPacketMCP(Optional ByRef PacketID As Byte = 0)
        Dim buf() As Byte
        Dim strbuf As String
        Dim veto As Boolean

        If (frmChat.sckMCP.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
            Clear()
            Exit Sub
        End If

        ' resize temporary data buffer
        ReDim buf(m_bufsize + 2)

        ' copy packet data Length to temporary buffer
        BitConverter.GetBytes(Convert.ToInt16(m_bufsize + 3)).CopyTo(buf, 0)

        buf(2) = PacketID ' packet identification number

        ' copy data from buffer to temporary buffer
        If (m_bufsize > 0) Then
            m_buf.CopyTo(buf, 3)
        End If

        strbuf = System.Text.Encoding.Default.GetString(buf)

        If (Not RunInAll("Event_PacketSent", "MCP", PacketID, m_bufsize + 3, strbuf)) Then
            If (MDebug("all")) Then
                frmChat.AddChat(COLOR_BLUE, "MCP SEND 0x" & ZeroOffset(PacketID, 2))
            End If

            Send(frmChat.sckMCP.SocketHandle, buf, m_bufsize + 3, 0)

            CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stMCP, PacketID, m_bufsize + 3, buf)

            WritePacketData(modEnum.enuPL_ServerTypes.stMCP, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 3, buf)
        End If

        ' clear buffer contents
        Clear()
    End Sub
	
    Public Sub SendPacket(Optional ByRef PacketID As Byte = 0)
        Dim L As Integer
        Dim buf() As Byte
        Dim i As Short
        Dim strbuf As String
        Dim veto As Boolean

        If (frmChat.sckBNet.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
            Clear()
            Exit Sub
        End If

        If (getCripple()) Then
            Select Case (PacketID)
                Case SID_CHATCOMMAND, SID_JOINCHANNEL
                    Clear()
                    Exit Sub
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
        BitConverter.GetBytes(Convert.ToInt16(m_bufsize + 4)).CopyTo(buf, 2)

        ' copy data from buffer to temporary buffer
        If (m_bufsize > 0) Then
            m_buf.CopyTo(buf, 4)
        End If

        strbuf = System.Text.Encoding.Default.GetString(buf)

        If (Not RunInAll("Event_PacketSent", "BNCS", PacketID, m_bufsize + 4, strbuf)) Then

            If (MDebug("all")) Then
                frmChat.AddChat(COLOR_BLUE, "BNET SEND 0x" & ZeroOffset(PacketID, 2))
            End If

            Send(frmChat.sckBNet.SocketHandle, buf, m_bufsize + 4, 0)

            CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stBNCS, PacketID, m_bufsize + 4, buf)

            WritePacketData(modEnum.enuPL_ServerTypes.stBNCS, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 4, buf)

            'Send Warden Everything thats Sent to Bnet
            Call modWarden.WardenData(WardenInstance, buf, True)
        End If
        ' clear buffer contents
        Clear()
    End Sub
	
    Public Sub vLSendPacket(Optional ByRef PacketID As Byte = 0)
        Dim buf() As Byte
        Dim strbuf As String
        Dim veto As Boolean

        If (frmChat.sckBNLS.CtlState <> MSWinsockLib.StateConstants.sckConnected) Then
            Clear()
            Exit Sub
        End If

        ' resize temporary data buffer
        ReDim buf(m_bufsize + 2)

        ' copy packet data Length to temporary buffer
        BitConverter.GetBytes(Convert.ToInt16(m_bufsize + 3)).CopyTo(buf, 0)

        buf(2) = PacketID ' packet identification number

        ' copy data from buffer to temporary buffer
        If (m_bufsize > 0) Then
            m_buf.CopyTo(buf, 3)
        End If

        strbuf = System.Text.Encoding.Default.GetString(buf)

        If (Not RunInAll("Event_PacketSent", "BNLS", PacketID, m_bufsize + 3, strbuf)) Then
            If (MDebug("all")) Then
                frmChat.AddChat(COLOR_BLUE, "BNLS SEND 0x" & ZeroOffset(PacketID, 2))
            End If

            Send(frmChat.sckBNLS.SocketHandle, buf, m_bufsize + 3, 0)

            CachePacket(modEnum.enuPL_DirectionTypes.CtoS, modEnum.enuPL_ServerTypes.stBNLS, PacketID, m_bufsize + 3, buf)

            WritePacketData(modEnum.enuPL_ServerTypes.stBNLS, modEnum.enuPL_DirectionTypes.CtoS, PacketID, m_bufsize + 3, buf)
        End If

        ' clear buffer contents
        Clear()
    End Sub
End Class