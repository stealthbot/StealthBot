Option Strict Off
Option Explicit On
Module modPacketBuffer
	' modDataBuffer.bas
	' Copyright (C) 2008 Eric Evans
	
	
	Private Const MAX_PACKET_CACHE_SIZE As Short = 100
	
	Private Structure PACKETCACHEITEM
		Dim Direction As modEnum.enuPL_DirectionTypes
		Dim PKT_Type As modEnum.enuPL_ServerTypes
		Dim ID As Byte
		Dim Length As Short
        Dim Data() As Byte
		Dim TimeDate As Date
	End Structure
	
	Public Enum STRINGENCODING
		ANSI_Renamed = 1
		UTF8 = 2
		UTF16 = 3
	End Enum
	
	Private m_cache() As PACKETCACHEITEM
	Private m_cache_count As Short
	
    Public Function CachePacket(ByRef Direction As modEnum.enuPL_DirectionTypes, ByRef PKT_Type As modEnum.enuPL_ServerTypes, ByRef ID As Byte, ByRef Length As Short, ByRef Data() As Byte) As Object

        Dim pkt As PACKETCACHEITEM

        With pkt
            .Direction = Direction
            .PKT_Type = PKT_Type
            .ID = ID
            .Length = Length
            .Data = Data
            .TimeDate = Now
        End With

        Dim I As Short
        If (m_cache_count + 1 >= MAX_PACKET_CACHE_SIZE) Then

            For I = 0 To m_cache_count - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object m_cache(I). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                m_cache(I) = m_cache(I + 1)
            Next I

            'UPGRADE_WARNING: Couldn't resolve default property of object m_cache(m_cache_count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_cache(m_cache_count) = pkt
        Else
            If (m_cache_count = 0) Then
                ReDim m_cache(0)
            Else
                ReDim Preserve m_cache(m_cache_count + 1)
            End If

            'UPGRADE_WARNING: Couldn't resolve default property of object m_cache(m_cache_count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_cache(m_cache_count) = pkt

            m_cache_count = m_cache_count + 1
        End If

    End Function
	
	Public Sub DumpPacketCache()
		
		Dim pkt As PACKETCACHEITEM
		Dim I As Short
		Dim Traffic As Boolean
		
		Traffic = LogPacketTraffic
		
		LogPacketTraffic = True
		
		For I = 0 To m_cache_count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object pkt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pkt = m_cache(I)
			
			WritePacketData(pkt.PKT_Type, pkt.Direction, pkt.ID, pkt.Length, pkt.Data, pkt.TimeDate)
		Next I
		
		LogPacketTraffic = Traffic
		
	End Sub
	
	' Written 2007-06-08 to produce packet logs or do other things
    Public Sub WritePacketData(ByVal Server As modEnum.enuPL_ServerTypes, ByVal Direction As modEnum.enuPL_DirectionTypes, ByVal PacketID As Integer, ByVal PacketLen As Integer, ByRef PacketData() As Byte, Optional ByVal TimeDate As Date = #12:00:00 AM#)

        Dim serverType As String
        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim str_Renamed As String

        Select Case (Server)
            Case modEnum.enuPL_ServerTypes.stBNCS : serverType = "BNCS"
            Case modEnum.enuPL_ServerTypes.stBNLS : serverType = "BNLS"
            Case modEnum.enuPL_ServerTypes.stMCP : serverType = "MCP"
        End Select

        If (Direction = modEnum.enuPL_DirectionTypes.StoC) Then
            str_Renamed = str_Renamed & serverType & " S -> C " & " -- Packet ID " & Right("00" & Hex(PacketID), 2) & "h (" & PacketID & "d) Length " & PacketLen & vbNewLine & vbNewLine
        Else
            str_Renamed = str_Renamed & serverType & " C -> S " & " -- Packet ID " & Right("00" & Hex(PacketID), 2) & "h (" & PacketID & "d) Length " & PacketLen & vbNewLine & vbNewLine
        End If

        str_Renamed = str_Renamed & DebugOutput(PacketData) & vbNewLine

        g_Logger.WriteSckData(str_Renamed, TimeDate)
    End Sub
	
	Public Function DWordToString(ByVal Data As Integer) As String
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As New VB6.FixedLengthString(4)
		
		Call CopyMemory(str_Renamed.Value, Data, 4)
	End Function
	
	Public Function WordToString(ByVal Data As Short) As String
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As New VB6.FixedLengthString(2)
		
		Call CopyMemory(str_Renamed.Value, Data, 2)
		
		WordToString = str_Renamed.Value
	End Function
	
	Public Function StringToDWord(ByVal Data As String) As Integer
		Dim DWord As Integer
		
		Call CopyMemory(DWord, Data, 4)
		
		StringToDWord = DWord
	End Function
	
	Public Function StringToWord(ByVal Data As String) As Integer
		Dim Word As Integer
		
		Call CopyMemory(Word, Data, 2)
		
		StringToWord = Word
	End Function
End Module