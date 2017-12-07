Attribute VB_Name = "modPacketBuffer"
' modDataBuffer.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Private Const MAX_PACKET_CACHE_SIZE = 100

Public Enum enuServerTypes
    stGEN = 0
    stBNCS = 1
    stBNLS = 2
    stMCP = 3
    stBNFTP = 4
    stPROXY = 5
End Enum

Public Enum enuPacketHeaderTypes
    phtNONE = 0
    phtBNCS = 4
    phtMCP = 3
End Enum

Public Enum enuDirectionTypes
    CtoS = 1
    StoC = 2
End Enum

Private Type PACKETCACHEITEM
    Data()    As Byte
    PktLen    As Long
    HasPktID  As Boolean
    PktID     As Byte
    PktType   As enuServerTypes
    Direction As enuDirectionTypes
    TimeDate  As Date
End Type

Public Enum STRINGENCODING
    ANSI = 1
    UTF8 = 2
    UTF16 = 3
End Enum

Private m_cache()     As PACKETCACHEITEM
Private m_cache_count As Integer

Private Function MakePacket(ByRef Data() As Byte, ByVal PktLen As Long, _
        ByVal HasPktID As Boolean, ByVal PktID As Byte, ByVal PktType As enuServerTypes, ByVal Direction As enuDirectionTypes, _
        Optional ByVal TimeDate As Date) As PACKETCACHEITEM
    If IsMissing(TimeDate) Then TimeDate = Now

    With MakePacket
        .Data = Data
        .PktLen = PktLen
        .HasPktID = HasPktID
        .PktID = PktID
        .PktType = PktType
        .Direction = Direction
        .TimeDate = TimeDate
    End With

End Function

Public Function NamePacketType(ByVal PktType As enuServerTypes) As String
    Select Case PktType
        Case stGEN:   NamePacketType = "SCRIPTING"
        Case stBNCS:  NamePacketType = "BNCS"
        Case stBNLS:  NamePacketType = "BNLS"
        Case stMCP:   NamePacketType = "MCP"
        Case stBNFTP: NamePacketType = "BNFTP"
        Case stPROXY: NamePacketType = "PROXY"
    End Select
End Function

Private Sub CachePacket(ByRef Pkt As PACKETCACHEITEM)

    If (m_cache_count + 1 > MAX_PACKET_CACHE_SIZE) Then
        Dim i As Integer

        For i = 0 To m_cache_count - 2
            m_cache(i) = m_cache(i + 1)
        Next i

        m_cache(m_cache_count - 1) = Pkt
    Else
        If (m_cache_count = 0) Then
            ReDim m_cache(0)
        Else
            ReDim Preserve m_cache(0 To m_cache_count + 1)
        End If

        m_cache(m_cache_count) = Pkt

        m_cache_count = m_cache_count + 1
    End If

End Sub

Public Function DumpPacketCache() As Integer

    Dim Pkt     As PACKETCACHEITEM
    Dim i       As Integer
    Dim Traffic As Boolean

    Traffic = LogPacketTraffic

    LogPacketTraffic = True

    For i = 0 To m_cache_count - 1
        Pkt = m_cache(i)

        Call WritePacketData(Pkt)
    Next i

    LogPacketTraffic = Traffic

    DumpPacketCache = m_cache_count

End Function

' Written 2007-06-08 to produce packet logs or do other things
Private Sub WritePacketData(ByRef Pkt As PACKETCACHEITEM)

    Dim sDir        As String
    Dim sPacketID   As String
    Dim sPacketLen  As String
    Dim sID         As String
    Dim str         As String
    
    sID = NamePacketType(Pkt.PktType)
    
    Select Case Pkt.Direction
        Case CtoS: sDir = "C -> S"
        Case StoC: sDir = "S -> C"
    End Select
    
    sPacketID = vbNullString
    If Pkt.HasPktID Then
        sPacketID = StringFormat("ID 0x{0} -- ", ZeroOffset(Pkt.PktID, 2))
    End If
    
    sPacketLen = StringFormat("Length {0} b", Pkt.PktLen)
    
    str = StringFormat("{0} {1} -- {2}{3}{4}{5}{6}", _
            sID, sDir, sPacketID, sPacketLen, vbNewLine, _
            DebugOutput(Pkt.Data), vbNewLine)
    
    g_Logger.WriteSckData str, Pkt.TimeDate
End Sub

Public Function DebugOutput(ByVal Data As Variant, Optional ByVal Start As Long = 0, Optional ByVal Length As Long = -1) As String

    Dim Buffer() As Byte
    Dim x1 As Long, y1 As Long
    Dim iLen As Long, iPos As Long
    Dim sHex As String, sRaw As String, c As Byte
    Dim sOut As String
    Dim offset As Long, sOffset As String
    Dim Brk As Integer

    If VarType(Data) = vbString Then
        Buffer = StringToByteArr(Data)
    ElseIf VarType(Data) = vbArray + vbByte Then
        Buffer = Data
    Else
        Exit Function
    End If

    If LBound(Buffer) > UBound(Buffer) Then Exit Function

    'build random string to display
    '    y1 = 256
    '    sIn = String(y1, 0)
    '    For x1 = 1 To 256
    '        Mid(sIn, x1, 1) = Chr(x1 - 1)
    '        Mid(sIn, x1, 1) = Chr(255 * Rnd())
    '    Next x1
    iLen = UBound(Buffer) + 1 - Start

    If Length >= 0 Then
        iLen = IIf(Length > iLen - Start, iLen - Start, Length)
    Else
        iLen = iLen - Start
    End If
    If iLen <= 0 Then Exit Function

    sOut = vbNullString
    offset = 0

    For x1 = 0 To ((iLen - 1) \ 16)
        sOffset = ZeroOffset(offset, 4)
        sHex = String$(49, " ")
        sRaw = "........ ........"
        For y1 = 1 To 16
            iPos = 16 * x1 + y1
            Brk = Abs(CInt(y1 > 8))
            If iPos > iLen Then
                Mid$(sRaw, y1 + Brk) = String$(17 - y1 - Brk + 1, " ")
                Exit For
            End If

            c = Buffer(iPos - 1 + Start)
            Mid$(sHex, 3 * (y1 - 1) + 1 + Brk, 2) = ZeroOffset(c, 2) & " "
            Select Case c
                Case Is < 32, Is >= 127
                Case Else
                    Mid$(sRaw, y1 + Brk, 1) = ChrW$(c)
            End Select
        Next y1
        If LenB(sOut) > 0 Then sOut = sOut & vbCrLf
        sOut = StringFormat("{0}{1}: {2} |{3}|", sOut, sOffset, sHex, sRaw)
        offset = offset + 16
    Next x1

    'sDebugBuf = sDebugBuf & vbCrLf & vbCrLf & sOut
    DebugOutput = sOut
End Function

' SendData() function, returns true if not vetoed and there was data to send; handles saving/logging
' arguments:
' Data(): buffer
' DataLen: length of Data
' HasPktID: whether a PktID parameter should be shown in packet logs
' PktID: the packet ID value
' SocketType: which socket to send on (if not valid or not connected, send fails)
' PacketType: value sent to NamePacketType() shown in packet logs
' HeaderType: what kind of header to prepend
Public Function SendData(ByRef Data() As Byte, ByVal DataLen As Long, _
        ByVal HasPktID As Boolean, Optional ByVal PktID As Byte, Optional ByVal Socket As Winsock, _
        Optional ByVal PktType As enuServerTypes, Optional ByVal HeaderType As enuPacketHeaderTypes) As Boolean
    Dim buf()    As Byte
    Dim HLen     As Byte
    Dim PktLen   As Long
    Dim sID      As String
    Dim sData    As String
    Dim Pkt      As PACKETCACHEITEM

    SendData = False

    If Socket Is Nothing Then Exit Function

    HLen = CByte(HeaderType)
    sID = NamePacketType(PktType)

    If (Socket.State <> sckConnected) Then
        ' not connected
        Exit Function
    End If

    PktLen = DataLen + HLen

    If PktLen <= 0 Then
        ' no data
        Exit Function
    End If

    ' resize temporary data buffer
    ReDim buf(PktLen - 1)

    ' copy packet data Length to temporary buffer
    Select Case HeaderType
        Case phtBNCS:
            buf(0) = &HFF                 ' (BYTE) 0xFF
            buf(1) = PktID                ' (BYTE) ID
            CopyMemory buf(2), PktLen, 2  ' (WORD) Length
        Case phtMCP
            CopyMemory buf(0), PktLen, 2  ' (WORD) Length
            buf(2) = PktID                ' (BYTE) ID
        Case Else
            ' nop
    End Select

    ' copy data from buffer to temporary buffer
    If (DataLen > 0) Then
        CopyMemory buf(HLen), Data(0), DataLen
    End If

    sData = ByteArrToString(buf)

    SendData = Not RunInAll("Event_PacketSent", sID, PktID, PktLen, sData)
    If SendData Then
        If (MDebug("all")) Then
            frmChat.AddChat COLOR_BLUE, StringFormat("{0} RECV 0x{1}", sID, ZeroOffset(PktID, 2))
        End If

        Socket.SendData buf

        ' only log if sent
        If PktType <> stGEN Then
            Pkt = MakePacket(buf, PktLen, HasPktID, PktID, PktType, CtoS)
            Call CachePacket(Pkt)
            Call WritePacketData(Pkt)
        End If
    End If
End Function

' HandleRecvData() function, returns true if not vetoed and there was data to recv; handles saving/logging
' arguments:
' Data(): buffer
' DataLen: length of Data
' HasPktID: whether a PktID parameter should be shown in packet logs
' PktID: the packet ID value
' PacketType: value sent to NamePacketType() shown in packet logs
' HeaderType: what kind of header this packet has
' ScriptSource: True if this was the result of SSC.ForcePacketParse()
Public Function HandleRecvData(ByRef Data() As Byte, ByVal DataLen As Long, ByVal HasPktID As Boolean, ByVal PktID As Byte, _
        Optional ByVal PktType As enuServerTypes, Optional ByVal HeaderType As enuPacketHeaderTypes, Optional ByVal ScriptSource As Boolean = False) As Boolean
    Dim buf() As Byte
    Dim sID   As String
    Dim sData As String
    Dim Pkt   As PACKETCACHEITEM

    HandleRecvData = False

    If DataLen = 0 Then Exit Function
    If LBound(Data) > UBound(Data) Then Exit Function

    ReDim buf(0 To DataLen - 1)
    CopyMemory buf(0), Data(0), DataLen

    sID = NamePacketType(PktType)
    If (MDebug("all")) Then
        frmChat.AddChat COLOR_BLUE, StringFormat("{0} RECV 0x{1}", sID, ZeroOffset(PktID, 2))
    End If

    sData = ByteArrToString(buf)

    If ScriptSource Then
        ' source is SSC.ForcePacketParse(), packet is going to be parsed as-is
        HandleRecvData = True
    Else
        ' source is socket, log then SSC event for vetoes
        Pkt = MakePacket(buf, DataLen, HasPktID, PktID, PktType, StoC)
        Call CachePacket(Pkt)
        Call WritePacketData(Pkt)

        HandleRecvData = Not RunInAll("Event_PacketReceived", sID, PktID, DataLen, sData)
    End If

    If HandleRecvData Then
        ' packet is going to be parsed
        Call RunInAll("Event_PacketParsed", sID, PktID, DataLen, sData)
    End If
End Function
