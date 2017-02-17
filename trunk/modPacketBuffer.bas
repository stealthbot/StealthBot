Attribute VB_Name = "modPacketBuffer"
' modDataBuffer.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Private Const MAX_PACKET_CACHE_SIZE = 100

Private Type PACKETCACHEITEM
    Direction As enuPL_DirectionTypes
    PKT_Type  As enuPL_ServerTypes
    ID        As Byte
    length    As Integer
    Data()    As Byte
    TimeDate  As Date
End Type

Public Enum STRINGENCODING
    ANSI = 1
    UTF8 = 2
    UTF16 = 3
End Enum

Private m_cache()     As PACKETCACHEITEM
Private m_cache_count As Integer

Public Function CachePacket(ByVal PKT_Type As enuPL_ServerTypes, ByVal Direction As enuPL_DirectionTypes, ByVal ID As Byte, ByVal length As Integer, ByRef Data() As Byte)

    Dim pkt As PACKETCACHEITEM
    
    With pkt
        .Direction = Direction
        .PKT_Type = PKT_Type
        .ID = ID
        .length = length
        .Data = Data
        .TimeDate = Now
    End With
    
    If (m_cache_count + 1 >= MAX_PACKET_CACHE_SIZE) Then
        Dim i As Integer
        
        For i = 0 To m_cache_count - 1
            m_cache(i) = m_cache(i + 1)
        Next i
        
        m_cache(m_cache_count) = pkt
    Else
        If (m_cache_count = 0) Then
            ReDim m_cache(0)
        Else
            ReDim Preserve m_cache(0 To m_cache_count + 1)
        End If
        
        m_cache(m_cache_count) = pkt
        
        m_cache_count = m_cache_count + 1
    End If

End Function

Public Sub DumpPacketCache()
    
    Dim pkt     As PACKETCACHEITEM
    Dim i       As Integer
    Dim Traffic As Boolean
    
    Traffic = LogPacketTraffic
    
    LogPacketTraffic = True
    
    For i = 0 To m_cache_count - 1
        pkt = m_cache(i)
        
        WritePacketData pkt.PKT_Type, pkt.Direction, pkt.ID, pkt.length, pkt.Data, _
            pkt.TimeDate
    Next i
    
    LogPacketTraffic = Traffic
    
End Sub

' Written 2007-06-08 to produce packet logs or do other things
Public Sub WritePacketData(ByVal PKT_Type As enuPL_ServerTypes, ByVal Direction As enuPL_DirectionTypes, ByVal PacketID As Long, ByVal PacketLen As Long, ByRef PacketBuffer() As Byte, Optional ByVal TimeDate As Date)

    Dim sServerType As String
    Dim sDir        As String
    Dim sPacketID   As String
    Dim sPacketLen  As String
    Dim str         As String

    Select Case (PKT_Type)
        Case stBNCS:  sServerType = "BNCS"
        Case stBNLS:  sServerType = "BNLS"
        Case stMCP:   sServerType = "REALM"
        Case stPROXY: sServerType = "PROXY"
    End Select
    
    Select Case Direction
        Case CtoS: sDir = "C -> S"
        Case StoC: sDir = "S -> C"
    End Select
    
    sPacketID = vbNullString
    If PKT_Type <> stPROXY Then
        sPacketID = StringFormat("ID 0x{0} -- ", ZeroOffset(PacketID, 2))
    End If
    
    sPacketLen = StringFormat("Length {0} b", PacketLen)
    
    str = StringFormat("{0} {1} -- {2}{3}{4}{5}{6}", _
            sServerType, sDir, sPacketID, sPacketLen, vbNewLine, _
            DebugOutputBuffer(PacketBuffer), vbNewLine)
    
    g_Logger.WriteSckData str, TimeDate
End Sub

Public Function DWordToString(ByVal Data As Long) As String
    Dim str As String * 4

    Call CopyMemory(ByVal str, Data, 4)
End Function

Public Function WordToString(ByVal Data As Integer) As String
    Dim str As String * 2

    Call CopyMemory(ByVal str, Data, 2)
    
    WordToString = str
End Function

Public Function StringToDWord(ByVal Data As String) As Long
    Dim DWord As Long
    
    Call CopyMemory(DWord, ByVal Data, 4)
    
    StringToDWord = DWord
End Function

Public Function StringToWord(ByVal Data As String) As Long
    Dim Word As Long

    Call CopyMemory(Word, ByVal Data, 2)
    
    StringToWord = Word
End Function

Public Function DebugOutput(ByVal sIn As String, Optional ByVal Start As Long = 1, Optional ByVal length As Long = -1) As String

    Dim x1 As Long, y1 As Long
    Dim iLen As Long, iPos As Long
    Dim sB As String, st As String, c As String
    Dim sOut As String
    Dim offset As Long, sOffset As String
    'build random string to display
    '    y1 = 256
    '    sIn = String(y1, 0)
    '    For x1 = 1 To 256
    '        Mid(sIn, x1, 1) = Chr(x1 - 1)
    '        Mid(sIn, x1, 1) = Chr(255 * Rnd())
    '    Next x1
    If length >= 0 Then
        sIn = Mid$(sIn, Start, length)
    Else
        sIn = Mid$(sIn, Start)
    End If
    
    iLen = Len(sIn)

    If iLen = 0 Then Exit Function
    sOut = vbNullString
    offset = 0

    For x1 = 0 To ((iLen - 1) \ 16)
        sOffset = Right$("0000" & Hex$(offset), 4)
        sB = String$(48, " ")
        st = "................"
        For y1 = 1 To 16
            iPos = 16 * x1 + y1
            If iPos > iLen Then Exit For

            c = Mid$(sIn, iPos, 1)
            Mid$(sB, 3 * (y1 - 1) + 1, 2) = Right$("00" & Hex$(Asc(c)), 2) & " "
            Select Case Asc(c)
                Case 0, 9, 10, 13
                Case Else
                    Mid$(st, y1, 1) = c
            End Select
        Next y1
        If LenB(sOut) > 0 Then sOut = sOut & vbCrLf
        sOut = sOut & sOffset & ":  "
        sOut = sOut & sB & "  " & st
        offset = offset + 16
    Next x1

    'sDebugBuf = sDebugBuf & vbCrLf & vbCrLf & sOut
    DebugOutput = sOut
End Function

Public Function DebugOutputBuffer(ByRef Data() As Byte, Optional ByVal Start As Long = 0, Optional ByVal length As Long = -1) As String
    Dim sData As String
    sData = StrConv(Data(), vbUnicode, 1033)
    DebugOutputBuffer = DebugOutput(sData, Start + 1, length)
End Function

