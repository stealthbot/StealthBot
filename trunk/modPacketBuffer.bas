Attribute VB_Name = "modPacketBuffer"
' modPacketBuffer.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Private Const MAX_PACKET_CACHE_SIZE = 100 ' ...

' ...
Private Type PACKETCACHEITEM
    Direction As enuPL_DirectionTypes
    PKT_Type  As enuPL_ServerTypes
    id        As Byte
    Length    As Integer
    Data      As String
    DateTime  As Date
End Type

' ...
Public Enum STRINGENCODING
    ANSI = 1
    UTF8 = 2
    UTF16 = 3
End Enum

Private m_cache()     As PACKETCACHEITEM  ' ...
Private m_cache_count As Integer          ' ...

Public Function CachePacket(Direction As enuPL_DirectionTypes, PKT_Type As enuPL_ServerTypes, id As Byte, Length As Integer, Data As String)

    Dim pkt As PACKETCACHEITEM ' ...
    
    ' ...
    With pkt
        .Direction = Direction
        .PKT_Type = PKT_Type
        .id = id
        .Length = Length
        .Data = Data
        .DateTime = Now
    End With
    
    ' ...
    If (m_cache_count + 1 >= MAX_PACKET_CACHE_SIZE) Then
        Dim I As Integer ' ...
        
        ' ...
        For I = 0 To m_cache_count - 1
            m_cache(I) = m_cache(I + 1)
        Next I
        
        ' ...
        m_cache(m_cache_count) = pkt
    Else
        ' ...
        If (m_cache_count = 0) Then
            ReDim m_cache(0)
        Else
            ReDim Preserve m_cache(0 To m_cache_count + 1)
        End If
        
        ' ...
        m_cache(m_cache_count) = pkt
        
        ' ...
        m_cache_count = m_cache_count + 1
    End If

End Function

' ...
Public Sub DumpPacketCache()
    
    Dim pkt     As PACKETCACHEITEM ' ...
    Dim I       As Integer ' ...
    Dim Traffic As Boolean ' ...
    
    ' ...
    Traffic = LogPacketTraffic
    
    ' ...
    LogPacketTraffic = True
    
    ' ...
    For I = 0 To m_cache_count - 1
        ' ...
        pkt = m_cache(I)
        
        ' ...
        WritePacketData pkt.PKT_Type, pkt.Direction, pkt.id, pkt.Length, pkt.Data, pkt.DateTime
    Next I
    
    ' ...
    LogPacketTraffic = Traffic
    
End Sub

' Written 2007-06-08 to produce packet logs or do other things
Public Sub WritePacketData(ByVal Server As enuPL_ServerTypes, ByVal Direction As enuPL_DirectionTypes, ByVal PacketID As Long, ByVal PacketLen As Long, ByRef PacketData As String, Optional ByVal DateTime As Date)

    Dim serverType As String ' ...
    Dim str        As String ' ...

    ' ...
    Select Case (Server)
        Case stBNCS: serverType = "BNCS"
        Case stBNLS: serverType = "BNLS"
        Case stMCP:  serverType = "MCP"
    End Select
    
    ' ...
    If (Direction = StoC) Then
        str = str & _
            serverType & " S -> C " & " -- Packet ID " & Right$("00" & Hex(PacketID), _
                2) & "h (" & PacketID & "d) Length " & PacketLen & _
                    vbNewLine & vbNewLine
    Else
        str = str & _
            serverType & " C -> S " & " -- Packet ID " & Right$("00" & Hex(PacketID), _
                2) & "h (" & PacketID & "d) Length " & PacketLen & _
                    vbNewLine & vbNewLine
    End If
    
    str = str & DebugOutput(PacketData) & _
            vbNewLine
    
    g_Logger.WriteSckData str
    
End Sub

' ...
Public Function DWordToString(ByVal Data As Long) As String
    Dim str As String * 4 ' ...

    ' ...
    Call CopyMemory(ByVal str, Data, 4)
End Function

' ...
Public Function WordToString(ByVal Data As Integer) As String
    Dim str As String * 2 ' ...

    ' ...
    Call CopyMemory(ByVal str, Data, 2)
    
    ' ...
    WordToString = str
End Function

' ...
Public Function StringToDWord(ByVal Data As String) As Long
    Dim DWord As Long ' ...
    
    ' ...
    Call CopyMemory(DWord, ByVal Data, 4)
    
    ' ...
    StringToDWord = DWord
End Function

' ...
Public Function StringToWord(ByVal Data As String) As Long
    Dim Word As Long ' ...

    ' ...
    Call CopyMemory(Word, ByVal Data, 2)
    
    ' ...
    StringToWord = Word
End Function
