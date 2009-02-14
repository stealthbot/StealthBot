Attribute VB_Name = "modPacketBuffer"
' modPacketBuffer.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Private Const MAX_PACKET_CACHE_SIZE = 100 ' ...

' ...
Private Type PACKETCACHEITEM
    Direction As enuPL_DirectionTypes
    PKT_Type  As enuPL_ServerTypes
    ID        As Byte
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

Public Function CachePacket(Direction As enuPL_DirectionTypes, PKT_Type As enuPL_ServerTypes, ID As Byte, Length As Integer, Data As String)

    Dim pkt As PACKETCACHEITEM ' ...
    
    ' ...
    With pkt
        .Direction = Direction
        .PKT_Type = PKT_Type
        .ID = ID
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
        LogPacketRaw pkt.PKT_Type, pkt.Direction, pkt.ID, pkt.Length, pkt.Data, pkt.DateTime
    Next I
    
    ' ...
    LogPacketTraffic = Traffic
    
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
