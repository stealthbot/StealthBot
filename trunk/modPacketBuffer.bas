Attribute VB_Name = "modPacketBuffer"
' modPacketBuffer.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Public Const MAX_PACKET_CACHE_SIZE = 30 ' ...

Public pkt As PACKETCACHEITEM ' ...

Private m_cache As Collection ' ...

' ...
Public Type PACKETCACHEITEM
    Type As enuPL_ServerTypes
    ID   As Integer
    Len  As Integer
    Data As String
End Type

' ...
Public Enum STRINGENCODING
    ANSI = 1
    UTF8 = 2
    UTF16 = 3
End Enum

' ...
Public Function PacketCache() As Collection

    ' ...
    If (m_cache Is Nothing) Then
        Set m_cache = New Collection
    End If
    
    ' ...
    Set PacketCache = m_cache

End Function

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
