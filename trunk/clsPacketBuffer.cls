VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' PacketBuffer.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_buf()  As Byte
Private m_bufpos As Integer

' ...
Private Sub Class_Initialize()
    ' clear buffer contents
    Call Clear
End Sub ' end function Class_Initialize

' ...
Private Sub Class_Terminate()
    ' clear buffer contents
    Call Clear
End Sub ' end function Class_Terminate

' ...
Public Function InsertByte(ByVal Data As Byte)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufpos + 1))

    ' copy data to buffer
    m_buf(m_bufpos) = Data
    
    ' store buffer length
    m_bufpos = (m_bufpos + 1)
End Function ' end function InsertByte

' ...
Public Function InsertByteArr(ByRef Data() As Byte)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufpos + (UBound(Data) + 1)))

    ' copy data to buffer
    Call CopyMemory(m_buf(m_bufpos), Data(0), (UBound(Data) + 1))

    ' store buffer length
    m_bufpos = (m_bufpos + (UBound(Data) + 1))
End Function ' end function InsertByte

' ...
Public Function GetByte() As Byte

End Function ' end function GetByte

' ...
Public Function InsertWord(ByVal Data As Integer)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufpos + 2))

    ' copy data to buffer
    Call CopyMemory(m_buf(m_bufpos), Data, 2)
    
    ' store buffer length
    m_bufpos = (m_bufpos + 2)
End Function ' end function InsertWord

' ...
Public Function GetWord() As Integer

End Function ' end function GetWord

' ...
Public Function InsertDWord(ByVal Data As Long)
    ' resize data buffer
    ReDim Preserve m_buf(0 To (m_bufpos + 4))

    ' copy data to buffer
    Call CopyMemory(m_buf(m_bufpos), Data, 4)
    
    ' store buffer length
    m_bufpos = (m_bufpos + 4)
End Function ' end function InsertDWord

' ...
Public Function GetDWord() As Long

End Function ' end function GetDWord

' ...
Public Function InsertNonNTString(ByVal Data As String)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufpos + Len(Data)))
    
    ' copy data to buffer
    Call CopyMemory(m_buf(m_bufpos), ByVal Data, Len(Data))
    
    ' store buffer length
    m_bufpos = (m_bufpos + Len(Data))
End Function ' end function InsertNonNTString

' ...
Public Function InsertNTString(ByRef Data As String, _
    Optional ByVal Encoding As STRINGENCODING = STRINGENCODING.ANSI)
    
    Dim arrStr() As Byte ' ...
    
    ' ...
    Select Case (Encoding)
        Case STRINGENCODING.ANSI
            arrStr() = StrConv(Data, vbFromUnicode)
            
        Case STRINGENCODING.UTF8
            arrStr() = modUTF8.UTF8Encode(Data)
            
        Case STRINGENCODING.UTF16
            arrStr() = StrConv(Data, vbUnicode)
    End Select
    
    ' resize buffer and include terminating null character
    ReDim Preserve m_buf(0 To (m_bufpos + (UBound(arrStr) + 2)))
    
    ' ...
    If (Data <> vbNullString) Then
        ' copy data to buffer
        Call CopyMemory(m_buf(m_bufpos), arrStr(0), (UBound(arrStr) + 1))
    End If

    ' store buffer length
    m_bufpos = (m_bufpos + (UBound(arrStr) + 2))
End Function ' end function InsertNTString

' ...
Public Function GetString() As String

End Function ' end function GetString

' ...
Public Function Length() As Long
    Length = m_bufpos
End Function

' ...
Public Function Clear()
    ' resize buffer
    ReDim m_buf(0)
    
    ' clear first index
    m_buf(0) = 0
    
    ' reset buffer length
    m_bufpos = 0
End Function ' end function Clear

' ...
Public Function SendPacketMCP(Optional PacketID As Byte)
    Dim buf() As Byte ' ...
    
    ' ...
    If (frmChat.sckMCP.State <> 0) Then
        ' resize temporary data buffer
        ReDim buf(m_bufpos + 2)

         ' copy packet data length to temporary buffer
        Call CopyMemory(buf(0), m_bufpos + 3, 2)
        
        buf(2) = PacketID ' packet identification number
    
        ' ...
        If (m_bufpos) Then
            ' copy data from buffer to temporary buffer
            Call CopyMemory(buf(3), m_buf(0), m_bufpos)
        End If
        
        ' ...
        If (MDebug("all")) Then
            frmChat.AddChat COLOR_BLUE, "MCP SEND 0x" & _
                Right$("00" & Hex(PacketID), 2)
        End If
        
        ' ...
        Call frmChat.sckMCP.SendData(buf)
    End If
    
    ' clear buffer contents
    Call Clear
End Function ' end function SendPacketMCP

' ...
Public Function SendPacket(Optional PacketID As Byte)
    Dim l     As Long    ' ...
    Dim buf() As Byte    ' ...
    Dim i     As Integer ' ...
    Dim log   As String  ' ...
    
    ' resize temporary data buffer
    ReDim buf(m_bufpos + 3)
    
    ' ...
    If (MDebug("all")) Then
        frmChat.AddChat COLOR_BLUE, "BNET SEND 0x" & Right$("00" & _
            Hex(PacketID), 2)
    End If
    
    ' ...
    If (frmChat.sckBNet.State <> 0) Then
        buf(0) = &HFF     ' header
        buf(1) = PacketID ' packet identification number
        
        ' copy packet data length to temporary buffer
        Call CopyMemory(buf(2), m_bufpos + 4, 2)
        
        ' ...
        If (m_bufpos) Then
            ' copy data from buffer to temporary buffer
            Call CopyMemory(buf(4), m_buf(0), m_bufpos)
        End If
        
        ' ...
        Call frmChat.sckBNet.SendData(buf)
    End If
    
    ' ...
    log = String$(m_bufpos + 4, vbNullChar)
    
    ' ...
    Call CopyMemory(ByVal log, buf(0), m_bufpos + 4)
    
    ' ...
    Call LogPacketRaw(stBNCS, CtoS, PacketID, m_bufpos + 4, log)
    
    ' clear buffer contents
    Call Clear
End Function ' end function SendPacket

' ...
Public Function vLSendPacket(Optional PacketID As Byte)
    Dim buf() As Byte   ' ...
    Dim log   As String ' ...
    
    ' ...
    If (MDebug("all")) Then
        frmChat.AddChat COLOR_BLUE, "BNLS SEND 0x" & _
            Right$("00" & Hex(PacketID), 2)
    End If
    
    ' ...
    If (frmChat.sckBNLS.State <> 0) Then
        ' resize temporary data buffer
        ReDim buf(m_bufpos + 2)
    
        ' copy packet data length to temporary buffer
        Call CopyMemory(buf(0), m_bufpos + 3, 2)

        buf(2) = PacketID ' packet identification number
        
        ' ...
        If (m_bufpos) Then
            ' copy data from buffer to temporary buffer
            Call CopyMemory(buf(3), m_buf(0), m_bufpos)
        End If
        
        ' ...
        Call frmChat.sckBNLS.SendData(buf)
    End If
    
    ' ...
    log = String$(m_bufpos + 3, vbNullChar)
    
    ' ...
    Call CopyMemory(ByVal log, buf(0), m_bufpos + 3)
    
    ' ...
    Call LogPacketRaw(stBNLS, CtoS, PacketID, m_bufpos + 3, log)
    
    ' clear buffer contents
    Call Clear
End Function ' end function vLSendPacket

' ...
Public Function GetPacket(PacketID As Byte) As String
    Dim buf As String ' ...
    
    ' ...
    buf = String$(m_bufpos + 1, vbNull)
    
    ' ...
    Call CopyMemory(ByVal buf, m_buf(0), m_bufpos + 1)

    ' ...
    GetPacket = Chr$(&HFF) & Chr$(PacketID) & WordToString((m_bufpos + 1) + 4) & _
        buf
    
    ' clear buffer contents
    Call Clear
End Function ' end function GetPacket

' ...
Public Function GetBNLSPacket(PacketID As Byte) As String
    Dim buf As String ' ...
    
    ' ...
    buf = String$(m_bufpos + 1, vbNull)
    
    ' ...
    Call CopyMemory(ByVal buf, m_buf(0), m_bufpos + 1)

    ' ...
    GetBNLSPacket = WordToString((m_bufpos + 1) + 3) & Chr$(PacketID) & _
        buf
    
    ' clear buffer contents
    Call Clear
End Function ' end function GetBNLSPacket

