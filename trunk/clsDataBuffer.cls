VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' clsDataBuffer.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_buf()    As Byte
Private m_bufpos   As Integer
Private m_bufsize  As Integer
Private m_cripple  As Boolean

' ...
Private Sub Class_Initialize()

    ' clear buffer contents
    Clear
    
End Sub ' end function Class_Initialize

' ...
Private Sub Class_Terminate()

    ' clear buffer contents
    Clear
    
End Sub ' end function Class_Terminate

Public Sub setCripple()

    m_cripple = True

End Sub

Public Function getCripple() As Boolean

    getCripple = m_cripple

End Function

Public Property Let Data(ByVal str As String)

    ' ...
    ReDim m_buf(0 To Len(str))

    ' ...
    CopyMemory m_buf(0), ByVal str, Len(str)

    ' ...
    m_bufsize = Len(str)

End Property

Public Property Get Data() As String

    ' ...
    Data = String(m_bufsize, Chr$(0))

    ' ...
    CopyMemory ByVal Data, m_buf(0), m_bufsize

End Property

' ...
Public Function InsertByte(ByVal Data As Byte)

    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 1))

    ' copy data to buffer
    m_buf(m_bufsize) = Data
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 1)
    
End Function ' end function InsertByte

' ...
Public Function InsertByteArr(ByRef Data() As Byte)

    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + (UBound(Data) + 1)))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data(0), (UBound(Data) + 1)

    ' store buffer Length
    m_bufsize = (m_bufsize + (UBound(Data) + 1))
    
End Function ' end function InsertByte

' ...
Public Sub GetByteArr(ByRef Data() As Byte, Optional ByVal length As Integer = -1, Optional Peek As Boolean = False)

    ' ...
    If (length = -1) Then
        length = m_bufsize
    End If
    
    ' ...
    If ((m_bufpos + length) > m_bufsize) Then
        Exit Sub
    End If
    
    ' resize buffer
    ReDim Data(0 To length)

    ' copy data to buffer
    CopyMemory Data(0), m_buf(m_bufpos), length

    ' store buffer Length
    If (Not Peek) Then m_bufpos = (m_bufpos + length)
    
End Sub      ' end function InsertByte

' ...
Public Function GetByte(Optional Peek As Boolean = False) As Byte

    ' ...
    If ((m_bufpos + 1) > m_bufsize) Then
        Exit Function
    End If

    ' ...
    GetByte = m_buf(m_bufpos)

    ' ...
    If (Not Peek) Then m_bufpos = (m_bufpos + 1)

End Function ' end function GetByte

' ...
Public Function InsertWord(ByVal Data As Integer)

    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 2))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data, 2
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 2)
    
End Function ' end function InsertWord

' ...
Public Function GetWord(Optional Peek As Boolean = False) As Integer

    ' ...
    If ((m_bufpos + 2) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory GetWord, m_buf(m_bufpos), 2

    ' ...
    If (Not Peek) Then m_bufpos = (m_bufpos + 2)

End Function ' end function GetWord

' ...
Public Function InsertDWord(ByVal Data As Long)

    ' resize data buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 4))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data, 4
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 4)
    
End Function ' end function InsertDWord

' ...
Public Function GetDWORD(Optional Peek As Boolean = False) As Long

    ' ...
    If ((m_bufpos + 4) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory GetDWORD, m_buf(m_bufpos), 4

    ' ...
    If (Not Peek) Then m_bufpos = (m_bufpos + 4)

End Function ' end function GetDWord

' ...
Public Function GetFileTime(Optional Peek As Boolean = False) As Date

    Dim ft As FILETIME ' ...

    ' ...
    If ((m_bufpos + 8) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory ft, m_buf(m_bufpos), 8

    ' ...
    If (Not Peek) Then m_bufpos = (m_bufpos + 8)

    ' ...
    GetFileTime = FileTimeToDate(ft)

End Function ' end function GetFileTime

' ...
Public Function InsertNonNTString(ByVal Data As String)

    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + Len(Data)))
    
    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), ByVal Data, Len(Data)
    
    ' store buffer Length
    m_bufsize = (m_bufsize + Len(Data))
    
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
    ReDim Preserve m_buf(0 To (m_bufsize + (UBound(arrStr) + 2)))
    
    ' copy data to buffer
    If (Data <> vbNullString) Then
        CopyMemory m_buf(m_bufsize), arrStr(0), (UBound(arrStr) + 1)
    End If

    ' store buffer Length
    m_bufsize = (m_bufsize + (UBound(arrStr) + 2))
    
End Function ' end function InsertNTString

' ...
Public Function GetString(Optional ByVal Encoding As STRINGENCODING = STRINGENCODING.ANSI, Optional Peek As Boolean = False) As String

    Dim I As Integer ' ...

    ' ...
    For I = m_bufpos To m_bufsize
        If (m_buf(I) = &H0) Then
            Exit For
        End If
    Next I

    ' ...
    If (I < m_bufsize) Then
        ' ...
        GetString = String$(I - m_bufpos, Chr$(0))
            
        ' copy data to buffer
        CopyMemory ByVal GetString, m_buf(m_bufpos), I - m_bufpos
        
        ' ...
        If (Not Peek) Then m_bufpos = I + 1
    End If

End Function ' end function GetString

' ...
Public Function GetRaw(Optional ByVal length As Integer = -1, Optional Peek As Boolean = False) As String

    ' ...
    If (length = -1) Then
        length = m_bufsize
    End If
    
    ' ...
    If ((m_bufpos + length) > m_bufsize) Then
        Exit Function
    End If

    ' ...
    GetRaw = String$(length, Chr$(0))
            
    ' copy data to buffer
    CopyMemory ByVal GetRaw, m_buf(m_bufpos), length
    
    ' ...
    If (Not Peek) Then m_bufpos = m_bufpos + length

End Function ' end function GetString

' ...
Public Property Get length() As Long

    length = m_bufsize
    
End Property

' ...
Public Property Get Position() As Long

    Position = m_bufpos
    
End Property

' ...
Public Property Let Position(ByVal lng As Long)

    m_bufpos = lng
    
End Property

' ...
Public Function Clear()

    ' resize buffer
    ReDim m_buf(0)
    
    ' clear first index
    m_buf(0) = 0
    
    ' reset buffer Length
    m_bufsize = 0
    
    ' reset buffer position
    m_bufpos = 0
End Function ' end function Clear

Public Function DebugOutput() As String

    DebugOutput = modWar3Clan.DebugOutput(Data)

End Function

' ...
Public Function SendPacketMCP(Optional PacketID As Byte)

    Dim buf()  As Byte   ' ...
    Dim strbuf As String ' ...
    Dim veto   As Boolean
    
    ' ...
    If (frmChat.sckMCP.State <> sckConnected) Then
        Clear
        Exit Function
    End If
    
    ' resize temporary data buffer
    ReDim buf(m_bufsize + 2)

     ' copy packet data Length to temporary buffer
    CopyMemory buf(0), m_bufsize + 3, 2
    
    ' ...
    buf(2) = PacketID ' packet identification number

    ' copy data from buffer to temporary buffer
    If (m_bufsize) Then
        CopyMemory buf(3), m_buf(0), m_bufsize
    End If
    
    ' ...
    strbuf = String$(m_bufsize + 3, vbNullChar)
        
    ' ...
    CopyMemory ByVal strbuf, buf(0), m_bufsize + 3

    ' ...
    veto = RunInAll("Event_PacketSent", "MCP", PacketID, m_bufsize + 3, strbuf)
    
    If (Not veto) Then
    
        ' ...
        If (MDebug("all")) Then
            frmChat.AddChat COLOR_BLUE, "MCP SEND 0x" & _
                Right$("00" & Hex(PacketID), 2)
        End If
    
        ' ...
        Send frmChat.sckMCP.SocketHandle, strbuf, m_bufsize + 3, 0

        ' ...
        CachePacket CtoS, stMCP, PacketID, m_bufsize + 3, strbuf
    
        ' ...
        WritePacketData stMCP, CtoS, PacketID, m_bufsize + 3, strbuf
    End If
    
    ' clear buffer contents
    Clear
    
End Function ' end function SendPacketMCP

' ...
Public Function SendPacket(Optional PacketID As Byte)

    Dim L      As Long    ' ...
    Dim buf()  As Byte    ' ...
    Dim I      As Integer ' ...
    Dim strbuf As String  ' ...
    Dim veto   As Boolean
    
    ' ...
    If (frmChat.sckBNet.State <> sckConnected) Then
        Clear
        Exit Function
    End If
    
    If (getCripple) Then
        Select Case (PacketID)
            Case &HE, &HC:
                Clear
                Exit Function
        End Select
    End If
    
    ' resize temporary data buffer
    ReDim buf(m_bufsize + 3)
    
    ' ...
    buf(0) = &HFF     ' header
    buf(1) = PacketID ' packet identification number
    
    ' copy packet data Length to temporary buffer
    CopyMemory buf(2), m_bufsize + 4, 2
    
    ' copy data from buffer to temporary buffer
    If (m_bufsize) Then
        CopyMemory buf(4), m_buf(0), m_bufsize
    End If
    
    ' ...
    strbuf = String$(m_bufsize + 4, vbNullChar)
        
    ' ...
    CopyMemory ByVal strbuf, buf(0), m_bufsize + 4
    
    ' ...
    veto = RunInAll("Event_PacketSent", "BNCS", PacketID, m_bufsize + 4, strbuf)
    
    If (Not veto) Then
    
        ' ...
        If (MDebug("all")) Then
            frmChat.AddChat COLOR_BLUE, "BNET SEND 0x" & Right$("00" & _
                Hex(PacketID), 2)
        End If
    
        ' ...
        Send frmChat.sckBNet.SocketHandle, strbuf, m_bufsize + 4, 0

        ' ...
        CachePacket CtoS, stBNCS, PacketID, m_bufsize + 4, strbuf
    
        ' ...
        WritePacketData stBNCS, CtoS, PacketID, m_bufsize + 4, strbuf
    
        'Send Warden Everything thats Sent to Bnet
        Call modWarden.WardenClientData(warden_context, strbuf)
    End If
    ' clear buffer contents
    Clear
    
End Function ' end function SendPacket

' ...
Public Function vLSendPacket(Optional PacketID As Byte)

    Dim buf()  As Byte   ' ...
    Dim strbuf As String ' ...
    Dim veto   As Boolean
    
    ' ...
    If (frmChat.sckBNLS.State <> sckConnected) Then
        Clear
        Exit Function
    End If
    
    ' resize temporary data buffer
    ReDim buf(m_bufsize + 2)

    ' copy packet data Length to temporary buffer
    CopyMemory buf(0), m_bufsize + 3, 2

    ' ...
    buf(2) = PacketID ' packet identification number
    
    ' copy data from buffer to temporary buffer
    If (m_bufsize) Then
        CopyMemory buf(3), m_buf(0), m_bufsize
    End If
    
    ' ...
    strbuf = String$(m_bufsize + 3, vbNullChar)
        
    ' ...
    CopyMemory ByVal strbuf, buf(0), m_bufsize + 3

    ' ...
    veto = RunInAll("Event_PacketSent", "BNLS", PacketID, m_bufsize + 3, strbuf)
    
    If (Not veto) Then
        ' ...
        If (MDebug("all")) Then
            frmChat.AddChat COLOR_BLUE, "BNLS SEND 0x" & _
                Right$("00" & Hex(PacketID), 2)
        End If
    
        ' ...
        Send frmChat.sckBNLS.SocketHandle, strbuf, m_bufsize + 3, 0

        ' ...
        CachePacket CtoS, stBNLS, PacketID, m_bufsize + 3, strbuf
    
        ' ...
        WritePacketData stBNLS, CtoS, PacketID, m_bufsize + 3, strbuf
    End If
    
    ' clear buffer contents
    Clear
    
End Function ' end function vLSendPacket

