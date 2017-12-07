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

Option Explicit

Private m_buf()    As Byte
Private m_bufpos   As Integer
Private m_bufsize  As Integer
Private m_cripple  As Boolean

Private Sub Class_Initialize()
    ' clear buffer contents
    Clear
End Sub

Private Sub Class_Terminate()
    ' clear buffer contents
    Clear
End Sub

Public Sub SetCripple()
    m_cripple = True
End Sub

Public Function GetCripple() As Boolean
    GetCripple = m_cripple
End Function

Public Property Let Data(ByVal str As String)
    If LenB(str) = 0 Then
        Call Clear
    Else
        Call SetDataAsByteArr(StringToByteArr(str))
    End If
End Property

Public Property Get Data() As String
    If m_bufsize = 0 Then
        Data = vbNullString
    Else
        Data = ByteArrToString(GetDataAsByteArr())
    End If
End Property

Public Sub SetDataAsByteArr(ByRef buf() As Byte)
    If UBound(buf) < LBound(buf) Then
        Clear
        Exit Sub
    End If
    
    ' resize buffer
    ReDim m_buf(0 To (UBound(buf) + 2))

    ' copy data to buffer
    CopyMemory m_buf(0), buf(0), (UBound(buf) + 1)

    ' store buffer Length
    m_bufsize = UBound(buf) + 1
End Sub

Public Function GetDataAsByteArr() As Byte()
    Dim buf() As Byte
    
    If m_bufsize = 0 Then
        GetDataAsByteArr = vbNullString
        Exit Function
    End If
    
    ReDim buf(0 To (m_bufsize - 1))
    
    CopyMemory buf(0), m_buf(0), m_bufsize
    
    GetDataAsByteArr = buf
End Function

Public Sub InsertByte(ByVal Data As Byte)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 1))

    ' copy data to buffer
    m_buf(m_bufsize) = Data
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 1)
End Sub

Public Sub InsertByteArr(ByRef Data() As Byte)
    If UBound(Data) < LBound(Data) Then
        Exit Sub
    End If
    
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + (UBound(Data) + 2)))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data(0), (UBound(Data) + 1)

    ' store buffer Length
    m_bufsize = (m_bufsize + (UBound(Data) + 1))
End Sub

Public Sub GetByteArr(ByRef Data() As Byte, Optional ByVal Length As Integer = -1, Optional Peek As Boolean = False)
    If (Length = -1) Then
        Length = m_bufsize - m_bufpos
    End If
    
    If ((m_bufpos + Length) > m_bufsize) Then
        Exit Sub
    ElseIf ((m_bufpos) = m_bufsize) Then
        Data = vbNullString
        Exit Sub
    End If
    
    ' resize buffer
    ReDim Data(0 To Length - 1)

    ' copy data to buffer
    CopyMemory Data(0), m_buf(m_bufpos), Length

    ' store buffer Length
    If (Not Peek) Then m_bufpos = (m_bufpos + Length)
End Sub

Public Function GetByte(Optional Peek As Boolean = False) As Byte
    If ((m_bufpos + 1) > m_bufsize) Then
        Exit Function
    End If

    GetByte = m_buf(m_bufpos)

    If (Not Peek) Then m_bufpos = (m_bufpos + 1)
End Function

Public Function InsertWord(ByVal Data As Integer)
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 2))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data, 2
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 2)
End Function

Public Function GetWord(Optional Peek As Boolean = False) As Integer
    If ((m_bufpos + 2) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory GetWord, m_buf(m_bufpos), 2

    If (Not Peek) Then m_bufpos = (m_bufpos + 2)
End Function

Public Sub InsertDWord(ByVal Data As Long)
    ' resize data buffer
    ReDim Preserve m_buf(0 To (m_bufsize + 4))

    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), Data, 4
    
    ' store buffer Length
    m_bufsize = (m_bufsize + 4)
End Sub

Public Function GetDWord(Optional Peek As Boolean = False) As Long
    If ((m_bufpos + 4) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory GetDWord, m_buf(m_bufpos), 4

    If (Not Peek) Then m_bufpos = (m_bufpos + 4)
End Function

Public Sub InsertBool(ByVal Data As Boolean)
    InsertDWord Abs(CLng(Data))
End Sub

Public Function GetBool(Optional Peek As Boolean = False) As Boolean
    GetBool = Not CBool(GetDWord(Peek) = 0)
End Function

Public Function GetFileTime(Optional Peek As Boolean = False) As Date
    Dim ft As FILETIME

    If ((m_bufpos + 8) > m_bufsize) Then
        Exit Function
    End If

    ' copy data to buffer
    CopyMemory ft, m_buf(m_bufpos), 8

    If (Not Peek) Then m_bufpos = (m_bufpos + 8)

    GetFileTime = FileTimeToDate(ft)
End Function

Public Sub InsertNonNTString(ByVal Data As String)
    Dim arrStr() As Byte
    
    arrStr() = StringToByteArr(Data)
    
    ' resize buffer
    ReDim Preserve m_buf(0 To (m_bufsize + UBound(arrStr) + 1))
    
    ' copy data to buffer
    CopyMemory m_buf(m_bufsize), arrStr(0), UBound(arrStr) + 1
    
    ' store buffer Length
    m_bufsize = (m_bufsize + UBound(arrStr) + 1)
End Sub

Public Sub InsertNTString(ByRef Data As String, _
    Optional ByVal Encoding As STRINGENCODING = STRINGENCODING.ANSI)
    
    Dim arrStr() As Byte
    
    Select Case (Encoding)
        Case STRINGENCODING.ANSI
            ' convert from Unicode (VB6 string) to Byte()
            arrStr() = modChat.UTF8Encode(Data, 1252)
            
        Case STRINGENCODING.UTF8
            ' encode result as Unicode (encoded in UTF-8)
            arrStr() = modChat.UTF8Encode(Data)
            
        Case STRINGENCODING.UTF16
            ' use Unicode (VB6 string) as-is
            arrStr() = Data
    End Select
    
    ' resize buffer and include terminating null character
    ReDim Preserve m_buf(0 To (m_bufsize + (UBound(arrStr) + 1)))
    
    ' copy data to buffer
    If (LenB(Data) > 0) Then
        CopyMemory m_buf(m_bufsize), arrStr(0), (UBound(arrStr) + 1)
    End If

    ' store buffer Length
    m_bufsize = (m_bufsize + (UBound(arrStr) + 1))
End Sub

Public Function GetString(Optional ByVal Encoding As STRINGENCODING = STRINGENCODING.ANSI, Optional Peek As Boolean = False) As String
    Dim i As Integer
    Dim arrStr() As Byte

    For i = m_bufpos To m_bufsize
        If (m_buf(i) = &H0) Then
            Exit For
        End If
    Next i

    If i = m_bufpos Then
        GetString = vbNullString
        If (Not Peek) Then m_bufpos = m_bufpos + 1
        Exit Function
    End If
    
    ' resize buffer and exclude terminating null character
    ReDim arrStr(0 To i - m_bufpos - 1)

    If (i < m_bufsize) Then
        ' copy data to buffer
        CopyMemory arrStr(0), m_buf(m_bufpos), i - m_bufpos

        If (Not Peek) Then m_bufpos = i + 1
    End If

    Select Case (Encoding)
        Case STRINGENCODING.ANSI
            ' convert from Byte() to Unicode (VB6 string)
            GetString = modChat.UTF8Decode(arrStr(), 1252)

        Case STRINGENCODING.UTF8
            ' decode result as Unicode (encoded in UTF-8)
            GetString = modChat.UTF8Decode(arrStr())

        Case STRINGENCODING.UTF16
            ' use Unicode (VB6 string) as-is
            GetString = arrStr()

    End Select
End Function

Public Function GetStringArr(Optional ByVal Encoding As STRINGENCODING = STRINGENCODING.ANSI) As String()
    Dim sArray() As String
    Dim sValue   As String
    Dim i        As Integer

    ' keep adding strings until reaching a 0-length string ("double null")
    i = 0
    Do
        sValue = GetString(Encoding)
        If LenB(sValue) > 0 Then
            ReDim Preserve sArray(0 To i)
            sArray(i) = sValue
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    ' make a 0-length array of string
    If i = 0 Then sArray() = Split(vbNullString)

    GetStringArr = sArray()
End Function

Public Function GetRaw(Optional ByVal Length As Integer = -1, Optional Peek As Boolean = False) As String
    Dim arrStr() As Byte
    
    If (Length = -1) Then
        Length = m_bufsize - m_bufpos
    End If
    
    If ((m_bufpos + Length) > m_bufsize) Then
        Exit Function
    End If

    GetRaw = String$(Length, Chr$(0))

    ' resize buffer
    ReDim arrStr(0 To Length - 1)

    ' copy data to buffer
    CopyMemory arrStr(0), m_buf(m_bufpos), Length
    
    GetRaw = ByteArrToString(arrStr())
    
    If (Not Peek) Then m_bufpos = m_bufpos + Length
End Function ' end function GetString

Public Property Get Length() As Long
    Length = m_bufsize
End Property

Public Property Get Position() As Long
    Position = m_bufpos
End Property

Public Property Let Position(ByVal lng As Long)
    m_bufpos = lng
End Property

Public Function Clear()
    ' resize buffer
    ReDim m_buf(0)
    
    ' clear first index
    m_buf(0) = 0
    
    ' reset buffer Length
    m_bufsize = 0
    
    ' reset buffer position
    m_bufpos = 0
End Function

Public Function DebugOutput(Optional ByVal Start As Long = 0, Optional ByVal Length As Long = -1) As String
    If Length < 0 Or Length + Start > m_bufsize Then Length = m_bufsize - Start
    DebugOutput = modPacketBuffer.DebugOutput(Data, Start, Length)
End Function

Public Function SendData(ByVal Socket As Winsock, Optional ByVal PacketID As Byte, Optional ByVal ServerType As enuServerTypes = stGEN, Optional ByVal HeaderType As enuPacketHeaderTypes = phtNONE) As Boolean
    If (Socket Is frmChat.sckBNet) Then
        ' catch packets being sent with their header
        If HeaderType = phtNONE And m_bufsize >= 4 Then PacketID = m_buf(1)

        ' scripts can't send these via the DataBuffer
        If GetCripple Then
            Select Case (PacketID)
                Case SID_CHATCOMMAND, SID_JOINCHANNEL:
                    Clear
                    Exit Function
            End Select
        End If

        ' These two packets cause the client to leave chat, and do not have any responses.
        '  (SID_NOTIFYJOIN is not valid unless it's at least 10 bytes long)
        If (PacketID = SID_LEAVECHAT Or (PacketID = SID_NOTIFYJOIN And m_bufsize >= 10)) Then
            ' there's no response to this one!
            Call Event_LeftChatEnvironment
        End If
    End If

    SendData = modPacketBuffer.SendData(m_buf, m_bufsize, Not IsMissing(PacketID), PacketID, Socket, ServerType, HeaderType)

    If (Socket Is frmChat.sckBNet) Then
        If m_bufsize > 0 Then
            'Send Warden Everything thats Sent to Bnet
            Call modWarden.WardenData(WardenInstance, m_buf, True, PacketID, m_bufsize)
        End If
    End If

    Clear
End Function

Public Function SendPacketMCP(Optional ByVal PacketID As Byte) As Boolean
    SendPacketMCP = Me.SendData(frmChat.sckMCP, PacketID, stMCP, phtMCP)
End Function

Public Function SendPacket(Optional ByVal PacketID As Byte) As Boolean
    SendPacket = Me.SendData(frmChat.sckBNet, PacketID, stBNCS, phtBNCS)
End Function

Public Function vLSendPacket(Optional ByVal PacketID As Byte) As Boolean
    vLSendPacket = Me.SendData(frmChat.sckBNLS, PacketID, stBNLS, phtMCP)
End Function

Public Sub GetDataAndAppend(ByVal Socket As Winsock, Optional ByVal Length As Long)
    Dim buf() As Byte
    ' read buffer as Byte()
    Socket.GetData buf(), vbArray + vbByte, Length
    ' add data to buffer
    Me.InsertByteArr buf()
End Sub

Public Function HandleRecvData(Optional ByRef PacketID As Byte, Optional ByRef PacketLength As Long, Optional ByVal ServerType As enuServerTypes = stBNCS, Optional ByVal HeaderType As enuPacketHeaderTypes = phtBNCS, Optional ByVal ScriptSource As Boolean = False) As Boolean
    Dim HasPacketID As Boolean

    Select Case HeaderType
        Case phtBNCS:
            Me.GetByte
            PacketID = Me.GetByte
            PacketLength = Me.GetWord
            HasPacketID = True
        Case phtMCP:
            PacketLength = Me.GetWord
            PacketID = Me.GetByte
            HasPacketID = True
        Case Else:
            HasPacketID = False
    End Select

    HandleRecvData = modPacketBuffer.HandleRecvData(m_buf, m_bufsize, HasPacketID, PacketID, ServerType, HeaderType, ScriptSource)
End Function

Public Function FindByte(ByVal Value As Integer) As Long
    Dim i As Integer

    For i = m_bufpos To m_bufsize
        If (m_buf(i) = Value) Then
            FindByte = i + 1
            Exit Function
        End If
    Next i
    
    FindByte = 0
End Function

Public Function IsFullPacket(ByVal HeaderLenStart As Integer) As Boolean
    Dim lngPacketLen As Long, ByteIndex As Long
    
    IsFullPacket = False
    
    If m_bufsize > 0 Then
        lngPacketLen = GetPacketLength(HeaderLenStart)
        
        If (lngPacketLen = 0) Then
            Exit Function
        End If

        If (m_bufsize >= lngPacketLen) Then
            If lngPacketLen < 10000 Then
                IsFullPacket = True
            Else
                frmChat.AddChat g_Color.ErrorMessageText, "Error: Packet Length of unusually high Length detected! Packet " & _
                        "dropped. Buffer content at this time: " & vbCrLf & DebugOutput()

                Call Clear
            End If
        End If
       'Else
       '    frmChat.AddChat g_Color.ErrorMessageText, "Error: The front of the buffer is not a valid packet!"
       '
       '    If MDebug("showdrops") Then
       '        frmChat.AddChat g_Color.ErrorMessageText, "Error: The front of the buffer is not " & _
       '                "a valid packet!"
       '        frmChat.AddChat g_Color.ErrorMessageText, "The following data is being purged:"
       '
       '        If ByteIndex > 0 Then
       '            frmChat.AddChat g_Color.ErrorMessageText, Space$(1) & DebugOutput()
       '        Else
       '            frmChat.AddChat g_Color.ErrorMessageText, Space$(1) & DebugOutput()
       '        End If
       '    End If
       '
       '    If ByteIndex > 0 Then
       '        m_bufpos = ByteIndex
       '        Call Clear
       '    Else
       '        Call Clear
       '    End If
       'End If
    End If
End Function

Public Function GetPacketLength(ByVal HeaderLenStart As Integer) As Long
    Dim Value As Long
    If HeaderLenStart = 0 Then
        Value = GetWord(True)
    ElseIf HeaderLenStart = 2 Then
        Value = GetDWord(True)
        Value = CLng(Value \ &H10000)
    ElseIf HeaderLenStart = 1 Then
        ' this one's unlikely but...
        Value = GetDWord(True)
        Value = CLng(Value \ &H100) And &HFFFF&
    End If
    GetPacketLength = Value
End Function

Public Function TakePacket(ByVal HeaderLenStart As Integer) As clsDataBuffer
    Dim tmpbuf() As Byte
    Dim lngPacketLen As Long
    Dim pBuff As clsDataBuffer

    lngPacketLen = GetPacketLength(HeaderLenStart)
    
    ' returns new buffer
    Set TakePacket = New clsDataBuffer
    
    ' set buffer to packet portion
    TakePacket.Clear
    If lngPacketLen >= 0 Then
        Call GetByteArr(tmpbuf, lngPacketLen, False)
        Call TakePacket.SetDataAsByteArr(tmpbuf)
    End If
    
    ' clear or set remaining buffer to remove the read part
    If m_bufpos = m_bufsize Then
        Call Clear
    Else
        Call GetByteArr(tmpbuf, m_bufsize - m_bufpos, False)
        Call SetDataAsByteArr(tmpbuf)
        m_bufpos = 0
    End If
End Function


