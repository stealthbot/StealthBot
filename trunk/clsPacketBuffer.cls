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
Option Explicit

Private Buffer As String

Public Function InsertDWORDArray(Data() As Long)
Dim i As Integer
For i = LBound(Data) To UBound(Data) Step 1
    Buffer = Buffer & MakeDWORD(Data(i))
Next i
End Function

Public Function InsertDWord(ByVal Data As Long)
Buffer = Buffer & MakeDWORD(Data)
End Function

Public Function InsertWord(ByVal Data As Integer)
Buffer = Buffer & MakeWORD(Data)
End Function

Public Function InsertByte(ByVal Data As Integer)
Buffer = Buffer & Chr(Data)
End Function

Public Function InsertNTString(ByVal Data As String)
Buffer = Buffer & Data & Chr(0)
End Function

Public Function InsertNonNTString(ByVal Data As String)
Buffer = Buffer & Data
End Function

Public Function MakeDWORD(Data As Long) As String
    Dim tmp As String
    tmp = Right("00000000" & Hex(Data), 8)
    Dim A As String, b As String, c As String, D As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    c = Mid(tmp, 5, 2)
    D = Mid(tmp, 7, 2)
    MakeDWORD = Chr(Val("&H" & D))
    MakeDWORD = MakeDWORD & Chr(Val("&H" & c))
    MakeDWORD = MakeDWORD & Chr(Val("&H" & b))
    MakeDWORD = MakeDWORD & Chr(Val("&H" & A))
End Function

Public Function MakeWORD(Data As Integer) As String
    Dim tmp As String
    tmp = Right("0000" & Hex(Data), 4)
    Dim A As String, b As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    MakeWORD = Chr(Val("&H" & b))
    MakeWORD = MakeWORD & Chr(Val("&H" & A))
End Function

Public Function GetDWord(ByVal Data As String) As Long
    Dim tmp As String
    tmp = StrToHex(Data)
    Dim A As String, b As String, c As String, D As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    c = Mid(tmp, 5, 2)
    D = Mid(tmp, 7, 2)
    tmp = D & c & b & A
    GetDWord = Val("&H" & tmp)
End Function

Public Function GetWord(ByVal Data As String) As Long
    Dim tmp As String
    tmp = ToHex(Data)
    Dim A As String, b As String
    A = Mid(tmp, 1, 2)
    b = Mid(tmp, 3, 2)
    tmp = b & A
    GetWord = Val("&H" & tmp)
End Function

Public Function Clear()
    Buffer = vbNullString
End Function

Public Function SendPacketMCP(Optional PacketID As Byte)
    On Error Resume Next
    If MDebug("all") > 0 Then
        frmChat.AddChat COLOR_BLUE, "MCP SEND 0x" & Right$("00" & Hex(PacketID), 2)
    End If
    
    If frmChat.sckMCP.State <> 0 Then
        frmChat.sckMCP.SendData MakeWORD(Len(Buffer) + 3) & Chr(PacketID) & Buffer
    End If
    Clear
End Function

Public Function SendPacket(Optional PacketID As Byte)
    On Error Resume Next
    
    Dim l     As Long
    Dim Buf() As Byte
    Dim i     As Integer
    
    If MDebug("all") Then
        frmChat.AddChat COLOR_BLUE, "BNET SEND 0x" & Right$("00" & Hex(PacketID), 2)
    End If
    
    l = Len(Buffer)
    
    If frmChat.sckBNet.State <> 0 Then
        Buffer = Chr(&HFF) & Chr(PacketID) & MakeWORD(l + 4) & Buffer
        
        For i = 0 To LenB(Buffer)
            ReDim Preserve Buf(i)
        
            Buf(i) = Asc(Mid$(Buffer, i + 1, 1))
        Next i
        
        'Call SendBytes(frmChat.sckBNet.SocketHandle, byt(), Len(Buffer), 0)
        
        frmChat.sckBNet.SendData Buffer
        
        'Debug.Print DebugOutput(Chr(&HFF) & Chr(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer)
    End If
    
    LogPacketRaw stBNCS, CtoS, PacketID, l + 4, Buffer
    
    Clear
End Function

Public Function vLSendPacket(Optional PacketID As Byte)
    On Error Resume Next
    Dim l As Long
    
    l = Len(Buffer)
    
    If MDebug("all") Then frmChat.AddChat COLOR_BLUE, "BNLS SEND 0x" & Right$("00" & Hex(PacketID), 2)
    
    If frmChat.sckBNLS.State <> 0 Then
        Buffer = MakeWORD(l + 3) & Chr(PacketID) & Buffer
        frmChat.sckBNLS.SendData Buffer
    End If
    
    LogPacketRaw stBNLS, CtoS, PacketID, l + 3, Buffer
    
    Clear
End Function

Public Function DumpHex() As String
    Dim i As Integer
    DumpHex = StrToHex(MakeWORD(Len(Buffer) + 3)) & Space(1) & Hex(&H1)
    For i = 1 To Len(Buffer)
        DumpHex = DumpHex & Space(1) & Right("00" & Hex(Asc(Mid(Buffer, i, 1))), 2)
    Next i
End Function

Public Function ToHex(Data As String) As String
    Dim i As Integer
    For i = 1 To Len(Data)
        ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
    Next i
End Function

Public Function GetPacket(PacketID As Byte) As String
    GetPacket = Chr(&HFF) & Chr(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
    Call Clear
End Function
'09-24-07 - Added by Hdx. Pretty self evidant.
Public Function GetBNLSPacket(PacketID As Byte) As String
    GetBNLSPacket = MakeWORD(Len(Buffer) + 3) & Chr(PacketID) & Buffer
    Call Clear
End Function

