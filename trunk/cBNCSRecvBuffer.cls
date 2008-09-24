VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBNCSRecvBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cBNCSBufferfer.cls
' cuphead@valhallalegends.com
'  Includes debugging code by Stealth (WriteLog etc)
'  Includes a modified FullPacket sub also by Stealth that checks for a valid
'   Battle.net packet at the front of the buffer (11/8/06)
Option Explicit

Private strData As String

Public Sub AddData(Data As String)
'    If InStr(Data, Chr(&HFF)) <> 1 Then
'        If LenB(strData) > 0 Then
'            strData = strData & Data
'        Else
'            If MDebug("showdrops") Then
'                frmChat.AddChat vbRed, "Packet data dropped due to bad formatting:"
'                frmChat.AddChat vbRed, DebugOutput(Data)
'            End If
'        End If
'    Else
        strData = strData & Data
'    End If
    'WriteLog "Data received. Added to buffer:", True
    'WriteLog Data
End Sub

'Public Sub WriteLog(ByVal s As String, Optional ByVal NoDebug As Boolean = False)
'    If Dir$(App.Path & "\packetlog.txt") = "" Then
'        Open App.Path & "\packetlog.txt" For Output As #1
'        Close #1
'    End If
'
'    Open App.Path & "\packetlog.txt" For Append As #1
'        If NoDebug Then
'            Print #1, s
'        Else
'            Print #1, DebugOutput(s) & vbCrLf
'        End If
'    Close #1
'End Sub

Public Function GetBuffer() As String
    GetBuffer = strData
End Function

Public Function FullPacket() As Boolean
    Dim lngPacketLen As Long, l As Long
    
    FullPacket = False
    
    If Len(strData) > 0 Then
        l = InStr(strData, Chr(&HFF))
        
        If l = 1 Then
            lngPacketLen = StringToWord(Mid$(strData, 3, 2))
            
            If (lngPacketLen = 0) Then
                Exit Function
            End If

            If (Len(strData) >= lngPacketLen) Then
                If lngPacketLen < 10000 Then
                    FullPacket = True
                Else
                    frmChat.AddChat vbRed, "Error: Packet Length of unusually high Length detected! Packet " & _
                            "dropped. Buffer content at this time: " & vbCrLf & DebugOutput(strData)
                            
                    Call ClearBuffer
                End If
            End If
        Else
            frmChat.AddChat vbRed, "Error: The front of the buffer is not a valid packet!"
        
            If MDebug("showdrops") Then
                frmChat.AddChat RTBColors.ErrorMessageText, "Error: The front of the buffer is not " & _
                        "a valid packet!"
                frmChat.AddChat RTBColors.ErrorMessageText, "The following data is being purged:"
                
                If l > 0 Then
                    frmChat.AddChat Space$(1) & DebugOutput(Mid$(strData, 1, l - 1))
                Else
                    frmChat.AddChat Space$(1) & DebugOutput(strData)
                End If
            End If
            
            If l > 0 Then
                strData = Mid$(strData, l)
            Else
                strData = ""
            End If
        End If
    End If
End Function

Public Function GetPacket() As String
    Dim lngPacketLen As Long

    lngPacketLen = StringToWord(Mid$(strData, 3, 2))
    
    'WriteLog "Pulling a packet. Length: " & lngPacketLen, True
    
    If lngPacketLen > 0 Then
        'frmChat.AddChat RTBColors.ErrorMessageText, "-> Warning: Invalid [low] packet Length specified! Flushing the BNCS receive buffer"
        GetPacket = Mid$(strData, 1, lngPacketLen)
        
        'WriteLog GetPacket
        strData = Mid$(strData, lngPacketLen + 1)
        'WriteLog "Buffer now contains: ", True
        'WriteLog GetBuffer
    End If
End Function

Public Sub ClearBuffer()
    strData = vbNullString
End Sub

Public Sub VoidTrimBuffer()
    Dim LastPacketStart As Long
    
    LastPacketStart = InStrRev(strData, Chr(&HFF) & Chr(&HF), , vbBinaryCompare)
    
    If LastPacketStart > 0 Then
        strData = Mid$(strData, LastPacketStart)
    End If
End Sub

'Private Function StringToWord(Data As String) As Long
'    Dim tmp As String
'    Dim A As String, b As String
'
'    tmp = ToHex(Data)
'    A = Mid(tmp, 3, 2)
'    b = Mid(tmp, 1, 2)
'    tmp = A & b
'    StringToWord = Val("&H" & tmp)
'End Function

Private Function ToHex(Data As String) As String
    Dim i As Integer
    
    For i = 1 To Len(Data)
        ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
    Next i
End Function


'Implementation in _DataArrival:
'    Dim strData As String
'    sckBNCS.GetData strData
'
'    BNCSData.AddData strData
    
'    While BNCSData.FullPacket = True
'        Call ParsingRoutine(BNCSData.GetPacket)
'    Wend

