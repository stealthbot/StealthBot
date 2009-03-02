VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBNLSRecvBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cBNCSBufferfer.cls
' cuphead@valhallalegends.com
Option Explicit

Private strData As String

Public Sub AddData(Data As String)
    strData = strData & Data
End Sub

Public Function FullPacket() As Boolean
    Dim lngPacketLen As Long
    
    FullPacket = False
    
    If (Len(strData) > 0) Then
        lngPacketLen = StringToWord(Mid(strData, 1, 2))
        
        If (Len(strData) >= lngPacketLen) Then
            FullPacket = True
        End If
    End If
End Function

Public Function GetPacket() As String
    Dim lngPacketLen As Long

    lngPacketLen = StringToWord(Mid(strData, 1, 2))
    GetPacket = Mid(strData, 1, lngPacketLen)
    strData = Mid(strData, lngPacketLen + 1)
End Function

Public Sub ClearBuffer()
    strData = vbNullString
End Sub

'Private Function StringToWord(Data As String) As Long
'    Dim tmp As String
'    Dim A As String
'    Dim b As String
'
'    tmp = ToHex(Data)
'
'    A = Mid$(tmp, 3, 2)
'    b = Mid$(tmp, 1, 2)
'
'    tmp = A & b
'
'    StringToWord = Val("&H" & tmp)
'End Function

Private Function ToHex(Data As String) As String
    Dim I As Integer
    
    For I = 1 To Len(Data)
        ToHex = ToHex & Right("00" & hex(Asc(Mid(Data, I, 1))), 2)
    Next I
End Function

