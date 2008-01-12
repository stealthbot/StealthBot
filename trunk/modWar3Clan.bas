Attribute VB_Name = "modWar3Clan"
Option Explicit
'10-18-07 - Hdx - Removed ClanInfoSplit - What was he thinking >.<

'Public sDebugBuf As String
Public AwaitingClanList As Byte
Public AwaitingClanMembership As Byte
Public AwaitingClanInfo As Byte
Public LastRemoval As Long

Public Type udtClan
    Token As String * 4
    Creator As String
    DWName As String * 4
    Name As String
    MyRank As Byte
    isNew As Byte
    isUsed As Boolean
End Type

Public Clan As udtClan

Public Function IsW3() As Boolean

    IsW3 = (BotVars.Product = "PX3W" Or BotVars.Product = "3RAW")

End Function

Public Sub InviteToClan(Username As String) '//Works
    With PBuffer
        .InsertDWORD &H1
        .InsertNTString Username
        .SendPacket &H77
    End With
End Sub

Public Sub SetClanMOTD(Message As String) '//Works
    With PBuffer
        .InsertDWORD &H0
        .InsertNTString Message
        .SendPacket &H7B
    End With
End Sub

Public Sub PromoteMember(Username As String)
    With PBuffer
        .InsertDWORD &HB
        .InsertNonNTString Username
        .InsertWORD &H3
        .SendPacket &H7A
    End With
End Sub

Public Sub DemoteMember(Username As String)
    With PBuffer
        .InsertDWORD &H1
        .InsertNTString Username
        .InsertBYTE &H1
        .SendPacket &H7A
    End With
End Sub

Public Function GetRank(ByVal i As Byte) As String
    Select Case i
        Case &H4: GetRank = "Chieftain"     'Chief
        Case &H3: GetRank = "Shaman"        'Shaman
        Case &H2: GetRank = "Grunt"         'Grunt
        Case &H1: GetRank = "Peon"          'Peon
        Case &H0: GetRank = "Recruit"       'Recruit
        Case Else: GetRank = "Unknown"
    End Select
End Function

Public Function DebugOutput(ByVal sIn As String) As String

    Dim x1 As Long, y1 As Long
    Dim iLen As Long, iPos As Long
    Dim sB As String, sT As String
    Dim sOut As String
    Dim Offset As Long, sOffset As String
    'build random string to display
    '    y1 = 256
    '    sIn = String(y1, 0)
    '    For x1 = 1 To 256
    '        Mid(sIn, x1, 1) = Chr(x1 - 1)
    '        Mid(sIn, x1, 1) = Chr(255 * Rnd())
    '    Next x1
    iLen = Len(sIn)

    If iLen = 0 Then Exit Function
    sOut = vbNullString
    Offset = 0

    For x1 = 0 To ((iLen - 1) \ 16)
        sOffset = Right$("0000" & Hex(Offset), 4)
        sB = String(48, " ")
        sT = "................"
        For y1 = 1 To 16
            iPos = 16 * x1 + y1
            If iPos > iLen Then Exit For

            Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
            Select Case Asc(Mid(sIn, iPos, 1))
                Case 0, 9, 10, 13
                Case Else
                    Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
            End Select
        Next y1
        If Len(sOut) > 0 Then sOut = sOut & vbCrLf
        sOut = sOut & sOffset & ":  "
        sOut = sOut & sB & "  " & sT
        Offset = Offset + 16
    Next x1

    'sDebugBuf = sDebugBuf & vbCrLf & vbCrLf & sOut
    DebugOutput = sOut
End Function

Public Function TimeSinceLastRemoval() As Long
    Dim l As Long
    
    If LastRemoval > 0 Then
        l = GetTickCount
        TimeSinceLastRemoval = (l - LastRemoval) / 1000
    Else
        TimeSinceLastRemoval = 31
    End If
End Function
