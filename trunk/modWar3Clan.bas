Attribute VB_Name = "modWar3Clan"
Option Explicit
'10-18-07 - Hdx - Removed ClanInfoSplit - What was he thinking >.<

'Public sDebugBuf As String
Public AwaitingClanList       As Byte
Public AwaitingClanMembership As Byte
Public AwaitingClanInfo       As Byte
Public LastRemoval            As Long

Public Type udtClan
    Token   As String * 4
    Creator As String
    DWName  As String * 4
    Name    As String
    MyRank  As Byte
    isNew   As Byte
    isUsed  As Boolean
End Type

Public Clan As udtClan

Public Function IsW3() As Boolean

    IsW3 = (BotVars.Product = "PX3W" Or BotVars.Product = "3RAW")

End Function

Public Sub RequestClanList()
    AwaitingClanList = 1
    
    g_Clan.Clear
    frmChat.lvClanList.ListItems.Clear
    
    PBuffer.InsertDWord &H1
    PBuffer.SendPacket SID_CLANMEMBERLIST
End Sub

Public Sub DisbandClan()
    With PBuffer
        .InsertDWord &H1
        .SendPacket SID_CLANDISBAND
    End With
End Sub

Public Sub InviteToClan(Username As String) '//Works
    If (LenB(Username) = 0) Then Exit Sub
    With PBuffer
        .InsertDWord &H1
        .InsertNTString Username
        .SendPacket SID_CLANINVITATION
    End With
End Sub

Public Sub RequestClanMOTD(Optional ByVal cookie As Long = &H0)
    PBuffer.InsertDWord cookie
    PBuffer.SendPacket SID_CLANMOTD
End Sub

Public Sub SetClanMOTD(Message As String) '//Works
    With PBuffer
        .InsertDWord &H0
        .InsertNTString Message
        .SendPacket SID_CLANSETMOTD
    End With
End Sub

Public Sub PromoteMember(Username As String, Rank As Integer)
    With PBuffer
        .InsertDWord &H3
        .InsertNTString Username
        .InsertByte Rank
        .SendPacket SID_CLANRANKCHANGE
    End With
End Sub

Public Sub DemoteMember(Username As String, Rank As Integer)
    With PBuffer
        .InsertDWord &H1
        .InsertNTString Username
        .InsertByte Rank
        .SendPacket SID_CLANRANKCHANGE
    End With
End Sub

Public Sub RemoveMember(Username As String)
    With PBuffer
        .InsertDWord &H1
        .InsertNTString Username
        .SendPacket SID_CLANREMOVEMEMBER
    End With
End Sub

Public Sub MakeMemberChieftain(Username As String)
    With PBuffer
        .InsertDWord &H1
        .InsertNTString Username
        .SendPacket SID_CLANMAKECHIEFTAIN
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

Public Function TimeSinceLastRemoval() As Long
    Dim L As Long
    
    If LastRemoval > 0 Then
        L = GetTickCount
        
        TimeSinceLastRemoval = ((L - LastRemoval) / 1000)
    Else
        TimeSinceLastRemoval = 30
    End If
End Function
