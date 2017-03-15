Attribute VB_Name = "modWar3Clan"
Option Explicit
'10-18-07 - Hdx - Removed ClanInfoSplit - What was he thinking >.<

'Public sDebugBuf As String
Public AwaitingClanList       As Byte
Public AwaitingClanMembership As Byte
Public LastRemoval            As Currency

Public Type udtClan
    Token   As Long
    Creator As String
    DWName  As String * 4
    Name    As String
    MyRank  As Byte
    IsNew   As Boolean
    IsUsed  As Boolean
End Type

Public Enum enuClanResponseValue
    ClanResponseSuccess = 0
    ClanResponseNameInUse = 1
    ClanResponseTooSoon = 2
    ClanResponseNotEnoughMembers = 3
    ClanResponseDecline = 4
    ClanResponseUnavailable = 5
    ClanResponseAccept = 6
    ClanResponseNotAuthorized = 7
    ClanResponseNotAllowed = 8
    ClanResponseIsFull = 9
    ClanResponseBadTag = 10
    ClanResponseBadName = 11
    ClanResponseUserNotFound = 12
End Enum

Public Clan As udtClan

Public Function IsW3() As Boolean

    IsW3 = (BotVars.Product = "PX3W" Or BotVars.Product = "3RAW")

End Function

Public Sub RequestClanList(Optional ByVal Cookie As Long = &H1)
    Dim pBuf As clsDataBuffer

    AwaitingClanList = 1

    g_Clan.Clear
    frmChat.lvClanList.ListItems.Clear

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANMEMBERLIST
    End With
    Set pBuf = Nothing
End Sub

Public Sub DisbandClan(Optional ByVal Cookie As Long = &H1)
    Dim pBuf As clsDataBuffer

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANDISBAND
    End With
    Set pBuf = Nothing
End Sub

Public Sub InviteToClan(ByVal Username As String, Optional ByVal Cookie As Long = &H1) '//Works
    Dim pBuf As clsDataBuffer

    If (LenB(Username) = 0) Then Exit Sub

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANINVITATION
    End With
    Set pBuf = Nothing
End Sub

Public Sub InvitationResponse(ByVal Response As enuClanResponseValue, Optional ByVal Token As Long, Optional ByVal DWName As String, Optional ByVal Creator As String, Optional ByVal IsNew As Boolean = False)
    Dim pBuf As clsDataBuffer

    If IsMissing(Token) Then Token = Clan.Token
    If IsMissing(DWName) Then DWName = Clan.DWName
    If IsMissing(Creator) Then Creator = Clan.Creator
    
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Token
        .InsertNonNTString DWName
        .InsertNTString Creator
        .InsertByte Response
        
        If IsNew Then
            .SendPacket SID_CLANCREATIONINVITATION
        Else
            .SendPacket SID_CLANINVITATIONRESPONSE
        End If
    End With
    Set pBuf = Nothing
End Sub

Public Sub RequestClanMOTD(Optional ByVal Cookie As Long = &H1)
    Dim pBuf As clsDataBuffer

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANMOTD
    End With
    Set pBuf = Nothing
End Sub

Public Sub SetClanMOTD(ByVal Message As String, Optional ByVal Cookie As Long = &H1) '//Works
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Message
        .SendPacket SID_CLANSETMOTD
    End With
    Set pBuf = Nothing
End Sub

Public Sub PromoteMember(ByVal Username As String, ByVal Rank As Integer)
    Call ChangeRankMember(Username, Rank, &H3)
End Sub

Public Sub DemoteMember(ByVal Username As String, ByVal Rank As Integer)
    Call ChangeRankMember(Username, Rank, &H1)
End Sub

Public Sub ChangeRankMember(ByVal Username As String, ByVal Rank As Integer, ByVal Cookie As Long)
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .InsertByte Rank
        .SendPacket SID_CLANRANKCHANGE
    End With
    Set pBuf = Nothing
End Sub

Public Sub RemoveMember(ByVal Username As String, Optional ByVal Cookie As Long = &H1)
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANREMOVEMEMBER
    End With
    Set pBuf = Nothing
End Sub

Public Sub MakeMemberChieftain(ByVal Username As String, Optional ByVal Cookie As Long = &H1)
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANMAKECHIEFTAIN
    End With
    Set pBuf = Nothing
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
    Dim L As Currency
    
    If LastRemoval > 0 Then
        L = GetTickCountMS()
        
        TimeSinceLastRemoval = (L - LastRemoval) \ 1000
    Else
        TimeSinceLastRemoval = 30
    End If
End Function
