VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanPacketHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsClanPacketHandler - project StealthBot - authored by Stealth (stealth@stealthbot.net)

'Special thanks:
'-  Ethereal packetlogger was used in my own research
'-  thanks to Arta[vL] and BNetDocs (http://bnetdocs.valhallalegends.com) for additional assistance

Public Event Info(ByVal ClanTag As String, ByVal Rank As enuClanRank)
Public Event MyRankChange(ByVal OldRank As enuClanRank, ByVal NewRank As enuClanRank, ByVal Initiator As String)
Public Event MemberUpdate(ByVal Member As clsClanMemberObj)
Public Event MemberLeaves(ByVal Member As String)
Public Event RemovedFromClan(ByVal Status As Byte)
Public Event InvitationReceived(ByVal Cookie As Long, ByVal ClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal IsNewClan As Boolean, ByRef Users() As String)
Public Event CandidateListReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue, ByRef Users() As String)
Public Event InviteMultipleReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue, ByRef Users() As String)
Public Event DisbandClanReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event MakeChieftainReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event InviteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event DemoteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event PromoteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event RemoveMemberReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
Public Event GetMemberList(ByVal Cookie As Long, ByRef Members() As clsClanMemberObj)
Public Event GetMOTD(ByVal Cookie As Long, ByVal Message As String)
Public Event UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)

Public LastRemoval As Currency

'10-18-07 - Hdx - Changed to use clsPacketDebuffer
Public Sub ParseClanPacket(ByVal PacketID As Byte, ByVal inBuf As clsDataBuffer)
    On Error GoTo ERROR_HANDLER

    Dim oRequest     As udtServerRequest
    Dim Cookie       As Long
    Dim Result       As enuClanResponseValue
    Dim Count        As Byte
    Dim i            As Byte
    Dim ClanTag      As String
    Dim ClanName     As String
    Dim Candidates() As String
    Dim Username     As String
    Dim Message      As String
    Dim OldRank      As enuClanRank
    Dim NewRank      As enuClanRank
    Dim Rank         As enuClanRank
    Dim JoinDate     As Date
    Dim Member       As clsClanMemberObj
    Dim Members()    As clsClanMemberObj
    Dim IsServerReq  As Boolean
    Dim IsPromotion  As Boolean

    IsServerReq = False

    Select Case PacketID
        Case SID_CLANFINDCANDIDATES                 ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            Count = inBuf.GetByte                   ' (BYTE) Candidate count
            Candidates = inBuf.GetStringArr         ' (STRING[]) Candidate names
            IsServerReq = True

        Case SID_CLANINVITEMULTIPLE                 ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            Candidates = inBuf.GetStringArr         ' (STRING[]) Failed candidate names
            IsServerReq = True

        Case SID_CLANCREATIONINVITATION             ' client request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            ClanTag = DWordToString(inBuf.GetDWORD) ' (BYTE) Clan tag
            ClanName = inBuf.GetString              ' (STRING) Clan name
            Username = inBuf.GetString              ' (STRING) Inviter
            Count = inBuf.GetByte                   ' (BYTE) Candidate count
            Candidates = inBuf.GetStringArr         ' (STRING[]) Candidate names

        Case SID_CLANDISBAND                        ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            IsServerReq = True

        Case SID_CLANMAKECHIEFTAIN                  ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            IsServerReq = True

        Case SID_CLANINFO                           ' information event
            inBuf.GetByte                           ' (BYTE) 0
            ClanTag = DWordToString(inBuf.GetDWORD) ' (DWORD) Clan tag
            Rank = inBuf.GetByte                    ' (BYTE) My rank

        Case SID_CLANQUITNOTIFY                     ' information event
            Result = inBuf.GetByte                  ' (BYTE) 1

        Case SID_CLANINVITATION                     ' server request (awaits another user response)
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            IsServerReq = True

        Case SID_CLANREMOVEMEMBER                   ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            IsServerReq = True

        Case SID_CLANINVITATIONRESPONSE             ' client request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            ClanTag = DWordToString(inBuf.GetDWORD) ' (BYTE) Clan tag
            ClanName = inBuf.GetString              ' (STRING) Clan name
            Username = inBuf.GetString              ' (STRING) Inviter
            Candidates = Split(vbNullString) ' empty string()

        Case SID_CLANRANKCHANGE                     ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            IsServerReq = True

        Case SID_CLANMOTD                           ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            inBuf.GetDWORD                          ' (DWORD) 0
            Message = inBuf.GetString               ' (STRING) Message
            IsServerReq = True

        Case SID_CLANMEMBERLIST                     ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Count = inBuf.GetByte                   ' (BYTE) Member count
            ReDim Members(0 To Count - 1)           ' (MEMBER[]) Members
            For i = 0 To Count - 1
                Set Members(i) = New clsClanMemberObj
                With Members(i)
                    .Name = inBuf.GetString         ' (STRING) Username
                    .Rank = inBuf.GetByte           ' (BYTE) Rank
                    .Status = inBuf.GetByte         ' (BYTE) Status
                    .Location = inBuf.GetString     ' (STRING) Location
                End With
            Next i
            IsServerReq = True

        Case SID_CLANMEMBERREMOVED                  ' information event
            Result = inBuf.GetByte                  ' (BYTE) 1

        Case SID_CLANMEMBERSTATUSCHANGE             ' information event
            Set Member = New clsClanMemberObj
            With Member
                .Name = inBuf.GetString             ' (STRING) Username
                .Rank = inBuf.GetByte               ' (BYTE) Rank
                .Status = inBuf.GetByte             ' (BYTE) Status
                .Location = inBuf.GetString         ' (STRING) Location
            End With

        Case SID_CLANMEMBERRANKCHANGE               ' information event
            OldRank = inBuf.GetByte                 ' (BYTE) Old rank
            NewRank = inBuf.GetByte                 ' (BYTE) New rank
            Username = inBuf.GetString              ' (STRING) Initiator

        Case SID_CLANMEMBERINFORMATION              ' standard server request
            Cookie = inBuf.GetDWORD                 ' (DWORD) Cookie
            Result = inBuf.GetByte                  ' (BYTE) Result
            ClanName = inBuf.GetString              ' (STRING) Clan name
            Rank = inBuf.GetByte                    ' (BYTE) Rank
            JoinDate = inBuf.GetFileTime            ' (FILETIME) Join date
            IsServerReq = True
    End Select
    
    If IsServerReq Then
        If Not FindServerRequest(oRequest, Cookie, PacketID, , False) Then
            ' we didn't request this
            ' TODO: error?
            Exit Sub
        End If
    End If

    Select Case PacketID
        Case SID_CLANFINDCANDIDATES
            RaiseEvent CandidateListReply(Cookie, Result, Candidates())

        Case SID_CLANINVITEMULTIPLE
            RaiseEvent InviteMultipleReply(Cookie, Result, Candidates())

        Case SID_CLANCREATIONINVITATION, SID_CLANINVITATIONRESPONSE
            RaiseEvent InvitationReceived(Cookie, ClanTag, ClanName, Username, (PacketID = SID_CLANCREATIONINVITATION), Candidates())

        Case SID_CLANDISBAND
            RaiseEvent DisbandClanReply(Cookie, Result)

        Case SID_CLANMAKECHIEFTAIN
            RaiseEvent MakeChieftainReply(Cookie, Result)

        Case SID_CLANINFO
            RaiseEvent Info(ClanTag, Rank)

        Case SID_CLANQUITNOTIFY
            RaiseEvent RemovedFromClan(Result)

        Case SID_CLANINVITATION
            RaiseEvent InviteUserReply(Cookie, Result)

        Case SID_CLANREMOVEMEMBER
            RaiseEvent RemoveMemberReply(Cookie, Result)

        Case SID_CLANRANKCHANGE
            IsPromotion = CBool(oRequest.Tag(2))
            If IsPromotion Then
                RaiseEvent PromoteUserReply(Cookie, Result)
            Else
                RaiseEvent DemoteUserReply(Cookie, Result)
            End If

        Case SID_CLANMOTD
            RaiseEvent GetMOTD(Cookie, Message)

        Case SID_CLANMEMBERLIST
            RaiseEvent GetMemberList(Cookie, Members())

        Case SID_CLANMEMBERSTATUSCHANGE
            RaiseEvent MemberUpdate(Member)

        Case SID_CLANMEMBERRANKCHANGE
            RaiseEvent MyRankChange(OldRank, NewRank, Username)

        Case Else
            RaiseEvent UnknownClanEvent(PacketID, inBuf.DebugOutput)

    End Select

    Set inBuf = Nothing
    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: " & Err.Description & " in ParseClanPacket()."

    Exit Sub
End Sub

Public Function IsW3() As Boolean

    IsW3 = (BotVars.Product = "PX3W" Or BotVars.Product = "3RAW")

End Function

Public Sub RequestClanList(Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long
    
    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Nothing
        .PacketID = SID_CLANMEMBERLIST
        .PacketCommand = 0
        .Tag = Empty
    End With

    Cookie = SaveServerRequest(oRequest)

    g_Clan.Clear
    frmChat.lvClanList.ListItems.Clear

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANMEMBERLIST
    End With
    Set pBuf = Nothing
End Sub

Public Sub DisbandClan(Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long

    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANDISBAND
        .PacketCommand = 0
        .Tag = Nothing
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANDISBAND
    End With
    Set pBuf = Nothing
End Sub

Public Sub InviteToClan(ByVal Username As String, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long

    If (LenB(Username) = 0) Then Exit Sub

    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANINVITATION
        .PacketCommand = 0
        .Tag = Username
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANINVITATION
    End With
    Set pBuf = Nothing
End Sub

Public Sub InvitationResponse(ByVal IsNew As Boolean, ByVal Cookie As Long, ByVal ClanTag As String, ByVal Inviter As String, ByVal Response As enuClanResponseValue)
    Dim pBuf As clsDataBuffer
    
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertDWord StringToDWord(ClanTag)
        .InsertNTString Inviter
        .InsertByte Response

        If IsNew Then
            .SendPacket SID_CLANCREATIONINVITATION
        Else
            .SendPacket SID_CLANINVITATIONRESPONSE
        End If
    End With
    Set pBuf = Nothing
End Sub

Public Sub CreateFindCandidates(ByVal ClanTag As String, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long

    If (LenB(ClanTag) = 0 Or Len(ClanTag) > 4) Then Exit Sub

    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANFINDCANDIDATES
        .PacketCommand = 0
        .Tag = Username
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertDWord StringToDWord(ClanTag)
        .SendPacket SID_CLANFINDCANDIDATES
    End With
    Set pBuf = Nothing
End Sub

Public Sub CreateInviteMultiple(ByVal ClanTag As String, ByVal ClanName As String, ByRef Users() As String, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long
    Dim i As Integer

    If (LenB(ClanTag) = 0 Or Len(ClanTag) > 4) Then Exit Sub

    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANINVITEMULTIPLE
        .PacketCommand = 0
        .Tag = Username
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString ClanName
        .InsertDWord StringToDWord(ClanTag)
        .InsertByte UBound(Users) + 1
        For i = 0 To UBound(Users)
            .InsertNTString Users(i)
        Next i
        .SendPacket SID_CLANINVITEMULTIPLE
    End With
    Set pBuf = Nothing
End Sub

Public Sub RequestClanMOTD(Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long

    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANMOTD
        .PacketCommand = 0
        .Tag = Empty
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .SendPacket SID_CLANMOTD
    End With
    Set pBuf = Nothing
End Sub

Public Sub SetClanMOTD(ByVal Message As String, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    'Dim oRequest As udtServerRequest
    'Dim Cookie As Long
    
    'With oRequest
    '    .ResponseReceived = False
    '    .HandlerType = eType
    '    Set .Command = Command
    '    .PacketID = SID_CLANSETMOTD
    '    .PacketCommand = 0
    '    .Tag = Message
    'End With

    'Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord 0&
        .InsertNTString Message
        .SendPacket SID_CLANSETMOTD
    End With
    Set pBuf = Nothing
End Sub

Public Sub PromoteMember(ByVal Username As String, ByVal Rank As Integer, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Call ChangeRankMember(Username, Rank, True)
End Sub

Public Sub DemoteMember(ByVal Username As String, ByVal Rank As Integer, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Call ChangeRankMember(Username, Rank, False)
End Sub

Private Sub ChangeRankMember(ByVal Username As String, ByVal Rank As Integer, ByVal IsPromotion As Boolean, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long
    
    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANRANKCHANGE
        .PacketCommand = 0
        .Tag = Array(Username, Rank, IsPromotion)
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .InsertByte Rank
        .SendPacket SID_CLANRANKCHANGE
    End With
    Set pBuf = Nothing
End Sub

Public Sub RemoveMember(ByVal Username As String, ByVal IsSelf As Boolean, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long
    
    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANREMOVEMEMBER
        .PacketCommand = 0
        .Tag = Array(Username, IsSelf)
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANREMOVEMEMBER
    End With
    Set pBuf = Nothing
End Sub

Public Sub MakeMemberChieftain(ByVal Username As String, Optional ByVal eType As enuServerRequestHandlerType = reqScriptingCall, Optional ByVal Command As clsCommandObj)
    Dim pBuf As clsDataBuffer
    Dim oRequest As udtServerRequest
    Dim Cookie As Long
    
    With oRequest
        .ResponseReceived = False
        .HandlerType = eType
        Set .Command = Command
        .PacketID = SID_CLANMAKECHIEFTAIN
        .PacketCommand = 0
        .Tag = Username
    End With

    Cookie = SaveServerRequest(oRequest)

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord Cookie
        .InsertNTString Username
        .SendPacket SID_CLANMAKECHIEFTAIN
    End With
    Set pBuf = Nothing
End Sub

Public Function GetRank(ByVal i As enuClanRank) As String
    Select Case i
        Case clrankChieftain: GetRank = "Chieftain"     'Chief
        Case clrankShaman:    GetRank = "Shaman"        'Shaman
        Case clrankGrunt:     GetRank = "Grunt"         'Grunt
        Case clrankPeon:      GetRank = "Peon"          'Peon
        Case clrankRecruit:   GetRank = "Recruit"       'Recruit
        Case Else:            GetRank = "Unknown"
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

