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

Const SID_FINDCLANCANDIDATES& = &H70
'// Const SID_INVITEMULTIPLEUSERS& = &H71 -- not supported
Const SID_DISBANDCLAN& = &H73
Const SID_CLANINFO& = &H75
Const SID_REMOVEDFROMCLAN& = &H76
Const SID_CLANREQUEST& = &H77
Const SID_REMOVEMEMBER& = &H78
Const SID_CLANINVITE& = &H79
Const SID_CLANINVITE2& = &H72
Const SID_CLANMEMBERLIST& = &H7D
Const SID_CLANMEMBERUPDATE& = &H7F
Const SID_MEMBERLEFT& = &H7E
Const SID_CHANGERANK& = &H7A
Const SID_NEWRANKRECEIVED& = &H81
Const SID_CLANUSERINFO& = &H82 '// arbitrary name


Public Event CandidateList(ByVal Status As Byte, ByRef Users() As String)
Public Event DisbandClanReply(ByVal Success As Boolean)
Public Event ClanInfo(ByVal ClanTag As String, ByVal RawClanTag As String, ByVal Rank As Byte)
Public Event InviteUserReply(ByVal Status As Byte)
Public Event ClanInvitation(ByVal Token As String, ByVal ClanTag As String, ByVal RawClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal NewClan As Boolean)
Public Event ClanMemberUpdate(ByVal Username As String, ByVal Rank As Byte, ByVal IsOnline As Integer)
Public Event ClanMemberList(ByRef Members() As String)
Public Event UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)
Public Event DemoteUserReply(ByVal Success As Boolean)
Public Event PromoteUserReply(ByVal Success As Boolean)
Public Event RemoveUserReply(ByVal Result As Byte)
Public Event MyRankChange(ByVal NewRank As Byte)
Public Event MemberLeaves(ByVal Member As String)
Public Event RemovedFromClan(ByVal Status As Byte)


Public Sub RequestClanList()
    AwaitingClanList = 1
    PBuffer.InsertDWORD &H1
    PBuffer.SendPacket &H7D
End Sub

Public Sub ParseClanPacket(ByVal PacketID As Byte, ByVal Data As String)
    
    Dim ary() As String
    'Dim dwTemp As Long
    Dim ClanTag As String * 4
    Dim sTemp As String ', Token As String
    Dim bTemp As Byte
    
    Select Case PacketID
    
        Case SID_FINDCLANCANDIDATES 'Clan candidates
        
            bTemp = Asc(Mid$(Data, 1, 1))
            
            If bTemp = 0 Then
                ary() = Split(Mid$(Data, 7), Chr(0))
            End If
            
            RaiseEvent CandidateList(bTemp, ary())
            
    
        Case SID_CLANMEMBERUPDATE 'Clan Info Update
        
            sTemp = Mid$(Data, 1, InStr(1, Data, Chr(0)))
            bTemp = Asc(Mid$(Data, InStr(1, Data, Chr(0)) + 1, 1))
            
            RaiseEvent ClanMemberUpdate(sTemp, bTemp, Asc(Mid$(Data, Len(Data) - 1, 1)))
            
            
        Case SID_CLANINFO   'Clan Info
        
            ClanTag = Mid$(Data, 2, 4)
            bTemp = Asc(Mid$(Data, 6, 1))
            
            RaiseEvent ClanInfo(StrReverse(ClanTag), ClanTag, bTemp)
            
        Case SID_REMOVEDFROMCLAN
        
            bTemp = Asc(Mid$(Data, 1, 1))
            RaiseEvent RemovedFromClan(bTemp)
            
            
        Case SID_CHANGERANK 'Action Response
            
            'demote: 1; promote: 3
            
            bTemp = Conv(Mid$(Data, 1, 4)) 'may overflow
            
            Select Case Asc(Mid$(Data, 5, 1))
                Case 0  'success
                    Select Case bTemp
                        Case 1: RaiseEvent DemoteUserReply(True)
                        Case 3: RaiseEvent PromoteUserReply(True)
                    End Select
                    
                Case 2, 7, 8  'too soon
                    Select Case bTemp
                        Case 1: RaiseEvent DemoteUserReply(False)
                        Case 3: RaiseEvent PromoteUserReply(False)
                    End Select
                    
            End Select
            
            
        Case SID_MEMBERLEFT
            RaiseEvent MemberLeaves(KillNull(Mid$(Data, 1)))
            
            
        Case SID_CLANMEMBERLIST 'Clan listing
            
            ClanInfoSplit Right(Data, Len(Data) - 5), ary()
            RaiseEvent ClanMemberList(ary())
            
            
        Case SID_REMOVEMEMBER
        
            bTemp = Asc(Mid$(Data, 1, 1))
            RaiseEvent RemoveUserReply(bTemp)
            
            
        Case SID_DISBANDCLAN
            
            RaiseEvent DisbandClanReply((Val(Mid$(Data, 5, 4)) = 0))
            
            
        Case SID_CLANINVITE, SID_CLANINVITE2
        
            sTemp = Mid$(Data, 1, 4)
            ClanTag = Mid$(Data, 5, 4)
            ary() = Split(Mid$(Data, 9), Chr(0))
            
            RaiseEvent ClanInvitation(sTemp, StrReverse(ClanTag), ClanTag, ary(0), ary(1), (PacketID = SID_CLANINVITE2))
            
        Case SID_CLANREQUEST
        
            bTemp = Asc(Mid$(Data, 5, 1))
            RaiseEvent InviteUserReply(bTemp)
            
        Case SID_NEWRANKRECEIVED
'            (BYTE) - Old rank
'            (BYTE) - New rank
'            (STRING) - User who changed your rank
            bTemp = Asc(Mid$(Data, 2, 1))
            RaiseEvent MyRankChange(bTemp)
            
        Case SID_CLANUSERINFO
            
        
        Case Else
        
            RaiseEvent UnknownClanEvent(PacketID, DebugOutput(Data))
            
    End Select

End Sub
