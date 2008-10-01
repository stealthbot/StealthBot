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
Const SID_CLANMOTD& = &H7C


Public Event CandidateList(ByVal Status As Byte, ByRef Users() As String)
Public Event DisbandClanReply(ByVal Success As Boolean)
Public Event ClanInfo(ByVal ClanTag As String, ByVal RawClanTag As String, ByVal Rank As Byte)
Public Event InviteUserReply(ByVal Status As Byte)
Public Event ClanInvitation(ByVal Token As String, ByVal ClanTag As String, ByVal RawClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal NewClan As Boolean)
Public Event ClanMemberUpdate(ByVal Username As String, ByVal Rank As Byte, ByVal IsOnline As Byte, ByVal Location As String)
Public Event ClanMemberList(ByRef Members() As String)
Public Event UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)
Public Event DemoteUserReply(ByVal Success As Boolean)
Public Event PromoteUserReply(ByVal Success As Boolean)
Public Event RemoveUserReply(ByVal result As Byte)
Public Event MyRankChange(ByVal NewRank As Byte)
Public Event MemberLeaves(ByVal Member As String)
Public Event RemovedFromClan(ByVal Status As Byte)
Public Event ClanMOTD(ByVal cookie As Long, ByVal Message As String)

'10-18-07 - Hdx - Changed to use clsPacketDebuffer
Public Sub ParseClanPacket(ByVal PacketID As Byte, ByVal Data As String)
    On Error GoTo ERROR_HANDLER
    
    Dim ary() As String
    'Dim dwTemp As Long
    Dim ClanTag As String * 4
    Dim sTemp As String ', Token As String
    Dim sTemp2 As String
    Dim iTemp As Integer
    Dim bTemp As Byte
    Dim bRank As Byte
    Dim bStatus As Byte
    
    Dim inBuf As New clsPacketDebuffer
    inBuf.Debuffer = Data
    
    Select Case PacketID
    
        Case SID_FINDCLANCANDIDATES 'Clan candidates
            inBuf.DebuffRaw 4 'Cookie
            bStatus = inBuf.DebuffBYTE
            bRank = inBuf.DebuffBYTE
            If (bRank > 0) Then
                ReDim ary(0 To bRank - 1)
                For iTemp = 1 To bRank
                    ary(iTemp - 1) = inBuf.DebuffNTString
                Next iTemp
            End If
            RaiseEvent CandidateList(bStatus, ary())
            
    
        Case SID_CLANMEMBERUPDATE 'Clan Info Update
            sTemp = inBuf.DebuffNTString
            bRank = inBuf.DebuffBYTE
            bStatus = inBuf.DebuffBYTE
            sTemp2 = inBuf.DebuffNTString
            RaiseEvent ClanMemberUpdate(sTemp, bRank, bStatus, sTemp2)
            
            
        Case SID_CLANINFO   'Clan Info
            inBuf.RemoveBYTE 'Unknown (0)
            ClanTag = inBuf.DebuffRaw(4)
            bRank = inBuf.DebuffBYTE
            RaiseEvent ClanInfo(KillNull(StrReverse(ClanTag)), ClanTag, bRank)
            
        Case SID_CHANGERANK 'Action Response
            
            'demote: 1; promote: 3
            
            iTemp = inBuf.DebuffDWORD
            bStatus = inBuf.DebuffBYTE

            Select Case bStatus
                Case 0  'success
                    Select Case iTemp
                        Case 1: RaiseEvent DemoteUserReply(True)
                        Case 3: RaiseEvent PromoteUserReply(True)
                    End Select
                    
                Case 2, 7, 8  'too soon
                    Select Case iTemp
                        Case 1: RaiseEvent DemoteUserReply(False)
                        Case 3: RaiseEvent PromoteUserReply(False)
                    End Select
                    
            End Select
            
        Case SID_CLANMEMBERLIST 'Clan listing
            inBuf.RemoveDWORD
            bTemp = inBuf.DebuffBYTE
            ReDim ary(0 To bTemp * 4 - 1)
            For iTemp = 0 To bTemp - 1
              ary(iTemp * 4) = inBuf.DebuffNTString()
              ary(iTemp * 4 + 1) = inBuf.DebuffBYTE()
              ary(iTemp * 4 + 2) = inBuf.DebuffBYTE()
              ary(iTemp * 4 + 3) = inBuf.DebuffNTString()
            Next iTemp
            RaiseEvent ClanMemberList(ary())
            
            
        Case SID_REMOVEMEMBER:    RaiseEvent RemoveUserReply(inBuf.GetByte)
        Case SID_MEMBERLEFT:      RaiseEvent MemberLeaves(inBuf.GetNTString)
        Case SID_REMOVEDFROMCLAN: RaiseEvent RemovedFromClan(inBuf.DebuffBYTE)
            
        Case SID_DISBANDCLAN
            inBuf.RemoveDWORD 'cookie
            RaiseEvent DisbandClanReply((inBuf.GetByte = 0))
            
            
        Case SID_CLANINVITE, SID_CLANINVITE2
        
            sTemp = inBuf.DebuffRaw(4)
            ClanTag = inBuf.DebuffRaw(4)
            ReDim ary(2)
            ary(0) = inBuf.DebuffNTString
            ary(1) = inBuf.DebuffNTString
            
            RaiseEvent ClanInvitation(sTemp, StrReverse(ClanTag), ClanTag, ary(0), ary(1), (PacketID = SID_CLANINVITE2))
            
        Case SID_CLANREQUEST
            inBuf.RemoveDWORD
            RaiseEvent InviteUserReply(inBuf.GetByte)
            
        Case SID_NEWRANKRECEIVED
'            (BYTE) - Old rank
'            (BYTE) - New rank
'            (STRING) - User who changed your rank
            inBuf.RemoveBYTE
            RaiseEvent MyRankChange(inBuf.GetByte)
            
        Case SID_CLANUSERINFO
            'frmChat.AddChat vbRed, "!"
        
        Case SID_CLANMOTD
            Dim cookie  As Long
            Dim Message As String
            
            cookie = inBuf.DebuffDWORD
            inBuf.RemoveDWORD
            Message = inBuf.GetNTString
        
            RaiseEvent ClanMOTD(cookie, Message)

        Case Else
        
            RaiseEvent UnknownClanEvent(PacketID, DebugOutput(Data))
            
    End Select
    inBuf.Clear
    Set inBuf = Nothing
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.description & " in ParseClanPacket()."
    
    Exit Sub

  'Debug.Print "ParseClanPacket Error: " & Err.Number & ": " & Err.description
  'Debug.Print DebugOutput(Data)
  'Err.Clear
End Sub
