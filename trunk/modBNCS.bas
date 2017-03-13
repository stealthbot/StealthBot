Attribute VB_Name = "modBNCS"
Option Explicit
Private Const OBJECT_NAME As String = "modBNCS"

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

' Packet IDs
Public Const SID_NULL                   As Byte = &H0
Public Const SID_CLIENTID               As Byte = &H5
Public Const SID_STARTVERSIONING        As Byte = &H6
Public Const SID_REPORTVERSION          As Byte = &H7
Public Const SID_ENTERCHAT              As Byte = &HA
Public Const SID_GETCHANNELLIST         As Byte = &HB
Public Const SID_JOINCHANNEL            As Byte = &HC
Public Const SID_CHATCOMMAND            As Byte = &HE
Public Const SID_CHATEVENT              As Byte = &HF
Public Const SID_LEAVECHAT              As Byte = &H10
Public Const SID_LOCALEINFO             As Byte = &H12
Public Const SID_UDPPINGRESPONSE        As Byte = &H14
Public Const SID_MESSAGEBOX             As Byte = &H19
Public Const SID_LOGONCHALLENGEEX       As Byte = &H1D
Public Const SID_CLIENTID2              As Byte = &H1E
Public Const SID_NOTIFYJOIN             As Byte = &H22
Public Const SID_PING                   As Byte = &H25
Public Const SID_READUSERDATA           As Byte = &H26
Public Const SID_WRITEUSERDATA          As Byte = &H27
Public Const SID_LOGONCHALLENGE         As Byte = &H28
Public Const SID_GETICONDATA            As Byte = &H2D
Public Const SID_CDKEY                  As Byte = &H30
Public Const SID_CHANGEPASSWORD         As Byte = &H31
Public Const SID_PROFILE                As Byte = &H35
Public Const SID_CDKEY2                 As Byte = &H36
Public Const SID_LOGONRESPONSE2         As Byte = &H3A
Public Const SID_CREATEACCOUNT2         As Byte = &H3D
Public Const SID_LOGONREALMEX           As Byte = &H3E
Public Const SID_QUERYREALMS2           As Byte = &H40
Public Const SID_WARCRAFTGENERAL        As Byte = &H44
Public Const SID_EXTRAWORK              As Byte = &H4C
Public Const SID_AUTH_INFO              As Byte = &H50
Public Const SID_AUTH_CHECK             As Byte = &H51
Public Const SID_AUTH_ACCOUNTCREATE     As Byte = &H52
Public Const SID_AUTH_ACCOUNTLOGON      As Byte = &H53
Public Const SID_AUTH_ACCOUNTLOGONPROOF As Byte = &H54
Public Const SID_AUTH_ACCOUNTCHANGE     As Byte = &H55
Public Const SID_AUTH_ACCOUNTCHANGEPROOF As Byte = &H56
Public Const SID_SETEMAIL               As Byte = &H59
Public Const SID_RESETPASSWORD          As Byte = &H5A
Public Const SID_CHANGEEMAIL            As Byte = &H5B
Public Const SID_FRIENDSLIST            As Byte = &H65
Public Const SID_FRIENDSUPDATE          As Byte = &H66
Public Const SID_FRIENDSADD             As Byte = &H67
Public Const SID_FRIENDSREMOVE          As Byte = &H68
Public Const SID_FRIENDSPOSITION        As Byte = &H69
Public Const SID_CLANFINDCANDIDATES     As Byte = &H70
Public Const SID_CLANINVITEMULTIPLE     As Byte = &H71
Public Const SID_CLANCREATIONINVITATION As Byte = &H72
Public Const SID_CLANDISBAND            As Byte = &H73
Public Const SID_CLANMAKECHIEFTAIN      As Byte = &H74
Public Const SID_CLANINFO               As Byte = &H75
Public Const SID_CLANQUITNOTIFY         As Byte = &H76
Public Const SID_CLANINVITATION         As Byte = &H77
Public Const SID_CLANREMOVEMEMBER       As Byte = &H78
Public Const SID_CLANINVITATIONRESPONSE As Byte = &H79
Public Const SID_CLANRANKCHANGE         As Byte = &H7A
Public Const SID_CLANSETMOTD            As Byte = &H7B
Public Const SID_CLANMOTD               As Byte = &H7C
Public Const SID_CLANMEMBERLIST         As Byte = &H7D
Public Const SID_CLANMEMBERREMOVED      As Byte = &H7E
Public Const SID_CLANMEMBERSTATUSCHANGE As Byte = &H7F
Public Const SID_CLANMEMBERRANKCHANGE   As Byte = &H81
Public Const SID_CLANMEMBERINFORMATION  As Byte = &H82


' SID_CHATEVENT EVENT IDs
Public Const ID_USER = &H1
Public Const ID_JOIN = &H2
Public Const ID_LEAVE = &H3
Public Const ID_WHISPER = &H4
Public Const ID_TALK = &H5
Public Const ID_BROADCAST = &H6
Public Const ID_CHANNEL = &H7
Public Const ID_USERFLAGS = &H9
Public Const ID_WHISPERSENT = &HA
Public Const ID_CHANNELFULL = &HD
Public Const ID_CHANNELDOESNOTEXIST = &HE
Public Const ID_CHANNELRESTRICTED = &HF
Public Const ID_INFO = &H12
Public Const ID_ERROR = &H13
Public Const ID_EMOTE = &H17
' Additional event constants for logging
Public Const ID_CONNECTED = &H18
Public Const ID_DISCONNECTED = &H19

' battle.net user flag constants
Public Const USER_BLIZZREP    As Long = &H1
Public Const USER_CHANNELOP   As Long = &H2
Public Const USER_SPEAKER     As Long = &H4
Public Const USER_SYSOP       As Long = &H8
Public Const USER_NOUDP       As Long = &H10
Public Const USER_BEEPENABLED As Long = &H100
Public Const USER_KBKOFFICIAL As Long = &H1000
Public Const USER_JAILED      As Long = &H100000
Public Const USER_SQUELCHED   As Long = &H20
Public Const USER_PGLPLAYER   As Long = &H200
Public Const USER_GFOFFICIAL  As Long = &H100000
Public Const USER_GFPLAYER    As Long = &H200000
Public Const USER_GUEST       As Long = &H40
Public Const USER_PGLOFFICIAL As Long = &H400
Public Const USER_KBKPLAYER   As Long = &H800


Public Const BNCS_NLS As Long = 1 'New:    SID_AUTH_*
Public Const BNCS_OLS As Long = 2 'Old:    SID_CLIENTID2
Public Const BNCS_LLS As Long = 3 'Legacy: SID_CLIENTID

Public Const BNCSSERVER_XSHA As Long = 0
Public Const BNCSSERVER_SRP2 As Long = 2

Public Const ACCOUNT_MODE_LOGON As String = "LOGON"
Public Const ACCOUNT_MODE_CREAT As String = "CREATE"
Public Const ACCOUNT_MODE_CHPWD As String = "CHANGEPASS"
Public Const ACCOUNT_MODE_RSPWD As String = "RESETPASS"
Public Const ACCOUNT_MODE_CHREG As String = "CHANGEEMAIL"

Public Const EMAIL_ACT_VALUE    As String = "VALUE"
Public Const EMAIL_ACT_PROMPT   As String = "PROMPT"
Public Const EMAIL_ACT_NEVERASK As String = "NEVERASK"
Public Const EMAIL_ACT_ASKLATER As String = "ASKLATER"

Public Const PLATFORM_INTEL   As Long = &H49583836 'IX86
Public Const PLATFORM_POWERPC As Long = &H504D4143 'PMAC
Public Const PLATFORM_OSX     As Long = &H584D4143 'XMAC


Public ds As New clsDataStorage 'Need to rename this -.-

Public Function BNCSRecvPacket(ByVal pBuff As clsDataBuffer) As Boolean
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim PacketID As Byte
    Dim PacketLen As Long
    
    BNCSRecvPacket = True
    With pBuff
        .GetByte
        PacketID = .GetByte
        PacketLen = .GetWord
    End With
    
    If pBuff.Length >= 0 Then
        If MDebug("all") Then
            frmChat.AddChat COLOR_BLUE, "BNET RECV 0x" & ZeroOffset(PacketID, 2)
        End If
        
        Call CachePacket(stBNCS, StoC, PacketID, PacketLen, pBuff.GetDataAsByteArr)
        Call WritePacketData(stBNCS, StoC, PacketID, PacketLen, pBuff.GetDataAsByteArr)
                
        If (RunInAll("Event_PacketReceived", "BNCS", PacketID, PacketLen, pBuff.Data)) Then
            Exit Function
        End If
        
        'This will be taken out when Warden is moved to a script like I want.
        If (modWarden.WardenData(WardenInstance, pBuff.GetDataAsByteArr, False)) Then
            Exit Function
        End If
    
        Select Case PacketID
            Case SID_NULL:                    'Don't Throw Unknown Error                   '0x00
            Case SID_CLIENTID:                'Don't Throw Unknown Error                   '0x05
            Case SID_STARTVERSIONING:         Call RECV_SID_STARTVERSIONING(pBuff)         '0x06
            Case SID_REPORTVERSION:           Call RECV_SID_REPORTVERSION(pBuff)           '0x07
            Case SID_ENTERCHAT:               Call RECV_SID_ENTERCHAT(pBuff)               '0x0A
            Case SID_GETCHANNELLIST:          Call RECV_SID_GETCHANNELLIST(pBuff)          '0x0B
            Case SID_CHATEVENT:               Call RECV_SID_CHATEVENT(pBuff)               '0x0F
            Case SID_MESSAGEBOX:              Call RECV_SID_MESSAGEBOX(pBuff)              '0x19
            Case SID_LOGONCHALLENGEEX:        Call RECV_SID_LOGONCHALLENGEEX(pBuff)        '0x1D
            Case SID_PING:                    Call RECV_SID_PING(pBuff)                    '0x25
            Case SID_READUSERDATA:            Call RECV_SID_READUSERDATA(pBuff)            '0x26
            Case SID_LOGONCHALLENGE:          Call RECV_SID_LOGONCHALLENGE(pBuff)          '0x28
            Case SID_GETICONDATA:             'Don't Throw Unknown Error                   '0x2D
            Case SID_CDKEY:                   Call RECV_SID_CDKEY(pBuff)                   '0x30
            Case SID_CHANGEPASSWORD:          Call RECV_SID_CHANGEPASSWORD(pBuff)          '0x31
            Case SID_CDKEY2:                  Call RECV_SID_CDKEY2(pBuff)                  '0x36
            Case SID_LOGONRESPONSE2:          Call RECV_SID_LOGONRESPONSE2(pBuff)          '0x3A
            Case SID_CREATEACCOUNT2:          Call RECV_SID_CREATEACCOUNT2(pBuff)          '0x3D
            Case SID_LOGONREALMEX:            Call RECV_SID_LOGONREALMEX(pBuff)            '0x3C
            Case SID_QUERYREALMS2:            Call RECV_SID_QUERYREALMS2(pBuff)            '0x40
            Case SID_EXTRAWORK:               'Don't Throw Unknown Error                   '0x4C
            Case SID_AUTH_INFO:               Call RECV_SID_AUTH_INFO(pBuff)               '0x50
            Case SID_AUTH_CHECK:              Call RECV_SID_AUTH_CHECK(pBuff)              '0x51
            Case SID_AUTH_ACCOUNTCREATE:      Call RECV_SID_AUTH_ACCOUNTCREATE(pBuff)      '0x52
            Case SID_AUTH_ACCOUNTLOGON:       Call RECV_SID_AUTH_ACCOUNTLOGON(pBuff)       '0x53
            Case SID_AUTH_ACCOUNTLOGONPROOF:  Call RECV_SID_AUTH_ACCOUNTLOGONPROOF(pBuff)  '0x54
            Case SID_AUTH_ACCOUNTCHANGE:      Call RECV_SID_AUTH_ACCOUNTCHANGE(pBuff)      '0x55
            Case SID_AUTH_ACCOUNTCHANGEPROOF: Call RECV_SID_AUTH_ACCOUNTCHANGEPROOF(pBuff) '0x56
            Case SID_SETEMAIL:                Call RECV_SID_SETEMAIL(pBuff)                '0x59
            
            Case Is >= &H65 'Friends List or Clan-related packet
                ' Hand the packet off to the appropriate handler
                If PacketID >= &H70 Then
                    ' added in response to the clan channel takeover exploit
                    ' discovered 11/7/05
                    If IsW3 Then
                        frmChat.ParseClanPacket PacketID, pBuff
                    End If
                Else
                    If (g_request_receipt) Then
                        g_request_receipt = False
                        
                        If (Caching) Then
                            frmChat.cacheTimer_Timer
                        End If
                        
                        Exit Function
                    End If
                
                    frmChat.ParseFriendsPacket PacketID, pBuff
                End If
        
            Case Else:
                BNCSRecvPacket = False
                If (MDebug("debug") And (MDebug("all") Or MDebug("unknown"))) Then
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unhandled packet 0x{0}", ZeroOffset(CLng(PacketID), 2)))
                    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Packet data: {0}{1}", vbNewLine, pBuff.DebugOutput))
                End If
        
        End Select
    End If
    
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.BNCSRecvPacket()", Err.Number, Err.Description, OBJECT_NAME))
End Function

'*********************************
' SID_CLIENTID (0x05) C->S
'*********************************
' (DWORD) Registration Version
' (DWORD) Registration Authority
' (DWORD) Account Number
' (DWORD) Registration Token
' (STRING) LAN computer name
' (STRING) LAN username
'*********************************
'For legacy login system (JSTR, SSHR).
'*********************************
Public Sub SEND_SID_CLIENTID()
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertDWord 0
        .InsertDWord 0
        .InsertDWord 0
        .InsertDWord 0
        .InsertNTString GetComputerLanName
        .InsertNTString GetComputerUsername
        .SendPacket SID_CLIENTID
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CLIENTID()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_STARTVERSIONING (0x06) S->C
'*******************************
' (FILETIME) MPQ Filetime
' (STRING) MPQ Filename
' (STRING) ValueString
'*******************************
Public Sub RECV_SID_STARTVERSIONING(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    With pBuff
        ds.CRevFileTime = .GetRaw(8)
        ds.CRevFileName = .GetString
        ds.CRevSeed = .GetString
    End With
    
    Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Checking version...")
    If (MDebug("all") Or MDebug("crev")) Then
        frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Name: {0}", ds.CRevFileName)
        frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Time: {0}", ds.CRevFileTime)
        If (InStr(1, ds.CRevFileName, "lockdown", vbTextCompare) > 0) Then
            frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Seed: {0}", StrToHex(ds.CRevSeed))
        Else
            frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Seed: {0}", ds.CRevSeed)
        End If
    End If
    
    ' If the server does not recognize the version byte we sent it, it will send back an empty seed string.
    If (LenB(ds.CRevSeed) = 0) Then
        Call HandleEmptyCRevSeed
        
        Exit Sub
    End If
    
    If (BotVars.BNLS) Then
        Call modBNLS.SEND_BNLS_VERSIONCHECKEX2(ds.CRevFileTimeRaw, ds.CRevFileName, ds.CRevSeed)
    Else
        Call SEND_SID_REPORTVERSION
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_STARTVERSIONING (0x06) C->S
'*******************************
' (DWORD) Platform ID
' (DWORD) Product ID
' (DWORD) Version Byte
' (DWORD) Unknown (0)
'*******************************
Public Sub SEND_SID_STARTVERSIONING(Optional lVerByte As Long = 0)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    
    With pBuff
        .InsertDWord GetDWORDOverride(Config.PlatformID)                      'Platform ID
        .InsertDWord GetDWORD(BotVars.Product)                                'Product ID
        .InsertDWord IIf(lVerByte = 0, GetVerByte(BotVars.Product), lVerByte) 'VersionByte
        .InsertDWord 0  'Unknown
        .SendPacket SID_STARTVERSIONING
    End With
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
End Sub
'*******************************
'SID_REPORTVERSION (0x07) S->C
'*******************************
' (DWORD) Result
' (STRING) Patch path
'*******************************
Private Sub RECV_SID_REPORTVERSION(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lResult  As Long
    Dim sInfo    As String
    Dim bSuccess As Boolean

    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString
    bSuccess = False

    Select Case lResult
        Case 0: Call Event_VersionCheck(1, sInfo) 'Failed Version Check
        Case 1: Call Event_VersionCheck(1, sInfo) 'Old Game Version
        Case 2: 'Success
            bSuccess = True
            'Call Event_VersionCheck(0, sInfo)
        Case 3: Call Event_VersionCheck(1, sInfo) '"Reinstall Required", Invalid version
        Case Else:
            frmChat.AddChat RTBColors.ErrorMessageText, "Unknown SID_REPORTVERSION Response: 0x" & ZeroOffset(lResult, 8)
    End Select

    If Config.IgnoreVersionCheck Then bSuccess = True

    If (frmChat.sckBNet.State = sckConnected And bSuccess) Then
        If (GetCDKeyCount > 0) Then
            'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Sending CDKey information...")
            Select Case GetLogonSystem()
                Case BNCS_OLS: Call SEND_SID_CDKEY2
                Case BNCS_LLS: Call SEND_SID_CDKEY
                Case Else:
                    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem())
                    frmChat.AddChat RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error."
                    frmChat.DoDisconnect
            End Select
        Else
            Call Event_VersionCheck(0, sInfo) ' display success here

            ds.AccountEntry = True
            ds.AccountEntryPending = False
            frmChat.tmrIdleTimer.Enabled = True

            Call DoAccountAction
        End If
    Else
        frmChat.DoDisconnect
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_REPORTVERSION()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_REPORTVERSION (0x07) C->S
'*******************************
' (DWORD) Platform ID
' (DWORD) Product ID
' (DWORD) Version Byte
' (DWORD) EXE Version
' (DWORD) EXE Hash
' (STRING) EXE Information
'*******************************
Public Sub SEND_SID_REPORTVERSION(Optional lVerByte As Long = 0)
On Error GoTo ERROR_HANDLER:

    If (Not BotVars.BNLS) Then
        If (Not CompileCheckrevision()) Then
            frmChat.DoDisconnect
            Exit Sub
        End If
    End If
        
    If (ds.CRevChecksum = 0 Or ds.CRevVersion = 0 Or LenB(ds.CRevResult) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Check Revision Failed, sanity failed"
        frmChat.DoDisconnect
        Exit Sub
    End If
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertDWord GetDWORDOverride(Config.PlatformID)                      'Platform ID
        .InsertDWord GetDWORD(BotVars.Product)                                'Product ID
        .InsertDWord IIf(lVerByte = 0, GetVerByte(BotVars.Product), lVerByte) 'VersionByte
        .InsertDWord ds.CRevVersion                                           'Exe Version
        .InsertDWord ds.CRevChecksum                                          'Checksum
        .InsertNTString ds.CRevResult                                         'Result
        .SendPacket SID_REPORTVERSION
    End With
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_STARTVERSIONING()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_ENTERCHAT (0x0A) S->C
'*******************************
' (String) Unique name
' (String) Statstring
' (String) Account name
'*******************************
Private Sub RECV_SID_ENTERCHAT(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Call Event_LoggedOnAs(pBuff.GetString, pBuff.GetString, pBuff.GetString)

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_ENTERCHAT (0x0A) C->S
'*******************************
' (STRING) Username *
' (STRING) Statstring **
'*******************************
Private Sub SEND_SID_ENTERCHAT()
On Error GoTo ERROR_HANDLER:
    Dim pBuff As New clsDataBuffer
    pBuff.InsertNTString BotVars.Username
    pBuff.InsertNTString Config.CustomStatstring
    pBuff.SendPacket SID_ENTERCHAT
    Set pBuff = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_LEAVECHAT (0x10) C->S
'*******************************
' [blank]
'*******************************
Public Sub SEND_SID_LEAVECHAT()
On Error GoTo ERROR_HANDLER:
    Dim pBuff As New clsDataBuffer
    pBuff.SendPacket SID_LEAVECHAT
    Set pBuff = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_ENTERCHAT()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_GETCHANNELLIST (0x0B) S->C
'*******************************
' (String[]) Channels
'*******************************
Private Sub RECV_SID_GETCHANNELLIST(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Dim sChannels() As String
    Dim sChannel As String
    
    sChannels() = Split(vbNullString)
    
    Do
        sChannel = pBuff.GetString(UTF8)
        
        If LenB(sChannel) > 0 Then
            ReDim Preserve sChannels(UBound(sChannels) + 1)
            sChannels(UBound(sChannels)) = Trim$(sChannel)
        End If
    Loop While LenB(sChannel) > 0
    
    Call Event_ChannelList(sChannels)
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_GETCHANNELLIST()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_GETCHANNELLIST (0x0B) S->C
'*******************************
' (DWORD) Product ID
'*******************************
Private Sub SEND_SID_GETCHANNELLIST()
On Error GoTo ERROR_HANDLER:
    Dim pBuff As New clsDataBuffer
    pBuff.InsertDWord GetDWORD(BotVars.Product)
    pBuff.SendPacket SID_GETCHANNELLIST
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_GETCHANNELLIST()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CHATCOMMAND (0x0E) S->C
'*******************************
' (STRING) Text
'*******************************
Public Sub SEND_SID_CHATCOMMAND(sText As String)
On Error GoTo ERROR_HANDLER:

    If (LenB(sText) = 0) Then Exit Sub
    
    Dim pBuff As New clsDataBuffer
    pBuff.InsertNTString sText
    pBuff.SendPacket SID_CHATCOMMAND
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CHATCOMMAND()", Err.Number, Err.Description, OBJECT_NAME))
End Sub


'*******************************
'SID_CHATEVENT (0x0F) S->C
'*******************************
' (DWORD) Event ID
' (DWORD) User's Flags
' (DWORD) Ping
' (DWORD) IP Address (Defunct)
' (DWORD) Account number (Defunct)
' (DWORD) Registration Authority (Defunct)
' (STRING) Username
' (STRING) Text
'*******************************
Private Sub RECV_SID_CHATEVENT(pBuff As clsDataBuffer)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim EventID   As Long
    Dim lFlags    As Long
    Dim lPing     As Long
    Dim sUsername As String
    Dim sText     As String
    Dim Encoding  As STRINGENCODING
    
    Dim sProduct As String
    Dim sParsed  As String
    Dim sClanTag As String
    Dim sW3Icon  As String
    
    EventID = pBuff.GetDWORD
    lFlags = pBuff.GetDWORD
    lPing = pBuff.GetDWORD
    pBuff.GetDWORD                  'IP Address
    pBuff.GetDWORD                  'Account Number
    pBuff.GetDWORD                  'Reg Auth
    
    Select Case EventID
        Case ID_JOIN, ID_LEAVE, ID_USER, ID_USERFLAGS: ' user events: always encode statstring ANSI
            Encoding = ANSI
        Case Else
            If (frmChat.mnuUTF8.Checked) Then
                Encoding = UTF8
            Else
                Encoding = ANSI
            End If
    End Select
    
    sUsername = pBuff.GetString(Encoding)
    sText = pBuff.GetString(Encoding)
    
    Select Case EventID
        Case ID_JOIN:        Call Event_UserJoins(sUsername, lFlags, sText, lPing)
        Case ID_LEAVE:       Call Event_UserLeaves(sUsername, lFlags)
        Case ID_USER:        Call Event_UserInChannel(sUsername, lFlags, sText, lPing)
        Case ID_WHISPER:     Call Event_WhisperFromUser(sUsername, lFlags, sText, lPing)
        Case ID_TALK:        Call Event_UserTalk(sUsername, lFlags, sText, lPing)
        Case ID_BROADCAST:   Call Event_ServerInfo(sUsername, StringFormat("BROADCAST from {0}: {1}", sUsername, sText))
        Case ID_CHANNEL:     Call Event_JoinedChannel(sText, lFlags)
        Case ID_USERFLAGS:   Call Event_FlagsUpdate(sUsername, lFlags, sText, lPing)
        Case ID_WHISPERSENT: Call Event_WhisperToUser(sUsername, lFlags, sText, lPing)
        Case ID_CHANNELFULL, ID_CHANNELDOESNOTEXIST, ID_CHANNELRESTRICTED: Call Event_ChannelJoinError(EventID, sText)
        Case ID_INFO:        Call Event_ServerInfo(sUsername, sText)
        Case ID_ERROR:       Call Event_ServerError(sText)
        Case ID_EMOTE:       Call Event_UserEmote(sUsername, lFlags, sText)
        Case Else:
            If MDebug("debug") Then
                Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Unhandled SID_CHATEVENT Event: 0x{0}", ZeroOffset(EventID, 8)))
                Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Packet data: {0}{1}", vbNewLine, pBuff.DebugOutput))
            End If
    End Select
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CHATEVENT()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*********************************
' SID_LOCALEINFO (0x12) C->S
'*********************************
' (FILETIME) System time
' (FILETIME) Local time
' (DWORD) Timezone bias
' (DWORD) SystemDefaultLCID
' (DWORD) UserDefaultLCID
' (DWORD) UserDefaultLangID
' (STRING) Abbreviated language name
' (STRING) Country Code
' (STRING) Abbreviated country name
' (STRING) Country Name
'*********************************
Public Sub SEND_SID_LOCALEINFO()
On Error GoTo ERROR_HANDLER:
    Const LOCALE_SABBREVLANGNAME As Long = &H3
    Const LOCALE_USER_DEFAULT    As Long = &H400
    Dim LanguageAbr As String
    Dim CountryCode As String
    Dim CountryAbr  As String
    Dim CountryName As String
    Dim lRet        As String
    
    Dim st As SYSTEMTIME
    Dim ft As FILETIME
    
    Dim pBuff As New clsDataBuffer
    
    LanguageAbr = String$(256, 0)
    Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVLANGNAME, LanguageAbr, Len(LanguageAbr))
    LanguageAbr = KillNull(LanguageAbr)
    
    Call GetCountryData(CountryAbr, CountryName, CountryCode)
    If (Len(LanguageAbr) = 0) Then LanguageAbr = "ENU"
    If (Len(CountryCode) = 0) Then CountryCode = "1"
    If (Not Len(CountryAbr) = 3) Then CountryAbr = "USA"
    If (LenB(CountryName) = 0) Then CountryName = "United States"
    
    With pBuff
        Call GetSystemTime(st)
        Call SystemTimeToFileTime(st, ft)
        .InsertDWord ft.dwLowDateTime                                 'SystemTime
        .InsertDWord ft.dwHighDateTime                                'SystemTime
        
        Call GetLocalTime(st)
        Call SystemTimeToFileTime(st, ft)
        .InsertDWord ft.dwLowDateTime                                 'LocalTime
        .InsertDWord ft.dwHighDateTime                                'LocalTime
        
        .InsertDWord GetTimeZoneBias                                  'Time Zone Bias
        If Config.ForceDefaultLocaleID Then
            .InsertDWord 1033                                         'SystemDefaultLCID
            .InsertDWord 1033                                         'UserDefaultLCID
            .InsertDWord 1033                                         'UserDefaultLangID
        Else
            .InsertDWord CLng(GetSystemDefaultLCID)                   'SystemDefaultLCID
            .InsertDWord CLng(GetUserDefaultLCID)                     'UserDefaultLCID
            .InsertDWord CLng(GetUserDefaultLangID)                   'UserDefaultLangID
        End If
        
        .InsertNTString LanguageAbr                                   'Language Abbrev
        .InsertNTString CountryCode                                   'Country Code
        .InsertNTString CountryAbr                                    'Country Abbrev
        .InsertNTString CountryName                                   'Country Name
        
        .SendPacket SID_LOCALEINFO
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOCALEINFO()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_UDPPINGRESPONSE (0x14) C->S
'*******************************
' (DWORD) UDP value
'*******************************
Private Sub SEND_SID_UDPPINGRESPONSE()
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer

    pBuff.InsertDWord GetDWORDOverride(Config.UDPString, &H626E6574)    'default: bnet
    pBuff.SendPacket SID_UDPPINGRESPONSE
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_UDPPINGRESPONSE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_MESSAGEBOX (0x19) S->C
'*******************************
' (DWORD) Style
' (String) Text
' (String) Caption
'*******************************
Private Sub RECV_SID_MESSAGEBOX(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Call Event_MessageBox(pBuff.GetDWORD, pBuff.GetString(UTF8), pBuff.GetString(UTF8))
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_MESSAGEBOX()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'********************************
'SID_LOGONCHALLENGEEX (0x1D) S->C
'********************************
' (DWORD) UDP Token
' (DWORD) Server Token
'********************************
Private Sub RECV_SID_LOGONCHALLENGEEX(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    ds.UDPValue = pBuff.GetDWORD
    ds.ServerToken = pBuff.GetDWORD
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONCHALLENGEEX()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*********************************
' SID_CLIENTID2 (0x1E) C->S
'*********************************
' (DWORD) Server Version
' For server version 1:
'  (DWORD) Registration Version
'  (DWORD) Registration Authority
' For server version 0:
'  (DWORD) Registration Authority
'  (DWORD) Registration Version
' (DWORD) Account Number
' (DWORD) Registration Token
' (STRING) LAN computer name
' (STRING) LAN username
'*********************************
'This is eww, I don't like hard coding,
'but to get this crap I would need to
'use Storm.dll, which we don't want to
'distribute with the bot.
'*********************************
Public Sub SEND_SID_CLIENTID2()
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertDWord 1
        .InsertDWord 0
        .InsertDWord 0
        .InsertDWord 0
        .InsertDWord 0
        .InsertNTString GetComputerLanName
        .InsertNTString GetComputerUsername
        .SendPacket SID_CLIENTID2
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CLIENTID2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_PING (0x25) S->C
'*******************************
' (DWORD) Ping value
'*******************************
Private Sub RECV_SID_PING(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Dim SendResponse As Boolean
    Dim Cookie As Long

    Cookie = pBuff.GetDWORD
    
    SendResponse = False
    If GetTickCountMS() >= ds.LastPingResponse + 1000 Then
        SendResponse = True
        ds.LastPingResponse = GetTickCountMS()
    End If
    'frmChat.AddChat vbWhite, StringFormat("PING uTicks={0} LPR={1} C={2} SendResponse={3}", uTicks, ds.LastPingResponse, cookie, SendResponse)

    If (frmChat.tmrIdleTimer.Enabled) Then
        ' reached account entry/idle timer enabled
        If (SendResponse) Then
            Call SEND_SID_PING(Cookie)
        End If
    Else
        ' during initial auth
        If (BotVars.Spoof = 0) Then
            Call SEND_SID_PING(Cookie)
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_PING()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_PING (0x25) C->S
'*******************************
' (DWORD) Ping value
'*******************************
Private Sub SEND_SID_PING(ByVal lPingValue As Long)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    
    SetNagelStatus frmChat.sckBNet.SocketHandle, False
    
    pBuff.InsertDWord lPingValue
    pBuff.SendPacket SID_PING
    
    SetNagelStatus frmChat.sckBNet.SocketHandle, True
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_PING()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_READUSERDATA (0x26) S->C
'*******************************
' (DWORD) Number of accounts
' (DWORD) Number of keys
' (DWORD) Request ID
' (STRING[]) Requested key values
'*******************************
Private Sub RECV_SID_READUSERDATA(pBuff As clsDataBuffer)
    On Error GoTo ERROR_HANDLER:

    Dim i As Integer
    
    Dim iNumKeys As Long
    Dim iRequest As Long
    Dim aValues() As String

    pBuff.GetDWORD                  ' (DWORD) Number of accounts
    iNumKeys = pBuff.GetDWORD()     ' (DWORD) Number of keys
    iRequest = pBuff.GetDWORD()     ' (DWORD) Request ID

    If iNumKeys < 1 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Notice: Received user data request with no returned keys. Cookie: " & CStr(iRequest)
    Else
        ReDim aValues(iNumKeys - 1)
    
        ' Read each of the keys
        For i = 0 To UBound(aValues)
            aValues(i) = pBuff.GetString(IIf(frmChat.mnuUTF8.Checked, UTF8, ANSI))
        Next i
    End If
    
    ' Find the request for this ID and hand it off to the event handler
    i = UBound(UserDataRequests)
    If i >= iRequest Then
    
        ' Process the request
        With UserDataRequests(iRequest)
            If .ResponseReceived Then
                frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Notice: Received extra data response for user: {0}, # of keys: {1}", .Account, iNumKeys)
            End If
        
            .ResponseReceived = True            ' Flag this request as received
            .Values = aValues                   ' Link the values
        End With
        
        ' Raise UserDataReceived event (also raises KeyReturn in scripting)
        Event_UserDataReceived UserDataRequests(iRequest)
        
        ' Shrink the array if needed
        If i > 1 Then
            For i = i To 1 Step -1
                If UserDataRequests(i).ResponseReceived Then
                    ReDim Preserve UserDataRequests(i - 1)
                Else
                    Exit For
                End If
            Next
        End If
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Notice: Received unsolicited user data, # of keys: {0}, Cookie: {1}", iNumKeys, iRequest)
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_READUSERDATA()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'********************************
'SID_LOGONCHALLENGE (0x28) S->C
'********************************
' (DWORD) Server Token
'********************************
Private Sub RECV_SID_LOGONCHALLENGE(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    ds.ServerToken = pBuff.GetDWORD
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONCHALLENGE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CDKEY (0x30) S->C
'*******************************
' (DWORD) Result
' (STRING) Key owner
'*******************************
Private Sub RECV_SID_CDKEY(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lResult As Long
    Dim sInfo   As String
    
    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString(UTF8)

    Select Case lResult
        Case 1:
            Call Event_VersionCheck(0, sInfo) ' display success here

            ds.AccountEntry = True
            ds.AccountEntryPending = False
            frmChat.tmrIdleTimer.Enabled = True

            Call DoAccountAction

            Exit Sub
        Case 2: Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
        Case 3: Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
        Case 4: Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
        Case 5: Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
        Case Else: frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown SID_CDKEY Response 0x{0}: {1}", ZeroOffset(lResult, 8), sInfo)
    End Select

    'Call frmChat.DoDisconnect

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CDKEY()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CDKEY (0x30) C->S
'*******************************
' (DWORD) Spawn (0/1)
' (STRING) Key
' (STRING) Key owner
'*******************************
Public Sub SEND_SID_CDKEY()
On Error GoTo ERROR_HANDLER:
    Dim oKey As New clsKeyDecoder
    Dim pBuff As New clsDataBuffer
    
    oKey.Initialize BotVars.CDKey
    If Not oKey.IsValid Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid."
        frmChat.DoDisconnect
        Exit Sub
    End If
    
    With pBuff
        .InsertBool (CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
        .InsertNTString BotVars.CDKey
        
        If (LenB(Config.CDKeyOwnerName) > 0) Then
            .InsertNTString Config.CDKeyOwnerName
        Else
            .InsertNTString Config.Username
        End If
        .SendPacket SID_CDKEY
    End With
    
    Set pBuff = Nothing
    Set oKey = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CDKEY()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_CHANGEPASSWORD (0x31) S->C
'**********************************
' (DWORD) Status
'**********************************
Private Sub RECV_SID_CHANGEPASSWORD(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    
    lResult = pBuff.GetDWORD

    ds.AccountEntryPending = False
    frmChat.tmrAccountLock.Enabled = False

    Select Case lResult
        Case &H0:
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, &H0, vbNullString)

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_LOGON)
            Else
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            End If

        Case Else
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, lResult + &H3100, vbNullString)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_CHPWD
            Else
                frmChat.DoDisconnect
            End If
    End Select

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CHANGEPASSWORD()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CHANGEPASSWORD (0x31) C->S
'*******************************
' (DWORD) Client Token
' (DWORD) Server Token
' (DWORD) [5] Old Password Hash
' (DWORD) [5] New Password Hash
' (STRING) Username
'*******************************
Public Sub SEND_SID_CHANGEPASSWORD()
On Error GoTo ERROR_HANDLER:
    Dim sHash As String
    Dim sHash2 As String
    Dim pBuff As New clsDataBuffer

    frmChat.tmrAccountLock.Enabled = True
    frmChat.tmrAccountLock.Tag = ACCOUNT_MODE_CHPWD
    If Not Config.UseLowerCasePassword Then
        sHash = DoubleHashPassword(Config.Password, ds.ClientToken, ds.ServerToken)
        sHash2 = DoubleHashPassword(Config.NewPassword, ds.ClientToken, ds.ServerToken)
    Else
        sHash = DoubleHashPassword(LCase$(Config.Password), ds.ClientToken, ds.ServerToken)
        sHash2 = DoubleHashPassword(LCase$(Config.NewPassword), ds.ClientToken, ds.ServerToken)
    End If

    With pBuff
        .InsertDWord ds.ClientToken
        .InsertDWord ds.ServerToken
        .InsertNonNTString sHash
        .InsertNonNTString sHash2
        .InsertNTString Config.Username
        .SendPacket SID_CHANGEPASSWORD
    End With

    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CHANGEPASSWORD()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CDKEY2 (0x36) S->C
'*******************************
' (DWORD) Result
' (STRING) Key owner
'*******************************
Private Sub RECV_SID_CDKEY2(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lResult As Long
    Dim sInfo   As String
    
    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString(UTF8)
    
    Select Case lResult
        Case 1:
            Call Event_VersionCheck(0, sInfo) ' display success here

            ds.AccountEntry = True
            ds.AccountEntryPending = False
            frmChat.tmrIdleTimer.Enabled = True

            Call DoAccountAction

            Exit Sub
        Case 2: Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
        Case 3: Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
        Case 4: Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
        Case 5: Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
        Case Else: frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown SID_CDKEY2 Response 0x{0}: {1}", ZeroOffset(lResult, 8), sInfo)
    End Select

    'Call frmChat.DoDisconnect
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CDKEY2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CDKEY2 (0x36) C->S
'*******************************
' (DWORD) Spawn (0/1)
' (DWORD) Key Length
' (DWORD) CDKey Product
' (DWORD) CDKey Value1
' (DWORD) Server Token
' (DWORD) Client Token
' (DWORD) [5] Hashed Data
' (STRING) Key owner
'*******************************
Public Sub SEND_SID_CDKEY2()
On Error GoTo ERROR_HANDLER:
    Dim oKey     As New clsKeyDecoder
    Dim pBuff As New clsDataBuffer
    
    oKey.Initialize BotVars.CDKey
    If Not oKey.IsValid Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid."
        frmChat.DoDisconnect
        Exit Sub
    End If

    If Not oKey.CalculateHash(ds.ClientToken, ds.ServerToken, BNCS_OLS) Then Exit Sub
    
    With pBuff
        .InsertBool (CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
        .InsertDWord oKey.KeyLength
        .InsertDWord oKey.ProductValue
        .InsertDWord oKey.PublicValue
        .InsertDWord ds.ServerToken
        .InsertDWord ds.ClientToken
        .InsertNonNTString oKey.Hash
        
        If (LenB(Config.CDKeyOwnerName) > 0) Then
            .InsertNTString Config.CDKeyOwnerName
        Else
            .InsertNTString Config.Username
        End If
        .SendPacket SID_CDKEY2
    End With

    Set pBuff = Nothing
    Set oKey = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CDKEY2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub


'*******************************
'SID_LOGONRESPONSE2 (0x3A) S->C
'*******************************
' (DWORD) Result
' (STRING) Reason
'*******************************
Private Sub RECV_SID_LOGONRESPONSE2(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lResult As Long
    Dim sInfo  As String
    
    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString(UTF8)

    ds.AccountEntryPending = False
    frmChat.tmrAccountLock.Enabled = False

    Select Case lResult
        Case &H0  'Logon Successful
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, &H0, sInfo)

            BotVars.Username = Config.Username
            BotVars.Password = Config.Password
            ds.AccountEntry = False

            If frmAccountManager.Visible Then
                frmAccountManager.LeftAccountEntryMode
            End If

            If (Not ds.WaitingForEmail) Then
                If (Dii And BotVars.UseRealm) Then
                    'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Asking Battle.net for a list of Realm servers...")
                    Call DoQueryRealms
                Else
                    Call SendEnterChatSequence
                End If
            Else
                Call DoRegisterEmail
            End If

        Case &H1  'Nonexistent account.
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, &H1, sInfo)

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_CREAT)
            ElseIf Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_CREAT
            Else
                frmChat.DoDisconnect
            End If

        Case Else
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, lResult, sInfo)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            Else
                frmChat.DoDisconnect
            End If
    End Select
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONRESPONSE2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_LOGONRESPONSE2 (0x3A) C->S
'*******************************
' (DWORD) Client Token
' (DWORD) Server Token
' (DWORD) [5] Password Hash
' (STRING) Username
'*******************************
Public Sub SEND_SID_LOGONRESPONSE2()
On Error GoTo ERROR_HANDLER:
    Dim sHash As String
    Dim pBuff As New clsDataBuffer

    frmChat.tmrAccountLock.Enabled = True
    frmChat.tmrAccountLock.Tag = ACCOUNT_MODE_LOGON
    If Not Config.UseLowerCasePassword Then
        sHash = DoubleHashPassword(Config.Password, ds.ClientToken, ds.ServerToken)
    Else
        sHash = DoubleHashPassword(LCase$(Config.Password), ds.ClientToken, ds.ServerToken)
    End If
    
    With pBuff
        .InsertDWord ds.ClientToken
        .InsertDWord ds.ServerToken
        .InsertNonNTString sHash
        .InsertNTString Config.Username
        .SendPacket SID_LOGONRESPONSE2
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOGONRESPONSE2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_CREATEACCOUNT2 (0x3D) S->C
'*******************************
' (DWORD) Status
' (STRING) Account name suggestion
'*******************************
Private Sub RECV_SID_CREATEACCOUNT2(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Dim lResult As Long
    Dim sInfo   As String
    Dim sOut    As String
    
    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString(UTF8)

    ds.AccountEntryPending = False

    Select Case lResult
        Case &H0:
            Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H0, sInfo)

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_LOGON)
            Else
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            End If

            Exit Sub

        Case &H1:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H7, sInfo)
        Case &H2:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H8, sInfo)
        Case &H3:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H9, sInfo)
        Case &H4:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H4, sInfo)
        Case &H5:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H3D05, sInfo)
        Case &H6:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &HA, sInfo)
        Case &H7:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &HB, sInfo)
        Case &H8:  Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &HC, sInfo)
        Case Else: Call Event_LogonEvent(ACCOUNT_MODE_CREAT, lResult, sInfo)
    End Select

    If Config.ManageOnAccountError Then
        frmAccountManager.ShowMode ACCOUNT_MODE_CREAT
    Else
        frmChat.DoDisconnect
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_CREATEACCOUNT2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_CREATEACCOUNT2 (0x3D) C->S
'**************************************
' (DWORD) [5] Password hash
' (STRING) Username
'**************************************
Public Sub SEND_SID_CREATEACCOUNT2()
On Error GoTo ERROR_HANDLER:
    
    Dim sHash As String
    If Not Config.UseLowerCasePassword Then
        sHash = HashPassword(Config.Password)
    Else
        sHash = HashPassword(LCase$(Config.Password))
    End If
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNonNTString sHash
        .InsertNTString Config.Username
        .SendPacket SID_CREATEACCOUNT2
    End With
    Set pBuff = Nothing
        
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CREATEACCOUNT2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_LOGONREALMEX (0x3E) S->C
'*******************************
' (DWORD) MCP Cookie
' (DWORD) MCP Status
' (DWORD) [2] MCP Chunk 1
' (DWORD) IP
' (DWORD) Port
' (DWORD) [12] MCP Chunk 2
' (STRING) Battle.net unique name
'*******************************
Private Sub RECV_SID_LOGONREALMEX(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lError   As Long
    Dim sMCPData As String
    Dim sTitle   As String
    Dim ptrIP    As Long
    Dim sIP      As String
    Dim lPort    As Long
    Dim sUniq    As String
    Dim x        As Integer

    If (Len(pBuff.GetRaw(, True)) > 8) Then
        sMCPData = pBuff.GetRaw(16)
        
        sIP = GetAddressFromLong(pBuff.GetDWORD)
        
        lPort = ntohs(pBuff.GetDWORD)
        
        sMCPData = StringFormat("{0}{1}", sMCPData, pBuff.GetRaw(48))
        sUniq = pBuff.GetString(UTF8)
        
        If (Not frmChat.sckMCP.State = 0) Then frmChat.sckMCP.Close
        
        If Not ds.MCPHandler Is Nothing Then
            Call ds.MCPHandler.SetStartupData(sMCPData, sUniq, sIP, lPort)
            sTitle = ds.MCPHandler.RealmServerTitle(ds.MCPHandler.RealmServerSelectedIndex)
            
            If (ProxyConnInfo(stMCP).IsUsingProxy) Then
                frmChat.AddChat RTBColors.InformationText, "[REALM] [PROXY] Connecting to the SOCKS" & ProxyConnInfo(stMCP).Version & " proxy server at " & ProxyConnInfo(stMCP).ProxyIP & ":" & ProxyConnInfo(stMCP).ProxyPort & "..."
            Else
                frmChat.AddChat RTBColors.InformationText, StringFormat("[REALM] Connecting to the Diablo II Realm {0} at {1}:{2}...", sTitle, sIP, lPort)
            End If
            
            With frmChat.sckMCP
                If (ProxyConnInfo(stMCP).IsUsingProxy) Then
                    .RemoteHost = ProxyConnInfo(stMCP).ProxyIP
                    .RemotePort = ProxyConnInfo(stMCP).ProxyPort
                Else
                    .RemoteHost = sIP
                    .RemotePort = lPort
                End If
                .Connect
            End With
        End If
    Else
        pBuff.GetDWORD
        lError = pBuff.GetDWORD
        
        Select Case lError
            Case &H80000001: frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] The Diablo II Realm is currently unavailable. Please try again later."
            Case &H80000002: frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Diablo II Realm logon has failed. Please try again later."
            Case Else:       frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[REALM] Logon to the Diablo II Realm has failed for an unknown reason (0x{0}). Please try again later.", ZeroOffset(lError, 8))
        End Select
        
        If Not ds.MCPHandler Is Nothing Then
            If ds.MCPHandler.FormActive Then
                frmRealm.UnloadRealmError
            End If
        End If
        
        Call SendEnterChatSequence
        frmChat.mnuRealmSwitch.Enabled = True
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_LOGONREALMEX()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_LOGONREALMEX (0x3E) C->S
'*******************************
' (DWORD) Client Token
' (DWORD) [5] Hashed realm password
' (STRING) Realm title
'*******************************
Public Sub SEND_SID_LOGONREALMEX(sRealmTitle As String, sRealmServerPassword As String)
On Error GoTo ERROR_HANDLER:
    
    If (LenB(sRealmTitle) = 0) Then Exit Sub
    
    Dim pBuff As New clsDataBuffer
    pBuff.InsertDWord ds.ClientToken
    pBuff.InsertNonNTString DoubleHashPassword(sRealmServerPassword, ds.ClientToken, ds.ServerToken)
    pBuff.InsertNTString sRealmTitle
    pBuff.SendPacket SID_LOGONREALMEX
    Set pBuff = Nothing
    
    'Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Logging on to the Diablo II Realm...")
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_LOGONREALMEX()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_QUERYREALMS2 (0x40) S->C
'*******************************
' (DWORD) Unknown
' (DWORD) Count
' For Each Realm:
'   (DWORD) Unknown
'   (STRING) Realm title
'   (STRING) Realm description
'*******************************
Private Sub RECV_SID_QUERYREALMS2(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lCount      As Long
    Dim i           As Integer
    Dim List()      As Variant
    Dim Server(0 To 1) As String
    
    pBuff.GetDWORD 'Unknown
    lCount = pBuff.GetDWORD
    
    If (MDebug("debug") And (MDebug("all") Or MDebug("info"))) Then
        frmChat.AddChat RTBColors.InformationText, "Received Realm List:"
    End If
    
    If lCount > 0 Then
        ReDim List(lCount - 1)
        
        For i = 0 To lCount - 1
            pBuff.GetDWORD 'Unknown
            
            Server(0) = pBuff.GetString(UTF8)
            Server(1) = pBuff.GetString(UTF8)
            List(i) = Server()
            
            If (MDebug("debug") And (MDebug("all") Or MDebug("info"))) Then
                frmChat.AddChat RTBColors.InformationText, StringFormat("  {0}: {1}", Server(0), Server(1))
            End If
        Next i
        
        If Not ds.MCPHandler Is Nothing Then
            Call ds.MCPHandler.SetRealmServerInfo(List)
        End If
    End If
    
    'Call frmChat.AddChat(RTBColors.SuccessText, "[BNCS] Received Realm list.")
    
    Call ds.MCPHandler.HandleQueryRealmServersResponse
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Recv_SID_QUERYREALMS2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_QUERYREALMS2 (0x40) C->S
'*******************************
' [Blank]
'*******************************
Public Sub SEND_SID_QUERYREALMS2()
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    pBuff.SendPacket SID_QUERYREALMS2
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_QUERYREALMS2()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_AUTH_INFO (0x50) S->C
'*******************************
' (DWORD) Logon Type
' (DWORD) Server Token
' (DWORD) UDPValue
' (FILETIME) CRev Archive File Time
' (STRING) CRev Archive File Name
' (STRING) CRev Seed Values
' WAR3/W3XP Only:
'   (VOID) 128-byte Server signature
'*******************************
Private Sub RECV_SID_AUTH_INFO(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim RemoteHostIP As String

    ds.LogonType = pBuff.GetDWORD
    ds.ServerToken = pBuff.GetDWORD
    ds.UDPValue = pBuff.GetDWORD
    ds.CRevFileTime = pBuff.GetRaw(8)
    ds.CRevFileName = pBuff.GetString
    ds.CRevSeed = pBuff.GetString
    ds.ServerSig = pBuff.GetRaw(128)
    
    Call frmChat.AddChat(RTBColors.InformationText, "[BNCS] Checking version...")
    
    If (MDebug("all") Or MDebug("crev")) Then
        frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Name: {0}", ds.CRevFileName)
        frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Time: {0}", ds.CRevFileTime)
        If (InStr(1, ds.CRevFileName, "lockdown", vbTextCompare) > 0) Then
            frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Seed: {0}", StrToHex(ds.CRevSeed))
        Else
            frmChat.AddChat RTBColors.InformationText, StringFormat("CRev Seed: {0}", ds.CRevSeed)
        End If
    End If
    
    ' If the server does not recognize the version byte we sent it, it will send back an empty seed string.
    If (LenB(ds.CRevSeed) = 0) Then
        Call HandleEmptyCRevSeed
        
        Exit Sub
    End If
    
    If (Len(ds.ServerSig) = 128) Then
        If (ProxyConnInfo(stBNCS).IsUsingProxy) Then
            RemoteHostIP = ProxyConnInfo(stBNCS).RemoteHostIP
        Else
            RemoteHostIP = frmChat.sckBNet.RemoteHostIP
        End If
        
        If (StrComp(RemoteHostIP, "0.0.0.255", vbBinaryCompare) = 0) Then
            frmChat.AddChat RTBColors.InformationText, "[BNCS] Note! A server signature was received but cannot be validated because of the proxy configuration."
        Else
            If (ds.NLS.VerifyServerSignature(RemoteHostIP, ds.ServerSig)) Then
                frmChat.AddChat RTBColors.SuccessText, "[BNCS] Server signature validated!"
            Else
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Warning! The server signature is invalid. This may not be a valid server."
            End If
        End If
    ElseIf (GetProductKey = "W3") Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Warning! The server signature is missing. This may not be a valid server."
    End If
    
    If (BotVars.BNLS) Then
        modBNLS.SEND_BNLS_VERSIONCHECKEX2 ds.CRevFileTimeRaw, ds.CRevFileName, ds.CRevSeed
    Else
        modBNCS.SEND_SID_AUTH_CHECK
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_INFO()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_AUTH_INFO (0x50) C->S
'*******************************
' (DWORD) Protocol ID (0)
' (DWORD) Platform ID
' (DWORD) Product ID
' (DWORD) Version Byte
' (DWORD) Product language
' (DWORD) Local IP for NAT compatibility*
' (DWORD) Time zone bias*
' (DWORD) Locale ID*
' (DWORD) Language ID*
' (STRING) Country abreviation
' (STRING) Country
'*******************************
Public Sub SEND_SID_AUTH_INFO(Optional lVerByte As Long = 0)
On Error GoTo ERROR_HANDLER:

    Dim LocalIP     As Long
    Dim CountryAbr  As String
    Dim CountryName As String
    
    Dim pBuff As New clsDataBuffer
    
    LocalIP = inet_addr(frmChat.sckBNet.LocalIP)

    Call GetCountryData(CountryAbr, CountryName, vbNull)
    If (Not Len(CountryAbr) = 3) Then CountryAbr = "USA"
    If (LenB(CountryName) = 0) Then CountryName = "United States"
    
    With pBuff
    
        .InsertDWord Config.ProtocolID                                        'ProtocolID
        .InsertDWord GetDWORDOverride(Config.PlatformID, PLATFORM_INTEL)      'Platform ID
        .InsertDWord GetDWORD(BotVars.Product)                                'Product ID
        .InsertDWord IIf(lVerByte = 0, GetVerByte(BotVars.Product), lVerByte) 'VersionByte
        .InsertDWord GetDWORDOverride(Config.ProductLanguage)                 'Product Language
        .InsertDWord LocalIP                                                  'Local IP
        .InsertDWord GetTimeZoneBias                                          'Time Zone Bias
        If Config.ForceDefaultLocaleID Then
            .InsertDWord 1033                                                 'LocalID
            .InsertDWord 1033                                                 'LangID
        Else
            .InsertDWord CLng(GetUserDefaultLCID)                             'LocalID
            .InsertDWord CLng(GetUserDefaultLangID)                           'LangID
        End If
        .InsertNTString CountryAbr                                            'Country abreviation
        .InsertNTString CountryName                                           'Country Name
        .SendPacket SID_AUTH_INFO
    End With
    
    Set pBuff = Nothing
    
    If (BotVars.Spoof = 1) Then
        Call SEND_SID_PING(0)
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_INFO()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_AUTH_CHECK (0x51) S->C
'*******************************
' (DWORD) Result
' (STRING) Additional Information
'*******************************
Private Sub RECV_SID_AUTH_CHECK(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    Dim lResult  As Long
    Dim sInfo    As String
    Dim bSuccess As Boolean
    
    lResult = pBuff.GetDWORD
    sInfo = pBuff.GetString(UTF8)

    bSuccess = False
    
    Select Case lResult
        Case &H0:
            Call Event_VersionCheck(0, sInfo)
            bSuccess = True

        Case &H100, &H101: Call Event_VersionCheck(1, sInfo) 'Outdated/Invalid Version
        Case &H200: Call Event_VersionCheck(2, sInfo) 'Invalid CDKey
        Case &H201: Call Event_VersionCheck(6, sInfo) 'CDKey is In Use
        Case &H202: Call Event_VersionCheck(5, sInfo) 'CDKey is Banned
        Case &H203: Call Event_VersionCheck(4, sInfo) 'CDKey is for the wrong product
        Case &H210: Call Event_VersionCheck(7, sInfo) 'Invalid Exp CDKey
        Case &H211: Call Event_VersionCheck(8, sInfo) 'Exp CDKey is In Use
        Case &H212: Call Event_VersionCheck(9, sInfo) 'Exp CDKey is Banned
        Case &H213: Call Event_VersionCheck(10, sInfo) 'Exp CDKey is for the wrong product
        Case Else:
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unknown 0x51 Response: 0x" & ZeroOffset(lResult, 8))
    End Select

    If Config.IgnoreVersionCheck Then bSuccess = True

    If (frmChat.sckBNet.State = sckConnected And bSuccess) Then
        ds.AccountEntry = True
        ds.AccountEntryPending = False
        frmChat.tmrIdleTimer.Enabled = True

        Call DoAccountAction

        Exit Sub
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_CHECK()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_AUTH_CHECK (0x51) C->S
'*******************************
' (DWORD) Client Token
' (DWORD) EXE Version
' (DWORD) EXE Hash
' (DWORD) Number of CD-keys in this packet
' (BOOLEAN) Spawn CD-key
' For Each Key:
'   (DWORD) Key Length
'   (DWORD) CD-key's product value
'   (DWORD) CD-key's public value
'   (DWORD) Unknown (0)
'   (DWORD) [5] Hashed Key Data
' (STRING) Exe Information
' (STRING) CD-Key owner name
'*******************************
Public Sub SEND_SID_AUTH_CHECK()
On Error GoTo ERROR_HANDLER:
    
    Dim pBuff    As New clsDataBuffer
    Dim i        As Long
    Dim keys     As Long
    Dim sKey     As String
    Dim oKey     As New clsKeyDecoder
    
    If (Not BotVars.BNLS) Then
        If (Not CompileCheckrevision()) Then
            frmChat.DoDisconnect
            Exit Sub
        End If
    End If
        
    If (ds.CRevChecksum = 0 Or ds.CRevVersion = 0 Or LenB(ds.CRevResult) = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Check Revision Failed, sanity failed"
        frmChat.DoDisconnect
        Exit Sub
    End If
    
    keys = GetCDKeyCount
    
    With pBuff
        .InsertDWord ds.ClientToken  'Client Token
        .InsertDWord ds.CRevVersion  'CRev Version
        .InsertDWord ds.CRevChecksum 'CRev Checksum
        .InsertDWord keys            'CDKey Count
        .InsertBool (CanSpawn(BotVars.Product, oKey.KeyLength) And Config.UseSpawn)
        
        For i = 1 To keys
            If (i = 1) Then
                sKey = BotVars.CDKey
            ElseIf (i = 2) Then
                sKey = BotVars.ExpKey
            Else
                sKey = ReadCfg$("Main", StringFormat("CDKey{0}", i))
            End If
            
            'Initialize the key decoder and validate the key.
            oKey.Initialize sKey
            If Not oKey.IsValid Then
                frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid."
                frmChat.DoDisconnect
                Exit Sub
            End If
            
            'Calculate the hash
            If Not oKey.CalculateHash(ds.ClientToken, ds.ServerToken, BNCS_NLS) Then Exit Sub
            
            .InsertDWord oKey.KeyLength
            .InsertDWord oKey.ProductValue
            .InsertDWord oKey.PublicValue
            .InsertDWord 0&
            .InsertNonNTString oKey.Hash
        Next i
        
        .InsertNTString ds.CRevResult
        If (LenB(Config.CDKeyOwnerName) > 0) Then
            .InsertNTString Config.CDKeyOwnerName
        Else
            .InsertNTString Config.Username
        End If
        
        .SendPacket SID_AUTH_CHECK
    End With

    Set pBuff = Nothing
    Set oKey = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_CHECK()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCREATE (0x52) S->C
'**********************************
' (DWORD) Status
'**********************************
Private Sub RECV_SID_AUTH_ACCOUNTCREATE(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    
    lResult = pBuff.GetDWORD

    ds.AccountEntryPending = False

    Select Case lResult
        Case &H0:
            Call Event_LogonEvent(ACCOUNT_MODE_CREAT, &H0, vbNullString)

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_LOGON)
            Else
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            End If

        Case Else
            Call Event_LogonEvent(ACCOUNT_MODE_CREAT, lResult, vbNullString)
    End Select

    If Config.ManageOnAccountError Then
        frmAccountManager.ShowMode ACCOUNT_MODE_CREAT
    Else
        frmChat.DoDisconnect
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTCREATE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCREATE (0x52) C->S
'**********************************
' (BYTE[32]) Salt (s)
' (BYTE[32]) Verifier (v)
' (STRING) Username
'**********************************
Public Sub SEND_SID_AUTH_ACCOUNTCREATE()
On Error GoTo ERROR_HANDLER:

    Dim oNLS As New clsNLS

    Call oNLS.Initialize(Config.Username, Config.Password)
    Call oNLS.GenerateSaltAndVerifier

    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNonNTString oNLS.Srp_Salt
        .InsertNonNTString oNLS.Srp_v
        .InsertNTString oNLS.Username
        .SendPacket SID_AUTH_ACCOUNTCREATE
    End With
    Set pBuff = Nothing

    Call oNLS.Terminate
    Set oNLS = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTCREATE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTLOGON (0x53) S->C
'**********************************
' (DWORD) Status
' (BYTE[32]) Salt (s)
' (BYTE[32]) Server Key (B)
'**********************************
Private Sub RECV_SID_AUTH_ACCOUNTLOGON(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    
    lResult = pBuff.GetDWORD
    ds.NLS.Srp_Salt = pBuff.GetRaw(32)
    ds.NLS.Srp_B = pBuff.GetRaw(32)
    
    Select Case lResult
        Case &H0: 'Accepted, requires proof.
            SEND_SID_AUTH_ACCOUNTLOGONPROOF
            Exit Sub

        Case &H1: 'Account doesn't exist.
            ds.AccountEntryPending = False
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, &H1, vbNullString)

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_CREAT)
            ElseIf Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_CREAT
            Else
                frmChat.DoDisconnect
            End If

        Case Else
            ds.AccountEntryPending = False
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, lResult, vbNullString)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            Else
                frmChat.DoDisconnect
            End If
    End Select

    Call ds.NLS.Terminate
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTLOGON()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTLOGON (0x53) C->S
'**********************************
' (BYTE[32]) Client Key (A)
' (STRING) Username
'**********************************
Private Sub SEND_SID_AUTH_ACCOUNTLOGON()
On Error GoTo ERROR_HANDLER:

    frmChat.tmrAccountLock.Enabled = True
    frmChat.tmrAccountLock.Tag = ACCOUNT_MODE_LOGON
    Call ds.NLS.Initialize(Config.Username, Config.Password)

    Dim pBuff As New clsDataBuffer
    pBuff.InsertNonNTString ds.NLS.Srp_A
    pBuff.InsertNTString ds.NLS.Username
    pBuff.SendPacket SID_AUTH_ACCOUNTLOGON
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTLOGON()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_AUTH_ACCOUNTLOGONPROOF (0x54) S->C
'**************************************
' (DWORD) Status
' (BYTE[20]) Server Password Proof (M2)
' (STRING) Additional information
'**************************************
Private Sub RECV_SID_AUTH_ACCOUNTLOGONPROOF(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    Dim M2      As String
    Dim sInfo   As String
    
    lResult = pBuff.GetDWORD
    M2 = pBuff.GetRaw(20)
    sInfo = pBuff.GetString(UTF8)

    ds.AccountEntryPending = False
    frmChat.tmrAccountLock.Enabled = False

    Select Case lResult
        Case &H0 'Logon successful.
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, &H0, sInfo)

            If (Not ds.NLS.SrpVerifyM2(M2)) Then
                frmChat.AddChat RTBColors.InformationText, "[BNCS] Warning! The server sent an invalid password proof. It may be a fake server."
            End If

            BotVars.Username = Config.Username
            BotVars.Password = Config.Password
            ds.AccountEntry = False

            If frmAccountManager.Visible Then
                frmAccountManager.LeftAccountEntryMode
            End If

            Call SendEnterChatSequence

        Case &HE 'Email registration requried
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, &H0, sInfo)

            If (Not ds.NLS.SrpVerifyM2(M2)) Then
                frmChat.AddChat RTBColors.InformationText, "[BNCS] Warning! The server sent an invalid password proof. It may be a fake server."
            End If

            BotVars.Username = Config.Username
            BotVars.Password = Config.Password
            ds.AccountEntry = False

            If frmAccountManager.Visible Then
                frmAccountManager.LeftAccountEntryMode
            End If

            Call DoRegisterEmail

        Case Else
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, lResult, sInfo)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            Else
                frmChat.DoDisconnect
            End If
    End Select

    Call ds.NLS.Terminate
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTLOGONPROOF()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_AUTH_ACCOUNTLOGONPROOF (0x54) C->S
'**************************************
' (BYTE[20]) Client Password Proof (M1)
'**************************************
Private Sub SEND_SID_AUTH_ACCOUNTLOGONPROOF()
On Error GoTo ERROR_HANDLER:
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNonNTString ds.NLS.Srp_M1
        .SendPacket SID_AUTH_ACCOUNTLOGONPROOF
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTLOGONPROOF()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCHANGE (0x55) S->C
'**********************************
' (DWORD) Status
' (BYTE[32]) Salt (s)
' (BYTE[32]) Server Key (B)
'**********************************
Private Sub RECV_SID_AUTH_ACCOUNTCHANGE(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    
    lResult = pBuff.GetDWORD
    ds.NLS.Srp_Salt = pBuff.GetRaw(32)
    ds.NLS.Srp_B = pBuff.GetRaw(32)

    Select Case lResult
        Case &H0:
            SEND_SID_AUTH_ACCOUNTCHANGEPROOF
            Exit Sub

        Case Else
            ds.AccountEntryPending = False
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, lResult, vbNullString)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_CHPWD
            Else
                frmChat.DoDisconnect
            End If
    End Select

    Call ds.NLS.Terminate

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTCHANGE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCHANGE (0x55) C->S
'**********************************
' (BYTE[32]) Client Key (A)
' (STRING) Username
'**********************************
Public Sub SEND_SID_AUTH_ACCOUNTCHANGE()
On Error GoTo ERROR_HANDLER:

    frmChat.tmrAccountLock.Enabled = True
    frmChat.tmrAccountLock.Tag = ACCOUNT_MODE_CHPWD
    Call ds.NLS.Initialize(Config.Username, Config.Password)

    Dim pBuff As New clsDataBuffer
    pBuff.InsertNonNTString ds.NLS.Srp_A
    pBuff.InsertNTString ds.NLS.Username
    pBuff.SendPacket SID_AUTH_ACCOUNTCHANGE
    Set pBuff = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTCHANGE()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCHANGEPROOF (0x56) S->C
'**********************************
' (DWORD) Status
' (BYTE[20]) Server Old Password Proof (M2)
'**********************************
Private Sub RECV_SID_AUTH_ACCOUNTCHANGEPROOF(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lResult As Long
    Dim M2      As String
    
    lResult = pBuff.GetDWORD
    M2 = pBuff.GetRaw(20)

    ds.AccountEntryPending = False
    frmChat.tmrAccountLock.Enabled = False

    Select Case lResult
        Case &H0 'Change successful.
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, &H0, vbNullString)

            Config.Password = Config.NewPassword
            Config.NewPassword = vbNullString
            Config.Save

            If (Not ds.NLS.SrpVerifyM2(M2)) Then
                frmChat.AddChat RTBColors.InformationText, "[BNCS] Warning! The server sent an invalid password proof. It may be a fake server."
            End If

            If Not frmAccountManager.Visible Then
                Call DoAccountAction(ACCOUNT_MODE_LOGON)
            Else
                frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
            End If

        Case Else
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, lResult, vbNullString)

            If Config.ManageOnAccountError Then
                frmAccountManager.ShowMode ACCOUNT_MODE_CHPWD
            Else
                frmChat.DoDisconnect
            End If

    End Select

    Call ds.NLS.Terminate

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_AUTH_ACCOUNTCHANGEPROOF()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**********************************
'SID_AUTH_ACCOUNTCHANGEPROOF (0x56) C->S
'**********************************
' (BYTE[20]) Client Password Proof (M1)
' (BYTE[32]) Salt (s)
' (BYTE[32]) Verifier (v)
'**********************************
Public Sub SEND_SID_AUTH_ACCOUNTCHANGEPROOF()
On Error GoTo ERROR_HANDLER:

    Dim oNLS As New clsNLS

    Call oNLS.Initialize(Config.Username, Config.NewPassword)
    Call oNLS.GenerateSaltAndVerifier

    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNonNTString ds.NLS.Srp_M1
        .InsertNonNTString oNLS.Srp_Salt
        .InsertNonNTString oNLS.Srp_v
        .SendPacket SID_AUTH_ACCOUNTCHANGEPROOF
    End With
    Set pBuff = Nothing

    Call oNLS.Terminate
    Set oNLS = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_AUTH_ACCOUNTCHANGEPROOF()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'*******************************
'SID_SETEMAIL (0x59) S->C
'*******************************
' [Blank]
'*******************************
Private Sub RECV_SID_SETEMAIL(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    ' do not call into EmailReg here,
    ' let receiving a successful account logon response call into it!
    ' XXX: is there any case in which SID_SETEMAIL will be received after a successful account logon (instead of before)?
    ds.WaitingForEmail = True
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.RECV_SID_SETEMAIL()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_SETEMAIL (0x59) C->S
'**************************************
' (STRING) Email Address
'**************************************
Public Sub SEND_SID_SETEMAIL(sEMailAddress As String)
On Error GoTo ERROR_HANDLER:
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNTString sEMailAddress
        .SendPacket SID_SETEMAIL
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_SETEMAIL()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_RESETPASSWORD (0x5A) C->S
'**************************************
' (STRING) Username
' (STRING) Email Address
'**************************************
Public Sub SEND_SID_RESETPASSWORD()
On Error GoTo ERROR_HANDLER:
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNTString Config.Username
        .InsertNTString Config.RegisterEmailDefault
        .SendPacket SID_RESETPASSWORD
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_RESETPASSWORD()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'**************************************
'SID_CHANGEEMAIL (0x5B) C->S
'**************************************
' (STRING) Username
' (STRING) Email Address
' (STRING) New Email Address
'**************************************
Public Sub SEND_SID_CHANGEEMAIL()
On Error GoTo ERROR_HANDLER:
    
    Dim pBuff As New clsDataBuffer
    With pBuff
        .InsertNTString Config.Username
        .InsertNTString Config.RegisterEmailDefault
        .InsertNTString Config.RegisterEmailChange
        .SendPacket SID_CHANGEEMAIL
    End With
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SEND_SID_CHANGEEMAIL()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'=======================================================================================================
'This function will open the form to prompt the user for their email, or if the overrides are set, automatically register an email.
Private Sub DoRegisterEmail()
On Error GoTo ERROR_HANDLER:

    Call frmEMailReg.DoRegisterEmail(Config.RegisterEmailAction, Config.RegisterEmailDefault)
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.DoRegisterEmail()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'=======================================================================================================
'This function will attempt to complete the CRev request that Bnet has sent to us.
'Returns True if successful.
Private Function CompileCheckrevision() As Boolean
On Error GoTo ERROR_HANDLER:
    Dim lVersion  As Long
    Dim lChecksum As Long
    Dim sResult   As String
    Dim sHeader   As String
    Dim sFile     As String
    
    sHeader = StringFormat("CRev_{0}", GetProductKey)
    If (Warden_CheckRevision(ds.CRevFileName, ds.CRevFileTimeRaw, ds.CRevSeed, sHeader, lVersion, lChecksum, sResult)) Then
        ds.CRevChecksum = lChecksum
        ds.CRevResult = sResult
        ds.CRevVersion = lVersion
        CompileCheckrevision = True
    Else
        Call frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Local hashing failed.")
        CompileCheckrevision = False
    End If
    Exit Function
ERROR_HANDLER:
    CompileCheckrevision = False
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.CompileCheckrevision()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Function GetCDKeyCount(Optional sProduct As String = vbNullString) As Long
On Error GoTo ERROR_HANDLER:

    Dim sOverride As String
    Dim lRet      As Long

    If (LenB(sProduct) = 0) Then sProduct = BotVars.Product
    
    lRet = GetProductInfo(sProduct).KeyCount
    
    GetCDKeyCount = lRet
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.GetCDKeyCount()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Function GetLogonSystem(Optional sProduct As String = vbNullString) As Long
On Error GoTo ERROR_HANDLER:

    Dim sOverride As String
    Dim tLng      As Long
    Dim lRet      As Long
    
    ' Temporary short-circuit:
    '  Return BNCS_NLS because no other login sequences are supported
    '  -andy
    'GetLogonSystem = BNCS_NLS
    'Exit Function
    
    If (LenB(sProduct) = 0) Then sProduct = BotVars.Product
    
    lRet = GetProductInfo(sProduct).LogonSystem
    
    tLng = Config.GetLogonSystem(GetProductKey(sProduct))
    If tLng > -1 Then
        Select Case tLng
            Case BNCS_NLS: lRet = BNCS_NLS
            Case BNCS_LLS: lRet = BNCS_LLS
            Case BNCS_OLS: lRet = BNCS_OLS
        End Select
    End If
    
    GetLogonSystem = lRet
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.GetLogonSystem()", Err.Number, Err.Description, OBJECT_NAME))
End Function

'Converts the normalized (forward: IX86, STAR, etc) string representation of a DWORD into it's numeric equivalent.
Private Function GetDWORDOverride(ByVal sDwordString As String, Optional ByVal lDefault As Long = 0) As Long
On Error GoTo ERROR_HANDLER:

    Dim lRet      As Long
    lRet = lDefault
    
    If ((LenB(sDwordString) > 0) And (Len(sDwordString) < 5)) Then
        lRet = GetDWORD(StrReverse(sDwordString))
    End If
    
    GetDWORDOverride = lRet
    Exit Function
ERROR_HANDLER:
    GetDWORDOverride = lRet
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.GetDWORDOverride()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Private Function GetDWORD(sData As String) As Long
On Error GoTo ERROR_HANDLER:
    
    sData = Left$(StringFormat("{0}{1}", sData, String$(4, Chr$(0))), 4)
    CopyMemory GetDWORD, ByVal sData, 4
    
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.GetDWORD()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Sub SendEnterChatSequence()
On Error GoTo ERROR_HANDLER:
    If ds.EnteredChatFirstTime Then
        Call DoChannelJoinHome(False)
    Else
        ds.EnteredChatFirstTime = True
        
        If ((Not BotVars.Product = "VD2D") And (Not BotVars.Product = "PX2D") And _
            (Not BotVars.Product = "PX3W") And (Not BotVars.Product = "3RAW")) Then
            
            If (Not BotVars.UseUDP) Then
                SEND_SID_UDPPINGRESPONSE
                'We dont use ICONDATA .SendPacket SID_GETICONDATA
            End If
        End If
        
        SEND_SID_ENTERCHAT
        SEND_SID_GETCHANNELLIST
        
        BotVars.Gateway = Config.PredefinedGateway
        If (LenB(BotVars.Gateway) = 0) Then
            If ((Not BotVars.Product = "VD2D") And (Not BotVars.Product = "PX2D") And _
                (Not BotVars.Product = "PX3W") And (Not BotVars.Product = "3RAW")) Then
                ' join nowhere to force non-W3-non-D2 to enter chat environment
                ' so they can use /whoami (see Event_ChannelJoinError for where this completes)
                Call FullJoin(BotVars.Product & BotVars.Username & Config.HomeChannel, 0)
            Else
                SEND_SID_CHATCOMMAND "/whoami"
            End If
        Else
            'PvPGN: Straight home
            Call DoChannelJoinHome
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.SendEnterChatSequence()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

' do channel join home (first time?)
Public Sub DoChannelJoinHome(Optional FirstTime As Boolean = True)
On Error GoTo ERROR_HANDLER:
    
    If (FirstTime And Config.DefaultChannelJoin) Or (LenB(Config.HomeChannel) = 0 And LenB(BotVars.LastChannel) = 0) Then
        ' product home override or home/last channel are both empty
        Call DoChannelJoinProductHome
    End If
    
    If (LenB(BotVars.LastChannel) > 0) Then
        ' go to "last channel" (for /reconnect and re-entering chat)
        Call FullJoin(BotVars.LastChannel, 2)
    ElseIf (LenB(Config.HomeChannel) > 0) Then
        ' go home
        Call FullJoin(Config.HomeChannel, 2)
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinHome()", Err.Number, Err.Description, OBJECT_NAME))
End Sub


Public Sub DoChannelJoinProductHome()
    On Error GoTo ERROR_HANDLER:
    
    Dim iJoinType As Integer
    Dim pi As udtProductInfo
    
    pi = GetProductInfo(BotVars.Product)
    
    ' D2 uses a different join type
    If pi.ShortCode = PRODUCT_D2DV Or pi.ShortCode = PRODUCT_D2XP Then
        iJoinType = 5
    Else
        iJoinType = 1
    End If
    
    Call FullJoin(pi.ChannelName, iJoinType)
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinProductHome()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub DoQueryRealms()
On Error GoTo ERROR_HANDLER:

    Set ds.MCPHandler = New clsMCPHandler
    
    SEND_SID_QUERYREALMS2

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.DoChannelJoinHome()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub DoAccountAction(Optional ByVal Mode As String = vbNullString)
On Error GoTo ERROR_HANDLER:

    ' not able to change account now?
    If frmChat.sckBNet.State <> sckConnected Or Not ds.AccountEntry Then
        Exit Sub
    End If

    ' set mode to config mode if not provided
    If LenB(Mode) = 0 Then
        Mode = Config.AccountMode
    End If

    ' check if an action is pending
    If ds.AccountEntryPending Then
        Exit Sub
    End If
    ds.AccountEntryPending = True

    ' execute action
    Select Case UCase$(Mode)
        Case ACCOUNT_MODE_CREAT
            Call Event_LogonEvent(ACCOUNT_MODE_CREAT, -1&)

            If LenB(Config.Username) = 0 Or LenB(Config.Password) = 0 Then
                ds.AccountEntryPending = False
                Call Event_LogonEvent(ACCOUNT_MODE_CREAT, -3&)

                If Config.ManageOnAccountError Then
                    frmAccountManager.ShowMode ACCOUNT_MODE_CREAT
                Else
                    frmChat.DoDisconnect
                End If

                Exit Sub
            End If

            If (ds.LogonType = BNCSSERVER_SRP2) Then
                modBNCS.SEND_SID_AUTH_ACCOUNTCREATE
            Else
                modBNCS.SEND_SID_CREATEACCOUNT2
            End If

        Case ACCOUNT_MODE_CHPWD
            Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, -1&)

            If LenB(Config.Username) = 0 Or LenB(Config.Password) = 0 Or LenB(Config.NewPassword) = 0 Then
                ds.AccountEntryPending = False
                Call Event_LogonEvent(ACCOUNT_MODE_CHPWD, -3&)

                If Config.ManageOnAccountError Then
                    frmAccountManager.ShowMode ACCOUNT_MODE_CHPWD
                Else
                    frmChat.DoDisconnect
                End If

                Exit Sub
            End If

            If (ds.LogonType = BNCSSERVER_SRP2) Then
                modBNCS.SEND_SID_AUTH_ACCOUNTCHANGE
            Else
                modBNCS.SEND_SID_CHANGEPASSWORD
            End If

        Case ACCOUNT_MODE_RSPWD
            If LenB(Config.Username) = 0 Or LenB(Config.RegisterEmailDefault) = 0 Then
                ds.AccountEntryPending = False
                Call Event_LogonEvent(ACCOUNT_MODE_RSPWD, -3&)

                If Config.ManageOnAccountError Then
                    frmAccountManager.ShowMode ACCOUNT_MODE_RSPWD
                Else
                    frmChat.DoDisconnect
                End If

                Exit Sub
            End If

            modBNCS.SEND_SID_RESETPASSWORD

            ' no response! assume success
            ds.AccountEntryPending = False

            Call Event_LogonEvent(ACCOUNT_MODE_RSPWD, &H0)

            frmAccountManager.ShowMode ACCOUNT_MODE_LOGON

        Case ACCOUNT_MODE_CHREG
            If LenB(Config.Username) = 0 Or LenB(Config.RegisterEmailDefault) = 0 Or LenB(Config.RegisterEmailChange) = 0 Then
                ds.AccountEntryPending = False
                Call Event_LogonEvent(ACCOUNT_MODE_CHREG, -3&)

                If Config.ManageOnAccountError Then
                    frmAccountManager.ShowMode ACCOUNT_MODE_CHREG
                Else
                    frmChat.DoDisconnect
                End If

                Exit Sub
            End If

            modBNCS.SEND_SID_CHANGEEMAIL

            ' no response! assume success
            ds.AccountEntryPending = False

            Call Event_LogonEvent(ACCOUNT_MODE_CHREG, &H0)

            Config.RegisterEmailDefault = Config.RegisterEmailChange
            Config.RegisterEmailChange = vbNullString
            Config.Save

            frmAccountManager.ShowMode ACCOUNT_MODE_LOGON

        Case Else ' ACCOUNT_MODE_LOGON
            Call Event_LogonEvent(ACCOUNT_MODE_LOGON, -1&)

            If LenB(Config.Username) = 0 Or LenB(Config.Password) = 0 Then
                ds.AccountEntryPending = False
                Call Event_LogonEvent(ACCOUNT_MODE_LOGON, -3&)

                If Config.ManageOnAccountError Then
                    frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
                Else
                    frmChat.DoDisconnect
                End If

                Exit Sub
            End If

            If (ds.LogonType = BNCSSERVER_SRP2) Then
                modBNCS.SEND_SID_AUTH_ACCOUNTLOGON
            Else
                modBNCS.SEND_SID_LOGONRESPONSE2
            End If

    End Select

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.DoAccountAction()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Function CanSpawn(ByVal sProduct As String, ByVal iKeyLength As Integer) As Boolean
    sProduct = GetProductInfo(sProduct).Code
    
    Select Case sProduct
        Case PRODUCT_STAR, PRODUCT_JSTR, PRODUCT_W2BN:
            CanSpawn = CBool(iKeyLength <> 26)
            Exit Function
    End Select
    CanSpawn = False
End Function

Public Sub HandleEmptyCRevSeed()
    frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] CheckRevision seed was returned empty! This is usually due to an unrecognized verison byte."
    If (BotVars.BNLS) Then
        frmChat.HandleBnlsError "[BNCS] The BNLS server you are using may be misconfigured."
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] You can reset your version bytes to the latest by going to Bot -> Update Version Bytes"
        frmChat.DoDisconnect
    End If
End Sub
