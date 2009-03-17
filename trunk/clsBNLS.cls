VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBNLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'BNLS Class - By NeTWorKz]ZeR0[ And Stealth]ZeR0[
'
'BNLS Login Framework to be used in StealthBot,
'and hopefully many, many more bots to come...
'
'Networkz is no longer here.. but I still am. This is now StealthBot-exclusive code --
' unless he comes back and asks for it, I suppose :/  --Andy

Option Explicit

Public Sub ServToken(ByRef s As String, ByVal Setting As Byte)
    Static Token As String
    
    Select Case Setting
        Case 0: s = Token
        Case 1: Token = s
    End Select
End Sub

Public Sub Send_0x01(ByVal sCDKey As String) 'BNLS_CDKEY
    ' dis paket es gross.
    Packet.InsertDWord ds.GetServerToken
    Packet.InsertNTString sCDKey
    Packet.vLSendPacket &H1
End Sub

Public Sub Send_0x02(sUser As String, sPass As String) 'BNLS_LOGONCHALLENGE
    ' sos dis 1
    Packet.InsertNTString sUser
    Packet.InsertNTString LCase(sPass)
    Packet.vLSendPacket &H2
End Sub

Public Sub Send_0x03(sData As String) 'BNLS_LOGONPROOF
    Packet.InsertNonNTString sData '(DWORD[16])  Data from SID_AUTH_ACCOUNTLOGON (0x53)
    Packet.vLSendPacket &H3
End Sub

Public Sub Send_0x04(sUser As String, sPass As String) 'BNLS_CREATEACCOUNT
    ' n dis 1
    Packet.InsertNTString sUser
    Packet.InsertNTString LCase(sPass)
    Packet.vLSendPacket &H4
End Sub

Public Sub Send_0x05(ByVal sUser As String, ByVal sOldPass As String, ByVal sNewPass As String) 'BNLS_CHANGECHALLENGE
    ' dis 1 2
    Packet.InsertNTString sUser
    Packet.InsertNTString sOldPass
    Packet.InsertNTString sNewPass
    Packet.vLSendPacket &H5
End Sub

Public Sub Send_0x06(ByVal sData As String) 'BNLS_CHANGEPROOF
    Packet.InsertNonNTString sData '(DWORD[16])  Data from SID_AUTH_ACCOUNTCHANGE (0x55).
    Packet.vLSendPacket &H6
End Sub

Public Sub Send_0x07(ByVal sUser As String, ByVal sPass As String) 'BNLS_UPGRADECHALLENGE
    ' dis 1 duznt do nuffin
    Packet.InsertNTString sUser
    Packet.InsertNTString LCase(sPass)
    Packet.InsertNTString LCase(sPass)
    Packet.vLSendPacket &H7
End Sub

Public Sub Send_0x08(ByVal sData As String) 'BNLS_UPGRADEPROOF
    Packet.InsertNonNTString sData '(DWORD)      Session key from SID_AUTH_ACCOUNTUPGRADE (0x57).
    Packet.vLSendPacket &H8
End Sub

Public Sub Send_0x09(ByVal IX86Num As Long, ByVal sCheckSumFormula As String) 'BNLS_VERSIONCHECK
'    If BotVars.Product = "PX3W" Or BotVars.Product = "3RAW" Then
'        Send_0x0D &H2
'    End If
    
    Packet.InsertDWord GetBNLSProductID(BotVars.Product)
    Packet.InsertDWord IX86Num
    
    'Debug.Print "Sending BNLS IX86Num of " & IX86Num
    
    Packet.InsertNTString sCheckSumFormula
    Packet.vLSendPacket &H9
End Sub

Public Sub Send_0x0A(ByVal sData As String) 'BNLS_CONFIRMLOGON
    Packet.InsertNonNTString sData  '(DWORD[5])   Password proof from Battle.net
    Packet.vLSendPacket &HA
End Sub

Public Sub Send_0x0B(sData As String, Optional sFlags As Long = &H1)  'BNLS_HASHDATA
    Packet.InsertDWord Len(sData)
    Packet.InsertDWord sFlags
    Packet.InsertNTString sData
    'If sFlags = &H2 Then
    '    Packet.InsertDWORD ClientKey   '(DWORD)      Client Key (Double Hash only)
    '    Packet.InsertDWORD ServerToken '(DWORD)      Server Key (Double Hash only)
    'End If
    Packet.vLSendPacket &HB
End Sub

Public Sub Send_0x0C(sCDKey As String) 'BNLS_CDKEY_EX
    Packet.InsertDWord &H0                  'Cookie
    Packet.InsertByte &H1                   'Number of CDKeys
    Packet.InsertDWord &H2                  'Flags
    Packet.InsertDWord ds.GetServerToken    'Server Key
    Packet.InsertDWord ds.GetGTC            'Client Key
    Packet.InsertNTString sCDKey            'CDKey
    Packet.vLSendPacket &HC
End Sub

Public Sub Send_0x0D(ByVal sNLS As Long) 'BNLS_CHOOSENLSREVISION
    'Debug.Print "--> Sending NLS revision " & sNLS
    Packet.InsertDWord sNLS
    Packet.vLSendPacket &HD
End Sub

Public Sub Send_0x0F(ByVal sCheckSum As Long) 'BNLS_AUTHORIZEPROOF
    Packet.InsertDWord sCheckSum
    Packet.vLSendPacket &HF
End Sub

Public Sub Send_0x10(ByVal sProdID As String) 'BNLS_REQUESTVERSIONBYTE
    Packet.InsertDWord GetBNLSProductID(sProdID)
    Packet.vLSendPacket &H10
End Sub

'Public Sub Send_0x11(sServerIP As Long, sServerSignature As String) 'BNLS_VERIFYSERVER
'    Dim sServerSig As String * 128
'    CopyMemory sServerSig, sServerSignature, 128
'    Packet.InsertDWORD sServerIP
'    Packet.InsertNonNTString sServerSig
'    Packet.vLSendPacket &H11
'End Sub

Public Sub Send_0x12(Optional ByVal sSlots As Long = 1) 'BNLS_RESERVESERVERSLOTS
    Packet.InsertDWord sSlots
    Packet.vLSendPacket &H12
End Sub

Public Sub Send_0x13(ByVal sIndex As Long, ByVal sNLS As Long, ByVal s0x53Data As String, _
    ByVal sAcctDBData As String) 'BNLS_SERVERLOGONCHALLENGE
    
    Packet.InsertDWord sIndex
    Packet.InsertDWord sNLS
    Packet.InsertNonNTString sAcctDBData '(DWORD[16])  Data from account database.
    Packet.InsertNonNTString s0x53Data   '(DWORD[8])   Data client's SID_AUTH_ACCOUNTLOGON (0x53) request
    Packet.vLSendPacket &H13
End Sub

Public Sub Send_0x14(ByVal sIndex As Long, ByVal sAccount As String, ByVal sData As String) 'BNLS_SERVERLOGONPROOF
    Packet.InsertDWord sIndex
    Packet.InsertNonNTString sData '(DWORD[5])   Data from client's SID_AUTH_ACCOUNTLOGONPROOF (0x54).
    Packet.InsertNTString sAccount
    Packet.vLSendPacket &H14
End Sub

Public Sub Send_0x1A(ByVal ProductID As Long, ByVal flags As Long, ByVal cookie As Long, _
    ByVal Timestamp As String, ByVal MPQName As String, ByVal ChecksumFormula As String)    'BNLS_VERSIONCHECKEX2

    Packet.InsertDWord ProductID
    Packet.InsertDWord flags
    Packet.InsertDWord cookie
    Packet.InsertNonNTString Timestamp
    Packet.InsertNTString MPQName
    Packet.InsertNTString ChecksumFormula
    Packet.vLSendPacket &H1A
End Sub

Public Sub SendGenPacket(ByVal s As String, ByVal PacketID As Byte) 'Send back packets that consist
                                                                    'entirely of data from BNLS.
    Packet.InsertNonNTString s
    Packet.SendPacket PacketID
    
End Sub

Public Sub ParsePacket(Data As String) 'Parse Incoming BNLS Packets
    On Error GoTo ERROR_HANDLER
    
    Dim pD        As clsPacketDebuffer
    
    Dim sTemp     As String
    Dim lTemp     As Long
    Dim lCheckSum As Long
    Dim Splt()    As String
    Dim PacketID  As Byte
    
    Set pD = New clsPacketDebuffer
    
    pD.DebuffPacket Mid$(Data, 4)
    
    PacketID = Asc(Mid(Data, 3, 1))
    
    If (MDebug("all")) Then
        frmChat.AddChat COLOR_BLUE, "BNLS RECV 0x" & _
            Right$("00" & hex(PacketID), 2)
    End If
    
    CachePacket StoC, stBNLS, PacketID, Len(Data), Data
    
    ' Added 2007-06-08 for a packet logging menu feature to aid tech support
    WritePacketData BNLS, StoC, PacketID, Len(Data), Data
    
    ' ...
    RunInAll "Event_PacketReceived", "BNLS", PacketID, Len(Data), Data
    
    If (GetVeto) Then
        Exit Sub
    End If
    
    'Debug.Print "_-------------------------------------_"
    'Debug.Print DebugOutput(Data)
    
    Select Case (PacketID)
        Case &HD 'BNLS_CHOOSENLSREVISION
            'Debug.Print Conv(Mid$(Data, 4, 4))
            '(BOOL) success
            'Implementation to cover future b.net upgrades?
            'Not required. No implementation at this time.
            
        Case &HE 'BNLS_AUTHORIZE
            lTemp = Conv(Mid(Data, 4, 4))
            lCheckSum = BNLSChecksum("gn1ftx14oc", lTemp)
            'Call frmChat.Event_BNLSConnected
            Call Send_0x0F(lCheckSum)
        
        Case &HF 'BNLS_AUTHORIZEPROOF
            Passed0x0F = 1
            Call frmChat.Event_BNLSAuthEvent(True)
            Call frmChat.Event_BNetConnecting

            frmChat.sckBNet.Connect 'BNLS is authorized, proceed to initiate BNet connection.
            
        Case &H10 'BNLS_REQUESTVERSIONBYTE
            lTemp = Conv(Mid$(Data, 4, 4))
            
            If lTemp <> &H0 Then
                lTemp = Conv(Mid(Data, 8, 4))
                Send0x50 lTemp
            Else
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNLS] Version byte request failed!"
                CloseConnection
            End If
            
        Case &H11 'BNLS_VERIFYSERVER
            'Nothing?
            Send_0x09 ds.GetMPQRev, ds.GetHashCmd
            
        Case &H13
            sTemp = Mid(Data, 8, 128)
            'SendGenPacket sTemp, &H53
            
        Case &H1A ' BNLS_VERSIONCHECKEX2
            If Len(Data) > 12 Then  ' We have a response!
                pD.Advance 4
                
                ds.SetVersion pD.DebuffDWORD
                ds.SetCheckSum pD.DebuffDWORD
                ds.SetExeInfo pD.DebuffNTString
                
                pD.Advance 4
                ds.SetVerByte pD.DebuffDWORD
                
                ' removed cd key hashing @ bnls ~Pyro
                'Call Send_0x01(BotVars.CDKey)
                
                Call Send_BNLS0x51

            Else
                frmChat.Event_BNLSDataError 2
                CloseConnection
            End If
'            (BOOL) Success (TRUE if successful, FALSE otherwise).
'               If this is FALSE, the next DWORD is the provided cookie,
'               following which the message ends.
'            (DWORD) Version.
'            (DWORD) Checksum.
'            (String) Version check stat string.
'            (DWORD) Cookie.
'            (DWORD) The latest version code for this product.
            
        Case &H1 'BNLS_CDKEY
            
            If Len(Data) < 8 Then
                Call frmChat.Event_BNLSDataError(Conv(Mid(Data, 4, 4)))
                CloseConnection
            Else
            
                ds.SetGTC Conv(Mid$(Data, 8, 4))
                
                ds.SetKeyHash Mid$(Data, 12, 36), 1
                
                'If BotVars.Product = "PX3W" Or BotVars.Product = "PX2D" Then
                    ' no more bnls key hashing
                    'Call Send_0x0C(BotVars.ExpKey)
                'Else
                    'Call Send_BNLS0x51
                'End If
                
            End If
            
        Case &H2 'BNLS_LOGONCHALLENGE
            sTemp = Mid$(Data, 4)
            
            With PBuffer
                .InsertNonNTString sTemp
                .InsertNTString BotVars.Username
                .SendPacket &H53
            End With
            
        Case &H3 'BNLS_LOGONPROOF
            sTemp = pD.DebuffRaw(20)
            lTemp = Conv(Mid$(sTemp, 1, 4))
            
            If (lTemp = 0) Then
                ' New condition added 1/22/06
                '  All zeroes means that the logon WILL fail and the connection
                '  should be aborted and retried. This is a Battle.net bug.
                With frmChat
                    If (AttemptedFirstReconnect = False) Then
                        AttemptedFirstReconnect = True
                        .AddChat RTBColors.ErrorMessageText, "[BNLS] The BNLS server has reported that " & _
                            "this logon will fail because of a Battle.net bug."
                        
                        .AddChat RTBColors.ErrorMessageText, "[BNLS] This is nothing to worry about. " & _
                            "The bot will try connecting again and should succeed."
                            
                        .DoDisconnect
                        
                        Call Pause(1)
                        
                        .DoConnect
                    Else
                        .AddChat RTBColors.ErrorMessageText, "[BNLS] The BNLS server has reported that " & _
                            "this logon will fail because of a Battle.net bug."
                            
                        .AddChat RTBColors.ErrorMessageText, "[BNLS] Please try your connection at a " & _
                            "later time."
                            
                        .DoDisconnect
                    End If
                End With
            Else
                SendGenPacket sTemp, &H54
            End If
            
        Case &H4 'BNLS_CREATEACCOUNT
            sTemp = Mid$(Data, 4, 64)
            PBuffer.InsertNonNTString sTemp
            PBuffer.InsertNTString BotVars.Username
            PBuffer.SendPacket &H52
            
        Case &H5 'BNLS_CHANGECHALLENGE
            sTemp = Mid(Data, 4, 32)
            SendGenPacket sTemp, &H55
            
        Case &H6 'BNLS_CHANGEPROOF
            sTemp = Mid(Data, 4, 84)
            SendGenPacket sTemp, &H56
            
        Case &H7 'BNLS_UPGRADECHALLENGE
            '(BOOL) Success code
            'Currently, no error conditions are defined, so this is always TRUE
            Packet.SendPacket &H57
            
        Case &H8 'BNLS_UPGRADEPROOF
            sTemp = Mid(Data, 4, 88)
            SendGenPacket sTemp, &H58
            
        Case &H9 'BNLS_VERSIONCHECK
            If Len(Data) > 12 Then
                ds.SetVersion Conv(Mid$(Data, 8, 4))
                ds.SetCheckSum Conv(Mid$(Data, 12, 4))
                ds.SetExeInfo KillNull(Right$(Data, Len(Data) - 15))
                
                ' no more bnls key hashing
                'Call Send_0x01(BotVars.CDKey)

            Else
                frmChat.Event_BNLSDataError 2
                CloseConnection
            End If
            '(bool) Success
            'If Success is TRUE:
            '(DWORD) Version.
            '(DWORD) Checksum.
            '(STRING) Version check stat string
            
        Case &HA 'BNLS_CONFIRMLOGON
            '(bool) Success
            
        Case &HB 'BNLS_HASHDATA
            sTemp = Mid(Data, 4, 20)
            'Call Send_BNLS0x3A(sTemp)
            
        Case &HC 'BNLS_CDKEY_EX
            'DWORD Cookie       4567
            'BYTE   #Keys       8
            'BYTE   #Successful 9
            'DWORD  Successcode 10111213
            'DWORD  Clientkey   14151617
            If (Len(Data) < 40) Then
                frmChat.Event_BNLSDataError 2
                
                Call CloseConnection
            Else
                ds.SetKeyHash Mid(Data, 18, 36), 2 '9 DWORD KeyData
                Call Send_BNLS0x51
            End If
        
    End Select
    
    Exit Sub

ERROR_HANDLER:
    Exit Sub
    
End Sub

Sub CloseConnection(Optional ByVal DoNotShow As Byte = 0)
    If (frmChat.sckBNLS.State <> 0) Then: frmChat.sckBNLS.Close
    If (frmChat.sckBNet.State <> 0) Then: frmChat.sckBNet.Close
    If (frmChat.sckMCP.State <> 0) Then: frmChat.sckMCP.Close
    
    If (DoNotShow = 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, _
            "All connections closed."
    End If
    
    Passed0x0F = 0
    
    SetTitle "Disconnected"
    
    On Error Resume Next
    
    RunInAll "Event_ServerError", "All connections closed."
End Sub

Private Sub Send_BNLS0x51()
    Dim KeyHash As String       ' CDKey hash
    Dim Value1 As Long          ' CDKey Value 1
    Dim ProductID As Long       ' CDKey Product ID
    Dim ClientToken As Long
    Dim lngWardenSeed As Long   ' seed for warden
    
    ClientToken = GetTickCount()
    ds.SetGTC ClientToken
    
    With PBuffer
        .InsertDWord ds.GetGTC
        .InsertDWord ds.GetVersion
        .InsertDWord ds.GetChecksum
        
        If ((BotVars.Product = "PX2D") Or _
            (BotVars.Product = "PX3W")) Then
           
            .InsertDWord &H2
        Else
            .InsertDWord &H1
        End If
        
        .InsertDWord &H0
        
        ' no more bnls key hashing
        '.InsertNonNTString ds.GetKeyHash(1)
        
        Call DecodeCDKey(BotVars.CDKey, ds.GetServerToken, ds.GetGTC, KeyHash, Value1, ProductID, ds.GetMPQRev)
        
        .InsertDWord Len(BotVars.CDKey)
        .InsertDWord ProductID
        .InsertDWord Value1
        .InsertDWord &H0

        .InsertNonNTString KeyHash
        
        'Add warden support
        If BotVars.Product = "RATS" Or BotVars.Product = "PXES" Then
            Call CopyMemory(lngWardenSeed, ByVal KeyHash, 4)
            Call modWarden.WardenInit(lngWardenSeed)
        End If
        
        If ((BotVars.Product = "PX2D") Or _
            (BotVars.Product = "PX3W")) Then
            
            ' no more bnls key hashing
            '.InsertNonNTString ds.GetKeyHash(2)
            
            Call DecodeCDKey(BotVars.ExpKey, ds.GetServerToken, ds.GetGTC, KeyHash, Value1, ProductID, ds.GetMPQRev)
        
            .InsertDWord Len(BotVars.ExpKey)
            .InsertDWord ProductID
            .InsertDWord Value1
            .InsertDWord &H0
            .InsertNonNTString KeyHash
        End If
        
        'Debug.Print ds.getExeInfo
        .InsertNTString ds.getExeInfo
        
        If (LenB(ReadCfg("Override", "OwnerName")) > 0) Then
            .InsertNTString ReadCfg("Override", "OwnerName")
        Else
            .InsertNTString BotVars.Username
        End If
        
        'ds.List
        
        .SendPacket &H51
    End With
End Sub
