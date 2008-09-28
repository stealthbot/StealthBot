Attribute VB_Name = "modOutboundPackets"
'StealthBot project - modOutboundPackets
' Packet creation, checkrevision and NLS code
' March 2005
' Andy T andy@stealthbot.net
Option Explicit

Public g_username As String

Public Sub Send0x50(Optional lVerByte As Long)
    Dim CAbbr As String, cName As String
    Dim s As String

    With PBuffer
        .InsertDWord &H0 'Protocol ID (Zero)
        
        s = StrReverse(UCase(ReadCFG("Override", GetProductKey() & "PlatID")))
        
        If Len(s) > 0 Then
            If Len(s) > 4 Then s = Mid$(s, 1, 4)
            
            .InsertNonNTString s & BotVars.Product 'Platform ID and Product ID
        Else
            .InsertNonNTString "68XI" & BotVars.Product 'Platform ID and Product ID
        End If
        
        If (lVerByte > 0) Then
            .InsertDWord lVerByte
        Else
            .InsertDWord GetVerByte(BotVars.Product) 'Version byte
        End If
        
        .InsertDWord &H0 ' Product Language
        .InsertDWord &H0 ' Local IP (defunct)
        .InsertDWord GetTimeZoneBias ' Timezone bias
        
        If LenB(ReadCFG("Override", "ForceDefaultLocaleID")) > 0 Then
            .InsertDWord 1033 ' US English
            .InsertDWord 1033
        Else
            .InsertDWord CLng(GetUserDefaultLCID)
            .InsertDWord CLng(GetUserDefaultLangID)
        End If
        
        Call GetCountryData(CAbbr, cName)
        
        If Len(CAbbr) = 3 Then
            .InsertNTString CAbbr
        Else
            .InsertNTString "USA"
        End If
        
        If LenB(cName) > 0 Then
            .InsertNTString cName
        Else
            .InsertNTString "United States"
        End If
        
        .SendPacket &H50
        
        If BotVars.Spoof = 1 Then
            .InsertDWord &H0
            .SendPacket &H25
        End If
    End With
End Sub


Public Sub Send0x51(ByVal ServerToken As Long)
    Dim ClientToken As Long     ' Client token
    Dim Path As String          ' Hash file path
    Dim EXEInfo As String       ' EXE Information
    Dim Checksum As Long        ' EXE Checksum
    Dim Version As Long         ' EXE Version
    Dim ValueString As String   ' "Hash Command" -- Values()tring used in checkrevision
    Dim Value1 As Long          ' CDKey Value 1
    Dim ProductID As Long       ' CDKey Product ID
    Dim MPQRevision As Long     ' MPQ Revision
    Dim HashPaths(2) As String  ' Hash file paths
    Dim KeyHash As String       ' CDKey hash
    Dim X As Long               ' Stop control variable
    
    Path = GetGamePath(BotVars.Product)
    ClientToken = GetTickCount()
    
    Select Case BotVars.Product
        Case "VD2D", "PX2D"
            HashPaths(0) = Path & "Game.exe"
            HashPaths(1) = Path & "BNClient.dll"
            HashPaths(2) = Path & "D2Client.dll"
            
        Case "NB2W"
            HashPaths(0) = Path & "Warcraft II BNE.exe"
            HashPaths(1) = Path & "Storm.dll"
            HashPaths(2) = Path & "Battle.snp"
            
        Case "RATS", "PXES"
            HashPaths(0) = Path & "Starcraft.exe"
            HashPaths(1) = Path & "Storm.dll"
            HashPaths(2) = Path & "Battle.snp"
            
        Case "PX3W", "3RAW"
            HashPaths(0) = Path & "war3.exe"
            HashPaths(1) = Path & "Storm.dll"
            HashPaths(2) = Path & "game.dll"   ' Might be the wrong filename..
            
        Case Else
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] You are attempting to connect with an unknown client. Connection aborted."
            frmChat.DoDisconnect
            Exit Sub
            
    End Select
    
    X = X Or CheckPath(HashPaths(0))
    X = X Or CheckPath(HashPaths(1))
    X = X Or CheckPath(HashPaths(2))
    
    If X = 1 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Connection terminated due to missing hash files."
        frmChat.DoDisconnect
    Else
        ' Do CheckRevision
        ValueString = ds.GetHashCmd
        MPQRevision = ds.GetMPQRev
        
        If checkRevision(ValueString, HashPaths(0), HashPaths(1), HashPaths(2), MPQRevision, Checksum) Then
            Version = getExeInfo(HashPaths(0), EXEInfo)
            
            With PBuffer
                .InsertDWord ClientToken    ' Client token
                .InsertDWord Version        ' CheckRevision version
                .InsertDWord Checksum       ' CheckRevision checksum
                
                ' Number of CDKeys
                If BotVars.Product = "PX2D" Or BotVars.Product = "PX3W" Then
                    .InsertDWord &H2
                Else
                    .InsertDWord &H1
                End If
                
                ' For each cdkey...
                Call DecodeCDKey(BotVars.CDKey, ServerToken, ClientToken, KeyHash, Value1, ProductID, MPQRevision)
                
                .InsertDWord &H0
                .InsertDWord Len(BotVars.CDKey)
                .InsertDWord ProductID
                .InsertDWord Value1
                .InsertDWord &H0
                .InsertNonNTString KeyHash
                
                'Warden code commented out for local hashing as I'm not sure if KeyHash actually contains the cd-key hash
                'Uncomment this when it's verified, hashing seems to be completely broken for me, so I couldn't test it
                '   ~ FrOzeN
                
                'Add warden support
                'If BotVars.Product = "RATS" Or BotVars.Product = "PXES" Then
                '    modWarden.InitializeWarden Left$(KeyHash, 4)
                'End If
                
                If BotVars.Product = "PX2D" Or BotVars.Product = "PX3W" Then
                    Call DecodeCDKey(BotVars.ExpKey, ServerToken, ClientToken, KeyHash, Value1, ProductID, MPQRevision)
                    
                    .InsertDWord Len(BotVars.ExpKey)
                    .InsertDWord ProductID
                    .InsertDWord Value1
                    .InsertDWord &H0
                    .InsertNonNTString KeyHash
                End If

                .InsertNTString EXEInfo
                
                If (LenB(ReadCFG("Override", "OwnerName")) > 0) Then
                    .InsertNTString ReadCFG("Override", "OwnerName")
                Else
                    .InsertNTString g_username
                End If
                
                .SendPacket &H51
            End With
        Else
            frmChat.AddChat RTBColors.ErrorMessageText, "CheckRevision failed!"
            frmChat.DoDisconnect
        End If
    End If
End Sub

' Uses BNCSUtil to decode and hash your cdkey
' Returns the ClientToken that should be used to connect
Public Sub DecodeCDKey(ByVal sCDKey As String, ByVal ServerToken As Long, ByVal ClientToken As Long, ByRef KeyHash As String, ByRef Value1 As Long, ByRef ProductID As Long, ByVal MPQRevision As Long)
    Dim KDh As Long                     ' Key Decoder handler
    Dim HashSize As Long                ' CDKey hash size in bytes
    Dim result As Long                  ' kd_init() result
    
    sCDKey = Replace(sCDKey, "-", vbNullString)
    sCDKey = Replace(sCDKey, " ", vbNullString)
    sCDKey = KillNull(sCDKey)
    
    result = kd_init()
    
    If result = 0 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "BNCSUtil: kd_init() failed! Please use BNLS to connect."
        frmChat.DoDisconnect
        
    Else
        KDh = kd_create(sCDKey, Len(sCDKey))
            
        If (kd_isValid(KDh) = 0) Then
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Your CD-Key is invalid."
            frmChat.DoDisconnect
            
        Else
            HashSize = kd_calculateHash(KDh, ClientToken, ServerToken)
        
            If HashSize <= 0 Then
                frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid. [kd_calculateHash() <= 0]"
                frmChat.AddChat RTBColors.ErrorMessageText, "Please make sure you typed your CD-Key correctly. This error is often generated when the CD-Key is not the correct Length."
                frmChat.DoDisconnect
                
            Else
                KeyHash = String$(HashSize, vbNullChar)
                Call kd_getHash(KDh, KeyHash)
                
                Value1 = kd_val1(KDh)
                ProductID = kd_product(KDh)
                
            End If
            
        End If
        
    End If
    
    If KDh > 0 Then
        Call kd_free(KDh)
    End If
End Sub

Public Sub Send0x3A(ByVal ServerToken As Long)
    Dim PasswordHash As String
    
    If LenB(ReadCFG("Override", "UppercasePassword")) > 0 Then
        PasswordHash = doubleHashPassword(BotVars.Password, ds.GetGTC, ServerToken)
    Else
        PasswordHash = doubleHashPassword(LCase(BotVars.Password), ds.GetGTC, ServerToken)
    End If
    
    With PBuffer
        .InsertDWord ds.GetGTC
        .InsertDWord ds.GetServerToken
        .InsertNonNTString PasswordHash
        .InsertNTString g_username
        .SendPacket &H3A
    End With
End Sub

Public Sub Send0x0A()
    Dim num As Integer

    If BotVars.Product <> "VD2D" And BotVars.Product <> "PX2D" And _
        BotVars.Product <> "PX3W" And BotVars.Product <> "3RAW" Then
        
        If Not BotVars.UseUDP Then
            With PBuffer
                If Len(ReadCFG("Override", "UDPString")) = 4 Then
                    .InsertNonNTString ReadCFG("Override", "UDPString")
                Else
                    .InsertNonNTString "tenb"
                End If
                
                .SendPacket &H14
                .SendPacket &H2D
            End With
        End If
    End If
    
    ' ...
    Randomize
    
    ' ...
    num = (1 + Rnd() * 1000)
    
    ' ...
    With PBuffer
        .InsertNTString g_username
        .InsertNTString vbNullString
        .SendPacket &HA
    End With
    
    RequestSystemKeys
    
    With PBuffer
        .InsertNonNTString BotVars.Product
        .SendPacket &HB
    End With
    
    FullJoin BotVars.HomeChannel & ":" & num, 0
    
    With PBuffer
        .InsertNTString "/whoami"
        .SendPacket &HE
    End With
    
    Call InsertDummyQueueEntry
End Sub

Public Sub AttemptAccountCreation()
    Dim PasswordHash As String
    
    PasswordHash = hashPassword(LCase(BotVars.Password))
    
    With PBuffer
        .InsertNonNTString PasswordHash
        .InsertNTString g_username
        .SendPacket &H3D
    End With
End Sub


Public Sub Send0x09(ByVal ServerToken As Long, ByVal UDPValue As Long)
    If BotVars.UseUDP Then
        With PBuffer
            .InsertDWord ServerToken
            .InsertDWord UDPValue
            .SendPacket &H9
        End With
    End If
End Sub


Public Sub CreateNLSObject()
    Dim Handle As Long
    
    Call DestroyNLSObject
    Handle = nls_init(g_username, BotVars.Password)
    
    If Handle = 0 Then
        frmChat.AddChat "Warning: nls_init() failed! Please update BNCSutil.dll."
        frmChat.DoDisconnect
    Else
        ds.NLSHandle = Handle
    End If
End Sub


Public Sub DestroyNLSObject()
    If ds.NLSHandle > 0 Then
        Call nls_free(ds.NLSHandle)
        ds.NLSHandle = 0
    End If
End Sub


Public Sub Send0x52()
    Dim result As Long
    Dim sBuffer As String
    
    sBuffer = String$(NLS_ACCOUNTCREATE_ + Len(g_username), vbNullChar)
    result = nls_account_create(ds.NLSHandle, sBuffer, NLS_ACCOUNTCREATE_ + Len(g_username))
    
    With PBuffer
        .InsertNonNTString sBuffer
        .SendPacket &H52
    End With
End Sub

Public Sub Send0x53()
    Dim sBuffer As String
    
    sBuffer = String$(NLS_GET_A_, vbNullChar)
    Call nls_get_A(ds.NLSHandle, sBuffer)
    
    With PBuffer
        .InsertNonNTString sBuffer
        .InsertNTString g_username
        .SendPacket &H53
    End With
End Sub

Public Sub Send0x54(ByVal Salt As String, ByVal ServerKey As String)
    Dim sBuffer As String
    
    sBuffer = String$(NLS_GET_M1_, vbNullChar)
    
    Call nls_get_M1(ds.NLSHandle, sBuffer, ServerKey, Salt)
    
    With PBuffer
        .InsertNonNTString sBuffer
        .SendPacket &H54
    End With
End Sub

Public Sub Send0x5E(ByRef PacketData As String) 'SID_WARDEN
    Dim strWardenPacket As String
    strWardenPacket = modWarden.HandleWarden(PacketData)

    If LenB(strWardenPacket) > 0 Then
        PBuffer.InsertNonNTString strWardenPacket
        PBuffer.SendPacket &H5E
    End If
    
    Static blWarden As Boolean
    If Not blWarden Then
        blWarden = True
        
        If LenB(strWardenPacket) > 0 Then
            MsgBox "Warden working"
        Else
            MsgBox "Warden broken"
        End If
    End If
End Sub
