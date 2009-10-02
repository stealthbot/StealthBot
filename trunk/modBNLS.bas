Attribute VB_Name = "modBNLS"
Option Explicit
Private Const OBJECT_NAME As String = "modBNLS"

Private Const BNLS_CHOOSENLSREVISION  As Byte = &HD
Private Const BNLS_AUTHORIZE          As Byte = &HE
Private Const BNLS_AUTHORIZEPROOF     As Byte = &HF
Private Const BNLS_REQUESTVERSIONBYTE As Byte = &H10
Private Const BNLS_VERSIONCHECKEX2    As Byte = &H1A

Public BNLSAuthorized As Boolean

Public Function BNLSRecvPacket(ByVal sData As String) As Boolean
    Static pBuff As New clsDataBuffer
    
    Dim PacketID As Byte
    
    BNLSRecvPacket = True
    With pBuff
        .Clear
        .Data = sData
        .GetWord
        PacketID = .GetByte
    End With
    
    If (MDebug("all")) Then
        frmChat.AddChat COLOR_BLUE, StringFormat("BNLS RECV 0x{0}", ZeroOffset(PacketID, 2))
    End If
    
    CachePacket StoC, stBNLS, PacketID, Len(sData), sData
    
    ' Added 2007-06-08 for a packet logging menu feature to aid tech support
    WritePacketData BNLS, StoC, PacketID, Len(sData), sData
    
    If (RunInAll("Event_PacketReceived", "BNLS", PacketID, Len(sData), sData)) Then
        Exit Function
    End If
    
    Select Case PacketID
        
        Case BNLS_CHOOSENLSREVISION:  Call RECV_BNLS_CHOOSENLSREVISION(pBuff)
        Case BNLS_AUTHORIZE:          Call RECV_BNLS_AUTHORIZE(pBuff)
        Case BNLS_AUTHORIZEPROOF:     Call RECV_BNLS_AUTHORIZEPROOF(pBuff)
        Case BNLS_REQUESTVERSIONBYTE: Call RECV_BNLS_REQUESTVERSIONBYTE(pBuff)
        Case BNLS_VERSIONCHECKEX2:    Call RECV_BNLS_VERSIONCHECKEX2(pBuff)
        
        Case Else:
            BNLSRecvPacket = False
            If (MDebug("debug") And (MDebug("all") Or MDebug("unknown"))) Then
                Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNLS] Unhandled packet 0x{0}", ZeroOffset(CLng(PacketID), 2)))
                Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNLS] Packet data: {0}{1}", vbNewLine, DebugOutput(sData)))
            End If
    
    End Select
    
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.BNLSRecvPacket()", Err.Number, Err.description, OBJECT_NAME))
End Function

'**********************************
'BNLS_CHOOSENLSREVISION (0x0D) S->C
'**********************************
' (DWORD) Success (1 = success)
'**********************************
Public Sub RECV_BNLS_CHOOSENLSREVISION(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    If (pBuff.GetDWORD = 1) Then
        If (MDebug("debug")) Then
            frmChat.AddChat RTBColors.InformationText, "[BNLS] NLS Revision Accepted"
        End If
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNLS] NLS Revision Rejected"
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_CHOOSENLSREVISION()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'**********************************
'BNLS_CHOOSENLSREVISION (0x0D) C->S
'**********************************
' (DWORD) NLS revision number.
'**********************************
Public Sub SEND_BNLS_CHOOSENLSREVISION(ByVal lNLSRev As Long)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    pBuff.InsertDWord lNLSRev
    pBuff.vLSendPacket BNLS_CHOOSENLSREVISION
    Set pBuff = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_CHOOSENLSREVISION()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'*******************************
'BNLS_AUTHORIZE (0x0E) S->C
'*******************************
' (DWORD) Server Token
'*******************************
Private Sub RECV_BNLS_AUTHORIZE(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    Call SEND_BNLS_AUTHORIZEPROOF(pBuff.GetDWORD)

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_AUTHORIZE()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'*******************************
'BNLS_AUTHORIZE (0x0E) C->S
'*******************************
' (String) Bot ID
'*******************************
Public Sub SEND_BNLS_AUTHORIZE(Optional sBotID As String = vbNullString)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    pBuff.InsertNTString IIf(LenB(sBotID) = 0, "stealth", sBotID)
    pBuff.vLSendPacket BNLS_AUTHORIZE
    Set pBuff = Nothing

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_AUTHORIZE()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'*******************************
'BNLS_AUTHORIZEPROOF (0x0F) S->C
'*******************************
' (DWORD) Server Token
'*******************************
Private Sub RECV_BNLS_AUTHORIZEPROOF(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    BNLSAuthorized = True
    Call frmChat.Event_BNLSAuthEvent(True)
    Call frmChat.Event_BNetConnecting
    frmChat.sckBNet.Connect

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_AUTHORIZEPROOF()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'*******************************
'BNLS_AUTHORIZEPROOF (0x0F) C->S
'*******************************
' (DWORD) Password Checksum
'*******************************
Private Sub SEND_BNLS_AUTHORIZEPROOF(lServerToken As Long, Optional sPassword As String = vbNullString)
On Error GoTo ERROR_HANDLER:

    Dim cCRC      As New clsCRC32
    Dim lChecksum As Long
    Dim pBuff     As New clsDataBuffer
    
    lChecksum = cCRC.CRC32(StringFormat("{0}{1}", _
        IIf(LenB(sPassword) = 0, "gn1ftx14oc", sPassword), _
        ZeroOffset(lServerToken, 8)))
    
    
    pBuff.InsertDWord lChecksum
    pBuff.vLSendPacket BNLS_AUTHORIZEPROOF
    
    Set cCRC = Nothing
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_AUTHORIZEPROOF()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'************************************
'BNLS_REQUESTVERSIONBYTE (0x10) S->C
'************************************
' (DWORD) Product ID (0 if failure)
' (DWORD) Version Byte (Not included if failure)
'************************************
Private Sub RECV_BNLS_REQUESTVERSIONBYTE(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:
    
    Dim lVerByte     As Long
    Dim lLogonSystem As Long
    
    If (Not pBuff.GetDWORD = 0) Then
        lVerByte = pBuff.GetDWORD
        
        lLogonSystem = modBNCS.GetLogonSystem()
        If (lLogonSystem = modBNCS.BNCS_NLS) Then
            modBNCS.SEND_SID_AUTH_INFO lVerByte
        Else 'TO-DO: Non-NLS connection
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", lLogonSystem)
            frmChat.DoDisconnect
        End If
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNLS] Version byte request failed!"
        frmChat.DoDisconnect
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_REQUESTVERSIONBYTE()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'************************************
'BNLS_REQUESTVERSIONBYTE (0x10) C->S
'************************************
' (DWORD) ProductID
'************************************
Public Sub SEND_BNLS_REQUESTVERSIONBYTE(Optional sProduct As String = vbNullString)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    
    pBuff.InsertDWord GetBNLSProductID(IIf(LenB(sProduct) = 0, BotVars.Product, sProduct))
    pBuff.vLSendPacket BNLS_REQUESTVERSIONBYTE
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_REQUESTVERSIONBYTE()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'************************************
'BNLS_VERSIONCHECKEX2 (0x1A) S->C
'************************************
' (DWORD) Success*
' (DWORD) Version.
' (DWORD) Checksum.
' (STRING) Version check stat string.
' (DWORD) Cookie.
' (DWORD) Version Byte
'************************************
Private Sub RECV_BNLS_VERSIONCHECKEX2(pBuff As clsDataBuffer)
On Error GoTo ERROR_HANDLER:

    If (pBuff.GetDWORD = 1) Then
        ds.CRevVersion = pBuff.GetDWORD
        ds.CRevChecksum = pBuff.GetDWORD
        ds.CRevResult = pBuff.GetString
        modBNCS.SEND_SID_AUTH_CHECK
    Else
        frmChat.Event_BNLSDataError 2
        frmChat.DoDisconnect
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.RECV_BNLS_VERSIONCHECKEX2()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'************************************
'BNLS_VERSIONCHECKEX2 (0x1A) C->S
'************************************
' (DWORD) Product ID.*
' (DWORD) Flags.**
' (DWORD) Cookie.
' (FILETIME) CRev Archive File Time
' (STRING) CRev Archive File Name
' (STRING) CRev Seed Values
'************************************
Public Sub SEND_BNLS_VERSIONCHECKEX2(sCRevFileTime As String, sCRevFileName As String, sCRevSeeds As String, Optional sProduct As String = vbNullString, Optional lFlags As Long = 0, Optional lCookie As Long = 1)
On Error GoTo ERROR_HANDLER:

    Dim pBuff As New clsDataBuffer
    
    With pBuff
        .InsertDWord GetBNLSProductID(sProduct)
        .InsertDWord lFlags
        .InsertDWord lCookie
        .InsertNonNTString sCRevFileTime
        .InsertNTString sCRevFileName
        .InsertNTString sCRevSeeds
        .vLSendPacket BNLS_VERSIONCHECKEX2
    End With
    
    Set pBuff = Nothing
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in {2}.SEND_BNLS_VERSIONCHECKEX2()", Err.Number, Err.description, OBJECT_NAME))
End Sub

'===================================================================================================
Public Function GetBNLSProductID(Optional ByVal sProdID As String = vbNullString) As Long
    If (LenB(sProdID) = 0) Then sProdID = BotVars.Product

    Select Case (UCase$(sProdID))
        Case "RATS", "STAR": GetBNLSProductID = &H1
        Case "PXES", "SEXP": GetBNLSProductID = &H2
        Case "NB2W", "W2BN": GetBNLSProductID = &H3
        Case "VD2D", "D2DV": GetBNLSProductID = &H4
        Case "PX2D", "D2XP": GetBNLSProductID = &H5
        Case "RTSJ", "JSTR": GetBNLSProductID = &H6
        Case "3RAW", "WAR3": GetBNLSProductID = &H7
        Case "PX3W", "W3XP": GetBNLSProductID = &H8
        Case "LTRD", "DRTL": GetBNLSProductID = &H9
        Case "RSHD", "DSHR": GetBNLSProductID = &HA
        Case "RHSS", "SSHR": GetBNLSProductID = &HB
        Case Else:           GetBNLSProductID = &H0
    End Select
    
End Function
