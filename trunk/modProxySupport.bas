Attribute VB_Name = "modProxySupport"
Option Explicit

'modProxySupport - project StealthBot
'February 2004 - by Stealth [stealth at stealthbot dot net]

Public Const PROXY_CLIENT_ERROR    As Long = &H400
Public Const PROXY_METHOD_ERROR    As Long = &H410
Public Const PROXY_LOGON_ERROR     As Long = &H420
Public Const PROXY_LOGON_SUCCESS   As Long = &H421
Public Const PROXY_LOGGING_ON      As Long = &H422
Public Const PROXY_REQUEST_ERROR   As Long = &H430
Public Const PROXY_REQUEST_SUCCESS As Long = &H431
Public Const PROXY_REQUESTING_CONN As Long = &H432

Public Const PROXY_SETTING_SOCKS5 As String = "SOCKS5"
Public Const PROXY_SETTING_SOCKS4 As String = "SOCKS4"

Private Const SOCKS4_VER           As Byte = 4
Private Const SOCKS5_VER           As Byte = 5
Private Const SOCKS5_VER_USERPASS  As Byte = 1

Private Const SOCKS_REQ_TCPCONN    As Byte = 1

Private Const SOCKS4_REQ_SUCCESS   As Byte = 90
Private Const SOCKS4_REQ_FAIL      As Byte = 91
Private Const SOCKS4_REQ_ID_RUN    As Byte = 92
Private Const SOCKS4_REQ_ID_FAIL   As Byte = 93

Private Const SOCKS5_MET_ANON      As Byte = &H0
Private Const SOCKS5_MET_USERPASS  As Byte = &H2
Private Const SOCKS5_MET_NONE      As Byte = &HFF

Private Const SOCKS5_REQ_SUCCESS   As Byte = 0
Private Const SOCKS5_REQ_FAIL      As Byte = 1
Private Const SOCKS5_REQ_RULESET   As Byte = 2
Private Const SOCKS5_REQ_NONET     As Byte = 3
Private Const SOCKS5_REQ_NOHOST    As Byte = 4
Private Const SOCKS5_REQ_REFUSED   As Byte = 5
Private Const SOCKS5_REQ_TTLEXP    As Byte = 6
Private Const SOCKS5_REQ_PROTOCOL  As Byte = 7
Private Const SOCKS5_REQ_ADDRTYPE  As Byte = 8

Private Const SOCKS5_ADDR_IPV4     As Byte = 1
Private Const SOCKS5_ADDR_DOMAIN   As Byte = 3
Private Const SOCKS5_ADDR_IPV6     As Byte = 4

' Thanks to macyui for the original implementation of this subroutine.
'   It has been heavily modified to conform with my needs.

Public Sub InitProxyConnection(ByVal ds As Winsock, ByRef ConnInfo As udtProxyConnectionInfo, ByVal RemoteIP As String, ByVal RemotePort As Long)
    ConnInfo.RemoteHost = vbNullString
    ConnInfo.RemoteHostIP = RemoteIP
    ConnInfo.RemotePort = RemotePort
    
    If Not IsValidIPAddress(ConnInfo.RemoteHostIP) Then
        ' not a valid IP - store it in hostname instead
        ConnInfo.RemoteHost = ConnInfo.RemoteHostIP
        If ConnInfo.RemoteResolveHost Then
            ' config is set to ask proxy server, use SOCKS4a/SOCKS5 address type domain
            ConnInfo.RemoteHostIP = "0.0.0.255"
        Else
            ' attempt to resolve it ourself
            ConnInfo.RemoteHostIP = ResolveHost(ConnInfo.RemoteHostIP)
        End If
    End If
    
    If LenB(ConnInfo.RemoteHostIP) = 0 Then
        UpdateProxyStatus ConnInfo, psNotConnected, PROXY_CLIENT_ERROR, WSAGetLastError()
        frmChat.DoDisconnect
        Exit Sub
    End If
    
    If ConnInfo.Version = 5 Then
        UpdateProxyStatus ConnInfo, psRequestingMethod
        SEND_SOCKS5_GREET ds
    Else
        UpdateProxyStatus ConnInfo, psRequestingConn, PROXY_REQUESTING_CONN
        SEND_SOCKS4_CONN ds, ConnInfo.RemoteHostIP, ConnInfo.RemotePort, ConnInfo.RemoteHost
    End If
End Sub

Public Sub ProxyRequestSuccess(ByRef ConnInfo As udtProxyConnectionInfo)
    Select Case ConnInfo.serverType
        Case stBNCS
            frmChat.InitBNetConnection
        Case stBNLS
            frmChat.InitBNLSConnection
        Case stMCP
            frmChat.InitMCPConnection
    End Select
End Sub

Public Sub SEND_SOCKS4_CONN(ByVal ds As Winsock, ByVal IP As String, ByVal Port As Long, ByVal HostName As String)
    Dim UseDom    As Boolean
    Dim lIP       As Long
    Dim pBuff     As clsDataBuffer
    
    Set pBuff = New clsDataBuffer
    
    lIP = inet_addr(IP)
    UseDom = (StrComp(IP, "0.0.0.255", vbBinaryCompare) = 0)
    HostName = Left$(HostName, 255)
    
    With pBuff
        .InsertByte SOCKS4_VER        ' (BYTE) Version [4]
        .InsertByte SOCKS_REQ_TCPCONN ' (BYTE) Request [Connect]
        .InsertWord htons(Port)       ' (WORD) Port
        .InsertDWord lIP              ' (DWORD) IP
        .InsertNTString vbNullString  ' (STRING) identd value [null]
        If UseDom Then
            .InsertNTString HostName  ' (STRING) hostname (if V4a)
        End If
    End With
    
    Call ProxySendPacket(ds, pBuff)
    
    Set pBuff = Nothing
End Sub

Public Sub SEND_SOCKS5_GREET(ByVal ds As Winsock)
    Dim pBuff As clsDataBuffer
    
    Set pBuff = New clsDataBuffer
    
    With pBuff
        .InsertByte SOCKS5_VER          ' (BYTE) Version [5]
        .InsertByte 2                   ' (BYTE) Auth methods supported count
        .InsertByte SOCKS5_MET_ANON     ' (BYTE)[] method #1: 0 (NO AUTH)
        .InsertByte SOCKS5_MET_USERPASS ' (BYTE)[] method #2: 2 (USER/PASS)
    End With
        
    Call ProxySendPacket(ds, pBuff)
    
    Set pBuff = Nothing
End Sub

Public Sub SEND_SOCKS5_LOGON(ByVal ds As Winsock, ByVal Username As String, ByVal Password As String)
    Dim pBuff As clsDataBuffer
    
    Username = Left$(Username, 255)
    Password = Left$(Password, 255)
    
    Set pBuff = New clsDataBuffer
    
    With pBuff
        .InsertByte SOCKS5_VER_USERPASS ' (BYTE) Version [1]
        .InsertByte Len(Username)       ' (BYTE) Username length
        .InsertNonNTString Username     ' (STRING) Username
        .InsertByte Len(Password)       ' (BYTE) Password length
        .InsertNonNTString Password     ' (STRING) Password
    End With
    
    Call ProxySendPacket(ds, pBuff)

    Set pBuff = Nothing
End Sub

Public Sub SEND_SOCKS5_CONN(ByVal ds As Winsock, ByVal IP As String, ByVal Port As Long, ByVal HostName As String)
    Dim UseDom    As Boolean
    Dim lIP       As Long
    Dim pBuff     As clsDataBuffer
    
    Set pBuff = New clsDataBuffer
    
    lIP = inet_addr(IP)
    UseDom = (StrComp(IP, "0.0.0.255", vbBinaryCompare) = 0)
    HostName = Left$(HostName, 255)
    
    With pBuff
        .InsertByte SOCKS5_VER        ' (BYTE) Version [4]
        .InsertByte SOCKS_REQ_TCPCONN ' (BYTE) Request [Connect]
        .InsertByte 0                 ' (BYTE) null
        If UseDom Then
            .InsertByte SOCKS5_ADDR_DOMAIN ' (BYTE) Address type: Domain [3]
            .InsertByte Len(HostName)      ' (BYTE) Domain length
            .InsertNonNTString HostName    ' (STRING) Domain
        Else
            .InsertByte SOCKS5_ADDR_IPV4   ' (BYTE) Address type: IPv4 [1]
            .InsertDWord lIP               ' (DWORD) IP
        End If
        .InsertWord htons(Port)       ' (WORD) Port
    End With
    
    Call ProxySendPacket(ds, pBuff)
    
    Set pBuff = Nothing
End Sub

Private Sub ProxySendPacket(ByVal ds As Winsock, ByVal pBuff As clsDataBuffer)
    ds.SendData pBuff.GetDataAsByteArr
    
    Call CachePacket(stPROXY, CtoS, 0, pBuff.length, pBuff.GetDataAsByteArr)
    Call WritePacketData(stPROXY, CtoS, 0, pBuff.length, pBuff.GetDataAsByteArr)
End Sub

Public Sub ProxyRecvPacket(ByVal ds As Winsock, ByRef ConnInfo As udtProxyConnectionInfo, ByVal pBuff As clsDataBuffer)
    Dim Status    As Byte
    Dim Method    As Byte
    Dim AddrType  As Byte
    Dim DomainLen As Byte
    
    Call CachePacket(stPROXY, StoC, 0, pBuff.length, pBuff.GetDataAsByteArr)
    Call WritePacketData(stPROXY, StoC, 0, pBuff.length, pBuff.GetDataAsByteArr)
    
    If ConnInfo.Version = 5 Then
        ' three possible packets: method, logon, and request
        Select Case ConnInfo.Status
            Case psRequestingMethod
                pBuff.GetByte          ' (BYTE) server version (5)
                Method = pBuff.GetByte ' (BYTE) method
                Select Case Method
                    Case SOCKS5_MET_ANON ' 0
                        UpdateProxyStatus ConnInfo, psRequestingConn, PROXY_REQUESTING_CONN
                        SEND_SOCKS5_CONN ds, ConnInfo.RemoteHostIP, ConnInfo.RemotePort, ConnInfo.RemoteHost
                        
                    Case SOCKS5_MET_USERPASS ' 2
                        If LenB(ConnInfo.Username) > 0 Or LenB(ConnInfo.Password) > 0 Then
                            UpdateProxyStatus ConnInfo, psLoggingOn, PROXY_LOGGING_ON
                            SEND_SOCKS5_LOGON ds, ConnInfo.Username, ConnInfo.Password
                        Else
                            UpdateProxyStatus ConnInfo, psNotConnected, PROXY_METHOD_ERROR, Method
                            frmChat.DoDisconnect
                        End If
                        
                    Case Else
                        UpdateProxyStatus ConnInfo, psNotConnected, PROXY_METHOD_ERROR, Method
                        frmChat.DoDisconnect
                End Select
                
            Case psLoggingOn
                pBuff.GetByte          ' (BYTE) u/p version (1)
                Status = pBuff.GetByte ' (BYTE) status
                If Status = SOCKS5_REQ_SUCCESS Then
                    UpdateProxyStatus ConnInfo, psRequestingConn, PROXY_LOGON_SUCCESS
                    UpdateProxyStatus ConnInfo, psRequestingConn, PROXY_REQUESTING_CONN
                    SEND_SOCKS5_CONN ds, ConnInfo.RemoteHostIP, ConnInfo.RemotePort, ConnInfo.RemoteHost
                Else
                    UpdateProxyStatus ConnInfo, psNotConnected, PROXY_LOGON_ERROR, Status
                    frmChat.DoDisconnect
                End If
                
            Case psRequestingConn
                pBuff.GetByte            ' (BYTE) server version (5)
                Status = pBuff.GetByte   ' (BYTE) status
                pBuff.GetByte            ' (BYTE) null
                AddrType = pBuff.GetByte ' (BYTE) address type
                Select Case AddrType
                    Case SOCKS5_ADDR_IPV4
                        pBuff.GetDWORD
                    Case SOCKS5_ADDR_DOMAIN
                        DomainLen = pBuff.GetByte
                        Call pBuff.GetRaw(DomainLen)
                    Case SOCKS5_ADDR_IPV6
                        Call pBuff.GetRaw(16)
                End Select
                pBuff.GetWord
        
                If Status = SOCKS5_REQ_SUCCESS Then
                    UpdateProxyStatus ConnInfo, psOnline, PROXY_REQUEST_SUCCESS
                    ProxyRequestSuccess ConnInfo
                Else
                    UpdateProxyStatus ConnInfo, psNotConnected, PROXY_REQUEST_ERROR, Status
                    frmChat.DoDisconnect
                End If
                
        End Select
    Else
        ' only possible packet is the response to the proxy request...
        pBuff.GetByte          ' (BYTE) null
        Status = pBuff.GetByte ' (BYTE) status
        pBuff.GetWord          ' (WORD) port [unused]
        pBuff.GetDWORD         ' (DWORD) ip (unused)
        
        If Status = SOCKS4_REQ_SUCCESS Then
            UpdateProxyStatus ConnInfo, psOnline, PROXY_REQUEST_SUCCESS
            ProxyRequestSuccess ConnInfo
        Else
            UpdateProxyStatus ConnInfo, psNotConnected, PROXY_REQUEST_ERROR, Status
            frmChat.DoDisconnect
        End If
    End If
    
    pBuff.Clear
End Sub

Private Sub UpdateProxyStatus(ByRef ConnInfo As udtProxyConnectionInfo, ByVal NewStatus As enuProxyStatus, Optional ByVal AddtlInfo As Long = 0, Optional ByVal ServerStatus As Long = 0)
    ConnInfo.Status = NewStatus
    
    Dim sOut As String
    Dim lColor As Long
    Dim sHost As String
    Dim sTitle As String
    
    Select Case AddtlInfo
        Case PROXY_CLIENT_ERROR
            lColor = RTBColors.ErrorMessageText
            sOut = "Unable to resolve hostname. You may ask the proxy server to resolve the hostname with ProxyServerResolve=Y. Error code: 0x" & Right(String$(8, "0") & Hex(ServerStatus), 8)
            
        Case PROXY_METHOD_ERROR
            lColor = RTBColors.ErrorMessageText
            sOut = "Method error: "
            Select Case ServerStatus
                Case SOCKS5_MET_USERPASS
                    sOut = sOut & "server is requesting a username and password but none is saved. Please set the ProxyUsername and ProxyPassword settings and try again."
                Case SOCKS5_MET_NONE
                    sOut = sOut & "server does not support the authentication methods requested."
                Case Else
                    sOut = sOut & "server is requesting an unknown/unhandled authentication method: 0x" & Right(String$(2, "0") & Hex(ServerStatus), 2)
            End Select
            
        Case PROXY_LOGGING_ON
            lColor = RTBColors.InformationText
            sOut = "Logging on..."
            
        Case PROXY_LOGON_ERROR
            lColor = RTBColors.ErrorMessageText
            sOut = "Logon error: "
            
        Case PROXY_LOGON_SUCCESS
            lColor = RTBColors.SuccessText
            sOut = "Logon successful."
        
        Case PROXY_REQUESTING_CONN
            lColor = RTBColors.InformationText
            sHost = ConnInfo.RemoteHost
            If LenB(sHost) = 0 Then sHost = ConnInfo.RemoteHostIP
            Select Case ConnInfo.serverType
                Case stBNCS
                    sOut = "Requesting connection to the Battle.net server at " & sHost & "..."
                Case stBNLS
                    sOut = "Requesting connection to the BNLS server at " & sHost & "..."
                Case stMCP
                    sTitle = ds.MCPHandler.RealmServerTitle(ds.MCPHandler.RealmServerConnected)
                    sOut = StringFormat("Requesting connection to the Diablo II Realm {0} at {1}:{2}...", sTitle, sHost, ConnInfo.RemotePort)
            End Select
            
        Case PROXY_REQUEST_ERROR
            lColor = RTBColors.ErrorMessageText
            sOut = "Request rejected: "
            If BotVars.ProxyIsSocks5 Then
                Select Case ServerStatus
                    Case SOCKS5_REQ_SUCCESS '00' succeeded
                        sOut = sOut & "SOCKS5 request granted."
                    Case SOCKS5_REQ_FAIL '01' general SOCKS server failure
                        sOut = sOut & "general server failure."
                    Case SOCKS5_REQ_RULESET '02' connection not allowed by ruleset
                        sOut = sOut & "connection not allowed by server ruleset."
                    Case SOCKS5_REQ_NONET '03' Network unreachable
                        sOut = sOut & "destination network unreachable."
                    Case SOCKS5_REQ_NOHOST '04' Host unreachable
                        sOut = sOut & "destination host unreachable."
                    Case SOCKS5_REQ_REFUSED '05' Connection refused
                        sOut = sOut & "connection refused."
                    Case SOCKS5_REQ_TTLEXP '06' TTL expired
                        sOut = sOut & "server-side TTL expiration."
                    Case SOCKS5_REQ_PROTOCOL '07' Command not supported
                        sOut = sOut & "command not supported."
                    Case SOCKS5_REQ_ADDRTYPE '08' Address type not supported
                        sOut = sOut & "address type not supported."
                    Case Else
                        sOut = sOut & "unknown/unhandled ID: 0x" & Right(String$(2, "0") & Hex(ServerStatus), 2)
                End Select
            Else
                Select Case ServerStatus
                    Case SOCKS4_REQ_SUCCESS '90
                        sOut = sOut & "SOCKS4 request granted."
                    Case SOCKS4_REQ_FAIL '91
                        sOut = sOut & "request rejected or failed."
                    Case SOCKS4_REQ_ID_RUN '92
                        sOut = sOut & "the SOCKS4 server cannot connect to identd on the client."
                    Case SOCKS4_REQ_ID_FAIL '93
                        sOut = sOut & "client program and identd report different userids."
                    Case Else
                        sOut = sOut & "unknown/unhandled ID: 0x" & Right(String$(2, "0") & Hex(ServerStatus), 2)
                End Select
            End If

        
        Case PROXY_REQUEST_SUCCESS
            lColor = RTBColors.SuccessText
            If BotVars.ProxyIsSocks5 Then
                sOut = "SOCKS5 request granted."
            Else
                sOut = "SOCKS4 request granted."
            End If
            
    End Select
    
    If LenB(sOut) > 0 Then
        Select Case ConnInfo.serverType
            Case stBNCS: frmChat.AddChat lColor, "[BNCS] [PROXY] " & sOut
            Case stBNLS: frmChat.AddChat lColor, "[BNLS] [PROXY] " & sOut
            Case stMCP:  frmChat.AddChat lColor, "[REALM] [PROXY] " & sOut
        End Select
    End If
End Sub
