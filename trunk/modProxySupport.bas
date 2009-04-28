Attribute VB_Name = "modProxySupport"
Option Explicit

'modProxySupport - project StealthBot
'February 2004 - by Stealth [stealth at stealthbot dot net]

Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inaddr As Long) As Long
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long

Public Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Const PROXY_LOGIN_FAILED = &H401
Public Const PROXY_LOGGING_IN = &H402
Public Const PROXY_IS_NOT_PUBLIC = &H403
Public Const PROXY_LOGIN_SUCCESS = &H404
Public Const PROXY_ATTEMPTING_LOGIN = &H405

' Thanks to macyui for the original implementation of this subroutine.
'   It has been heavily modified to conform with my needs.

Public Sub LogonToProxy(ByRef ds As Winsock, ByVal sServerIP As String, ByVal lPort As Long, ByVal Socks5 As Boolean)
    Dim INet(0 To 3) As Byte
    Dim lServer As Long
    Dim buf As String
    Dim lPort1, lPort2 As Long
    Dim HostInfo As HOSTENT
    Dim ptrIP As Long
    
    If Socks5 Then
        buf = Chr(5) & Chr(1)
    Else
        buf = Chr(4) & Chr(1) 'connect request
    End If
    
    'Do we have an IP address or a hostname?
    If Not IsValidIPAddress(sServerIP) Then
        sServerIP = ResolveHost(sServerIP)
        
        If sServerIP = vbNullString Then
            Call frmChat.AddChat(vbRed, "[PROXY] Unable to resolve hostname. Error code: 0x" & Right(String(8, "0") & Hex(WSAGetLastError()), 8))
            Call frmChat.DoDisconnect
            Exit Sub
        End If
    End If
    
    'Make the IP into a long.
    lServer = inet_addr(sServerIP)
            
    'Copy ip long into a struct with 4 byte members
    CopyMemory INet(0), lServer, 4
        
    If Socks5 Then
        'Method (0x00 No authentication)
        buf = buf & Chr(0)
        'protocol version VER
        buf = buf & Chr(5)
        'command CMD 0x01 CONNECT
        buf = buf & Chr(1)
        'reserved
        buf = buf & Chr(0)
        'IPv4 address type (0x01)
        buf = buf & Chr(1)
        
        'dest address
        buf = buf & Chr(INet(0)) & Chr(INet(1)) & Chr(INet(2)) & Chr(INet(3))
        
        lPort1 = CLng(lPort) \ 256
        lPort2 = CLng(lPort) - htons(lPort1)
        
        'dest port
        buf = buf & Chr(lPort1) & Chr(lPort2)
    Else
        lPort1 = CLng(lPort) \ 256
        lPort2 = CLng(lPort) - htons(lPort1)
        
        buf = buf & Chr(lPort1) & Chr(lPort2)
    
        buf = buf & Chr(INet(0)) & Chr(INet(1)) & Chr(INet(2)) & Chr(INet(3))
        
        buf = buf & Chr(0)
    End If
    
    ds.SendData buf
End Sub

Public Sub ParseProxyPacket(ByVal Data As String)
    Dim s0ut As String
    Dim lColor As Long
    Dim Message As Byte
    Static ReceivedMethodReply As Boolean
    
    Select Case Asc(Mid$(Data, 1, 1))
        Case &H0, &H4, &H5
            Message = Asc(Mid$(Data, 2, 1))
            
            If Not BotVars.ProxyIsSocks5 Then
                Select Case Message
                    Case 90
                        s0ut = "[PROXY] SOCKS4 request granted."
                        lColor = RTBColors.SuccessText
                        UpdateProxyStatus psOnline
                        frmChat.InitBNetConnection
                    Case 91
                        s0ut = "[PROXY] Request rejected or failed."
                        UpdateProxyStatus psNotConnected
                        lColor = RTBColors.ErrorMessageText
                    Case 92
                        s0ut = "[PROXY] Request rejected: the SOCKS server cannot connect to identd on the client."
                        UpdateProxyStatus psNotConnected
                        lColor = RTBColors.ErrorMessageText
                    Case 93
                        s0ut = "[PROXY] Request rejected: client program and identd report different userids."
                        UpdateProxyStatus psNotConnected
                        lColor = RTBColors.ErrorMessageText
                    Case Else
                        s0ut = "[PROXY] Unknown/unhandled proxy message. ID: " & Message
                        lColor = RTBColors.InformationText
                        UpdateProxyStatus psNotConnected
                        
                End Select
            
                If Message >= 91 Then
                    If frmChat.sckBNet.State <> 0 Then
                        frmChat.sckBNet.Close
                    End If
                End If
            Else
                'o  REP    Reply field:
                s0ut = "[PROXY] "
                
                Select Case Message
                    Case 0 '00' succeeded
                        If ReceivedMethodReply Then
                            ReceivedMethodReply = False
                            s0ut = s0ut & "SOCKS5 login succeeded."
                            lColor = RTBColors.SuccessText
                            
                            UpdateProxyStatus psOnline
                            
                            frmChat.InitBNetConnection
                        Else
                            ReceivedMethodReply = True
                        End If
                        
                    Case 1 '01' general SOCKS server failure
                        s0ut = s0ut & "General server failure."
                        lColor = RTBColors.ErrorMessageText
                    Case 2 '02' connection not allowed by ruleset
                        s0ut = s0ut & "Connection not allowed by server ruleset."
                        lColor = RTBColors.ErrorMessageText
                    Case 3 '03' Network unreachable
                        s0ut = s0ut & "Destination network unreachable."
                        lColor = RTBColors.ErrorMessageText
                    Case 4 '04' Host unreachable
                        s0ut = s0ut & "Destination host unreachable."
                        lColor = RTBColors.ErrorMessageText
                    Case 5 '05' Connection refused
                        s0ut = s0ut & "Connection refused."
                        lColor = RTBColors.ErrorMessageText
                    Case 6 '06' TTL expired
                        s0ut = s0ut & "Server-side TTL expiration."
                        lColor = RTBColors.ErrorMessageText
                    Case 7 '07' Command not supported
                        s0ut = s0ut & "Command not supported."
                        lColor = RTBColors.ErrorMessageText
                    Case 8 '08' Address type not supported
                        s0ut = s0ut & "Address type not supported."
                        lColor = RTBColors.ErrorMessageText
                    Case Else '09' to X'FF' unassigned
                        s0ut = s0ut & "Unknown response code."
                        lColor = RTBColors.ErrorMessageText
                End Select
                
                If Message > 0 Then UpdateProxyStatus psNotConnected
                
                s0ut = s0ut & " (" & Message & ")"
            End If
            
        Case Else
            s0ut = "[PROXY] Unknown/unhandled proxy message. Message ID: " & Asc(Mid$(Data, 2, 1)) & ", VN Byte: " & Asc(Mid$(Data, 1, 1))
            lColor = RTBColors.InformationText
            UpdateProxyStatus psNotConnected
            
    End Select
    
    If lColor > 0 Then frmChat.AddChat lColor, s0ut
End Sub

Sub UpdateProxyStatus(ByVal NewStatus As enuProxyStatus, Optional ByVal AddtlInfo As Long)
    BotVars.ProxyStatus = NewStatus
    
    Dim sOut As String
    Dim lColor As Long
    
    Select Case AddtlInfo
        Case PROXY_LOGIN_FAILED
            lColor = RTBColors.ErrorMessageText
            sOut = "[PROXY] Login failed."
        
        Case PROXY_LOGGING_IN
            lColor = RTBColors.SuccessText
            sOut = "[PROXY] Logging in..."
            
        Case PROXY_IS_NOT_PUBLIC
            lColor = RTBColors.ErrorMessageText
            sOut = "[PROXY] This proxy does not allow anonymous logons."
            
        Case PROXY_LOGIN_SUCCESS
            lColor = RTBColors.SuccessText
            sOut = "[PROXY] Login successful."
        
        Case PROXY_ATTEMPTING_LOGIN
            lColor = RTBColors.InformationText
            sOut = "[PROXY] Attempting to log onto proxy..."
            
    End Select
    
    If LenB(sOut) > 0 Then frmChat.AddChat lColor, sOut
End Sub
