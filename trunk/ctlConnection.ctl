VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ctlConnection 
   BackColor       =   &H000000FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1290
   ScaleWidth      =   1755
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wsBnls 
      Left            =   480
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsBnet 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module by hdx

Option Explicit
Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
        
        
Private Const IPPROTO_TCP            As Long = &H6
Private Const TCP_NODELAY            As Long = &H1
Private Const LOCALE_SABBREVCTRYNAME As Long = &H7
Private Const LOCALE_SENGCOUNTRY     As Long = &H1002
Private Const LOCALE_SABBREVLANGNAME As Long = &H3
Private Const LOCALE_SNATIVECTRYNAME As Long = &H8

Private Const PRODID       As Long = &H4452544C 'LTRD
Private Const PLATID       As Long = &H49583836 '68XI

Private strUsername       As String
Private strPassword       As String
Private strServer         As String
Private strBNLS           As String
Private UserList()        As String
Attribute UserList.VB_VarHelpID = -1
Private Index             As Integer
Private ServerToken       As Long
Private ClientToken       As Long
Private VersionByte       As Long
Public Errored           As Integer

Private junk As String
Private Type SYSTEMTIME
   wYear           As Integer
   wMonth          As Integer
   wDayOfWeek      As Integer
   wDay            As Integer
   wHour           As Integer
   wMinute         As Integer
   wSecond         As Integer
   wMilliseconds   As Integer
End Type
Private Type FILETIME
   dwLowDateTime  As Long
   dwHighDateTime As Long
End Type

Public Event BNLSClose()
Public Event BNETClose()
Public Event BNLSConnect()
Public Event BNETConnect()
Public Event BNLSError(ByVal Number As Integer, ByVal Description As String)
Public Event BNETError(ByVal Number As Integer, ByVal Description As String)
Public Event OnVersionCheck(ByVal result As Long, PatchFile As String)
Public Event OnLogin(ByVal Success As Boolean)
Public Event OnChatJoin(ByVal UniqueName As String)
Public Event UserInfo(ByVal Index As Integer, ByVal Username As String, ByVal Online As Boolean, ByVal Client As String, ByVal Channel As String)



Public Property Get Username() As String:           Username = strUsername: End Property
Public Property Let Username(ByVal usr As String):  strUsername = usr:      End Property
Public Property Get Password() As String:           Password = strPassword: End Property
Public Property Let Password(ByVal Pass As String): strPassword = Pass:     End Property
Public Property Get Server() As String:             Server = strServer:     End Property
Public Property Let Server(ByVal svr As String):    strServer = svr:        End Property
Public Property Get BNLS() As String:               BNLS = strBNLS:         End Property
Public Property Let BNLS(ByVal svr As String):      strBNLS = svr:          End Property
Public Property Get VerByte() As Long:              VerByte = VersionByte:  End Property
Public Property Let VerByte(ByVal vb As Long):      VersionByte = vb:       End Property

Public Sub SetList(List() As String)
    UserList = List
End Sub
Public Function GetList() As String()
    GetList = UserList
End Function
Public Function Online() As Boolean
  Online = tmr.Enabled
End Function

Public Sub Connect()
    ClientToken = GetTickCount
    Debug.Print "[BNLS] Connecting " & strBNLS & ":9367"
    wsBnls.Close
    wsBnls.Connect strBNLS, 9367
End Sub

Public Sub Disconnect()
    If wsBnls.State = sckConnected Then RaiseEvent BNLSClose
    If wsBnet.State = sckConnected Then RaiseEvent BNETClose
    wsBnls.Close
    wsBnet.Close
    tmr.Enabled = False
    Index = 0
End Sub

Private Sub tmr_Timer()
    If UBound(UserList) < Index Then Index = 0
    Call Send0x0E("/whereis " & UserList(Index))
End Sub

Private Sub UserControl_Initialize()
    Dim optval  As Boolean
    optval = False
    Call setsockopt(wsBnet.SocketHandle, IPPROTO_TCP, TCP_NODELAY, VarPtr(optval), LenB(optval))
    Call setsockopt(wsBnls.SocketHandle, IPPROTO_TCP, TCP_NODELAY, VarPtr(optval), LenB(optval))
    VersionByte = &H2A
    ReDim UserList(0)
End Sub

Private Sub wsBnet_Close()
    Debug.Print "[BNET] Closed."
    RaiseEvent BNLSClose
    tmr.Enabled = False
    Index = 0
    wsBnet.Close
End Sub

Private Sub wsBnet_Connect()
    Debug.Print "[BNET] Connected."
    RaiseEvent BNETConnect
    wsBnet.SendData Chr(1)
    Call Send0x1E
    Call Send0x12
    Call Send0x06
End Sub

Private Sub wsBnet_DataArrival(ByVal bytesTotal As Long)
    Static inBuff As String
    Dim strTmp As String
    strTmp = String(bytesTotal, Chr$(&H0))
    wsBnet.GetData strTmp, vbString, bytesTotal
    inBuff = inBuff & strTmp
    Do While Len(inBuff) >= 4
        Dim length As Long, ID As Long
        If Left(inBuff, 1) <> Chr$(&HFF) Then
          Debug.Print "Dumping ban packet!"
          Debug.Print DebugOutput(inBuff)
          inBuff = vbNullString
          Exit Sub
        End If
        Dim hi, lo As Integer
        lo = Asc(Mid(inBuff, 4, 1))
        hi = Asc(Mid(inBuff, 3, 1))
        length = Val("&H" & Hex(lo) & Right("00" & Hex(hi), 2))
        If Len(inBuff) < length Then Exit Sub
        ID = Asc(Mid(inBuff, 2, 1))
        Debug.Print "[BNET] Recived 0x" & Right("00" & Hex(ID), 2)
        Dim PBuffer As New clsBuff
        PBuffer.All = Left(inBuff, length)
        PBuffer.Pop = 4 'remove the header
        
        Select Case ID
            Case &H5
            
            Case &H6
                Dim MPQName As String, VS As String
                PBuffer.Pop = 8 'remove the filetime
                MPQName = PBuffer.NTString
                VS = PBuffer.NTString
                Debug.Print "MPQ Name: " & MPQName
                Debug.Print "Value String: " & VS
                Send0x18BNLS MPQName, VS
                
            Case &H7
                Dim result As Long, path As String
                result = PBuffer.DWORD
                path = PBuffer.NTString
                RaiseEvent OnVersionCheck(result, path)
                Select Case result
                    Case 0
                        Debug.Print "[BNET] Version Check failed. " & path
                        wsBnet.Close
                        wsBnls.Close
                    Case 1
                        Debug.Print "[BNET] Old game version. " & path
                        wsBnet.Close
                        wsBnls.Close
                    Case 2
                        Debug.Print "[BNET] Version Check Passed."
                        Call Send0x0BBNLS(strPassword)
                    Case 3
                        Debug.Print "[BNET] Downgrade required."
                        wsBnet.Close
                        wsBnls.Close
                End Select
                
            Case &HA
                Dim Name As String, stat As String, acct As String
                Name = PBuffer.NTString
                stat = PBuffer.NTString
                acct = PBuffer.NTString
                Debug.Print "[BNET] Entered chat as " & Name & " (" & stat & ")"
                RaiseEvent OnChatJoin(Name)
                tmr.Enabled = True
                
            Case &HF
                Dim eventID As Long, Username As String, text As String
                eventID = PBuffer.DWORD
                PBuffer.Pop = 20
                Username = PBuffer.NTString
                text = PBuffer.NTString
                Dim curName As String, Game As String, Channel As String
                curName = Split(text, Space(1))(0)
                
                If eventID = 18 Then
                    If InStr(LCase(text), " is refusing messages ") > 0 Then Exit Sub
                    If InStr(LCase(text), " is away ") > 0 Then Exit Sub
                    If InStr(LCase(text), " in the ") > 0 Then
                        Channel = Mid(text, InStr(LCase(text), " in the ") + 8)
                        Channel = Left(Channel, Len(Channel) - 1)
                    End If
                    If InStr(LCase(text), " using ") > 0 Then
                        Game = Mid(text, InStr(LCase(text), " using ") + 7)
                        Game = Left(Game, InStr(LCase(Game), " in ") - 1)
                    End If
                    RaiseEvent UserInfo(Index, curName, True, Game, Channel)
                ElseIf eventID = 19 Then
                    RaiseEvent UserInfo(Index, UserList(Index), False, vbNullString, vbNullString)
                End If
                Index = Index + 1
                
                
            Case &H1D
                PBuffer.Pop = 4 'UDP Token
                ServerToken = PBuffer.DWORD
                Debug.Print "Server token: 0x" & Hex(ServerToken)
                
            Case &H25
                Dim PBuff As New clsBuff
                PBuff.DWORD = PBuffer.DWORD
                PBuff.AddBNETHeader &H25
                SendBNET PBuff.All
                Debug.Print "[BNET] Send 0x25"
                
            Case &H29
                If PBuffer.DWORD = 1 Then
                    RaiseEvent OnLogin(True)
                    Debug.Print "[BNET] Login Successfull"
                    Call Send0x0A
                Else
                    RaiseEvent OnLogin(False)
                    Debug.Print "[BNET] Login failed"
                    wsBnet.Close
                    wsBnls.Close
                End If
            
            Case Else
                Debug.Print "Unhandeled packet 0x" & Right("00" & Hex(ID), 2)
        End Select
        
        
        inBuff = Mid(inBuff, length + 1)
    Loop
End Sub

Private Sub wsBnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent BNETError(Number, Description)
    Debug.Print "[BNET] Error: " & Number & ": " & Description
    Errored = Errored + 1
End Sub

Private Sub wsBnls_Close()
    Debug.Print "[BNLS] Closed."
    RaiseEvent BNLSClose
    wsBnls.Close
End Sub

Private Sub wsBnls_Connect()
    RaiseEvent BNLSConnect
    Debug.Print "[BNLS] Connected."
    Debug.Print "[BNET] Connecting " & strServer & ":6112"
    wsBnet.Close
    wsBnet.Connect strServer, 6112
End Sub

Private Sub wsBnls_DataArrival(ByVal bytesTotal As Long)
    Static inBuff As String
    Dim strTmp As String
    strTmp = String(bytesTotal, Chr$(&H0))
    wsBnls.GetData strTmp, vbString, bytesTotal
    inBuff = inBuff & strTmp
    Do While Len(inBuff) >= 3
        Dim length As Long, ID As Long
        length = Val("&H" & Hex(Asc(Mid(inBuff, 2, 1))) & Hex(Asc(Mid(inBuff, 1, 1))))
        If Len(inBuff) < length Then Exit Sub
        ID = Asc(Mid(inBuff, 3, 1))
        
        Debug.Print "[BNLS] Recived 0x" & Right("00" & Hex(ID), 2)
        
        Dim PBuffer As New clsBuff
        PBuffer.All = Left(inBuff, length)
        PBuffer.Pop = 3
        
        Select Case ID
            Case &HB
                Send0x29 PBuffer.void
        
            Case &H18
                Dim exeVer As Long, check As Long, info As String, vb As Long
                PBuffer.Pop = 4
                exeVer = PBuffer.DWORD
                check = PBuffer.DWORD
                info = PBuffer.NTString
                PBuffer.Pop = 4
                VersionByte = PBuffer.DWORD
                Debug.Print "EXE Version: 0x" & Right("0000000" & Hex(exeVer), 8)
                Debug.Print "Checksum: 0x" & Right("0000000" & Hex(check), 8)
                Debug.Print "EXE Info: " & info
                Debug.Print "VerByte: 0x" & Right("00" & Hex(VersionByte), 2)
                Call Send0x07(exeVer, check, info, vb)
                
            Case Else
                Debug.Print "[BNLS] Unhandeled Packet 0x" & Right("00" & Hex(ID), 2)
                
        End Select
        inBuff = Mid(inBuff, length + 1)
    Loop
End Sub

Private Sub wsBnls_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent BNETError(Number, Description)
    Debug.Print "[BNLS] Error: " & Number & ": " & Description
End Sub

Private Sub Send0x06()
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = PLATID
        .DWORD = PRODID
        .DWORD = VersionByte
        .DWORD = &H0
        .AddBNETHeader 6
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x06"
End Sub

Private Sub Send0x07(ver As Long, check As Long, info As String, vb As Long)
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = PLATID
        .DWORD = PRODID
        .DWORD = VersionByte
        .DWORD = ver
        .DWORD = check
        .NTString = info
        .AddBNETHeader &H7
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x07"
End Sub

Private Sub Send0x0A()
    Dim PBuffer As New clsBuff
    With PBuffer
        .NTString = strUsername
        .NTString = "LTRD 0 0 0 0 0 0 0 0 LTRD"
        .AddBNETHeader &HA
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x0A"
End Sub

Private Sub Send0x0C()
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = 2
        .NTString = "The Void"
        .AddBNETHeader &HC
        SendBNET .All
    End With
    Debug.Print "[BNET] Send 0x0C"
End Sub

Private Sub Send0x0E(text As String)
    Dim PBuffer As New clsBuff
    With PBuffer
        .NTString = text
        .AddBNETHeader &HE
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x0E"
End Sub


Private Sub Send0x12()
    Dim PBuffer As New clsBuff
    Dim LCID As Long
    LCID = GetUserDefaultLCID
    With PBuffer
        .void = GetFTString(False)
        .void = GetFTString(True)
        .DWORD = &H0
        .DWORD = GetSystemDefaultLCID
        .DWORD = LCID
        .DWORD = GetUserDefaultLangID
        .NTString = LocaleInfo(LCID, LOCALE_SABBREVLANGNAME)
        .NTString = LocaleInfo(LCID, LOCALE_SNATIVECTRYNAME)
        .NTString = LocaleInfo(LCID, LOCALE_SABBREVCTRYNAME)
        .NTString = LocaleInfo(LCID, LOCALE_SENGCOUNTRY)
        .AddBNETHeader &H12
        SendBNET .All
        Debug.Print "[BNET] Sent 0x12"
    End With
End Sub


Private Sub Send0x1E()
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = 0
        .DWORD = 0
        .DWORD = 0
        .DWORD = 0
        .DWORD = 0
        .NTString = GetCompUserName(False)
        .NTString = GetCompUserName(True)
        .AddBNETHeader &H1E
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x1E"
End Sub

Private Sub Send0x29(hash As String)
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = ClientToken
        .DWORD = ServerToken
        .void = hash
        .NTString = strUsername
        .AddBNETHeader &H29
        SendBNET .All
    End With
    Debug.Print "[BNET] Sent 0x29"
End Sub


Private Sub Send0x0BBNLS(ByVal Pass As String)
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = Len(Pass)
        .DWORD = 2
        .void = LCase(Pass)
        .DWORD = ClientToken
        .DWORD = ServerToken
        .AddBNLSHeader &HB
        SendBNLS .All
    End With
    Debug.Print "[BNLS] Sent 0x0B"
End Sub


Private Sub Send0x18BNLS(MPQ As String, Value As String)
    Dim PBuffer As New clsBuff
    With PBuffer
        .DWORD = &H9
        .DWORD = Val(Mid(MPQ, InStr(MPQ, ".") - 1, 1))
        .DWORD = 0
        .DWORD = 0
        .NTString = Value
        .AddBNLSHeader &H18
        SendBNLS .All
    End With
    Debug.Print "[BNLS] Sent 0x18"
End Sub



Private Sub SendBNET(buff As String)
    If wsBnet.State = sckConnected Then wsBnet.SendData buff
End Sub
Private Sub SendBNLS(buff As String)
    If wsBnls.State = sckConnected Then wsBnls.SendData buff
End Sub







Private Function GetFTString(Optional LocalTime As Boolean = False) As String
    Dim SyT As SYSTEMTIME, FT As FILETIME
    If LocalTime Then
        GetLocalTime SyT
    Else
        GetSystemTime SyT
    End If
    Call SystemTimeToFileTime(SyT, FT)
    Dim buff As New clsBuff
    buff.DWORD = FT.dwLowDateTime
    buff.DWORD = FT.dwHighDateTime
    GetFTString = buff.All
End Function

Private Function LocaleInfo(ByVal locale As Long, ByVal lc_type As Long) As String
    Dim length As Long, Buf As String * 1024
    length = GetLocaleInfo(locale, lc_type, Buf, Len(Buf))
    LocaleInfo = Left$(Buf, length - 1)
End Function

Private Function GetCompUserName(Optional User As Boolean = False) As String
        Dim strBuff As String, Rut As Long
        strBuff = String(255, Chr(&H0))
        Rut = IIf(User, GetUserName(strBuff, Len(strBuff)), GetComputerName(strBuff, Len(strBuff)))

        Rut = InStr(strBuff, Chr$(&H0))
        GetCompUserName = Left(strBuff, Rut - 1)
End Function

Private Function DebugOutput(ByVal sIn As String) As String
   Dim x1 As Long, y1 As Long
   Dim iLen As Long, iPos As Long
   Dim sB As String, sT As String
   Dim sOut As String
   
   iLen = Len(sIn)
   If iLen = 0 Then Exit Function
   sOut = ""
   For x1 = 0 To ((iLen - 1) \ 16)
       sB = String(48, " ")
       sT = "................"
       For y1 = 1 To 16
           iPos = 16 * x1 + y1
           If iPos > iLen Then Exit For
           Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
           Select Case Asc(Mid(sIn, iPos, 1))
           Case 32 To 255
               Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
           End Select
       Next y1
       If Len(sOut) > 0 Then sOut = sOut & vbCrLf
       sOut = sOut & sB & "  " & sT
   Next x1
   DebugOutput = sOut
End Function
