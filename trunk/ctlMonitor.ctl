VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ctlMonitor 
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
      Left            =   840
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsBnls 
      Left            =   480
      Top             =   480
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
Attribute VB_Name = "ctlMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'User Monitor by Hdx.
'09-24-07 - Changed to use Stealth's buffer classes. And fixed for lockdown
Option Explicit

Private Const PRODID       As Long = &H4452544C 'LTRD
Private Const PLATID       As Long = &H49583836 '68XI

Private strUsername       As String
Private strPassword       As String
Private strServer         As String
Private strBNLS           As String
Attribute strBNLS.VB_VarHelpID = -1
Private currentIndex      As Integer
Private ServerToken       As Long
Private ClientToken       As Long
Private VersionByte       As Long
Private colUserInfo       As Collection


Public Event BNLSClose()
Public Event BNETClose()
Public Event BNLSConnect()
Public Event BNETConnect()
Public Event BNLSError(ByVal Number As Integer, ByVal Description As String)
Public Event BNETError(ByVal Number As Integer, ByVal Description As String)
Public Event OnVersionCheck(ByVal result As Long, PatchFile As String)
Public Event OnLogin(ByVal Success As Boolean)
Public Event OnChatJoin(ByVal UniqueName As String)
Public Event UserInfo(user As clsFriend)

Public Sub LoadMonitorConfig()
  strUsername = ReadCFG("Monitor", "Username")
  strPassword = ReadCFG("Monitor", "Password")
  strServer = ReadCFG("Main", "Server")
  strBNLS = "Hdx.JBLS.org"
  
  VersionByte = &H2A
  Dim l As Integer
  l = Val("&h" & ReadCFG("Monitor", "Verbyte"))
  If (l > 0) Then VersionByte = l
  LoadList
End Sub

Public Sub LoadList()
    Dim l As Integer, X As Integer
    If (colUserInfo Is Nothing) Then Set colUserInfo = New Collection
    l = Val(ReadCFG("Monitor", "ListCount"))
    If (l < 1) Then Exit Sub
    For X = 1 To l
      Dim tmpUser As New clsFriend
      tmpUser.Username = ReadCFG("Monitor", "User" & X)
      colUserInfo.Add tmpUser, LCase(tmpUser.Username)
    Next X
End Sub
Public Function getList() As Collection
    Set getList = colUserInfo
End Function

Public Sub SaveList()
    Dim X As Integer
    If (colUserInfo Is Nothing) Then Exit Sub
    WriteINI "Monitor", "ListCount", colUserInfo
    For X = 0 To colUserInfo.Count
      WriteINI "Monitor", "User" & (X + 1), colUserInfo.Item(X).Username
    Next X
End Sub

Public Sub Connect()
    ClientToken = GetTickCount
    Debug.Print "[BNLS] Connecting " & strBNLS & ":9367"
    wsBnls.Close
    wsBnls.Connect strBNLS, 9368
End Sub

Public Sub Disconnect()
    If wsBnls.State = sckConnected Then RaiseEvent BNLSClose
    If wsBnet.State = sckConnected Then RaiseEvent BNETClose
    wsBnls.Close
    wsBnet.Close
    tmr.Enabled = False
    currentIndex = 1
End Sub

Private Sub tmr_Timer()
    If (currentIndex > colUserInfo.Count) Then currentIndex = 1
    Debug.Print colUserInfo.Count
    Call Send0x0E("/whereis " & colUserInfo.Item(1).Username)
End Sub

Private Sub UserControl_Initialize()
    Call SetNagelStatus(wsBnet.SocketHandle, False)
    Call SetNagelStatus(wsBnls.SocketHandle, False)
    VersionByte = &H2A
    Set colUserInfo = New Collection
End Sub

Private Sub wsBnet_Close()
    Debug.Print "[BNET] Closed."
    RaiseEvent BNLSClose
    tmr.Enabled = False
    currentIndex = 1
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
          inBuff = vbNullString
          Exit Sub
        End If
        Dim hi, lo As Integer
        lo = Asc(Mid(inBuff, 4, 1))
        hi = Asc(Mid(inBuff, 3, 1))
        length = Val("&H" & Hex(lo) & Right("00" & Hex(hi), 2))
        If Len(inBuff) < length Then Exit Sub
        ID = Asc(Mid(inBuff, 2, 1))
        'Debug.Print "[BNET] Recived 0x" & Right("00" & Hex(ID), 2)
        Dim PBuffer As New clsPacketDebuffer
        PBuffer.DebuffPacket Left(inBuff, length)
        PBuffer.Advance 4  'remove the header
    
        Select Case ID
            Case &H5
    
            Case &H6
                Dim MPQName As String, VS As String
                PBuffer.Advance 8 'remove the filetime
                MPQName = PBuffer.DebuffNTString
                VS = PBuffer.DebuffNTString
                Send0x1ABNLS MPQName, VS
    
            Case &H7
                Dim result As Long, Path As String
                result = PBuffer.DebuffDWORD
                Path = PBuffer.DebuffNTString
                RaiseEvent OnVersionCheck(result, Path)
                If (result = 2) Then
                    Send0x29
                Else
                    wsBnet.Close
                    wsBnls.Close
                End If
     
            Case &HA
                Dim Name As String, stat As String, acct As String
                Name = PBuffer.DebuffNTString
                stat = PBuffer.DebuffNTString
                acct = PBuffer.DebuffNTString
                'Debug.Print "[BNET] Entered chat as " & Name & " (" & stat & ")"
                RaiseEvent OnChatJoin(Name)
                tmr.Enabled = True
     
            Case &HF
                Dim eventID As Long, Username As String, text As String
                eventID = PBuffer.DebuffDWORD
                PBuffer.Advance 20
                Username = PBuffer.DebuffNTString
                text = PBuffer.DebuffNTString
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
                ElseIf eventID = 19 Then
                End If
                currentIndex = currentIndex + 1
      
      
            Case &H1D
                PBuffer.Advance 4  'UDP Token
                ServerToken = PBuffer.DebuffDWORD
                Debug.Print "Server token: 0x" & Hex(ServerToken)
      
            Case &H25
                Dim pBuff As New PacketBuffer
                pBuff.InsertDWORD PBuffer.DebuffDWORD
                SendBNET pBuff.GetPacket(&H25)
                Debug.Print "[BNET] Send 0x25"
      
            Case &H29
                If PBuffer.DebuffDWORD = 1 Then
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
    'Debug.Print "[BNET] Error: " & Number & ": " & Description
End Sub
Private Sub wsBnls_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent BNLSError(Number, Description)
    'Debug.Print "[BNLS] Error: " & Number & ": " & Description
End Sub

Private Sub wsBnls_Close()
    'Debug.Print "[BNLS] Closed."
    RaiseEvent BNLSClose
    wsBnls.Close
End Sub

Private Sub wsBnls_Connect()
    RaiseEvent BNLSConnect
    'Debug.Print "[BNLS] Connected."
    'Debug.Print "[BNET] Connecting " & strServer & ":6112"
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

        'Debug.Print "[BNLS] Recived 0x" & Right("00" & Hex(ID), 2)

        Dim PBuffer As New clsPacketDebuffer
        PBuffer.DebuffPacket Left(inBuff, length)
        PBuffer.Advance 3

        Select Case ID
            Case &H1A
                Dim exeVer As Long, Check As Long, info As String, vb As Long
                PBuffer.Advance 4
                exeVer = PBuffer.DebuffDWORD
                Check = PBuffer.DebuffDWORD
                info = PBuffer.DebuffNTString
                PBuffer.Advance 4
                VersionByte = PBuffer.DebuffDWORD
                'Debug.Print "EXE Version: 0x" & Right("0000000" & Hex(exeVer), 8)
                'Debug.Print "Checksum: 0x" & Right("0000000" & Hex(check), 8)
                'Debug.Print "EXE Info: " & info
                'Debug.Print "VerByte: 0x" & Right("00" & Hex(VersionByte), 2)
                Call Send0x07(exeVer, Check, info, vb)

            Case Else
                'Debug.Print "[BNLS] Unhandeled Packet 0x" & Right("00" & Hex(ID), 2)

        End Select
        inBuff = Mid(inBuff, length + 1)
    Loop
End Sub

Private Sub Send0x06()
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertDWORD PLATID
        .InsertDWORD PRODID
        .InsertDWORD VersionByte
        .InsertDWORD &H0
        SendBNET .GetPacket(&H6)
    End With
    'Debug.Print "[BNET] Sent 0x06"
End Sub

Private Sub Send0x07(ver As Long, Check As Long, info As String, vb As Long)
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertDWORD PLATID
        .InsertDWORD PRODID
        .InsertDWORD VersionByte
        .InsertDWORD ver
        .InsertDWORD Check
        .InsertNTString info
        SendBNET .GetPacket(7)
    End With
    'Debug.Print "[BNET] Sent 0x07"
End Sub

Private Sub Send0x0A()
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertNTString strUsername
        .InsertNTString "LTRD 0 0 0 0 0 0 0 0 LTRD"
        SendBNET .GetPacket(&HA)
    End With
    'Debug.Print "[BNET] Sent 0x0A"
End Sub

Private Sub Send0x0E(text As String)
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertNTString text
        SendBNET .GetPacket(&HE)
    End With
    'Debug.Print "[BNET] Sent 0x0E"
End Sub


Private Sub Send0x12()
    Dim PBuffer As New PacketBuffer
    Dim LCID As Long
    LCID = GetUserDefaultLCID
    With PBuffer
        .InsertNonNTString GetFTString(False)
        .InsertNonNTString GetFTString(True)
        .InsertDWORD &H0
        .InsertDWORD GetSystemDefaultLCID
        .InsertDWORD LCID
        .InsertDWORD GetUserDefaultLangID
        .InsertNTString LocaleInfo(LCID, LOCALE_SABBREVLANGNAME)
        .InsertNTString LocaleInfo(LCID, LOCALE_SNATIVECTRYNAME)
        .InsertNTString LocaleInfo(LCID, LOCALE_SABBREVCTRYNAME)
        .InsertNTString LocaleInfo(LCID, LOCALE_SENGCOUNTRY)
        SendBNET .GetPacket(&H12)
        'Debug.Print "[BNET] Sent 0x12"
    End With
End Sub


Private Sub Send0x1E()
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertNTString GetCompUserName(False)
        .InsertNTString GetCompUserName(True)
        SendBNET .GetPacket(&H1E)
    End With
    'Debug.Print "[BNET] Sent 0x1E"
End Sub

Private Sub Send0x29()
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertDWORD ClientToken
        .InsertDWORD ServerToken
        .InsertNonNTString doubleHashPassword(strPassword, ClientToken, ServerToken)
        .InsertNTString strUsername
        SendBNET .GetPacket(&H29)
    End With
    'Debug.Print "[BNET] Sent 0x29"
End Sub

Private Sub Send0x1ABNLS(MPQ As String, Value As String)
    Dim PBuffer As New PacketBuffer
    With PBuffer
        .InsertDWORD &H9
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertDWORD 0
        .InsertNTString MPQ
        .InsertNTString Value
        SendBNLS .GetBNLSPacket(&H1A)
    End With
    'Debug.Print "[BNLS] Sent 0x1A"
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
    Dim buff As New PacketBuffer
    GetFTString = buff.MakeDWORD(FT.dwLowDateTime) & buff.MakeDWORD(FT.dwHighDateTime)
End Function

Private Function LocaleInfo(ByVal locale As Long, ByVal lc_type As Long) As String
    Dim length As Long, buf As String * 1024
    length = GetLocaleInfo(locale, lc_type, buf, Len(buf))
    LocaleInfo = Left$(buf, length - 1)
End Function

Private Function GetCompUserName(Optional user As Boolean = False) As String
    Dim strBuff As String, Rut As Long
    strBuff = String(255, Chr(&H0))
    Rut = IIf(user, GetUserName(strBuff, Len(strBuff)), GetComputerName(strBuff, Len(strBuff)))

    Rut = InStr(strBuff, Chr$(&H0))
    GetCompUserName = Left(strBuff, Rut - 1)
End Function

