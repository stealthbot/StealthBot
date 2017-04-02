Attribute VB_Name = "modAPI"
Option Explicit

'modAPI - project StealthBot
'February 2004 - Stealth [stealth at stealthbot dot net]

Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long

Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inaddr As Long) As Long
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Public Declare Function lstrlen Lib "Kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function SetSockOpt Lib "wsock32.dll" Alias "setsockopt" (ByVal lSocketHandle As Long, ByVal lSocketLevel As Long, ByVal lOptName As Long, vOptVal As Any, ByVal lOptLen As Long) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long

Public Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long

Public Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function RegisterWindowMessage Lib "user32" Alias _
        "RegisterWindowMessageA" (ByVal lpString As String) As Long

Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

'Public Declare Function z Lib "bnetauth.dll" Alias "Z" (ByVal FileExe As String, ByVal FileStormDll As String, ByVal FileBnetDll As String, ByVal HashText As String, ByRef Version As Long, ByRef Checksum As Long, ByVal EXEInfo As String, ByVal mpqName As String) As Long
'Public Declare Function c Lib "bnetauth.dll" Alias "C" (ByVal outbuf As String, ByVal serverhash As Long, ByVal prodid As Long, ByVal val1 As Long, ByVal val2 As Long, ByVal Seed As Long) As Long '
'Public Declare Function a Lib "bnetauth.dll" Alias "A" (ByVal outbuf As String, ByVal ServerKey As Long, ByVal Password As String) As Long
'Public Declare Function A2 Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Key As Long) As Long
'Public Declare Function X Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Password As String) As Long

Public Declare Function Send Lib "ws2_32.dll" Alias "send" _
   (ByVal s As Long, _
    ByVal buf As String, _
    ByVal DataLen As Long, _
    ByVal Flags As Long) As Long
    
Public Declare Function SendBytes Lib "ws2_32.dll" Alias "send" _
   (ByVal s As Long, _
    ByRef buf() As Byte, _
    ByVal DataLen As Long, _
    ByVal Flags As Long) As Long
    
Public Declare Function DeleteURLCacheEntry Lib "Wininet.dll" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long

Public Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Public Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Integer
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Const CB_LIMITTEXT = &H141
Public Const SB_BOTTOM = 7
Public Const EM_SCROLL = &HB5
Public Const IDC_HAND = 32649&

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd _
    As Long, ByVal nIndex As Long) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
    
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Needed for the RTB scroll lock
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Integer, ByRef lpMinPos As Integer, ByRef lpMaxPos As Integer) As Boolean
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Const MAX_USERNAME_LENGTH As Long = 256

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Const LOCALE_SABBREVCTRYNAME As Long = &H7
Public Const LOCALE_SENGCOUNTRY     As Long = &H1002
Public Const LOCALE_SABBREVLANGNAME As Long = &H3
Public Const LOCALE_SNATIVECTRYNAME As Long = &H8

' some stuff needed for warden
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal numBytes As Long)
Public Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LoadLibraryA Lib "kernel32" (ByVal strFilePath As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Sub free(ByVal dwPtr As Long)
    Dim lngHandle   As Long
    Call CopyMemory(lngHandle, ByVal dwPtr - 4, 4)
    Call GlobalUnlock(lngHandle)
    Call GlobalFree(lngHandle)
End Sub

Public Function malloc(ByVal dwSize As Long) As Long
    Dim lngHandle   As Long
    lngHandle = GlobalAlloc(0, dwSize + 4)
    malloc = GlobalLock(lngHandle) + 4
    Call CopyMemory(ByVal malloc - 4, lngHandle, 4)
End Function

Public Function GetComputerLanName() As String
    Dim buff As String
    Dim Length As Long
    buff = String(MAX_COMPUTERNAME_LENGTH + 1, Chr$(0))
    Length = Len(buff)
    If (GetComputerName(buff, Length)) Then
        GetComputerLanName = Left(buff, Length)
    Else
        GetComputerLanName = vbNullString
    End If
End Function

Public Function GetComputerUsername() As String
    Dim buff As String
    Dim Length As Long
    buff = String(MAX_USERNAME_LENGTH + 1, Chr$(0))
    Length = Len(buff)
    If (GetUserName(buff, Length)) Then
        GetComputerUsername = KillNull(buff)
    Else
        GetComputerUsername = vbNullString
    End If
End Function

Public Function GetAddressFromLong(ByVal lServer As Long) As String
    Dim ptrIP    As Long
    Dim Length   As Integer
    Dim arrStr() As Byte

    ptrIP = inet_ntoa(lServer)
    Length = lstrlen(ptrIP)

    ReDim arrStr(0 To Length) ' include NT
    CopyMemory arrStr(0), ByVal ptrIP, Length ' don't copy NT

    GetAddressFromLong = NTByteArrToString(arrStr)
End Function

Public Function ResolveHost(ByVal strHostName As String) As String
    Dim lServer As Long
    Dim HostInfo As HOSTENT
    Dim ptrIP As Long
    Dim sIP As String
    
    'Do we have an IP address or a hostname?
    If Not IsValidIPAddress(strHostName) Then
        'Resolve the IP.
        lServer = gethostbyname(strHostName)

        If lServer = 0 Then
            ResolveHost = vbNullString
            Exit Function
        Else
            'Copy data to HOSTENT struct.
            CopyMemory HostInfo, ByVal lServer, Len(HostInfo)
            
            If HostInfo.h_addrtype = 2 Then
                CopyMemory ptrIP, ByVal HostInfo.h_addr_list, 4
                CopyMemory lServer, ByVal ptrIP, 4
                sIP = GetAddressFromLong(lServer)
                ResolveHost = sIP
            Else
                ResolveHost = vbNullString
                Exit Function
            End If
        End If
    Else
        ResolveHost = strHostName
    End If
End Function

Public Function IsValidIPAddress(ByVal sIn As String) As Boolean
    Dim cp As Long

    cp = inet_addr(sIn)
    IsValidIPAddress = (cp <> -1)
End Function
