Attribute VB_Name = "modAPI"
Option Explicit

'modAPI - project StealthBot
'February 2004 - Stealth [stealth at stealthbot dot net]

Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
   
Public Declare Function SetSockOpt Lib "ws2_32.dll" Alias "setsockopt" (ByVal lSocketHandle As Long, ByVal lSocketLevel As Long, ByVal lOptName As Long, vOptVal As Any, ByVal lOptLen As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer

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

Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long

'Public Declare Function z Lib "bnetauth.dll" Alias "Z" (ByVal FileExe As String, ByVal FileStormDll As String, ByVal FileBnetDll As String, ByVal HashText As String, ByRef Version As Long, ByRef Checksum As Long, ByVal EXEInfo As String, ByVal mpqName As String) As Long
'Public Declare Function c Lib "bnetauth.dll" Alias "C" (ByVal outbuf As String, ByVal serverhash As Long, ByVal prodid As Long, ByVal val1 As Long, ByVal val2 As Long, ByVal Seed As Long) As Long '
'Public Declare Function a Lib "bnetauth.dll" Alias "A" (ByVal outbuf As String, ByVal ServerKey As Long, ByVal Password As String) As Long
'Public Declare Function A2 Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Key As Long) As Long
'Public Declare Function X Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Password As String) As Long

Public Declare Function Send Lib "ws2_32.dll" Alias "send" _
   (ByVal s As Long, _
    ByVal Buf As String, _
    ByVal datalen As Long, _
    ByVal Flags As Long) As Long
    
Public Declare Function SendBytes Lib "ws2_32.dll" Alias "send" _
   (ByVal s As Long, _
    ByRef Buf() As Byte, _
    ByVal datalen As Long, _
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

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal length As Long)
    
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Needed for the RTB scroll lock
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Integer, ByRef lpMinPos As Integer, ByRef lpMaxPos As Integer) As Boolean
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

'Public Declare Sub SRP_Init Lib "SRPx86.dll" Alias "srp_initialize" (ByRef SRP As SRP_TYPE, ByVal Username As String, ByVal Password As String)
'
'Public Declare Sub SRP_Done Lib "SRPx86.dll" Alias "srp_destroy" (ByRef SRP As SRP_TYPE)
'
'Public Declare Sub get_A Lib "SRPx86.dll" (ByRef SRP As SRP_TYPE, ByRef a() As Byte)
'Public Declare Sub get_M1 Lib "SRPx86.dll" (ByRef SRP As SRP_TYPE, ByRef s() As Byte, ByRef b() As Byte, ByRef M1() As Byte)
'Public Declare Sub get_v Lib "SRPx86.dll" (ByRef SRP As SRP_TYPE, ByRef s() As Byte, ByRef v() As Byte)
'
'Public Const BIGINT_SIZE As Long = 32
'Public Const SHA_DIGESTSIZE As Long = 20
'
'Public Type MP_DIGIT
'    digit As Integer  'I'm assuming a 16-bit datatype?
'End Type
'
'Public Type MP_INT
'    Used As Long
'    Alloc As Long
'    Sign As Long
'    dp As MP_DIGIT
'End Type
'
'Public Type SRP_TYPE
'    Username As String
'    Password As String
'    a As MP_INT
'End Type

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public Const LOCALE_SABBREVCTRYNAME As Long = &H7
Public Const LOCALE_SENGCOUNTRY     As Long = &H1002
Public Const LOCALE_SABBREVLANGNAME As Long = &H3
Public Const LOCALE_SNATIVECTRYNAME As Long = &H8
