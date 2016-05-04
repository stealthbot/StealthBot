Option Strict Off
Option Explicit On
Module modAPI
	
	'modAPI - project StealthBot
	'February 2004 - Stealth [stealth at stealthbot dot net]
	
	
	Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Integer)
	
	Public Declare Function GetForegroundWindow Lib "user32" () As Integer
    Public Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Integer, ByRef lprcClip As Integer, ByVal lpfnEnum As MonitorEnumProc_Callback, ByRef dwData As Integer) As Integer
	
    Public Declare Function SetSockOpt Lib "ws2_32.dll" Alias "setsockopt" (ByVal lSocketHandle As Integer, ByVal lSocketLevel As Integer, ByVal lOptName As Integer, ByRef vOptVal As Integer, ByVal lOptLen As Integer) As Integer
	Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Integer) As Integer
	Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Short
	
	'UPGRADE_WARNING: Structure FLASHWINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function FlashWindowEx Lib "user32" (ByRef pfwi As FLASHWINFO) As Boolean
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	
	Declare Function RegisterWindowMessage Lib "user32"  Alias "RegisterWindowMessageA"(ByVal lpString As String) As Integer
	
	Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Integer
	Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
	Public Declare Function GetLocaleInfo Lib "kernel32"  Alias "GetLocaleInfoA"(ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
	
    Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Integer, ByVal nIDEvent As Integer, ByVal uElapse As Integer, ByVal lpTimerFunc As TimerProc_Callback) As Integer
	Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Integer, ByVal nIDEvent As Integer) As Integer
	
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	
	Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Integer, ByVal wBar As Integer, ByVal bShow As Integer) As Integer
	
	'Public Declare Function z Lib "bnetauth.dll" Alias "Z" (ByVal FileExe As String, ByVal FileStormDll As String, ByVal FileBnetDll As String, ByVal HashText As String, ByRef Version As Long, ByRef Checksum As Long, ByVal EXEInfo As String, ByVal mpqName As String) As Long
	'Public Declare Function c Lib "bnetauth.dll" Alias "C" (ByVal outbuf As String, ByVal serverhash As Long, ByVal prodid As Long, ByVal val1 As Long, ByVal val2 As Long, ByVal Seed As Long) As Long '
	'Public Declare Function a Lib "bnetauth.dll" Alias "A" (ByVal outbuf As String, ByVal ServerKey As Long, ByVal Password As String) As Long
	'Public Declare Function A2 Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Key As Long) As Long
	'Public Declare Function X Lib "bnetauth.dll" (ByVal outbuf As String, ByVal Password As String) As Long
	
    Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Integer, ByVal buf() As Byte, ByVal datalen As Integer, ByVal Flags As Integer) As Integer
	
	Public Declare Function SendBytes Lib "ws2_32.dll"  Alias "send"(ByVal s As Integer, ByRef buf() As Byte, ByVal datalen As Integer, ByVal Flags As Integer) As Integer
	
	Public Declare Function DeleteURLCacheEntry Lib "Wininet.dll"  Alias "DeleteUrlCacheEntryA"(ByVal lpszUrlName As String) As Integer
	
	Public Declare Function URLDownloadToFile Lib "urlmon"  Alias "URLDownloadToFileA"(ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer
	
	Public Declare Function CallWindowProc Lib "user32"  Alias "CallWindowProcA"(ByVal lpPrevWndFunc As Integer, ByVal hWnd As Integer, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	
	'UPGRADE_NOTE: Beep was upgraded to Beep_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Declare Function Beep_Renamed Lib "kernel32" (ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Integer
	
	Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	
	Public Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	Public Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Object) As Integer
	Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	Public Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hWnd As Integer) As Integer
	Public Declare Function PostMessage Lib "user32"  Alias "PostMessageA"(ByVal hWnd As Integer, ByVal wMsg As Short, ByVal wParam As Integer, ByVal lParam As Integer) As Short
	Public Declare Function SendMessageByString Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	
	Public Declare Function FindWindowEx Lib "user32"  Alias "FindWindowExA"(ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
	
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	Public Declare Function SetCursorPos Lib "user32" (ByVal X As Integer, ByVal y As Integer) As Integer
	
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Public Declare Function SetWindowProc Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As WindowProc_Callback) As Integer
	
	Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Object, ByRef source As Object, ByVal length As Integer)
	
	Public Declare Function ShellExecute Lib "shell32"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	' Needed for the RTB scroll lock
	Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Integer, ByVal nBar As Short, ByRef lpMinPos As Short, ByRef lpMaxPos As Short) As Boolean
	Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Integer) As Integer
	
	Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	
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
	
	Public Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal sBuffer As String, ByRef lSize As Integer) As Integer
	Public Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Integer
	Public Declare Function GetFocus Lib "user" () As Short
	
	Public Const LOCALE_SABBREVCTRYNAME As Integer = &H7
	Public Const LOCALE_SENGCOUNTRY As Integer = &H1002
	Public Const LOCALE_SABBREVLANGNAME As Integer = &H3
	Public Const LOCALE_SNATIVECTRYNAME As Integer = &H8
	
	' some stuff needed for warden
    Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Object, ByVal numBytes As Integer)
	Public Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Integer, ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Public Declare Function LoadLibraryA Lib "kernel32" (ByVal strFilePath As String) As Integer
	Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
	
	Public Sub free(ByVal dwPtr As Integer)
		Dim lngHandle As Integer
		Call CopyMemory(lngHandle, dwPtr - 4, 4)
		Call GlobalUnlock(lngHandle)
		Call GlobalFree(lngHandle)
	End Sub
	
	Public Function malloc(ByVal dwSize As Integer) As Integer
		Dim lngHandle As Integer
		lngHandle = GlobalAlloc(0, dwSize + 4)
		malloc = GlobalLock(lngHandle) + 4
		Call CopyMemory(malloc - 4, lngHandle, 4)
	End Function
End Module