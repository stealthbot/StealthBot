Attribute VB_Name = "modWinampTitle"
'// Original code by Yoni[vL], thanks to NSDN [Nullsoft]
'// Ported to Visual Basic 6 by Stealth (stealth@stealthbot.net)
'// Thanks to Skywing[vL], Kp and thuscelackpiss who helped me straighten out ReadProcessMemory

'// modWinampTitle.bas
Option Explicit

'// Constants
Private Const WM_USER& = &H400
Private Const WM_WA_IPC = WM_USER
Private Const IPC_GETLISTPOS& = 125
Private Const IPC_GETPLAYLISTTITLE& = 212
Private Const PROCESS_VM_READ = (&H10)

'// API declarations
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal blnheritHandle As Long, _
    ByVal dwAppProcessId As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
    ByVal lpBaseAddress As Long, _
    ByVal lpBuffer As String, ByVal nSize As Long, _
    ByRef lpNumberOfBytesWritten As Long) As Long
    
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
    ByRef lpdwProcessId As Long) As Long
    
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

'// The following are mine (Yoni's ;)
Public Function GetWinamphWnd() As Long
    Dim lRet As Long
    
    lRet = (FindWindow("Winamp v1.x", vbNullString))
    
    If lRet = 0 Then
        lRet = (FindWindow("STUDIO", vbNullString))
    End If
    
    GetWinamphWnd = lRet
End Function

'// Winamp 2.05+
Private Function GetWinampListPos(ByVal hWndWinamp As Long) As Long
    GetWinampListPos = SendMessage(hWndWinamp, WM_WA_IPC, 0&, IPC_GETLISTPOS)
End Function

'// Winamp 2.04+, pointer is in Winamp address space
Private Function GetWinampSongTitleRemote(ByVal hWndWinamp As Long, ByVal Index As Long) As Long
   GetWinampSongTitleRemote = SendMessage(hWndWinamp, WM_WA_IPC, Index, IPC_GETPLAYLISTTITLE)
End Function

Private Function GetWinampSongTitleLocal(ByVal hWndWinamp As Long, ByVal Index As Long) As String '// Winamp 2.04+, pointer is in local address space
    
    Dim SongTitle As String
    Dim WinampProcessID As Long
    'Dim WinampProcessHandle As Long
    Dim ProcessHandle As Long
    Dim SongTitleRemote As Long
    Dim Ret As String
    
    SongTitle = String(1024, vbNullChar)
    
    '// Get process ID
    Call GetWindowThreadProcessId(hWndWinamp, WinampProcessID)
    
    '// Open process
    ProcessHandle = OpenProcess(PROCESS_VM_READ, False, WinampProcessID)
    
    If (ProcessHandle > 0) Then
        '// Get pointer
        SongTitleRemote = GetWinampSongTitleRemote(hWndWinamp, Index)
        
        If (SongTitleRemote > 0) Then
            '// Try to read it
            If (ReadProcessMemory(ProcessHandle, SongTitleRemote, SongTitle, ByVal Len(SongTitle), 0&) > 0) Then
                '// Success
                Ret = Left$(SongTitle, InStr(1, SongTitle, Chr(0)) - 1)
            End If
        End If
        
        Call CloseHandle(ProcessHandle)
    End If
    
    GetWinampSongTitleLocal = Ret
End Function

Public Function GetCurrentSongTitle(Optional ShowTrackNumber As Boolean = False) As String
    Dim hWnd As Long
    Dim sBuf As String
    Dim bSuccess As Boolean
    
    hWnd = GetWinamphWnd()
    sBuf = GetWinampSongTitleLocal(hWnd, GetWinampListPos(hWnd))
    
    bSuccess = (Len(sBuf) > 0)
    
    If ShowTrackNumber And bSuccess Then
        sBuf = (GetWinampListPos(hWnd) + 1) & ": " & sBuf
    End If
    
    If Not bSuccess Then
        sBuf = "(Winamp not loaded)"
    End If
    
    If iTunesReady Then
        sBuf = "(Using iTunes)"
    End If
    
    GetCurrentSongTitle = sBuf
End Function
