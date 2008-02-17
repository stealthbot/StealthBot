VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ...

Option Explicit

' Windows API Functions
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Integer
Private Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WA_DEFAULT_PATH As String = "C:\Program Files\Winamp\winamp.exe"

Private Const WM_COMMAND = &H111                     'Used in SendMessage call
Private Const WM_USER = &H400                        'Used in SendMessage call
Private Const PROCESS_VM_READ = (&H10)

'--------------------------------------'
'         User Message Constants       '
'--------------------------------------'
Private Const WA_GETVERSION = 0
Private Const WA_CLEARPLAYLIST = 101
Private Const WA_GETSTATUS = 104
Private Const WA_GETTRACKPOSITION = 105
Private Const WA_GETTRACKLENGTH = 105
Private Const WA_SEEKTOPOSITION = 106
Private Const WA_SETVOLUME = 122
Private Const WA_SETBALANCE = 123
Private Const WA_GETEQDATA = 127
Private Const WA_SETEQDATA = 128
Private Const WA_SENDCUSTOMDATA = 273


'--------------------------------------'
'      Command Message Constants       '
'--------------------------------------'
Private Const WM_WA_IPC = 1024
Private Const IPC_SETPLAYLISTPOS = 121
Private Const IPC_SETVOLUME = 122
Private Const IPC_GETLISTLENGTH = 124
Private Const IPC_GETLISTPOS& = 125
Private Const IPC_GETPLAYLISTTITLE& = 212
Private Const IPC_GET_EXTENDED_FILE_INFO = 290
Private Const WA_STOPAFTERTRACK = 40157
Private Const WA_FASTFORWARD = 40148 '5 secs
Private Const WA_FASTREWIND = 40144  '5 secs
Private Const WA_PLAYLISTHOME = 40154
Private Const WA_PLAYLISTEND = 40158
Private Const WA_DIALOGOPENFILE = 40029
Private Const WA_DIALOGOPENURL = 40155
Private Const WA_DIALOGFILEINFO = 40188
Private Const WA_TIMEDISPLAYELAPSED = 40037
Private Const WA_TIMEDISPLAYREMAINING = 40038
Private Const WA_TOGGLEPREFERENCES = 40012
Private Const WA_DIALOGVISUALOPTIONS = 40190
Private Const WA_DIALOGVISUALPLUGINOPTIONS = 40191
Private Const WA_STARTVISUALPLUGIN = 40192
Private Const WA_TOGGLEABOUT = 40041
Private Const WA_TOGGLEAUTOSCROLL = 40189
Private Const WA_TOGGLEALWAYSONTOP = 40019
Private Const WA_TOGGLEWINDOWSHADE = 40064
Private Const WA_TOGGLEPLAYLISTWINDOWSHADE = 40266
Private Const WA_TOGGLEDOUBLESIZE = 40165
Private Const WA_TOGGLEEQ = 40036
Private Const WA_TOGGLEPLAYLIST = 40040
Private Const WA_TOGGLEMAINWINDOW = 40258
Private Const WA_TOGGLEMINIBROWSER = 40298
Private Const WA_TOGGLEEASYMOVE = 40186
Private Const WA_VOLUMEUP = 40058 'increase 1%
Private Const WA_VOLUMEDOWN = 40059 'decrease 1%
Private Const WA_TOGGLEREPEAT = 40022
Private Const WA_TOGGLESHUFFLE = 40023
Private Const WA_DIALOGJUMPTOTIME = 40193
Private Const WA_DIALOGJUMPTOFILE = 40194
Private Const WA_DIALOGSKINSELECTOR = 40219
Private Const WA_DIALOGCONFIGUREVISUALPLUGIN = 40221
Private Const WA_RELOADSKIN = 40291
Private Const WA_CLOSE = 40001
Private Const WA_PREVTRACK   As Long = 40044 ' ...
Private Const WA_NEXTTRACK   As Long = 40048 ' ...
Private Const WA_PLAY        As Long = 40045 ' ...
Private Const WA_PAUSE       As Long = 40046 ' ...
Private Const WA_STOP        As Long = 40047 ' ...
Private Const WA_FADEOUTSTOP As Long = 40147 ' ...

Private Type extendedFileInfoStruct
    filename As String
    metadata As String
    retlen   As Integer
    ret      As String
End Type

Private m_hWnd As Long

' ...
Private Sub Class_Initialize()
    ' ...
    m_hWnd = GetWindowHandle()
End Sub

' ...
Private Function GetWindowHandle() As Long
    Dim lRet As Long ' ...
    
    ' ...
    lRet = (FindWindow("Winamp v1.x", vbNullString))
    
    ' ...
    If (lRet = False) Then
        ' ...
        lRet = (FindWindow("STUDIO", vbNullString))
    End If
    
    ' ...
    GetWindowHandle = lRet
End Function

' Winamp 2.04+ Only
Private Function GetWinampSongTitle(Optional ByVal entryNumber As Integer = 0) As String
    Dim SongTitle       As String ' ...
    Dim WinampProcessID As Long   ' ...
    Dim ProcessHandle   As Long   ' ...
    Dim SongTitleRemote As Long   ' ...
    Dim ret             As String ' ...
    
    ' ...
    If (m_hWnd = 0) Then
        ' ...
        m_hWnd = GetWindowHandle()
        
        ' ...
        If (m_hWnd = 0) Then
            Exit Function
        End If
    End If
    
    ' ...
    SongTitle = String$(1024, vbNullChar)
    
    ' ...
    If (entryNumber = 0) Then
        entryNumber = PlaylistPosition() - 1
    End If
    
    '// Get process ID
    Call GetWindowThreadProcessId(m_hWnd, WinampProcessID)
    
    '// Open process
    ProcessHandle = OpenProcess(PROCESS_VM_READ, False, WinampProcessID)
    
    ' ...
    If (ProcessHandle > 0) Then
        '// Get pointer
        SongTitleRemote = SendMessage(m_hWnd, WM_WA_IPC, entryNumber, IPC_GETPLAYLISTTITLE)
        
        If (SongTitleRemote > 0) Then
            '// Try to read it
            If (ReadProcessMemory(ProcessHandle, SongTitleRemote, SongTitle, ByVal Len(SongTitle), 0&) > 0) Then
                ret = Left$(SongTitle, InStr(1, SongTitle, Chr$(0)) - 1)
            End If
        End If
        
        ' ...
        Call CloseHandle(ProcessHandle)
    End If
    
    ' ...
    GetWinampSongTitle = ret
End Function

' ...
Public Function OpenPlayer(Optional FilePath As String) As Boolean
    If (FilePath <> vbNullString) Then
        If (Dir$(FilePath) <> vbNullString) Then
            Call Shell(FilePath, vbNormalFocus)
            
            OpenPlayer = True
        End If
    Else
        If (Dir$(WA_DEFAULT_PATH) <> vbNullString) Then
            Call Shell(WA_DEFAULT_PATH, vbNormalFocus)
            
            OpenPlayer = True
        End If
    End If
End Function

' ...
Public Sub PlayTrack(Optional Track As Integer)
    ' ...
    If (m_hWnd = 0) Then
        ' ...
        m_hWnd = GetWindowHandle()
        
        ' ...
        If (m_hWnd = 0) Then
            Exit Sub
        End If
    End If
    
    If (Track) Then
        Call QuitPlayback
        
        Call SendMessage(m_hWnd, WM_USER, Track - 1, IPC_SETPLAYLISTPOS)
    End If

    Call SendMessage(m_hWnd, WM_COMMAND, WA_PLAY, 0)
End Sub

' ...
Public Sub PausePlayback()
    Call SendMessage(m_hWnd, WM_COMMAND, WA_PAUSE, 0)
End Sub

' ...
Public Sub QuitPlayback()
    ' ...
    If (m_hWnd = 0) Then
        ' ...
        m_hWnd = GetWindowHandle()
        
        ' ...
        If (m_hWnd = 0) Then
            Exit Sub
        End If
    End If

    Call SendMessage(m_hWnd, WM_COMMAND, WA_STOP, 0)
End Sub

' Winamp 2.05+ Only
Public Property Get PlaylistPosition() As Long
    ' ...
    If (m_hWnd = 0) Then
        ' ...
        m_hWnd = GetWindowHandle()
        
        ' ...
        If (m_hWnd = 0) Then
            Exit Property
        End If
    End If
    
    ' ...
    PlaylistPosition = (SendMessage(m_hWnd, WM_WA_IPC, 0&, IPC_GETLISTPOS) + 1)
End Property

' ...
Public Property Get PlaylistCount() As Integer
    PlaylistCount = SendMessage(m_hWnd, WM_WA_IPC, 0&, IPC_GETLISTLENGTH)
End Property

' ...
Public Property Get TrackName() As String
    ' ...
    TrackName = GetWinampSongTitle()
End Property

Public Property Get TrackTime() As Long
    TrackTime = 0
End Property

Public Property Get TrackLength() As Long
    TrackLength = 0
End Property

' ...
Public Property Get IsPlaying() As Boolean

End Property
