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
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Integer
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WA_DEFAULT_PATH As String = "C:\Program Files\Winamp\winamp.exe"

Private Const WM_COMMAND = &H111                     'Used in SendMessage call
Private Const WM_USER = &H400                        'Used in SendMessage call
Private Const PROCESS_VM_READ = (&H10)
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_COPYDATA = &H4A

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
Private Const IPC_GET_SHUFFLE = 250
Private Const IPC_GET_REPEAT = 251
Private Const IPC_SET_SHUFFLE = 252
Private Const IPC_SET_REPEAT = 253
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

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As String
End Type

Private Type FILEINFO
    file  As String
    Index As Long
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

Public Property Get Name() As String
    Name = "Winamp"
End Property

Public Function IsLoaded() As Boolean
    If (GetWindowHandle() = 0) Then
        IsLoaded = False
    Else
        IsLoaded = True
    End If
End Function

' ...
Public Function OpenPlayer(Optional FilePath As String) As Boolean
    ' ...
    If ((Dir$(FilePath) <> vbNullString) And (FilePath <> vbNullString)) Then
        ' ...
        OpenPlayer = True
        Call Shell(FilePath, vbNormalFocus)
    Else
        If (Dir$(WA_DEFAULT_PATH) <> vbNullString) Then
            ' ...
            OpenPlayer = True
            Call Shell(WA_DEFAULT_PATH, vbNormalFocus)
        End If
    End If
End Function

Public Sub NextTrack()

   PlayTrack PlaylistPosition + 1
   
End Sub

Public Sub PreviousTrack()

    PlayTrack PlaylistPosition - 1

End Sub

' ...
Public Sub PlayTrack(Optional Track As Variant = vbNullString)
    ' ...
    If (m_hWnd = 0) Then
        ' ...
        m_hWnd = GetWindowHandle()
        
        ' ...
        If (m_hWnd = 0) Then
            Exit Sub
        End If
    End If
    
    ' ...
    If (Track <> vbNullString) Then
        ' ...
        Call QuitPlayback
        
        ' ...
        If (StrictIsNumeric(Track)) Then
            ' ...
            Call SendMessage(m_hWnd, WM_USER, Track - 1, IPC_SETPLAYLISTPOS)
        Else
            Dim lngJumpto  As Long    ' ...
            Dim lngEdit    As Long    ' ...
            Dim lngListBox As Long    ' ...
            Dim iterations As Integer ' ...
   
            ' ...
            Call PostMessage(m_hWnd, WA_SENDCUSTOMDATA, WA_DIALOGJUMPTOFILE, 0)
            
            ' ...
            Do
                ' ...
                lngJumpto = FindWindow("#32770", "Jump to file")
                
                ' ...
                lngEdit = FindWindowEx(lngJumpto, 0, "Edit", vbNullString)
                
                ' ...
                lngListBox = FindWindowEx(lngJumpto, 0, "ListBox", vbNullString)
                
                ' ...
                iterations = (iterations + 1)
                
                ' ...
                DoEvents
            Loop Until lngListBox <> 0 And lngJumpto <> 0 And lngEdit <> 0 Or iterations > 3000
            
            ' ...
            Call SendMessageByString(lngEdit, &HC, 0, Track)
            
            ' ...
            Call PostMessage(lngListBox, WM_LBUTTONDBLCLK, 0, 0)
        End If
    End If

    ' ...
    Call SendMessage(m_hWnd, WM_COMMAND, WA_PLAY, 0)
End Sub

' ...
Public Property Let Volume(lngVol As Long)
    Call SendMessage(m_hWnd, WM_WA_IPC, (lngVol * 2.55), WA_SETVOLUME)
End Property

' ...
Public Property Get Volume() As Long
    Volume = SendMessage(m_hWnd, WM_WA_IPC, -666, WA_SETVOLUME)
End Property

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

Public Sub FadeOutToStop()
    Call SendMessage(m_hWnd, WM_COMMAND, WA_FADEOUTSTOP, 0)
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

' ...
Public Property Let Shuffle(blnShuffle As Boolean)
    Call SendMessage(m_hWnd, WM_WA_IPC, blnShuffle, IPC_SET_SHUFFLE)
End Property

' ...
Public Property Get Shuffle() As Boolean
    Shuffle = SendMessage(m_hWnd, WM_WA_IPC, 0&, IPC_GET_SHUFFLE)
End Property

' ...
Public Property Let Repeat(blnRepeat As Boolean)
    Call SendMessage(m_hWnd, WM_WA_IPC, blnRepeat, IPC_SET_REPEAT)
End Property

' ...
Public Property Get Repeat() As Boolean
    Repeat = SendMessage(m_hWnd, WM_WA_IPC, 0&, IPC_GET_REPEAT)
End Property

Public Property Get TrackTime() As Long
    TrackTime = (SendMessage(m_hWnd, WM_WA_IPC, 0&, WA_GETTRACKPOSITION) / 1000)
End Property

Public Property Get TrackLength() As Long
    TrackLength = SendMessage(m_hWnd, WM_WA_IPC, 1&, WA_GETTRACKLENGTH)
End Property

' ...
Public Property Get IsPlaying() As Boolean
    IsPlaying = (SendMessage(m_hWnd, WM_WA_IPC, 0, WA_GETSTATUS) = 1)
End Property

' ...
Public Property Get IsPaused() As Boolean
    IsPaused = (SendMessage(m_hWnd, WM_WA_IPC, 0, WA_GETSTATUS) = 3)
End Property
