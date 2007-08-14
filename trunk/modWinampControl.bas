Attribute VB_Name = "modWinampControl"
'Windows API Functions
Public Const WM_COMMAND = &H111                     'Used in SendMessage call
Public Const WM_USER = &H400                        'Used in SendMessage call

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Integer
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const WA_DEFAULT_PATH As String = "C:\Program Files\Winamp\winamp.exe"

'--------------------------------------'
'         User Message Constants       '
'--------------------------------------'
'Public Const WA_GETVERSION = 0
'Public Const WA_CLEARPLAYLIST = 101
'Public Const WA_GETSTATUS = 104
'Public Const WA_GETTRACKPOSITION = 105
'Public Const WA_GETTRACKLENGTH = 105
'Public Const WA_SEEKTOPOSITION = 106
Public Const WA_SETVOLUME = 122
'Public Const WA_SETBALANCE = 123
'Public Const WA_GETEQDATA = 127
'Public Const WA_SETEQDATA = 128
Public Const WA_SENDCUSTOMDATA = 273
'--------------------------------------'
'      Command Message Constants       '
'--------------------------------------'
'Public Const IPC_SETVOLUME = 122
'Public Const WA_STOPAFTERTRACK = 40157
''Public Const WA_FASTFORWARD = 40148                 '5 secs
'Public Const WA_FASTREWIND = 40144                  '5 secs
'Public Const WA_PLAYLISTHOME = 40154
'Public Const WA_PLAYLISTEND = 40158
'Public Const WA_DIALOGOPENFILE = 40029
'Public Const WA_DIALOGOPENURL = 40155
'Public Const WA_DIALOGFILEINFO = 40188
'Public Const WA_TIMEDISPLAYELAPSED = 40037
'Public Const WA_TIMEDISPLAYREMAINING = 40038
'Public Const WA_TOGGLEPREFERENCES = 40012
'Public Const WA_DIALOGVISUALOPTIONS = 40190
'Public Const WA_DIALOGVISUALPLUGINOPTIONS = 40191
'Public Const WA_STARTVISUALPLUGIN = 40192
'Public Const WA_TOGGLEABOUT = 40041
'Public Const WA_TOGGLEAUTOSCROLL = 40189
'Public Const WA_TOGGLEALWAYSONTOP = 40019
'Public Const WA_TOGGLEWINDOWSHADE = 40064
'Public Const WA_TOGGLEPLAYLISTWINDOWSHADE = 40266
'Public Const WA_TOGGLEDOUBLESIZE = 40165
'Public Const WA_TOGGLEEQ = 40036
'Public Const WA_TOGGLEPLAYLIST = 40040
'Public Const WA_TOGGLEMAINWINDOW = 40258
'Public Const WA_TOGGLEMINIBROWSER = 40298
'Public Const WA_TOGGLEEASYMOVE = 40186
'Public Const WA_VOLUMEUP = 40058                    'increase 1%
'Public Const WA_VOLUMEDOWN = 40059                  'decrease 1%
Public Const WA_TOGGLEREPEAT = 40022
Public Const WA_TOGGLESHUFFLE = 40023
'Public Const WA_DIALOGJUMPTOTIME = 40193
Public Const WA_DIALOGJUMPTOFILE = 40194
'Public Const WA_DIALOGSKINSELECTOR = 40219
'Public Const WA_DIALOGCONFIGUREVISUALPLUGIN = 40221
'Public Const WA_RELOADSKIN = 40291
'Public Const WA_CLOSE = 40001
Public Const WM_WA_IPC = 1024

' Updated 4/10/06 by Andy (@stealthbot.net):
'   - Removed underscore replacement in the filename (Winamp 5.2x no longer allows them)
'   - Improved response from Winamp by waiting for lngEdit and lngJumpto to be drawn
'   - Added an infinite loop block
Public Sub WinampJumpToFile(ByVal strFile As String)
    Dim lngWinamp As Long
    Dim lngJumpto As Long
    Dim lngEdit As Long
    Dim lngListBox As Long
    Dim Iterations As Integer
    
    'strFile = Replace(strFile, " ", "_")
    lngWinamp = GetWinamphWnd()
    
    If lngWinamp > 0 Then
        PostMessage lngWinamp, WA_SENDCUSTOMDATA, WA_DIALOGJUMPTOFILE, 0
        
        Do
            lngJumpto = FindWindow("#32770", "Jump to file")
            lngEdit = FindWindowEx(lngJumpto, 0, "Edit", vbNullString)
            lngListBox = FindWindowEx(lngJumpto, 0, "ListBox", vbNullString)
            Iterations = Iterations + 1
            DoEvents
        Loop Until lngListBox <> 0 And lngJumpto <> 0 And lngEdit <> 0 Or Iterations > 3000
        
        SendMessageByString lngEdit, &HC, 0, strFile ' & Chr(vbKeyReturn)
        
        'PostMessage lngJumpto, 245, 0, 0
        Pause 500, True, True
        
        PostMessage lngListBox, &H203, 0, 0
        
        'If SendMessage(lngListBox, &H18B, 0, 0) = 0 Then
        '    PostMessage lngJumpto, &H10, 0, 0
        'Else
        '    SendMessage lngListBox, &H203, 0, 0
        'End If
    End If
End Sub

Public Function LoadWinamp(Optional Path As String) As String
    If Len(Path) > 0 Then
        If Dir(Path) <> vbNullString Then
            Shell Path, vbNormalFocus
            LoadWinamp = "Winamp loaded."
        Else
            LoadWinamp = "Winamp is not installed at the path you have specified in config.ini."
        End If
    Else
        If Dir(WA_DEFAULT_PATH) <> vbNullString Then
            Shell WA_DEFAULT_PATH, vbNormalFocus
            LoadWinamp = "Winamp loaded."
        Else
            LoadWinamp = "Winamp is not installed in its default folder, and you have not specified an alternate path in config.ini."
        End If
    End If
End Function
