Option Strict Off
Option Explicit On
Friend Class clsWinamp
	
	
	' Windows API Functions
	Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal blnheritHandle As Integer, ByVal dwAppProcessId As Integer) As Integer
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
	Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
	Private Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	Private Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hWnd As Integer) As Integer
	Private Declare Function PostMessage Lib "user32"  Alias "PostMessageA"(ByVal hWnd As Integer, ByVal wMsg As Short, ByVal wParam As Integer, ByVal lParam As Integer) As Short
	Private Declare Function SendMessageByString Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	Private Declare Function FindWindowEx Lib "user32"  Alias "FindWindowExA"(ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
	
	Private Const WA_DEFAULT_PATH As String = "C:\Program Files\Winamp\winamp.exe"
	
	Private Const WM_COMMAND As Integer = &H111 'Used in SendMessage call
	Private Const WM_USER As Integer = &H400 'Used in SendMessage call
	Private Const PROCESS_VM_READ As Integer = (&H10)
	Private Const WM_LBUTTONDBLCLK As Integer = &H203
	Private Const WM_COPYDATA As Integer = &H4A
	
	'--------------------------------------'
	'         User Message Constants       '
	'--------------------------------------'
	Private Const WA_GETVERSION As Short = 0
	Private Const WA_CLEARPLAYLIST As Short = 101
	Private Const WA_GETSTATUS As Short = 104
	Private Const WA_GETTRACKPOSITION As Short = 105
	Private Const WA_GETTRACKLENGTH As Short = 105
	Private Const WA_SEEKTOPOSITION As Short = 106
	Private Const WA_SETVOLUME As Short = 122
	Private Const WA_SETBALANCE As Short = 123
	Private Const WA_GETEQDATA As Short = 127
	Private Const WA_SETEQDATA As Short = 128
	Private Const WA_SENDCUSTOMDATA As Short = 273
	
	'--------------------------------------'
	'      Command Message Constants       '
	'--------------------------------------'
	Private Const WM_WA_IPC As Short = 1024
	Private Const IPC_SETPLAYLISTPOS As Short = 121
	Private Const IPC_SETVOLUME As Short = 122
	Private Const IPC_GETLISTLENGTH As Short = 124
	Private Const IPC_GETLISTPOS As Integer = 125
	Private Const IPC_GETPLAYLISTTITLE As Integer = 212
	Private Const IPC_GET_SHUFFLE As Short = 250
	Private Const IPC_GET_REPEAT As Short = 251
	Private Const IPC_SET_SHUFFLE As Short = 252
	Private Const IPC_SET_REPEAT As Short = 253
	Private Const IPC_GET_EXTENDED_FILE_INFO As Short = 290
	Private Const WA_STOPAFTERTRACK As Integer = 40157
	Private Const WA_FASTFORWARD As Integer = 40148 '5 secs
	Private Const WA_FASTREWIND As Integer = 40144 '5 secs
	Private Const WA_PLAYLISTHOME As Integer = 40154
	Private Const WA_PLAYLISTEND As Integer = 40158
	Private Const WA_DIALOGOPENFILE As Integer = 40029
	Private Const WA_DIALOGOPENURL As Integer = 40155
	Private Const WA_DIALOGFILEINFO As Integer = 40188
	Private Const WA_TIMEDISPLAYELAPSED As Integer = 40037
	Private Const WA_TIMEDISPLAYREMAINING As Integer = 40038
	Private Const WA_TOGGLEPREFERENCES As Integer = 40012
	Private Const WA_DIALOGVISUALOPTIONS As Integer = 40190
	Private Const WA_DIALOGVISUALPLUGINOPTIONS As Integer = 40191
	Private Const WA_STARTVISUALPLUGIN As Integer = 40192
	Private Const WA_TOGGLEABOUT As Integer = 40041
	Private Const WA_TOGGLEAUTOSCROLL As Integer = 40189
	Private Const WA_TOGGLEALWAYSONTOP As Integer = 40019
	Private Const WA_TOGGLEWINDOWSHADE As Integer = 40064
	Private Const WA_TOGGLEPLAYLISTWINDOWSHADE As Integer = 40266
	Private Const WA_TOGGLEDOUBLESIZE As Integer = 40165
	Private Const WA_TOGGLEEQ As Integer = 40036
	Private Const WA_TOGGLEPLAYLIST As Integer = 40040
	Private Const WA_TOGGLEMAINWINDOW As Integer = 40258
	Private Const WA_TOGGLEMINIBROWSER As Integer = 40298
	Private Const WA_TOGGLEEASYMOVE As Integer = 40186
	Private Const WA_VOLUMEUP As Integer = 40058 'increase 1%
	Private Const WA_VOLUMEDOWN As Integer = 40059 'decrease 1%
	Private Const WA_TOGGLEREPEAT As Integer = 40022
	Private Const WA_TOGGLESHUFFLE As Integer = 40023
	Private Const WA_DIALOGJUMPTOTIME As Integer = 40193
	Private Const WA_DIALOGJUMPTOFILE As Integer = 40194
	Private Const WA_DIALOGSKINSELECTOR As Integer = 40219
	Private Const WA_DIALOGCONFIGUREVISUALPLUGIN As Integer = 40221
	Private Const WA_RELOADSKIN As Integer = 40291
	Private Const WA_CLOSE As Integer = 40001
	Private Const WA_PREVTRACK As Integer = 40044
	Private Const WA_NEXTTRACK As Integer = 40048
	Private Const WA_PLAY As Integer = 40045
	Private Const WA_PAUSE As Integer = 40046
	Private Const WA_STOP As Integer = 40047
	Private Const WA_FADEOUTSTOP As Integer = 40147
	
	Private Structure extendedFileInfoStruct
		Dim FileName As String
		Dim metadata As String
		Dim retlen As Short
		Dim ret As String
	End Structure
	
	Private Structure COPYDATASTRUCT
		Dim dwData As Integer
		Dim cbData As Integer
		Dim lpData As String
	End Structure
	
	Private Structure FILEINFO
		Dim File As String
		Dim Index As Integer
	End Structure
	
	Private m_hWnd As Integer
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_hWnd = GetWindowHandle()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Function GetWindowHandle() As Integer
		Dim lRet As Integer
		
		lRet = (FindWindow("Winamp v1.x", vbNullString))
		
		If (lRet = False) Then
			lRet = (FindWindow("STUDIO", vbNullString))
		End If
		
		GetWindowHandle = lRet
	End Function
	
	' Winamp 2.04+ Only
	Private Function GetWinampSongTitle(Optional ByVal entryNumber As Short = 0) As String
		Dim SongTitle As String
		Dim WinampProcessID As Integer
		Dim ProcessHandle As Integer
		Dim SongTitleRemote As Integer
		Dim ret As String
		
		If (IsLoaded() = False) Then
			Exit Function
		End If
		
		SongTitle = New String(vbNullChar, 1024)
		
		If (entryNumber = 0) Then
			entryNumber = PlaylistPosition() - 1
		End If
		
		'// Get process ID
		GetWindowThreadProcessId(m_hWnd, WinampProcessID)
		
		'// Open process
		ProcessHandle = CInt(OpenProcess(PROCESS_VM_READ, False, WinampProcessID))
		
		If (ProcessHandle > 0) Then
			'// Get pointer
			SongTitleRemote = CInt(CStr(SendMessage(m_hWnd, WM_WA_IPC, entryNumber, IPC_GETPLAYLISTTITLE)))
			
			If (SongTitleRemote > 0) Then
				'// Try to read it
				If (ReadProcessMemory(ProcessHandle, SongTitleRemote, SongTitle, Len(SongTitle), 0) > 0) Then
					ret = Left(SongTitle, InStr(1, SongTitle, Chr(0)) - 1)
				End If
			End If
			
			CloseHandle(ProcessHandle)
		End If
		
		GetWinampSongTitle = ret
	End Function
	
	Public ReadOnly Property Name() As String
		Get
			Name = "Winamp"
		End Get
	End Property
	
	
	Public Property Volume() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			Volume = (CInt(SendMessage(m_hWnd, WM_WA_IPC, -666, WA_SETVOLUME) / 2.55))
		End Get
		Set(ByVal Value As Integer)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			SendMessage(m_hWnd, WM_WA_IPC, Value * 2.55, WA_SETVOLUME)
		End Set
	End Property
	
	' Winamp 2.05+ Only
	Public ReadOnly Property PlaylistPosition() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			PlaylistPosition = CInt(SendMessage(m_hWnd, WM_WA_IPC, 0, IPC_GETLISTPOS) + 1)
		End Get
	End Property
	
	Public ReadOnly Property PlaylistCount() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			PlaylistCount = CInt(SendMessage(m_hWnd, WM_WA_IPC, 0, IPC_GETLISTLENGTH))
		End Get
	End Property
	
	Public ReadOnly Property TrackName() As String
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			TrackName = GetWinampSongTitle()
		End Get
	End Property
	
	
	Public Property Shuffle() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			Shuffle = CBool(SendMessage(m_hWnd, WM_WA_IPC, 0, IPC_GET_SHUFFLE))
		End Get
		Set(ByVal Value As Boolean)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			SendMessage(m_hWnd, WM_WA_IPC, Value, IPC_SET_SHUFFLE)
		End Set
	End Property
	
	
	Public Property Repeat() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			Repeat = CBool(SendMessage(m_hWnd, WM_WA_IPC, 0, IPC_GET_REPEAT))
		End Get
		Set(ByVal Value As Boolean)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			SendMessage(m_hWnd, WM_WA_IPC, Value, IPC_SET_REPEAT)
		End Set
	End Property
	
	Public ReadOnly Property TrackTime() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			TrackTime = CInt(SendMessage(m_hWnd, WM_WA_IPC, 0, WA_GETTRACKPOSITION) / 1000)
		End Get
	End Property
	
	Public ReadOnly Property TrackLength() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			TrackLength = CInt(SendMessage(m_hWnd, WM_WA_IPC, 1, WA_GETTRACKLENGTH))
		End Get
	End Property
	
	Public ReadOnly Property IsPlaying() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			IsPlaying = CBool(SendMessage(m_hWnd, WM_WA_IPC, 0, WA_GETSTATUS) = 1)
		End Get
	End Property
	
	Public ReadOnly Property IsPaused() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			IsPaused = CBool(SendMessage(m_hWnd, WM_WA_IPC, 0, WA_GETSTATUS) = 3)
		End Get
	End Property
	
	Public Function IsLoaded() As Boolean
		m_hWnd = GetWindowHandle()
		
		If (m_hWnd = 0) Then
			IsLoaded = False
		Else
			IsLoaded = True
		End If
	End Function
	
	Public Function Start(Optional ByRef filePath As String = "") As Boolean
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If ((Dir(filePath) <> vbNullString) And (filePath <> vbNullString)) Then
			Start = True
			
			Shell(filePath, AppWinStyle.NormalFocus)
		Else
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If (Dir(WA_DEFAULT_PATH) <> vbNullString) Then
				Start = True
				
				Shell(WA_DEFAULT_PATH, AppWinStyle.NormalFocus)
			End If
		End If
	End Function
	
	Public Sub NextTrack()
		PlayTrack(PlaylistPosition + 1)
	End Sub
	
	Public Sub PreviousTrack()
		PlayTrack(PlaylistPosition - 1)
	End Sub
	
	Public Sub PlayTrack(Optional ByVal Track As Object = vbNullString)
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim lngJumpto As Integer
		Dim lngEdit As Integer
		Dim lngListBox As Integer
		Dim iterations As Short
		If (Track <> vbNullString) Then
			QuitPlayback()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrictIsNumeric(Track)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SendMessage(m_hWnd, WM_USER, Track - 1, IPC_SETPLAYLISTPOS)
			Else
				
				PostMessage(m_hWnd, WA_SENDCUSTOMDATA, WA_DIALOGJUMPTOFILE, 0)
				
				Do 
					lngJumpto = FindWindow("#32770", "Jump to file")
					
					lngEdit = FindWindowEx(lngJumpto, 0, "Edit", vbNullString)
					
					lngListBox = FindWindowEx(lngJumpto, 0, "ListBox", vbNullString)
					
					iterations = (iterations + 1)
					
					System.Windows.Forms.Application.DoEvents()
				Loop Until lngListBox <> 0 And lngJumpto <> 0 And lngEdit <> 0 Or iterations > 3000
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SendMessageByString(lngEdit, &HC, 0, Track)
				
				PostMessage(lngListBox, WM_LBUTTONDBLCLK, 0, 0)
			End If
		End If
		SendMessage(m_hWnd, WM_COMMAND, WA_PLAY, 0)
	End Sub
	
	Public Sub PausePlayback()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		SendMessage(m_hWnd, WM_COMMAND, WA_PAUSE, 0)
	End Sub
	
	Public Sub QuitPlayback()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		SendMessage(m_hWnd, WM_COMMAND, WA_STOP, 0)
	End Sub
	
	Public Sub FadeOutToStop()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		SendMessage(m_hWnd, WM_COMMAND, WA_FADEOUTSTOP, 0)
	End Sub
End Class