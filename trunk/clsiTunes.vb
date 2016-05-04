Option Strict Off
Option Explicit On
Friend Class clsiTunes
	' clsiTunes.cls
	
	
	Private m_iTunesObj As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		If (IsLoaded() = True) Then
			Start()
		End If
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_iTunesObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_iTunesObj = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub Start()
		CreateiTunesObj()
	End Sub
	
	Private Function GetWindowHandle() As Integer
		Dim lRet As Integer
		
		lRet = (FindWindow("iTunes", "iTunes"))
		
		GetWindowHandle = lRet
	End Function
	
	Private Sub CreateiTunesObj()
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_ISSUE: App property App.OleRequestPendingTimeout was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'App.OleRequestPendingTimeout = (30 * 1000)
		
		m_iTunesObj = CreateObject("iTunes.Application")
		
		Exit Sub
		
ERROR_HANDLER: 
		'UPGRADE_NOTE: Object m_iTunesObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_iTunesObj = Nothing
		Exit Sub
	End Sub
	
	Public ReadOnly Property Name() As String
		Get
			Name = "iTunes"
		End Get
	End Property
	
	Public ReadOnly Property TrackName() As String
		Get
			If (IsPlaying() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (m_iTunesObj.CurrentTrack Is Nothing) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackName = m_iTunesObj.CurrentTrack.Artist & " - " & m_iTunesObj.CurrentTrack.Name
		End Get
	End Property
	
	Public ReadOnly Property PlaylistCount() As Integer
		Get
			Dim TrackCollection As Object
			
			If (IsLoaded() = False) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.LibraryPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackCollection = m_iTunesObj.LibraryPlaylist.Tracks
			
			'UPGRADE_WARNING: Couldn't resolve default property of object TrackCollection.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PlaylistCount = TrackCollection.Count
		End Get
	End Property
	
	Public ReadOnly Property PlaylistPosition() As Integer
		Get
			If (IsPlaying() = False) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PlaylistPosition = m_iTunesObj.CurrentTrack.PlayOrderIndex
		End Get
	End Property
	
	Public ReadOnly Property TrackTime() As Integer
		Get
			If (IsPlaying() = False) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.PlayerPosition. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackTime = m_iTunesObj.PlayerPosition
		End Get
	End Property
	
	Public ReadOnly Property IsPlaying() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.PlayerState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IsPlaying = (m_iTunesObj.PlayerState > 0)
		End Get
	End Property
	
	Public ReadOnly Property IsPaused() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.PlayerPosition. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.PlayerState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IsPaused = ((m_iTunesObj.PlayerState = 0) And (Not (m_iTunesObj.PlayerPosition = 0)))
		End Get
	End Property
	
	Public ReadOnly Property TrackLength() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (m_iTunesObj.CurrentTrack Is Nothing) Then
				Exit Property
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackLength = m_iTunesObj.CurrentTrack.Finish
		End Get
	End Property
	
	
	Public Property Volume() As Integer
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.SoundVolume. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Volume = m_iTunesObj.SoundVolume
		End Get
		Set(ByVal Value As Integer)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.SoundVolume. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iTunesObj.SoundVolume = Value
		End Set
	End Property
	
	
	Public Property Shuffle() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Shuffle = m_iTunesObj.CurrentPlaylist.Shuffle
		End Get
		Set(ByVal Value As Boolean)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iTunesObj.CurrentPlaylist.Shuffle = Value
		End Set
	End Property
	
	
	Public Property Repeat() As Boolean
		Get
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Repeat = CBool(m_iTunesObj.CurrentPlaylist.SongRepeat)
		End Get
		Set(ByVal Value As Boolean)
			If (IsLoaded() = False) Then
				Exit Property
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.CurrentPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iTunesObj.CurrentPlaylist.SongRepeat = CShort(Value)
		End Set
	End Property
	
	Public Function IsLoaded() As Boolean
		On Error GoTo ERROR_HANDLER
		
		If (GetWindowHandle() = 0) Then
			Exit Function
		Else
			Start()
		End If
		IsLoaded = True
		
		Exit Function
		
ERROR_HANDLER: 
		Exit Function
	End Function
	
	Public Sub PlayTrack(Optional ByVal Track As Object = vbNullString)
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim TrackCollection As Object
		Dim colTracks As Object
		If (Track = vbNullString) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.Play. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iTunesObj.Play()
			'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf (StrictIsNumeric(Track)) Then 
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.LibraryPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackCollection = m_iTunesObj.LibraryPlaylist.Tracks
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Track. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TrackCollection.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TrackCollection.Item(CShort(Track)).Play()
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.LibraryPlaylist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			colTracks = m_iTunesObj.LibraryPlaylist.Search(Track, 5)
			If (Not (colTracks Is Nothing)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object colTracks.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call colTracks.Item(1).Play()
			End If
		End If
	End Sub
	
	Public Sub NextTrack()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.NextTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_iTunesObj.NextTrack()
	End Sub
	
	Public Sub PreviousTrack()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.PreviousTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_iTunesObj.PreviousTrack()
	End Sub
	
	Public Sub PausePlayback()
		If (IsPlaying() = False) Then
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.Pause. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_iTunesObj.Pause()
	End Sub
	
	Public Sub QuitPlayback()
		If (IsPlaying() = False) Then
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object m_iTunesObj.Stop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_iTunesObj.Stop()
	End Sub
	
	Public Sub FadeOutToStop()
		If (IsLoaded() = False) Then
			Exit Sub
		End If
		
		' iTunes can't fade-out, so we'll just stop it.
		QuitPlayback()
	End Sub
End Class