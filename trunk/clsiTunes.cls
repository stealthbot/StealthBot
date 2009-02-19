VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsiTunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsiTunes.cls
' ...

Option Explicit

Private m_iTunesObj As Object ' ...

Private Sub Class_Initialize()

    ' ...
    If (IsLoaded() = True) Then
        Start
    End If
    
End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_iTunesObj = Nothing
    
End Sub

Public Sub Start()

    ' ...
    CreateiTunesObj

End Sub

' ...
Private Function GetWindowHandle() As Long

    Dim lRet As Long ' ...
    
    ' ...
    lRet = (FindWindow("iTunes", "iTunes"))
    
    ' ...
    GetWindowHandle = lRet
    
End Function

Private Sub CreateiTunesObj()

    On Error GoTo ERROR_HANDLER
    
    ' ...
    App.OleRequestPendingTimeout = (30 * 1000)
    
    ' ...
    Set m_iTunesObj = CreateObject("iTunes.Application")
    
    ' ...
    Exit Sub
    
ERROR_HANDLER:
    ' ...
    Set m_iTunesObj = Nothing

    ' ...
    Exit Sub
    
End Sub

Public Property Get Name() As String

    ' ...
    Name = "iTunes"
    
End Property

Public Function IsLoaded() As Boolean

    On Error GoTo ERROR_HANDLER
    
    ' ...
    If (GetWindowHandle() = 0) Then
        Exit Function
    Else
        Start
    End If

    ' ...
    IsLoaded = True
    
    ' ...
    Exit Function
    
ERROR_HANDLER:
    Exit Function
    
End Function

Public Sub PlayTrack(ByVal Track As Variant)

    ' ...
    If (IsLoaded() = False) Then
        Exit Sub
    End If
    
    If (Track = vbNullString) Then
        
        m_iTunesObj.Play
    
    ElseIf (StrictIsNumeric(Track)) Then
        
        Dim TrackCollection As Object
        
        Set TrackCollection = _
            m_iTunesObj.LibraryPlaylist.Tracks
            
        TrackCollection.Item(CInt(Track)).Play
    
    Else
    
        Dim colTracks As Object ' ...
        
        ' ...
        Set colTracks = _
            m_iTunesObj.LibraryPlaylist.Search(Track, 5)
        
        ' ...
        If (Not (colTracks Is Nothing)) Then
            Call colTracks.Item(1).Play
        End If
        
    End If

End Sub

Public Sub NextTrack()

    ' ...
    If (IsLoaded() = False) Then
        Exit Sub
    End If

    ' ...
    m_iTunesObj.NextTrack

End Sub

Public Sub PreviousTrack()

    ' ...
    If (IsLoaded() = False) Then
        Exit Sub
    End If

    ' ...
    m_iTunesObj.PreviousTrack

End Sub

Public Sub PausePlayback()

    ' ...
    If (IsPlaying() = False) Then
        Exit Sub
    End If

    ' ...
    m_iTunesObj.PlayPause

End Sub

Public Sub QuitPlayback()

    ' ...
    If (IsPlaying() = False) Then
        Exit Sub
    End If

    ' ...
    m_iTunesObj.Stop
    
End Sub

' ...
Public Property Get TrackName() As String

    ' ...
    If (IsPlaying() = False) Then
        Exit Property
    End If
    
    ' ...
    If (m_iTunesObj.CurrentTrack Is Nothing) Then
        Exit Property
    End If

    ' ...
    TrackName = _
        m_iTunesObj.CurrentTrack.Artist & " - " & m_iTunesObj.CurrentTrack.Name
        
End Property

' ...
Public Property Get PlaylistCount() As Long

    Dim TrackCollection As Object ' ...
    
    ' ...
    If (IsLoaded() = False) Then
        Exit Property
    End If

    ' ...
    Set TrackCollection = _
        m_iTunesObj.LibraryPlaylist.Tracks
        
    ' ...
    PlaylistCount = TrackCollection.Count

End Property

Public Property Get PlaylistPosition() As Long

    ' ...
    If (IsPlaying() = False) Then
        Exit Property
    End If

    ' ...
    PlaylistPosition = _
        m_iTunesObj.CurrentTrack.PlayOrderIndex

End Property

Public Property Get TrackTime() As Long

    ' ...
    If (IsPlaying() = False) Then
        Exit Property
    End If

    ' ...
    TrackTime = m_iTunesObj.PlayerPosition

End Property

Public Property Get IsPlaying() As Boolean

    ' ...
    If (IsLoaded() = False) Then
        Exit Property
    End If

    ' ...
    IsPlaying = (m_iTunesObj.PlayerState > 0)
    
End Property

' ...
Public Property Get IsPaused() As Boolean

    ' ...
    If (IsLoaded() = False) Then
        Exit Property
    End If

    ' ...
    IsPaused = False

End Property

Public Property Get TrackLength() As Long

    ' ...
    If (IsLoaded() = False) Then
        Exit Property
    End If
    
    ' ...
    If (m_iTunesObj.CurrentTrack Is Nothing) Then
        Exit Property
    End If

    ' ...
    TrackLength = m_iTunesObj.CurrentTrack.Finish

End Property

