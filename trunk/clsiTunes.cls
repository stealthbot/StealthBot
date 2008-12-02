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

Private m_iTunesObj As Object

Private Sub Class_Initialize()
    ' ...
    Call CreateiTunesObj
End Sub

Private Sub Class_Terminate()
    Set m_iTunesObj = Nothing
End Sub

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
    Name = "iTunes"
End Property

Public Function IsLoaded() As Boolean
    IsLoaded = Not CBool(m_iTunesObj Is Nothing)
End Function

Public Sub PlayTrack(ByVal Track As Variant)
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If
    
    If (StrictIsNumeric(Track)) Then
        
        Dim TrackCollection As Object
        
        Set TrackCollection = _
            m_iTunesObj.LibraryPlaylist.Tracks
            
        Call TrackCollection.Item(CInt(Track)).Play
        
    ElseIf (Track = vbNullString) Then
        
        Call m_iTunesObj.Play
    
    Else
        
        Dim colTracks As Object ' ...
        
        ' ...
        Set colTracks = m_iTunesObj.LibraryPlaylist.Search(Track, 5)
        
        ' ...
        If (Not (colTracks Is Nothing)) Then
            Call colTracks.Item(1).Play
        End If
        
    End If
End Sub

Public Sub PausePlayback()
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    ' ...
    Call m_iTunesObj.PlayPause
End Sub

Public Sub QuitPlayback()
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    ' ...
    Call m_iTunesObj.Stop
End Sub

' ...
Public Property Get TrackName() As String
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    ' ...
    TrackName = m_iTunesObj.CurrentTrack.Artist & " - " & m_iTunesObj.CurrentTrack.Name
End Property

' ...
Public Property Get PlaylistCount() As Integer
    Dim TrackCollection As Object ' ...
    
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    ' ...
    Set TrackCollection = _
        m_iTunesObj.LibraryPlaylist.Tracks
        
    ' ...
    PlaylistCount = TrackCollection.Count
End Property

Public Property Get PlaylistPosition() As Long
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    PlaylistPosition = m_iTunesObj.CurrentTrack.PlayOrderIndex
End Property

Public Property Get TrackTime() As Long
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    TrackTime = m_iTunesObj.PlayerPosition
End Property

Public Property Get IsPlaying() As Boolean
    IsPlaying = (m_iTunesObj.PlayerState > 0)
End Property

' ...
Public Property Get IsPaused() As Boolean
    IsPaused = False
End Property

Public Property Get TrackLength() As Long
    ' ...
    If (Not (m_iTunesObj Is Nothing)) Then
        Call CreateiTunesObj
    End If

    ' ...
    TrackLength = m_iTunesObj.CurrentTrack.Finish
End Property

