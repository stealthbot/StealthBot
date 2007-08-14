Attribute VB_Name = "modITunes"
Option Explicit

'modITunes - ITunes Control Module
' by AndyT [ andy@stealthbot.net ]

Private iTunesObj As Object

Public Function InitITunes() As Boolean
    On Error GoTo iTunesInitFail

    If Not iTunesReady Then
        If iTunesObj Is Nothing Then
            Set iTunesObj = CreateObject("iTunes.Application")
        End If
    End If
    
    InitITunes = iTunesReady
    
iTunesInitExit:
    Exit Function
    
iTunesInitFail:
    InitITunes = False
    Resume iTunesInitExit
    
End Function

Public Function iTunesReady() As Boolean
    iTunesReady = (Not (iTunesObj Is Nothing))
End Function

'// if the file is not already present it will be added to the library.
'// full filepath expected
'// todo: add a base filepath
Public Function iTunesPlayFile(ByVal sFilePath As String) As Long
    InitITunes
    
    iTunesPlayFile = iTunesObj.PlayFile(sFilePath)
End Function

Public Function iTunesPlay() As Long
    InitITunes
    
    iTunesPlay = iTunesObj.Play()
End Function


'// using PlayPause() which toggles the paused state of the track
Public Function iTunesPause() As Long
    InitITunes
    
    iTunesPause = iTunesObj.PlayPause()
End Function

Public Function iTunesStop() As Long
    InitITunes
    
    iTunesStop = iTunesObj.Stop()
End Function

Public Function iTunesNext() As Long
    InitITunes
    
    iTunesNext = iTunesObj.NextTrack()
End Function

Public Function iTunesBack() As Long
    InitITunes
    
    iTunesBack = iTunesObj.BackTrack()
End Function


Public Sub iTunesUnready()
    Set iTunesObj = Nothing
End Sub

'Private Function GetCurrentTrack() As Object
'    Dim objTrack As Object
'
'    Call iTunesObj.CurrentTrack(objTrack)
'
'    Set GetCurrentTrack = objTrack
'End Function
'
'Public Function iTunesCurrentSongTitle() As String
'    InitITunes
'
'    Dim objTrack As Object, objPlaylist As Object
'    Dim sBuf As String, sOut As String
'
'    Set objTrack = GetCurrentTrack
'
'    If Not (objTrack Is Nothing) Then
'        sBuf = String(256, vbNullChar)
'
'        Call objTrack.Artist(sBuf)
'
'        sOut = Trim(sBuf)
'
'        iTunesCurrentSongTitle = sOut
'    End If
'End Function
