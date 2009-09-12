Attribute VB_Name = "modCommandsMP3"
Option Explicit
'This module will house all commands related to manipluating the computer's Media player

Public Sub OnAllowMp3(Command As clsCommandObj)
    ' This command will enable or disable the use of media player-related commands.
    
    If (BotVars.DisableMP3Commands) Then
        Command.Respond "Allowing MP3 commands."
        BotVars.DisableMP3Commands = False
    Else
        Command.Respond "MP3 commands are now disabled."
        BotVars.DisableMP3Commands = True
    End If
End Sub

Public Sub OnFOS(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        MediaPlayer.FadeOutToStop
        Command.Respond "Fade-out stop."
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnLoadWinamp(Command As clsCommandObj)
    ' This command will run Winamp from the default directory, or the directory
    ' specified within the configuration file.
    Dim winamp As New clsWinamp
    
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (winamp.Start(ReadCfg("Other", "WinampPath"))) Then
        Command.Respond "Winamp loaded."
    Else
        Command.Respond "There was an error loading Winamp."
    End If
End Sub

Public Sub OnMP3(Command As clsCommandObj)

    Dim tmpbuf       As String  ' temporary output buffer
    Dim TrackName    As String  ' ...
    Dim ListPosition As Long    ' ...
    Dim ListCount    As Long    ' ...
    Dim TrackTime    As Long    ' ...
    Dim TrackLength  As Long    ' ...
    
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        If (LenB(MediaPlayer.TrackName) > 0) Then
            Command.Respond StringFormat("Current MP3 [{0}/{1}]: {2} ({3}/{4}{5})", _
              MediaPlayer.PlaylistPosition, _
              MediaPlayer.PlaylistCount, _
              MediaPlayer.TrackName, _
              SecondsToString(MediaPlayer.TrackTime), _
              SecondsToString(MediaPlayer.TrackLength), _
              IIf(MediaPlayer.IsPaused, ", paused", vbNullString) _
            )
        Else
            Command.Respond MediaPlayer.Name & " is not currently playing any media."
        End If
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnNext(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        Call MediaPlayer.NextTrack
        Command.Respond "Skipped forwards."
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnPause(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        Call MediaPlayer.PausePlayback
        Command.Respond "Paused/Resumed play."
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnPlay(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        If (LenB(Command.Argument("Song")) > 0) Then
            MediaPlayer.PlayTrack Command.Argument("Song")
            Command.Respond "Now playing track: " & MediaPlayer.TrackName
        Else
            MediaPlayer.PlayTrack
            Command.Respond "Playback started."
        End If
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnPrevious(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        MediaPlayer.PreviousTrack
        Command.Respond "Skipped backwards."
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnRepeat(Command As clsCommandObj)
    ' This command will toggle the usage of the selected media player's
    ' repeat feature.
    
    If (BotVars.DisableMP3Commands) Then Exit Sub
        
    If (MediaPlayer.IsLoaded()) Then
        If (MediaPlayer.Repeat) Then
            MediaPlayer.Repeat = False
            Command.Respond "The repeat option has been disabled for the selected media player."
        Else
            MediaPlayer.Repeat = True
            Command.Respond "The repeat option has been enabled for the selected media player."
        End If
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnSetVol(Command As clsCommandObj)
    ' This command will set the volume of the media player to the level
    ' specified by the user.
    
    Dim lngVolume As Long
    
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (Command.IsValid) Then
        If (MediaPlayer.IsLoaded()) Then
            lngVolume = Command.Argument("Volume")
            If (lngVolume < 0) Then lngVolume = 0
            If (lngVolume > 100) Then lngVolume = 100
            
            MediaPlayer.Volume = lngVolume
            Command.Respond StringFormat("Volume set to {0}%.", lngVolume)
        Else
            Command.Respond MediaPlayer.Name & " is not loaded."
        End If
    Else
        Command.Respond "Error: You must specify a volume level (0-100)."
    End If
End Sub

Public Sub OnShuffle(Command As clsCommandObj)
    ' This command will toggle the usage of the selected media player's
    ' shuffling feature.
    
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        If (MediaPlayer.Shuffle) Then
            MediaPlayer.Shuffle = False
            Command.Respond "The shuffle option has been disabled for the selected media player."
        Else
            MediaPlayer.Shuffle = True
            Command.Respond "The shuffle option has been enabled for the selected media player."
        End If
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnStop(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    If (MediaPlayer.IsLoaded()) Then
        Call MediaPlayer.QuitPlayback
        Command.Respond "Stopped playback."
    Else
        Command.Respond MediaPlayer.Name & " is not loaded."
    End If
End Sub

Public Sub OnUseiTunes(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    BotVars.MediaPlayer = "iTunes"
    Command.Respond "iTunes is ready."
    
    Call WriteINI("Other", "MediaPlayer", "iTunes")
End Sub

Public Sub OnUseWinamp(Command As clsCommandObj)
    If (BotVars.DisableMP3Commands) Then Exit Sub
    
    BotVars.MediaPlayer = "Winamp"
    Command.Respond "Winamp is ready."
    
    Call WriteINI("Other", "MediaPlayer", "Winamp")
End Sub

Private Function SecondsToString(ByVal seconds As Long) As String
    Dim temp  As Long
    Dim mins  As Long
    Dim hours As Long
    temp = seconds
    
    hours = temp \ 3600
    temp = temp - (hours * 3600)
    mins = temp \ 60
    temp = temp - (mins * 60)
    
    SecondsToString = IIf(hours, Right$("00" & hours, 2) & ":", vbNullString) & _
        Right$("00" & mins, 2) & ":" & Right$("00" & temp, 2)
End Function



