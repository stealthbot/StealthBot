Option Strict Off
Option Explicit On
Module modCommandsMP3
	'This module will house all commands related to manipluating the computer's Media player
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnAllowMp3(ByRef Command_Renamed As clsCommandObj)
		' This command will enable or disable the use of media player-related commands.
		
		If (BotVars.DisableMP3Commands) Then
			Command_Renamed.Respond("Allowing MP3 commands.")
			BotVars.DisableMP3Commands = False
		Else
			Command_Renamed.Respond("MP3 commands are now disabled.")
			BotVars.DisableMP3Commands = True
		End If
		
		If Config.Mp3Commands <> (Not BotVars.DisableMP3Commands) Then
			Config.Mp3Commands = Not BotVars.DisableMP3Commands
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnFOS(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.FadeOutToStop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MediaPlayer.FadeOutToStop()
			Command_Renamed.Respond("Fade-out stop.")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnLoadWinamp(ByRef Command_Renamed As clsCommandObj)
		' This command will run Winamp from the default directory, or the directory
		' specified within the configuration file.
		Dim winamp As New clsWinamp
		
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		If (winamp.Start(Config.MediaPlayerPath)) Then
			Command_Renamed.Respond("Winamp loaded.")
		Else
			Command_Renamed.Respond("There was an error loading Winamp.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnMP3(ByRef Command_Renamed As clsCommandObj)
		
		Dim tmpbuf As String ' temporary output buffer
		Dim TrackName As String
		Dim ListPosition As Integer
		Dim ListCount As Integer
		Dim TrackTime As Integer
		Dim TrackLength As Integer
		
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Len(MediaPlayer.TrackName) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsPaused. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PlaylistCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PlaylistPosition. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Command_Renamed.Respond(StringFormat("Current MP3 [{0}/{1}]: {2} ({3}/{4}{5})", MediaPlayer.PlaylistPosition, MediaPlayer.PlaylistCount, MediaPlayer.TrackName, SecondsToString(MediaPlayer.TrackTime), SecondsToString(MediaPlayer.TrackLength), IIf(MediaPlayer.IsPaused, ", paused", vbNullString)))
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Command_Renamed.Respond(MediaPlayer.Name & " is not currently playing any media.")
            End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnNext(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.NextTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call MediaPlayer.NextTrack()
			Command_Renamed.Respond("Skipped forwards.")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPause(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PausePlayback. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call MediaPlayer.PausePlayback()
			Command_Renamed.Respond("Paused/Resumed play.")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPlay(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
            If (Len(Command_Renamed.Argument("Song")) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PlayTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MediaPlayer.PlayTrack(Command_Renamed.Argument("Song"))
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Command_Renamed.Respond("Now playing track: " & MediaPlayer.TrackName)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PlayTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MediaPlayer.PlayTrack()
                Command_Renamed.Respond("Playback started.")
            End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPrevious(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.PreviousTrack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MediaPlayer.PreviousTrack()
			Command_Renamed.Respond("Skipped backwards.")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnRepeat(ByRef Command_Renamed As clsCommandObj)
		' This command will toggle the usage of the selected media player's
		' repeat feature.
		
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Repeat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (MediaPlayer.Repeat) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Repeat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MediaPlayer.Repeat = False
				Command_Renamed.Respond("The repeat option has been disabled for the selected media player.")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Repeat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MediaPlayer.Repeat = True
				Command_Renamed.Respond("The repeat option has been enabled for the selected media player.")
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetVol(ByRef Command_Renamed As clsCommandObj)
		' This command will set the volume of the media player to the level
		' specified by the user.
		
		Dim lngVolume As Integer
		
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		If (Command_Renamed.IsValid) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (MediaPlayer.IsLoaded()) Then
				lngVolume = CInt(Command_Renamed.Argument("Volume"))
				If (lngVolume < 0) Then lngVolume = 0
				If (lngVolume > 100) Then lngVolume = 100
				
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Volume. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MediaPlayer.Volume = lngVolume
				Command_Renamed.Respond(StringFormat("Volume set to {0}%.", lngVolume))
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
			End If
		Else
			Command_Renamed.Respond("Error: You must specify a volume level (0-100).")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnShuffle(ByRef Command_Renamed As clsCommandObj)
		' This command will toggle the usage of the selected media player's
		' shuffling feature.
		
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Shuffle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (MediaPlayer.Shuffle) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Shuffle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MediaPlayer.Shuffle = False
				Command_Renamed.Respond("The shuffle option has been disabled for the selected media player.")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Shuffle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MediaPlayer.Shuffle = True
				Command_Renamed.Respond("The shuffle option has been enabled for the selected media player.")
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnStop(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsLoaded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (MediaPlayer.IsLoaded()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.QuitPlayback. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call MediaPlayer.QuitPlayback()
			Command_Renamed.Respond("Stopped playback.")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Command_Renamed.Respond(MediaPlayer.Name & " is not loaded.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUseiTunes(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		BotVars.MediaPlayer = "iTunes"
		Command_Renamed.Respond("iTunes is ready.")
		
		If Config.MediaPlayer <> BotVars.MediaPlayer Then
			Config.MediaPlayer = BotVars.MediaPlayer
			Call Config.Save()
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnUseWinamp(ByRef Command_Renamed As clsCommandObj)
		If (BotVars.DisableMP3Commands) Then Exit Sub
		
		BotVars.MediaPlayer = "Winamp"
		Command_Renamed.Respond("Winamp is ready.")
		
		If Config.MediaPlayer <> BotVars.MediaPlayer Then
			Config.MediaPlayer = BotVars.MediaPlayer
			Call Config.Save()
		End If
	End Sub
	
	Private Function SecondsToString(ByVal seconds As Integer) As String
		Dim temp As Integer
		Dim mins As Integer
		Dim hours As Integer
		temp = seconds
		
		hours = temp \ 3600
		temp = temp - (hours * 3600)
		mins = temp \ 60
		temp = temp - (mins * 60)
		
		SecondsToString = IIf(hours, Right("00" & hours, 2) & ":", vbNullString) & Right("00" & mins, 2) & ":" & Right("00" & temp, 2)
	End Function
End Module