Option Strict Off
Option Explicit On
Module modMediaPlayer
	' modMediaPlayer.bas
	' Copyright (C) 2008 Eric Evans
	
	
	' fake polymorphism by returning dynamic object to client
	Public Function MediaPlayer() As Object
		
		' determine selected media player type
		If (StrComp(BotVars.MediaPlayer, "iTunes", CompareMethod.Text) = 0) Then
			MediaPlayer = New clsiTunes
		Else
			MediaPlayer = New clsWinamp
		End If
		
	End Function
End Module