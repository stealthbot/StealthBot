Attribute VB_Name = "modMediaPlayer"
' modMediaPlayer.bas
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

' fake polymorphism by returning dynamic object to client
Public Function MediaPlayer() As Object
    ' determine selected media player type
    If (StrComp(BotVars.MediaPlayer, "iTunes", vbTextCompare) = 0) Then
        ' return itunes object
        Set MediaPlayer = New clsiTunes
    Else
        ' return winamp object
        Set MediaPlayer = New clsWinamp
    End If
End Function
