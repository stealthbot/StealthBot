Attribute VB_Name = "modCommandsInfo"
Option Explicit
'This module will hold all of the 'Info' Commands
'Commands that return information, but have really no functionality

Public Function OnOwner(Command As clsCommandObj) As Boolean
    If (LenB(BotVars.BotOwner)) Then
        Command.Respond "This bot's owner is " & BotVars.BotOwner & "."
    Else
        Command.Respond "There is no owner currently set."
    End If
End Function

Public Function OnPing(Command As clsCommandObj) As Boolean
    Dim Latency As Long
    If (Command.IsValid) Then
        Latency = GetPing(Command.Argument("Username"))
        If (Latency >= -1) Then
            Command.Respond Command.Argument("Username") & "'s ping at login was " & Latency & "ms."
        Else
            Command.Respond "I can not see " & Command.Argument("Username") & " in the channel."
        End If
    Else
        Command.Respond "Please specify a user to ping."
    End If
End Function

Public Function OnPingMe(Command As clsCommandObj) As Boolean
    Dim Latency As Long
    If (Command.IsLocal) Then
        If (g_Online) Then
            Command.Respond "Your ping at login was " & GetPing(GetCurrentUsername) & "ms."
        Else
            Command.Respond "Error: You are not logged on."
        End If
    Else
        Latency = GetPing(Command.Username)
        If (Latency >= -1) Then
            Command.Respond "Your ping at login was " & Latency & "ms."
        Else
            Command.Respond "I can not see you in the channel."
        End If
    End If
End Function

Public Function OnTime(Command As clsCommandObj) As Boolean
    Command.Respond "The current time on this computer is " & Time & " on " & Format(Date, "MM-dd-yyyy") & "."
End Function






