Attribute VB_Name = "modCommandsClan"
Option Explicit
'This modules contains all the commands that have to do with Warcraft III's Clan system

Public Function OnClan(Command As clsCommandObj) As Boolean
    ' This command will allow the use of Battle.net's /clan command without requiring
    ' users be given the ability to use the bot's say command.

    If (Command.IsValid) Then
        Select Case (LCase$(Command.Argument("SubCommand")))
            Case "public", "pub"
                If (LCase$(Left$(g_Channel.Name, 5)) = "clan ") Then
                    If (g_Channel.Self.IsOperator) Then
                        Command.Respond "The clan channel is now public."
                        Call frmChat.AddQ("/c pub", PRIORITY.CHANNEL_MODERATION_MESSAGE, Command.Username)
                    Else
                        Command.Respond "Error: The bot must have ops to change the clan privacy status."
                    End If
                Else
                    Command.Respond "Error: The bot must be in a clan channel to change the clan privacy status."
                End If
            Case "private", "priv"
                If (LCase$(Left$(g_Channel.Name, 5)) = "clan ") Then
                    If (g_Channel.Self.IsOperator) Then
                        Command.Respond "The clan channel is now private."
                        Call frmChat.AddQ("/c priv", PRIORITY.CHANNEL_MODERATION_MESSAGE, Command.Username)
                    Else
                        Command.Respond "Error: The bot must have ops to change the clan privacy status."
                    End If
                Else
                    Command.Respond "Error: The bot must be in a clan channel to change the clan privacy status."
                End If
            Case "motd"
                If (IsW3) Then
                    If (g_Clan.Self.Rank > 2) Then
                        If (LenB(Command.Argument("Message")) = 0) Then
                            Command.Respond "You must specify a message to set."
                        Else
                            Command.Respond "The clan MOTD has been set to: " & Command.Argument("Message")
                            Call frmChat.AddQ("/c motd " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
                        End If
                    ElseIf (g_Clan.Self.Rank > 0) Then
                        Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to set the MOTD"
                    End If
                End If
            Case "mail"
                If (g_Clan.Self.Rank > 2) Then
                    If (LenB(Command.Argument("Message")) = 0) Then
                        Command.Respond "You must specify a message to send."
                    Else
                        Command.Respond "An e-mail has been sent to everyone in the clan who choose to receive them."
                        Call frmChat.AddQ("/c mail " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
                    End If
                ElseIf (g_Clan.Self.Rank > 0) Then
                    Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to send Clan mail."
                End If
        End Select
    End If
End Function

Public Function OnDemote(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        If (IsW3) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon > 1) Then
                            Call DemoteMember(reverseUsername(liUser.Text), liUser.SmallIcon - 1)
                        Else
                            Command.Respond "Error: The specified user is already at the lowest demoteable ranking."
                        End If
                    Else
                        Command.Respond "Error: The specified user is not currently a member of this clan."
                    End If
                Else
                    Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to promote."
                End If
            Else
                Command.Respond "Error: I am not a member of a clan."
            End If
        End If
    End If
End Function

Public Function OnDisbandClan(Command As clsCommandObj)
    If (IsW3) Then
        If (LenB(g_Clan.Self.Name) > 0) Then
            If (g_Clan.Self.Rank >= 4) Then
                Call DisbandClan
            Else
                Command.Respond "Error: The bot must be a Chieftain to execute this command."
            End If
        Else
            Command.Respond "Error: I am not a member of a clan."
        End If
    End If
End Function

Public Function OnInvite(Command As clsCommandObj) As Boolean
    ' This command will send an invitation to the specified user to join the
    ' clan that the bot is currently either a Shaman or Chieftain of.  This
    ' command will only work if the bot is logged on using WarCraft III, and
    ' is either a Shaman, or a Chieftain of the clan in question.
    
    If (IsW3) Then
        If (g_Clan.Self.Rank >= 3) Then
            If (Command.IsValid) Then
                Call InviteToClan(Command.Argument("Username"))
                Command.Respond Command.Argument("Username") & ": Clan invitation sent."
            Else
                Command.Respond "Error: You must specify a username to invite."
            End If
        Else
            Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to invite users."
        End If
    End If
End Function

Public Function OnMakeChieftain(Command As clsCommandObj) As Boolean
    If (IsW3) Then
        If (g_Clan.Self.Rank >= 4) Then
            If (Command.IsValid) Then
                Call MakeMemberChieftain(reverseUsername(Command.Argument("Username")))
            Else
                Command.Respond "Error: You must specify a username to promote."
            End If
        Else
            Command.Respond "Error: The bot must be the Chieftain in it's clan to use this command."
        End If
    End If
End Function

Public Function OnMOTD(Command As clsCommandObj) As Boolean
    If (LenB(g_Clan.Self.Name) > 0) Then
        Command.Respond StringFormatA("Clan {0}'s MOTD: {1}", g_Clan.Name, g_Clan.MOTD)
    Else
        Command.Respond "Error: I am not a member of a clan."
    End If
End Function

Public Function OnPromote(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        If (IsW3) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon < 3) Then
                            Call PromoteMember(reverseUsername(liUser.Text), liUser.SmallIcon + 1)
                        Else
                            Command.Respond "Error: The specified user is already at the highest promotable ranking."
                        End If
                    Else
                        Command.Respond "Error: The specified user is not currently a member of this clan."
                    End If
                Else
                    Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to promote."
                End If
            Else
                Command.Respond "Error: I am not a member of a clan."
            End If
        End If
    End If
End Function


Public Function OnSetMOTD(Command As clsCommandObj)
    ' This command will set the clan channel's Message Of The Day.  This
    ' command will only work if the bot is logged on using WarCraft III,
    ' and is either a Shaman or a Chieftain of the clan in question.
    
    If (Command.IsValid) Then
        If (IsW3) Then
            If (g_Clan.Self.Rank > 2) Then
                Command.Respond "The clan MOTD has been set to: " & Command.Argument("Message")
                Call frmChat.AddQ("/c motd " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
            ElseIf (g_Clan.Self.Rank > 0) Then
                Command.Respond "Error: The bot must be a Shaman or Chieftain in it's clan to set the MOTD"
            End If
        End If
    Else
        Command.Respond "You must specify a message to set."
    End If
End Function


