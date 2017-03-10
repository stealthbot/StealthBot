Attribute VB_Name = "modCommandsClan"
Option Explicit
'This modules contains all the commands that have to do with Warcraft III's Clan system

Public Sub OnClan(Command As clsCommandObj)
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
                        Command.Respond "Error: The bot must be a shaman or chieftain in its clan to set the MOTD"
                    End If
                End If
            Case "mail"
                If (g_Clan.Self.Rank > 2) Then
                    If (LenB(Command.Argument("Message")) = 0) Then
                        Command.Respond "You must specify a message to send."
                    Else
                        Command.Respond "Emails have been sent to everyone in the clan who have choosen to receive them."
                        Call frmChat.AddQ("/c mail " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
                    End If
                ElseIf (g_Clan.Self.Rank > 0) Then
                    Command.Respond "Error: The bot must be a shaman or chieftain in its clan to send Clan mail."
                End If
        End Select
    End If
End Sub

Public Sub OnDemote(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (IsW3) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon > 1) Then
                            Call DemoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.SmallIcon - 1)
                        Else
                            Command.Respond "Error: The specified user is already at the lowest demoteable ranking."
                        End If
                    Else
                        Command.Respond "Error: The specified user is not currently a member of this clan."
                    End If
                Else
                    Command.Respond "Error: The bot must be a shaman or chieftain in its clan to demote."
                End If
            Else
                Command.Respond "Error: The bot must be a member of a clan."
            End If
        End If
    End If
End Sub

Public Sub OnDisbandClan(Command As clsCommandObj)
    If (IsW3) Then
        If (LenB(g_Clan.Self.Name) > 0) Then
            If (g_Clan.Self.Rank >= 4) Then
                Call DisbandClan
            Else
                Command.Respond "Error: The bot must be a chieftain to execute this command."
            End If
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    End If
End Sub

Public Sub OnInvite(Command As clsCommandObj)
    ' This command will send an invitation to the specified user to join the
    ' clan that the bot is currently either a Shaman or Chieftain of.  This
    ' command will only work if the bot is logged on using WarCraft III, and
    ' is either a Shaman, or a Chieftain of the clan in question.
    
    If (IsW3) Then
        If (g_Clan.Self.Rank >= 3) Then
            If (Command.IsValid Or LenB(Command.Argument("Username")) = 0) Then
                Call InviteToClan(Command.Argument("Username"))
                Command.Respond Command.Argument("Username") & ": Clan invitation sent."
            Else
                Command.Respond "Error: You must specify a username to invite."
            End If
        Else
            Command.Respond "Error: The bot must be a shaman or chieftain in its clan to invite users."
        End If
    End If
End Sub

Public Sub OnMakeChieftain(Command As clsCommandObj)
    If (IsW3) Then
        If (g_Clan.Self.Rank >= 4) Then
            If (Command.IsValid) Then
                Call MakeMemberChieftain(ReverseConvertUsernameGateway(Command.Argument("Username")))
            Else
                Command.Respond "Error: You must specify a username to promote."
            End If
        Else
            Command.Respond "Error: The bot must be the chieftain in its clan to use this command."
        End If
    End If
End Sub

Public Sub OnMOTD(Command As clsCommandObj)
    If (LenB(g_Clan.Self.Name) > 0) Then
        Command.Respond StringFormat("Clan {0}'s MOTD: {1}", g_Clan.Name, g_Clan.MOTD)
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnPromote(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (IsW3) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon < 3) Then
                            Call PromoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.SmallIcon + 1)
                        Else
                            Command.Respond "Error: The specified user is already at the highest promotable ranking."
                        End If
                    Else
                        Command.Respond "Error: The specified user is not currently a member of this clan."
                    End If
                Else
                    Command.Respond "Error: The bot must be a shaman or chieftain in its clan to promote."
                End If
            Else
                Command.Respond "Error: The bot must be a member of a clan."
            End If
        End If
    End If
End Sub


Public Sub OnSetMOTD(Command As clsCommandObj)
    ' This command will set the clan channel's Message Of The Day.  This
    ' command will only work if the bot is logged on using WarCraft III,
    ' and is either a Shaman or a Chieftain of the clan in question.
    
    If (Command.IsValid) Then
        If (IsW3) Then
            If (g_Clan.Self.Rank > 2) Then
                Command.Respond "The clan MOTD has been set to: " & Command.Argument("Message")
                Call frmChat.AddQ("/c motd " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
            ElseIf (g_Clan.Self.Rank > 0) Then
                Command.Respond "Error: The bot must be a shaman or chieftain in its clan to set the MOTD"
            End If
        End If
    Else
        Command.Respond "You must specify a message to set."
    End If
End Sub


