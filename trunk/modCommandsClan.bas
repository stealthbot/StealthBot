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
                If (g_Clan.InClan) Then
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
                Else
                    Command.Respond "Error: The bot must be a member of a clan."
                End If
            Case "mail"
                If (g_Clan.InClan) Then
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
                Else
                    Command.Respond "Error: The bot must be a member of a clan."
                End If
        End Select
    End If
End Sub

Public Sub OnDemote(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (g_Clan.InClan) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon > 1) Then
                            Call frmChat.ClanHandler.DemoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.SmallIcon - 1, reqUserCommand, Command)
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
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    End If
End Sub

Public Sub OnDisbandClan(Command As clsCommandObj)
    If (g_Clan.InClan) Then
        If (LenB(g_Clan.Self.Name) > 0) Then
            If (g_Clan.Self.Rank >= clrankChieftain) Then
                Call frmChat.ClanHandler.DisbandClan(reqUserCommand, Command)
            Else
                Command.Respond "Error: The bot must be the chieftain in its clan to use this command."
            End If
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnInvite(Command As clsCommandObj)
    ' This command will send an invitation to the specified user to join the
    ' clan that the bot is currently either a Shaman or Chieftain of.  This
    ' command will only work if the bot is logged on using WarCraft III, and
    ' is either a Shaman, or a Chieftain of the clan in question.
    If (g_Clan.InClan) Then
        If (g_Clan.Self.Rank >= clrankShaman) Then
            If (Command.IsValid Or LenB(Command.Argument("Username")) = 0) Then
                Call frmChat.ClanHandler.InviteToClan(Command.Argument("Username"), reqUserCommand, Command)
                Command.Respond Command.Argument("Username") & ": Clan invitation sent."
            Else
                Command.Respond "Error: You must specify a username to invite."
            End If
        Else
            Command.Respond "Error: The bot must be a shaman or chieftain in its clan to invite users."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnLeaveClan(Command As clsCommandObj)
    If (g_Clan.InClan) Then
        If (g_Clan.Self.Rank < clrankChieftain) Then
            Call frmChat.ClanHandler.RemoveMember(g_Clan.Self.Name, True, reqUserCommand, Command)
        Else
            Command.Respond "Error: The bot cannot be the chieftain in its clan and leave the clan."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnMakeChieftain(Command As clsCommandObj)
    If (g_Clan.InClan) Then
        If (g_Clan.Self.Rank >= clrankChieftain) Then
            If (Command.IsValid) Then
                Call frmChat.ClanHandler.MakeMemberChieftain(ReverseConvertUsernameGateway(Command.Argument("Username")), reqUserCommand, Command)
            Else
                Command.Respond "Error: You must specify a username to promote."
            End If
        Else
            Command.Respond "Error: The bot must be the chieftain in its clan to use this command."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnMOTD(Command As clsCommandObj)
    If (g_Clan.InClan) Then
        If (LenB(g_Clan.Self.Name) > 0) Then
            Command.Respond StringFormat("Clan {0}'s MOTD: {1}", g_Clan.Name, g_Clan.MOTD)
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnPromote(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (g_Clan.InClan) Then
            If (LenB(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank > 2) Then
                    Dim liUser As ListItem
                    Set liUser = frmChat.lvClanList.FindItem(Command.Argument("Username"))
    
                    If (Not liUser Is Nothing) Then
                        If (liUser.SmallIcon < 3) Then
                            Call frmChat.ClanHandler.PromoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.SmallIcon + 1, reqUserCommand, Command)
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
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    End If
End Sub

Public Sub OnRemoveMember(Command As clsCommandObj)
    If (g_Clan.InClan) Then
        If (g_Clan.Self.Rank >= clrankShaman) Then
            If (Command.IsValid) Then
                Call frmChat.ClanHandler.RemoveMember(ReverseConvertUsernameGateway(Command.Argument("Username")), False, reqUserCommand, Command)
            Else
                Command.Respond "Error: You must specify a username to remove."
            End If
        Else
            Command.Respond "Error: The bot must be a shaman or chieftain in its clan to remove users."
        End If
    Else
        Command.Respond "Error: The bot must be a member of a clan."
    End If
End Sub

Public Sub OnSetMOTD(Command As clsCommandObj)
    ' This command will set the clan channel's Message Of The Day.  This
    ' command will only work if the bot is logged on using WarCraft III,
    ' and is either a Shaman or a Chieftain of the clan in question.
    
    If (Command.IsValid) Then
        If (g_Clan.InClan) Then
            If (g_Clan.Self.Rank > 2) Then
                Command.Respond "The clan MOTD has been set to: " & Command.Argument("Message")
                Call frmChat.AddQ("/c motd " & Command.Argument("Message"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
                Call frmChat.ClanHandler.RequestClanMOTD(reqInternal)
            ElseIf (g_Clan.Self.Rank > 0) Then
                Command.Respond "Error: The bot must be a shaman or chieftain in its clan to set the MOTD."
            End If
        Else
            Command.Respond "Error: The bot must be a member of a clan."
        End If
    Else
        Command.Respond "Error: You must specify a message to set."
    End If
End Sub


