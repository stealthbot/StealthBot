Option Strict Off
Option Explicit On
Module modCommandsClan
	'This modules contains all the commands that have to do with Warcraft III's Clan system
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnClan(ByRef Command_Renamed As clsCommandObj)
		' This command will allow the use of Battle.net's /clan command without requiring
		' users be given the ability to use the bot's say command.
		
		If (Command_Renamed.IsValid) Then
			Select Case (LCase(Command_Renamed.Argument("SubCommand")))
				Case "public", "pub"
					If (LCase(Left(g_Channel.Name, 5)) = "clan ") Then
						If (g_Channel.Self.IsOperator) Then
							Command_Renamed.Respond("The clan channel is now public.")
							Call frmChat.AddQ("/c pub", (modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE), (Command_Renamed.Username))
						Else
							Command_Renamed.Respond("Error: The bot must have ops to change the clan privacy status.")
						End If
					Else
						Command_Renamed.Respond("Error: The bot must be in a clan channel to change the clan privacy status.")
					End If
				Case "private", "priv"
					If (LCase(Left(g_Channel.Name, 5)) = "clan ") Then
						If (g_Channel.Self.IsOperator) Then
							Command_Renamed.Respond("The clan channel is now private.")
							Call frmChat.AddQ("/c priv", (modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE), (Command_Renamed.Username))
						Else
							Command_Renamed.Respond("Error: The bot must have ops to change the clan privacy status.")
						End If
					Else
						Command_Renamed.Respond("Error: The bot must be in a clan channel to change the clan privacy status.")
					End If
				Case "motd"
					If (IsW3) Then
						If (g_Clan.Self.Rank > 2) Then
                            If (Len(Command_Renamed.Argument("Message")) = 0) Then
                                Command_Renamed.Respond("You must specify a message to set.")
                            Else
                                Command_Renamed.Respond("The clan MOTD has been set to: " & Command_Renamed.Argument("Message"))
                                Call frmChat.AddQ("/c motd " & Command_Renamed.Argument("Message"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
                            End If
						ElseIf (g_Clan.Self.Rank > 0) Then 
							Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to set the MOTD")
						End If
					End If
				Case "mail"
					If (g_Clan.Self.Rank > 2) Then
                        If (Len(Command_Renamed.Argument("Message")) = 0) Then
                            Command_Renamed.Respond("You must specify a message to send.")
                        Else
                            Command_Renamed.Respond("E-mails have been sent to everyone in the clan who have choosen to receive them.")
                            Call frmChat.AddQ("/c mail " & Command_Renamed.Argument("Message"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
                        End If
					ElseIf (g_Clan.Self.Rank > 0) Then 
						Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to send Clan mail.")
					End If
			End Select
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDemote(ByRef Command_Renamed As clsCommandObj)
		Dim liUser As System.Windows.Forms.ListViewItem
		If (Command_Renamed.IsValid) Then
			If (IsW3) Then
                If (Len(g_Clan.Self.Name) > 0) Then
                    If (g_Clan.Self.Rank > 2) Then
                        liUser = frmChat.lvClanList.FindItemWithText(Command_Renamed.Argument("Username"))

                        If (Not liUser Is Nothing) Then
                            'UPGRADE_ISSUE: MSComctlLib.ListItem property liUser.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            If (liUser.ImageIndex > 1) Then
                                'UPGRADE_ISSUE: MSComctlLib.ListItem property liUser.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                Call DemoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.ImageIndex - 1)
                            Else
                                Command_Renamed.Respond("Error: The specified user is already at the lowest demoteable ranking.")
                            End If
                        Else
                            Command_Renamed.Respond("Error: The specified user is not currently a member of this clan.")
                        End If
                    Else
                        Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to demote.")
                    End If
                Else
                    Command_Renamed.Respond("Error: The bot must be a member of a clan.")
                End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnDisbandClan(ByRef Command_Renamed As clsCommandObj)
		If (IsW3) Then
            If (Len(g_Clan.Self.Name) > 0) Then
                If (g_Clan.Self.Rank >= 4) Then
                    Call DisbandClan()
                Else
                    Command_Renamed.Respond("Error: The bot must be a chieftain to execute this command.")
                End If
            Else
                Command_Renamed.Respond("Error: The bot must be a member of a clan.")
            End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnInvite(ByRef Command_Renamed As clsCommandObj)
		' This command will send an invitation to the specified user to join the
		' clan that the bot is currently either a Shaman or Chieftain of.  This
		' command will only work if the bot is logged on using WarCraft III, and
		' is either a Shaman, or a Chieftain of the clan in question.
		
		If (IsW3) Then
			If (g_Clan.Self.Rank >= 3) Then
                If (Command_Renamed.IsValid Or Len(Command_Renamed.Argument("Username")) = 0) Then
                    Call InviteToClan(Command_Renamed.Argument("Username"))
                    Command_Renamed.Respond(Command_Renamed.Argument("Username") & ": Clan invitation sent.")
                Else
                    Command_Renamed.Respond("Error: You must specify a username to invite.")
                End If
			Else
				Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to invite users.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnMakeChieftain(ByRef Command_Renamed As clsCommandObj)
		If (IsW3) Then
			If (g_Clan.Self.Rank >= 4) Then
				If (Command_Renamed.IsValid) Then
					Call MakeMemberChieftain(ReverseConvertUsernameGateway(Command_Renamed.Argument("Username")))
				Else
					Command_Renamed.Respond("Error: You must specify a username to promote.")
				End If
			Else
				Command_Renamed.Respond("Error: The bot must be the chieftain in its clan to use this command.")
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnMOTD(ByRef Command_Renamed As clsCommandObj)
        If (Len(g_Clan.Self.Name) > 0) Then
            Command_Renamed.Respond(StringFormat("Clan {0}'s MOTD: {1}", g_Clan.Name, g_Clan.MOTD))
        Else
            Command_Renamed.Respond("Error: The bot must be a member of a clan.")
        End If
	End Sub
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnPromote(ByRef Command_Renamed As clsCommandObj)
		Dim liUser As System.Windows.Forms.ListViewItem
		If (Command_Renamed.IsValid) Then
			If (IsW3) Then
                If (Len(g_Clan.Self.Name) > 0) Then
                    If (g_Clan.Self.Rank > 2) Then
                        liUser = frmChat.lvClanList.FindItemWithText(Command_Renamed.Argument("Username"))

                        If (Not liUser Is Nothing) Then
                            'UPGRADE_ISSUE: MSComctlLib.ListItem property liUser.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            If (liUser.ImageIndex < 3) Then
                                'UPGRADE_ISSUE: MSComctlLib.ListItem property liUser.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                Call PromoteMember(ReverseConvertUsernameGateway(liUser.Text), liUser.ImageIndex + 1)
                            Else
                                Command_Renamed.Respond("Error: The specified user is already at the highest promotable ranking.")
                            End If
                        Else
                            Command_Renamed.Respond("Error: The specified user is not currently a member of this clan.")
                        End If
                    Else
                        Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to promote.")
                    End If
                Else
                    Command_Renamed.Respond("Error: The bot must be a member of a clan.")
                End If
			End If
		End If
	End Sub
	
	
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub OnSetMOTD(ByRef Command_Renamed As clsCommandObj)
		' This command will set the clan channel's Message Of The Day.  This
		' command will only work if the bot is logged on using WarCraft III,
		' and is either a Shaman or a Chieftain of the clan in question.
		
		If (Command_Renamed.IsValid) Then
			If (IsW3) Then
				If (g_Clan.Self.Rank > 2) Then
					Command_Renamed.Respond("The clan MOTD has been set to: " & Command_Renamed.Argument("Message"))
					Call frmChat.AddQ("/c motd " & Command_Renamed.Argument("Message"), (modQueueObj.PRIORITY.COMMAND_RESPONSE_MESSAGE), (Command_Renamed.Username))
				ElseIf (g_Clan.Self.Rank > 0) Then 
					Command_Renamed.Respond("Error: The bot must be a shaman or chieftain in its clan to set the MOTD")
				End If
			End If
		Else
			Command_Renamed.Respond("You must specify a message to set.")
		End If
	End Sub
End Module