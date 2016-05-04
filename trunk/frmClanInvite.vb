Option Strict Off
Option Explicit On
Friend Class frmClanInvite
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdAccept_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAccept.Click
		frmChat.AddChat(RTBColors.SuccessText, "[CLAN] Invitation accepted.")
		
		With PBuffer
			.InsertNonNTString(Clan.Token)
			.InsertNonNTString(Clan.DWName)
			.InsertNTString(Clan.Creator)
			.InsertByte(&H6)
			
			If Clan.isNew = 1 Then
				.SendPacket(SID_CLANCREATIONINVITATION)
				Clan.isNew = 0
			Else
				.SendPacket(SID_CLANINVITATIONRESPONSE)
			End If
		End With
		AwaitingClanMembership = 1
		
		Me.Close()
	End Sub
	
	Sub cmdDecline_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDecline.Click
		frmChat.AddChat(RTBColors.ErrorMessageText, "[CLAN] Invitation declined.")
		
		With PBuffer
			.InsertNonNTString(Clan.Token)
			.InsertNonNTString(Clan.DWName)
			.InsertNTString(Clan.Creator)
			.InsertByte(&H4)
			
			If Clan.isNew = 1 Then
				.SendPacket(SID_CLANCREATIONINVITATION)
				Clan.isNew = 0
			Else
				.SendPacket(SID_CLANINVITATIONRESPONSE)
			End If
		End With
		
		Me.Close()
	End Sub
	
	Private Sub frmClanInvite_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		lblUser.Text = Clan.Creator
		lblClan.Text = "Clan " & StrReverse(Clan.DWName)
		Me.Icon = frmChat.Icon
		cmdAccept.Enabled = False
		cmdDecline.Enabled = False
		
		'UPGRADE_WARNING: Add a delegate for AddressOf ClanInviteTimerProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		ClanAcceptTimerID = SetTimer(Me.Handle.ToInt32, 0, 2000, AddressOf ClanInviteTimerProc)
	End Sub
	
	Private Sub frmClanInvite_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If ClanAcceptTimerID > 0 Then
			KillTimer(0, ClanAcceptTimerID)
		End If
	End Sub
End Class