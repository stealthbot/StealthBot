Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmEMailReg
	Inherits System.Windows.Forms.Form
	'StealthBot EMail Registration Form
	'stealth@stealthbot.net
	Private ClosedProperly As Boolean
	
	' this is the main functionality of email registration
	' depending on the "Action", (result of clicking a button in the prompt OR the config values)
	' will do the specified task, then continue logon sequence
	Public Sub DoRegisterEmail(ByVal EMailAction As String, Optional ByVal EMailValue As String = vbNullString)
		Select Case EMailAction
			Case "ASKLATER"
				' "ASKLATER"/ask later: do nothing here
				frmChat.AddChat(RTBColors.SuccessText, "[EMAIL] E-mail address registration ignored. You may be prompted later.")
				
				ContinueLogonSequence()
				
			Case "NEVERASK"
				' "NEVERASK"/never ask: register an empty email address
				frmChat.AddChat(RTBColors.SuccessText, "[EMAIL] E-mail address registration declined.")
				
				modBNCS.SEND_SID_SETEMAIL(vbNullString)
				
				ContinueLogonSequence()
				
			Case Else
				' "VALUE" or "PROMPT" [default behavior]: use the provided value, or prompt with the form if empty
				' note that "VALUE" and "PROMPT" behave the same
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If LenB(EMailValue) = 0 Then
					' prompt: show email registration form
					' this is the default behavior if no config value is specified
					' (then depending on the user's selection, another one of this functions' actions will happen)
					Show()
					
					On Error Resume Next
					txtAddress.Focus()
				Else
					' value: send the provided email
					frmChat.AddChat(RTBColors.SuccessText, "[EMAIL] E-mail address registered.")
					
					SEND_SID_SETEMAIL(EMailValue)
					
					ContinueLogonSequence()
				End If
				
		End Select
	End Sub
	
	' this function continues logon sequence from where it left off
	Private Sub ContinueLogonSequence()
		If g_Connected And Not g_Online Then
			If Dii And BotVars.UseRealm Then
				Call DoQueryRealms()
			Else
				Call SendEnterChatSequence()
			End If
		End If
		
		ClosedProperly = True
		ds.WaitingForEmail = False
	End Sub
	
	Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If LenB(txtAddress.Text) > 0 Then
			Call DoRegisterEmail("VALUE", (txtAddress.Text))
			
			Me.Close()
		End If
	End Sub
	
	Private Sub cmdIgnore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdIgnore.Click
		Call DoRegisterEmail("NEVERASK")
		
		Me.Close()
	End Sub
	
	Private Sub cmdAskLater_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAskLater.Click
		Call DoRegisterEmail("ASKLATER")
		
		Me.Close()
	End Sub
	
	Private Sub frmEMailReg_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		ClosedProperly = False
		
		Label1.Text = "Battle.net would like to know if you want to register an e-mail address " & "with your account. If you want to do so, type a valid e-mail address in the box " & "below. If you don't want to register an e-mail address, click ""Never Ask Again""." & "To be asked again on your next login, click ""Ask Me Later"". ""OK"" and ""Never Ask Again"" are permanent."
		
		Label1.Text = Label1.Text & vbCrLf & vbCrLf & "Choose an option below to proceed."
		
		Label2.Text = "Click here for more information."
		
		txtAddress.Text = vbNullString
		cmdGo.Enabled = False
	End Sub
	
	Private Sub frmEMailReg_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If Not ClosedProperly And g_Online Then
			Call DoRegisterEmail("ASKLATER")
		End If
	End Sub
	
	Private Sub txtAddress_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAddress.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdGo_Click(cmdGo, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdAskLater_Click(cmdAskLater, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtAddress.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtAddress_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAddress.TextChanged
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		cmdGo.Enabled = (LenB(txtAddress.Text) > 0)
	End Sub
End Class