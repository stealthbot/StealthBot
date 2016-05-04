Option Strict Off
Option Explicit On
Friend Class frmCustomInputBox
	Inherits System.Windows.Forms.Form
	
	Private CurrentPos As Short
	Private InputMessages(9) As String
	'UPGRADE_WARNING: Lower bound of array InputValues was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private InputValues(8) As String
	Private Captions(9) As String
	'UPGRADE_WARNING: Lower bound of array MaxLengthVal was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private MaxLengthVal(8) As Short
	Private IsExpansion As Boolean
	Private NeedsExpansionKey As Boolean
	
	Private Const STEP_WELC As Short = 0
	Private Const STEP_NAME As Short = 1
	Private Const STEP_PASS As Short = 2
	Private Const STEP_PROD As Short = 3
	Private Const STEP_KEY1 As Short = 4
	Private Const STEP_KEY2 As Short = 5
	Private Const STEP_CHAN As Short = 6
	Private Const STEP_SERV As Short = 7
	Private Const STEP_OWNR As Short = 8
	Private Const STEP_DONE As Short = 9
	
	Private Sub frmCustomInputBox_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		
		Me.Icon = frmChat.Icon
		
		InputMessages(STEP_WELC) = "Welcome to StealthBot's Step-By-Step Setup. Click Next to begin. You may click Cancel at any point and no changes will be made."
		InputMessages(STEP_NAME) = "Please enter the username you'd like your bot to use. If it's not already existent, the bot will create it."
		InputMessages(STEP_PASS) = "Enter the password corresponding with the account you just entered."
		InputMessages(STEP_PROD) = "Which game would you like the bot to connect with?"
		InputMessages(STEP_KEY1) = "Please enter a valid CDKey for the game %s."
		InputMessages(STEP_KEY2) = "Please enter a valid CDKey for the %s expansion. (Both a Regular and an Expansion key are required.)"
		InputMessages(STEP_CHAN) = "What channel would you like the bot to use as its home? An empty channel will take you to the server's default when you log on."
		InputMessages(STEP_SERV) = "To which Battle.net gateway will you be connecting? (USEast, USWest, Asia, Europe)"
		InputMessages(STEP_OWNR) = "Enter your main Battle.net account name. This will act as the bot's ""owner"" account -- you can leave it blank." & vbCrLf & vbCrLf & "IMPORTANT: Correct the Owner Name later if the bot sees your account name differently once it has connected -- it must be EXACT and include any @Realm stuff that the bot sees on your name."
		InputMessages(STEP_DONE) = "Congratulations, you're all set up! Enjoy your StealthBot, and remember to visit http://www.stealthbot.net if you have problems."
		
		Captions(STEP_WELC) = "Welcome to StealthBot!"
		Captions(STEP_NAME) = "Set Username"
		Captions(STEP_PASS) = "Set Password"
		Captions(STEP_PROD) = "Game Selection"
		Captions(STEP_KEY1) = "Set CD-Key"
		Captions(STEP_KEY2) = "Set Expansion CD-Key"
		Captions(STEP_CHAN) = "Set Home Channel"
		Captions(STEP_SERV) = "Gateway Selection"
		Captions(STEP_OWNR) = "Set Bot Owner"
		Captions(STEP_DONE) = "Finished!"
		
		MaxLengthVal(STEP_NAME) = 15
		MaxLengthVal(STEP_KEY1) = 60
		MaxLengthVal(STEP_KEY2) = 60
		MaxLengthVal(STEP_CHAN) = 31
		MaxLengthVal(STEP_OWNR) = 30
		
		For i = 1 To 8
			InputValues(i) = vbNullString
		Next i
		
		NeedsExpansionKey = False
		IsExpansion = False
		
		With cboGame
			.Items.Clear()
			.Items.Add("WarCraft II: Battle.net Edition")
			.Items.Add("StarCraft")
			.Items.Add("StarCraft: Brood War")
			.Items.Add("Diablo II")
			.Items.Add("Diablo II: Lord of Destruction")
			.Items.Add("WarCraft III")
			.Items.Add("WarCraft III: The Frozen Throne")
			.Text = "- choose one -"
			.Visible = False
		End With
		
		With cboServer
			.Items.Clear()
			.Items.Add("USEast / Azeroth")
			.Items.Add("USWest / Lordaeron")
			.Items.Add("Asia / Kalimdor")
			.Items.Add("Europe / Northrend")
			.Text = "- choose one -"
			.Visible = False
		End With
		
		CurrentPos = 0
		ShowCurrentPos()
	End Sub
	
	Private Sub cmdBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBack.Click
		If CurrentPos > STEP_WELC And CurrentPos < STEP_DONE Then InputValues(CurrentPos) = txtInput.Text
		CurrentPos = CurrentPos - 1
		Call ShowCurrentPos(True)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click
		Dim SettingsForm As Object
		If CurrentPos < STEP_DONE Then
			If CurrentPos > STEP_WELC And CurrentPos < STEP_DONE Then InputValues(CurrentPos) = txtInput.Text
			CurrentPos = CurrentPos + 1
			ShowCurrentPos()
		Else
			If frmChat.SettingsForm Is Nothing Then
				frmChat.SettingsForm = New frmSettings
				frmChat.SettingsForm.Show()
			End If
			
			With frmChat.SettingsForm
				.ShowPanel(modEnum.enuSettingsPanels.spConnectionConfig)
				.txtUsername.Text = InputValues(STEP_NAME)
				.txtPassword.Text = InputValues(STEP_PASS)
				
				'            .AddItem "(Choose One)"
				'            .AddItem "Warcraft II Battle.Net Edition"
				'            .AddItem "Starcraft"
				'            .AddItem "Starcraft: Brood War"
				'            .AddItem "Diablo II"
				'            .AddItem "Diablo II: Lord of Destruction"
				'            .AddItem "Warcraft III"
				'            .AddItem "Warcraft III: The Frozen Throne"
				
				Select Case cboGame.SelectedIndex
					Case 0
						.optW2BN.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 1
						.optSTAR.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 2
						.optSEXP.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 3
						.optD2DV.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 4
						.optD2XP.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 5
						.optWAR3.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
					Case 6
						.optW3XP.Checked = True
						'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
				End Select
				
				.txtCdKey.Text = InputValues(STEP_KEY1)
				'UPGRADE_WARNING: Untranslated statement in cmdNext_Click. Please check source code.
				.txtExpKey.Text = InputValues(STEP_KEY2)
				.txtHomeChan.Text = InputValues(STEP_CHAN)
				
				'            .AddItem "(Choose One)"
				'            .AddItem "USEast / Azeroth"
				'            .AddItem "USWest / Lordaeron"
				'            .AddItem "Asia / Kalimdor"
				'            .AddItem "Europe / Northrend"
				
				Select Case cboServer.SelectedIndex
					Case 0 : .cboServer.Text = "useast.battle.net"
					Case 1 : .cboServer.Text = "uswest.battle.net"
					Case 2 : .cboServer.Text = "asia.battle.net"
					Case 3 : .cboServer.Text = "europe.battle.net"
				End Select
				
				.txtOwner.Text = InputValues(STEP_OWNR)
				
				Me.Close()
			End With
		End If
	End Sub
	
	Private Sub ShowCurrentPos(Optional ByVal GoingBackwards As Boolean = False)
		Dim s As String
		Dim InputPresent As Boolean
		
		'debug.print "CurrentPos: " & CurrentPos
		'debug.print "GoingBackwards: " & GoingBackwards
		
		If CurrentPos = STEP_KEY2 Then
			If Not NeedsExpansionKey Then
				If GoingBackwards Then
					CurrentPos = STEP_KEY1
				Else
					CurrentPos = STEP_CHAN
				End If
			End If
			
			lblText.Text = FormatOutput(InputMessages(CurrentPos))
		Else
			lblText.Text = FormatOutput(InputMessages(CurrentPos))
		End If
		
		Me.Text = Captions(CurrentPos)
		
		' textbox visible: all steps except STEP_PROD, STEP_SERV
		txtInput.Visible = (CurrentPos > STEP_WELC And CurrentPos < STEP_DONE And CurrentPos <> STEP_PROD And CurrentPos <> STEP_SERV)
		' server visible IF STEP_SERV
		cboServer.Visible = (CurrentPos = STEP_SERV)
		' game visible IF STEP_PROD
		cboGame.Visible = (CurrentPos = STEP_PROD)
		' back enabled IF > STEP_WELC
		cmdBack.Enabled = (CurrentPos > STEP_WELC)
		' password char IF STEP_PASS
		txtInput.PasswordChar = IIf(CurrentPos = STEP_PASS, "*", vbNullString)
		
		' get saved input for this value
		InputPresent = True
		If CurrentPos > STEP_WELC And CurrentPos < STEP_DONE Then
			' get saved value
			txtInput.Text = InputValues(CurrentPos)
			' get max length value
			'UPGRADE_WARNING: TextBox property txtInput.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			txtInput.Maxlength = MaxLengthVal(CurrentPos)
			' find "input present" state
			InputPresent = False
			If CurrentPos = STEP_PROD And StrComp(cboGame.Text, "- choose one -", CompareMethod.Binary) <> 0 Then
				InputPresent = True
			ElseIf CurrentPos = STEP_SERV And StrComp(cboServer.Text, "- choose one -", CompareMethod.Binary) <> 0 Then 
				InputPresent = True
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			ElseIf LenB(txtInput.Text) > 0 Then 
				InputPresent = True
			ElseIf CurrentPos = STEP_CHAN Or CurrentPos = STEP_OWNR Then 
				InputPresent = True
			End If
		End If
		' use "input present" state
		cmdNext.Enabled = InputPresent
		
		' if STEP_DONE, set special caption for next button
		If CurrentPos = STEP_DONE Then
			cmdNext.Text = "&Finish!"
		Else
			cmdNext.Text = ">> &Next"
		End If
		
		' set focus
		On Error Resume Next
		If txtInput.Visible Then
			txtInput.Focus()
			txtInput.SelectionStart = 0
			txtInput.SelectionLength = Len(txtInput.Text)
		End If
		If cboGame.Visible Then cboGame.Focus()
		If cboServer.Visible Then cboServer.Focus()
	End Sub
	
	Function FormatOutput(ByVal sIn As String) As String
		FormatOutput = sIn
		If CurrentPos = STEP_KEY1 Then
			If IsExpansion Then
				' (STEP_KEY1) if IsExpansion, then "%s" is non-expansion name item
				FormatOutput = Replace(sIn, "%s", VB6.GetItemString(cboGame, cboGame.SelectedIndex - 1))
			Else
				' (STEP_KEY1) else "%s" is currently selected name item
				FormatOutput = Replace(sIn, "%s", VB6.GetItemString(cboGame, cboGame.SelectedIndex))
			End If
		ElseIf CurrentPos = STEP_KEY2 Then 
			If NeedsExpansionKey Then
				' (STEP_KEY2) if needs expansion key, then "%s" is currently selected name item
				FormatOutput = Replace(sIn, "%s", VB6.GetItemString(cboGame, cboGame.SelectedIndex))
			End If
		End If
	End Function
	
	Sub cboGame_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboGame.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = 0
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboGame.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub cboGame_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGame.SelectedIndexChanged
		' needs expansion key: D2XP, W3XP
		NeedsExpansionKey = (cboGame.SelectedIndex = 4 Or cboGame.SelectedIndex = 6)
		' is expansion (name of KEY1 is non-expansion game's name): D2XP, W3XP, SEXP
		IsExpansion = (cboGame.SelectedIndex = 2 Or cboGame.SelectedIndex = 4 Or cboGame.SelectedIndex = 6)
		' enabled: product is not "- choose one -"
		cmdNext.Enabled = (StrComp(cboGame.Text, "- choose one -", CompareMethod.Binary) <> 0)
		'debug.print "NeedsExpansionKey is now " & CBool(cboGame.ListIndex = 4 Or cboGame.ListIndex = 6)
	End Sub
	
	Sub cboServer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboServer.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = 0
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboServer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub cboServer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboServer.SelectedIndexChanged
		' enabled: server is not "- choose one -"
		cmdNext.Enabled = (StrComp(cboServer.Text, "- choose one -", CompareMethod.Binary) <> 0)
	End Sub
	
	Private Sub txtInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdNext_Click(cmdNext, New System.EventArgs())
			KeyAscii = 0
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdCancel_Click(cmdCancel, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtInput.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub txtInput_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInput.TextChanged
		' enabled: text field has input
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		cmdNext.Enabled = (LenB(txtInput.Text) > 0)
		' always enabled for STEP_CHAN, STEP_OWNR
		If CurrentPos = STEP_CHAN Or CurrentPos = STEP_OWNR Then cmdNext.Enabled = True
	End Sub
End Class