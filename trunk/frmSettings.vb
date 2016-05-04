Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmSettings
	Inherits System.Windows.Forms.Form
	'// frmSettings.frm - project StealthBot - january 2004 - authored by Stealth (stealth@stealthbot.net)
	'// To switch between panels on the UI display, click the current top panel
	'//     and use CTRL+K to send it to the back. Rinse and repeat until you are looking
	'//     at the panel you want to edit.
	' u 7/13/09 to fix topic 42703 -andy
	
	Private mColors() As Integer
	Private FirstRun As Byte
	Private ModifiedColors As Boolean
	Private InitChanFont As String
	Private InitChatFont As String
	Private InitChanSize As Short
	Private InitChatSize As Short
	Private PanelsInitialized As Boolean
	Private OldBotOwner As String
	
	Const SC As Byte = 0
	Const BW As Byte = 1
	Const D2 As Byte = 2
	Const D2X As Byte = 3
	Const W3 As Byte = 4
	Const W3X As Byte = 5
	Const W2 As Byte = 6
	
	Private Const BNLS_NOT_SET As String = "No server set"
	
	Private Sub frmSettings_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		Dim nRoot As vbalTreeViewLib6.cTreeViewNode
		Dim nCurrent As vbalTreeViewLib6.cTreeViewNode
		Dim nTopLevel As vbalTreeViewLib6.cTreeViewNodes
		Dim nOptLevel As vbalTreeViewLib6.cTreeViewNodes
		Dim lMouseOver As Integer
		Dim s As String
		Dim serverList() As String
		Dim j, i, K As Integer
		Dim colBNLS As Collection
		
		'##########################################
		' TREEVIEW INITIALIZATION CODE
		'##########################################
		
		lMouseOver = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		
		With tvw
			
			.nodes.Clear()
			
			nRoot = .nodes.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwFirst, "root", "StealthBot Settings")
			nRoot.MouseOverForeColor = System.Convert.ToUInt32(lMouseOver)
			
			nTopLevel = nRoot.Children
			
			nCurrent = nTopLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "connection", "Connection Settings")
			nCurrent.MouseOverForeColor = System.Convert.ToUInt32(lMouseOver)
			
			nOptLevel = nCurrent.Children
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "conn_config", "Basic Setup") 'general setup
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "conn_advanced", "Advanced") 'proxies/spoofing/bnls
			
			nCurrent.Expanded = True
			
			nCurrent = nTopLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "interface", "Interface Settings")
			nCurrent.MouseOverForeColor = System.Convert.ToUInt32(lMouseOver)
			
			nOptLevel = nCurrent.Children
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "int_general", "General Settings")
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "int_fonts", "Fonts and Colors")
			
			nCurrent.Expanded = True
			
			nCurrent = nTopLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "general", "General Settings")
			nCurrent.MouseOverForeColor = System.Convert.ToUInt32(lMouseOver)
			
			nOptLevel = nCurrent.Children
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "op_moderation", "Moderation")
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "op_logging", "Logging")
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "op_greets", "Greet Message")
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "op_idles", "Idle Message")
			nOptLevel.Add( , vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "op_misc", "Miscellaneous")
			
			nCurrent.Expanded = True
			
			'UPGRADE_NOTE: Object nOptLevel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			nOptLevel = Nothing
			'UPGRADE_NOTE: Object nCurrent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			nCurrent = Nothing
			
			'UPGRADE_NOTE: Object nTopLevel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			nTopLevel = Nothing
			
			nRoot.Expanded = True
			
			'UPGRADE_NOTE: Object nRoot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			nRoot = Nothing
			
		End With
		
		'##########################################
		' PROFILE SELECTOR COMBO BOX STUFF
		'##########################################
		
		With cboProfile
			.Text = "[default profile]"
		End With
		
		colProfiles = New Collection
		
		'Call LoadProfileList(cboProfile)
		
		
		'##########################################
		' INTERFACE DISPLAY
		'##########################################
		
		lblSplash.Text = vbCrLf & "If you're new to bots, click " & Chr(39) & "Step-By-Step Configuration" & Chr(39) & " below for a walkthrough to get the bot set up." & vbCrLf & vbCrLf & "Otherwise, click a section on the left to change settings."
		
		With cboSpoof
			.Items.Add("Disabled")
			.Items.Add("0ms (send first)")
			.Items.Add("-1ms (ignore)")
			.SelectedIndex = 0
		End With
		
		With cboConnMethod
			.Items.Add("BNLS - Battle.net Logon Server")
			.Items.Add("ADVANCED - Local hashing")
			.SelectedIndex = 0
		End With
		
		Dim BNLSServers As Collection
		With cboBNLSServer
			
			.Items.Add("Automatic (Server Finder)")
			
			' If the user has a server set, add it.
			If Len(Config.BNLSServer) > 0 Then
				AddBNLSServer(Config.BNLSServer)
			End If
			
			' Figure out the default selection
			If Config.BNLSFinder Or Len(Config.BNLSServer) = 0 Then
				.SelectedIndex = 0
			ElseIf Len(Config.BNLSServer) > 0 Then 
				.SelectedIndex = GetBnlsIndex(Config.BNLSServer)
			Else
				.SelectedIndex = 0
			End If
			
			' Add servers from the user's local list.
			colBNLS = ListFileLoad(GetFilePath(FILE_BNLS_LIST))
			
			If colBNLS.Count() > 0 Then
				For i = 1 To colBNLS.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object colBNLS.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddBNLSServer(colBNLS.Item(i))
				Next i
			End If
			
			'UPGRADE_NOTE: Object BNLSServers may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			BNLSServers = Nothing
		End With
		
		With cboLogging
			.Items.Add("Full - text is logged and a dated logfile is saved during operation.")
			.Items.Add("Temporary - text is logged. The logfile is deleted on shutdown.")
			.Items.Add("Disabled")
			.SelectedIndex = 0
		End With
		
		With cboTimestamp
			.Items.Add("[HH:MM:SS PM] - Seconds with time of day")
			.Items.Add("[HH:MM:SS] - Seconds")
			.Items.Add("[HH:MM:SS:mmm] - Milliseconds")
			.Items.Add("No timestamp")
			.SelectedIndex = 0
		End With
		
		lblGreetVars.Text = "Greet Message Variables: (Suggest more! email stealth@stealthbot.net) " & vbNewLine & "%c = Current channel" & vbNewLine & "%0 = Username of the person who joins" & vbNewLine & "%1 = Bot's current username" & vbNewLine & "%p = Ping of user who joins" & vbNewLine & "%v = The bot's current version" & vbNewLine & "%a = Database access of the person who joins" & vbNewLine & "%f = Database flags of the person who joins" & vbNewLine & "%t = Current time" & vbNewLine & "%d = Current date"
		
		lblIdleVars.Text = "Idle message variables: (Suggest more! email stealth@stealthbot.net) " & vbNewLine & "%c = Current channel" & vbNewLine & "%me = Current username" & vbNewLine & "%v = Bot version" & vbNewLine & "%botup = Bot uptime" & vbNewLine & "%cpuup = System uptime" & vbNewLine & "%mp3 = Current MP3" & vbNewLine & "%quote = Random quote" & vbNewLine & "%rnd = Random person in the channel" & vbNewLine
		
		'##########################################
		'COLOR STUFF
		'##########################################
		Call LoadColors()
		
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		cDLGOpen.Filter = "StealthBot ColorList Files|*.sclf"
		cDLGSave.Filter = "StealthBot ColorList Files|*.sclf"
		cboColorList.SelectedIndex = 0
		'##########################################
		'##########################################
		ShowCurrentColor()
		
		Call InitAllPanels()
		
		PanelsInitialized = True
		
		'##########################################
		'   LAST SETTINGS PANEL
		'##########################################
		
		If Config.LastSettingsPanel = -1 Then
			ShowPanel(modEnum.enuSettingsPanels.spSplash, 1)
		Else
			ShowPanel(CShort(Config.LastSettingsPanel))
		End If
	End Sub
	
	Private Function GetBnlsIndex(ByVal sServerName As String) As Short
		Dim i As Short
		GetBnlsIndex = -1
		
		For i = 1 To (cboBNLSServer.Items.Count - 1)
			If StrComp(sServerName, VB6.GetItemString(cboBNLSServer, i), CompareMethod.Text) = 0 Then
				GetBnlsIndex = i
				Exit Function
			End If
		Next 
	End Function
	
	' Adds the server to the list (if it does not already exist) and returns its position.
	Private Function AddBNLSServer(ByVal sServerHost As String) As Short
		Dim i As Short ' counter
		sServerHost = Trim(sServerHost)
		
		If Len(sServerHost) = 0 Then
			AddBNLSServer = -1
			Exit Function
		End If
		
		' Check if the server is already in the list.
		i = GetBnlsIndex(sServerHost)
		If i = -1 Then
			cboBNLSServer.Items.Add(sServerHost)
			i = GetBnlsIndex(sServerHost)
		End If
		
		AddBNLSServer = i
	End Function
	
	Private Sub SetHashPath(ByVal Path As String)
		lblHashPath.Text = Path
		ToolTip1.SetToolTip(lblHashPath, Path)
		
		' Try to adjust font size for long paths
		If ((Len(Path) > 80) And (Not InStr(Path, Space(1)) > 0)) Then
			If Len(Path) > 100 Then
				lblHashPath.Font = VB6.FontChangeSize(lblHashPath.Font, 6)
			Else
				lblHashPath.Font = VB6.FontChangeSize(lblHashPath.Font, 7)
			End If
		End If
		
	End Sub
	
	Function KeyToIndex(ByVal sKey As String) As Byte
		
		Select Case sKey
			Case "splash" : KeyToIndex = 8
				
			Case "conn_config" : KeyToIndex = 0
			Case "conn_advanced" : KeyToIndex = 1
				
			Case "int_general" : KeyToIndex = 2
			Case "int_fonts" : KeyToIndex = 3
				
			Case "op_moderation" : KeyToIndex = 4
			Case "op_logging" : KeyToIndex = 9
			Case "op_greets" : KeyToIndex = 5
			Case "op_idles" : KeyToIndex = 6
			Case "op_misc" : KeyToIndex = 7
				
			Case Else : KeyToIndex = 8
		End Select
		
	End Function
	
	Sub ShowPanel(ByVal Index As modEnum.enuSettingsPanels, Optional ByVal Mode As Byte = 0)
		
		Static ActivePanel As Short
		
		Dim nod As vbalTreeViewLib6.cTreeViewNode
		Dim i As Short
		If PanelsInitialized Then
			If Index <> 8 Then
				For i = 1 To tvw.NodeCount
					nod = tvw.nodes.Item(i)
					If Not nod.Selected And KeyToIndex(nod.Key) = Index Then
						nod.Selected = True
						Exit For
					End If
				Next i
				'UPGRADE_NOTE: Object nod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				nod = Nothing
			End If
			If Mode = 1 Then
				fraPanel(KeyToIndex("splash")).BringToFront()
				ActivePanel = KeyToIndex("splash")
			Else
				'fraPanel(ActivePanel).ZOrder vbSendToBack
				fraPanel(Index).BringToFront()
				ActivePanel = Index
				Config.LastSettingsPanel = ActivePanel
				
				'Debug.Print "Writing: " & ActivePanel
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboConnMethod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboConnMethod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboConnMethod.SelectedIndexChanged
		'Disables the server selection box when hashing is selected
		cboBNLSServer.Enabled = CBool(cboConnMethod.SelectedIndex = 0)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdReadme_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReadme.Click
		OpenReadme()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		If ModifiedColors Then Call cmdSaveColor_Click(cmdSaveColor, New System.EventArgs())
		
		If Not InvalidConfigValues Then
			If SaveSettings Then
				Me.Close()
			End If
		End If
	End Sub
	
	Private Sub cmdSaveColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSaveColor.Click
		ModifiedColors = True
		RecordCurrentColor()
		lblColorStatus.Text = "Color settings saved."
	End Sub
	
	Private Sub cmdStepByStep_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStepByStep.Click
		frmCustomInputBox.Show()
	End Sub
	
	Private Sub cmdWebsite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdWebsite.Click
		Call frmChat.mnuHelpWebsite_Click(Nothing, New System.EventArgs())
	End Sub
	
	Private Sub frmSettings_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call frmChat.DeconstructSettings()
		'UPGRADE_NOTE: Object colProfiles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colProfiles = Nothing
	End Sub
	
	Sub lblAddCurrentKey_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblAddCurrentKey.Click
		Dim keys As Collection
		Dim Item As Object
		Dim Key1 As String
		Dim Key2 As String
		
		' Load the list
		keys = ListFileLoad(GetFilePath("Keys.txt"))
		
		Key1 = UCase(CDKeyReplacements(txtCDKey.Text))
		Key2 = UCase(CDKeyReplacements(txtExpKey.Text))
		
		' if it's already there, do nothing.
		For	Each Item In keys
			'UPGRADE_WARNING: Couldn't resolve default property of object Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If StrComp(CStr(Item), Key1, CompareMethod.Text) = 0 Then Key1 = vbNullString
			'UPGRADE_WARNING: Couldn't resolve default property of object Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If StrComp(CStr(Item), Key2, CompareMethod.Text) = 0 Then Key2 = vbNullString
		Next Item
		
		' Add the keys
        If Len(Key1) > 0 Then keys.Add(Key1)
        If Len(Key2) > 0 Then keys.Add(Key2)
		
		' Save the list
        If Len(Key1) > 0 Or Len(Key2) > 0 Then
            ListFileSave(GetFilePath("Keys.txt"), keys)
        End If
		
		'UPGRADE_NOTE: Object keys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		keys = Nothing
	End Sub
	
	Private Sub lblManageKeys_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblManageKeys.Click
        If Len(txtCDKey.Text) > 0 Then
            Call lblAddCurrentKey_Click(lblAddCurrentKey, New System.EventArgs())
        End If
		
		frmManageKeys.Show()
	End Sub
	
	Sub lblAddCurrentKey_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblAddCurrentKey.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblAddCurrentKey.ForeColor = System.Drawing.Color.Blue
		lblManageKeys.ForeColor = System.Drawing.Color.White
	End Sub
	
	Sub lblManageKeys_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblManageKeys.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblAddCurrentKey.ForeColor = System.Drawing.Color.White
		lblManageKeys.ForeColor = System.Drawing.Color.Blue
	End Sub
	
	Sub frmSettings_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblAddCurrentKey.ForeColor = System.Drawing.Color.White
		lblManageKeys.ForeColor = System.Drawing.Color.White
	End Sub
	
	'UPGRADE_ISSUE: Frame event fraPanel.MouseMove was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Sub fraPanel_MouseMove(ByRef Index As Short, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		lblAddCurrentKey.ForeColor = System.Drawing.Color.White
		lblManageKeys.ForeColor = System.Drawing.Color.White
	End Sub
	
	'Private Sub tvw_Click()
	'    Call tvw_SelectedNodeChanged
	'End Sub
	'
	'Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
	'    Call tvw_SelectedNodeChanged
	'End Sub
	'
	'Private Sub tvw_NodeClick(node As vbalTreeViewLib6.cTreeViewNode)
	'    Call tvw_SelectedNodeChanged
	'End Sub
	
	Private Sub tvw_SelectedNodeChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tvw.SelectedNodeChanged
		If FirstRun = 0 Then
			ShowPanel(KeyToIndex(tvw.SelectedItem.Key))
		Else
			FirstRun = 0
		End If
	End Sub
	
	Private Function DoCDKeyLengthCheck(ByVal sKey As String, ByVal sProd As String) As Boolean
		sKey = CDKeyReplacements(sKey)
		
		DoCDKeyLengthCheck = True
		If Config.IgnoreCDKeyLength Then Exit Function
		
		Select Case sProd
			Case PRODUCT_STAR, PRODUCT_SEXP
				If ((Len(sKey) <> 13) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
				
			Case PRODUCT_D2DV, PRODUCT_D2XP
				If ((Len(sKey) <> 16) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
				
			Case PRODUCT_W2BN
				If (Len(sKey) <> 16) Then DoCDKeyLengthCheck = False
				
			Case PRODUCT_WAR3, PRODUCT_W3XP
				If (Len(sKey) <> 26) Then DoCDKeyLengthCheck = False
				
			Case PRODUCT_SSHR, PRODUCT_DSHR, PRODUCT_DRTL
				DoCDKeyLengthCheck = True
				
			Case PRODUCT_JSTR
				If (Len(sKey) <> 13) Then DoCDKeyLengthCheck = False
				
			Case Else
				DoCDKeyLengthCheck = False
		End Select
	End Function
	
	Private Function SaveSettings() As Boolean
		Dim s As String
		Dim Clients(6) As String
		Dim i, j As Integer
		Dim colBNLS As New Collection
		
		' First, CDKey Length check and corresponding stuff that needs to run first:
		Select Case True
			Case optSTAR.Checked
				If (chkSHR.CheckState) Then
					s = PRODUCT_SSHR
				ElseIf (chkJPN.CheckState) Then 
					s = PRODUCT_JSTR
				Else
					s = PRODUCT_STAR
				End If
			Case optSEXP.Checked : s = PRODUCT_SEXP
			Case optD2DV.Checked : s = PRODUCT_D2DV
			Case optD2XP.Checked : s = PRODUCT_D2XP
			Case optWAR3.Checked : s = PRODUCT_WAR3
			Case optW3XP.Checked : s = PRODUCT_W3XP
			Case optW2BN.Checked : s = PRODUCT_W2BN
			Case optDRTL.Checked
				If (chkSHR.CheckState) Then
					s = PRODUCT_DSHR
				Else
					s = PRODUCT_DRTL
				End If
				'Case optCHAT.Value: s = PRODUCT_CHAT
		End Select
		
		If Not DoCDKeyLengthCheck(txtCDKey.Text, s) Then
			If MsgBox("Your CD key is of an invalid Length for the product you have chosen. Do you want to save anyway?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "StealthBot Settings") = MsgBoxResult.No Then
				ShowPanel(modEnum.enuSettingsPanels.spConnectionConfig)
				txtCDKey.Focus()
				SaveSettings = False
				Exit Function
			End If
		End If
		
		If txtExpKey.Enabled And Not DoCDKeyLengthCheck(txtExpKey.Text, s) Then
			If MsgBox("Your expansion CD key is of an invalid Length for the product you have chosen. Do you want to anyway?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "StealthBot Settings") = MsgBoxResult.No Then
				ShowPanel(modEnum.enuSettingsPanels.spConnectionConfig)
				txtExpKey.Focus()
				SaveSettings = False
				Exit Function
			End If
		End If
		
		Config.Game = s
		
		' The rest of the basic config now
		Config.Username = txtUsername.Text
		Config.Password = txtPassword.Text
		Config.CDKey = CDKeyReplacements(txtCDKey.Text)
		Config.ExpKey = CDKeyReplacements(txtExpKey.Text)
		Config.HomeChannel = txtHomeChan.Text
		Config.Server = cboServer.Text
		
		Config.UseSpawn = CBool(CanSpawn(Config.Game, Len(Config.CDKey)) And CBool(chkSpawn.CheckState))
		Config.UseD2Realms = CBool(chkUseRealm.CheckState)
		
		' Advanced connection settings
		Config.UseBNLS = CBool(cboConnMethod.SelectedIndex = 0)
		Config.BNLSFinder = CBool(cboBNLSServer.SelectedIndex = 0)
		
		If cboBNLSServer.SelectedIndex > 0 Then
			Config.BNLSServer = cboBNLSServer.Text
		ElseIf cboBNLSServer.SelectedIndex = -1 Then 
			If StrComp(cboBNLSServer.Text, BNLS_NOT_SET, CompareMethod.Binary) <> 0 Then
				Config.BNLSServer = cboBNLSServer.Text
				AddBNLSServer(Config.BNLSServer)
			End If
		End If
		
		' Save the BNLS server list
		With cboBNLSServer
			j = -1
			
			' Check if the set server is in the list
			For j = 1 To .Items.Count
				If StrComp(.Text, VB6.GetItemString(cboBNLSServer, j)) = 0 Then
					j = -1
					Exit For
				End If
			Next j
			
			If j >= 0 Or .Items.Count > 0 Then
				For j = 1 To .Items.Count
					If Not (Len(VB6.GetItemString(cboBNLSServer, j)) = 0) Then colBNLS.Add(VB6.GetItemString(cboBNLSServer, j))
				Next 
				
				ListFileSave(GetFilePath(FILE_BNLS_LIST), colBNLS)
			End If
		End With
		
		Config.AutoConnect = CBool(chkConnectOnStartup.CheckState)
		Config.RegisterEmailDefault = Trim(txtEmail.Text)
		Config.PingSpoofing = cboSpoof.SelectedIndex
		Config.UseProxy = CBool(chkUseProxies.CheckState)
		Config.ProxyPort = StringToNumber(txtProxyPort.Text, (Config.ProxyPort))
		Config.ProxyIP = Trim(txtProxyIP.Text)
		Config.ProxyType = IIf(CBool(optSocks5.Checked), "SOCKS5", "SOCKS4")
		Config.UseUDP = CBool(chkUDP.CheckState)
		
		Config.ReconnectDelay = StringToNumber(txtReconDelay.Text, (Config.ReconnectDelay))
		
		' General Interface settings
		Config.ShowSplashScreen = CBool(chkSplash.CheckState)
		Config.MinimizeToTray = Not CBool(chkNoTray.CheckState)
		Config.FlashOnEvents = CBool(chkFlash.CheckState)
		Config.MinimizeOnStartup = CBool(chkMinimizeOnStartup.CheckState)
		
		Config.UseUTF8 = CBool(chkUTF8.CheckState)
		Config.ShowJoinLeaves = CBool(chkJoinLeaves.CheckState)
		Config.ChatFilters = CBool(chkFilter.CheckState)
		Config.UrlDetection = CBool(chkURLDetect.CheckState)
		Config.NameAutoComplete = Not CBool(chkNoAutocomplete.CheckState)
		
		Config.NameColoring = Not CBool(chkNoColoring.CheckState)
		Config.ShowStatsIcons = CBool(chkShowUserGameStatsIcons.CheckState)
		Config.ShowFlagIcons = CBool(chkShowUserFlagsIcons.CheckState)
		
		Config.DisablePrefixBox = CBool(chkDisablePrefix.CheckState)
		Config.DisableSuffixBox = CBool(chkDisableSuffix.CheckState)
		Config.TimestampMode = cboTimestamp.SelectedIndex
		
		' Font and color
		SaveFontSettings()
		
		' Moderation settings
		Config.IPBans = CBool(chkIPBans.CheckState)
		Config.UDPBan = CBool(chkPlugban.CheckState)
		Config.KickOnYell = CBool(chkKOY.CheckState)
		Config.BanEvasion = CBool(chkBanEvasion.CheckState)
		
		Config.QuietTime = CBool(chkQuiet.CheckState)
		Config.QuietTimeKick = CBool(chkQuietKick.CheckState)
		
		Config.PhraseBans = CBool(chkPhrasebans.CheckState)
		Config.PhraseKick = CBool(chkPhraseKick.CheckState)
		
		Config.PingBan = CBool(chkPingBan.CheckState)
		Config.PingBanLevel = StringToNumber(txtPingLevel.Text, (Config.PingBanLevel))
		
		Config.ChannelProtection = CBool(chkProtect.CheckState)
		Config.ChannelProtectionMessage = txtProtectMsg.Text
		
		Config.IdleBan = CBool(chkIdlebans.CheckState)
		Config.IdleBanKick = CBool(chkIdleKick.CheckState)
		Config.IdleBanDelay = StringToNumber(txtIdleBanDelay.Text, (Config.IdleBanDelay))
		
		Call SaveClientBans()
		
		Config.PeonBan = CBool(chkPeonbans.CheckState)
		
		Config.LevelBanW3 = StringToNumber(txtBanW3.Text, (Config.LevelBanW3))
		Config.LevelBanD2 = StringToNumber(txtBanD2.Text, (Config.LevelBanD2))
		Config.LevelBanMessage = txtLevelBanMsg.Text
		
		'// Logging options
		If (cboLogging.SelectedIndex = 0) Then
			Config.LoggingMode = 2
		ElseIf (cboLogging.SelectedIndex = 1) Then 
			Config.LoggingMode = 1
		ElseIf (cboLogging.SelectedIndex = 2) Then 
			Config.LoggingMode = 0
		End If
		
		Config.LogDBActions = CBool(chkLogDBActions.CheckState)
		Config.LogCommands = CBool(chkLogAllCommands.CheckState)
		
		Config.MaxBacklogSize = StringToNumber(txtMaxBackLogSize.Text, (Config.MaxBacklogSize))
		Config.MaxLogFileSize = StringToNumber(txtMaxLogSize.Text, (Config.MaxLogFileSize))
		
		' Greet Message Settings
		Config.GreetMessageText = txtGreetMsg.Text
		Config.WhisperGreet = CBool(chkWhisperGreet.CheckState)
		Config.GreetMessage = CBool(chkGreetMsg.CheckState)
		
		' Idle message settings
		Config.IdleMessage = CBool(chkIdles.CheckState)
		If IsNumeric(txtIdleWait.Text) Then
			Config.IdleMessageDelay = CInt(CDbl(txtIdleWait.Text) * 2)
		End If
		
		Select Case True
			Case optMsg.Checked : Config.IdleMessageType = "msg"
			Case optUptime.Checked : Config.IdleMessageType = "uptime"
			Case optMP3.Checked : Config.IdleMessageType = "mp3"
			Case optQuote.Checked : Config.IdleMessageType = "quote"
			Case Else : Config.IdleMessageType = "msg"
		End Select
		Config.IdleMessageText = txtIdleMsg.Text
		
		
		'// Misc General Settings
		Config.Mp3Commands = CBool(chkAllowMP3.CheckState)
		Config.ProfileAmp = CBool(chkPAmp.CheckState)
		Config.BotMail = CBool(chkMail.CheckState)
		
		Config.BotOwner = txtOwner.Text
		Config.Trigger = txtTrigger.Text
		
		Config.WhisperCommands = CBool(chkWhisperCmds.CheckState)
		Config.ShowOfflineFriends = CBool(chkShowOffline.CheckState)
		Config.UseBackupChannel = CBool(chkBackup.CheckState)
		Config.BackupChannel = txtBackupChan.Text
		
		For i = 0 To 3
			If optNaming(i).Checked Then Exit For ' i = index of opt checked
		Next i
		If i = 4 Then i = 0 ' if none were checked, then set to default
		Config.NamespaceConvention = i
		
		Config.UseD2Naming = CBool(chkD2Naming.CheckState)
		
		'// Save the config instance to disk
		Call Config.Save()
		
		'// Load the config into the form
		Call frmChat.ReloadConfig(1)
		
		'// RESIZE FORM TO FIX ANY UI CHANGES
		Call frmChat.frmChat_Resize(Nothing, New System.EventArgs())
		
		'// Take care of the colors.
		If ModifiedColors Then
			Call SaveColors()
			Call GetColorLists()
		End If
		
		SaveSettings = True
	End Function
	
	' check for potential invalid config stuffs
	Public Function InvalidConfigValues() As Boolean
		
		Dim s As String
		
		If optW3XP.Checked Or optD2XP.Checked Then
            If Len(txtExpKey.Text) = 0 Then
                If optW3XP.Checked Then
                    s = "Warcraft III and a Frozen Throne"
                Else
                    s = "Diablo II and a Lord of Destruction"
                End If

                MsgBox("You must enter both a " & s & " CD-key to connect with an Expansion game.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                ShowPanel(modEnum.enuSettingsPanels.spConnectionConfig)
                txtExpKey.Focus()
                InvalidConfigValues = True
            End If
		End If
		
	End Function
	
	
	'##########################################
	' COLOR-RELATED CODE
	'##########################################
	
	Sub CAdd(ByVal colorName As String, ByRef ColorValue As Integer, Optional ByRef Append As Byte = 0)
		mColors(UBound(mColors)) = ColorValue
		
		If Append = 1 Then colorName = "Event | " & colorName
		
		cboColorList.Items.Add(colorName)
		
		ReDim Preserve mColors(UBound(mColors) + 1)
	End Sub
	
	Private Sub cmdExport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExport.Click
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With New SaveFileDialog
            .FileName = vbNullString
            .ShowDialog()
            If .FileName <> vbNullString Then
                SaveColors(.FileName)
                MsgBox("ColorList exported.", MsgBoxStyle.OkOnly)
            End If
        End With
	End Sub
	
	Private Sub cmdImport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImport.Click
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With New OpenFileDialog
            .FileName = vbNullString
            .ShowDialog()
            If .FileName <> vbNullString Then
                GetColorLists((.FileName))
                cboColorList.Items.Clear()
                Call frmSettings_Load(Me, New System.EventArgs())
            End If
        End With
	End Sub
	
	Private Sub cmdGetRGB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGetRGB.Click
		On Error Resume Next
		txtValue.Text = CStr(RGB(CInt(txtR.Text), CInt(txtG.Text), CInt(txtB.Text)))
	End Sub
	
	Private Sub cmdHTMLGen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHTMLGen.Click
		If VB.Left(txtHTML.Text, 1) = "#" Then txtHTML.Text = Mid(txtHTML.Text, 2)
		
		txtValue.Text = CStr(HTMLToRGBColor(txtHTML.Text))
	End Sub
	
	Private Sub cmdDefaults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDefaults.Click
		If MsgBox("Are you sure you want to restore the default Values()?" & vbCrLf & "(All current color data will be lost unless exported)", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation) = MsgBoxResult.Yes Then
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Len(Dir(GetFilePath(FILE_COLORS))) > 0 Then
                Kill(GetFilePath(FILE_COLORS))
                Call GetColorLists()
                Call LoadColors()
            End If
		End If
	End Sub
	
	Private Sub SaveColors(Optional ByRef sPath As String = "")
		Dim f As Short
		Dim i As Short
		
		f = FreeFile
		
        If Len(sPath) = 0 Then sPath = GetFilePath(FILE_COLORS)
		
		FileOpen(f, sPath, OpenMode.Random, , , 4)
		
		For i = LBound(mColors) To UBound(mColors)
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(f, CInt(mColors(i)), i + 1)
			'Debug.Print "Putting color; " & i & ":" & mColors(i)
		Next i
		
		FileClose(f)
	End Sub
	
	'UPGRADE_WARNING: Event txtValue.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtValue_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtValue.TextChanged
		On Error Resume Next
		lblEg.BackColor = System.Drawing.ColorTranslator.FromOle(Val(txtValue.Text))
	End Sub
	
	Private Sub cmdColorPicker_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdColorPicker.Click
		cDLGColor.ShowDialog()
		txtValue.Text = System.Drawing.ColorTranslator.ToOle(cDLGColor.Color).ToString
	End Sub
	
	
	Sub ShowCurrentColor()
		On Error GoTo ShowCurrentColor_Error
		
		lblEg.BackColor = System.Drawing.ColorTranslator.FromOle(mColors(cboColorList.SelectedIndex))
		txtValue.Text = CStr(mColors(cboColorList.SelectedIndex))
		lblCurrentValue.Text = CStr(mColors(cboColorList.SelectedIndex))
		
ShowCurrentColor_Exit: 
		Exit Sub
		
ShowCurrentColor_Error: 
		
		Debug.Print("Error " & Err.Number & " (" & Err.Description & ") in procedure ShowCurrentColor of Form frmSettings")
		Resume ShowCurrentColor_Exit
	End Sub
	
	Private Sub RecordCurrentColor()
		If cboColorList.SelectedIndex > -1 Then
			mColors(cboColorList.SelectedIndex) = Val(txtValue.Text)
			'Debug.Print "Recording current color."
		End If
	End Sub
	
	Private Sub cboColorList_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboColorList.Enter
		ShowCurrentColor()
	End Sub
	
	'UPGRADE_ISSUE: ComboBox event cboColorList.Scroll was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub cboColorList_Scroll()
		ShowCurrentColor()
	End Sub
	
	'UPGRADE_WARNING: Event cboColorList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboColorList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboColorList.SelectedIndexChanged
		ModifiedColors = True
		lblColorStatus.Text = "Be sure to click 'Save Changes to This Color' before proceeding."
		ShowCurrentColor()
	End Sub
	
	
	
	'##########################################
	' ENABLE/DISABLE CODE
	'##########################################
	
	'UPGRADE_WARNING: Event chkUseProxies.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkUseProxies_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUseProxies.CheckStateChanged
		txtProxyIP.Enabled = chkUseProxies.CheckState
		txtProxyPort.Enabled = chkUseProxies.CheckState
		optSocks4.Enabled = chkUseProxies.CheckState
		optSocks5.Enabled = chkUseProxies.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkBackup.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkBackup_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkBackup.CheckStateChanged
		txtBackupChan.Enabled = chkBackup.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkIdlebans.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIdlebans_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIdlebans.CheckStateChanged
		chkIdleKick.Enabled = chkIdlebans.CheckState
		txtIdleBanDelay.Enabled = chkIdlebans.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkIdles.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIdles_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIdles.CheckStateChanged
		optMsg.Enabled = chkIdles.CheckState
		optUptime.Enabled = chkIdles.CheckState
		optMP3.Enabled = chkIdles.CheckState
		optQuote.Enabled = chkIdles.CheckState
		txtIdleWait.Enabled = chkIdles.CheckState
		txtIdleMsg.Enabled = (optMsg.Enabled And optMsg.Checked)
	End Sub
	
	'UPGRADE_WARNING: Event optMsg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMsg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMsg.CheckedChanged
		If eventSender.Checked Then
			txtIdleMsg.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optUptime.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optUptime_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUptime.CheckedChanged
		If eventSender.Checked Then
			txtIdleMsg.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMP3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMP3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMP3.CheckedChanged
		If eventSender.Checked Then
			txtIdleMsg.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optQuote.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optQuote_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optQuote.CheckedChanged
		If eventSender.Checked Then
			txtIdleMsg.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSTAR.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optSTAR_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSTAR.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = True
			chkSpawn.Enabled = True
			chkJPN.Visible = True
			txtCDKey.Enabled = True
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = False
			If (chkSHR.CheckState) Then
				SetHashPath(GetGamePath("RHSS"))
				chkSpawn.Enabled = False
				txtCDKey.Enabled = False
			ElseIf (chkJPN.CheckState) Then 
				SetHashPath(GetGamePath("RTSJ"))
			Else
				SetHashPath(GetGamePath("RATS"))
			End If
			chkUDP.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optWAR3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optWAR3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optWAR3.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = False
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = False
			SetHashPath(GetGamePath("3RAW"))
			chkUDP.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optD2DV.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optD2DV_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optD2DV.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = False
			chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = True
			SetHashPath(GetGamePath("VD2D"))
			chkUDP.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optW2BN.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optW2BN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optW2BN.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = True
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = False
			SetHashPath(GetGamePath("NB2W"))
			chkUDP.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSEXP.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optSEXP_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSEXP.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = False
			chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = False
			SetHashPath(GetGamePath("RATS"))
			chkUDP.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optD2XP.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optD2XP_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optD2XP.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = False
			chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = True
			chkUseRealm.Enabled = True
			SetHashPath(GetGamePath("PX2D"))
			chkUDP.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optW3XP.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optW3XP_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optW3XP.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = False
			chkSpawn.Enabled = False
			chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkJPN.Visible = False
			txtCDKey.Enabled = True
			txtExpKey.Enabled = True
			chkUseRealm.Enabled = False
			SetHashPath(GetGamePath("PX3W"))
			chkUDP.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optDRTL.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub optDRTL_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDRTL.CheckedChanged
		If eventSender.Checked Then
			chkSHR.Visible = True
			chkSpawn.Enabled = False
			chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkJPN.Visible = False
			txtCDKey.Enabled = False
			txtExpKey.Enabled = False
			chkUseRealm.Enabled = False
			If (chkSHR.CheckState) Then
				SetHashPath(GetGamePath("RHSD"))
			Else
				SetHashPath(GetGamePath("LTRD"))
			End If
			chkUDP.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkJPN.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkJPN_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkJPN.CheckStateChanged
		Dim Checked As Boolean
		Checked = CBool(chkJPN.CheckState)
		If (Checked) Then chkSHR.CheckState = System.Windows.Forms.CheckState.Unchecked
		If (optSTAR.Checked) Then
			If (Checked) Then
				SetHashPath(GetGamePath("RTSJ"))
			Else
				SetHashPath(GetGamePath("RATS"))
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkSHR.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSHR_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSHR.CheckStateChanged
		Dim Checked As Boolean
		Checked = CBool(chkSHR.CheckState)
		If (Checked) Then chkJPN.CheckState = System.Windows.Forms.CheckState.Unchecked
		If (optSTAR.Checked) Then
			chkSpawn.Enabled = Not Checked
			txtCDKey.Enabled = Not Checked
			If (Checked) Then
				SetHashPath(GetGamePath("RHSS"))
			Else
				SetHashPath(GetGamePath("RATS"))
			End If
		ElseIf (optDRTL.Checked) Then 
			If (Checked) Then
				SetHashPath(GetGamePath("RHSD"))
			Else
				SetHashPath(GetGamePath("LTRD"))
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkGreetMsg.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkGreetMsg_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkGreetMsg.CheckStateChanged
		chkWhisperGreet.Enabled = chkGreetMsg.CheckState
		txtGreetMsg.Enabled = chkGreetMsg.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkProtect.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkProtect_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProtect.CheckStateChanged
		txtProtectMsg.Enabled = chkProtect.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkQuiet.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkQuiet_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkQuiet.CheckStateChanged
		chkQuietKick.Enabled = chkQuiet.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkPhrasebans.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkPhrasebans_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPhrasebans.CheckStateChanged
		chkPhraseKick.Enabled = chkPhrasebans.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkPingBan.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkPingBan_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPingBan.CheckStateChanged
		txtPingLevel.Enabled = chkPingBan.CheckState
	End Sub
	
	'##########################################
	' INIT SUBS
	'##########################################
	
	Private Sub InitAllPanels()
		InitGenMisc()
		InitConnAdvanced()
		InitFontsColors()
		InitGenInterface()
		InitGenMod()
		InitBasicConfig()
		InitGenGreets()
		InitGenIdles()
		InitLogging()
		
		If Config.FileExists Then
			ShowPanel(modEnum.enuSettingsPanels.spConnectionConfig)
			FirstRun = 1
		End If
	End Sub
	
	Private Sub InitBasicConfig()
		Dim i As Short
		Dim AddCurrent As Boolean
		Dim Item As String
		Dim AdditionalServerList As Collection
		
		txtUsername.Text = Config.Username
		txtPassword.Text = Config.Password
		txtCDKey.Text = Config.CDKey
		txtExpKey.Text = Config.ExpKey
		
		txtHomeChan.Text = Config.HomeChannel
		
		With cboServer
			' add the 4 default servers
			.Items.Add("useast.battle.net")
			.Items.Add("uswest.battle.net")
			.Items.Add("europe.battle.net")
			.Items.Add("asia.battle.net")
			
			' get additional servers
			AdditionalServerList = ListFileLoad(GetFilePath(FILE_SERVER_LIST))
			
			' if additional servers, add a blank line, then add them
			If AdditionalServerList.Count() > 0 Then
				.Items.Add("")
				
				For i = 1 To AdditionalServerList.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object AdditionalServerList.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Items.Add(AdditionalServerList.Item(i))
				Next i
			End If
			
			' check if "currently selected" is in list
			AddCurrent = True
			
			For i = 0 To .Items.Count - 1
				Item = VB6.GetItemString(cboServer, i)
				
				If StrComp(Item, Config.Server, CompareMethod.Binary) = 0 Then
					AddCurrent = False
					.SelectedIndex = i
				End If
			Next i
			
			' if not, add it (first)
			If AddCurrent Then
				.Items.Insert(0, Config.Server)
				.SelectedIndex = 0
			End If
		End With
		
		Select Case GetProductInfo(Config.Game).Code
			Case PRODUCT_STAR : Call optSTAR_CheckedChanged(optSTAR, New System.EventArgs()) : optSTAR.Checked = True : chkSHR.CheckState = System.Windows.Forms.CheckState.Unchecked : chkJPN.CheckState = System.Windows.Forms.CheckState.Unchecked
			Case PRODUCT_SEXP : Call optSEXP_CheckedChanged(optSEXP, New System.EventArgs()) : optSEXP.Checked = True
			Case PRODUCT_D2DV : Call optD2DV_CheckedChanged(optD2DV, New System.EventArgs()) : optD2DV.Checked = True
			Case PRODUCT_D2XP : Call optD2XP_CheckedChanged(optD2XP, New System.EventArgs()) : optD2XP.Checked = True
			Case PRODUCT_W2BN : Call optW2BN_CheckedChanged(optW2BN, New System.EventArgs()) : optW2BN.Checked = True
			Case PRODUCT_WAR3 : Call optWAR3_CheckedChanged(optWAR3, New System.EventArgs()) : optWAR3.Checked = True
			Case PRODUCT_W3XP : Call optW3XP_CheckedChanged(optW3XP, New System.EventArgs()) : optW3XP.Checked = True
			Case PRODUCT_DRTL : Call optDRTL_CheckedChanged(optDRTL, New System.EventArgs()) : optDRTL.Checked = True : chkSHR.CheckState = System.Windows.Forms.CheckState.Unchecked
			Case PRODUCT_DSHR : Call optDRTL_CheckedChanged(optDRTL, New System.EventArgs()) : optDRTL.Checked = True : chkSHR.CheckState = System.Windows.Forms.CheckState.Checked
			Case PRODUCT_SSHR : Call optSTAR_CheckedChanged(optSTAR, New System.EventArgs()) : optSTAR.Checked = True : chkSHR.CheckState = System.Windows.Forms.CheckState.Checked ' unchecks jpn
			Case PRODUCT_JSTR : Call optSTAR_CheckedChanged(optSTAR, New System.EventArgs()) : optSTAR.Checked = True : chkJPN.CheckState = System.Windows.Forms.CheckState.Checked ' unchecks shr
			Case Else : Call optSTAR_CheckedChanged(optSTAR, New System.EventArgs()) : optSTAR.Checked = True : chkSHR.CheckState = System.Windows.Forms.CheckState.Unchecked : chkJPN.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
		chkSpawn.CheckState = System.Math.Abs(CInt(Config.UseSpawn))
		chkUseRealm.CheckState = System.Math.Abs(CInt(Config.UseD2Realms))
		
	End Sub
	
	Private Sub InitConnAdvanced()
		' Connection method
		cboConnMethod.SelectedIndex = CShort(System.Math.Abs(CInt(Not Config.UseBNLS)))
		cboBNLSServer.Enabled = System.Math.Abs(CInt(Config.UseBNLS))
		
		' Set selected BNLS server
		If Config.BNLSFinder Then
			cboBNLSServer.SelectedIndex = 0
		ElseIf Len(Config.BNLSServer) = 0 Then 
			cboBNLSServer.Text = BNLS_NOT_SET
		ElseIf Len(Config.BNLSServer) > 0 Then 
			cboBNLSServer.SelectedIndex = GetBnlsIndex(Config.BNLSServer)
		End If
		
		txtEmail.Text = Config.RegisterEmailDefault
		
		cboSpoof.SelectedIndex = Config.PingSpoofing
		chkUDP.CheckState = System.Math.Abs(CInt(Config.UseUDP))
		
		chkConnectOnStartup.CheckState = System.Math.Abs(CInt(Config.AutoConnect))
		txtReconDelay.Text = CStr(Config.ReconnectDelay)
        If Len(txtReconDelay.Text) = 0 Then txtReconDelay.Text = CStr(1000)
		
		chkUseProxies.CheckState = System.Math.Abs(CInt(Config.UseProxy))
		Call chkUseProxies_CheckStateChanged(chkUseProxies, New System.EventArgs())
		
		txtProxyPort.Text = CStr(Config.ProxyPort)
		txtProxyIP.Text = Config.ProxyIP
		
		Select Case UCase(Config.ProxyType)
			Case "SOCKS5"
				optSocks5.Checked = True
				optSocks4.Checked = False
			Case "SOCKS4"
				optSocks5.Checked = False
				optSocks4.Checked = True
		End Select
		
		' Adjust "BNLS server" label 2 pixels down
		lbl5(12).Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lbl5(12).Top) + (2 * VB6.TwipsPerPixelY))
	End Sub
	
	Private Sub InitGenInterface()
		chkSplash.CheckState = System.Math.Abs(CInt(Config.ShowSplashScreen))
		chkNoTray.CheckState = System.Math.Abs(CInt(Not Config.MinimizeToTray))
		chkFlash.CheckState = System.Math.Abs(CInt(Config.FlashOnEvents))
		chkMinimizeOnStartup.CheckState = System.Math.Abs(CInt(Config.MinimizeOnStartup))
		
		chkUTF8.CheckState = System.Math.Abs(CInt(Config.UseUTF8))
		chkJoinLeaves.CheckState = System.Math.Abs(CInt(Config.ShowJoinLeaves))
		chkFilter.CheckState = System.Math.Abs(CInt(Config.ChatFilters))
		chkURLDetect.CheckState = System.Math.Abs(CInt(Config.UrlDetection))
		chkNoAutocomplete.CheckState = System.Math.Abs(CInt(Not Config.NameAutoComplete))
		
		chkNoColoring.CheckState = System.Math.Abs(CInt(Not Config.NameColoring))
		chkShowUserGameStatsIcons.CheckState = System.Math.Abs(CInt(Config.ShowStatsIcons))
		chkShowUserFlagsIcons.CheckState = System.Math.Abs(CInt(Config.ShowFlagIcons))
		
		chkDisablePrefix.CheckState = System.Math.Abs(CInt(Config.DisablePrefixBox))
		chkDisableSuffix.CheckState = System.Math.Abs(CInt(Config.DisableSuffixBox))
		cboTimestamp.SelectedIndex = Config.TimestampMode
	End Sub
	
	Private Sub InitFontsColors()
		txtChatFont.Text = frmChat.rtbChat.Font.Name
		InitChatFont = txtChatFont.Text
		
		txtChanFont.Text = frmChat.lvChannel.Font.Name
		InitChanFont = txtChanFont.Text
		
		txtChatSize.Text = CStr(CShort(frmChat.rtbChat.Font.SizeInPoints))
		InitChatSize = CShort(frmChat.rtbChat.Font.SizeInPoints)
		
		txtChanSize.Text = CStr(CShort(frmChat.lvChannel.Font.SizeInPoints))
		InitChanSize = CShort(frmChat.lvChannel.Font.SizeInPoints)
		
		cboColorList.SelectedIndex = 0
	End Sub
	
	Private Sub InitGenMod()
		chkIPBans.CheckState = System.Math.Abs(CInt(Config.IPBans))
		chkPlugban.CheckState = System.Math.Abs(CInt(Config.UDPBan))
		chkKOY.CheckState = System.Math.Abs(CInt(Config.KickOnYell))
		chkBanEvasion.CheckState = System.Math.Abs(CInt(Config.BanEvasion))
		
		chkQuiet.CheckState = System.Math.Abs(CInt(Config.QuietTime))
		chkQuietKick.CheckState = System.Math.Abs(CInt(Config.QuietTimeKick))
		Call chkQuiet_CheckStateChanged(chkQuiet, New System.EventArgs())
		
		chkPhrasebans.CheckState = System.Math.Abs(CInt(Config.PhraseBans))
		chkPhraseKick.CheckState = System.Math.Abs(CInt(Config.PhraseKick))
		Call chkPhrasebans_CheckStateChanged(chkPhrasebans, New System.EventArgs())
		
		chkPingBan.CheckState = System.Math.Abs(CInt(Config.PingBan))
		txtPingLevel.Text = CStr(Config.PingBanLevel)
		Call chkPingBan_CheckStateChanged(chkPingBan, New System.EventArgs())
		
		chkProtect.CheckState = System.Math.Abs(CInt(Config.ChannelProtection))
		txtProtectMsg.Text = Config.ChannelProtectionMessage
		Call chkProtect_CheckStateChanged(chkProtect, New System.EventArgs())
		
		chkIdlebans.CheckState = System.Math.Abs(CInt(Config.IdleBan))
		chkIdleKick.CheckState = System.Math.Abs(CInt(Config.IdleBanKick))
		txtIdleBanDelay.Text = CStr(Config.IdleBanDelay)
		Call chkIdlebans_CheckStateChanged(chkIdlebans, New System.EventArgs())
		
		' grab client ban settings from database
		chkCBan(SC).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_STAR)))
		chkCBan(BW).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_SEXP)))
		chkCBan(D2).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_D2DV)))
		chkCBan(D2X).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_D2XP)))
		chkCBan(W2).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_W2BN)))
		chkCBan(W3).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_WAR3)))
		chkCBan(W3X).CheckState = System.Math.Abs(CInt(IsClientBanned(PRODUCT_W3XP)))
		
		chkPeonbans.CheckState = System.Math.Abs(CInt(Config.PeonBan))
		
		txtBanD2.Text = CStr(Config.LevelBanD2)
		txtBanW3.Text = CStr(Config.LevelBanW3)
		txtLevelBanMsg.Text = Config.LevelBanMessage
	End Sub
	
	Private Sub InitLogging()
		Select Case Config.LoggingMode
			Case 0
				cboLogging.SelectedIndex = 2
			Case 1
				cboLogging.SelectedIndex = 1
			Case Else
				cboLogging.SelectedIndex = 0
		End Select
		
		chkLogDBActions.CheckState = System.Math.Abs(CInt(Config.LogDBActions))
		chkLogAllCommands.CheckState = System.Math.Abs(CInt(Config.LogCommands))
		
		txtMaxBackLogSize.Text = CStr(Config.MaxBacklogSize)
		txtMaxLogSize.Text = CStr(Config.MaxLogFileSize)
	End Sub
	
	Private Sub InitGenGreets()
		txtGreetMsg.Text = Config.GreetMessageText
		chkGreetMsg.CheckState = System.Math.Abs(CInt(Config.GreetMessage))
		Call chkGreetMsg_CheckStateChanged(chkGreetMsg, New System.EventArgs())
		
		chkWhisperGreet.CheckState = System.Math.Abs(CInt(Config.WhisperGreet))
	End Sub
	
	Private Sub InitGenIdles()
		txtIdleWait.Text = CStr(Config.IdleMessageDelay / 2)
		
		Select Case Config.IdleMessageType
			Case "msg", vbNullString
				optMsg.Checked = True
				Call optMsg_CheckedChanged(optMsg, New System.EventArgs())
			Case "quote"
				optQuote.Checked = True
				Call optQuote_CheckedChanged(optQuote, New System.EventArgs())
			Case "uptime"
				optUptime.Checked = True
				Call optUptime_CheckedChanged(optUptime, New System.EventArgs())
			Case "mp3"
				optMP3.Checked = True
				Call optMP3_CheckedChanged(optMP3, New System.EventArgs())
			Case Else
				optMsg.Checked = True
				Call optMsg_CheckedChanged(optMsg, New System.EventArgs())
		End Select
		
		txtIdleMsg.Text = Config.IdleMessageText
        If Len(txtIdleMsg.Text) = 0 Then txtIdleMsg.Text = "/me is a %v by Stealth - http://www.stealthbot.net"
		
		chkIdles.CheckState = System.Math.Abs(CInt(Config.IdleMessage))
		Call chkIdles_CheckStateChanged(chkIdles, New System.EventArgs())
		
	End Sub
	
	Private Sub InitGenMisc()
		chkAllowMP3.CheckState = System.Math.Abs(CInt(Config.Mp3Commands))
		chkPAmp.CheckState = System.Math.Abs(CInt(Config.ProfileAmp))
		chkMail.CheckState = System.Math.Abs(CInt(Config.BotMail))
		
		chkWhisperCmds.CheckState = System.Math.Abs(CInt(Config.WhisperCommands))
		chkShowOffline.CheckState = System.Math.Abs(CInt(Config.ShowOfflineFriends))
		
		txtOwner.Text = Config.BotOwner
		txtTrigger.Text = Config.Trigger
		
		chkBackup.CheckState = System.Math.Abs(CInt(Config.UseBackupChannel))
		Call chkBackup_CheckStateChanged(chkBackup, New System.EventArgs())
		txtBackupChan.Text = Config.BackupChannel
		
		optNaming(Config.NamespaceConvention).Checked = True
		chkD2Naming.CheckState = System.Math.Abs(CInt(Config.UseD2Naming))
	End Sub
	
	' END INIT SUBS
	
	
	Function CDKeyReplacements(ByVal inString As String) As String
		inString = Replace(inString, "-", "")
		inString = Replace(inString, " ", "")
		CDKeyReplacements = Trim(inString)
	End Function
	
	' Returns true if the client is banned
	Private Function IsClientBanned(ByVal sProductCode As String) As Boolean
		IsClientBanned = (InStr(1, GetAccess(UCase(sProductCode), "GAME").Flags, "B", CompareMethod.Binary) > 0)
	End Function
	
	Private Sub SaveClientBans()
		Dim Clients(6) As String
		Dim i, j As Short
		
		Clients(SC) = PRODUCT_STAR
		Clients(BW) = PRODUCT_SEXP
		Clients(D2) = PRODUCT_D2DV
		Clients(D2X) = PRODUCT_D2XP
		Clients(W3) = PRODUCT_WAR3
		Clients(W3X) = PRODUCT_W3XP
		Clients(W2) = PRODUCT_W2BN
		
		For i = 0 To 6
			If (chkCBan(i).CheckState = 1) Then
				If (GetAccess(Clients(i), "GAME").Username = vbNullString) Then
					
					' redefine array size
					ReDim Preserve DB(UBound(DB) + 1)
					
					With DB(UBound(DB))
						.Username = Clients(i)
						.Flags = "B"
						.ModifiedBy = "(console)"
						.ModifiedOn = Now
						.AddedBy = "(console)"
						.AddedOn = Now
						.Type = "GAME"
					End With
					
					' commit modifications
					Call WriteDatabase(GetFilePath(FILE_USERDB))
					
					' log actions
					If (BotVars.LogDBActions) Then
						Call LogDbAction(modEnum.enuDBActions.AddEntry, "console", DB(UBound(DB)).Username, "game", DB(UBound(DB)).Rank, DB(UBound(DB)).Flags)
					End If
				Else
					For j = LBound(DB) To UBound(DB)
						If ((StrComp(DB(j).Username, Clients(i), CompareMethod.Text) = 0) And (StrComp(DB(j).Type, "GAME", CompareMethod.Text) = 0)) Then
							
							If (InStr(1, DB(j).Flags, "B", CompareMethod.Binary) = 0) Then
								With DB(j)
									.Username = Clients(i)
									.Flags = "B" & .Flags
									.ModifiedBy = "(console)"
									.ModifiedOn = Now
								End With
								
								' log actions
								If (BotVars.LogDBActions) Then
									Call LogDbAction(modEnum.enuDBActions.ModEntry, "console", DB(j).Username, "game", DB(j).Rank, DB(j).Flags)
								End If
								
								' commit modifications
								Call WriteDatabase(GetFilePath(FILE_USERDB))
								
								' break loop
								Exit For
							End If
						End If
					Next j
				End If
			Else
				If (GetAccess(Clients(i), "GAME").Username <> vbNullString) Then
					
					For j = LBound(DB) To UBound(DB)
						If ((StrComp(DB(j).Username, Clients(i), CompareMethod.Text) = 0) And (StrComp(DB(j).Type, "GAME", CompareMethod.Text) = 0)) Then
							
							If ((Len(DB(j).Flags) > 1) Or (DB(j).Rank > 0) Or (Len(DB(j).Groups) > 1)) Then
								
								With DB(j)
									.Username = Clients(i)
									.Flags = Replace(.Flags, "B", vbNullString)
									.ModifiedBy = "(console)"
									.ModifiedOn = Now
								End With
								
								' log actions
								If (BotVars.LogDBActions) Then
									Call LogDbAction(modEnum.enuDBActions.ModEntry, "console", DB(j).Username, "game", DB(j).Rank, DB(j).Flags)
								End If
								
								' commit modifications
								Call WriteDatabase(GetFilePath(FILE_USERDB))
							Else
								Call RemoveItem(Clients(i), "users", "GAME")
								
								' log actions
								If (BotVars.LogDBActions) Then
									Call LogDbAction(modEnum.enuDBActions.RemEntry, "console", DB(j).Username, "game", DB(j).Rank, DB(j).Flags)
								End If
								
								' reload database entries
								Call LoadDatabase()
							End If
							
							' break loop
							Exit For
						End If
					Next j
				End If
			End If
		Next i
	End Sub
	
	Sub SaveFontSettings()
		Dim ResizeChatElements As Boolean
		Dim ResizeChannelElements As Boolean
		
		If (StrComp(InitChatFont, txtChatFont.Text, CompareMethod.Text) <> 0) Then
			Config.ChatFont = txtChatFont.Text
			
			'frmChat.rtbChat.Font.Name = Config.ChatFont
			frmChat.cboSend.Font = VB6.FontChangeName(frmChat.cboSend.Font, Config.ChatFont)
			frmChat.txtPre.Font = VB6.FontChangeName(frmChat.txtPre.Font, Config.ChatFont)
			frmChat.txtPost.Font = VB6.FontChangeName(frmChat.txtPost.Font, Config.ChatFont)
			'frmChat.rtbWhispers.Font.Name = Config.ChatFont
			ResizeChatElements = True
		End If
		
		If Not InitChatSize = CShort(txtChatSize.Text) Then
			Config.ChatFontSize = Val(txtChatSize.Text)
			
			'frmChat.rtbChat.Font.Size = Config.ChatFontSize
			frmChat.cboSend.Font = VB6.FontChangeSize(frmChat.cboSend.Font, Config.ChatFontSize)
			frmChat.txtPre.Font = VB6.FontChangeSize(frmChat.txtPre.Font, Config.ChatFontSize)
			frmChat.txtPost.Font = VB6.FontChangeSize(frmChat.txtPost.Font, Config.ChatFontSize)
			'frmChat.rtbWhispers.Font.Size = Config.ChatFontSize
			ResizeChatElements = True
		End If
		
		If (StrComp(InitChanFont, txtChanFont.Text, CompareMethod.Text) <> 0) Then
			Config.ChannelListFont = txtChanFont.Text
			
			frmChat.lvChannel.Font = VB6.FontChangeName(frmChat.lvChannel.Font, Config.ChannelListFont)
			frmChat.lvClanList.Font = VB6.FontChangeName(frmChat.lvClanList.Font, Config.ChannelListFont)
			frmChat.lvFriendList.Font = VB6.FontChangeName(frmChat.lvFriendList.Font, Config.ChannelListFont)
			frmChat.ListviewTabs.Font = VB6.FontChangeName(frmChat.ListviewTabs.Font, Config.ChannelListFont)
			frmChat.lblCurrentChannel.Font = VB6.FontChangeName(frmChat.lblCurrentChannel.Font, Config.ChannelListFont)
			ResizeChannelElements = True
		End If
		
		If Not InitChanSize = CShort(txtChanSize.Text) Then
			Config.ChannelListFontSize = Val(txtChanSize.Text)
			
			frmChat.lvChannel.Font = VB6.FontChangeSize(frmChat.lvChannel.Font, Config.ChannelListFontSize)
			frmChat.lvClanList.Font = VB6.FontChangeSize(frmChat.lvClanList.Font, Config.ChannelListFontSize)
			frmChat.lvFriendList.Font = VB6.FontChangeSize(frmChat.lvFriendList.Font, Config.ChannelListFontSize)
			frmChat.ListviewTabs.Font = VB6.FontChangeSize(frmChat.ListviewTabs.Font, Config.ChannelListFontSize)
			frmChat.lblCurrentChannel.Font = VB6.FontChangeSize(frmChat.lblCurrentChannel.Font, Config.ChannelListFontSize)
			ResizeChannelElements = True
		End If
		
		Dim lblHeight As Single
		If ResizeChannelElements Then
			
			frmChat.lblCurrentChannel.AutoSize = True
			lblHeight = VB6.PixelsToTwipsY(frmChat.lblCurrentChannel.Height) + 40
			frmChat.lblCurrentChannel.AutoSize = False
			frmChat.lblCurrentChannel.Height = VB6.TwipsToPixelsY(lblHeight)
			ResizeChatElements = True
		End If
		
		If ResizeChatElements Then
			Call ChangeRTBFont((frmChat.rtbChat), Config.ChatFont, Config.ChannelListFontSize)
			Call ChangeRTBFont((frmChat.rtbWhispers), Config.ChatFont, Config.ChannelListFontSize)
			
			frmChat.frmChat_Resize(Nothing, New System.EventArgs())
		End If
	End Sub
	
	Private Sub ChangeRTBFont(ByRef rtb As System.Windows.Forms.RichTextBox, ByVal NewFont As String, ByVal NewSize As Short)
		Dim tmpBuffer As String
		
		With rtb
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
			.SelectionFont = VB6.FontChangeSize(.SelectionFont, NewSize)
			.SelectionFont = VB6.FontChangeName(.SelectionFont, NewFont)
			tmpBuffer = .RTF
			.Text = vbNullString
			.Font = VB6.FontChangeName(.Font, NewFont)
			.Font = VB6.FontChangeSize(.Font, NewSize)
			'UPGRADE_WARNING: TextRTF was upgraded to Text and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Text = tmpBuffer
			.SelectionStart = Len(.Text)
		End With
	End Sub
	
	Sub LoadColors()
		ReDim mColors(0)
		cboColorList.Items.Clear()
		
		With FormColors
			CAdd("Current Channel Label | Background", .ChannelLabelBack)
			CAdd("Current Channel Label | Text", .ChannelLabelText)
			CAdd("Channel List | Background", .ChannelListBack)
			CAdd("Channel List | Normal Users", .ChannelListText)
			CAdd("Channel List | Self", .ChannelListSelf)
			CAdd("Channel List | Idle Users", .ChannelListIdle)
			CAdd("Channel List | Squelched Users", .ChannelListSquelched)
			CAdd("Channel List | Operators", .ChannelListOps)
			CAdd("Chat Window | Background", .RTBBack)
			CAdd("Send Boxes | Background", .SendBoxesBack)
			CAdd("Send Boxes | Text", .SendBoxesText)
		End With
		
		With RTBColors
			CAdd("Talk - Bot Username", .TalkBotUsername, 1)
			CAdd("Talk - Normal Usernames", .TalkUsernameNormal, 1)
			CAdd("Talk - Op Usernames", .TalkUsernameOp, 1)
			CAdd("Talk - Normal Text", .TalkNormalText, 1)
			CAdd("Talk - Carat Color", .Carats, 1)
			CAdd("Emote - Text", .EmoteText, 1)
			CAdd("Emote - Username", .EmoteUsernames, 1)
			CAdd("Information - Neutral", .InformationText, 1)
			CAdd("Information - Success", .SuccessText, 1)
			CAdd("Information - Errors", .ErrorMessageText, 1)
			CAdd("Information - Timestamps", .TimeStamps, 1)
			CAdd("Information - Server Information", .ServerInfoText, 1)
			CAdd("Information - Console Messages", .ConsoleText, 1)
			CAdd("Join/Leave - Text", .JoinText, 1)
			CAdd("Join/Leave - Username", .JoinUsername, 1)
			CAdd("Channel Join - Name", .JoinedChannelName, 1)
			CAdd("Channel Join - Text", .JoinedChannelText, 1)
			CAdd("Whisper - Carat Color", .WhisperCarats, 1)
			CAdd("Whisper - Text", .WhisperText, 1)
			CAdd("Whisper - Usernames", .WhisperUsernames, 1)
		End With
	End Sub
	
	Private Function StringToNumber(ByVal sNumber As String, Optional ByVal lDefault As Integer = 0) As Integer
		sNumber = Trim(sNumber)
		
		If IsNumeric(sNumber) Then
			StringToNumber = CInt(sNumber)
		Else
			StringToNumber = lDefault
		End If
	End Function
End Class