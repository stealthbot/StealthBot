Option Strict Off
Option Explicit On
Friend Class frmRealm
	Inherits System.Windows.Forms.Form
	
	Private m_Unload_SuccessfulLogin As Boolean
	' ticks until auto choose
	Private m_Ticks As Integer
	' current auto choose target
	Private m_Choice As Short
	' are we on expansion?
	Private m_IsExpansion As Boolean
	' save selected item for refreshes and menu clicking
	Private m_Selection As Short
	' save new character name so m_Selection gets set to the new character we just created
	Private m_NewCharacterName As String
	
	'Private Const WM_NCDESTROY = &H82
	
	'    Unknown& = &H0
	'    Amazon& = &H1
	'    Sorceress& = &H2
	'    Necromancer& = &H3
	'    Paladin& = &H4
	'    Barbarian& = &H5
	'    Druid& = &H6
	'    Assassin& = &H7
	
	Private Sub frmRealm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		
		Me.Icon = frmChat.Icon
		
		' must have a MCPHandler
		If ds.MCPHandler Is Nothing Then
			ds.MCPHandler = New clsMCPHandler
			ds.MCPHandler.IsRealmError = True
			Me.Close()
		End If
		
		' must be on D2
		If (BotVars.Product <> "PX2D" And BotVars.Product <> "VD2D") Then
			ds.MCPHandler.IsRealmError = True
			Me.Close()
		End If
		
		' store if expansion
		m_IsExpansion = (BotVars.Product = "PX2D")
		
		' this is for deciding whether to enter chat after a form close
		m_Unload_SuccessfulLogin = False
		
		With lblRealm(0) ' detail
			.Text = "Please wait..."
			.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
		End With
		
		' subclass for listview...
#If COMPILE_DEBUG = 0 Then
		HookWindowProc(Handle.ToInt32)
		'm_OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SkipDragLVItem)
#End If
		
		' read auto choose settings from handler
		m_Ticks = ds.MCPHandler.AutoChooseWait
		m_Choice = ds.MCPHandler.AutoChooseTarget
		m_Selection = m_Choice
		
		' UI setup
		Call RealmStartupResponse()
		
		' set up char creation defaults
		chkExpansion.Enabled = m_IsExpansion
		chkExpansion.CheckState = IIf(m_IsExpansion, 1, 0)
		chkLadder.CheckState = System.Windows.Forms.CheckState.Checked
		optNewCharType(6).Enabled = m_IsExpansion
		optNewCharType(7).Enabled = m_IsExpansion
		optNewCharType(1).Checked = True
		Call optNewCharType_CheckedChanged(optNewCharType.Item(1), New System.EventArgs())
		
		' view existing
		optViewExisting.Checked = True
		Call CharListResponse()
		
		' setup timer
		If m_Ticks >= 0 Then
			tmrLoginTimeout.Enabled = True
			tmrLoginTimeout_Tick(tmrLoginTimeout, New System.EventArgs())
			
			' hide delete button here as it's been made visible (over seconds remaining in timer)
			' will be made visible by stopping timer (if character is selected)
			btnUpgrade.Visible = False
			btnDelete.Visible = False
		End If
		
		' MCP handler state
		ds.MCPHandler.FormActive = True
	End Sub
	
	Private Sub frmRealm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		Call StopLoginTimer()
	End Sub
	
	' Display message names.
	'Public Function SkipDragLVItem(ByVal hWnd As Long, ByVal Msg As Long, _
	''        ByVal wParam As Long, ByVal lParam As Long) As Long
	'
	'    Dim nm_hdr As NMHDR
	'
	'    ' If we're being destroyed,
	'    ' restore the original WindowProc.
	'    If Msg = WM_NCDESTROY Then
	'        SetWindowLong hWnd, GWL_WNDPROC, OldWindowProc
	'    ElseIf Msg = WM_NOTIFY Then
	'        ' Copy info into the NMHDR structure.
	'        CopyMemory nm_hdr, ByVal lParam, Len(nm_hdr)
	'
	'    End If
	'
	'    NewWindowProc = CallWindowProc( _
	''        OldWindowProc, hWnd, Msg, wParam, _
	''        lParam)
	'End Function
	
	Private Sub AddCharacterItem(ByRef CharacterName As String, ByRef CharacterStats As clsUserStats, ByRef CharacterExpires As Date, ByVal Index As Short)
		
		Dim Expired As Boolean
		Dim NewItem As System.Windows.Forms.ListViewItem
		
		With lvwChars
            If (Len(CharacterName) > 0) Then
                If (FindCharacter(CharacterName) < 0) Then
                    'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    NewItem = .Items.Add(CharacterName, CharacterStats.CharacterTitleAndName, CharacterStats.CharacterClassID + 1)

                    NewItem.Tag = CStr(Index)

                    If CharacterStats.IsExpansionCharacter Then
                        NewItem.ForeColor = System.Drawing.Color.Lime
                    End If

                    If IsDateExpired(CharacterExpires) Then
                        NewItem.ForeColor = System.Drawing.Color.Red
                    End If

                End If
            End If
		End With
	End Sub
	
	Private Sub frmRealm_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		lvwChars.Items.Clear()
		
		tmrLoginTimeout.Enabled = False
		m_Ticks = -1
		
		If ((Not (m_Unload_SuccessfulLogin)) Or (ds.MCPHandler.IsRealmError)) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] Logon cancelled.")
			
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If frmChat.sckMCP.CtlState <> 0 Then
				frmChat.sckMCP.Close()
			End If
			
			SendEnterChatSequence()
			frmChat.mnuRealmSwitch.Enabled = True
			m_Unload_SuccessfulLogin = True
			ds.MCPHandler.IsRealmError = False
		End If
		
		ds.MCPHandler.FormActive = False
		ds.MCPHandler.IsRealmError = False
	End Sub
	
	'UPGRADE_WARNING: Event cboOtherRealms.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboOtherRealms_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboOtherRealms.SelectedIndexChanged
		Dim CurrRealmIndex As Short
		Dim CurrRealmTitle As String
		Dim NewRealmTitle As String
		Dim RealmPassword As String
		
		Call StopLoginTimer()
		
		CurrRealmIndex = ds.MCPHandler.RealmServerSelectedIndex
		CurrRealmTitle = ds.MCPHandler.RealmServerTitle(CurrRealmIndex)
		
		If (StrComp(cboOtherRealms.Text, CurrRealmTitle, CompareMethod.Text) <> 0) Then
			NewRealmTitle = cboOtherRealms.Text
			
			DisableGUI()
			
			' close connection and switch realms
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If frmChat.sckMCP.CtlState <> 0 Then
				frmChat.sckMCP.Close()
			End If
			
			Call ds.MCPHandler.RealmServerLogon(NewRealmTitle)
			
			CharListResponse()
		End If
	End Sub
	
	Private Sub cboOtherRealms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboOtherRealms.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Call StopLoginTimer()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub DisableGUI()
		btnChoose.Enabled = False
		btnDisconnect.Enabled = False
		btnDelete.Visible = False
		btnUpgrade.Visible = False
		cboOtherRealms.Enabled = False
		optCreateNew.Enabled = False
		cmdCreate.Enabled = False
		lvwChars.Items.Clear()
	End Sub
	
	' labels are all in a control array so clicking on any one of them calls this:
	Private Sub label_Click(ByRef Index As Short)
		Call StopLoginTimer()
	End Sub
	
	Private Sub lvwChars_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwChars.DoubleClick
		btnChoose_Click(btnChoose, New System.EventArgs())
	End Sub
	
	'UPGRADE_ISSUE: MSComctlLib.ListView event lvwChars.ItemClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lvwChars_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
		Dim ExpireText As String
		Dim DetailText As String
		
		If Not (lvwChars.FocusedItem Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object lvwChars.SelectedItem.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_Selection = lvwChars.FocusedItem.Tag
			
			ExpireText = GetCharacterExpireText(lvwChars.FocusedItem.Tag)
			DetailText = GetCharacterDetailText(lvwChars.FocusedItem.Tag)
			
			With lblRealm(0) ' detail
				.Text = DetailText
				.ForeColor = System.Drawing.Color.White
			End With
			
			With lblRealm(1) ' expires
				.Text = ExpireText
				If IsDateExpired(ds.MCPHandler.CharacterExpires(lvwChars.FocusedItem.Tag)) Then
					.ForeColor = System.Drawing.Color.Red
				Else
					.ForeColor = System.Drawing.Color.Yellow
				End If
			End With
			
			btnChoose.Enabled = CanChooseCharacter(lvwChars.FocusedItem.Tag)
			btnDelete.Visible = True
			btnUpgrade.Visible = CanUpgradeCharacter(lvwChars.FocusedItem.Tag)
		Else
			' clear detail
			lblRealm(0).Text = vbNullString
			' clear expires
			lblRealm(1).Text = vbNullString
			btnChoose.Enabled = False
			btnDelete.Visible = False
			btnUpgrade.Visible = False
		End If
	End Sub
	
	Private Sub lvwChars_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvwChars.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Call StopLoginTimer()
	End Sub
	
	Private Sub lvwChars_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvwChars.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If (KeyCode = System.Windows.Forms.Keys.Return) Then
			Call btnChoose_Click(btnChoose, New System.EventArgs())
		ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then 
			Call UnloadNormal()
		ElseIf KeyCode = System.Windows.Forms.Keys.Delete Then 
			Call HandleCharacterDelete(lvwChars.FocusedItem)
		ElseIf KeyCode = System.Windows.Forms.Keys.U And Shift = VB6.ShiftConstants.CtrlMask Then 
			Call HandleCharacterUpgrade(lvwChars.FocusedItem)
		End If
	End Sub
	
	Private Sub cmdCreate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCreate.Click
		Dim i As Short
		Dim Flags As Integer
		
		If lvwChars.Items.Count > 7 Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] Your account is full! Delete a character before trying to create another.")
		Else
			If Len(txtCharName.Text) > 2 Then
				Flags = 0
				If chkLadder.CheckState = 1 Then Flags = Flags Or &H40
				If chkExpansion.CheckState = 1 Then Flags = Flags Or &H20
				If chkHardcore.CheckState = 1 Then Flags = Flags Or &H4
				
				For i = 1 To 7
					If optNewCharType(i).Checked = True Then
						m_NewCharacterName = txtCharName.Text
						
						ds.MCPHandler.SEND_MCP_CHARCREATE((i - 1), Flags, txtCharName.Text)
						
						Exit For
					End If
				Next i
			End If
		End If
	End Sub
	
	Private Sub lvwChars_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvwChars.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call StopLoginTimer()
	End Sub
	
	Private Sub lvwChars_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvwChars.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Item As System.Windows.Forms.ListViewItem
		
		m_Selection = -1
		
		If Button = VB6.MouseButtonConstants.RightButton Then
			Item = lvwChars.GetItemAt(x, y)
			
			If Not (Item Is Nothing) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Item.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_Selection = Item.Tag
				
				mnuPopUpgrade.Visible = CanUpgradeCharacter(m_Selection)
				
				'UPGRADE_ISSUE: Form method frmRealm.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				PopupMenu(mnuPop)
			End If
		End If
	End Sub
	
	Public Sub RealmStartupResponse()
		Dim i As Short
		Dim RealmIndex As Short
		Dim RealmTitle As String
		Dim RealmDescr As String
		Dim RealmIP As String
		Dim RealmPort As Short
		
		RealmIndex = ds.MCPHandler.RealmServerSelectedIndex
		RealmTitle = ds.MCPHandler.RealmServerTitle(RealmIndex)
		RealmDescr = ds.MCPHandler.RealmServerDescription(RealmIndex)
		RealmIP = ds.MCPHandler.RealmSelectedServerIP
		RealmPort = ds.MCPHandler.RealmSelectedServerPort
		
		Me.Text = "Realm " & RealmTitle & " - " & RealmDescr & " (" & RealmIP & ":" & CStr(RealmPort) & ")"
		
		' build "other realm" list
		If ds.MCPHandler.RealmServerCount > 1 Then
			cboOtherRealms.Items.Clear()
			
			For i = 0 To ds.MCPHandler.RealmServerCount - 1
				cboOtherRealms.Items.Add(ds.MCPHandler.RealmServerTitle(i))
				
				If (StrComp(RealmTitle, ds.MCPHandler.RealmServerTitle(i), CompareMethod.Text) = 0) Then
					cboOtherRealms.SelectedIndex = i
				End If
			Next i
			
			cboOtherRealms.Text = RealmTitle
			
			' other realms label
			lblRealm(5).Visible = True
			cboOtherRealms.Visible = True
		Else
			' other realms label
			lblRealm(5).Visible = False
			cboOtherRealms.Visible = False
		End If
	End Sub
	
	' form callback for character list
	' also if this form needs to check existing character list
	' updates UI after checking character list
	Public Sub CharListResponse()
		Dim i As Short
		Dim NewSelection As Short
		
		lvwChars.Items.Clear()
		
		optCreateNew.Checked = False
		
		Call StopLoginTimer()
		
		fraCreateNew.Visible = False
		lvwChars.Visible = True
		' clear expires
		lblRealm(1).Text = vbNullString
		btnDelete.Visible = False
		btnUpgrade.Visible = False
		
		cboOtherRealms.Enabled = True
		btnDisconnect.Enabled = True
		cmdCreate.Enabled = True
		btnChoose.Enabled = False
		
		If ds.MCPHandler.RetrievingCharacterList Then
			With lblRealm(0) ' detail
				.Text = "Retrieving characters..."
				.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
			End With
		ElseIf Not ds.MCPHandler.RealmServerConnected Then 
			With lblRealm(0) ' detail
				.Text = "Switching realm..."
				.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
			End With
		Else
			For i = 0 To ds.MCPHandler.CharacterCount - 1
				
				Call AddCharacterItem(ds.MCPHandler.CharacterName(i), ds.MCPHandler.CharacterStats(i), ds.MCPHandler.CharacterExpires(i), i)
				
			Next i
			
			optCreateNew.Enabled = (ds.MCPHandler.CharacterCount < 8)

            If Len(m_NewCharacterName) > 0 Then
                NewSelection = FindCharacter(m_NewCharacterName)

                If NewSelection >= 0 Then m_Selection = NewSelection

                m_NewCharacterName = vbNullString
            End If
			
			If m_Selection < 0 Then m_Selection = 0
			
			With lvwChars.Items
				If .Count > m_Selection Then
					'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.Item(m_Selection + 1).Selected = True
					'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_ISSUE: MSComctlLib.ListView event lvwChars.ItemClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
					lvwChars_ItemClick(.Item(m_Selection + 1))
					'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: MSComctlLib.IListItem method lvwChars.ListItems.Item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					.Item(m_Selection + 1).EnsureVisible()
				End If
			End With
			
			If ds.MCPHandler.CharacterCount = 0 Then
				With lblRealm(0)
					.Text = "No characters found."
					.ForeColor = System.Drawing.Color.Red
				End With
			End If
		End If
	End Sub
	
	' form callback for character create
	Public Sub CharCreateResponse(ByVal Success As Boolean, ByVal Message As String)
		If Success Then
			' go back to character list and refresh
			HandleRefresh()
			
			' clear field for returning to this panel
			txtCharName.Text = vbNullString
		Else
			' clear saved name so we don't try to look for it later
			m_NewCharacterName = vbNullString
			
			' focus on the textbox since the character create failed
			With txtCharName
				.SelectionStart = 0
				.SelectionLength = Len(.Text)
				On Error Resume Next
				.Focus()
			End With
		End If
	End Sub
	
	' form callback for character delete
	Public Sub CharDeleteResponse(ByVal Success As Boolean, ByVal Message As String)
		If Success Then
			'If m_Selection >= 0 Then
			'    lvwChars.ListItems.Remove m_Selection + 1
			'Else
			' always re-request (or MCPHandler.CharacterList will be different)
			HandleRefresh()
			'End If
		End If
	End Sub
	
	' form callback for character upgrade
	Public Sub CharUpgradeResponse(ByVal Success As Boolean, ByVal Message As String)
		If Success Then
			'If m_Selection >= 0 Then
			'    lvwChars.ListItems.Item(m_Selection + 1).ForeColor = vbGreen
			'    Call lvwChars_ItemClick(lvwChars.ListItems.Item(m_Selection + 1))
			'Else
			' always re-request (or MCPHandler.CharacterList will be different)
			HandleRefresh()
			'End If
		End If
	End Sub
	
	' form callback for character logon
	Public Sub CharLogonResponse(ByVal Success As Boolean, ByVal Message As String)
		If Success Then
			UnloadNormal()
		End If
	End Sub
	
	' form callback for BNCS close/DoDisconnect
	Public Sub UnloadAfterBNCSClose()
		m_Unload_SuccessfulLogin = True
		ds.MCPHandler.IsRealmError = False
		
#If COMPILE_DEBUG = 0 Then
		UnhookWindowProc(Handle.ToInt32)
#End If
		
		Me.Close()
	End Sub
	
	Public Sub UnloadRealmError()
		If ds.MCPHandler Is Nothing Then
			ds.MCPHandler = New clsMCPHandler
		End If
		
		ds.MCPHandler.IsRealmError = True
		
#If COMPILE_DEBUG = 0 Then
		UnhookWindowProc(Handle.ToInt32)
#End If
		
		Me.Close()
	End Sub
	
	Private Sub UnloadNormal()
#If COMPILE_DEBUG = 0 Then
		UnhookWindowProc(Handle.ToInt32)
#End If
		
		Me.Close()
	End Sub
	
	Public Sub mnuPopDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopDelete.Click
		If (m_Selection >= 0 And optViewExisting.Checked) Then
			If lvwChars.Items.Count > m_Selection Then
				'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				Call HandleCharacterDelete(lvwChars.Items.Item(m_Selection + 1))
			End If
		End If
	End Sub
	
	Public Sub mnuPopUpgrade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopUpgrade.Click
		If (m_Selection >= 0) And CanUpgradeCharacter(m_Selection And optViewExisting.Checked) Then
			If lvwChars.Items.Count > m_Selection Then
				'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				Call HandleCharacterUpgrade(lvwChars.Items.Item(m_Selection + 1))
			End If
		End If
	End Sub
	
	Private Sub btnDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDelete.Click
		Call StopLoginTimer()
		
		If (Not lvwChars.FocusedItem Is Nothing And optViewExisting.Checked) Then
			Call HandleCharacterDelete(lvwChars.FocusedItem)
		End If
	End Sub
	
	Private Sub btnUpgrade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpgrade.Click
		Call StopLoginTimer()
		
		If (Not lvwChars.FocusedItem Is Nothing And optViewExisting.Checked) Then
			If CanUpgradeCharacter(lvwChars.FocusedItem.Tag) Then
				Call HandleCharacterUpgrade(lvwChars.FocusedItem)
			End If
		End If
	End Sub
	
	Private Sub HandleRefresh()
		lvwChars.Items.Clear()
		
		ds.MCPHandler.ClearInternalCharacters()
		
		optViewExisting.Checked = True
		Call CharListResponse()
		
		ds.MCPHandler.DoRequestCharacters()
	End Sub
	
	Private Sub HandleCharacterDelete(ByVal CharacterItem As System.Windows.Forms.ListViewItem)
		Dim Result As MsgBoxResult
		
		If Not CharacterItem Is Nothing And optViewExisting.Checked Then
			Result = MsgBox(CharacterItem.Text & " will be deleted. This action is irreversable! " & vbNewLine & "Are you sure you want to do that?", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Realm Confirm Delete")
			
			If Result = MsgBoxResult.Yes Then
				ds.MCPHandler.SEND_MCP_CHARDELETE(CharacterItem.Name)
			End If
		End If
	End Sub
	
	Private Sub HandleCharacterUpgrade(ByVal CharacterItem As System.Windows.Forms.ListViewItem)
		Dim Result As MsgBoxResult
		
		If Not CharacterItem Is Nothing And optViewExisting.Checked Then
			If CanUpgradeCharacter(CharacterItem.Tag) Then
				Result = MsgBox(CharacterItem.Text & " will be upgraded to a Lord of Destruction character. This action is irreversable! " & vbNewLine & "Are you sure you want to do that?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Realm Confirm Upgrade")
				
				If Result = MsgBoxResult.Yes Then
					ds.MCPHandler.SEND_MCP_CHARUPGRADE(CharacterItem.Name)
				End If
			End If
		End If
	End Sub
	
	Private Sub btnDisconnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDisconnect.Click
		Call StopLoginTimer()
		
		UnloadNormal()
	End Sub
	
	Private Sub btnChoose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnChoose.Click
		Call StopLoginTimer()
		
		With lvwChars
			If Not (.FocusedItem Is Nothing) Then
				If Not CanChooseCharacter(.FocusedItem.Tag) Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] You must use Diablo II: Lord of Destruction to choose that character.")
				Else
					Call ds.MCPHandler.SEND_MCP_CHARLOGON(.FocusedItem.Name)
					m_Unload_SuccessfulLogin = True
					ds.MCPHandler.IsRealmError = False
				End If
			End If
		End With
	End Sub
	
	'UPGRADE_WARNING: Event optNewCharType.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optNewCharType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNewCharType.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optNewCharType.GetIndex(eventSender)
			Dim i As Short
			
			'UPGRADE_WARNING: Lower bound of collection imlChars.ListImages has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			imgCharPortrait.Image = imlChars.Images.Item(Index + 1)
			
			For i = 1 To 7
				If i <> Index Then optNewCharType(i).Checked = False
			Next i
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkExpansion.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkExpansion_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkExpansion.CheckStateChanged
		Dim Enable As Boolean
		
		If Not m_IsExpansion Then
			chkExpansion.CheckState = System.Windows.Forms.CheckState.Unchecked
			Exit Sub
		End If
		
		Enable = (chkExpansion.CheckState <> 0)
		optNewCharType(6).Enabled = Enable
		optNewCharType(7).Enabled = Enable
		If Not Enable Then
			If optNewCharType(6).Checked Or optNewCharType(7).Checked Then
				optNewCharType(1).Checked = True
				Call optNewCharType_CheckedChanged(optNewCharType.Item(1), New System.EventArgs())
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optViewExisting.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optViewExisting_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optViewExisting.CheckedChanged
		If eventSender.Checked Then
			Call CharListResponse()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optCreateNew.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optCreateNew_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCreateNew.CheckedChanged
		If eventSender.Checked Then
			optViewExisting.Checked = False
			
			Call StopLoginTimer()
			
			fraCreateNew.Visible = True
			lvwChars.Visible = False
			With lblRealm(0) ' detail
				.Text = "Character creation."
				.ForeColor = System.Drawing.Color.Yellow
			End With
			' expires
			lblRealm(1).Text = vbNullString
			btnChoose.Enabled = False
			btnDelete.Visible = False
			btnUpgrade.Visible = False
		End If
	End Sub
	
	Sub StopLoginTimer()
		tmrLoginTimeout.Enabled = False
		' clear 3 parts of timer labels
		lblRealm(2).Text = vbNullString
		lblRealm(3).Text = vbNullString
		lblRealm(4).Text = vbNullString
	End Sub
	
	Private Sub tmrLoginTimeout_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrLoginTimeout.Tick
		Static indexValid As Short
		
		indexValid = m_Choice
		
		' if selecting nothing, find first unexpired account (no choose setting)
		Dim i As Short
		If (indexValid = -1) Then
			
			For i = 0 To ds.MCPHandler.CharacterCount - 1
				If Not IsDateExpired(ds.MCPHandler.CharacterExpires(i)) Or Not CanChooseCharacter(indexValid) Then
					indexValid = i
					Exit For
				End If
			Next i
			' if choose setting, then select only if not expired
		Else
			If IsDateExpired(ds.MCPHandler.CharacterExpires(indexValid)) Or Not CanChooseCharacter(indexValid) Then
				indexValid = -1
			End If
		End If
		
		If (indexValid >= 0) Then
			' warning label (part 1 of timer labels)
			'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			lblRealm(2).Text = lvwChars.Items.Item(indexValid + 1).Text & vbCrLf & " will be chosen automatically in"
			
			'If m_Selection < 0 Then
			'    lvwChars.ListItems.Item(indexValid + 1).Selected = True
			'    Call lvwChars_ItemClick(lvwChars.ListItems.Item(indexValid + 1))
			'End If
			
			If m_Ticks <= 1 Then
				tmrLoginTimeout.Enabled = False
				If Not CanChooseCharacter(indexValid) Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "[REALM] You must use Diablo II: Lord of Destruction to choose that character.")
				Else
					Call ds.MCPHandler.SEND_MCP_CHARLOGON(ds.MCPHandler.CharacterName(indexValid))
					m_Unload_SuccessfulLogin = True
					ds.MCPHandler.IsRealmError = False
				End If
			End If
		Else
			' warning label
			lblRealm(2).Text = "No unexpired characters found! Realm login will be cancelled in"
			
			If m_Ticks <= 1 Then
				tmrLoginTimeout.Enabled = False
				UnloadNormal()
			End If
		End If
		
		m_Ticks = m_Ticks - 1
		
		' seconds label (part 2 of timer labels)
		lblRealm(3).Text = CStr(m_Ticks)
		' seconds cap label (part 2 of timer labels)
		lblRealm(4).Text = "seconds."
	End Sub
	
	Private Function FindCharacter(ByVal sKey As String) As Short
		
		Dim i As Short
		
		With lvwChars.Items
			If .Count > 0 Then
				For i = 1 To .Count
					'UPGRADE_WARNING: Lower bound of collection lvwChars.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If .Item(i).Name = sKey Then
						FindCharacter = i - 1
						Exit Function
					End If
				Next i
			End If
		End With
		
		FindCharacter = -1
		
	End Function
	
	Private Function IsDateExpired(ByVal Expires As Date) As Boolean
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		IsDateExpired = (System.Math.Sign(DateDiff(Microsoft.VisualBasic.DateInterval.Second, UtcNow, Expires)) = -1)
	End Function
	
	Private Function CanUpgradeCharacter(ByVal CharIndex As Short) As Boolean
		Dim Stats As clsUserStats
		
		Stats = ds.MCPHandler.CharacterStats(CharIndex)
		
		CanUpgradeCharacter = (Not Stats.IsExpansionCharacter And m_IsExpansion)
	End Function
	
	Private Function CanChooseCharacter(ByVal CharIndex As Short) As Boolean
		Dim Stats As clsUserStats
		
		Stats = ds.MCPHandler.CharacterStats(CharIndex)
		
		' must be PX2D if isExpansion, otherwise doesn't matter
		CanChooseCharacter = (VB6.Imp(Stats.IsExpansionCharacter, m_IsExpansion))
	End Function
	
	Private Function GetCharacterExpireText(ByVal CharIndex As Short) As String
		Dim Expires As Date
		Dim ExpireType As String
		Dim ExpireDiff As Integer
		Dim ExpireVal As String
		Dim ExpireUnit As String
		
		Expires = ds.MCPHandler.CharacterExpires(CharIndex)
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		ExpireDiff = System.Math.Abs(DateDiff(Microsoft.VisualBasic.DateInterval.Hour, UtcNow, Expires))
		ExpireUnit = "hour"
		If ExpireDiff <> 1 Then ExpireUnit = ExpireUnit & "s"
		
		If ExpireDiff > 24 Then
			ExpireDiff = System.Math.Round(ExpireDiff / 24)
			ExpireUnit = "day"
			If ExpireDiff <> 1 Then ExpireUnit = ExpireUnit & "s"
		End If
		
		If IsDateExpired(Expires) Then
			ExpireType = "Expired"
			ExpireVal = CStr(ExpireDiff) & " " & ExpireUnit & " ago"
		Else
			ExpireType = "Expires"
			ExpireVal = "in " & CStr(ExpireDiff) & " " & ExpireUnit
		End If
		
		If ExpireDiff = 0 Then
			If IsDateExpired(Expires) Then
				ExpireVal = "minutes ago"
			Else
				ExpireType = "Expires"
				ExpireVal = "in minutes"
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCharacterExpireText = StringFormat("{0} on{1}{2}{1}({3})", ExpireType, vbCrLf, CStr(UtcToLocal(Expires)), ExpireVal)
	End Function
	
	Private Function GetCharacterDetailText(ByVal CharIndex As Short) As String
		Dim Stats As clsUserStats
		Dim NonExpansion As String
		Dim NonLadder As String
		Dim IsDead As String
		Dim NonHardcore As String
		
		Stats = ds.MCPHandler.CharacterStats(CharIndex)
		
		If Not Stats.IsLadderCharacter Then
			NonLadder = "non-"
		End If
		
		If Not Stats.IsHardcoreCharacter Then
			NonHardcore = "non-"
		ElseIf Stats.IsCharacterDead Then 
			IsDead = "dead "
		End If
		
		If Not Stats.IsExpansionCharacter Then
			NonExpansion = "non-"
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCharacterDetailText = StringFormat("{0} is a {1}ladder, {2}hardcore, {3}expansion {4} in {5}.", Stats.CharacterTitleAndName, NonLadder, NonHardcore, NonExpansion, Stats.CharacterClass, Stats.CurrentActAndDifficulty)
	End Function
End Class