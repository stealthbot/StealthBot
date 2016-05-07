Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmDBManager
	Inherits System.Windows.Forms.Form
	' frmDBManager.frm
	' Copyright (C) 2008 Eric Evans
	
	Public m_entrytype As String
	Public m_entryname As String
	
	' icons for groups list
	Private Const IC_EMPTY As Short = 0
	Private Const IC_UNCHECKED As Short = 2
	Private Const IC_CHECKED As Short = 3
	Private Const IC_PRIMARY As Short = 4
	
	' icons for database tree (for some reason it's 0-based!)
	Private Const IC_UNKNOWN As Short = 0
	Private Const IC_DATABASE As Short = 4
	Private Const IC_USER As Short = 5
	Private Const IC_GROUP As Short = 6
	Private Const IC_CLAN As Short = 7
	Private Const IC_GAME As Short = 8
	
	' temporary DB working copy (TODO: USE Collection OF clsDBEntryObj!!)
	Private m_DB() As udtDatabase
	' current entry index
	Private m_currententry As Short
	' current entry node
	Private m_currnode As vbalTreeViewLib6.cTreeViewNode
	' is this entry modified
	Private m_modified As Boolean
	' is this a new entry (unused: if label editing worked, we'd use this)
	Private m_new_entry As Boolean
	' root of DB ("Database")
	Private m_root As vbalTreeViewLib6.cTreeViewNode
	' target for user node right-click menu
	Private m_menutarget As vbalTreeViewLib6.cTreeViewNode
	' count of groups
	Private m_glistcount As Short
	' selected list item
	Private m_glistsel As System.Windows.Forms.ListViewItem
	' target for group list right-click menu
	Private m_gmenutarget As System.Windows.Forms.ListViewItem
	
	Private Sub frmDBManager_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		' this line is gay but for some reason I can't set the ImageList for vbalTV in the designer/VB properties -Ribose
		'UPGRADE_ISSUE: MSComctlLib.ImageList property Icons.hImageList was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        trvUsers.set_ImageList(icons)
		'TODO get images, because i broke the imagelist we had... :(
		
		' has our database been loaded?
		If (DB(0).Username = vbNullString) Then
			' load database, if for some reason that hasn't been done
			Call LoadDatabase()
		End If
		
		' store temporary copy of database
		m_DB = VB6.CopyArray(DB)
		
		' show database for default tab
		Call LoadView()
	End Sub ' end function Form_Load
	
	Public Sub ImportDatabase(ByRef strPath As String, ByRef dbType As Short)
		Dim f As Short
		Dim buf As String
		Dim n As vbalTreeViewLib6.cTreeViewNode
		Dim i As Short
		
		f = FreeFile
		
		Dim User As String
		Dim Msg As String
		If (dbType = 0) Then
			
			FileOpen(f, strPath, OpenMode.Input)
			Do While (EOF(f) = False)
				buf = LineInput(f)
				
				If (buf <> vbNullString) Then
					If (GetAccess(buf, "USER").Username <> vbNullString) Then
						For i = 0 To UBound(m_DB)
							If ((StrComp(m_DB(i).Username, buf, CompareMethod.Text) = 0) And (StrComp(m_DB(i).Type, "USER", CompareMethod.Text) = 0)) Then
								
								With m_DB(i)
									.Username = buf
									.Type = "USER"
									.ModifiedBy = "(console)"
									.ModifiedOn = Now
									
									If (InStr(1, .Flags, "S", CompareMethod.Binary) = 0) Then
										.Flags = .Flags & "S"
									End If
									
									If (Not (trvUsers.SelectedItem Is Nothing)) Then
										If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0) Then
											.Groups = .Groups & "," & trvUsers.SelectedItem.Text
										End If
									End If
									
									If (.Groups = vbNullString) Then
										.Groups = "%"
									End If
								End With
								
								Exit For
							End If
						Next i
					Else
						' redefine array to support new entry
						ReDim Preserve m_DB(UBound(m_DB) + 1)
						
						' create new database entry
						With m_DB(UBound(m_DB))
							.Username = buf
							.Type = "USER"
							.AddedBy = "(console)"
							.AddedOn = Now
							.ModifiedBy = "(console)"
							.ModifiedOn = Now
							.Flags = "S"
							
							If (Not (trvUsers.SelectedItem Is Nothing)) Then
								If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0) Then
									If (Not IsInGroup(m_DB(UBound(m_DB)).Groups, trvUsers.SelectedItem.Text)) Then
										.Groups = .Groups & "," & trvUsers.SelectedItem.Text
									End If
								End If
							End If
							
							If (.Groups = vbNullString) Then
								.Groups = "%"
							End If
						End With
					End If
				End If
			Loop  ' end loop
			FileClose(f)
			
		ElseIf ((dbType = 1) Or (dbType = 2)) Then 
			
			
			FileOpen(f, strPath, OpenMode.Input)
			Do While (EOF(f) = False)
				buf = LineInput(f)
				
				If (buf <> vbNullString) Then
					If (Not InStr(1, buf, Space(1), CompareMethod.Binary) = 0) Then
						User = VB.Left(buf, InStr(1, buf, Space(1), CompareMethod.Binary) - 1)
						
						Msg = Mid(buf, Len(User) + 1)
					Else
						User = buf
					End If
					
					If (GetAccess(User, "USER").Username <> vbNullString) Then
						For i = 0 To UBound(m_DB)
							If ((StrComp(m_DB(i).Username, User, CompareMethod.Text) = 0) And (StrComp(m_DB(i).Type, "USER", CompareMethod.Text) = 0)) Then
								
								With m_DB(i)
									.Username = User
									.Type = "USER"
									.ModifiedBy = "(console)"
									.ModifiedOn = Now
									
									If (InStr(1, .Flags, "B", CompareMethod.Binary) = 0) Then
										.Flags = .Flags & "B"
									End If
									
									If (Not (trvUsers.SelectedItem Is Nothing)) Then
										If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0) Then
											If (Not IsInGroup(m_DB(i).Groups, trvUsers.SelectedItem.Text)) Then
												.Groups = .Groups & "," & trvUsers.SelectedItem.Text
											End If
										End If
									End If
									
									If (.Groups = vbNullString) Then
										.Groups = "%"
									End If
								End With
							End If
						Next i
					Else
						' redefine array to support new entry
						ReDim Preserve m_DB(UBound(m_DB) + 1)
						
						' create new database entry
						With m_DB(UBound(m_DB))
							.Username = User
							.Type = "USER"
							.AddedBy = "(console)"
							.AddedOn = Now
							.ModifiedBy = "(console)"
							.ModifiedOn = Now
							.Flags = "B"
							.BanMessage = Msg
							
							If (Not (trvUsers.SelectedItem Is Nothing)) Then
								If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0) Then
									.Groups = trvUsers.SelectedItem.Text
								End If
							End If
							
							If (.Groups = vbNullString) Then
								.Groups = "%"
							End If
						End With
					End If
				End If
			Loop  ' end loop
			FileClose(f)
			
		End If
		
		LoadView()
	End Sub
	
	Private Sub btnCreateUser_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateUser.Click
		Dim userCount As Short
		Dim newNode As vbalTreeViewLib6.cTreeViewNode
		Dim gAcc As udtGetAccessResponse
		Dim Username As String
		
		m_entrytype = "USER"
		m_entryname = vbNullString
		
		Call VB6.ShowForm(frmDBNameEntry, VB6.FormShowConstants.Modal, Me)
		
        If (Len(m_entryname) > 0) Then

            Username = m_entryname

            If (GetAccess(Username, "USER").Username = vbNullString) Then
                ' redefine array to support new entry
                ReDim Preserve m_DB(UBound(m_DB) + 1)

                ' create new database entry
                With m_DB(UBound(m_DB))
                    .Username = Username
                    .Type = "USER"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With

                newNode = PlaceNewNode(Username, "USER", IC_USER)

                If (Not (newNode Is Nothing)) Then
                    ' change misc. settings
                    With newNode
                        '.Image = 0
                        .Tag = "USER"
                        .Selected = True
                    End With

                    'Call trvUsers_NodeClick(newNode)
                End If
            Else
                ' alert user that entry already exists
                MsgBox("There is already an entry of this type matching " & "the specified name.")
            End If
        End If
	End Sub
	
	Private Sub btnCreateGroup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateGroup.Click
		Dim newNode As vbalTreeViewLib6.cTreeViewNode
		Dim GroupName As String
		
		m_entrytype = "GROUP"
		m_entryname = vbNullString
		
		Call VB6.ShowForm(frmDBNameEntry, VB6.FormShowConstants.Modal, Me)
		
        If (Len(m_entryname) > 0) Then

            GroupName = m_entryname

            If (GetAccess(GroupName, "GROUP").Username = vbNullString) Then
                ReDim Preserve m_DB(UBound(m_DB) + 1)

                With m_DB(UBound(m_DB))
                    .Username = GroupName
                    .Type = "GROUP"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With

                newNode = PlaceNewNode(GroupName, "GROUP", IC_GROUP)

                Call UpdateGroupList()

                If (Not (newNode Is Nothing)) Then
                    ' change misc. settings
                    With newNode
                        .Tag = "GROUP"
                        .Selected = True
                    End With

                    'Call trvUsers_NodeClick(newNode)
                End If
            Else
                ' alert user that entry already exists
                MsgBox("There is already an entry of this type matching " & "the specified name.")
            End If
        End If
	End Sub
	
	Sub btnCreateClan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateClan.Click
		Dim newNode As vbalTreeViewLib6.cTreeViewNode
		Dim ClanName As String
		
		m_entrytype = "CLAN"
		m_entryname = vbNullString
		
		Call VB6.ShowForm(frmDBNameEntry, VB6.FormShowConstants.Modal, Me)
		
        If (Len(m_entryname) > 0) Then

            ClanName = m_entryname

            If (GetAccess(ClanName, "CLAN").Username = vbNullString) Then
                ReDim Preserve m_DB(UBound(m_DB) + 1)

                With m_DB(UBound(m_DB))
                    .Username = ClanName
                    .Type = "CLAN"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With

                newNode = PlaceNewNode(ClanName, "CLAN", IC_CLAN)

                If (Not (newNode Is Nothing)) Then
                    ' change misc. settings
                    With newNode
                        .Tag = "CLAN"
                        .Selected = True
                    End With

                    'Call trvUsers_NodeClick(newNode)
                End If
            Else
                ' alert user that entry already exists
                MsgBox("There is already an entry of this type matching " & "the specified name.")
            End If
        End If
	End Sub
	
	Sub btnCreateGame_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateGame.Click
		Dim newNode As vbalTreeViewLib6.cTreeViewNode
		Dim GameEntry As String
		
		m_entryname = vbNullString
		
		Call VB6.ShowForm(frmDBGameSelection, VB6.FormShowConstants.Modal, Me)
		
        If (Len(m_entryname) > 0) Then

            GameEntry = m_entryname

            If (GetAccess(GameEntry, "GAME").Username = vbNullString) Then
                ReDim Preserve m_DB(UBound(m_DB) + 1)

                With m_DB(UBound(m_DB))
                    .Username = GameEntry
                    .Type = "GAME"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With

                newNode = PlaceNewNode(GameEntry, "GAME", IC_GAME)

                If (Not (newNode Is Nothing)) Then
                    ' change misc. settings
                    With newNode
                        .Tag = "GAME"
                        .Selected = True
                    End With

                    'Call trvUsers_NodeClick(newNode)
                End If
            Else
                ' alert user that entry already exists
                MsgBox("There is already an entry of this type matching " & "the specified name.")
            End If
        End If
	End Sub
	
	Private Function PlaceNewNode(ByRef EntryName As String, ByRef EntryType As String, ByRef EntryImage As Short) As vbalTreeViewLib6.cTreeViewNode
		Dim NewParent As vbalTreeViewLib6.cTreeViewNode
		
		' by default create the node under the root node
		NewParent = m_root
		
		' do we have an item (hopefully a group) selected?
		If (Not (trvUsers.SelectedItem Is Nothing)) Then
			' is the item a group?
			If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0) Then
				' create new node under group node
				NewParent = trvUsers.SelectedItem
			Else
				' is our parent a group?
				If Not trvUsers.SelectedItem.Parent Is Nothing Then
					If (StrComp(trvUsers.SelectedItem.Parent.Tag, "GROUP", CompareMethod.Text) = 0) Then
						NewParent = trvUsers.SelectedItem.Parent
					End If
				End If
			End If
		End If
		
		PlaceNewNode = trvUsers.nodes.Add(NewParent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, EntryType & ": " & EntryName, EntryName, EntryImage, EntryImage)
		
		' set group settings on new database entry
		If (Not PlaceNewNode Is Nothing) Then
			If (StrComp(PlaceNewNode.Parent.Tag, "GROUP", CompareMethod.Text) = 0) Then
				With m_DB(UBound(m_DB))
					.Groups = PlaceNewNode.Parent.Text
				End With
			End If
		End If
	End Function
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		m_modified = False
		
		Me.Close()
	End Sub
	
	Private Sub btnSaveUser_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSaveUser.Click
		Dim i As Short
		Dim j As Short
		Dim OldGroups As String
		Dim NewGroups As String
		Dim OldPGroup As String
		Dim NewPGroup As String
		Dim pos As Short
		Dim NewParent As vbalTreeViewLib6.cTreeViewNode
		Dim node As vbalTreeViewLib6.cTreeViewNode
		
		' if we have no selected user... escape quick!
		If (trvUsers.SelectedItem Is Nothing) Then
			' break from function
			Exit Sub
		End If
		
		' can't "save" the "Database"/root node
		If (StrComp(trvUsers.SelectedItem.Tag, "DATABASE", CompareMethod.Text) = 0) Then
			Exit Sub
		End If
		
		' disable entry save command
		Call HandleSaved()
		
		' look for selected user in database
		For i = LBound(m_DB) To UBound(m_DB)
			' is this the user we were looking for?
			If (StrComp(trvUsers.SelectedItem.Text, m_DB(i).Username, CompareMethod.Text) = 0) Then
				If (StrComp(trvUsers.SelectedItem.Tag, m_DB(i).Type, CompareMethod.Text) = 0) Then
					' modifiy user data
					With m_DB(i)
						.Rank = Val(txtRank.Text)
						.Flags = txtFlags.Text
						.ModifiedBy = "(console)"
						.ModifiedOn = Now
						.BanMessage = txtBanMessage.Text
						
						' save old groups...
						OldGroups = .Groups
						
						' generate new groups string
						NewGroups = vbNullString
						
						If m_glistcount > 0 Then
							For j = 1 To lvGroups.Items.Count
								'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								With lvGroups.Items.Item(j)
									'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                    If .Checked And Not False Then
                                        'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                                        If .ForeColor.Equals(System.Drawing.Color.Yellow) Then
                                            ' place first
                                            'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                                            NewGroups = .Text & "," & NewGroups
                                        Else
                                            ' append
                                            'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                                            NewGroups = NewGroups & .Text & ","
                                        End If
                                    End If
								End With
							Next j
						End If
						
						' if ends with ",", trim it
						If (Len(NewGroups) > 1) Then
							NewGroups = VB.Left(NewGroups, Len(NewGroups) - 1)
						End If
						
						' store it
						.Groups = NewGroups
						If .Groups = vbNullString Then
							.Groups = "%"
						End If
						
						' now to check if we need to move this node!
						' did the "primary" group change?
						OldPGroup = GetPrimaryGroup(OldGroups)
						NewPGroup = GetPrimaryGroup(NewGroups)
						
						If (StrComp(OldPGroup, NewPGroup, CompareMethod.Text) <> 0) Then
							' move under new primary
							pos = FindNodeIndex(NewPGroup, "GROUP")
							' well, does it exist?
							If (pos > 0) Then
								' make node a child of existing group
								NewParent = trvUsers.nodes(pos)
							Else
								' put it under DB root
								NewParent = m_root
							End If
							
							' move node!!
							node = trvUsers.SelectedItem
							Call node.MoveNode(NewParent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild)
							node.Tag = .Type
							node.Selected = True
						End If
					End With
					
					Exit For
				End If
			End If
		Next i
	End Sub
	
	Private Sub HandleSaved()
		m_modified = False
		btnSaveUser.Enabled = False
		If m_currententry < 0 Then
			frmDatabase.Text = "Database"
			Me.Text = "Database"
		Else
			frmDatabase.Text = m_DB(m_currententry).Username & " (" & LCase(m_DB(m_currententry).Type) & ")"
			Me.Text = "Database - " & m_DB(m_currententry).Username & " (" & LCase(m_DB(m_currententry).Type) & ")"
		End If
	End Sub
	
	Private Sub HandleUnsaved()
		If m_currententry < 0 Then
			m_modified = False
			frmDatabase.Text = "Database"
			Me.Text = "Database"
		Else
			m_modified = True
			btnSaveUser.Enabled = True
			frmDatabase.Text = m_DB(m_currententry).Username & " (" & m_DB(m_currententry).Type & ")*"
			Me.Text = "Database - " & m_DB(m_currententry).Username & " (" & m_DB(m_currententry).Type & ")*"
		End If
	End Sub
	
	Private Sub btnSaveForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSaveForm.Click
		' save this user first
		If m_modified Then
			btnSaveUser_Click(btnSaveUser, New System.EventArgs())
		End If
		
		' write temporary database to official
		DB = VB6.CopyArray(m_DB)
		
		' save database
		Call WriteDatabase(GetFilePath(FILE_USERDB))
		
		' check channel to find potential banned users
		Call g_Channel.CheckUsers()
		
		' close database form
		Call Me.Close()
	End Sub
	
	Private Sub lvGroups_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvGroups.Click
		m_glistsel = lvGroups.FocusedItem
		
		If Not m_glistsel Is Nothing Then
			m_glistsel.Checked = Not m_glistsel.Checked
			Call lvGroups_ItemCheck(lvGroups, New System.Windows.Forms.ItemCheckEventArgs(m_glistsel.Index, System.Windows.Forms.CheckState.Indeterminate, m_glistsel.Checked))
		End If
	End Sub
	
	Private Sub lvGroups_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvGroups.DoubleClick
		Dim i As Short
		Dim Item As System.Windows.Forms.ListViewItem
		
		Item = lvGroups.FocusedItem
		
		'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        If False Then
            Exit Sub
        End If
		
		' set primary
		If Item.Checked Then
			Call SetLVPrimaryGroup(Item)
		End If
		
		' enable entry save command
		Call HandleUnsaved()
	End Sub
	
	Private Sub lvGroups_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lvGroups.ItemCheck
		Dim Item As System.Windows.Forms.ListViewItem = lvGroups.Items(eventArgs.Index)
		Dim i As Short
		Dim NewGroups As String
		
		Item.Selected = True
		
		'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        If False Then
            Item.Checked = False
            'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            Item.ImageIndex = IIf(m_glistcount = 0, IC_EMPTY, IC_UNCHECKED)
            Exit Sub
        End If
		
		If Item.Checked Then
			' if checked
			' if no primary group
			If GetLVPrimaryGroup() Is Nothing Then
				' set this
				Item.ForeColor = System.Drawing.Color.Yellow
				'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                Item.ImageIndex = IC_PRIMARY
			Else
				'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                Item.ImageIndex = IC_CHECKED
			End If
		Else
			' if not checked
			' select the first "checked" item to be new primary
			If Item.ForeColor.equals(System.Drawing.Color.Yellow) Then
				' unset this
				Item.ForeColor = System.Drawing.Color.White
				For i = 1 To lvGroups.Items.Count
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If lvGroups.Items.Item(i).Checked Then
						' set if found
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						With lvGroups.Items.Item(i)
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							.ForeColor = System.Drawing.Color.Yellow
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            .ImageIndex = IC_PRIMARY
						End With
						
						Exit For
					End If
				Next i
			End If
			'UPGRADE_ISSUE: MSComctlLib.ListItem property Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            Item.ImageIndex = IC_UNCHECKED
		End If
		
		' generate new groups string
		NewGroups = vbNullString
		
		For i = 1 To lvGroups.Items.Count
			'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			With lvGroups.Items.Item(i)
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                If .Checked And Not False Then
                    'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .ForeColor.Equals(System.Drawing.Color.Yellow) Then
                        ' place first
                        'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        NewGroups = .Text & "," & NewGroups
                    Else
                        ' append
                        'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        NewGroups = NewGroups & .Text & ","
                    End If
                End If
			End With
		Next i
		
		' if ends with ",", trim it
		If (Len(NewGroups) > 1) Then
			NewGroups = VB.Left(NewGroups, Len(NewGroups) - 1)
		End If
		
		' update inherits list
		Call UpdateInheritCaption(NewGroups)
		
		' enable entry save command
		Call HandleUnsaved()
	End Sub
	
	Private Sub lvGroups_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvGroups.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If m_glistsel.Text <> lvGroups.FocusedItem.Text Then
			m_glistsel = lvGroups.FocusedItem
		End If
		
		If KeyAscii = System.Windows.Forms.Keys.Space Then
			If Not m_glistsel Is Nothing Then
				m_glistsel.Checked = Not m_glistsel.Checked
				Call lvGroups_ItemCheck(lvGroups, New System.Windows.Forms.ItemCheckEventArgs(m_glistsel.Index, System.Windows.Forms.CheckState.Indeterminate, m_glistsel.Checked))
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub lvGroups_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvGroups.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		mnuSetPrimary.Visible = True
		mnuRename.Visible = False
		mnuDelete.Visible = False
		
		mnuSetPrimary.Enabled = False
		
		'UPGRADE_NOTE: Object m_gmenutarget may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_gmenutarget = Nothing
		
		If (Button = VB6.MouseButtonConstants.RightButton) Then
			m_gmenutarget = lvGroups.GetItemAt(x, y)
			
			If (Not m_gmenutarget Is Nothing) Then
				'UPGRADE_ISSUE: MSComctlLib.ListItem property m_gmenutarget.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                If (Not False) Then
                    mnuSetPrimary.Enabled = (Not m_gmenutarget.ForeColor.Equals(System.Drawing.Color.Yellow))

                    mnuContext.ShowDropDown()
                End If
			End If
		End If
	End Sub
	
	Public Sub mnuSetPrimary_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSetPrimary.Click
		Call SetLVPrimaryGroup(m_gmenutarget)
		
		Call HandleUnsaved()
	End Sub
	
	Public Sub mnuDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelete.Click
		If (Not (trvUsers.SelectedItem Is Nothing)) Then
			Call HandleDeleteEvent(m_menutarget)
		End If
	End Sub
	
	Private Sub btnDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDelete.Click
		If (Not (trvUsers.SelectedItem Is Nothing)) Then
			Call HandleDeleteEvent(trvUsers.SelectedItem)
		End If
	End Sub
	
	Public Sub mnuRename_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRename.Click
		Call QueryRenameEvent(m_menutarget)
	End Sub
	
	Private Sub btnRename_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRename.Click
		Call QueryRenameEvent(trvUsers.SelectedItem)
	End Sub
	
	'Private Sub mnuOpenDatabase_Click()
	'    ' open file dialog
	'    Call CommonDialog.ShowOpen
	'End Sub
	
	Private Sub HandleDeleteEvent(ByRef Target As vbalTreeViewLib6.cTreeViewNode)
		If (Target Is Nothing) Then
			Exit Sub
		End If
		
		If (StrComp(Target.Tag, "DATABASE", CompareMethod.Text) = 0) Then
			Exit Sub
		End If
		
		Dim response As MsgBoxResult
		Dim isGroup As Boolean
		
		If (m_modified And StrComp(Target.Text, m_currnode.Text, CompareMethod.Binary) = 0) Then
			' do not ask about "unsaved data" later!
			m_modified = False
		End If
		
		isGroup = (StrComp(Target.Tag, "GROUP", CompareMethod.Text) = 0)
		
		If (isGroup) Then
			response = MsgBox("Are you sure you wish to delete this group and " & "all of its members?", MsgBoxStyle.YesNo Or MsgBoxStyle.Information, "Database - Confirm Delete")
		End If
		
		If ((isGroup = False) Or ((isGroup) And (response = MsgBoxResult.Yes))) Then
			Call DB_remove(Target.Text, (Target.Tag))
			
			Target.Delete()
			
			If isGroup Then
				Call UpdateGroupList()
			End If
			
			'Call trvUsers_NodeClick(trvUsers.SelectedItem)
		End If
	End Sub
	
	Private Sub QueryRenameEvent(ByRef Target As vbalTreeViewLib6.cTreeViewNode)
		If (Not (Target Is Nothing)) Then
			' only "GROUP" entries can be renamed
			If (StrComp(Target.Tag, "GROUP", CompareMethod.Text) = 0) Then
				'trvUsers.SelectedItem.StartEdit
				
				m_entrytype = Target.Tag
				m_entryname = Target.Text
				
				Call VB6.ShowForm(frmDBNameEntry, VB6.FormShowConstants.Modal, Me)
				
				If HandleRenameEvent(Target, m_entryname) Then
					
                    If Len(m_entryname) > 0 Then
                        Target.Text = m_entryname
                        trvUsers.Refresh()

                        Call UpdateGroupList()

                        If m_modified Then
                            Call HandleUnsaved()
                        Else
                            Call HandleSaved()
                        End If
                    End If
					
				Else
					' alert user that entry already exists
					MsgBox("There is already an entry of this type matching " & "the specified name.")
				End If
			End If
		End If
	End Sub
	
	Private Function HandleRenameEvent(ByRef Target As vbalTreeViewLib6.cTreeViewNode, ByRef NewString As String) As Boolean
		Dim i As Short
		Dim WasUpdated As Boolean
		
		HandleRenameEvent = True
		WasUpdated = False
		
		If (Target Is Nothing) Then
			Exit Function
		End If
		
		If (StrComp(Target.Tag, "DATABASE", CompareMethod.Text) = 0) Then
			Exit Function
		End If
		
        If NewString = Target.Text Or Len(NewString) = 0 Then
            ' same name succeeds (no chnage); empty name success (cancelled)
            Exit Function
        End If
		
		For i = LBound(m_DB) To UBound(m_DB)
			If (StrComp(Target.Text, m_DB(i).Username, CompareMethod.Text) = 0) Then
				If (StrComp(Target.Tag, m_DB(i).Type, CompareMethod.Text) = 0) Then
                    If Len(GetAccess(NewString, m_DB(i).Type).Username) > 0 Then
                        ' already exists
                        HandleRenameEvent = False
                        Exit Function
                    End If
					
					' rename DB entry
					m_DB(i).Username = NewString
					WasUpdated = True
					
					Exit For
				End If
			End If
		Next i
		
		If Not WasUpdated Then
			' this didn't already exist!! shouldn't happen...
			HandleRenameEvent = False
			Exit Function
		End If
		
		Dim Splt() As String
		Dim j As Short
		If (StrComp(m_DB(i).Type, "GROUP", CompareMethod.Text) = 0) Then
			For i = LBound(m_DB) To UBound(m_DB)
				If (Len(m_DB(i).Groups) > 0) Then
					
					If (Not InStr(1, m_DB(i).Groups, ",", CompareMethod.Text) = 0) Then
						Splt = Split(m_DB(i).Groups, ",")
					Else
						ReDim Preserve Splt(0)
						
						Splt(0) = m_DB(i).Groups
					End If
					
					For j = LBound(Splt) To UBound(Splt)
						If (StrComp(Splt(j), "%", CompareMethod.Binary) <> 0) And (StrComp(Splt(j), Target.Text, CompareMethod.Text) = 0) Then
							Splt(j) = NewString
						End If
					Next j
					
					m_DB(i).Groups = Join(Splt, ",")
				End If
			Next i
		End If
	End Function
	
	' handle tab clicks and initial loading
	Private Sub LoadView()
		On Error GoTo ERROR_HANDLER
		
		Dim newNode As vbalTreeViewLib6.cTreeViewNode
		
		Dim i As Short
		Dim grp As String
		Dim j As Short
		Dim pos As Short
		Dim blnDuplicateFound As Boolean
		'UPGRADE_NOTE: TypeName was upgraded to TypeName_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim TypeName_Renamed As String
		Dim TypeImage As Short
		Dim NewParent As vbalTreeViewLib6.cTreeViewNode
		
		' clear treeview
		Call trvUsers.nodes.Clear()
		
		' create root node
		m_root = trvUsers.nodes.Add( ,  , "Database", "Database", IC_DATABASE, IC_DATABASE)
		' type DATABASE
		m_root.Tag = "DATABASE"
		
		' which tab index are we on?
		Dim K As Short
		Dim bln As Boolean
		For i = LBound(m_DB) To UBound(m_DB)
			' we're handling groups first; is this entry a group?
			If (StrComp(m_DB(i).Type, "GROUP", CompareMethod.Binary) = 0) Then
				' is this group a member of other groups?
				If (Len(m_DB(i).Groups) > 0) And (StrComp(m_DB(i).Groups, "%", CompareMethod.Binary) <> 0) Then
					' get the "primary" group (the first group) to put the node under
					grp = GetPrimaryGroup(m_DB(i).Groups)
					
					' has the group already been added or is database in an
					' incorrect order?
					pos = FindNodeIndex(grp, "GROUP")
					' well, does it exist?
					If (pos > 0) Then
						' make node a child of existing group
						NewParent = trvUsers.nodes(pos)
					Else
						' lets make this guy a parent node for now until we can find
						' his real parent.
						NewParent = m_root
					End If
					
					newNode = trvUsers.nodes.Add(NewParent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "GROUP: " & m_DB(i).Username, m_DB(i).Username, IC_GROUP, IC_GROUP)
				Else
					
					' create node
					NewParent = m_root
					
					newNode = trvUsers.nodes.Add(NewParent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, "GROUP: " & m_DB(i).Username, m_DB(i).Username, IC_GROUP, IC_GROUP)
					
					' Okay, is the group a lone ranger?  Or does he have children
					' that are already in the list?
					j = LBound(m_DB)
					Do 
						For j = j To (i - 1)
							' we're only concerned with groups, atm.
							If (StrComp(m_DB(j).Type, "GROUP", CompareMethod.Binary) = 0) Then
								' we only need to check for groups that are members of
								' other groups
								If (Len(m_DB(j).Groups) > 0) And (StrComp(m_DB(j).Groups, "%", CompareMethod.Binary) <> 0) Then
									' is entry member of multiple groups?
									If (InStr(1, m_DB(j).Groups, ",", CompareMethod.Binary) <> 0) Then
										' split up multiple groupings
										grp = Split(m_DB(j).Groups, ",", 2)(0)
									Else
										' no need for special handling...
										grp = m_DB(j).Groups
									End If
									
									' is the current group a member of our group?
									If (StrComp(grp, m_DB(i).Username, CompareMethod.Text) = 0) Then
										' indicate that we've found a match
										bln = True
										
										' break from loop
										Exit For
									End If
								End If
							End If
						Next j
						
						' is this node a baby's daddy?
						If (bln) Then
							' move node
							pos = FindNodeIndex(m_DB(j).Username, "GROUP")
							
                            If pos > 0 Then
                                trvUsers.Nodes(pos).MoveNode(newNode, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild)
                            End If
						End If
						
						' reset boolean
						bln = False
					Loop Until j = i
				End If
				
				If (Not (newNode Is Nothing)) Then
					' change misc. settings
					newNode.Tag = "GROUP"
				End If
			End If
		Next i
		
		' loop through database... this time looking for users, clans, and games (clans and games are "user like" in the tree)
		For i = LBound(m_DB) To UBound(m_DB)
			' is the entry a user?
			If (StrComp(m_DB(i).Type, "GROUP", CompareMethod.Text) <> 0) Then
				' find the type name, used for the treeview
				If (StrComp(m_DB(i).Type, "USER", CompareMethod.Text) = 0) Then
					TypeName_Renamed = "USER"
					TypeImage = IC_USER
				ElseIf (StrComp(m_DB(i).Type, "CLAN", CompareMethod.Text) = 0) Then 
					TypeName_Renamed = "CLAN"
					TypeImage = IC_CLAN
				ElseIf (StrComp(m_DB(i).Type, "GAME", CompareMethod.Text) = 0) Then 
					TypeName_Renamed = "GAME"
					TypeImage = IC_GAME
				Else
					TypeName_Renamed = "USER"
					TypeImage = IC_UNKNOWN
				End If
				
				' is the user a member of any groups?
				If (Len(m_DB(i).Groups) > 0) And (StrComp(m_DB(i).Groups, "%", CompareMethod.Binary) <> 0) Then
					' get the "primary" group (the first group) to put the node under
					grp = GetPrimaryGroup(m_DB(i).Groups)
					
					If (grp = vbNullString) Then
						pos = False
					Else
						' search for our group
						pos = FindNodeIndex(grp, "GROUP")
						
						' does our group exist?
						If (pos > 0) Then
							NewParent = trvUsers.nodes(pos)
						End If
					End If
				End If
				
				If (pos <= 0) Then
					' create new user node under root
					NewParent = m_root
				End If
				
				' create user node and move into group
				newNode = trvUsers.nodes.Add(NewParent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, TypeName_Renamed & ": " & m_DB(i).Username, m_DB(i).Username, TypeImage, TypeImage)
				
				If (Not (newNode Is Nothing)) Then
					' change misc. settings
					newNode.Tag = TypeName_Renamed
				End If
			End If
			
			' reset our variables
			pos = 0
		Next i
		
		' does our treeview contain any nodes?
		For i = 1 To trvUsers.NodeCount
			trvUsers.nodes(i).Expanded = True
		Next i
		
		Call UpdateGroupList()
		
		If trvUsers.NodeCount = 0 Then
			Call LockGUI()
			'Else
			'    Call trvUsers_NodeClick(trvUsers.nodes(1))
		End If
		
		If (blnDuplicateFound = True) Then
			MsgBox("There were one or more duplicate database entries found which could not be loaded.", MsgBoxStyle.Exclamation, "Error")
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		If (Err.Number = 35602) Then
			DB_remove(m_DB(i).Username, m_DB(i).Type)
			blnDuplicateFound = True
			Resume Next
		End If
		
		Exit Sub
	End Sub
	
	Private Sub LockGUI()
		Dim i As Short
		
		' set our default frame caption
		m_currententry = -1
		Call HandleSaved()
		
		' disable & clear rank
		txtRank.Enabled = False
		txtRank.Text = vbNullString
		
		' disable & clear flags
		txtFlags.Enabled = False
		txtFlags.Text = vbNullString
		
		' loop through listbox and clear selected items
		Call ClearGroupListChecks()
		
		' disable group lists
		'lvGroups.Enabled = False
		
		' disable & clear ban message
		txtBanMessage.Enabled = False
		txtBanMessage.Text = vbNullString
		
		' reset created on & modified on labels
		lblCreatedOn.Text = "(not applicable)"
		lblModifiedOn.Text = "(not applicable)"
		
		' reset created by & modified by labels
		lblCreatedBy.Text = vbNullString
		lblModifiedBy.Text = vbNullString
		
		' reset inherits caption
		lblInherit.Text = vbNullString
		
		' disable entry buttons
		btnRename.Enabled = False
		btnDelete.Enabled = False
	End Sub
	
	Private Sub UnlockGUI()
		Dim i As Short
		
		' enable rank field
		txtRank.Enabled = True
		
		' enable flags field
		txtFlags.Enabled = True
		
		' enable ban message field
		txtBanMessage.Enabled = True
		
		' enable entry rename/delete buttons
		btnRename.Enabled = (StrComp(trvUsers.SelectedItem.Tag, "GROUP", CompareMethod.Text) = 0)
		btnDelete.Enabled = True
		
		' enable group lists
		'lvGroups.Enabled = True
		
		' make sure save button and caption is up to date
		HandleSaved()
	End Sub
	
	' handle node collapse
	Private Sub trvUsers_Collapse(ByVal eventSender As System.Object, ByVal eventArgs As AxvbalTreeViewLib6.__vbalTreeView_CollapseEvent) Handles trvUsers.Collapse
		' refresh tree view
        Call trvUsers.Refresh()
	End Sub
	
	' handle node expand
	Private Sub trvUsers_Expand(ByVal eventSender As System.Object, ByVal eventArgs As AxvbalTreeViewLib6.__vbalTreeView_ExpandEvent) Handles trvUsers.Expand
		' refresh tree view
        Call trvUsers.Refresh()
	End Sub
	
	Private Sub trvUsers_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxvbalTreeViewLib6.__vbalTreeView_KeyDownEvent) Handles trvUsers.KeyDownEvent
		If (eventArgs.KeyCode = System.Windows.Forms.Keys.Delete) Then
			If (Not (trvUsers.SelectedItem Is Nothing)) Then
				Call HandleDeleteEvent(trvUsers.SelectedItem)
			End If
		ElseIf (eventArgs.KeyCode = System.Windows.Forms.Keys.F2) Then 
			Call QueryRenameEvent(trvUsers.SelectedItem)
		End If
	End Sub
	
	'Private Sub trvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
	'    Dim node As cTreeViewNode
	'
	'    If (Button = vbLeftButton) Then
	'        Set m_nodedragsrc = trvUsers.HitTest(x, y)
	'        'frmChat.AddChat vbYellow, "[MOUSE] DOWN DRAG=" & m_nodedragsrc.Text
	'    End If
	'End Sub
	'
	'Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
	'    'frmChat.AddChat vbYellow, "[MOUSE] MOVE" ' DRAG=" & m_nodedragsrc.Text
	'    If (Button = vbLeftButton And m_nodedrag And Not m_nodedragsrc Is Nothing) Then
	'        trvUsers_OLEDragDrop Nothing, vbDropEffectMove, Button, Shift, x, y 'm_dragtarget
	'    End If
	'End Sub
	
	'Private Sub trvUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
	'    'frmChat.AddChat vbYellow, "[MOUSE] UP" ' DRAG=" & m_nodedragsrc.Text
	'End Sub
	
	Private Sub trvUsers_NodeRightClick(ByVal eventSender As System.Object, ByVal eventArgs As AxvbalTreeViewLib6.__vbalTreeView_NodeRightClickEvent) Handles trvUsers.NodeRightClick
		mnuRename.Visible = True
		mnuDelete.Visible = True
		mnuSetPrimary.Visible = False
		
		mnuRename.Enabled = False
		mnuDelete.Enabled = False
		
		'UPGRADE_NOTE: Object m_menutarget may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_menutarget = Nothing
		
		If (Not (eventArgs.node Is Nothing)) Then
			If (StrComp(eventArgs.node.Tag, "DATABASE", CompareMethod.Text) <> 0) Then
				mnuRename.Enabled = (StrComp(eventArgs.node.Tag, "GROUP", CompareMethod.Text) = 0)
				mnuDelete.Enabled = True
				
				m_menutarget = eventArgs.node
			End If
		End If

        mnuContext.ShowDropDown()
	End Sub
	
	'Private Sub trvUsers_OLECompleteDrag(Effect As Long)
	'    'frmChat.AddChat vbYellow, "[DRAG] COMPLETE: E=" & Effect
	'End Sub
	'
	''// occurs when the user drops the object
	''// this is where you move the node and its children.
	''// this will not occur if Effect = vbDropEffectNone
	'Private Sub trvUsers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
	'    Dim strSourceKey As String
	'    Dim dragNode   As cTreeViewNode
	'    Dim dropNode   As cTreeViewNode
	'
	'    '// get the carried data
	'    'strSourceKey = Data.GetData(vbCFText)
	'    Set dragNode = m_nodedragsrc
	'    If dragNode Is Nothing Then
	'        Exit Sub
	'    End If
	'    Set dropNode = trvUsers.HitTest(x, y)
	'    If dropNode Is Nothing Then
	'        Set dropNode = m_root
	'    End If
	'    '// get the target node
	'    frmChat.AddChat vbYellow, "[DRAG] DROP: DROP=" & dropNode.Text & " E=" & Effect
	'    '// if the target node is not a folder or the root item
	'    '// then get it's parent (that is a folder or the root item)
	'    If (StrComp(dropNode.Tag, "GROUP", vbTextCompare) <> 0) And (StrComp(dropNode.Tag, "DATABASE", vbTextCompare) <> 0) Then
	'        '// the target must be a GROUP or the DATABASE
	'        Effect = vbDropEffectNone
	'        Exit Sub
	'    End If
	'
	'    'Set dragNode = trvUsers.nodes(strSourceKey)
	'    If Not dragNode Is Nothing Then
	'        frmChat.AddChat vbYellow, "[DRAG] DROP: DRAG=" & dragNode.Text & " E=" & Effect
	'
	'        '// move the source node to the target node
	'        Call dragNode.MoveNode(dropNode, etvwChild)
	'        dragNode.Tag = "USER"
	'        dragNode.Selected = True
	'    End If
	'    '// NOTE: You will also need to update the key to reflect the changes
	'    '// if you are using it
	'    '// we are not dragging from this control any more
	'    m_nodedrag = False
	'    '// cancel effect so that VB doesn't muck up your transfer
	'    Effect = 0
	'End Sub
	'
	''// occurs when the user starts dragging
	''// this is where you assign the effect and the data.
	'Private Sub trvUsers_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
	'    'Dim K As String
	'    AllowedEffects = vbDropEffectNone
	'
	'    If (Not m_nodedragsrc Is Nothing) Then
	'        If (StrComp(m_nodedragsrc.Tag, "DATABASE", vbTextCompare) <> 0) Then
	'            '// Set the effect to move
	'            AllowedEffects = vbDropEffectMove
	'            '// Assign the selected item's key to the DataObject
	'            'K = m_nodedragsrc.Key
	'            'Call Data.SetData(m_nodedragsrc.Key)
	'            '// we are dragging from this control
	'            m_nodedrag = True
	'        End If
	'    End If
	'
	'    'frmChat.AddChat vbYellow, "[DRAG] START: AE=" & AllowedEffects & ", D=" & K
	'End Sub
	'
	''// occurs when the object is dragged over the control.
	''// this is where you check to see if the mouse is over
	''// a valid drop object
	'Private Sub trvUsers_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
	'    Dim dropNode As cTreeViewNode
	'
	'    '// set the effect
	'    Effect = vbDropEffectMove
	'    '// get the node that the object is being dragged over
	'    Set dropNode = trvUsers.HitTest(x, y)
	'
	'    'Set m_nodedragdata = Data
	'
	'    If (Not dropNode Is Nothing And m_nodedrag) Then
	'        If (StrComp(dropNode.Tag, "GROUP", vbTextCompare) = 0) Or (StrComp(dropNode.Tag, "DATABASE", vbTextCompare) = 0) Then
	'            'dropNode.DropHighlighted = True
	'        Else
	'            '// the target must be a GROUP or the DATABASE
	'            Effect = vbDropEffectNone
	'            Exit Sub
	'        End If
	'    Else
	'        'm_root.DropHighlighted = True
	'        Exit Sub
	'    End If
	'
	'    frmChat.AddChat vbYellow, "[DRAG] OVER: DROP=" & dropNode.Text & " E=" & Effect & " S=" & State
	'End Sub
	
	Private Sub trvUsers_SelectedNodeChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trvUsers.SelectedNodeChanged
		Static skipupdate As Boolean
		Dim node As vbalTreeViewLib6.cTreeViewNode
		Dim tmp As udtGetAccessResponse
		Dim i As Short
		Dim Splt() As String
		Dim j As Short
		Dim response As MsgBoxResult
		
		node = trvUsers.SelectedItem
		
		If (node Is Nothing) Then
			Exit Sub
		End If
		
		If m_modified And Not skipupdate Then
			' check if we should allow this node change? (is Unsaved?)
			response = MsgBox("Are you sure you wish to discard changes to the " & m_currnode.Text & " (" & UCase(m_currnode.Tag) & ") database entry?", MsgBoxStyle.YesNo Or MsgBoxStyle.Information, "Database - Confirm Discard Changes")
			
			If response = MsgBoxResult.No Then
				skipupdate = True
				m_currnode.Selected = True
				skipupdate = False
				Exit Sub
			End If
		End If
		
		If skipupdate Then Exit Sub
		
		m_currnode = node
		
		Call LockGUI()
		
		node.Expanded = True
		
		If (StrComp(node.Tag, "DATABASE", CompareMethod.Text) = 0) Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tmp = GetAccess(node.Text, (node.Tag), m_currententry)
		
		' does entry have a rank?
		If (tmp.Rank > 0) Then
			' write rank to text box
			txtRank.Text = CStr(tmp.Rank)
		Else
			' clear rank from text box
			txtRank.Text = vbNullString
		End If
		
		' clear flags from text box
		txtFlags.Text = tmp.Flags
		
		If ((tmp.AddedBy = vbNullString) Or (tmp.AddedBy = "%")) Then
			lblCreatedOn.Text = "unknown"
			lblCreatedBy.Text = "by unknown"
		Else
			lblCreatedOn.Text = tmp.AddedOn & " Local Time"
			lblCreatedBy.Text = "by " & tmp.AddedBy
		End If
		
		If ((tmp.ModifiedBy = vbNullString) Or (tmp.ModifiedBy = "%")) Then
			lblModifiedOn.Text = "unknown"
			lblModifiedBy.Text = "by unknown"
		Else
			lblModifiedOn.Text = tmp.ModifiedOn & " Local Time"
			lblModifiedBy.Text = "by " & tmp.ModifiedBy
		End If
		
		' is entry a member of a group?
		If (Len(tmp.Groups) > 0) And (StrComp(tmp.Groups, "%", CompareMethod.Binary) <> 0) Then
			' is entry a member of multiple groups?
			If (InStr(1, tmp.Groups, ",", CompareMethod.Binary) <> 0) Then
				' store working copy of group memberships, splitting up
				' multiple groupings by the ',' delimiter.
				Splt = Split(tmp.Groups, ",")
			Else
				' redefine array size to store group name
				ReDim Preserve Splt(0)
				
				' store working copy of group membership
				Splt(0) = tmp.Groups
			End If
		End If
		
		Call UpdateInheritCaption(tmp.Groups)
		
		' loop through our listview, checking for matches
		If m_glistcount > 0 Then
			For j = 1 To lvGroups.Items.Count
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				With lvGroups.Items.Item(j)
					' loop through entry's group memberships
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.Checked = False
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    .ImageIndex = IC_UNCHECKED
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    '.Ghosted = False
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.ForeColor = System.Drawing.Color.White
					
					If (Len(tmp.Groups) > 0) And (StrComp(tmp.Groups, "%", CompareMethod.Binary) <> 0) Then
						For i = LBound(Splt) To UBound(Splt)
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							If (StrComp(Splt(i), "%", CompareMethod.Binary) <> 0) And (StrComp(Splt(i), .Text, CompareMethod.Text) = 0) Then
								' select group if entry is a member
								'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								.Checked = True
								'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                .ImageIndex = IC_CHECKED
								
								' highlight group if "primary" (first group)
								If (i = LBound(Splt)) Then
									'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									.ForeColor = System.Drawing.Color.Yellow
									'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                    .ImageIndex = IC_PRIMARY
								End If
								
								Exit For
							End If
						Next i
					End If
					
					If (StrComp(tmp.Type, "GROUP", CompareMethod.Text) = 0) Then
						' don't allow groups to contain themself
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						If (StrComp(tmp.Username, .Text, CompareMethod.Text) = 0) Then
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							.Checked = False
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            .ImageIndex = IC_UNCHECKED
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            '.Ghosted = True
							'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
						End If
					End If
				End With
			Next j
		End If
		
		If ((tmp.BanMessage <> vbNullString) And (tmp.BanMessage <> "%")) Then
			txtBanMessage.Text = tmp.BanMessage
		End If
		
		Call UnlockGUI()
		
		node.Selected = True
		
		' refresh tree view
        Call trvUsers.Refresh()
	End Sub
	
	'Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
	'Dim node As cTreeViewNode
	
	'If (Button = vbLeftButton) Then
	'    Set node = trvUsers.HitTest(x, y)
	'    If Not node Is Nothing Then
	'        node.Selected = True
	'        Call trvUsers_NodeClick(trvUsers.SelectedItem)
	'    End If
	'End If
	'End Sub
	
	'Private Sub trvUsers_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, _
	''    Shift As Integer, x As Single, y As Single, State As Integer)
	'
	'    Set trvUsers.SelectedItem = trvUsers.HitTest(x, y)
	'End Sub
	
	'Private Sub trvUsers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, _
	''    Shift As Integer, x As Single, y As Single)
	'
	'    On Error GoTo ERROR_HANDLER
	'
	'    Dim nodePrev As cTreeViewNode
	'    Dim nodeNow  As cTreeViewNode
	'
	'    Dim strKey   As String
	'    Dim res      As Integer
	'    Dim i        As Integer
	'    Dim found    As Integer '
	'
	'    If (Data.GetFormat(15) = True) Then
	'        With frmDBType
	'            .setFilePath Data.Files(1)
	'            .Show
	'        End With
	'    Else
	'        Set nodeNow = trvUsers.NodeFromDragData(Data)
	'
	'        Set nodePrev = trvUsers.SelectedItem
	'
	'        If (nodeNow.Index = 1) Then
	'            For i = LBound(m_DB) To UBound(m_DB)
	'                If (StrComp(m_DB(i).Username, nodePrev.Text, vbTextCompare) = 0) Then
	'                    If (StrComp(m_DB(i).Type, nodePrev.Tag, vbTextCompare) = 0) Then
	'                        If ((Len(m_DB(i).Groups) > 0) And (m_DB(i).Groups <> "%")) Then
	'                            Set nodePrev.Parent = nodeNow
	'                        End If
	'
	'                        m_DB(i).Groups = vbNullString
	'                        Exit For
	'                    End If
	'                End If
	'            Next i
	'        Else
	'            If (Not nodePrev.Index = 1) Then
	'                If (Not StrComp(nodeNow.Tag, "GROUP", vbTextCompare) = 0) Then
	'                    Set nodeNow = nodeNow.Parent
	'
	'                    If (nodeNow.Index = 1) Then
	'                        Set trvUsers.SelectedItem = nodeNow
	'
	'                        Call trvUsers_OLEDragDrop(Data, Effect, Button, Shift, x, y)
	'                        Exit Sub
	'                    End If
	'                End If
	'
	'                'If (IsInGroup(nodePrev.Text, nodeNow.Text) = False) Then
	'                '    For i = LBound(m_DB) To UBound(m_DB)
	'                '        If (StrComp(m_DB(i).Username, nodePrev.Text, vbTextCompare) = 0) Then
	'                '            If (StrComp(m_DB(i).Type, nodePrev.Tag, vbTextCompare) = 0) Then
	'                '                m_DB(i).Groups = nodeNow.Text
	'                '                Exit For
	'                '            End If
	'                '        End If
	'                '    Next i
	'                '
	'                '    Set nodePrev.Parent = nodeNow
	'                'End If
	'            End If
	'        End If
	'
	'        Call trvUsers_NodeClick(nodePrev)
	'        'Set trvUsers.DropHighlight = Nothing
	'    End If
	'
	'    Exit Sub
	'
	'ERROR_HANDLER:
	'    ' potential cycle introduction error
	'    If (Err.Number = 35614) Then
	'        MsgBox Err.description, vbCritical, "Error"
	'    End If
	'
	'    'Set trvUsers.DropHighlight = Nothing
	'
	'    Exit Sub
	'End Sub
	
	Private Function IsInGroup(ByRef Groups As String, ByVal GroupName As String) As Boolean
		Dim j As Short
		Dim Splt() As String
		
		IsInGroup = False
		
		If (Len(Groups) > 0) Then
			If (Not InStr(1, Groups, ",", CompareMethod.Binary) = 0) Then
				Splt = Split(Groups, ",")
			Else
				ReDim Splt(0)
				Splt(0) = Groups
			End If
			
			For j = LBound(Splt) To UBound(Splt)
				If (StrComp(Splt(j), "%", CompareMethod.Binary) <> 0) And (StrComp(GroupName, Splt(j), CompareMethod.Text) = 0) Then
					IsInGroup = True
					
					Exit Function
				End If
			Next j
		End If
	End Function
	
	Private Sub UpdateGroupList()
		Dim i As Short
		Dim Count As Short
		
		' clear group selection listing
		Call lvGroups.Items.Clear()
		
		m_glistcount = 0
		
		' go through group listing
		For i = LBound(m_DB) To UBound(m_DB)
			If (StrComp(m_DB(i).Type, "GROUP", CompareMethod.Text) = 0) Then
				m_glistcount = m_glistcount + 1
				' add group to group selection listbox
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
				With lvGroups.Items.Add(m_DB(i).Username, IC_UNCHECKED)
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
					.ForeColor = System.Drawing.Color.White
				End With
			End If
		Next i
		
		If m_glistcount = 0 Then
			'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
			With lvGroups.Items.Add("[none]", IC_EMPTY)
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Add.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                '.Ghosted = True
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
				.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
			End With
		End If
	End Sub
	
	Private Sub UpdateInheritCaption(ByRef Groups As String)
		Dim grpwlk As udtGetAccessResponse
		
		'UPGRADE_WARNING: Couldn't resolve default property of object grpwlk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		grpwlk = GetAccessGroupWalk(Groups)
		
		lblInherit.Text = vbNullString
        If grpwlk.Rank > 0 And Len(grpwlk.Flags) > 0 Then
            lblInherit.Text = "Inherits rank " & grpwlk.Rank & " and flags " & grpwlk.Flags & " from groups."
        ElseIf grpwlk.Rank > 0 Then
            lblInherit.Text = "Inherits rank " & grpwlk.Rank & " from groups."
        ElseIf Len(grpwlk.Flags) > 0 Then
            lblInherit.Text = "Inherits flags " & grpwlk.Flags & " from groups."
        End If
	End Sub
	
	Private Sub ClearGroupListChecks()
		Dim i As Short
		
		' loop through listbox and clear selected items
		For i = 1 To lvGroups.Items.Count
			'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			With lvGroups.Items.Item(i)
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Checked = False
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .ImageIndex = IIf(m_glistcount = 0, IC_EMPTY, IC_UNCHECKED)
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                '.Ghosted = True
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.ForeColor = System.Drawing.ColorTranslator.FromOle(&H888888)
			End With
		Next i
	End Sub
	
	Private Function GetPrimaryGroup(ByVal Groups As String) As String
		Dim grp As String
		Dim Splt() As String
		
		' is it not in a group?
        If (Len(Groups) = 0 Or StrComp(Groups, "%", CompareMethod.Binary) = 0) Then
            grp = vbNullString
            ' is entry member of multiple groups?
        ElseIf (InStr(1, Groups, ",", CompareMethod.Binary) <> 0) Then
            ' split up multiple groupings
            Splt = Split(Groups, ",")
            grp = Splt(0)
        Else
            ' no need for special handling...
            grp = Groups
        End If
		
		GetPrimaryGroup = grp
	End Function
	
	Private Function GetLVPrimaryGroup() As System.Windows.Forms.ListViewItem
		Dim i As Short
		
		For i = 1 To lvGroups.Items.Count
			With lvGroups.Items
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Item.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                If (.Item(i).ForeColor.Equals(System.Drawing.Color.Yellow) And Not False And .Item(i).Checked) Then
                    'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    GetLVPrimaryGroup = .Item(i)
                    Exit Function
                End If
			End With
		Next i
		
		'UPGRADE_NOTE: Object GetLVPrimaryGroup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GetLVPrimaryGroup = Nothing
	End Function
	
	Private Sub SetLVPrimaryGroup(ByRef ListItem As System.Windows.Forms.ListViewItem)
		Dim i As Short
		
		If (Not ListItem Is Nothing) Then
			For i = 1 To lvGroups.Items.Count
				'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				With lvGroups.Items.Item(i)
					'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If (StrComp(.Text, ListItem.Text, CompareMethod.Text) = 0) Then
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						.ForeColor = System.Drawing.Color.Yellow
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						.Checked = True
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .ImageIndex = IC_PRIMARY
						'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Item.Ghosted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    ElseIf (Not False) Then
                        'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        .ForeColor = System.Drawing.Color.White
                        'UPGRADE_WARNING: Lower bound of collection lvGroups.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        'UPGRADE_ISSUE: MSComctlLib.ListItem property lvGroups.ListItems.Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .ImageIndex = IIf(.Checked, IC_CHECKED, IC_UNCHECKED)
					End If
				End With
			Next i
		End If
	End Sub
	
	'UPGRADE_NOTE: Tag was upgraded to Tag_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function FindNodeIndex(ByVal nodeName As String, Optional ByRef Tag_Renamed As String = vbNullString) As Short
		Dim i As Short
		
		For i = 1 To trvUsers.NodeCount
			If (StrComp(trvUsers.nodes(i).Text, nodeName, CompareMethod.Text) = 0) Then
                If (Len(Tag_Renamed) > 0) Then
                    If (StrComp(trvUsers.Nodes(i).Tag, Tag_Renamed, CompareMethod.Text) = 0) Then
                        FindNodeIndex = i
                        Exit Function
                    End If
                Else
                    FindNodeIndex = i
                    Exit Function
                End If
			End If
		Next i
		
		FindNodeIndex = 0
	End Function
	
	'UPGRADE_WARNING: Event txtBanMessage.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBanMessage_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBanMessage.TextChanged
		' enable entry save button
		Call HandleUnsaved()
	End Sub
	
	Private Sub txtFlags_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlags.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		Const AZ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		
		' disallow entering space
		If (KeyAscii = 32) Then KeyAscii = 0
		
		' if key is A-Z, then make uppercase
		If (InStr(1, AZ, Chr(KeyAscii), CompareMethod.Text) > 0) Then
			If (BotVars.CaseSensitiveFlags = False) Then
				If (KeyAscii > 90) Then ' lowercase if greater than "Z"
					KeyAscii = Asc(UCase(Chr(KeyAscii)))
				End If
			End If
			' else disallow entering that character (if not a control character)
		ElseIf (KeyAscii > 32) Then 
			KeyAscii = 0
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtFlags.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtFlags_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlags.TextChanged
		' enable entry save button
		Call HandleUnsaved()
	End Sub
	
	Private Sub txtRank_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRank.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Const n09 As String = "0123456789"
		
		' disallow entering space
		If (KeyAscii = 32) Then KeyAscii = 0
		
		' if key is not 0-9, disallow entering that character (if not a control character)
		If (InStr(1, n09, Chr(KeyAscii), CompareMethod.Text) = 0 And KeyAscii > 32) Then
			KeyAscii = 0
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtRank.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRank_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRank.TextChanged
		Dim SelStart As Integer
		
		If (Val(txtRank.Text) > 200) Then
			With txtRank
				SelStart = .SelectionStart
				.Text = "200"
				.SelectionStart = SelStart
			End With
		End If
		
		' enable entry save button
		Call HandleUnsaved()
	End Sub
	
	Private Function GetAccess(ByVal Username As String, Optional ByRef dbType As String = vbNullString, Optional ByRef Index As Short = 0) As udtGetAccessResponse
		
		Dim i As Short
		Dim bln As Boolean
		
		dbType = UCase(dbType)
		
		For i = LBound(m_DB) To UBound(m_DB)
			If (StrComp(m_DB(i).Username, Username, CompareMethod.Text) = 0) Then
				If (Len(dbType) > 0) Then
					If (StrComp(m_DB(i).Type, dbType, CompareMethod.Text) = 0) Then
						bln = True
					End If
				Else
					bln = True
				End If
				
				If (bln = True) Then
					Index = i
					
					With GetAccess
						.Username = m_DB(i).Username
						.Rank = m_DB(i).Rank
						.Flags = m_DB(i).Flags
						.AddedBy = m_DB(i).AddedBy
						.AddedOn = m_DB(i).AddedOn
						.ModifiedBy = m_DB(i).ModifiedBy
						.ModifiedOn = m_DB(i).ModifiedOn
						.Type = m_DB(i).Type
						.Groups = m_DB(i).Groups
						.BanMessage = m_DB(i).BanMessage
					End With
					
					Exit Function
				End If
			End If
			
			bln = False
		Next i
		
		GetAccess.Rank = -1
	End Function
	
	' gets combined access of all Groups containing this item
	Private Function GetAccessGroupWalk(ByRef Groups As String) As udtGetAccessResponse
		Dim Splt() As String
		Dim Group As String
		Dim AllGroups As New Collection
		Dim tmp As udtGetAccessResponse
		Dim MaxRank As Short
		Dim CombFlags As String
		Dim i As Short
		Dim j As Short
		
        If Len(Groups) > 0 Then
            Splt = Split(Groups, ",")
            For j = LBound(Splt) To UBound(Splt)
                If (StrComp(Splt(j), "%", CompareMethod.Binary) <> 0) Then
                    On Error GoTo ERROR_HANDLER
                    Call AllGroups.Add(Splt(j), Splt(j))
                    On Error GoTo 0
                End If
            Next j
        End If
		
		i = 1
		Do While i <= AllGroups.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object AllGroups.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tmp = GetAccess(AllGroups.Item(i), "GROUP")
			If tmp.Rank > MaxRank Then MaxRank = tmp.Rank
			CombFlags = CombFlags & tmp.Flags
			
            If Len(tmp.Groups) > 0 And StrComp(tmp.Groups, "%", CompareMethod.Text) <> 0 Then
                Splt = Split(tmp.Groups, ",")
                For j = LBound(Splt) To UBound(Splt)
                    If (StrComp(Splt(j), "%", CompareMethod.Binary) <> 0) Then
                        On Error GoTo ERROR_HANDLER
                        Call AllGroups.Add(Splt(j), Splt(j))
                        On Error GoTo 0
                    End If
                Next j
            End If
			i = i + 1
		Loop 
		
		With GetAccessGroupWalk
			.Username = "(all groups)"
			.Rank = MaxRank
			.Flags = CombFlags
			.AddedBy = "(console)"
			.AddedOn = Now
			.ModifiedBy = vbNullString
			.ModifiedOn = System.Date.FromOADate(0)
			.Type = "GROUP"
			.Groups = vbNullString
			.BanMessage = vbNullString
		End With
		
		Exit Function
ERROR_HANDLER: 
		If (Err.Number = 457) Then
			Resume Next
		End If
	End Function
	
	Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = vbNullString) As Boolean
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		Dim found As Boolean
		
		dbType = UCase(dbType)
		
		Dim bln As Boolean
		For i = LBound(m_DB) To UBound(m_DB)
			If (StrComp(m_DB(i).Username, entry, CompareMethod.Text) = 0) Then
				
				If (Len(dbType)) Then
					If (StrComp(m_DB(i).Type, dbType, CompareMethod.Text) = 0) Then
						bln = True
					End If
				Else
					bln = True
				End If
				
				If (bln) Then
					found = True
					
					Exit For
				End If
			End If
			
			bln = False
		Next i
		
		Dim bak As udtDatabase
		Dim j As Short
		Dim res As Boolean
		Dim Splt() As String
		Dim innerfound As Boolean
		Dim K As Short
		If (found) Then
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object bak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bak = m_DB(i)
			
			' we aren't removing the last array
			' element, are we?
			If (UBound(m_DB) = 0) Then
				ReDim m_DB(0)
				
				With m_DB(0)
					.Username = vbNullString
					.Flags = vbNullString
					.Rank = 0
					.Groups = vbNullString
					.AddedBy = vbNullString
					.ModifiedBy = vbNullString
					.AddedOn = Now
					.ModifiedOn = Now
				End With
			Else
				For j = i To UBound(m_DB) - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object m_DB(j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_DB(j) = m_DB(j + 1)
				Next j
				
				' redefine array size
				ReDim Preserve m_DB(UBound(m_DB) - 1)
				
				' if we're removing a group, we need to also fix our
				' group memberships, in case anything is broken now
				If (StrComp(bak.Type, "GROUP", CompareMethod.Binary) = 0) Then
					
					' if we remove a user from the database during the
					' execution of the inner loop, we have to reset our
					' inner loop variables, otherwise we create errors
					' due to incorrect database indexes.  Because of this,
					' we have to dual-loop until our inner loop runs out
					' of matching users.
					Do 
						' reset loop variable
						res = False
						
						' loop through database checking for users that
						' were members of the group that we just removed
						For i = LBound(m_DB) To UBound(m_DB)
							If (Len(m_DB(i).Groups) > 0) Then
								If (InStr(1, m_DB(i).Groups, ",", CompareMethod.Binary) <> 0) Then
									
									Splt = Split(m_DB(i).Groups, ",")
									
									For j = LBound(Splt) To UBound(Splt)
										If (StrComp(Splt(j), "%", CompareMethod.Binary) <> 0) And (StrComp(bak.Username, Splt(j), CompareMethod.Text) = 0) Then
											innerfound = True
											
											Exit For
										End If
									Next j
									
									If (innerfound) Then
										
										For K = (j + 1) To UBound(Splt)
											Splt(K - 1) = Splt(K)
										Next K
										
										ReDim Preserve Splt(UBound(Splt) - 1)
										
										m_DB(i).Groups = Join(Splt, vbNullString)
									End If
								Else
									If (StrComp(bak.Username, m_DB(i).Groups, CompareMethod.Text) = 0) Then
										res = DB_remove(m_DB(i).Username, m_DB(i).Type)
										
										Exit For
									End If
								End If
							End If
						Next i
					Loop While (res)
				End If
			End If
			
			' commit modifications
			'Call WriteDatabase(GetFilePath(FILE_USERDB))
			
			DB_remove = True
			
			Exit Function
		End If
		
		DB_remove = False
		
		Exit Function
		
ERROR_HANDLER: 
		Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DB_remove().")
		
		DB_remove = False
		
		Exit Function
	End Function
End Class