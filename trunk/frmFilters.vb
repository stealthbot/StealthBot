Option Strict Off
Option Explicit On
Friend Class frmFilters
	Inherits System.Windows.Forms.Form
	Private OldMaxBlockIndex As Short
	Private OldMaxFilterIndex As Short
	' Updated later to erase old entries
	
	Private Sub cmdAdd_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdAdd.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call frmFilters_MouseMove(Me, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 0, 0))
	End Sub
	
	Private Sub cmdDone_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdDone.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call frmFilters_MouseMove(Me, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 0, 0))
	End Sub
	
	Private Sub cmdOutAdd_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdOutAdd.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call frmFilters_MouseMove(Me, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 0, 0))
	End Sub
	
	Private Sub cmdEdit_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdEdit.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Allows you to edit the selected item."
	End Sub
	
	Private Sub cmdOutAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOutAdd.Click
		If txtOutAdd(0).Text <> vbNullString And txtOutAdd(1).Text <> vbNullString Then
			'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			lvReplace.Items.Insert(lvReplace.Items.Count + 1, txtOutAdd(0).Text)
			'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			lvReplace.Items.Item(lvReplace.Items.Count).SubItems.Add(txtOutAdd(1).Text)
			txtOutAdd(0).Text = vbNullString
			txtOutAdd(1).Text = vbNullString
			txtOutAdd(0).Focus()
		End If
	End Sub
	
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		If lbText.SelectedIndex <> -1 Then
			txtAdd.Text = lbText.Text
			Call cmdRem_Click(cmdRem, New System.EventArgs())
		ElseIf lbBlock.SelectedIndex <> -1 Then 
			txtAdd.Text = lbBlock.Text
			Call cmdRem_Click(cmdRem, New System.EventArgs())
		ElseIf Not (lvReplace.FocusedItem Is Nothing) Then 
			'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			txtOutAdd(0).Text = lvReplace.Items.Item(lvReplace.FocusedItem.Index).Text
			'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems.Item().ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			txtOutAdd(1).Text = lvReplace.Items.Item(lvReplace.FocusedItem.Index).SubItems.Item(1).Text
			Call cmdOutRem_Click(cmdOutRem, New System.EventArgs())
		End If
	End Sub
	
	Private Sub cmdOutRem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOutRem.Click
		If Not (lvReplace.FocusedItem Is Nothing) Then
			If lvReplace.FocusedItem.Index > 0 Then lvReplace.Items.RemoveAt(lvReplace.FocusedItem.Index)
		End If
	End Sub
	
	Private Sub cmdOutRem_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdOutRem.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Removes the selected item from the list."
	End Sub
	
	Private Sub frmFilters_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		Dim i As Short
		Dim s As String
		
		optText.Checked = True
		
		s = ReadINI("TextFilters", "Total", FILE_FILTERS)
		If StrictIsNumeric(s) Then
			i = Val(s)
			OldMaxFilterIndex = i
			
			If i > 0 Then
				For i = 1 To i
					s = ReadINI("TextFilters", "Filter" & i, FILE_FILTERS)
					
                    If Len(s) > 0 Then
                        lbText.Items.Add(s)
                    End If
				Next i
			End If
		End If
		
		s = ReadINI("BlockList", "Total", FILE_FILTERS)
		If StrictIsNumeric(s) Then
			i = Val(s)
			OldMaxBlockIndex = i
			
			If i > 0 Then
				For i = 1 To i
					s = ReadINI("BlockList", "Filter" & i, FILE_FILTERS)
					
                    If Len(s) > 0 Then
                        lbBlock.Items.Add(s)
                    End If
				Next i
			End If
		End If
		
		s = ReadINI("Outgoing", "Total", FILE_FILTERS)
		If StrictIsNumeric(s) Then
			i = Val(s)
			
			For i = 1 To i
				s = Replace(ReadINI("Outgoing", "Find" & i, FILE_FILTERS), "¦", " ")
				
                If Len(s) > 0 Then
                    'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    lvReplace.Items.Insert(lvReplace.Items.Count + 1, s)
                    'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    lvReplace.Items.Item(lvReplace.Items.Count).SubItems.Add(Replace(ReadINI("Outgoing", "Replace" & i, FILE_FILTERS), "¦", " "))
                End If
			Next i
		End If
		
		Call optType_CheckedChanged(optType.Item(0), New System.EventArgs())
	End Sub
	
	Private Sub cmdDone_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDone.Click
		Me.Close()
	End Sub
	
	Private Sub cmdRem_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdRem.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Clicking this button will remove any usernames or filters that you have selected on either list."
	End Sub
	
	Private Sub cmdRem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRem.Click
		If lbText.SelectedIndex <> -1 Then lbText.Items.RemoveAt(lbText.SelectedIndex)
		If lbBlock.SelectedIndex <> -1 Then lbBlock.Items.RemoveAt(lbBlock.SelectedIndex)
	End Sub
	
	Private Sub lbBlock_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbBlock.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Adding usernames to this list will cause the bot to completely block any and all messages from their username. The filter system supports wildcards. Example: Adding 'floodbot*' to the list will cause the bot to stop any messages coming from anybody whose name starts with the word 'FloodBot'. The filter is not applied to whispers."
	End Sub
	
	Private Sub lbText_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbText.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Adding filters to this list will cause the bot to completely block any and all messages containing the text you added. Useful for blocking annoying floodbot or other spams."
	End Sub
	
	Private Sub frmFilters_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		lblMI.Text = "Incoming chat filters are toggled by pressing CTRL + F inside the bot." & vbNewLine & "Outgoing filters are permanently active."
	End Sub
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		If optBlock.Checked = True Then
			lbBlock.Items.Add(txtAdd.Text)
		Else
			lbText.Items.Add(txtAdd.Text)
		End If
		txtAdd.Text = vbNullString
	End Sub
	
	Private Sub frmFilters_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim i As Short
		i = -1
		
		' Write text filters
		If lbText.Items.Count > 0 Then
			For i = 1 To lbText.Items.Count
				WriteINI("TextFilters", "Filter" & i, VB6.GetItemString(lbText, i - 1), FILE_FILTERS)
			Next i
			
			WriteINI("TextFilters", "Total", CStr(lbText.Items.Count), FILE_FILTERS)
		Else
			WriteINI("TextFilters", "Total", CStr(0), FILE_FILTERS)
		End If
		
		' Erase old text filters
		If i >= 0 And i < OldMaxFilterIndex Then
			For i = i To OldMaxFilterIndex
				WriteINI("TextFilters", "Filter" & i, "-", FILE_FILTERS)
			Next i
		End If
		
		' Write block list
		i = -1
		If lbBlock.Items.Count <> 0 Then
			For i = 1 To lbBlock.Items.Count
				WriteINI("BlockList", "Filter" & i, VB6.GetItemString(lbBlock, i - 1), FILE_FILTERS)
			Next i
			
			WriteINI("BlockList", "Total", CStr(lbBlock.Items.Count), FILE_FILTERS)
		Else
			WriteINI("BlockList", "Total", CStr(0), FILE_FILTERS)
		End If
		
		' Erase old blocked items
		If i >= 0 And i < OldMaxBlockIndex Then
			For i = i To OldMaxBlockIndex
				WriteINI("BlockList", "Filter" & i, "-", FILE_FILTERS)
			Next i
		End If
		
		' Write new outgoing filters
		If lvReplace.Items.Count <> 0 Then
			For i = 1 To lvReplace.Items.Count
				'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				WriteINI("Outgoing", "Find" & i, Replace(lvReplace.Items.Item(i).Text, " ", "¦"), FILE_FILTERS)
				'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection lvReplace.ListItems.Item().ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				WriteINI("Outgoing", "Replace" & i, Replace(lvReplace.Items.Item(i).SubItems.Item(1).Text, " ", "¦"), FILE_FILTERS)
			Next i
			
			WriteINI("Outgoing", "Total", CStr(lvReplace.Items.Count), FILE_FILTERS)
		Else
			WriteINI("Outgoing", "Total", CStr(0), FILE_FILTERS)
		End If
		
		Call frmChat.LoadOutFilters()
		Call frmChat.LoadArray(LOAD_FILTERS, gFilters)
	End Sub
	
	Private Sub optBlock_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles optBlock.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call lbBlock_MouseMove(lbBlock, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 0, 0))
	End Sub
	
	Private Sub optText_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles optText.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call lbText_MouseMove(lbText, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 0, 0))
	End Sub
	
	'UPGRADE_WARNING: Event optType.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optType.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optType.GetIndex(eventSender)
			Dim i As Byte
			Dim A As Boolean
			
			If Index = 0 Then 'Incoming Filters Clicked
				optType(1).Checked = False
				optType(0).Checked = True
				A = True
			Else 'Outgoing Filters Clicked
				optType(0).Checked = False
				optType(1).Checked = True
				A = False
			End If
			
			lbText.Visible = A
			lbBlock.Visible = A
			txtAdd.Visible = A
			cmdAdd.Visible = A
			cmdRem.Visible = A
			
			For i = 0 To IncomingLbl.UBound
				IncomingLbl(i).Visible = A
			Next i
			
			For i = 0 To OutgoingLbl.UBound
				OutgoingLbl(i).Visible = Not A
			Next i
			
			optText.Visible = A
			optBlock.Visible = A
			lvReplace.Visible = Not A
			txtOutAdd(0).Visible = Not A
			txtOutAdd(1).Visible = Not A
			cmdOutRem.Visible = Not A
			cmdOutAdd.Visible = Not A
		End If
	End Sub
	
	Private Sub optType_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles optType.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = optType.GetIndex(eventSender)
		If Index = 0 Then
			lblMI.Text = "Change your incoming chat filters (message filters) here."
		Else
			lblMI.Text = "Change your outgoing chat filters here."
		End If
	End Sub
End Class