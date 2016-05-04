Option Strict Off
Option Explicit On
Friend Class frmDBNameEntry
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDBManager.m_entryname = vbNullString
		
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		frmDBManager.m_entryname = txtEntry.Text
		
		Me.Close()
	End Sub
	
	Private Sub frmDBNameEntry_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Text = StringFormat("New Entry - {0} Name", frmDBManager.m_entrytype)
		
		With txtEntry
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If LenB(frmDBManager.m_entryname) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblEntry.Text = StringFormat("Choose the name for this new {0} entry.", frmDBManager.m_entrytype)
				.Text = vbNullString
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lblEntry.Text = StringFormat("Rename this {0} entry.", frmDBManager.m_entrytype)
				.Text = frmDBManager.m_entryname
				.SelectionStart = 0
				.SelectionLength = Len(frmDBManager.m_entryname)
			End If
			
			If StrComp(frmDBManager.m_entrytype, "Clan", CompareMethod.Text) = 0 Then
				'UPGRADE_WARNING: TextBox property txtEntry.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				.Maxlength = 4
			Else
				'UPGRADE_WARNING: TextBox property txtEntry.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				.Maxlength = 30
			End If
			
			cmdOK.Enabled = CanSave()
		End With
	End Sub
	
	'UPGRADE_WARNING: Event txtEntry.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtEntry_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEntry.TextChanged
		With txtEntry
			cmdOK.Enabled = CanSave()
		End With
	End Sub
	
	Private Sub txtName_KeyPress(ByRef KeyAscii As Short)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdOK_Click(cmdOK, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdCancel_Click(cmdCancel, New System.EventArgs())
		End If
	End Sub
	
	Private Function CanSave() As Boolean
		With txtEntry
			CanSave = True
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CanSave = CanSave And (LenB(.Text) > 0)
			'UPGRADE_WARNING: TextBox property txtEntry.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			CanSave = CanSave And (Len(.Text) <= .Maxlength)
			CanSave = CanSave And (StrComp(.Text, "%", CompareMethod.Binary) <> 0)
			CanSave = CanSave And (InStr(1, .Text, " ", CompareMethod.Binary) = 0)
			CanSave = CanSave And (InStr(1, .Text, ",", CompareMethod.Binary) = 0)
			CanSave = CanSave And (InStr(1, .Text, Chr(34), CompareMethod.Binary) = 0)
		End With
	End Function
End Class