Option Strict Off
Option Explicit On
Friend Class frmManageKeys
	Inherits System.Windows.Forms.Form
	
	Private Const FILE_KEY_STORAGE As String = "Keys.txt"
	
	Private KeyProducts As Scripting.Dictionary
	
	Private m_editing As String
	
	Private Sub frmManageKeys_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		' Set default button states
		cmdAdd.Enabled = False
		cmdEdit.Enabled = False
		cmdDelete.Enabled = False
		cmdSetKey.Enabled = False
		
		' Create a lookup for product codes to icon indexes
		KeyProducts = New Scripting.Dictionary
		KeyProducts.Add(&H1, 3) ' STAR
		KeyProducts.Add(&H2, 3) ' STAR
		KeyProducts.Add(&H4, 4) ' W2BN
		KeyProducts.Add(&H5, 1) ' D2DV (beta, defunct)
		KeyProducts.Add(&H6, 5) ' D2DV
		KeyProducts.Add(&H7, 5) ' D2DV
		KeyProducts.Add(&H9, 1) ' D2DV (stress test, defunct)
		KeyProducts.Add(&HA, 6) ' D2XP
		KeyProducts.Add(&HC, 6) ' D2XP
		KeyProducts.Add(&HD, 1) ' WAR3 (beta, defunct)
		KeyProducts.Add(&HE, 7) ' WAR3
		KeyProducts.Add(&HF, 7) ' WAR3
		KeyProducts.Add(&H11, 1) ' W3XP (beta, defunct)
		KeyProducts.Add(&H12, 8) ' W3XP
		KeyProducts.Add(&H13, 1) ' W3XP (retail, disabled)
		KeyProducts.Add(&H17, 3) ' STAR (online upgrade)
		KeyProducts.Add(&H18, 5) ' D2DV (online upgrade)
		KeyProducts.Add(&H19, 6) ' D2XP (online upgrade)
		
		Call Local_LoadCDKeys()
		
		If lvKeys.Items.Count > 0 Then
			'UPGRADE_WARNING: Lower bound of collection lvKeys.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			lvKeys.FocusedItem = lvKeys.Items.Item(1)
			lvKeys_Click(lvKeys, New System.EventArgs())
		End If
	End Sub
	
	Private Sub frmManageKeys_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim SettingsForm As Object
		On Error Resume Next
		frmChat.SettingsForm.Show()
		frmChat.SettingsForm.Activate()
	End Sub
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		ProcessKey(txtActiveKey.Text)
		
		m_editing = vbNullString
		
		txtActiveKey.Text = vbNullString
		lvKeys.Focus()
	End Sub
	
	Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
		If Not (lvKeys.FocusedItem Is Nothing) Then
			lvKeys.Items.RemoveAt(lvKeys.FocusedItem.Index)
		End If
	End Sub
	
	Private Sub cmdDone_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDone.Click
        If Len(m_editing) > 0 Then
            ProcessKey(m_editing)
            m_editing = vbNullString
        End If
		
		Call Local_WriteCDKeys()
		
		Me.Close()
	End Sub
	
	Private Sub cmdSetKey_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSetKey.Click
		Dim SettingsForm As Object
		If ((Not (frmChat.SettingsForm Is Nothing)) And (Not (lvKeys.FocusedItem Is Nothing))) Then
			'UPGRADE_ISSUE: MSComctlLib.ListItem property lvKeys.SelectedItem.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            If ((lvKeys.FocusedItem.ImageIndex = 6) Or (lvKeys.FocusedItem.ImageIndex = 8)) Then
                frmChat.SettingsForm.txtExpKey.Text = lvKeys.FocusedItem.Tag
            Else
                frmChat.SettingsForm.txtCDKey.Text = lvKeys.FocusedItem.Tag
            End If
			
			Call cmdDone_Click(cmdDone, New System.EventArgs())
		End If
	End Sub
	
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		If Not (lvKeys.FocusedItem Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object lvKeys.SelectedItem.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_editing = lvKeys.FocusedItem.Tag
			lvKeys.Items.RemoveAt(lvKeys.FocusedItem.Index)
			With txtActiveKey
				.Text = m_editing
				.Focus()
				.SelectionStart = 0
				.SelectionLength = Len(.Text)
			End With
		End If
	End Sub
	
	Private Sub lvKeys_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvKeys.Click
		Dim Value As Boolean
		Value = (Not (lvKeys.FocusedItem Is Nothing))
		
		cmdEdit.Enabled = Value
		cmdDelete.Enabled = Value
		cmdSetKey.Enabled = Value
	End Sub
	
	Private Sub lvKeys_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvKeys.DoubleClick
		Call cmdSetKey_Click(cmdSetKey, New System.EventArgs())
	End Sub
	
	Private Sub lvKeys_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvKeys.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdSetKey_Click(cmdSetKey, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdDone_Click(cmdDone, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtActiveKey_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveKey.Enter
		cmdEdit.Enabled = False
		cmdDelete.Enabled = False
		cmdSetKey.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event txtActiveKey.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtActiveKey_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveKey.TextChanged
		cmdAdd.Enabled = (Len(txtActiveKey.Text) > 0)
	End Sub
	
	Private Sub txtActiveKey_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtActiveKey.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return And cmdAdd.Enabled Then
			Call cmdAdd_Click(cmdAdd, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
            If Len(m_editing) > 0 Then
                ProcessKey(m_editing)

                m_editing = vbNullString

                txtActiveKey.Text = vbNullString
                lvKeys.Focus()
            ElseIf Len(txtActiveKey.Text) > 0 Then
                txtActiveKey.Text = vbNullString
            Else
                Call cmdDone_Click(cmdDone, New System.EventArgs())
            End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub ProcessKey(ByVal sKey As String)
		Dim oKey As New clsKeyDecoder
		Dim KeyProduct As Integer
		
		oKey.Initialize(sKey)
		If Not oKey.IsValid Then
			KeyProduct = -1
		Else
			KeyProduct = oKey.ProductValue
		End If
		
		AddUnique(oKey.GetKeyForDisplay(), GetImageIndex(KeyProduct), oKey.Key)
		
		'UPGRADE_NOTE: Object oKey may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oKey = Nothing
	End Sub
	
	' Adds the specified text and image while checking for duplicates
	'UPGRADE_NOTE: Tag was upgraded to Tag_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub AddUnique(ByVal strNewValue As String, ByVal image As Short, ByVal Tag_Renamed As String)
		Dim Item As System.Windows.Forms.ListViewItem
		
		For	Each Item In lvKeys.Items
			If StrComp(Item.Tag, Tag_Renamed, CompareMethod.Text) = 0 Then Exit Sub
		Next Item
		
		AddItem(strNewValue, image, Tag_Renamed)
	End Sub
	
	' Adds the specified text and image regardless of duplicates
	'UPGRADE_NOTE: Tag was upgraded to Tag_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub AddItem(ByVal Text_Renamed As String, ByVal image As Short, ByVal Tag_Renamed As String)
		'UPGRADE_WARNING: Lower bound of collection lvKeys.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		With lvKeys.Items.Add(Text_Renamed, image)
			'UPGRADE_WARNING: Lower bound of collection lvKeys.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
			.Tag = Tag_Renamed
		End With
	End Sub
	
	' Returns the image used to identify the key.
	Private Function GetImageIndex(ByVal productCode As Integer) As Short
		If productCode = -1 Then GetImageIndex = 1 : Exit Function ' invalid
		
		If KeyProducts.Exists(productCode) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyProducts.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetImageIndex = KeyProducts.Item(productCode)
		Else
			GetImageIndex = 2 ' unrecognized
		End If
	End Function
	
	
	Private Sub Local_LoadCDKeys()
		Dim keys As Collection
		Dim sKey As Object
		keys = ListFileLoad(GetFilePath(FILE_KEY_STORAGE))
		
		For	Each sKey In keys
			'UPGRADE_WARNING: Couldn't resolve default property of object sKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = CStr(Trim(sKey))
			'UPGRADE_WARNING: Couldn't resolve default property of object sKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Len(sKey) > 0 Then ProcessKey(sKey)
		Next sKey
	End Sub
	
	Private Sub Local_WriteCDKeys()
		Dim keys As Collection
		Dim Item As System.Windows.Forms.ListViewItem
		
		keys = New Collection
		
		For	Each Item In lvKeys.Items
			keys.Add(Item.Tag)
		Next Item
		
		ListFileSave(GetFilePath(FILE_KEY_STORAGE), keys)
		
		'UPGRADE_NOTE: Object keys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		keys = Nothing
	End Sub
End Class