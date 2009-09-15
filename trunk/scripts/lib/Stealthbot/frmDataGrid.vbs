'// Declare forms
Private frmDataGrid

'// Controls
Private lvDataGrid



'// Private variables
Private COLUMN_WIDTH
Private p_loaded, p_form_min_width, p_form_max_width, p_form_min_height, p_form_max_height		
Private p_doubleClickCallback, p_clickCallback

Set frmDataGrid = Nothing

'// Control Events

Public Sub frmDataGrid_Initialize()

	COLUMN_WIDTH = 3200
	
	p_form_min_width = 400
	p_form_max_width = 1200
	
	p_form_min_height = 200
	p_form_max_height = 850
	
	p_loaded = false

	'// create our objects
	Call frmDataGrid.CreateObj("ListView", "lvDataGrid")
	Set lvDataGrid = frmDataGrid.GetObjByName("lvDataGrid")
	
	'// Form properties
	With frmDataGrid
		.BackColor = vbBlack
		.Height = 334 * 16
		.Width = 800 * 16
		.Caption = "DataGrid - [52]"
	End With

	'// Control properties
	With lvDataGrid
		.Height = 23 * 16
		.Width = 23 * 16
		.Top = 27 * 16
		.Left = 149 * 16
		.View = 3
		.FullRowSelect = True
		.HideColumnHeaders = False
		.GridLines = True
	End With

	p_loaded = True
	
End Sub	

Public Sub frmDataGrid_Resize()

	'// make sure the form has completely loaded
	'// * This insures that all of the form objects have references before we try to access them.
	'// * The problem is that this event will be fired as soon as the form is created by the script,
	'// * but you are not able to add controls to the form before its created. Because of this, any
	'// * code we write to modify the controls will fail since they technically do not exist. YOU
	'// * MUST SET p_loaded TO True AFTER ALL CONTROLS HAVE BEEN ADDED BEFORE THIS EVENT WILL FIRE.
	If Not p_loaded Then Exit Sub		

	'// make sure the form minimums and maximums are in effect
	'// * this will automatically resize the form
	If checkFormSize() = False Then Exit Sub

	'// Any custom code for this event should only be below this line
	
	resizeDataGrid

	

End Sub

Public Sub frmDataGrid_lvDataGrid_Click()
	Dim listItem, i, cw
	If p_clickCallback <> "" Then
		Set listItem = lvDataGrid.SelectedItem
		cw = listItem.Text
		For i = 1 To listItem.ListSubItems.Count
			cw = cw & "|" & listItem.ListSubItems(i)
		Next
		cw = StringFormat("Call {0}(Split(""{1}"", ""|""))", p_clickCallback, cw)
		Execute cw
	End If
End Sub

Public Sub frmDataGrid_lvDataGrid_DblClick()
	Dim listItem, i, cw
	If p_doubleClickCallback <> "" Then
		Set listItem = lvDataGrid.SelectedItem
		cw = listItem.Text
		For i = 1 To listItem.ListSubItems.Count
			cw = cw & "|" & listItem.ListSubItems(i)
		Next
		cw = StringFormat("Call {0}(Split(""{1}"", ""|""))", p_doubleClickCallback, cw)
		Execute cw
	End If
End Sub

'// Public Methods

Public Sub LoadRecordSet(ByRef recordSet)

	Call prepareForm()


	Dim oRS, z, i, listItem
	Set oRS = recordSet
	
	With lvDataGrid
	
		For i = 0 To oRS.Fields.Count - 1
		
			With .ColumnHeaders
				.Add , , oRS.Fields(i).Name, COLUMN_WIDTH, 0
			End With
			
		Next
		z = 1
		Do While Not oRS.EOF
		
			Set listItem = .ListItems.Add( , , z)
			
			For i = 0 To oRS.Fields.Count - 1
				On Error Resume Next
				listItem.SubItems(i + 1) = oRS(i)
				If Err.Number <> 0 Then
					listItem.SubItems(i + 1) = "[NULL]"
				End If
				On Error Goto 0
			Next
		
			oRS.MoveNext
			z = z + 1
		
		Loop
	
	End With

End Sub

Public Sub LoadDelimitedText(text, fieldDelimiter, lineDelimiter, headerText) 

	Call prepareForm()
	Dim z, i, listItem, aFields, aLines
	
	With lvDataGrid
	
		aLines = Split(text, lineDelimiter)
		
		With .ColumnHeaders
			If headerText <> "" Then
				aFields = Split(headerText, fieldDelimiter)
				For i = 0 To UBound(aFields)
					.Add , , aFields(i), COLUMN_WIDTH, 0
				Next
			Else
				aFields = Split(aLines(0), fieldDelimiter)
				For i = 0 To UBound(aFields)
					.Add , , "Column " & (i + 1), COLUMN_WIDTH, 0
				Next
			End If
		End With
	
		For z = 0 To UBound(aLines)
			aFields = Split(aLines(z), fieldDelimiter)
			Set listItem = .ListItems.Add( , , z + 1)
			For i = 0 To UBound(aFields) 
				listItem.SubItems(i + 1) = aFields(i)
			Next
		Next
		
	End With

End Sub

Public Sub LoadDictionary(ByRef dictionary)

	Call prepareForm()

	Dim z, i, listItem, dictionaryKeys
	
	dictionaryKeys = dictionary.Keys
	With lvDataGrid
		With .ColumnHeaders
			.Add , , "Key", COLUMN_WIDTH, 0
			.Add , , "Value", COLUMN_WIDTH, 0
		End With
		For z = 1 To UBound(dictionaryKeys) + 1
			Set listItem = .ListItems.Add( , , z)
			listItem.SubItems(1) = dictionaryKeys(z - 1)
			listItem.SubItems(2) = dictionary.Item(dictionaryKeys(z - 1))
		Next
	End With
End Sub

Public Function GetHeaders()
	If Not p_loaded Then
		GetHeaders = Array()
		Exit Function
	End If
	
	Dim i, s
	With lvDataGrid.ColumnHeaders
		s = .Item(1)
		For i = 2 To .Count
			s = s & "|" & .Item(i)
		Next
	End With

	GetHeaders = Split(s, "|")
End Function

Public Sub SetDoubleClickCallback(cb)
	p_doubleClickCallback = cb
End Sub

Public Sub SetClickCallback(cb)
	p_clickCallBack = cb
End Sub

'// Private Methods

Private Sub resizeDataGrid
	
	If Not p_loaded Then Exit Sub

	With lvDataGrid
		.Top = 5 * 16
		.Left = 5 * 16
		'.Height = frmDataGrid.ClientHeight - (10 * 16) '// <-- this one should work!!
		.Height = frmDataGrid.Height - (10 * 16) - (32 * 16)
		.Width = frmDataGrid.Width - (10 * 16) - (7 * 16)
	End With


End Sub		

Private Sub prepareForm()

	'// Create forms
	If (frmDataGrid Is Nothing) = True Then
		Call CreateObj("Form", "frmDataGrid")
	End If
	
	Call frmDataGrid.Show()
	
	With lvDataGrid
		.ListItems.Clear
		.ColumnHeaders.Clear		
		.ColumnHeaders.Add , , "Row", 640, 0
	End With
	
End Sub

'// makes sure the form conforms to the minumum and maximum form sizes.
'// * This will always return false when the form is minimized or maximized
Private Function checkFormSize()
	
	Dim bExit
	
	'// check if the window state is normal
	If frmDataGrid.WindowState = 1 Or frmDataGrid.WindowState = 2 Then	
		checkFormSize = False
		Exit Function
	End If

	'// Handle minimum form width
	
	If frmDataGrid.Width < (p_form_min_width * 16) Then
		frmDataGrid.Width = (p_form_min_width * 16)
		bExit = True
	End If

	'// Handle maximum form width
	If frmDataGrid.Width > (p_form_max_width * 16) Then
		frmDataGrid.Width = (p_form_max_width * 16)
		bExit = True
	End If

	'// Handle minimum form height
	If frmDataGrid.Height < (p_form_min_height * 16) Then
		frmDataGrid.Height = (p_form_min_height * 16)
		bExit = True
	End If

	'// Handle maximum form height
	If frmDataGrid.Height > (p_form_max_height * 16) Then
		frmDataGrid.Height = (p_form_max_height * 16)
		bExit = True
	End If
	
	If bExit = True Then
		checkFormSize = False
		Exit Function
	End If
	
	'// all good in the hood
	checkFormSize = True
	
End Function

