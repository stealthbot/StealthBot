
'// form
Private frmScriptUpdate

'// Controls
Private txtSearch, lvScripts, btnSearch, ddCategories, lblHelp, ilCategories, chkShowLibraries

Sub frmScriptUpdate_Initialize()

	'// Create our form controls
	Call frmScriptUpdate.CreateObj("TextBox", "txtSearch")
	Call frmScriptUpdate.CreateObj("ListView", "lvScripts")
	Call frmScriptUpdate.CreateObj("Button", "btnSearch")
	Call frmScriptUpdate.CreateObj("ComboBox", "ddCategories")
	Call frmScriptUpdate.CreateObj("Label", "lblHelp")
	Call frmScriptUpdate.CreateObj("ImageList", "ilCategories")
	Call frmScriptUpdate.CreateObj("CheckBox", "chkShowLibraries")
	
	Set txtSearch = frmScriptUpdate.GetObjByName("txtSearch")
	Set lvScripts = frmScriptUpdate.GetObjByName("lvScripts")
	Set btnSearch = frmScriptUpdate.GetObjByName("btnSearch")
	Set ddCategories = frmScriptUpdate.GetObjByName("ddCategories")
	Set lblHelp = frmScriptUpdate.GetObjByName("lblHelp")
	Set ilCategories = frmScriptUpdate.GetObjByName("ilCategories")
	Set chkShowLibraries = frmScriptUpdate.GetObjByName("chkShowLibraries")
	
	'// Form properties
	With frmScriptUpdate
		.Caption = "Online Script Repository"
		.Width = 800 * 16
		.Height = 597 * 16
	End With

	'// Control properties
	
	With ddCategories
		.Top = 5 * 16
		.Left = 5 * 16
		.Width = 160 * 16
		With .Font
			.Size = 10
		End With
	End With
	
	With txtSearch
		.Top = 5 * 16
		.Left = ddCategories.Left + ddCategories.Width + (5 * 16)
		.Width = 180 * 16
		.Height = 22 * 16
		With .Font
			.Size = 10
		End With
	End With

	With btnSearch
		.Top = 5 * 16
		.Left = txtSearch.Left + txtSearch.Width + (15 * 16)
		.Width = 60 * 16
		.Height = 22 * 16
		.Caption = "Search"
	End With
	
	With ilCategories
		.ImageHeight = 16
		.ImageWidth = 16
	End With

	With lvScripts
		.Top = 60 * 16
		.Left = 4 * 16
		.HideColumnHeaders = False
		.FullRowSelect = True
		.Sorted = True
		.GridLines = True
		'// alignment
		'// 0 : Left
		'// 1 : Right
		'// 2 : Center
		With .ColumnHeaders
			.Add , , "Script Name", 150 * 16, 0
			.Add , , "Category", 120 * 16, 0
			.Add , , "Version", 50 * 16, 1
			.Add , , "Author(s)", 80 * 16, 0
			.Add , , "Avg./Total Rating", 100 * 16, 1
			.Add , , "Description", 270 * 16, 0
		End With
		
	End With

	With lblHelp
		.AutoSize = False
		.Top = 8 * 16
		.Left = btnSearch.Left + btnSearch.Width + (18 * 16)
		.Width = 360 * 16
		.Height = 40 * 16
		With .Font
			.Size = 10
		End With
		.ForeColor = vbWhite
		.Caption = "Search the online script repository using the form to the left. Double click a script to open the script details."
	End With
	
	With chkShowLibraries
		.Caption = "Show script libraries?"
		.ToolTipText = "Scripting libraries that other scripts depend on will automatically download when required."
		.Top = 31 * 16
		.Left = 14 * 16
		.Height = 20 * 16
		.Width = 150 * 16
		With .Font
			.Size = 10
		End With
		.BackColor = vbBlack
		.ForeColor = vbWhite
		.Alignment = 1
	End With
	
	Call bind_ddCategories()
	
End Sub	

Sub frmScriptUpdate_Activate()
	lvScripts.SmallIcons = ilCategories
End Sub

Sub frmScriptUpdate_Unload(cancel)
	Set lvScripts.SmallIcons = Nothing
	lvScripts.ListItems.Clear
End Sub

Sub frmScriptUpdate_Resize()

	With lvScripts
		'.Height = Updater_FO.frmUpdater.ClientHeight - (10 * 16) '// <-- this one should work!!
		.Height = frmScriptUpdate.Height - (10 * 16) - (32 * 16) - (60 * 16) + (4 * 16)
		.Width = frmScriptUpdate.Width - (10 * 16) - (8 * 16)
	End With

End Sub

Sub bind_ddCategories()

	Dim xmlResponse, categoryNodeList, categoryNode
	
	lvScripts.ListItems.Clear
	
	Set xmlResponse = Updater.FetchCategories()
	Set categoryNodeList = xmlResponse.SelectNodes("/ScriptUpdate/Response/Categories/Category")
	
	'// Add the 'Search All Categories' item
	With ddCategories
		Call .AddItem("Search All Categories")
		.ItemData(.NewIndex) = -1
	End With
	
	For Each categoryNode In categoryNodeList
		'// bind the dropdown
		With ddCategories
			Call .AddItem(categoryNode.SelectSingleNode("@Name").Text)
			.ItemData(.NewIndex) = categoryNode.SelectSingleNode("@CategoryID").Text
		End With
		'// add the associated image
		
		With ilCategories.ListImages
			Call .Add(,categoryNode.SelectSingleNode("@LookupCode").Text,_
					   LoadPicture(StringFormat("{0}images\{1}", GetWorkingDirectory(), categoryNode.SelectSingleNode("@ImagePath").Text)))
		End With
	Next
	
	ddCategories.Text = "Search All Categories"
	
	Set categoryNode = Nothing
	Set categoryNodeList = Nothing
	Set xmlResponse = Nothing

End Sub

Sub bind_lvScripts(sbscriptNodeList)
	
	Dim sbscriptNode, authorNode, lvi, authors
	
	lvScripts.ListItems.Clear
	
	For Each sbscriptNode In sbscriptNodeList
		With lvScripts
			'// Name
			Set lvi = .ListItems.Add( , , sbscriptNode.SelectSingleNode("@Name").Text)
			'// Icon
			If sbscriptNode.SelectSingleNode("@CategoryLookupCode").Text <> "" Then
				lvi.SmallIcon = sbscriptNode.SelectSingleNode("@CategoryLookupCode").Text
			End If
			'// SBScriptID
			lvi.Tag = sbscriptNode.SelectSingleNode("@SBScriptID").Text
			'// Category
			lvi.SubItems(1) = sbscriptNode.SelectSingleNode("@CategoryName").Text
			'// Version
			lvi.SubItems(2) = sbscriptNode.SelectSingleNode("@Version").Text
			'// Author
			authors = ""
			For Each authorNode In sbscriptNode.SelectNodes("Authors/Author")
				authors = authors & authorNode.SelectSingleNode("@Name").Text & " "
			Next
			lvi.SubItems(3) = authors
			'// Avg./Total Rating
			lvi.SubItems(4) = StringFormat("{0} / {1}", sbscriptNode.SelectSingleNode("@AverageRating").Text, _
													    sbscriptNode.SelectSingleNode("@TotalRatings").Text)
			'// Description
			lvi.SubItems(5) = sbscriptNode.SelectSingleNode("Description").Text
		End With		
	Next
	
	Set sbscriptNode = Nothing

End Sub

Sub frmScriptUpdate_btnSearch_Click()

	
	Dim xmlResponse
	Dim categoryID, searchText, plural, showLibraries
	
	frmScriptUpdate.Caption = "Online Script Repository - Searching...."

	searchText = txtSearch.Text
	With ddCategories
		If .ListIndex <> -1 Then
			categoryID = .ItemData(.ListIndex)
		Else
			categoryID = 0
		End If
	End With
	
	showLibraries = chkShowLibraries.Value
	
	Set xmlResponse = Updater.SearchScripts(categoryID, searchText, showLibraries)
	Call bind_lvScripts(xmlResponse.SelectNodes("/ScriptUpdate/Response/SBScripts/SBScript"))
	
	plural = ""
	If xmlResponse.SelectNodes("/ScriptUpdate/Response/SBScripts/SBScript").Length > 1 Then
		plural = "s"
	End If
		
	frmScriptUpdate.Caption = StringFormat("Online Script Repository - Searching Complete - {0} script{1} found.", xmlResponse.SelectNodes("/ScriptUpdate/Response/SBScripts/SBScript").Length, plural)

	'// clean up
	Set xmlResponse = Nothing
	
End Sub

Sub frmScriptUpdate_lvScripts_DblClick()
	Call Updater.FetchScriptAndDependencies(lvScripts.SelectedItem.Tag)
End Sub

Sub frmScriptUpdate_lvScripts_ColumnClick(ByVal column)

	If lvScripts.SortKey <> column.Index Then
		lvScripts.SortKey = column.Index
		lvScripts.SortOrder = 0
	Else
		lvScripts.SortOrder = lvScripts.SortOrder Xor 1
	End If

End Sub