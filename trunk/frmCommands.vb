Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmCommands
	Inherits System.Windows.Forms.Form
	
	Private m_Commands As clsCommandDocObj
	'Private m_CommandsDoc As DOMDocument60
	Private m_SelectedElement As SelectedElement
	Private m_ClearingNodes As Boolean
	
	'// Enums
	Private Enum NodeType
		nCommand = 0
		nArgument = 1
		nRestriction = 2
	End Enum
	
	'// Stores information about the selected node in the treeview
	Private Structure SelectedElement
		Dim TheNodeType As NodeType
		Dim IsDirty As Boolean
		Dim commandName As String
		Dim argumentName As String
		Dim restrictionName As String
	End Structure
	
	Sub ClearTreeViewNodes(ByRef trv As AxvbalTreeViewLib6.AxvbalTreeView)
		
		m_ClearingNodes = True
		trv.nodes.Clear()
		m_ClearingNodes = False
		
	End Sub
	
	
	'UPGRADE_WARNING: Event cboCommandGroup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCommandGroup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCommandGroup.SelectedIndexChanged
		
		If PromptToSaveChanges() = True Then
			
			Call ResetForm()
			Call PopulateTreeView(getScriptOwner())
			
			'// getScriptOwner() contains this logic already
			'If cboCommandGroup.ListIndex = 0 Then
			'    Call PopulateTreeView
			'Else
			'    Call PopulateTreeView(getScriptOwner())
			'End If
		End If
		
	End Sub
	
	Private Sub cmdFlagAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFlagAdd.Click
		
		cboFlags.Items.Add(cboFlags.Text)
		cboFlags.Text = ""
		Call FormIsDirty()
		
	End Sub
	
	
	Private Sub cmdFlagRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFlagRemove.Click
		
		Dim i As Short
		
		For i = 0 To cboFlags.Items.Count - 1
			If (StrComp(cboFlags.Text, VB6.GetItemString(cboFlags, i), CompareMethod.Binary) = 0) Then
				cboFlags.Items.RemoveAt(i)
				Exit For
			End If
		Next i
		
		cboFlags.Text = ""
		Call FormIsDirty()
		
	End Sub
	
	Private Sub cmdAliasAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAliasAdd.Click
		
		cboAlias.Items.Add(cboAlias.Text)
		cboAlias.Text = ""
		Call FormIsDirty()
		
	End Sub
	
	Private Sub cmdAliasRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAliasRemove.Click
		
		Dim i As Short
		
		For i = 0 To cboAlias.Items.Count - 1
			If (StrComp(cboAlias.Text, VB6.GetItemString(cboAlias, i), CompareMethod.Text) = 0) Then
				cboAlias.Items.RemoveAt(i)
				Exit For
			End If
		Next i
		
		cboAlias.Text = ""
		Call FormIsDirty()
		
	End Sub
	
	
	'// 08/30/2008 JSM - Created
	Private Sub cmdDiscard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDiscard.Click
		
		Call PrepareForm(m_SelectedElement.TheNodeType)
		
	End Sub
	
	
	
	Private Sub cmdDeleteCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDeleteCommand.Click
		
		Dim scriptName As String
		Dim scriptIndex As Short
		
		scriptName = Mid(cboCommandGroup.Text, 1, InStr(1, cboCommandGroup.Text, "(") - 2)
		
		scriptIndex = cboCommandGroup.SelectedIndex
		
		If MsgBoxResult.Yes <> MsgBox(StringFormat("Are you sure you want to delete the {0} command for the {1} script?", m_SelectedElement.commandName, scriptName), MsgBoxStyle.YesNo + MsgBoxStyle.Question, Me.Text) Then
			Exit Sub
		End If
		
		Call m_Commands.OpenCommand(m_SelectedElement.commandName, scriptName)
		Call m_Commands.Delete()
		
		m_SelectedElement.IsDirty = False
		
		Call PopulateOwnerComboBox()
		Call ResetForm()
		Call PopulateTreeView(scriptName, scriptIndex)
		
	End Sub
	
	
	'// 08/30/2008 JSM - Created
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Call SaveForm()
	End Sub
	
	Private Sub frmCommands_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error GoTo ErrorHandler
		
		Dim colErrorList As Collection
		
		'// Load commands.xml
		m_Commands = New clsCommandDocObj
		
		'If Not clsCommandDocObj.ValidateXMLFromFiles(App.Path & "\commands.xml", App.Path & "\commands.xsd") Then
		'    Exit Sub
		'End If
		
		'Call m_CommandsDoc.Load(App.Path & "\commands.xml")
		
		'If Not clsCommandDocObj.CommandsSanityCheck(m_CommandsDoc) Then
		'    Exit Sub
		'End If
		
		Call ResetForm()
		Call PopulateOwnerComboBox()
		Call PopulateTreeView()
		
		Exit Sub
		
ErrorHandler: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, Err.Description)
		Call ResetForm()
		'// Disable our buttons
		cmdSave.Enabled = False
		cmdDiscard.Enabled = False
		Exit Sub
		
	End Sub
	
	Private Sub frmCommands_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object m_Commands may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_Commands = Nothing
	End Sub
	
	
	Private Sub PopulateOwnerComboBox()
		
		On Error Resume Next
		
		Dim i As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim commandCount As Short
		Dim scriptName As String
		
		cboCommandGroup.Items.Clear()
		
		'// get the script name and number of commands
		scriptName = "Internal Bot Commands"
		commandCount = m_Commands.GetCommandCount()
		
		'// add the item
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboCommandGroup.Items.Add(StringFormat("{0} ({1})", scriptName, commandCount))
		
		For i = 2 To frmChat.SControl.Modules.Count
			scriptName = modScripting.GetScriptName(CStr(i))
			str_Renamed = SharedScriptSupport.GetSettingsEntry("Public", scriptName)
			
			If (StrComp(str_Renamed, "False", CompareMethod.Text) <> 0) Then
				'// get the script name and number of commands
				commandCount = m_Commands.GetCommandCount(scriptName)
				'// only add the commands if there is at least 1 command to show
				If commandCount > 0 Then
					'// add the item
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cboCommandGroup.Items.Add(StringFormat("{0} ({1})", scriptName, commandCount))
				End If
			End If
		Next i
		cboCommandGroup.SelectedIndex = 0
		
	End Sub
	
	'UPGRADE_NOTE: clsCommandObj was upgraded to clsCommandObj_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub PopulateTreeView(Optional ByRef strScriptOwner As String = vbNullString, Optional ByRef intScriptIndex As Short = -1)
		Dim clsCommandObj_Renamed As Object
		
		Dim commandNodes As MSXML2.IXMLDOMNodeList
		Dim totalCommands As Short
		Dim commandNameArray() As Object
		
		Dim xpath As String
		
		Dim xmlCommand As MSXML2.IXMLDOMNode
		Dim xmlArgs As MSXML2.IXMLDOMNodeList
		Dim xmlArgRestricions As MSXML2.IXMLDOMNodeList
		
		Dim nCommand As vbalTreeViewLib6.cTreeViewNode
		Dim nArg As vbalTreeViewLib6.cTreeViewNode
		Dim nArgRestriction As vbalTreeViewLib6.cTreeViewNode
		
		Dim commandName As String
		Dim argumentName As String
		Dim restrictionName As String
		
		'// 08/30/2008 JSM - used to get the first command alphabetically
		Dim defaultNode As vbalTreeViewLib6.cTreeViewNode
		
		'// Counters
		Dim j As Short
		Dim i As Short
		Dim x As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strScriptOwner = clsCommandObj.CleanXPathVar(strScriptOwner)
		
		'// reset the treeview
		If trvCommands.nodes.Count > 0 Then
			trvCommands.nodes(1).Selected = True
		End If
		
		Call ClearTreeViewNodes(trvCommands)
		
		'// create xpath expression based on strScriptOwner
        If Len(strScriptOwner) = 0 Then
            xpath = "/commands/command[not(@owner)]"
            'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xpath = StringFormat("/commands/command[@owner='{0}']", strScriptOwner)
            'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
        End If
		
		'// get a list of all the commands
		commandNodes = m_Commands.XMLDocument.documentElement.selectNodes(xpath)
		ReDim commandNameArray(commandNodes.length)
		
		'// read them 1 at a time and add them to an array
		x = 0
		For	Each xmlCommand In m_Commands.XMLDocument.documentElement.selectNodes(xpath)
			'UPGRADE_WARNING: Couldn't resolve default property of object commandNameArray(x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			commandNameArray(x) = xmlCommand.Attributes.getNamedItem("name").Text
			x = x + 1
		Next xmlCommand
		
		'// sort the command names
		Call BubbleSort1(commandNameArray)
		
		'// loop through the sorted array and select the commands
		For x = LBound(commandNameArray) To UBound(commandNameArray)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object commandNameArray(x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			commandName = commandNameArray(x)
			'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandObj.CleanXPathVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			commandName = clsCommandObj.CleanXPathVar(commandName)
            If Len(commandName) > 0 Then
                '// create xpath expression based on strScriptOwner
                If Len(strScriptOwner) = 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xpath = StringFormat("/commands/command[@name='{0}' and not(@owner)]", commandName)
                    'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xpath = StringFormat("/commands/command[@name='{0}' and @owner='{1}']", commandName, strScriptOwner)
                    'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
                End If

                xmlCommand = m_Commands.XMLDocument.documentElement.selectSingleNode(xpath)

                commandName = xmlCommand.attributes.getNamedItem("name").text
                nCommand = trvCommands.Nodes.Add(trvCommands.Nodes.Parent, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, commandName, commandName)

                '// 08/30/2008 JSM - check if this command is the first alphabetically
                If defaultNode Is Nothing Then
                    defaultNode = nCommand
                Else
                    If StrComp(defaultNode.Text, nCommand.Text) > 0 Then
                        defaultNode = nCommand
                    End If
                End If

                xmlArgs = xmlCommand.selectNodes("arguments/argument")
                '// 08/29/2008 JSM - removed 'Not (xmlArgs Is Nothing)' condition. xmlArgs will always be
                '//                  something, even if nothing matches the XPath expression.
                For i = 0 To (xmlArgs.length - 1)

                    argumentName = xmlArgs(i).attributes.getNamedItem("name").text
                    If (Not xmlArgs(i).attributes.getNamedItem("optional") Is Nothing) Then
                        If (xmlArgs(i).attributes.getNamedItem("optional").text = "1") Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            argumentName = StringFormat("[{0}]", argumentName)
                        End If
                    End If

                    '// Add the datatype to the argument name
                    If (Not xmlArgs(i).attributes.getNamedItem("type") Is Nothing) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        argumentName = StringFormat("{0} ({1})", argumentName, xmlArgs(i).attributes.getNamedItem("type").text)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        argumentName = StringFormat("{0} ({1})", argumentName, "String")
                    End If

                    nArg = trvCommands.Nodes.Add(nCommand, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, commandName & "." & argumentName, argumentName)

                    xmlArgRestricions = xmlArgs(i).selectNodes("restrictions/restriction")

                    For j = 0 To (xmlArgRestricions.length - 1)
                        restrictionName = xmlArgRestricions(j).attributes.getNamedItem("name").text
                        nArgRestriction = trvCommands.Nodes.Add(nArg, vbalTreeViewLib6.ETreeViewRelationshipContants.etvwChild, commandName & "." & argumentName & "." & restrictionName, restrictionName)
                    Next j
                Next i
            End If '// Len(commandName) > 0
		Next x
		
		'// 08/30/2008 JSM - click the first command alphabetically
		' fixed to work with SelectedNodeChanged() -Ribose/2009-08-10
		If Not (defaultNode Is Nothing) Then
			defaultNode.Selected = True
		Else
			trvCommands_SelectedNodeChanged(trvCommands, New System.EventArgs())
		End If
		
		If (intScriptIndex >= 0) Then
			cboCommandGroup.SelectedIndex = intScriptIndex
		End If
		
	End Sub
	
	
	'// This function will prompt a user to save the changes (if necessary)
	'// 08/30/2008 JSM - Created
	Private Function PromptToSaveChanges() As Boolean
		
		Dim sMessage As String
		
		'// If the current form is dirty, lets show a save dialog
		With m_SelectedElement
			If Not .IsDirty Then
				PromptToSaveChanges = True
				Exit Function
			End If
			
			'// Get the message for the prompt
			Select Case .TheNodeType
				Case NodeType.nCommand
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sMessage = StringFormat("You have not saved your changes to {0}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
				Case NodeType.nArgument
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sMessage = StringFormat("You have not saved your changes to {1}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
				Case NodeType.nRestriction
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sMessage = StringFormat("You have not saved your changes to {2}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
			End Select
			
			'// Get the user response
			Select Case MsgBox(sMessage, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, Me.Text)
				Case MsgBoxResult.Yes
					Call SaveForm()
					PromptToSaveChanges = True
					Exit Function
				Case MsgBoxResult.No
					PromptToSaveChanges = True
					Exit Function
				Case MsgBoxResult.Cancel
					PromptToSaveChanges = False
					Exit Function
			End Select
			
		End With
		
		
	End Function
	
	'// 08/28/2008 JSM - Created
	' moved to _SelectedNodeChanged -Ribose
	' if no node is selected (such as none existing), now disables all fields -Ribose/2009-08-10
	Private Sub trvCommands_SelectedNodeChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trvCommands.SelectedNodeChanged
		
		On Error GoTo ErrorHandler
		
		Dim node As vbalTreeViewLib6.cTreeViewNode
		Dim nt As NodeType
		Dim commandName As String
		Dim argumentName As String
		Dim restrictionName As String
		
		Dim xpath As String
		
		If m_ClearingNodes Then Exit Sub
		
		node = trvCommands.SelectedItem
		If node Is Nothing Then
			Call ResetForm()
			Exit Sub
		End If
		
		'// This function will prompt the user to save changes if necessary. If the
		'// return value is false, then the use clicked cancel so we should gtfo of here.
		If PromptToSaveChanges() = False Then
			Exit Sub
		End If
		
		'// figure out what type of node was clicked on
		nt = GetNodeInfo(node, commandName, argumentName, restrictionName)
		
		'// Update m_SelectedElement so we know which element we are viewing
		m_SelectedElement.commandName = commandName
		m_SelectedElement.argumentName = argumentName
		m_SelectedElement.restrictionName = restrictionName
		
		'// load the command and set up the form
		Call m_Commands.OpenCommand(commandName, IIf(cboCommandGroup.SelectedIndex = 0, vbNullString, cboCommandGroup.Text))
		Call ResetForm()
		Call PrepareForm(nt)
		
		Exit Sub
		
ErrorHandler: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, Err.Description)
		Call ResetForm()
		'// Disable our buttons
		cmdSave.Enabled = False
		cmdDiscard.Enabled = False
		Exit Sub
		
	End Sub
	
	'// Call this sub whenever the form controls have been changed
	'// 08/30/2008 JSM - Created
	Private Sub FormIsDirty()
		m_SelectedElement.IsDirty = True
		cmdSave.Enabled = True
		cmdDiscard.Enabled = True
	End Sub
	
	'// Checks the hiarchy of the treenodes to determine what type of node it is.
	'// 08/29/2008 JSM - Created
	Private Function GetNodeInfo(ByRef node As vbalTreeViewLib6.cTreeViewNode, ByRef commandName As String, ByRef argumentName As String, ByRef restrictionName As String) As NodeType
		Dim s() As String
		
        If Len(node.Key) > 0 Then
            s = Split(node.Key, ".")
            Select Case UBound(s)
                Case 0
                    commandName = s(0)
                    argumentName = vbNullString
                    restrictionName = vbNullString
                    GetNodeInfo = NodeType.nCommand
                Case 1
                    commandName = s(0)
                    argumentName = s(1)
                    restrictionName = vbNullString
                    GetNodeInfo = NodeType.nArgument
                Case 2
                    commandName = s(0)
                    argumentName = s(1)
                    restrictionName = s(2)
                    GetNodeInfo = NodeType.nRestriction
            End Select
            '// strip the [ ] around optional parameters
            If VB.Left(argumentName, 1) = "[" Then
                argumentName = Mid(argumentName, 2, InStr(1, argumentName, "]") - 2)
            End If
            If (InStr(1, argumentName, "(") >= 0 And VB.Right(argumentName, 1) = ")") Then
                argumentName = Mid(argumentName, 1, InStr(1, argumentName, "(") - 2)
            End If
        End If
	End Function
	
	Private Function getFlags() As Object
		
		Dim i As Short
		Dim sTmp As String
		
		sTmp = ""
		For i = 0 To cboFlags.Items.Count - 1
			sTmp = sTmp & VB6.GetItemString(cboFlags, i)
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object getFlags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getFlags = sTmp
		
	End Function
	
	
	'// Saves the selected treeview node in the commands.xml
	'// 08/30/2008 JSM - Created
	Private Sub SaveForm()
		
		Dim parameter As clsCommandParamsObj
		Dim Restriction As clsCommandRestrictionObj
		
		Dim i As Object
		
		Select Case m_SelectedElement.TheNodeType
			Case NodeType.nCommand
				'// saving the command
				With m_Commands
					.description = txtDescription.Text
					.SpecialNotes = txtSpecialNotes.Text
					.RequiredRank = CShort(txtRank.Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object getFlags(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.RequiredFlags = getFlags()
					While .aliases.Count() <> 0
						.aliases.Remove(1)
					End While
					For i = 0 To cboAlias.Items.Count - 1
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.aliases.Add(VB6.GetItemString(cboAlias, i))
					Next i
					.IsEnabled = Not CBool(chkDisable.CheckState)
				End With
				
			Case NodeType.nArgument
				'// saving the parameter
				With m_Commands
					parameter = .GetParameterByName(m_SelectedElement.argumentName)
					With parameter
						.description = txtDescription.Text
						.SpecialNotes = txtSpecialNotes.Text
					End With
				End With
				
			Case NodeType.nRestriction
				'// saving the restriction
				With m_Commands
					parameter = m_Commands.GetParameterByName(m_SelectedElement.argumentName)
					With parameter
						Restriction = parameter.GetRestrictionByName(m_SelectedElement.restrictionName)
						With Restriction
							.RequiredRank = CShort(txtRank.Text)
							'UPGRADE_WARNING: Couldn't resolve default property of object getFlags(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.RequiredFlags = getFlags()
						End With
					End With
				End With
				
		End Select
		
		Call m_Commands.Save()
		Call ResetForm()
		Call PrepareForm(m_SelectedElement.TheNodeType)
		
	End Sub
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function PrepString(ByVal str_Renamed As String) As Object
		Dim retVal As String
		retVal = str_Renamed
		retVal = Replace(retVal, vbLf, vbCrLf)
		retVal = Replace(retVal, vbTab, vbNullString)
		'UPGRADE_WARNING: Couldn't resolve default property of object PrepString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrepString = retVal
	End Function
	
	Function getScriptOwner() As String
		
		'// return vbNullstring if internal commands is selected, otherwise the script name
		If cboCommandGroup.SelectedIndex = 0 Then
			getScriptOwner = vbNullString
		Else
			getScriptOwner = Mid(cboCommandGroup.Text, 1, InStr(1, cboCommandGroup.Text, "(") - 2)
		End If
		
	End Function
	
	
	
	'// When a node in the treeview is clicked, it should locate the XML element that was
	'// used to create the node and call this method to populate appropriate form controls.
	'// 08/29/2008 JSM - Created
	Private Sub PrepareForm(ByRef nt As NodeType)
		
		Dim parameter As clsCommandParamsObj
		Dim Restriction As clsCommandRestrictionObj
		
		Dim requirements As String
		Dim sItem As String
		Dim i As Short
		
		With m_SelectedElement
			
			Call m_Commands.OpenCommand(.commandName, getScriptOwner())
			
			'// lblSyntax
			lblSyntax.Text = m_Commands.SyntaxString
			'// lblRequirements
			lblRequirements.Text = m_Commands.RequirementsString
			
			Select Case nt
				Case NodeType.nCommand
					'// txtRank
					txtRank.Enabled = True
					lblRank.Enabled = True
					txtRank.Text = CStr(m_Commands.RequiredRank)
					'// cboAlias
					cboAlias.Enabled = True
					lblAlias.Enabled = True
					cmdAliasAdd.Enabled = True
					cmdAliasRemove.Enabled = True
					For i = 1 To m_Commands.aliases.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object m_Commands.aliases(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cboAlias.Items.Add(m_Commands.aliases.Item(i))
					Next i
					'// cboFlags
					cboFlags.Enabled = True
					lblFlags.Enabled = True
					cmdFlagAdd.Enabled = True
					cmdFlagRemove.Enabled = True
					For i = 1 To Len(m_Commands.RequiredFlags)
						cboFlags.Items.Add(Mid(m_Commands.RequiredFlags, i, 1))
					Next i
					If (cboFlags.Items.Count) Then
						cboFlags.Text = VB6.GetItemString(cboFlags, 0)
					End If
					'// txtDescription
					txtDescription.Enabled = True
					lblDescription.Enabled = True
					'UPGRADE_WARNING: Couldn't resolve default property of object PrepString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txtDescription.Text = PrepString(m_Commands.description)
					'// txtSpecialNotes
					txtSpecialNotes.Enabled = True
					lblSpecialNotes.Enabled = True
					'UPGRADE_WARNING: Couldn't resolve default property of object PrepString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txtSpecialNotes.Text = PrepString(m_Commands.SpecialNotes)
					'// chkDisable
					chkDisable.Enabled = True
					chkDisable.Visible = True
					chkDisable.CheckState = IIf(m_Commands.IsEnabled, System.Windows.Forms.CheckState.Unchecked, System.Windows.Forms.CheckState.Checked)
					'// custom captions
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fraCommand.Text = StringFormat("{0}", .commandName, .argumentName, .restrictionName)
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chkDisable.Text = StringFormat("Disable {0} command", .commandName, .argumentName, .restrictionName)
					'// only allow deleting script commands
					If cboCommandGroup.SelectedIndex > 0 Then
						cmdDeleteCommand.Enabled = True
					End If
					
				Case NodeType.nArgument
					parameter = m_Commands.GetParameterByName(.argumentName)
					
					'// txtDescription
					txtDescription.Enabled = True
					lblDescription.Enabled = True
					'UPGRADE_WARNING: Couldn't resolve default property of object PrepString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txtDescription.Text = PrepString(parameter.description)
					'// txtSpecialNotes
					txtSpecialNotes.Enabled = True
					lblSpecialNotes.Enabled = True
					'UPGRADE_WARNING: Couldn't resolve default property of object PrepString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txtSpecialNotes.Text = PrepString(parameter.SpecialNotes)
					'// custom captions
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fraCommand.Text = StringFormat("{0} => {1}{3}", .commandName, .argumentName, .restrictionName, IIf(parameter.IsOptional, " - Optional", ""))
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chkDisable.Text = StringFormat("Disable {1} argument", .commandName, .argumentName, .restrictionName)
					
				Case NodeType.nRestriction
					
					parameter = m_Commands.GetParameterByName(.argumentName)
					Restriction = parameter.GetRestrictionByName(.restrictionName)
					
					'// txtRank
					txtRank.Enabled = True
					lblRank.Enabled = True
					txtRank.Text = CStr(Restriction.RequiredRank)
					'// cboFlags
					cboFlags.Enabled = True
					lblFlags.Enabled = True
					For i = 1 To Len(Restriction.RequiredFlags)
						cboFlags.Items.Add(Mid(Restriction.RequiredFlags, i, 1))
					Next i
					If (cboFlags.Items.Count) Then
						cboFlags.Text = VB6.GetItemString(cboFlags, 0)
					End If
					'// special captions
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fraCommand.Text = StringFormat("{0} => {1}{3} => {2}", .commandName, .argumentName, .restrictionName, IIf(parameter.IsOptional, " - Optional", ""))
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chkDisable.Text = StringFormat("Disable {2} restriction", .commandName, .argumentName, .restrictionName)
			End Select
		End With
		
		'// Update m_SelectedElement so we know which element we are viewing
		m_SelectedElement.TheNodeType = nt
		m_SelectedElement.IsDirty = False
		
		'// Disable our buttons
		cmdSave.Enabled = False
		cmdDiscard.Enabled = False
		
	End Sub
	
	'// Clears and disables all edit controls. Treeview is left intact
	'// 08/29/2008 JSM - Created
	Private Sub ResetForm()
		
		txtRank.Text = vbNullString
		cboAlias.Items.Clear()
		cboFlags.Items.Clear()
		txtDescription.Text = vbNullString
		txtSpecialNotes.Text = vbNullString
		fraCommand.Text = vbNullString
		chkDisable.CheckState = System.Windows.Forms.CheckState.Unchecked
		
		txtRank.Enabled = False
		cboAlias.Enabled = False
		cboFlags.Enabled = False
		txtDescription.Enabled = False
		txtSpecialNotes.Enabled = False
		chkDisable.Enabled = False
		
		lblRank.Enabled = False
		lblAlias.Enabled = False
		lblFlags.Enabled = False
		lblDescription.Enabled = False
		lblSpecialNotes.Enabled = False
		
		cmdAliasAdd.Enabled = False
		cmdAliasRemove.Enabled = False
		cmdFlagAdd.Enabled = False
		cmdFlagRemove.Enabled = False
		
		chkDisable.Visible = False
		
		m_SelectedElement.IsDirty = False
		cmdSave.Enabled = False
		cmdDiscard.Enabled = False
		cmdDeleteCommand.Enabled = False
		
		lblSyntax.Text = vbNullString
		lblSyntax.ForeColor = System.Drawing.ColorTranslator.FromOle(RTBColors.ConsoleText)
		
	End Sub
	
	'UPGRADE_WARNING: Event cboFlags.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cboFlags.Change was upgraded to cboFlags.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cboFlags_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFlags.TextChanged
		
		If (BotVars.CaseSensitiveFlags = False) Then
			cboFlags.Text = UCase(cboFlags.Text)
			
			cboFlags.SelectionStart = Len(cboFlags.Text)
		End If
		
	End Sub
	
	'// 08/29/2008 JSM - Created
	Private Sub cboFlags_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboFlags.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim i As Short
		
		'// Enter
		If KeyCode = 13 Then
			'// Make sure it doesnt have a space
			If InStr(cboFlags.Text, " ") Then
				MsgBox("Flags cannot contain spaces.", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, Me.Text)
				cboFlags.SelectionStart = 1
				cboFlags.SelectionLength = Len(cboFlags.Text)
				Exit Sub
			End If
			'// Make sure its not already a flag
			For i = 0 To cboFlags.Items.Count - 1
				If VB6.GetItemString(cboFlags, i) = cboFlags.Text Then
					cboFlags.Text = ""
					Exit Sub
				End If
			Next i
			
			'// If we made it this far, it should be safe to add it to the list
			cboFlags.Items.Add(cboFlags.Text)
			cboFlags.Text = ""
			Call FormIsDirty()
		End If
		
		'// Delete
		If KeyCode = 46 Then
			For i = 0 To cboFlags.Items.Count - 1
				'// If the current text is already in the list, lets delete it. Otherwise,
				'// this code should behave like a normal delete keypress.
				If VB6.GetItemString(cboFlags, i) = cboFlags.Text Then
					cboFlags.Items.RemoveAt(i)
					cboFlags.Text = ""
					Call FormIsDirty()
					Exit Sub
				End If
			Next i
		End If
		
		
	End Sub
	
	
	'// 08/29/2008 JSM - Created
	Private Sub cboAlias_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboAlias.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim i As Short
		
		'// Enter
		If KeyCode = 13 Then
			'// Make sure it doesnt have a space
			If InStr(cboAlias.Text, " ") Then
				MsgBox("Aliases cannot contain spaces.", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, Me.Text)
				cboAlias.SelectionStart = 1
				cboAlias.SelectionLength = Len(cboAlias.Text)
				Exit Sub
			End If
			'// Make sure its not already an alias
			For i = 0 To cboAlias.Items.Count - 1
				If VB6.GetItemString(cboAlias, i) = cboAlias.Text Then
					cboAlias.Text = ""
					Exit Sub
				End If
			Next i
			
			'// TODO: Make sure its not an alias for another command. Must loop through the
			'// m_CommandsDoc elements to get all aliases and make sure its unique. This logic
			'// should probably be in its own function.
			
			'// If we made it this far, it should be safe to add it to the list
			cboAlias.Items.Add(cboAlias.Text)
			cboAlias.Text = ""
			Call FormIsDirty()
		End If
		
		'// Delete
		If KeyCode = 46 Then
			For i = 0 To cboAlias.Items.Count - 1
				'// If the current text is already in the list, lets delete it. Otherwise,
				'// this code should behave like a normal delete keypress.
				If VB6.GetItemString(cboAlias, i) = cboAlias.Text Then
					cboAlias.Items.RemoveAt(i)
					cboAlias.Text = ""
					Call FormIsDirty()
					Exit Sub
				End If
			Next i
		End If
		
	End Sub
	
	'// Must mark the element as dirty when these change
	'UPGRADE_WARNING: Event txtRank.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRank_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRank.TextChanged
		Call FormIsDirty()
	End Sub
	'UPGRADE_WARNING: Event txtDescription.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDescription_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescription.TextChanged
		Call FormIsDirty()
	End Sub
	'UPGRADE_WARNING: Event txtSpecialNotes.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSpecialNotes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpecialNotes.TextChanged
		Call FormIsDirty()
	End Sub
	'UPGRADE_WARNING: Event chkDisable.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDisable_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisable.CheckStateChanged
		Call FormIsDirty()
	End Sub
End Class