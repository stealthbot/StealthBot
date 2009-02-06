VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCommands 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Manager"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView trvCommands 
      Height          =   4935
      Left            =   120
      TabIndex        =   11
      Top             =   390
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8705
      _Version        =   393217
      Indentation     =   575
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraCommand 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "&Discard Changes"
         Height          =   300
         Left            =   2760
         TabIndex        =   14
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes"
         Height          =   300
         Left            =   600
         TabIndex        =   13
         Top             =   4680
         Width           =   1815
      End
      Begin VB.ComboBox cboFlags 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCommands.frx":0000
         Left            =   3600
         List            =   "frmCommands.frx":0002
         TabIndex        =   12
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cboAlias 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCommands.frx":0004
         Left            =   1725
         List            =   "frmCommands.frx":0006
         TabIndex        =   9
         Top             =   600
         Width           =   1605
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   6
         Top             =   630
         Width           =   1215
      End
      Begin VB.CheckBox chkDisable 
         BackColor       =   &H00000000&
         Caption         =   "Disable"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   4200
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtSpecialNotes 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3120
         Width           =   4695
      End
      Begin VB.Label lblAlias 
         BackStyle       =   0  'Transparent
         Caption         =   "Custom aliases:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1725
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "Rank (1 - 200):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFlags 
         BackStyle       =   0  'Transparent
         Caption         =   "Flags:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblSpecialNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "Special notes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command List"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   15
      Top             =   165
      Width           =   990
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CommandsDoc As DOMDocument60
Private m_SelectedElement As SelectedElement

'// Enums
Private Enum NodeType
    nCommand
    nArgument
    nRestriction
End Enum

'// Stores information about the selected node in the treeview
Private Type SelectedElement
    TheNodeType As NodeType
    TheXMLElement As IXMLDOMElement
    IsDirty As Boolean
    CommandName As String
    ArgumentName As String
    restrictionName As String
End Type

'// 08/30/2008 JSM - Created
Private Sub cmdDiscard_Click()
    Call PrepareForm(m_SelectedElement.TheNodeType, m_SelectedElement.TheXMLElement)
End Sub
'// 08/30/2008 JSM - Created
Private Sub cmdSave_Click()
    Call SaveForm
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    '// Load commands.xml
    Set m_CommandsDoc = New DOMDocument60
    
    If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
        Exit Sub
    End If
    '// 08/31/2008 JSM - ensure schema file is present
    If (Dir$(App.Path & "\commands.xsd") = vbNullString) Then
        Exit Sub
    End If
    
    
    If Not ValidateXML(App.Path & "\commands.xml", App.Path & "\commands.xsd") Then
        Exit Sub
    End If
    
    
    Call m_CommandsDoc.Load(App.Path & "\commands.xml")
    
    Call ResetForm
    Call PopulateTreeView
    
    Exit Sub
    
ErrorHandler:

    MsgBox Err.description, vbCritical + vbOKOnly, Me.Caption
    Call ResetForm
    '// Disable our buttons
    cmdSave.Enabled = False
    cmdDiscard.Enabled = False
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_CommandsDoc = Nothing
End Sub

'// Reads commands.xml and
Private Sub PopulateTreeView()
    Dim xmlCommand        As IXMLDOMNode
    Dim xmlArgs           As IXMLDOMNodeList
    Dim xmlArgRestricions As IXMLDOMNodeList

    Dim nCommand          As node
    Dim nArg              As node
    Dim nArgRestriction   As node
    
    Dim CommandName       As String
    Dim ArgumentName      As String
    Dim restrictionName   As String
    
    '// 08/30/2008 JSM - used to get the first command alphabetically
    Dim defaultNode       As node
    
    '// Counters
    Dim j                 As Integer
    Dim i                 As Integer

    '// reset the treeview
    trvCommands.Nodes.Clear

    '// loop through all child nodes
    For Each xmlCommand In m_CommandsDoc.documentElement.childNodes
        CommandName = xmlCommand.Attributes.getNamedItem("name").text
        Set nCommand = trvCommands.Nodes.Add(, , , CommandName)
        
        '// 08/30/2008 JSM - check if this command is the first alphabetically
        If defaultNode Is Nothing Then
            Set defaultNode = nCommand
        Else
            If StrComp(defaultNode.text, nCommand.text) > 0 Then
                Set defaultNode = nCommand
            End If
        End If
        
        Set xmlArgs = xmlCommand.selectNodes("arguments/argument")
        '// 08/29/2008 JSM - removed 'Not (xmlArgs Is Nothing)' condition. xmlArgs will always be
        '//                  something, even if nothing matches the XPath expression.
        For i = 0 To (xmlArgs.length - 1)
            ArgumentName = xmlArgs(i).Attributes.getNamedItem("name").text
            Set nArg = trvCommands.Nodes.Add(nCommand, tvwChild, , ArgumentName)
            Set xmlArgRestricions = xmlArgs(i).selectNodes("restrictions/restriction")
            
            For j = 0 To (xmlArgRestricions.length - 1)
                restrictionName = xmlArgRestricions(j).Attributes.getNamedItem("name").text
                Set nArgRestriction = trvCommands.Nodes.Add(nArg, tvwChild, , restrictionName)
            Next j
        Next i
        
    Next
    
    '// 08/30/2008 JSM - click the first command alphabetically
    trvCommands_NodeClick defaultNode
    
    
End Sub

'// This function will prompt a user to save the changes (if necessary)
'// 08/30/2008 JSM - Created
Private Function PromptToSaveChanges() As Boolean
    
    Dim sMessage As String
    Dim options() As Variant '// <-- boo

    '// If the current form is dirty, lets show a save dialog
    With m_SelectedElement
        If Not .IsDirty Then
            PromptToSaveChanges = True
            Exit Function
        End If
        
        '// Get the message for the prompt
        options = Array(.CommandName, .ArgumentName, .restrictionName)
        Select Case .TheNodeType
            Case NodeType.nCommand
                sMessage = StringFormat("You have not saved your changes to {0}. Do you want to save them now?", options)
            Case NodeType.nArgument
                sMessage = StringFormat("You have not saved your changes to {1}. Do you want to save them now?", options)
            Case NodeType.nRestriction
                sMessage = StringFormat("You have not saved your changes to {2}. Do you want to save them now?", options)
        End Select
        '// Get the user response
        Select Case MsgBox(sMessage, vbQuestion + vbYesNoCancel, Me.Caption)
            Case vbYes:
                Call SaveForm
                PromptToSaveChanges = True
                Exit Function
            Case vbNo:
                PromptToSaveChanges = True
                Exit Function
            Case vbCancel:
                PromptToSaveChanges = False
                Exit Function
        End Select
    
    End With


End Function



'// 08/28/2008 JSM - Created
Private Sub trvCommands_NodeClick(ByVal node As MSComctlLib.node)

    Dim nt As NodeType
    Dim CommandName As String
    Dim ArgumentName As String
    Dim restrictionName As String
    Dim options() As Variant '// <-- boo
    
    Dim xpath As String
    Dim xmlElement As IXMLDOMElement
    
    '// This function will prompt the user to save changes if necessary. If the
    '// return value is false, then the use clicked cancel so we should gtfo of here.
    If PromptToSaveChanges() = False Then
        Exit Sub
    End If
    
    '// figure out what type of node was clicked on
    nt = GetNodeInfo(node, CommandName, ArgumentName, restrictionName)
    '// create an array for the StringFormat function, this function will replace
    '// the {0} {1} and {2} with their respective Values() found below
    '//                {0}           {1}             {2}
    options = Array(CommandName, ArgumentName, restrictionName)
    
    Select Case nt
        Case NodeType.nCommand
            fraCommand.Caption = StringFormat("{0}", options)
            chkDisable.Caption = StringFormat("Disable {0} command", options)
            xpath = StringFormat("/commands/command[@name='{0}']", options)
        Case NodeType.nArgument
            fraCommand.Caption = StringFormat("{0} => {1}", options)
            chkDisable.Caption = StringFormat("Disable {1} argument", options)
            xpath = StringFormat("/commands/command[@name='{0}']/arguments/argument[@name='{1}']", options)
        Case NodeType.nRestriction
            fraCommand.Caption = StringFormat("{0} => {1} => {2}", options)
            chkDisable.Caption = StringFormat("Disable {2} restriction", options)
            xpath = StringFormat("/commands/command[@name='{0}']/arguments/argument[@name='{1}']/restrictions/restriction[@name='{2}']", options)
    End Select
    
    '// Update m_SelectedElement so we know which element we are viewing
    Let m_SelectedElement.CommandName = CommandName
    Let m_SelectedElement.ArgumentName = ArgumentName
    Let m_SelectedElement.restrictionName = restrictionName
    
    '// grab the node from the xpath
    Set xmlElement = m_CommandsDoc.selectSingleNode(xpath)
    Call PrepareForm(nt, xmlElement)

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
Private Function GetNodeInfo(node As MSComctlLib.node, ByRef CommandName As String, ByRef ArgumentName As String, ByRef restrictionName As String) As NodeType
    If node.Parent Is Nothing Then
        GetNodeInfo = NodeType.nCommand
        CommandName = node.text
        ArgumentName = ""
        restrictionName = ""
    ElseIf node.Parent.Parent Is Nothing Then
        GetNodeInfo = NodeType.nArgument
        CommandName = node.Parent.text
        ArgumentName = node.text
        restrictionName = ""
    Else
        GetNodeInfo = NodeType.nRestriction
        CommandName = node.Parent.Parent.text
        ArgumentName = node.Parent.text
        restrictionName = node.text
    End If
End Function


'// Saves the selected treeview node in the commands.xml
'// 08/30/2008 JSM - Created
Private Sub SaveForm()
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode

    Dim i As Integer
    
    With m_SelectedElement
        '// txtRank
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nRestriction Then
            Set xmlNode = .TheXMLElement.selectSingleNode("access/rank")
            If xmlNode Is Nothing Then
                Set xmlNode = .TheXMLElement.selectSingleNode("access")
                If xmlNode Is Nothing Then
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "access", "")
                    .TheXMLElement.appendChild xmlNewNode
                    Set xmlNode = xmlNewNode.cloneNode(True)
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "rank", "")
                    xmlNode.appendChild xmlNewNode
                Else
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "rank", "")
                    xmlNode.appendChild xmlNewNode
                End If
                Set xmlNode = .TheXMLElement.selectSingleNode("access/rank")
            End If
            xmlNode.text = txtRank.text
        End If
        
        '// txtDescription
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nArgument Or .TheNodeType = NodeType.nRestriction Then
            Set xmlNode = .TheXMLElement.selectSingleNode("documentation/description")
            If xmlNode Is Nothing Then
                Set xmlNode = .TheXMLElement.selectSingleNode("documentation")
                If xmlNode Is Nothing Then
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "documentation", "")
                    .TheXMLElement.appendChild xmlNewNode
                    Set xmlNode = xmlNewNode.cloneNode(True)
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "description", "")
                    xmlNode.appendChild xmlNewNode
                Else
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "description", "")
                    xmlNode.appendChild xmlNewNode
                End If
                Set xmlNode = .TheXMLElement.selectSingleNode("documentation/description")
            End If
            xmlNode.text = txtDescription.text
        End If
        
        '// txtSpecialNotes
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nArgument Then
            Set xmlNode = .TheXMLElement.selectSingleNode("documentation/specialnotes")
            If xmlNode Is Nothing Then
                Set xmlNode = .TheXMLElement.selectSingleNode("documentation")
                If xmlNode Is Nothing Then
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "documentation", "")
                    .TheXMLElement.appendChild xmlNewNode
                    Set xmlNode = xmlNewNode.cloneNode(True)
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "specialnotes", "")
                    xmlNode.appendChild xmlNewNode
                Else
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "specialnotes", "")
                    xmlNode.appendChild xmlNewNode
                End If
                Set xmlNode = .TheXMLElement.selectSingleNode("documentation/specialnotes")
            End If
            xmlNode.text = txtSpecialNotes.text
        End If

        '// cboAlias
        If .TheNodeType = NodeType.nCommand Then
            For Each xmlNode In .TheXMLElement.selectNodes("aliases/alias")
                .TheXMLElement.selectSingleNode("aliases").removeChild xmlNode
            Next xmlNode
            For i = 0 To cboAlias.ListCount - 1
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "alias", "")
                xmlNewNode.text = cboAlias.List(i)
                .TheXMLElement.selectSingleNode("aliases").appendChild xmlNewNode
            Next i
        End If
        
        '// cboFlags
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nRestriction Then
            
            '// remove existing flags
            For Each xmlNode In .TheXMLElement.selectNodes("access/flag")
                .TheXMLElement.selectSingleNode("access").removeChild xmlNode
            Next xmlNode
            
            '// make sue the access element exists
            Set xmlNode = .TheXMLElement.selectSingleNode("access")
            If xmlNode Is Nothing Then
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "access", "")
                .TheXMLElement.appendChild xmlNewNode
            End If
            
            Set xmlNode = .TheXMLElement.selectSingleNode("access/flags")
            If xmlNode Is Nothing Then
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "access/flags", "")
                .TheXMLElement.appendChild xmlNewNode
            End If

            '// loop through cboFlags and add the text
            For i = 0 To cboFlags.ListCount - 1
                Set xmlNewNode = xmlNode.selectSingleNode("flag[text()='" & cboFlags.List(i) & "']")
                
                If (xmlNewNode Is Nothing) Then
                    Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "flag", "")
                    xmlNewNode.text = cboFlags.List(i)
                    xmlNode.appendChild xmlNewNode
                End If
            Next i
        End If
        
        
        '// chkDisable
        If .TheNodeType = NodeType.nCommand Then
            If chkDisable.Value = 1 Then
                Call .TheXMLElement.setAttribute("enabled", "false")
            Else
                Call .TheXMLElement.setAttribute("enabled", "true")
            End If
        End If
        
        Call m_CommandsDoc.Save(App.Path & "\commands.xml")
        Call PrepareForm(.TheNodeType, .TheXMLElement)
        
    End With
    

    
    
End Sub


'// When a node in the treeview is clicked, it should locate the XML element that was
'// used to create the node and call this method to populate appropriate form controls.
'// 08/29/2008 JSM - Created
Private Sub PrepareForm(nt As NodeType, xmlElement As IXMLDOMElement)
    
    Dim xmlNode As IXMLDOMNode
    
    '// Reset controls
    Call ResetForm
    
    Select Case nt
        Case NodeType.nCommand
            '// txtRank
            txtRank.Enabled = True
            lblRank.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("access/rank")
            If Not (xmlNode Is Nothing) Then
                txtRank.text = xmlNode.text
            End If
            '// cboAlias
            cboAlias.Enabled = True
            lblAlias.Enabled = True
            For Each xmlNode In xmlElement.selectNodes("aliases/alias")
                cboAlias.AddItem xmlNode.text
            Next xmlNode
            '// cboFlags
            cboFlags.Enabled = True
            lblFlags.Enabled = True
            For Each xmlNode In xmlElement.selectNodes("access/flags/flag")
                cboFlags.AddItem xmlNode.text
            Next xmlNode
            If (cboFlags.ListCount) Then
                cboFlags.text = cboFlags.List(0)
            End If
            
            '// txtDescription
            txtDescription.Enabled = True
            lblDescription.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.text = xmlNode.text
            End If
            '// txtSpecialNotes
            txtSpecialNotes.Enabled = True
            lblSpecialNotes.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.text = xmlNode.text
            End If
            '// chkDisable
            chkDisable.Enabled = True
            chkDisable.Visible = True
            If LCase(xmlElement.getAttribute("enabled")) = "false" Then
                chkDisable.Value = 1
            Else
                chkDisable.Value = 0
            End If
            
        Case NodeType.nArgument
            '// txtRank
            'txtRank.Enabled = True
            'lblRank.Enabled = True
            'Set xmlNode = xmlElement.selectSingleNode("access/rank")
            'If Not (xmlNode Is Nothing) Then
            '    txtRank.text = xmlNode.text
            'End If
            '// cboAlias
            'cboAlias.Enabled = True
            'lblAlias.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("aliases/alias")
            '    cboAlias.AddItem xmlNode.text
            'Next xmlNode
            '// cboFlags
            'cboFlags.Enabled = True
            'lblFlags.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("access/flag")
            '    cboFlags.AddItem xmlNode.text
            'Next xmlNode
            '// txtDescription
            txtDescription.Enabled = True
            lblDescription.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.text = xmlNode.text
            End If
            '// txtSpecialNotes
            txtSpecialNotes.Enabled = True
            lblSpecialNotes.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.text = xmlNode.text
            End If
            '// chkDisable
            'chkDisable.Enabled = True
            'chkDisable.Visible = True
            
        Case NodeType.nRestriction
            '// txtRank
            txtRank.Enabled = True
            lblRank.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("access/rank")
            If Not (xmlNode Is Nothing) Then
                txtRank.text = xmlNode.text
            End If
            '// cboAlias
            'cboAlias.Enabled = True
            'lblAlias.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("aliases/alias")
            '    cboAlias.AddItem xmlNode.text
            'Next xmlNode
            '// cboFlags
            cboFlags.Enabled = True
            lblFlags.Enabled = True
            For Each xmlNode In xmlElement.selectNodes("access/flag")
                cboFlags.AddItem xmlNode.text
            Next xmlNode
            '// txtDescription
            txtDescription.Enabled = True
            lblDescription.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.text = xmlNode.text
            End If
            '// txtSpecialNotes
            'txtSpecialNotes.Enabled = True
            'lblSpecialNotes.Enabled = True
            'Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            'If Not (xmlNode Is Nothing) Then
            '    txtSpecialNotes.text = xmlNode.text
            'End If
            '// chkDisable
            'chkDisable.Enabled = True
            'chkDisable.Visible = True
            
    End Select
    
    '// Update m_SelectedElement so we know which element we are viewing
    Set m_SelectedElement.TheXMLElement = xmlElement
    Let m_SelectedElement.TheNodeType = nt
    Let m_SelectedElement.IsDirty = False
    
    '// Disable our buttons
    cmdSave.Enabled = False
    cmdDiscard.Enabled = False

End Sub

'// Clears and disables all edit controls. Treeview is left intact
'// 08/29/2008 JSM - Created
Private Sub ResetForm()
    
    txtRank.text = ""
    cboAlias.Clear
    cboFlags.Clear
    txtDescription.text = ""
    txtSpecialNotes.text = ""
    chkDisable.Value = 0
    
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
    
    chkDisable.Visible = False
    
   
        
End Sub

'// 08/29/2008 JSM - Created
Private Sub cboFlags_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    '// Enter
    If KeyCode = 13 Then
        '// Make sure it doesnt have a space
        If InStr(cboFlags.text, " ") Then
            MsgBox "Flags cannot contain spaces.", vbOKOnly + vbCritical, Me.Caption
            cboFlags.SelStart = 1
            cboFlags.SelLength = Len(cboFlags.text)
            Exit Sub
        End If
        '// Make sure its not already a flag
        For i = 0 To cboFlags.ListCount - 1
            If cboFlags.List(i) = cboFlags.text Then
                cboFlags.text = ""
                Exit Sub
            End If
        Next i
        
        '// If we made it this far, it should be safe to add it to the list
        cboFlags.AddItem cboFlags.text
        cboFlags.text = ""
        Call FormIsDirty
    End If

    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboFlags.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboFlags.List(i) = cboFlags.text Then
                cboFlags.RemoveItem i
                cboFlags.text = ""
                Call FormIsDirty
                Exit Sub
            End If
        Next i
    End If
    
    
End Sub


'// 08/29/2008 JSM - Created
Private Sub cboAlias_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    '// Enter
    If KeyCode = 13 Then
        '// Make sure it doesnt have a space
        If InStr(cboAlias.text, " ") Then
            MsgBox "Aliases cannot contain spaces.", vbOKOnly + vbCritical, Me.Caption
            cboAlias.SelStart = 1
            cboAlias.SelLength = Len(cboAlias.text)
            Exit Sub
        End If
        '// Make sure its not already an alias
        For i = 0 To cboAlias.ListCount - 1
            If cboAlias.List(i) = cboAlias.text Then
                cboAlias.text = ""
                Exit Sub
            End If
        Next i
            
        '// TODO: Make sure its not an alias for another command. Must loop through the
        '// m_CommandsDoc elements to get all aliases and make sure its unique. This logic
        '// should probably be in its own function.
        
        '// If we made it this far, it should be safe to add it to the list
        cboAlias.AddItem cboAlias.text
        cboAlias.text = ""
        Call FormIsDirty
    End If
    
    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboAlias.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboAlias.List(i) = cboAlias.text Then
                cboAlias.RemoveItem i
                cboAlias.text = ""
                Call FormIsDirty
                Exit Sub
            End If
        Next i
    End If
    
End Sub

'// Must mark the element as dirty when these change
Private Sub txtRank_Change()
    Call FormIsDirty
End Sub
Private Sub txtDescription_Change()
    Call FormIsDirty
End Sub
Private Sub txtSpecialNotes_Change()
    Call FormIsDirty
End Sub
Private Sub chkDisable_Click()
    Call FormIsDirty
End Sub

