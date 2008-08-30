VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommands 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Manager"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
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
   ScaleHeight     =   4920
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   300
      Left            =   6600
      TabIndex        =   12
      Top             =   4395
      Width           =   1095
   End
   Begin MSComctlLib.TreeView trvCommands 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8070
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
      Height          =   4695
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboFlags 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCommands.frx":0000
         Left            =   3240
         List            =   "frmCommands.frx":0002
         TabIndex        =   13
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cboAlias 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCommands.frx":0004
         Left            =   1605
         List            =   "frmCommands.frx":0006
         TabIndex        =   9
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "Disable"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
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
         Width           =   3975
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
         Width           =   3975
      End
      Begin VB.Label lblAlias 
         Caption         =   "Custom aliases:"
         Height          =   255
         Left            =   1605
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank (1 - 200):"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFlags 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblSpecialNotes 
         Caption         =   "Special notes:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CommandsDoc As MSXML2.DOMDocument

Private Enum NodeType
  nCommand
  nArgument
  nRestriction
End Enum

'// Checks the hiarchy of the treenodes to determine what type of node it is.
'// 08/29/2008 JSM - Created
Private Function GetNodeInfo(node As MSComctlLib.node, ByRef CommandName As String, ByRef argumentName As String, ByRef restrictionName As String) As NodeType
    If node.Parent Is Nothing Then
        GetNodeInfo = NodeType.nCommand
        CommandName = node.text
        argumentName = ""
        restrictionName = ""
    ElseIf node.Parent.Parent Is Nothing Then
        GetNodeInfo = NodeType.nArgument
        CommandName = node.Parent.text
        argumentName = node.text
        restrictionName = ""
    Else
        GetNodeInfo = NodeType.nRestriction
        CommandName = node.Parent.Parent.text
        argumentName = node.Parent.text
        restrictionName = node.text
    End If
End Function




Private Sub Form_Load()
    '// Load commands.xml
    Set m_CommandsDoc = New MSXML2.DOMDocument
    
    If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
        Exit Sub
    End If
    
    Call m_CommandsDoc.Load(App.Path & "\commands.xml")
    Call PopulateTreeView
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_CommandsDoc = Nothing
    
End Sub

'// Reads commands.xml and
Private Sub PopulateTreeView()
    Dim xmlCommand        As MSXML2.IXMLDOMNode
    Dim xmlArgs           As MSXML2.IXMLDOMNodeList
    Dim xmlArgRestricions As MSXML2.IXMLDOMNodeList

    Dim nCommand          As node
    Dim nArg              As node
    Dim nArgRestriction   As node
    
    Dim CommandName       As String
    Dim argumentName      As String
    Dim restrictionName   As String
    
    '// Counters
    Dim j                 As Integer
    Dim i                 As Integer

    '// reset the treeview
    trvCommands.Nodes.Clear

    '// loop through all child nodes
    For Each xmlCommand In m_CommandsDoc.documentElement.childNodes
        CommandName = xmlCommand.Attributes.getNamedItem("name").text
        Set nCommand = trvCommands.Nodes.Add(, , , CommandName)
        
        Set xmlArgs = xmlCommand.selectNodes("arguments/argument")
        '// 08/29/2008 JSM - removed 'Not (xmlArgs Is Nothing)' condition. xmlArgs will always be
        '//                  something, even if nothing matches the XPath expression.
        For i = 0 To (xmlArgs.length - 1)
            argumentName = xmlArgs(i).Attributes.getNamedItem("name").text
            Set nArg = trvCommands.Nodes.Add(nCommand, tvwChild, , argumentName)
            Set xmlArgRestricions = xmlArgs(i).selectNodes("restriction")
            
            For j = 0 To (xmlArgRestricions.length - 1)
                restrictionName = xmlArgRestricions(j).Attributes.getNamedItem("name").text
                Set nArgRestriction = trvCommands.Nodes.Add(nArg, tvwChild, , restrictionName)
            Next j
        Next i
        
    Next
    
End Sub


'// 08/28/2008 JSM - Created
Private Sub trvCommands_NodeClick(ByVal node As MSComctlLib.node)

    Dim nt As NodeType
    Dim CommandName As String
    Dim argumentName As String
    Dim restrictionName As String
    Dim options() As Variant '// <-- boo
    
    Dim xpath As String
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    '// figure out what type of node was clicked on
    nt = GetNodeInfo(node, CommandName, argumentName, restrictionName)
    '// create an array for the StringFormat function, this function will replace
    '// the {0} {1} and {2} with their respective values found below
    '//                {0}           {1}             {2}
    options = Array(CommandName, argumentName, restrictionName)
    
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
            xpath = StringFormat("/commands/command[@name='{0}']/arguments/argument[@name='{1}']/restriction[@name='{2}']", options)
    End Select
    
    '// grab the node from the xpath
    Set xmlElement = m_CommandsDoc.selectSingleNode(xpath)
    Call PrepareForm(nt, xmlElement)
    


End Sub

'// When a node in the treeview is clicked, it should locate the XML element that was
'// used to create the node and call this method to populate appropriate form controls.
'// 08/29/2008 JSM - Created
Private Sub PrepareForm(nt As NodeType, xmlElement As MSXML2.IXMLDOMElement)
    
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim xmlNode As MSXML2.IXMLDOMNode
    
    '// Reset controls
    Call ResetForm
    
    Select Case nt
        Case NodeType.nCommand
            '// txtRank
            txtRank.Enabled = True
            lblRank.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("access/rank")
            If Not (xmlNode Is Nothing) Then
                txtRank.text = xmlElement.selectSingleNode("access/rank").text
            End If
            '// cboAlias
            cboAlias.Enabled = True
            lblAlias.Enabled = True
            For Each xmlNode In xmlElement.selectNodes("alias")
                cboAlias.AddItem xmlNode.text
            Next xmlNode
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
                txtDescription.text = xmlElement.selectSingleNode("documentation/description").text
            End If
            '// txtSpecialNotes
            txtSpecialNotes.Enabled = True
            lblSpecialNotes.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.text = xmlElement.selectSingleNode("documentation/specialnotes").text
            End If
            '// chkDisable
            chkDisable.Enabled = True
            chkDisable.Visible = True
            
        Case NodeType.nArgument
            '// txtRank
            'txtRank.Enabled = True
            'lblRank.Enabled = True
            'Set xmlNode = xmlElement.selectSingleNode("access/rank")
            'If Not (xmlNode Is Nothing) Then
            '    txtRank.text = xmlElement.selectSingleNode("access/rank").text
            'End If
            '// cboAlias
            'cboAlias.Enabled = True
            'lblAlias.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("alias")
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
                txtDescription.text = xmlElement.selectSingleNode("documentation/description").text
            End If
            '// txtSpecialNotes
            txtSpecialNotes.Enabled = True
            lblSpecialNotes.Enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.text = xmlElement.selectSingleNode("documentation/specialnotes").text
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
                txtRank.text = xmlElement.selectSingleNode("access/rank").text
            End If
            '// cboAlias
            'cboAlias.Enabled = True
            'lblAlias.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("alias")
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
                txtDescription.text = xmlElement.selectSingleNode("documentation/description").text
            End If
            '// txtSpecialNotes
            'txtSpecialNotes.Enabled = True
            'lblSpecialNotes.Enabled = True
            'Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            'If Not (xmlNode Is Nothing) Then
            '    txtSpecialNotes.text = xmlElement.selectSingleNode("documentation/specialnotes").text
            'End If
            '// chkDisable
            chkDisable.Enabled = True
            chkDisable.Visible = True
            
    End Select

End Sub

'// Clears all edit controls. Treeview is left intact
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
    End If

    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboFlags.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboFlags.List(i) = cboFlags.text Then
                cboFlags.RemoveItem i
                cboFlags.text = ""
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
    End If
    
    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboAlias.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboAlias.List(i) = cboAlias.text Then
                cboAlias.RemoveItem i
                cboAlias.text = ""
                Exit Sub
            End If
        Next i
    End If
    
End Sub

