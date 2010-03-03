VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmCommands 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Manager"
   ClientHeight    =   7185
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
   Icon            =   "frmCommands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Command Syntax"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   6120
      Width           =   9135
      Begin VB.Label lblRequirements 
         BackStyle       =   0  'Transparent
         Caption         =   "Command Requirements"
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   240
         TabIndex        =   24
         Top             =   435
         Width           =   8655
      End
      Begin VB.Label lblSyntax 
         BackStyle       =   0  'Transparent
         Caption         =   "Command Syntax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   260
         TabIndex        =   23
         Top             =   240
         Width           =   8655
      End
   End
   Begin vbalTreeViewLib6.vbalTreeView trvCommands 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9128
      BackColor       =   10040064
      ForeColor       =   16777215
      LineStyle       =   0
      Style           =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboCommandGroup 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Frame fraCommand 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "&Delete Command"
         Height          =   300
         Left            =   3480
         TabIndex        =   14
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton cmdFlagRemove 
         Caption         =   "-"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   600
         Width           =   270
      End
      Begin VB.CommandButton cmdAliasAdd 
         Caption         =   "+"
         Height          =   315
         Left            =   2900
         TabIndex        =   4
         Top             =   600
         Width           =   270
      End
      Begin VB.CommandButton cmdDiscard 
         Caption         =   "Di&scard Changes"
         Height          =   300
         Left            =   1860
         TabIndex        =   13
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   5400
         Width           =   1455
      End
      Begin VB.ComboBox cboFlags 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   700
      End
      Begin VB.ComboBox cboAlias 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1605
         TabIndex        =   3
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   307
         Left            =   240
         MaxLength       =   25
         TabIndex        =   2
         Top             =   608
         Width           =   1215
      End
      Begin VB.CheckBox chkDisable 
         BackColor       =   &H00000000&
         Caption         =   "Disable"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   4920
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
         TabIndex        =   9
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtSpecialNotes 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3120
         Width           =   4695
      End
      Begin VB.CommandButton cmdFlagAdd 
         Caption         =   "+"
         Height          =   315
         Left            =   4380
         TabIndex        =   7
         Top             =   600
         Width           =   270
      End
      Begin VB.CommandButton cmdAliasRemove 
         Caption         =   "-"
         Height          =   315
         Left            =   3210
         TabIndex        =   5
         Top             =   600
         Width           =   270
      End
      Begin VB.Label lblAlias 
         BackStyle       =   0  'Transparent
         Caption         =   "Custom aliases:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1605
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "Rank (1 - 200):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFlags 
         BackStyle       =   0  'Transparent
         Caption         =   "Flags:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblSpecialNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "Special notes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Label lblCommandList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command List"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   21
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
Private Type SelectedElement
    TheNodeType As NodeType
    IsDirty As Boolean
    commandName As String
    argumentName As String
    restrictionName As String
End Type

Sub ClearTreeViewNodes(ByRef trv As vbalTreeView)
    
    m_ClearingNodes = True
    trv.nodes.Clear
    m_ClearingNodes = False

End Sub


Private Sub cboCommandGroup_Click()

    If PromptToSaveChanges() = True Then
        
        Call ResetForm
        Call PopulateTreeView(getScriptOwner())
        
        '// getScriptOwner() contains this logic already
        'If cboCommandGroup.ListIndex = 0 Then
        '    Call PopulateTreeView
        'Else
        '    Call PopulateTreeView(getScriptOwner())
        'End If
    End If
    
End Sub

Private Sub cmdFlagAdd_Click()

    cboFlags.AddItem cboFlags.Text
    cboFlags.Text = ""
    Call FormIsDirty

End Sub


Private Sub cmdFlagRemove_Click()

    Dim i As Integer
    
    For i = 0 To cboFlags.ListCount - 1
        If (StrComp(cboFlags.Text, cboFlags.List(i), vbBinaryCompare) = 0) Then
            cboFlags.RemoveItem i
            Exit For
        End If
    Next i
    
    cboFlags.Text = ""
    Call FormIsDirty

End Sub

Private Sub cmdAliasAdd_Click()

    cboAlias.AddItem cboAlias.Text
    cboAlias.Text = ""
    Call FormIsDirty

End Sub

Private Sub cmdAliasRemove_Click()

    Dim i As Integer
    
    For i = 0 To cboAlias.ListCount - 1
        If (StrComp(cboAlias.Text, cboAlias.List(i), vbTextCompare) = 0) Then
            cboAlias.RemoveItem i
            Exit For
        End If
    Next i
    
    cboAlias.Text = ""
    Call FormIsDirty

End Sub


'// 08/30/2008 JSM - Created
Private Sub cmdDiscard_Click()

    Call PrepareForm(m_SelectedElement.TheNodeType)
    
End Sub



Private Sub cmdDeleteCommand_Click()
    
    Dim scriptName As String
    Dim scriptIndex As Integer

    scriptName = Mid$(cboCommandGroup.Text, 1, InStr(1, cboCommandGroup.Text, "(") - 2)
    
    scriptIndex = cboCommandGroup.ListIndex

    If vbYes <> MsgBox(StringFormat("Are you sure you want to delete the {0} command for the {1} script?", m_SelectedElement.commandName, scriptName), vbYesNo + vbQuestion, frmCommands.Caption) Then
        Exit Sub
    End If
    
    Call m_Commands.OpenCommand(m_SelectedElement.commandName, scriptName)
    Call m_Commands.Delete
    
    m_SelectedElement.IsDirty = False
    
    Call PopulateOwnerComboBox
    Call ResetForm
    Call PopulateTreeView(scriptName, scriptIndex)

End Sub


'// 08/30/2008 JSM - Created
Private Sub cmdSave_Click()
    Call SaveForm
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim colErrorList As Collection

    '// Load commands.xml
    Set m_Commands = New clsCommandDocObj
    
    'If Not clsCommandDocObj.ValidateXMLFromFiles(App.Path & "\commands.xml", App.Path & "\commands.xsd") Then
    '    Exit Sub
    'End If
    
    'Call m_CommandsDoc.Load(App.Path & "\commands.xml")
    
    'If Not clsCommandDocObj.CommandsSanityCheck(m_CommandsDoc) Then
    '    Exit Sub
    'End If
    
    Call ResetForm
    Call PopulateOwnerComboBox
    Call PopulateTreeView
    
    Exit Sub
    
ErrorHandler:

    frmChat.AddChat RTBColors.ErrorMessageText, Err.description
    Call ResetForm
    '// Disable our buttons
    cmdSave.Enabled = False
    cmdDiscard.Enabled = False
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Commands = Nothing
End Sub


Private Sub PopulateOwnerComboBox()
    
    On Error Resume Next

    Dim i   As Integer
    Dim str As String
    Dim commandCount As Integer
    Dim scriptName As String

    cboCommandGroup.Clear
    
    '// get the script name and number of commands
    scriptName = "Internal Bot Commands"
    commandCount = m_Commands.GetCommandCount()

    '// add the item
    cboCommandGroup.AddItem StringFormat("{0} ({1})", scriptName, commandCount)
    
    For i = 2 To frmChat.SControl.Modules.Count
        scriptName = modScripting.GetScriptName(CStr(i))
        str = SharedScriptSupport.GetSettingsEntry("Public", scriptName)
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            '// get the script name and number of commands
            commandCount = m_Commands.GetCommandCount(scriptName)
            '// only add the commands if there is at least 1 command to show
            If commandCount > 0 Then
                '// add the item
                cboCommandGroup.AddItem StringFormat("{0} ({1})", scriptName, commandCount)
            End If
        End If
    Next i
    cboCommandGroup.ListIndex = 0
    
End Sub

Private Sub PopulateTreeView(Optional strScriptOwner As String = vbNullString, Optional intScriptIndex As Integer = -1)
    
    Dim commandNodes      As IXMLDOMNodeList
    Dim totalCommands     As Integer
    Dim commandNameArray() As Variant
    
    Dim xpath             As String
    
    Dim xmlCommand        As IXMLDOMNode
    Dim xmlArgs           As IXMLDOMNodeList
    Dim xmlArgRestricions As IXMLDOMNodeList

    Dim nCommand          As cTreeViewNode
    Dim nArg              As cTreeViewNode
    Dim nArgRestriction   As cTreeViewNode
    
    Dim commandName       As String
    Dim argumentName      As String
    Dim restrictionName   As String
    
    '// 08/30/2008 JSM - used to get the first command alphabetically
    Dim defaultNode       As cTreeViewNode
    
    '// Counters
    Dim j                 As Integer
    Dim i                 As Integer
    Dim x                 As Integer
    
    strScriptOwner = clsCommandObj.CleanXPathVar(strScriptOwner)

    '// reset the treeview
    If trvCommands.nodes.Count > 0 Then
        trvCommands.nodes(1).Selected = True
    End If
    
    Call ClearTreeViewNodes(trvCommands)
    
    '// create xpath expression based on strScriptOwner
    If LenB(strScriptOwner) = 0 Then
        xpath = "/commands/command[not(@owner)]"
        'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
    Else
        xpath = StringFormat("/commands/command[@owner='{0}']", strScriptOwner)
        'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
    End If
    
    '// get a list of all the commands
    Set commandNodes = m_Commands.XMLDocument.documentElement.selectNodes(xpath)
    ReDim commandNameArray(commandNodes.length)
    
    '// read them 1 at a time and add them to an array
    x = 0
    For Each xmlCommand In m_Commands.XMLDocument.documentElement.selectNodes(xpath)
        commandNameArray(x) = xmlCommand.Attributes.getNamedItem("name").Text
        x = x + 1
    Next xmlCommand
    
    '// sort the command names
    Call BubbleSort1(commandNameArray)

    '// loop through the sorted array and select the commands
    For x = LBound(commandNameArray) To UBound(commandNameArray)
        
        commandName = commandNameArray(x)
        commandName = clsCommandObj.CleanXPathVar(commandName)
        If LenB(commandName) > 0 Then
            '// create xpath expression based on strScriptOwner
            If LenB(strScriptOwner) = 0 Then
                xpath = StringFormat("/commands/command[@name='{0}' and not(@owner)]", commandName)
                'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
            Else
                xpath = StringFormat("/commands/command[@name='{0}' and @owner='{1}']", commandName, strScriptOwner)
                'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
            End If
    
            Set xmlCommand = m_Commands.XMLDocument.documentElement.selectSingleNode(xpath)
        
            commandName = xmlCommand.Attributes.getNamedItem("name").Text
            Set nCommand = trvCommands.nodes.Add(trvCommands.nodes.Parent, etvwChild, commandName, commandName)
            
            '// 08/30/2008 JSM - check if this command is the first alphabetically
            If defaultNode Is Nothing Then
                Set defaultNode = nCommand
            Else
                If StrComp(defaultNode.Text, nCommand.Text) > 0 Then
                    Set defaultNode = nCommand
                End If
            End If
            
            Set xmlArgs = xmlCommand.selectNodes("arguments/argument")
            '// 08/29/2008 JSM - removed 'Not (xmlArgs Is Nothing)' condition. xmlArgs will always be
            '//                  something, even if nothing matches the XPath expression.
            For i = 0 To (xmlArgs.length - 1)
            
                argumentName = xmlArgs(i).Attributes.getNamedItem("name").Text
                If (Not xmlArgs(i).Attributes.getNamedItem("optional") Is Nothing) Then
                    If (xmlArgs(i).Attributes.getNamedItem("optional").Text = "1") Then
                        argumentName = StringFormat("[{0}]", argumentName)
                    End If
                End If
                
                '// Add the datatype to the argument name
                If (Not xmlArgs(i).Attributes.getNamedItem("type") Is Nothing) Then
                    argumentName = StringFormat("{0} ({1})", argumentName, xmlArgs(i).Attributes.getNamedItem("type").Text)
                Else
                    argumentName = StringFormat("{0} ({1})", argumentName, "String")
                End If
                
                Set nArg = trvCommands.nodes.Add(nCommand, etvwChild, commandName & "." & argumentName, argumentName)
                
                Set xmlArgRestricions = xmlArgs(i).selectNodes("restrictions/restriction")
                
                For j = 0 To (xmlArgRestricions.length - 1)
                    restrictionName = xmlArgRestricions(j).Attributes.getNamedItem("name").Text
                    Set nArgRestriction = trvCommands.nodes.Add(nArg, etvwChild, commandName & "." & argumentName & "." & restrictionName, restrictionName)
                Next j
            Next i
        End If '// Len(commandName) > 0
    Next x
    
    '// 08/30/2008 JSM - click the first command alphabetically
    ' fixed to work with SelectedNodeChanged() -Ribose/2009-08-10
    If Not (defaultNode Is Nothing) Then
        defaultNode.Selected = True
    Else
        trvCommands_SelectedNodeChanged
    End If
    
    If (intScriptIndex >= 0) Then
        cboCommandGroup.ListIndex = intScriptIndex
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
                sMessage = StringFormat("You have not saved your changes to {0}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
            Case NodeType.nArgument
                sMessage = StringFormat("You have not saved your changes to {1}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
            Case NodeType.nRestriction
                sMessage = StringFormat("You have not saved your changes to {2}. Do you want to save them now?", .commandName, .argumentName, .restrictionName)
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
' moved to _SelectedNodeChanged -Ribose
' if no node is selected (such as none existing), now disables all fields -Ribose/2009-08-10
Private Sub trvCommands_SelectedNodeChanged()

    On Error GoTo ErrorHandler

    Dim node As cTreeViewNode
    Dim nt As NodeType
    Dim commandName As String
    Dim argumentName As String
    Dim restrictionName As String
    
    Dim xpath As String
    
    If m_ClearingNodes Then Exit Sub
    
    Set node = trvCommands.SelectedItem
    If node Is Nothing Then
        Call ResetForm
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
    Call m_Commands.OpenCommand(commandName, IIf(cboCommandGroup.ListIndex = 0, vbNullString, cboCommandGroup.Text))
    Call ResetForm
    Call PrepareForm(nt)
    
    Exit Sub
    
ErrorHandler:

    frmChat.AddChat RTBColors.ErrorMessageText, Err.description
    Call ResetForm
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
Private Function GetNodeInfo(node As cTreeViewNode, ByRef commandName As String, ByRef argumentName As String, ByRef restrictionName As String) As NodeType
    Dim s() As String
    
    If LenB(node.Key) > 0 Then
        s = Split(node.Key, ".")
        Select Case UBound(s)
            Case 0
                commandName = s(0)
                argumentName = vbNullString
                restrictionName = vbNullString
                GetNodeInfo = nCommand
            Case 1
                commandName = s(0)
                argumentName = s(1)
                restrictionName = vbNullString
                GetNodeInfo = nArgument
            Case 2
                commandName = s(0)
                argumentName = s(1)
                restrictionName = s(2)
                GetNodeInfo = nRestriction
        End Select
        '// strip the [ ] around optional parameters
        If Left$(argumentName, 1) = "[" Then
            argumentName = Mid$(argumentName, 2, InStr(1, argumentName, "]") - 2)
        End If
        If (InStr(1, argumentName, "(") >= 0 And Right$(argumentName, 1) = ")") Then
            argumentName = Mid$(argumentName, 1, InStr(1, argumentName, "(") - 2)
        End If
    End If
End Function

Private Function getFlags()

    Dim i As Integer
    Dim sTmp As String

    sTmp = ""
    For i = 0 To cboFlags.ListCount - 1
        sTmp = sTmp & cboFlags.List(i)
    Next i
    getFlags = sTmp

End Function


'// Saves the selected treeview node in the commands.xml
'// 08/30/2008 JSM - Created
Private Sub SaveForm()
    
    Dim parameter As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj
    
    Dim i

    Select Case m_SelectedElement.TheNodeType
        Case NodeType.nCommand:
            '// saving the command
            With m_Commands
                .description = txtDescription.Text
                .SpecialNotes = txtSpecialNotes.Text
                .RequiredRank = txtRank.Text
                .RequiredFlags = getFlags()
                While .aliases.Count <> 0
                    .aliases.Remove 1
                Wend
                For i = 0 To cboAlias.ListCount - 1
                    .aliases.Add cboAlias.List(i)
                Next i
                .IsEnabled = Not CBool(chkDisable.Value)
            End With
       
        Case NodeType.nArgument
            '// saving the parameter
            With m_Commands
                Set parameter = .GetParameterByName(m_SelectedElement.argumentName)
                With parameter
                    .description = txtDescription.Text
                    .SpecialNotes = txtSpecialNotes.Text
                End With
            End With
            
        Case NodeType.nRestriction
            '// saving the restriction
            With m_Commands
                Set parameter = m_Commands.GetParameterByName(m_SelectedElement.argumentName)
                With parameter
                    Set Restriction = parameter.GetRestrictionByName(m_SelectedElement.restrictionName)
                    With Restriction
                        .RequiredRank = txtRank.Text
                        .RequiredFlags = getFlags()
                    End With
                End With
            End With
    
    End Select
    
    Call m_Commands.Save
    Call ResetForm
    Call PrepareForm(m_SelectedElement.TheNodeType)
    
End Sub

Private Function PrepString(ByVal str As String)
    Dim retVal As String
    retVal = str
    retVal = Replace(retVal, vbLf, vbCrLf)
    retVal = Replace(retVal, vbTab, vbNullString)
    PrepString = retVal
End Function

Function getScriptOwner() As String
    
    '// return vbNullstring if internal commands is selected, otherwise the script name
    If cboCommandGroup.ListIndex = 0 Then
        getScriptOwner = vbNullString
    Else
        getScriptOwner = Mid$(cboCommandGroup.Text, 1, InStr(1, cboCommandGroup.Text, "(") - 2)
    End If
    
End Function



'// When a node in the treeview is clicked, it should locate the XML element that was
'// used to create the node and call this method to populate appropriate form controls.
'// 08/29/2008 JSM - Created
Private Sub PrepareForm(nt As NodeType)
    
    Dim parameter As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj
    
    Dim requirements As String
    Dim sItem As String
    Dim i As Integer
    
    With m_SelectedElement

        Call m_Commands.OpenCommand(.commandName, getScriptOwner())
        
        '// lblSyntax
        lblSyntax.Caption = m_Commands.SyntaxString
        '// lblRequirements
        lblRequirements.Caption = m_Commands.RequirementsString
    
        Select Case nt
            Case NodeType.nCommand
                '// txtRank
                txtRank.Enabled = True
                lblRank.Enabled = True
                txtRank.Text = m_Commands.RequiredRank
                '// cboAlias
                cboAlias.Enabled = True
                lblAlias.Enabled = True
                cmdAliasAdd.Enabled = True
                cmdAliasRemove.Enabled = True
                For i = 1 To m_Commands.aliases.Count
                    cboAlias.AddItem m_Commands.aliases(i)
                Next i
                '// cboFlags
                cboFlags.Enabled = True
                lblFlags.Enabled = True
                cmdFlagAdd.Enabled = True
                cmdFlagRemove.Enabled = True
                For i = 1 To Len(m_Commands.RequiredFlags)
                    cboFlags.AddItem Mid$(m_Commands.RequiredFlags, i, 1)
                Next i
                If (cboFlags.ListCount) Then
                    cboFlags.Text = cboFlags.List(0)
                End If
                '// txtDescription
                txtDescription.Enabled = True
                lblDescription.Enabled = True
                txtDescription.Text = PrepString(m_Commands.description)
                '// txtSpecialNotes
                txtSpecialNotes.Enabled = True
                lblSpecialNotes.Enabled = True
                txtSpecialNotes.Text = PrepString(m_Commands.SpecialNotes)
                '// chkDisable
                chkDisable.Enabled = True
                chkDisable.Visible = True
                chkDisable.Value = IIf(m_Commands.IsEnabled, vbUnchecked, vbChecked)
                '// custom captions
                fraCommand.Caption = StringFormat("{0}", .commandName, .argumentName, .restrictionName)
                chkDisable.Caption = StringFormat("Disable {0} command", .commandName, .argumentName, .restrictionName)
                '// only allow deleting script commands
                If cboCommandGroup.ListIndex > 0 Then
                    cmdDeleteCommand.Enabled = True
                End If
                
            Case NodeType.nArgument
                Set parameter = m_Commands.GetParameterByName(.argumentName)
            
                '// txtDescription
                txtDescription.Enabled = True
                lblDescription.Enabled = True
                txtDescription.Text = PrepString(parameter.description)
                '// txtSpecialNotes
                txtSpecialNotes.Enabled = True
                lblSpecialNotes.Enabled = True
                txtSpecialNotes.Text = PrepString(parameter.SpecialNotes)
                '// custom captions
                fraCommand.Caption = StringFormat("{0} => {1}{3}", .commandName, .argumentName, .restrictionName, IIf(parameter.IsOptional, " - Optional", ""))
                chkDisable.Caption = StringFormat("Disable {1} argument", .commandName, .argumentName, .restrictionName)
                
            Case NodeType.nRestriction
            
                Set parameter = m_Commands.GetParameterByName(.argumentName)
                Set Restriction = parameter.GetRestrictionByName(.restrictionName)
            
                '// txtRank
                txtRank.Enabled = True
                lblRank.Enabled = True
                txtRank.Text = Restriction.RequiredRank
                '// cboFlags
                cboFlags.Enabled = True
                lblFlags.Enabled = True
                For i = 1 To Len(Restriction.RequiredFlags)
                    cboFlags.AddItem Mid$(Restriction.RequiredFlags, i, 1)
                Next i
                If (cboFlags.ListCount) Then
                    cboFlags.Text = cboFlags.List(0)
                End If
                '// special captions
                fraCommand.Caption = StringFormat("{0} => {1}{3} => {2}", .commandName, .argumentName, .restrictionName, IIf(parameter.IsOptional, " - Optional", ""))
                chkDisable.Caption = StringFormat("Disable {2} restriction", .commandName, .argumentName, .restrictionName)
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
    cboAlias.Clear
    cboFlags.Clear
    txtDescription.Text = vbNullString
    txtSpecialNotes.Text = vbNullString
    fraCommand.Caption = vbNullString
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
    
    cmdAliasAdd.Enabled = False
    cmdAliasRemove.Enabled = False
    cmdFlagAdd.Enabled = False
    cmdFlagRemove.Enabled = False
    
    chkDisable.Visible = False
    
    m_SelectedElement.IsDirty = False
    cmdSave.Enabled = False
    cmdDiscard.Enabled = False
    cmdDeleteCommand.Enabled = False
    
    lblSyntax.Caption = vbNullString
    lblSyntax.ForeColor = RTBColors.ConsoleText
    
End Sub

Private Sub cboFlags_Change()

    If (BotVars.CaseSensitiveFlags = False) Then
        cboFlags.Text = UCase$(cboFlags.Text)
        
        cboFlags.selStart = Len(cboFlags.Text)
    End If
    
End Sub

'// 08/29/2008 JSM - Created
Private Sub cboFlags_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    '// Enter
    If KeyCode = 13 Then
        '// Make sure it doesnt have a space
        If InStr(cboFlags.Text, " ") Then
            MsgBox "Flags cannot contain spaces.", vbOKOnly + vbCritical, Me.Caption
            cboFlags.selStart = 1
            cboFlags.selLength = Len(cboFlags.Text)
            Exit Sub
        End If
        '// Make sure its not already a flag
        For i = 0 To cboFlags.ListCount - 1
            If cboFlags.List(i) = cboFlags.Text Then
                cboFlags.Text = ""
                Exit Sub
            End If
        Next i
        
        '// If we made it this far, it should be safe to add it to the list
        cboFlags.AddItem cboFlags.Text
        cboFlags.Text = ""
        Call FormIsDirty
    End If

    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboFlags.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboFlags.List(i) = cboFlags.Text Then
                cboFlags.RemoveItem i
                cboFlags.Text = ""
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
        If InStr(cboAlias.Text, " ") Then
            MsgBox "Aliases cannot contain spaces.", vbOKOnly + vbCritical, Me.Caption
            cboAlias.selStart = 1
            cboAlias.selLength = Len(cboAlias.Text)
            Exit Sub
        End If
        '// Make sure its not already an alias
        For i = 0 To cboAlias.ListCount - 1
            If cboAlias.List(i) = cboAlias.Text Then
                cboAlias.Text = ""
                Exit Sub
            End If
        Next i
            
        '// TODO: Make sure its not an alias for another command. Must loop through the
        '// m_CommandsDoc elements to get all aliases and make sure its unique. This logic
        '// should probably be in its own function.
        
        '// If we made it this far, it should be safe to add it to the list
        cboAlias.AddItem cboAlias.Text
        cboAlias.Text = ""
        Call FormIsDirty
    End If
    
    '// Delete
    If KeyCode = 46 Then
        For i = 0 To cboAlias.ListCount - 1
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            If cboAlias.List(i) = cboAlias.Text Then
                cboAlias.RemoveItem i
                cboAlias.Text = ""
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

