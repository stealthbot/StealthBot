VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmCommandManager 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Manager"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommandManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   7440
      TabIndex        =   30
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdSaveForm 
      Caption         =   "Apply and Cl&ose"
      Default         =   -1  'True
      Height          =   300
      Left            =   8160
      TabIndex        =   29
      Top             =   5880
      Width           =   1335
   End
   Begin vbalTreeViewLib6.vbalTreeView trvCommands 
      Height          =   4965
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8758
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
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Frame fraCommand 
      BackColor       =   &H00000000&
      Caption         =   "command"
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox chkMatchFatal 
         BackColor       =   &H00000000&
         Caption         =   "Fatal on restricted"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "If enabled, parameters that fail this restriction do not get executed."
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkMatchErrorGlobal 
         BackColor       =   &H00000000&
         Caption         =   "Error message global for this parameter"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Enable this to make the above error message apply to all restrictions under this parameter."
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtMatchError 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   200
         TabIndex        =   18
         ToolTipText     =   "Leave blank for no error."
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CheckBox chkMatchCase 
         BackColor       =   &H00000000&
         Caption         =   "Case sensitive"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtMatch 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Leave blank to match any parameter."
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton cmdAddAlias 
         Caption         =   "Add"
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtAddAlias 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveCommand 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3840
         TabIndex        =   26
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdDiscardCommand 
         Caption         =   "D&iscard"
         Height          =   300
         Left            =   2640
         TabIndex        =   27
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   22
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtFlags 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CheckBox chkDisable 
         BackColor       =   &H00000000&
         Caption         =   "Disable"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4800
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtSpecialNotes 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1920
         Width           =   4815
      End
      Begin MSComctlLib.ListView lvAliases 
         Height          =   1305
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2302
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "Icons"
         ForeColor       =   16777215
         BackColor       =   10040064
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblMatchError 
         BackStyle       =   0  'Transparent
         Caption         =   "&Error on no match"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblMatch 
         BackStyle       =   0  'Transparent
         Caption         =   "&Match message"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblRequirements 
         BackStyle       =   0  'Transparent
         Caption         =   "Command Requirements"
         ForeColor       =   &H00FFFFFF&
         Height          =   1260
         Left            =   2280
         TabIndex        =   25
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblSyntax 
         BackStyle       =   0  'Transparent
         Caption         =   ".find <username/rank> [upperrank]"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4800
      End
      Begin VB.Label lblAlias 
         BackStyle       =   0  'Transparent
         Caption         =   "&Aliases"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "Req &Rank (1-200)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblFlags 
         BackStyle       =   0  'Transparent
         Caption         =   "Req &Flags"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "D&escription"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblSpecialNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "S&pecial notes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   4815
      End
   End
   Begin VB.Label lblCommandList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command &List"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   990
   End
End
Attribute VB_Name = "frmCommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Enums
Private Enum NodeType
    nCommand = 0
    nParameter = 1
    nRestriction = 2
End Enum

Private Type CommandGroup
    ScriptName As String
    CommandCount As Integer
End Type

'// Stores information about the selected node in the treeview
Private Type SelectedElement
    TheNodeType As NodeType
    IsDirty As Boolean
    CommandName As String
    ParameterName As String
    RestrictionName As String
    CommandGroup As CommandGroup
End Type

Private m_Commands As clsCommandDocObj
Private m_CommandGroups() As CommandGroup
Private m_SelectedElement As SelectedElement
Private m_DocumentCopy As DOMDocument60
Private m_ClearingNodes As Boolean
Private m_PopulatingGroupList As Boolean
Private m_ResettingForm As Boolean

Private Sub ClearTreeViewNodes(ByRef trv As vbalTreeView)

    m_ClearingNodes = True
    trv.Nodes.Clear
    m_ClearingNodes = False

End Sub

Private Sub cboCommandGroup_Click()
    If Not m_PopulatingGroupList Then
        m_SelectedElement.CommandGroup = m_CommandGroups(cboCommandGroup.ListIndex)
        If PromptToSaveChanges() = True Then
            Call ResetCommandForm
            Call PopulateTreeView(GetScriptOwner())
        End If
    End If
End Sub

Private Sub cmdAddAlias_Click()
    Dim Exists  As Boolean
    Dim Alias   As String
    Dim Command As String
    Dim Other   As String
    Dim Script  As String
    Dim i       As Integer

    Exists = False
    Command = m_Commands.Name
    Script = m_SelectedElement.CommandGroup.ScriptName
    Alias = LCase$(txtAddAlias.Text)

    '// Make sure its not already an alias
    For i = 1 To lvAliases.ListItems.Count
        If StrComp(lvAliases.ListItems.Item(i), Alias, vbTextCompare) = 0 Then
            Exists = True
            Other = m_SelectedElement.CommandName
            Exit For
        End If
    Next i

    If Not Exists Then
        '// temporarily save the current OpenCommand
        Call m_Commands.Save(False)
        '// didn't find a match in our current alias list, look in whole document for enabled commands
        If (m_Commands.OpenCommand(Alias, vbNullChar, True, m_DocumentCopy)) Then
            If (StrComp(m_Commands.Name, Command, vbTextCompare) <> 0) And (StrComp(m_Commands.Owner, Script, vbTextCompare) <> 0) Then
                '// found a command with this alias as it's name or aliases, that isn't the current command (was checked above with active list)?
                Exists = True
                Other = m_Commands.Name
                Script = m_Commands.Owner
            End If
        End If
        '// reset OpenCommand to current
        Call m_Commands.OpenCommand(m_SelectedElement.CommandName, GetScriptOwner(), False, m_DocumentCopy)
    End If

    If Exists Then
        If Len(Script) > 0 Then Script = " for the " & Script & " script"
        Call MsgBox(StringFormat("That alias already exists as a name of the {0} command{1}.", Other, Script), vbOKOnly Or vbCritical, "Command Manager - Add Alias Error")
    Else
        '// If we made it this far, it should be safe to add it to the list
        lvAliases.ListItems.Add , , txtAddAlias.Text
        txtAddAlias.Text = vbNullString
        Call FormIsDirty
    End If
End Sub

Private Sub lvAliases_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lvAliases.ListItems.Count > 0 And Not lvAliases.SelectedItem Is Nothing Then
            Call lvAliases.ListItems.Remove(lvAliases.SelectedItem.Index)
            Call FormIsDirty
        End If
        KeyCode = 0
    End If
End Sub

'// 08/30/2008 JSM - Created
Private Sub cmdDiscardCommand_Click()

    Call PrepareCommandForm

End Sub

Private Sub cmdDeleteCommand_Click()

    Dim ScriptName  As String
    Dim ScriptIndex As Integer
    Dim i           As Integer

    ScriptName = GetScriptOwner()
    ScriptIndex = cboCommandGroup.ListIndex - 1
    
    ' don't allow internal commands to be deleted
    If ScriptIndex < 0 Then Exit Sub

    If vbYes <> MsgBox(StringFormat("Are you sure you want to delete the {0} command for the {1} script?", m_SelectedElement.CommandName, ScriptName), vbYesNo Or vbQuestion, "Command Manager - Confirm Delete") Then
        Exit Sub
    End If

    Call m_Commands.OpenCommand(m_SelectedElement.CommandName, ScriptName, False, m_DocumentCopy)
    Call m_Commands.Delete(False)

    m_SelectedElement.IsDirty = False

    ' regenerate script list (scripts with no commands are hidden)
    Call PopulateOwnerComboBox

    ' if we're deleting a command, combo box could change!
    ' set index to initial script name (or fall back to previously selected index - 1, set above)
    For i = LBound(m_CommandGroups) To UBound(m_CommandGroups)
        If m_CommandGroups(i).ScriptName = ScriptName Then
            ScriptIndex = i
            Exit For
        End If
    Next i
    ' make sure index is in bounds
    If ScriptIndex > UBound(m_CommandGroups) Then ScriptIndex = UBound(m_CommandGroups)
    If ScriptIndex < LBound(m_CommandGroups) Then ScriptIndex = LBound(m_CommandGroups)

    ' save selected group
    m_SelectedElement.CommandGroup = m_CommandGroups(ScriptIndex)

    ' repopulate tree
    Call ResetCommandForm
    Call PopulateTreeView(GetScriptOwner(), ScriptIndex)

End Sub


'// 08/30/2008 JSM - Created
Private Sub cmdSaveCommand_Click()
    Call SaveCommandForm
End Sub

Private Sub Form_Load()

    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim colErrorList As Collection

    '// Load commands.xml
    Set m_Commands = New clsCommandDocObj
    Set m_DocumentCopy = New DOMDocument60
    Call m_Commands.XMLDocument.Save(m_DocumentCopy)
    
    Call PopulateOwnerComboBox
    Call ResetCommandForm
    Call PopulateTreeView
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in frmCommandManager.Form_Load()."

    Call ResetCommandForm
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Commands = Nothing
End Sub


Private Sub PopulateOwnerComboBox()

    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim i            As Integer
    Dim str          As String
    Dim CommandGroup As CommandGroup
    Dim ScriptCount  As Integer

    m_PopulatingGroupList = True
    cboCommandGroup.Clear

    '// get the script name and number of commands
    CommandGroup.ScriptName = vbNullString
    CommandGroup.CommandCount = clsCommandDocObj.GetCommandCount(, m_DocumentCopy)

    ReDim m_CommandGroups(0)
    m_CommandGroups(0) = CommandGroup
    cboCommandGroup.AddItem StringFormat("Internal Bot Commands ({1})", CommandGroup.ScriptName, CommandGroup.CommandCount)

    ScriptCount = 0
    For i = 2 To frmChat.SControl.Modules.Count
        CommandGroup.ScriptName = modScripting.GetScriptName(CStr(i))
        str = SharedScriptSupport.GetSettingsEntry("Public", CommandGroup.ScriptName)
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            '// get the script name and number of commands
            CommandGroup.CommandCount = clsCommandDocObj.GetCommandCount(CommandGroup.ScriptName, m_DocumentCopy)
            '// only add the commands if there is at least 1 command to show
            If CommandGroup.CommandCount > 0 Then
                '// add the item
                ScriptCount = ScriptCount + 1
                ReDim Preserve m_CommandGroups(ScriptCount)
                m_CommandGroups(ScriptCount) = CommandGroup
                cboCommandGroup.AddItem StringFormat("{0} ({1})", CommandGroup.ScriptName, CommandGroup.CommandCount)
            End If
        End If
    Next i
    cboCommandGroup.ListIndex = 0
    m_PopulatingGroupList = False

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in frmCommandManager.PopulateOwnerComboBox()."

    Call ResetCommandForm
    Exit Sub

End Sub

Private Sub PopulateTreeView(Optional strScriptOwner As String = vbNullString, Optional intScriptIndex As Integer = -1)

    Dim CommandNames      As Collection
    Dim TotalCommands     As Integer
    Dim CommandNameArray() As Variant

    Dim xpath             As String

    Dim xmlCommand        As IXMLDOMNode
    Dim xmlArgs           As IXMLDOMNodeList
    Dim xmlArgRestricions As IXMLDOMNodeList

    Dim nCommand          As cTreeViewNode
    Dim nArg              As cTreeViewNode
    Dim nArgRestriction   As cTreeViewNode

    Dim CommandName       As String
    Dim ParameterName     As String
    Dim RestrictionName   As String

    '// 08/30/2008 JSM - used to get the first command alphabetically
    Dim DefaultNode       As cTreeViewNode

    '// Counters
    Dim j                 As Integer
    Dim i                 As Integer
    Dim x                 As Integer

    '// reset the treeview
    If trvCommands.Nodes.Count > 0 Then
        trvCommands.Nodes(1).Selected = True
    End If

    Call ClearTreeViewNodes(trvCommands)

    '// get a list of all the commands
    Set CommandNames = clsCommandDocObj.GetCommands(strScriptOwner, m_DocumentCopy)

    If CommandNames.Count = 0 Then Exit Sub

    ReDim CommandNameArray(CommandNames.Count - 1)

    '// read them 1 at a time and add them to an array
    For x = 1 To CommandNames.Count
        CommandNameArray(x - 1) = CommandNames.Item(x)
    Next x

    '// sort the command names
    Call InsertionSort(CommandNameArray)

    '// loop through the sorted array and select the commands
    For x = LBound(CommandNameArray) To UBound(CommandNameArray)

        CommandName = CommandNameArray(x)
        If LenB(CommandName) > 0 Then
            '// create xpath expression
            xpath = clsCommandObj.GetCommandXPath(CommandName, strScriptOwner, False)

            Set xmlCommand = m_DocumentCopy.documentElement.selectSingleNode(xpath)

            Set nCommand = trvCommands.Nodes.Add(trvCommands.Nodes.Parent, etvwChild, CommandName, CommandName)

            '// 08/30/2008 JSM - check if this command is the first alphabetically
            If DefaultNode Is Nothing Then
                Set DefaultNode = nCommand
            Else
                If StrComp(DefaultNode.Text, nCommand.Text) > 0 Then
                    Set DefaultNode = nCommand
                End If
            End If

            Set xmlArgs = xmlCommand.selectNodes("arguments/argument")
            '// 08/29/2008 JSM - removed 'Not (xmlArgs Is Nothing)' condition. xmlArgs will always be
            '//                  something, even if nothing matches the XPath expression.
            For i = 0 To (xmlArgs.Length - 1)

                ParameterName = xmlArgs(i).Attributes.getNamedItem("name").Text
                If (Not xmlArgs(i).Attributes.getNamedItem("optional") Is Nothing) Then
                    If (xmlArgs(i).Attributes.getNamedItem("optional").Text = "1") Then
                        ParameterName = StringFormat("[{0}]", ParameterName)
                    End If
                End If

                '// Add the datatype to the argument name
                If (Not xmlArgs(i).Attributes.getNamedItem("type") Is Nothing) Then
                    ParameterName = StringFormat("{0} ({1})", ParameterName, LCase$(xmlArgs(i).Attributes.getNamedItem("type").Text))
                Else
                    ParameterName = StringFormat("{0} ({1})", ParameterName, "string")
                End If

                Set nArg = trvCommands.Nodes.Add(nCommand, etvwChild, CommandName & "." & ParameterName, ParameterName)

                Set xmlArgRestricions = xmlArgs(i).selectNodes("restrictions/restriction")

                For j = 0 To (xmlArgRestricions.Length - 1)
                    RestrictionName = xmlArgRestricions(j).Attributes.getNamedItem("name").Text
                    Set nArgRestriction = trvCommands.Nodes.Add(nArg, etvwChild, CommandName & "." & ParameterName & "." & RestrictionName, RestrictionName)
                Next j
            Next i
        End If '// Len(commandName) > 0
    Next x

    '// 08/30/2008 JSM - click the first command alphabetically
    ' fixed to work with SelectedNodeChanged() -Ribose/2009-08-10
    If Not (DefaultNode Is Nothing) Then
        DefaultNode.Selected = True
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
                sMessage = StringFormat("You have not saved your changes to {0}. Do you want to save them now?", .CommandName, .ParameterName, .RestrictionName)
            Case NodeType.nParameter
                sMessage = StringFormat("You have not saved your changes to {1}. Do you want to save them now?", .CommandName, .ParameterName, .RestrictionName)
            Case NodeType.nRestriction
                sMessage = StringFormat("You have not saved your changes to {2}. Do you want to save them now?", .CommandName, .ParameterName, .RestrictionName)
        End Select
        
        '// Get the user response
        Select Case MsgBox(sMessage, vbQuestion Or vbYesNoCancel, "Command Manager - Save")
            Case vbYes:
                Call SaveCommandForm
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

    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim Node As cTreeViewNode
    Dim nt As NodeType
    Dim CommandName As String
    Dim ParameterName As String
    Dim RestrictionName As String

    Dim xpath As String

    If m_ClearingNodes Then Exit Sub

    Set Node = trvCommands.SelectedItem
    If Node Is Nothing Then
        Call ResetCommandForm
        Exit Sub
    End If

    '// This function will prompt the user to save changes if necessary. If the
    '// return value is false, then the use clicked cancel so we should gtfo of here.
    If PromptToSaveChanges() = False Then
        Exit Sub
    End If

    '// figure out what type of node was clicked on
    nt = GetNodeInfo(Node, CommandName, ParameterName, RestrictionName)

    '// Update m_SelectedElement so we know which element we are viewing
    With m_SelectedElement
        .IsDirty = False
        .TheNodeType = nt
        .CommandName = CommandName
        .ParameterName = ParameterName
        .RestrictionName = RestrictionName
        .CommandGroup = m_CommandGroups(cboCommandGroup.ListIndex)
    End With

    '// load the command and set up the form
    Call m_Commands.OpenCommand(CommandName, GetScriptName(), False, m_DocumentCopy)
    Call ResetCommandForm
    Call PrepareCommandForm

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in frmCommandManager.trvCommands_SelectedNodeChanged()."

    Call ResetCommandForm
    Exit Sub

End Sub

'// Call this sub whenever the form controls have been changed
'// 08/30/2008 JSM - Created
Private Sub FormIsDirty()
    If Not m_ResettingForm Then
        m_SelectedElement.IsDirty = True
        cmdSaveCommand.Enabled = True
        cmdDiscardCommand.Enabled = True
        fraCommand.Caption = GetCaptionString
        Me.Caption = "Command Manager - " & fraCommand.Caption
    End If
End Sub

'// Checks the hiarchy of the treenodes to determine what type of node it is.
'// 08/29/2008 JSM - Created
Private Function GetNodeInfo(Node As cTreeViewNode, ByRef CommandName As String, ByRef ParameterName As String, ByRef RestrictionName As String) As NodeType
    Dim s() As String
    
    If LenB(Node.Key) > 0 Then
        s = Split(Node.Key, ".")
        Select Case UBound(s)
            Case 0
                CommandName = s(0)
                ParameterName = vbNullString
                RestrictionName = vbNullString
                GetNodeInfo = nCommand
            Case 1
                CommandName = s(0)
                ParameterName = s(1)
                RestrictionName = vbNullString
                GetNodeInfo = nParameter
            Case 2
                CommandName = s(0)
                ParameterName = s(1)
                RestrictionName = s(2)
                GetNodeInfo = nRestriction
        End Select
        '// strip the [ ] around optional parameters
        If Left$(ParameterName, 1) = "[" Then
            ParameterName = Mid$(ParameterName, 2, InStr(1, ParameterName, "]") - 2)
        End If
        If (InStr(1, ParameterName, "(") >= 0 And Right$(ParameterName, 1) = ")") Then
            ParameterName = Mid$(ParameterName, 1, InStr(1, ParameterName, "(") - 2)
        End If
    End If
End Function


'// Saves the selected treeview node in the commands.xml
'// 08/30/2008 JSM - Created
Private Sub SaveCommandForm()

    Dim Parameter As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj

    Dim i

    Select Case m_SelectedElement.TheNodeType
        Case NodeType.nCommand:
            '// saving the command
            With m_Commands
                .Description = txtDescription.Text
                .SpecialNotes = txtSpecialNotes.Text
                If StrictIsNumeric(txtRank.Text) Then
                    .RequiredRank = CInt(txtRank.Text)
                Else
                    .RequiredRank = -1
                End If
                .RequiredFlags = txtFlags.Text
                While .Aliases.Count <> 0
                    .Aliases.Remove 1
                Wend
                For i = 1 To lvAliases.ListItems.Count
                    .Aliases.Add lvAliases.ListItems.Item(i)
                Next i
                .IsEnabled = Not CBool(chkDisable.Value)
            End With

        Case NodeType.nParameter
            '// saving the parameter
            With m_Commands
                Set Parameter = .GetParameterByName(m_SelectedElement.ParameterName)
                With Parameter
                    .Description = txtDescription.Text
                    .SpecialNotes = txtSpecialNotes.Text
                    .MatchMessage = txtMatch.Text
                    .MatchCaseSensitive = CBool(chkMatchCase.Value)
                    .MatchError = txtMatchError.Text
                End With
            End With

        Case NodeType.nRestriction
            '// saving the restriction
            With m_Commands
                Set Parameter = m_Commands.GetParameterByName(m_SelectedElement.ParameterName)
                With Parameter
                    Set Restriction = Parameter.GetRestrictionByName(m_SelectedElement.RestrictionName)
                    With Restriction
                        .Description = txtDescription.Text
                        .SpecialNotes = txtSpecialNotes.Text
                        If StrictIsNumeric(txtRank.Text) Then
                            .RequiredRank = CInt(txtRank.Text)
                        Else
                            .RequiredRank = -1
                        End If
                        .RequiredFlags = txtFlags.Text
                        .MatchMessage = txtMatch.Text
                        .MatchCaseSensitive = CBool(chkMatchCase.Value)
                        If chkMatchErrorGlobal.Value <> vbUnchecked Then
                            Parameter.RestrictionsSharedError = txtMatchError.Text
                        Else
                            .MatchError = txtMatchError.Text
                        End If
                        .Fatal = CBool(chkMatchFatal.Value)
                    End With
                End With
            End With

    End Select

    Call m_Commands.Save(False)
    Call ResetCommandForm
    Call PrepareCommandForm

End Sub

Private Function PrepString(ByVal str As String)
    Dim retVal As String
    retVal = str
    retVal = Replace(retVal, vbLf, vbCrLf)
    retVal = Replace(retVal, vbTab, vbNullString)
    PrepString = retVal
End Function

Private Function GetScriptOwner() As String
    GetScriptOwner = m_SelectedElement.CommandGroup.ScriptName
End Function

'// When a node in the treeview is clicked, it should locate the XML element that was
'// used to create the node and call this method to populate appropriate form controls.
'// 08/29/2008 JSM - Created
Private Sub PrepareCommandForm()

    Dim Parameter As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj

    Dim requirements As String
    Dim sItem As String
    Dim i As Integer

    m_ResettingForm = True

    With m_SelectedElement

        Call m_Commands.OpenCommand(.CommandName, GetScriptOwner(), False, m_DocumentCopy)

        '// lblSyntax
        lblSyntax.Caption = m_Commands.SyntaxString
        '// lblRequirements
        lblRequirements.Caption = m_Commands.RequirementsString

        Select Case .TheNodeType
            Case NodeType.nCommand
                '// txtRank
                txtRank.Enabled = True
                lblRank.Enabled = True
                If m_Commands.RequiredRank < 0 Then
                    txtRank.Text = vbNullString
                Else
                    txtRank.Text = CStr(m_Commands.RequiredRank)
                End If
                '// txtFlags
                txtFlags.Enabled = True
                lblFlags.Enabled = True
                txtFlags.Text = m_Commands.RequiredFlags
                '// lvAliases
                lvAliases.Visible = True
                lblAlias.Visible = True
                txtAddAlias.Visible = True
                cmdAddAlias.Visible = True
                lvAliases.ListItems.Clear
                For i = 1 To m_Commands.Aliases.Count
                    lvAliases.ListItems.Add , , m_Commands.Aliases(i)
                Next i
                '// txtDescription
                txtDescription.Enabled = True
                lblDescription.Enabled = True
                txtDescription.Text = PrepString(m_Commands.Description)
                '// txtSpecialNotes
                txtSpecialNotes.Enabled = True
                lblSpecialNotes.Enabled = True
                txtSpecialNotes.Text = PrepString(m_Commands.SpecialNotes)
                '// chkDisable
                chkDisable.Enabled = True
                chkDisable.Visible = True
                chkDisable.Value = Abs(CInt(Not m_Commands.IsEnabled))
                chkDisable.Caption = StringFormat("Disable {0} command", .CommandName)
                '// only allow deleting script commands
                cmdDeleteCommand.Visible = True
                cmdDeleteCommand.Enabled = (cboCommandGroup.ListIndex > 0)

            Case NodeType.nParameter
                Set Parameter = m_Commands.GetParameterByName(.ParameterName)

                '// txtDescription
                txtDescription.Enabled = True
                lblDescription.Enabled = True
                txtDescription.Text = PrepString(Parameter.Description)
                '// txtSpecialNotes
                txtSpecialNotes.Enabled = True
                lblSpecialNotes.Enabled = True
                txtSpecialNotes.Text = PrepString(Parameter.SpecialNotes)
                '// txtMatch
                txtMatch.Visible = True
                lblMatch.Visible = True
                txtMatch.Text = Parameter.MatchMessage
                lblMatch.Caption = "Parameter must &match"
                '// chkMatchCase
                chkMatchCase.Visible = True
                chkMatchCase.Value = Abs(CInt(Parameter.MatchCaseSensitive))
                '// txtMatchError
                txtMatchError.Visible = True
                lblMatchError.Visible = True
                txtMatchError.Text = Parameter.MatchError
                lblMatchError.Caption = "&Error on no match"
                '// custom captions
                lblRequirements.Caption = StringFormat("Rank and flags requirements do not apply to individual parameters.{1}{0}", lblRequirements.Caption, vbNewLine)
                cmdDeleteCommand.Visible = False

            Case NodeType.nRestriction

                Set Parameter = m_Commands.GetParameterByName(.ParameterName)
                Set Restriction = Parameter.GetRestrictionByName(.RestrictionName)

                '// txtDescription
                txtDescription.Enabled = True
                lblDescription.Enabled = True
                txtDescription.Text = PrepString(Restriction.Description)
                '// txtSpecialNotes
                txtSpecialNotes.Enabled = True
                lblSpecialNotes.Enabled = True
                txtSpecialNotes.Text = PrepString(Restriction.SpecialNotes)
                '// txtMatch
                txtMatch.Visible = True
                lblMatch.Visible = True
                txtMatch.Text = Restriction.MatchMessage
                lblMatch.Caption = "Restrict if &matches"
                '// chkMatchCase
                chkMatchCase.Visible = True
                chkMatchCase.Value = Abs(CInt(Restriction.MatchCaseSensitive))
                '// chkMatchErrorGlobal
                chkMatchErrorGlobal.Visible = True
                chkMatchErrorGlobal.Value = vbUnchecked
                '// txtMatchError
                txtMatchError.Visible = True
                lblMatchError.Visible = True
                txtMatchError.Text = Restriction.MatchError
                If LenB(Restriction.MatchError) = 0 And LenB(Parameter.RestrictionsSharedError) > 0 Then
                    txtMatchError.Text = Parameter.RestrictionsSharedError
                    chkMatchErrorGlobal.Value = vbChecked
                End If
                lblMatchError.Caption = "&Error on restricted"
                '// chkMatchFatal
                chkMatchFatal.Visible = True
                chkMatchFatal.Value = Abs(CInt(Restriction.Fatal))
                '// txtRank
                txtRank.Enabled = True
                lblRank.Enabled = True
                If Restriction.RequiredRank < 0 Then
                    txtRank.Text = vbNullString
                Else
                    txtRank.Text = CStr(Restriction.RequiredRank)
                End If
                '// txtFlags
                txtFlags.Enabled = True
                lblFlags.Enabled = True
                txtFlags.Text = Restriction.RequiredFlags
                '// special captions
                lblRequirements.Caption = StringFormat("This is the rank and flags required to bypass this restriction.{1}{0}", lblRequirements.Caption, vbNewLine)
                cmdDeleteCommand.Visible = False
        End Select
    End With

    '// Disable our buttons
    cmdSaveCommand.Enabled = False
    cmdDiscardCommand.Enabled = False

    fraCommand.Caption = GetCaptionString()
    Me.Caption = "Command Manager - " & fraCommand.Caption

    m_ResettingForm = False

End Sub

Private Function GetCaptionString() As String
    Dim Parameter As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj

    With m_SelectedElement
        Select Case .TheNodeType
            Case nCommand
                GetCaptionString = m_Commands.ToString
            Case nParameter
                Set Parameter = m_Commands.GetParameterByName(.ParameterName)
                GetCaptionString = StringFormat("{0} -> {1}", m_Commands.ToString, Parameter.ToString(True))
            Case nRestriction
                Set Parameter = m_Commands.GetParameterByName(.ParameterName)
                Set Restriction = Parameter.GetRestrictionByName(.RestrictionName)
                GetCaptionString = StringFormat("{0} -> {1} -> {2}", m_Commands.ToString, Parameter.ToString(True), Restriction.ToString)
        End Select
        If .IsDirty Then GetCaptionString = GetCaptionString & " *"
    End With
End Function

'// Clears and disables all edit controls. Treeview is left intact
'// 08/29/2008 JSM - Created
Private Sub ResetCommandForm()

    m_ResettingForm = True

    txtRank.Text = vbNullString
    lvAliases.ListItems.Clear
    txtAddAlias.Text = vbNullString
    txtFlags.Text = vbNullString
    txtDescription.Text = vbNullString
    txtSpecialNotes.Text = vbNullString
    fraCommand.Caption = vbNullString
    chkDisable.Value = vbUnchecked

    lblAlias.Visible = False
    lvAliases.Visible = False
    txtAddAlias.Visible = False
    cmdAddAlias.Visible = False
    chkDisable.Visible = False

    lblMatch.Visible = False
    lblMatchError.Visible = False
    txtMatch.Visible = False
    txtMatchError.Visible = False
    chkMatchCase.Visible = False
    chkMatchErrorGlobal.Visible = False
    chkMatchFatal.Visible = False

    lblRank.Enabled = False
    lblFlags.Enabled = False
    txtRank.Enabled = False
    txtFlags.Enabled = False

    lblDescription.Enabled = False
    lblSpecialNotes.Enabled = False
    txtDescription.Enabled = False
    txtSpecialNotes.Enabled = False

    m_SelectedElement.IsDirty = False
    cmdSaveCommand.Enabled = False
    cmdDiscardCommand.Enabled = False
    cmdDeleteCommand.Visible = False

    lblSyntax.Caption = vbNullString
    lblSyntax.ForeColor = RTBColors.ConsoleText

    fraCommand.Caption = vbNullString
    Me.Caption = "Command Manager"

    m_ResettingForm = False

End Sub

Private Sub txtAddAlias_KeyPress(KeyAscii As Integer)
    ' disallow entering space
    If (KeyAscii = vbKeySpace) Then KeyAscii = 0
End Sub

Private Sub txtAddAlias_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    '// Delete
    If KeyCode = vbKeyDelete Then
        If txtAddAlias.SelStart + 1 = Len(txtAddAlias.Text) Then
            '// If the current text is already in the list, lets delete it. Otherwise,
            '// this code should behave like a normal delete keypress.
            For i = 1 To lvAliases.ListItems.Count
                If StrComp(lvAliases.ListItems.Item(i), txtAddAlias.Text, vbTextCompare) = 0 Then
                    lvAliases.ListItems.Remove i
                    txtAddAlias.Text = vbNullString
                    KeyCode = 0
                    Call FormIsDirty
                    Exit Sub
                End If
            Next i
        End If
    End If
End Sub

'// Must mark the element as dirty when these change
Private Sub txtFlags_KeyPress(KeyAscii As Integer)
    ' disallow entering space
    If (KeyAscii = vbKeySpace) Then KeyAscii = 0
    
    ' if key is A-Z, then make uppercase
    If (InStr(1, AZ, ChrW$(KeyAscii), vbTextCompare) > 0) Then
        If (BotVars.CaseSensitiveFlags = False) Then
            If (KeyAscii > vbKeyZ) Then ' lowercase if greater than "Z"
                KeyAscii = AscW(UCase$(ChrW$(KeyAscii)))
            End If
        End If
        ' disallow repeating a flag already present
        If (InStr(1, txtFlags.Text, ChrW$(KeyAscii), vbBinaryCompare) > 0) Then
            KeyAscii = 0
        End If
    ' else disallow entering that character (if not a control character)
    ElseIf (KeyAscii > vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFlags_Change()
    ' enable entry save button
    Call FormIsDirty
End Sub

Private Sub txtRank_KeyPress(KeyAscii As Integer)
    ' disallow entering space
    If (KeyAscii = vbKeySpace) Then KeyAscii = 0
    
    ' if key is not 0-9, disallow entering that character (if not a control character)
    If (InStr(1, Num09, ChrW$(KeyAscii), vbTextCompare) = 0 And KeyAscii > vbKeySpace) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRank_Change()
    Dim SelStart As Long

    If StrictIsNumeric(txtRank.Text) Then
        If (CInt(txtRank.Text) > 200) Then
            With txtRank
                SelStart = .SelStart
                .Text = "200"
                .SelStart = SelStart
            End With
        End If
    End If

    ' enable entry save button
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

Private Sub txtMatch_Change()
    Call FormIsDirty
End Sub

Private Sub txtMatchError_Change()
    Call FormIsDirty
End Sub

Private Sub chkMatchCase_Click()
    Call FormIsDirty
End Sub

Private Sub chkMatchErrorGlobal_Click()
    Call FormIsDirty
End Sub

Private Sub chkMatchFatal_Click()
    Call FormIsDirty
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaveForm_Click()
    ' if the current focus is on the Add Alias textbox and user hits enter, they mean to add an alias...
    If (Len(txtAddAlias.Text) > 0) And (Me.ActiveControl Is txtAddAlias) Then
        Call cmdAddAlias_Click
        Exit Sub
    End If

    Call SaveCommandForm
    Call m_DocumentCopy.Save(m_Commands.XMLDocument)
    Call m_Commands.Save(True)
    Unload Me
End Sub
