VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.0#0"; "vbalTreeView6.ocx"
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

Private m_CommandsDoc As DOMDocument60
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
    TheXMLElement As IXMLDOMElement
    IsDirty As Boolean
    commandName As String
    argumentName As String
    restrictionName As String
End Type


'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
'    (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
'    ByVal lParam As Long) As Long
'Private Const WM_SETREDRAW As Long = &HB
'Private Const TV_FIRST As Long = &H1100
'Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
'Private Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
'Private Const TVGN_ROOT As Long = &H0


' Quicky clear the treeview identified by the hWnd parameter
Sub ClearTreeViewNodes(ByRef trv As vbalTreeView)
    
    m_ClearingNodes = True
    trv.nodes.Clear
    m_ClearingNodes = False
    
    '// Below code is no longer necesarry thanks to a better treeview. :) -Pyro
    'Dim hWnd As Long
    'Dim hItem As Long
    '
    'hWnd = trv.hWnd
    '
    '
    '' lock the window update to avoid flickering
    'SendMessageLong hWnd, WM_SETREDRAW, False, &O0
    '
    '' clear the treeview
    'Do
    '    hItem = SendMessageLong(hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    '    If hItem <= 0 Then Exit Do
    '    SendMessageLong hWnd, TVM_DELETEITEM, &O0, hItem
    'Loop
    '
    '' unlock the window
    'SendMessageLong hWnd, WM_SETREDRAW, True, &O0
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

Private Sub cmdAliasAdd_Click()

    ' ...
    cboAlias.AddItem cboAlias.Text
    
    ' ...
    cboAlias.Text = ""
    
    Call FormIsDirty

End Sub

Private Sub cmdAliasRemove_Click()

    Dim i As Integer ' ...
    
    ' ...
    For i = 0 To cboAlias.ListCount - 1
        If (StrComp(cboAlias.Text, cboAlias.List(i), vbTextCompare) = 0) Then
            cboAlias.RemoveItem i
            
            Exit For
        End If
    Next i
    
    ' ...
    cboAlias.Text = ""
    
    Call FormIsDirty

End Sub


'// 08/30/2008 JSM - Created
Private Sub cmdDiscard_Click()
    Call PrepareForm(m_SelectedElement.TheNodeType, m_SelectedElement.TheXMLElement)
End Sub

Private Sub cmdDeleteCommand_Click()
    
    Dim scriptName As String
    Dim doc As clsCommandDocObj

    scriptName = Mid$(cboCommandGroup.Text, 1, InStr(1, cboCommandGroup.Text, "(") - 2)


    If vbYes <> MsgBox(StringFormat("Are you sure you want to delete the {0} command for the {1} script?", m_SelectedElement.commandName, scriptName), vbYesNo + vbQuestion, frmCommands.Caption) Then
        Exit Sub
    End If
    
    
    
    Set doc = New clsCommandDocObj
    
    Call doc.OpenCommand(m_SelectedElement.commandName, scriptName)
    Call doc.Delete
    Set doc = Nothing
    
    m_SelectedElement.IsDirty = False
    
    Set m_CommandsDoc = Nothing
    Set m_CommandsDoc = New DOMDocument60
    Call m_CommandsDoc.Load(App.Path & "\commands.xml")
    
    Call PopulateOwnerComboBox
    Call ResetForm
    Call PopulateTreeView(scriptName)
    
    

End Sub



Private Sub cmdFlagAdd_Click()

    ' ...
    cboFlags.AddItem cboFlags.Text
    
    ' ...
    cboFlags.Text = ""
    
    Call FormIsDirty

End Sub

Private Sub cmdFlagRemove_Click()

    Dim i As Integer ' ...
    
    ' ...
    For i = 0 To cboFlags.ListCount - 1
        If (StrComp(cboFlags.Text, cboFlags.List(i), vbBinaryCompare) = 0) Then
            cboFlags.RemoveItem i
            
            Exit For
        End If
    Next i
    
    ' ...
    cboFlags.Text = ""
    
    Call FormIsDirty

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
    
    
    If Not clsCommandDocObj.ValidateXMLFromFiles(App.Path & "\commands.xml", App.Path & "\commands.xsd") Then
        Exit Sub
    End If
    
    
    Call m_CommandsDoc.Load(App.Path & "\commands.xml")
    
    'Change tree view background and foreground color.
    ' // (REMOVED 8/9/09: changed to a better treeview -Pyro)
    'Dim lStyle As Long
    'Dim tNode As node
    
    'For Each tNode In trvCommands.nodes
    '    tNode.BackColor = txtRank.BackColor
    'Next
    
    'SendMessage trvCommands.hWnd, 4381&, 0, txtRank.BackColor
    'lStyle = GetWindowLong(trvCommands.hWnd, -16&)
    'SetWindowLong trvCommands.hWnd, -16&, lStyle And (Not 2&)
    'SetWindowLong trvCommands.hWnd, -16&, lStyle
    
    Call ResetForm
    Call PopulateOwnerComboBox
    Call PopulateTreeView
    
    Exit Sub
    
ErrorHandler:

    MsgBox Err.description, vbCritical + vbOKOnly, Me.Caption
    Call ResetForm
    '// Disable our buttons
    cmdSave.enabled = False
    cmdDiscard.enabled = False
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_CommandsDoc = Nothing
End Sub


Private Sub PopulateOwnerComboBox()
    
    On Error Resume Next

    Dim i   As Integer
    Dim str As String
    Dim commandCount As Integer
    Dim scriptName As String
    Dim commandDoc As clsCommandDocObj
    Dim options() As Variant '// <-- boo
    
    Set commandDoc = New clsCommandDocObj

    cboCommandGroup.Clear
    
    
    '// get the script name and number of commands
    scriptName = "Internal Bot Commands"
    commandCount = commandDoc.GetCommandCount()
    options = Array(scriptName, commandCount)
    '// add the item
    cboCommandGroup.AddItem StringFormatA("{0} ({1})", options)
    
    For i = 2 To frmChat.SControl.Modules.Count
        scriptName = _
            modScripting.GetScriptName(CStr(i))
        str = _
            SharedScriptSupport.GetSettingsEntry("Public", scriptName)
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            '// get the script name and number of commands
            commandCount = commandDoc.GetCommandCount(scriptName)
            '// only add the commands if there is at least 1 command to show
            If commandCount > 0 Then
                options = Array(scriptName, commandCount)
                '// add the item
                cboCommandGroup.AddItem StringFormatA("{0} ({1})", options)
            End If
        End If
    Next i
    cboCommandGroup.ListIndex = 0
    
End Sub



Private Sub PopulateTreeView(Optional strScriptOwner As String = vbNullString)
    
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
    Dim X                 As Integer

    '// reset the treeview
    If trvCommands.nodes.Count > 0 Then
        trvCommands.nodes(1).Selected = True
    End If
    
    Call ClearTreeViewNodes(trvCommands)
    
    '// create xpath expression based on strScriptOwner
    If strScriptOwner = vbNullString Then
        xpath = "/commands/command[not(@owner)]"
        'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
    Else
        xpath = StringFormat("/commands/command[@owner='{0}']", strScriptOwner)
        'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
    End If
    
    '// get a list of all the commands
    Set commandNodes = m_CommandsDoc.documentElement.selectNodes(xpath)
    ReDim commandNameArray(commandNodes.length)
    
    
    '// read them 1 at a time and add them to an array
    X = 0
    For Each xmlCommand In m_CommandsDoc.documentElement.selectNodes(xpath)
        commandNameArray(X) = xmlCommand.Attributes.getNamedItem("name").Text
        X = X + 1
    Next
    
    '// sort the command names
    Call BubbleSort1(commandNameArray)
    

    '// loop through the sorted array and select the commands
    For X = LBound(commandNameArray) To UBound(commandNameArray)

        commandName = commandNameArray(X)
        If Len(commandName) > 0 Then
            '// create xpath expression based on strScriptOwner
            If strScriptOwner = vbNullString Then
                xpath = StringFormat("/commands/command[@name='{0}' and not(@owner)]", commandName)
                'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , "Internal Commands")
            Else
                xpath = StringFormat("/commands/command[@name='{0}' and @owner='{1}']", commandName, strScriptOwner)
                'Set nRoot = trvCommands.Nodes.Add(, etvwFirst, , strScriptOwner & " Commands")
            End If
    
            Set xmlCommand = m_CommandsDoc.documentElement.selectSingleNode(xpath)
        
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
    Next
    
    '// 08/30/2008 JSM - click the first command alphabetically
    ' fixed to work with SelectedNodeChanged() -Ribose/2009-08-10
    If Not (defaultNode Is Nothing) Then
        defaultNode.Selected = True
    Else
        trvCommands_SelectedNodeChanged
    End If
    
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
        options = Array(.commandName, .argumentName, .restrictionName)
        Select Case .TheNodeType
            Case NodeType.nCommand
                sMessage = StringFormatA("You have not saved your changes to {0}. Do you want to save them now?", options)
            Case NodeType.nArgument
                sMessage = StringFormatA("You have not saved your changes to {1}. Do you want to save them now?", options)
            Case NodeType.nRestriction
                sMessage = StringFormatA("You have not saved your changes to {2}. Do you want to save them now?", options)
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

    Dim node As cTreeViewNode
    Dim nt As NodeType
    Dim commandName As String
    Dim argumentName As String
    Dim restrictionName As String
    Dim options() As Variant '// <-- boo
    
    Dim xpath As String
    Dim xmlElement As IXMLDOMElement
    
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
    '// create an array for the StringFormat function, this function will replace
    '// the {0} {1} and {2} with their respective Values() found below
    '//                {0}           {1}             {2}
    options = Array(commandName, argumentName, restrictionName)
    
    Select Case nt
        Case NodeType.nCommand
            xpath = StringFormatA("/commands/command[@name='{0}']", options)
        Case NodeType.nArgument
            xpath = StringFormatA("/commands/command[@name='{0}']/arguments/argument[@name='{1}']", options)
        Case NodeType.nRestriction
            xpath = StringFormatA("/commands/command[@name='{0}']/arguments/argument[@name='{1}']/restrictions/restriction[@name='{2}']", options)
    End Select
    
    '// Update m_SelectedElement so we know which element we are viewing
    Let m_SelectedElement.commandName = commandName
    Let m_SelectedElement.argumentName = argumentName
    Let m_SelectedElement.restrictionName = restrictionName
    
    '// grab the node from the xpath
    Set xmlElement = m_CommandsDoc.selectSingleNode(xpath)
    
    Call ResetForm
    
    '// if there are no comment elements then we do not need to prepare the form
    If Not (xmlElement Is Nothing) Then
        Call PrepareForm(nt, xmlElement)
    End If
    
End Sub

'// Call this sub whenever the form controls have been changed
'// 08/30/2008 JSM - Created
Private Sub FormIsDirty()
    m_SelectedElement.IsDirty = True
    cmdSave.enabled = True
    cmdDiscard.enabled = True
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
            
            If (txtRank.Text <> vbNullString) Then
                xmlNode.Text = txtRank.Text
            Else
                For i = 0 To xmlNode.childNodes.length - 1
                    xmlNode.removeChild xmlNode.childNodes(i)
                Next i
            End If
        End If
        
        '// txtDescription
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nArgument Then
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
            xmlNode.Text = txtDescription.Text
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
            xmlNode.Text = txtSpecialNotes.Text
        End If

        '// cboAlias
        If .TheNodeType = NodeType.nCommand Then
            For Each xmlNode In .TheXMLElement.selectNodes("aliases/alias")
                .TheXMLElement.selectSingleNode("aliases").removeChild xmlNode
            Next xmlNode
            For i = 0 To cboAlias.ListCount - 1
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "alias", "")
                xmlNewNode.Text = cboAlias.List(i)
                .TheXMLElement.selectSingleNode("aliases").appendChild xmlNewNode
            Next i
        End If
        
        '// cboFlags
        If .TheNodeType = NodeType.nCommand Or .TheNodeType = NodeType.nRestriction Then
            
            '// remove existing flags
            For Each xmlNode In .TheXMLElement.selectNodes("access/flags")
                .TheXMLElement.selectSingleNode("access").removeChild xmlNode
            Next xmlNode
            
            '// make sue the access element exists
            Set xmlNode = .TheXMLElement.selectSingleNode("access")
            If xmlNode Is Nothing Then
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "access", "")
                Set xmlNode = .TheXMLElement.appendChild(xmlNewNode)
            End If
            
            Set xmlNode = .TheXMLElement.selectSingleNode("access/flags")
            If xmlNode Is Nothing Then
                Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "flags", "")
                Set xmlNode = _
                    .TheXMLElement.selectSingleNode("access").appendChild(xmlNewNode)
            End If

            If (cboFlags.ListCount > 0) Then
                '// loop through cboFlags and add the text
                For i = 0 To cboFlags.ListCount - 1
                    Set xmlNewNode = xmlNode.selectSingleNode("flag[text()='" & cboFlags.List(i) & "']")
                    
                    If (xmlNewNode Is Nothing) Then
                        Set xmlNewNode = m_CommandsDoc.createNode(NODE_ELEMENT, "flag", "")
                        xmlNewNode.Text = cboFlags.List(i)
                        xmlNode.appendChild xmlNewNode
                    End If
                Next i
            End If
        End If
        
        
        '// chkDisable
        If .TheNodeType = NodeType.nCommand Then
            If chkDisable.Value = 1 Then
                Call .TheXMLElement.setAttribute("enabled", "false")
            Else
                If (.TheXMLElement.getAttribute("enabled") <> vbNullString) Then
                    Call .TheXMLElement.setAttribute("enabled", "true")
                End If
            End If
        End If
        
        '// 08/302009 52 - getting rid of the clsXML class since it didnt write valid XML
		'// TODO: this needs to validate!!!
        Call m_CommandsDoc.Save(App.Path & "\commands.xml")
        
        'Call m_CommandsDoc.Save(App.Path & "\commands.xml")
        Call PrepareForm(.TheNodeType, .TheXMLElement)
        
    End With
    

    
    
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
Private Sub PrepareForm(nt As NodeType, xmlElement As IXMLDOMElement)
    
    Dim xmlNode As IXMLDOMNode
    Dim options() As Variant '// <-- boo
    Dim requirements As String
    
    
    Dim cmd As clsCommandDocObj
    Set cmd = New clsCommandDocObj
    
    Call cmd.OpenCommand(m_SelectedElement.commandName, getScriptOwner())
    
    '// lblSyntax
    lblSyntax.Caption = cmd.SyntaxString
    
    '// lblRequirements
    lblRequirements.Caption = cmd.RequirementsString
    
    
    
    
    options = Array(m_SelectedElement.commandName, m_SelectedElement.argumentName, m_SelectedElement.restrictionName)
    
    
    
    Select Case nt
        Case NodeType.nCommand
            '// txtRank
            txtRank.enabled = True
            lblRank.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("access/rank")
            If Not (xmlNode Is Nothing) Then
                txtRank.Text = xmlNode.Text
            Else
                txtRank.Text = vbNullString
            End If
            '// cboAlias
            cboAlias.enabled = True
            lblAlias.enabled = True
            cmdAliasAdd.enabled = True
            cmdAliasRemove.enabled = True
            For Each xmlNode In xmlElement.selectNodes("aliases/alias")
                cboAlias.AddItem xmlNode.Text
            Next xmlNode
            '// cboFlags
            cboFlags.enabled = True
            lblFlags.enabled = True
            cmdFlagAdd.enabled = True
            cmdFlagRemove.enabled = True
            For Each xmlNode In xmlElement.selectNodes("access/flags/flag")
                cboFlags.AddItem xmlNode.Text
            Next xmlNode
            If (cboFlags.ListCount) Then
                cboFlags.Text = cboFlags.List(0)
            End If
            
            '// txtDescription
            txtDescription.enabled = True
            lblDescription.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.Text = PrepString(xmlNode.Text)
            End If
            '// txtSpecialNotes
            txtSpecialNotes.enabled = True
            lblSpecialNotes.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.Text = PrepString(xmlNode.Text)
            End If
            '// chkDisable
            chkDisable.enabled = True
            chkDisable.Visible = True
            If LCase(xmlElement.getAttribute("enabled")) = "false" Then
                chkDisable.Value = 1
            Else
                chkDisable.Value = 0
            End If
            
            '// custom captions
            fraCommand.Caption = StringFormatA("{0}", options)
            chkDisable.Caption = StringFormatA("Disable {0} command", options)
            
            If cboCommandGroup.ListIndex > 0 Then
                cmdDeleteCommand.enabled = True
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
            txtDescription.enabled = True
            lblDescription.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.Text = PrepString(xmlNode.Text)
            End If
            '// txtSpecialNotes
            txtSpecialNotes.enabled = True
            lblSpecialNotes.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/specialnotes")
            If Not (xmlNode Is Nothing) Then
                txtSpecialNotes.Text = PrepString(xmlNode.Text)
            End If

            '// chkDisable
            'chkDisable.Enabled = True
            'chkDisable.Visible = True
            
            '// special captions
            If (Not xmlElement.Attributes.getNamedItem("optional") Is Nothing) Then
                If (xmlElement.Attributes.getNamedItem("optional").Text = "1") Then
                    fraCommand.Caption = StringFormatA("{0} => {1} - Optional", options)
                End If
            Else
                fraCommand.Caption = StringFormatA("{0} => {1}", options)
            End If
            
            
            
            chkDisable.Caption = StringFormatA("Disable {1} argument", options)
            
            
            
            
        Case NodeType.nRestriction
            '// txtRank
            txtRank.enabled = True
            lblRank.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("access/rank")
            If Not (xmlNode Is Nothing) Then
                txtRank.Text = xmlNode.Text
            End If
            '// cboAlias
            'cboAlias.Enabled = True
            'lblAlias.Enabled = True
            'For Each xmlNode In xmlElement.selectNodes("aliases/alias")
            '    cboAlias.AddItem xmlNode.text
            'Next xmlNode
            '// cboFlags
            cboFlags.enabled = True
            lblFlags.enabled = True
            For Each xmlNode In xmlElement.selectNodes("access/flags/flag")
                cboFlags.AddItem xmlNode.Text
            Next xmlNode
            '// txtDescription
            txtDescription.enabled = True
            lblDescription.enabled = True
            Set xmlNode = xmlElement.selectSingleNode("documentation/description")
            If Not (xmlNode Is Nothing) Then
                txtDescription.Text = PrepString(xmlNode.Text)
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
            
            '// special captions
            If (Not xmlElement.parentNode.parentNode.Attributes.getNamedItem("optional") Is Nothing) Then
                If (xmlElement.parentNode.parentNode.Attributes.getNamedItem("optional").Text = "1") Then
                    fraCommand.Caption = StringFormatA("{0} => {1} - Optional => {2}", options)
                End If
            Else
                fraCommand.Caption = StringFormatA("{0} => {1} => {2}", options)
            End If
            
            chkDisable.Caption = StringFormatA("Disable {2} restriction", options)
            
    End Select
    
    '// Update m_SelectedElement so we know which element we are viewing
    Set m_SelectedElement.TheXMLElement = xmlElement
    Let m_SelectedElement.TheNodeType = nt
    Let m_SelectedElement.IsDirty = False
    
    '// Disable our buttons
    cmdSave.enabled = False
    cmdDiscard.enabled = False

End Sub

'// Clears and disables all edit controls. Treeview is left intact
'// 08/29/2008 JSM - Created
Private Sub ResetForm()
    
    txtRank.Text = ""
    cboAlias.Clear
    cboFlags.Clear
    txtDescription.Text = ""
    txtSpecialNotes.Text = ""
    fraCommand.Caption = ""
    chkDisable.Value = 0
    
    txtRank.enabled = False
    cboAlias.enabled = False
    cboFlags.enabled = False
    txtDescription.enabled = False
    txtSpecialNotes.enabled = False
    chkDisable.enabled = False
    
    lblRank.enabled = False
    lblAlias.enabled = False
    lblFlags.enabled = False
    lblDescription.enabled = False
    lblSpecialNotes.enabled = False
    
    cmdAliasAdd.enabled = False
    cmdAliasRemove.enabled = False
    cmdFlagAdd.enabled = False
    cmdFlagRemove.enabled = False
    
    chkDisable.Visible = False
    
    
    m_SelectedElement.IsDirty = False
    cmdSave.enabled = False
    cmdDiscard.enabled = False
    cmdDeleteCommand.enabled = False
    
    lblSyntax.Caption = ""
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

