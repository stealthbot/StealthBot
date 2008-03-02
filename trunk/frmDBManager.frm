VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDBManager 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmDBManager"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2880
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin VB.Frame frmDatabase 
      Caption         =   "Eric[nK]"
      Height          =   4950
      Left            =   3600
      TabIndex        =   6
      Top             =   487
      Width           =   3025
      Begin VB.ListBox lstGroups 
         Height          =   2010
         Left            =   240
         TabIndex        =   20
         Top             =   2450
         Width           =   2535
      End
      Begin VB.TextBox txtFlags 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   8
         Top             =   580
         Width           =   1215
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   7
         Top             =   580
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1930
         TabIndex        =   9
         Top             =   4535
         Width           =   855
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1088
         TabIndex        =   10
         Top             =   4535
         Width           =   855
      End
      Begin VB.Label lblModifiedBy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   480
         TabIndex        =   19
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Label lblCreatedBy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   480
         TabIndex        =   18
         Top             =   1350
         Width           =   2415
      End
      Begin VB.Label lblCreatedOn 
         Caption         =   "(not applicable)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   130
         Left            =   360
         TabIndex        =   14
         Top             =   1180
         Width           =   2415
      End
      Begin VB.Label lblModifiedOn 
         Caption         =   "(not applicable)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Created on:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   970
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Last modified on:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1590
         Width           =   2535
      End
      Begin VB.Label lblGroup 
         Caption         =   "Member of Group(s):"
         Height          =   255
         Left            =   245
         TabIndex        =   13
         Top             =   2240
         Width           =   2535
      End
      Begin VB.Label lblFlags 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   340
         Width           =   1215
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank (1 - 200):"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   340
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList icons 
      Left            =   4080
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0552
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0AA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnCreateGroup 
      Caption         =   "Create Group"
      Height          =   375
      Left            =   1800
      Picture         =   "frmDBManager.frx":0FF6
      TabIndex        =   2
      ToolTipText     =   "Create Group"
      Top             =   5047
      Width           =   1695
   End
   Begin VB.CommandButton btnCreateUser 
      Caption         =   "Create User"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "frmDBManager.frx":145E
      TabIndex        =   1
      ToolTipText     =   "Create User"
      Top             =   5047
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tbsTabs 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   1
      Separators      =   -1  'True
      TabMinWidth     =   176
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Users and Groups"
            Object.ToolTipText     =   "User entries identify individual users which can be grouped for easier control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clans"
            Object.ToolTipText     =   "Clan entries allow access to be given based on WarCraft III clan membership"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Games"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Apply and Cl&ose"
      Height          =   300
      Index           =   0
      Left            =   5280
      TabIndex        =   4
      Top             =   5540
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   5540
      Width           =   735
   End
   Begin MSComctlLib.TreeView trvUsers 
      Height          =   4370
      Left            =   120
      TabIndex        =   0
      Top             =   577
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7699
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   575
      LineStyle       =   1
      Style           =   7
      ImageList       =   "icons"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDBManager.frm
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Public m_game          As String

Private m_DB()         As udtDatabase
Private m_modified     As Boolean
Private m_new_entry    As Boolean
Private m_DBDate       As Long
Private m_group_index  As Integer
Private m_group_change As Boolean

' ...
Private Sub Form_Load()
    ' has our database been loaded?
    If (DB(0).Username = vbNullString) Then
        ' load database, if for some reason that hasn't been done
        Call LoadDatabase
    End If
    
    ' store temporary copy of database
    m_DB() = DB()
    
    ' show database for default tab
    Call tbsTabs_Click
End Sub ' end function Form_Load

' ...
Private Sub btnCreateUser_Click()
    Static userCount As Integer ' ...
    
    Dim newNode      As node    ' ...
    Dim gAcc         As udtGetAccessResponse
    Dim Username     As String  ' ...
    
    ' ...
    Do
        ' ...
        Username = "New User #" & (userCount + 1)
    Loop While (GetAccess(Username, "User").Username <> vbNullString)
        
    ' redefine array to support new entry
    ReDim Preserve m_DB(UBound(m_DB) + 1)
    
    ' create new database entry
    With m_DB(UBound(m_DB))
        .Username = Username
        .Type = "USER"
        .AddedBy = "(console)"
        .AddedOn = Now
        .ModifiedBy = "(console)"
        .ModifiedOn = Now
    End With

    ' do we have an item (hopefully a group) selected?
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' is the item really just the root item?
        If (trvUsers.SelectedItem.index = 1) Then
            Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                tvwChild, "User: " & Username, Username, 3)
                
        Else
            ' is the item a group?
            If (StrComp(trvUsers.SelectedItem.tag, "Group", vbTextCompare) = 0) Then
                ' create new node under group node
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, tvwChild, _
                    "User: " & Username, Username, 3)
    
                ' ...
                With m_DB(UBound(m_DB))
                    .Groups = trvUsers.SelectedItem.text
                End With
            Else
                ' is our parent a group?
                If (StrComp(trvUsers.SelectedItem.Parent.tag, "Group", vbTextCompare) = 0) Then
                    ' create new node under group node
                    Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Parent.Key, tvwChild, _
                        "User: " & Username, Username, 3)
                
                    ' set group settings on new database entry
                    With m_DB(UBound(m_DB))
                        .Groups = trvUsers.SelectedItem.Parent.text
                    End With
                Else
                    ' create new node under root
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, "User: " & _
                        Username, Username, 3)
                End If
            End If
        End If
    Else
        ' lets just create the node under the root node
        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
            "User: " & Username, Username, 3)
    End If
    
    ' change misc. settings
    With newNode
        .tag = "User"
        .Selected = True
    End With
    
    ' ...
    m_new_entry = True
    
    ' open entry name for editing
    Call trvUsers.StartLabelEdit
    
    ' increment user count
    userCount = (userCount + 1)
End Sub

' ...
Private Sub btnCreateGroup_Click()
    Static groupCount As Integer ' ...
    Static clanCount  As Integer ' ...
    
    Dim newNode       As node    ' ...
    
    ' ...
    If (tbsTabs.SelectedItem.index = 1) Then ' Users and Groups Tab
        Dim GroupName As String ' ...
    
        ' ...
        Do
            ' ...
            GroupName = "New Group #" & (groupCount + 1)
        Loop While (GetAccess(GroupName, "Group").Username <> vbNullString)
        
        ' ...
        ReDim Preserve m_DB(UBound(m_DB) + 1)
        
        ' ...
        With m_DB(UBound(m_DB))
            .Username = GroupName
            .Type = "GROUP"
            .AddedBy = "(console)"
            .AddedOn = Now
            .ModifiedBy = "(console)"
            .ModifiedOn = Now
        End With
    
        ' do we have an item (hopefully a group) selected?
        If (Not (trvUsers.SelectedItem Is Nothing)) Then
            ' is the item reall just the root node?
            If (trvUsers.SelectedItem.index = 1) Then
                ' ...
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                    tvwChild, "Group: " & GroupName, GroupName, 1)
            Else
                ' ...
                If (StrComp(trvUsers.SelectedItem.tag, "Group", vbTextCompare) = 0) Then
                    ' ...
                    Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                        tvwChild, "Group: " & GroupName, GroupName, 1)
    
                    ' ...
                    With m_DB(UBound(m_DB))
                        .Groups = trvUsers.SelectedItem.text
                    End With
                Else
                    ' ...
                    Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Parent.Key, _
                        tvwChild, "Group: " & GroupName, GroupName, 1)
                        
                    ' ...
                    If (StrComp(trvUsers.SelectedItem.Parent.tag, "Group", vbTextCompare) = 0) Then
                        ' ...
                        With m_DB(UBound(m_DB))
                            .Groups = trvUsers.SelectedItem.Parent.text
                        End With
                    End If
                End If
            End If
        Else
            ' ...
            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                "Group: " & GroupName, GroupName, 1)
        End If
        
        ' change misc. settings
        With newNode
            .tag = "Group"
        End With
        
        ' increment group counter
        groupCount = (groupCount + 1)
        
    ElseIf (tbsTabs.SelectedItem.index = 2) Then ' Clan Tab
        Dim ClanName As String ' ...
    
        ' ...
        Do
            ' ...
            ClanName = "New Clan #" & (clanCount + 1)
        Loop While (GetAccess(GroupName, "Group").Username <> vbNullString)
        
        ' ...
        ReDim Preserve m_DB(UBound(m_DB) + 1)
        
        ' ...
        With m_DB(UBound(m_DB))
            .Username = ClanName
            .Type = "CLAN"
            .AddedBy = "(console)"
            .AddedOn = Now
            .ModifiedBy = "(console)"
            .ModifiedOn = Now
        End With
        
        ' ...
        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, "Clan: " & _
            ClanName, ClanName, 2)
            
        ' change misc. settings
        With newNode
            .tag = "Clan"
        End With
            
        ' increment clan counter
        clanCount = (clanCount + 1)
        
    ElseIf (tbsTabs.SelectedItem.index = 3) Then ' Game Tab
        ' ...
        Call frmGameSelection.Show(vbModal, frmDBManager)
        
        ' ...
        If (m_game <> vbNullString) Then
            If (GetAccess(m_game, "GAME").Username = vbNullString) Then
                ' ...
                ReDim Preserve m_DB(UBound(m_DB) + 1)
                
                ' ...
                With m_DB(UBound(m_DB))
                    .Username = m_game
                    .Type = "GAME"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With
                
                ' ...
                Set newNode = trvUsers.Nodes.Add("Database", tvwChild, "Game: " & _
                    m_game, m_game, 2)
                    
                ' change misc. settings
                With newNode
                    .tag = "Game"
                End With
            Else
                ' alert user that game entry already exists
                MsgBox "There is already an entry of this type matching " & _
                    "the specified name."
            End If
        End If
    End If
    
    ' ...
    If (Not (newNode Is Nothing)) Then
        ' change misc. settings
        With newNode
            .Selected = True
        End With

        ' ...
        If ((tbsTabs.SelectedItem.index = 1) Or (tbsTabs.SelectedItem.index = 2)) Then
            ' ...
            m_new_entry = True
        
            ' ...
            Call trvUsers.StartLabelEdit
        End If
    End If
End Sub

' ...
Private Sub btnCancel_Click()
    ' ...
    Call Unload(frmDBManager)
End Sub

' ...
Private Sub btnSave_Click(index As Integer)
    Dim i As Integer ' ...
    Dim j As Integer ' ...

    ' are we looking at a single entry or are we saving it all?
    If (index = 1) Then
        ' if we have no selected user... escape quick!
        If (trvUsers.SelectedItem Is Nothing) Then
            ' break from function
            Exit Sub
        End If
    
        ' look for selected user in database
        For i = LBound(m_DB()) To UBound(m_DB())
            ' is this the user we were looking for?
            If (StrComp(trvUsers.SelectedItem.text, m_DB(i).Username, vbTextCompare) = 0) Then
                ' modifiy user data
                With m_DB(i)
                    .Access = Val(txtRank.text)
                    .Flags = txtFlags.text
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With
                
                ' ...
                If (m_group_change) Then
                    ' ...
                    If (m_group_index > -1) Then
                        m_DB(i).Groups = lstGroups.List(m_group_index)
                        
                        ' ...
                        If (tbsTabs.SelectedItem.index = 1) Then
                            Set trvUsers.SelectedItem.Parent = _
                                    trvUsers.Nodes(Exists(m_DB(i).Groups, "Group"))
                        End If
                    Else
                        m_DB(i).Groups = vbNullString
                        
                        ' ...
                        If (tbsTabs.SelectedItem.index = 1) Then
                            Set trvUsers.SelectedItem.Parent = trvUsers.Nodes(1)
                        End If
                    End If
                    
                    ' ...
                    If (lstGroups.SelCount > 1) Then
                        ' ...
                        For j = 0 To (lstGroups.ListCount - 1)
                            ' ...
                            If (j <> m_group_index) Then
                                ' ...
                                If (lstGroups.Selected(j) = True) Then
                                    ' ...
                                    m_DB(i).Groups = m_DB(i).Groups & "," & _
                                        lstGroups.List(j)
                                End If
                            End If
                        Next j
                    End If
                    
                    ' ...
                    m_group_change = False
                End If
                
                ' break loop
                Exit For
            End If
        Next i
        
        ' disable entry save command
        btnSave(1).Enabled = False
    Else
        ' write temporary database to official
        DB() = m_DB()
        
        ' save database
        Call WriteDatabase(GetFilePath("users.txt"))
        
        ' check channel to find potential banned users
        Call checkUsers
        
        ' close database form
        Call Unload(frmDBManager)
    End If
End Sub

' ...
Private Sub lstGroups_Click()
    ' ...
    m_group_change = True '

    ' ...
    If (lstGroups.SelCount = 0) Then
        m_group_index = -1
    ElseIf (lstGroups.SelCount = 1) Then
        m_group_index = lstGroups.ListIndex
    Else
        MsgBox lstGroups.ListIndex
    End If '

    ' ...
    btnSave(1).Enabled = True
End Sub

' ...
Private Sub mnuDelete_Click()
    ' ...
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' ...
        Call HandleDeleteEvent(trvUsers.SelectedItem)
    End If
End Sub

' ...
Private Sub btnDelete_Click()
    ' ...
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' ...
        Call HandleDeleteEvent(trvUsers.SelectedItem)
    End If
End Sub

' ...
Private Sub mnuOpenDatabase_Click()
    ' open file dialog
    Call CommonDialog.ShowOpen
End Sub

' ...
Private Sub mnuRename_Click()
    ' open selected entry for editing
    Call trvUsers.StartLabelEdit
End Sub

' handle tab clicks and initial loading
Private Sub tbsTabs_Click()
    Dim newNode As node ' ...

    Dim i       As Integer ' ...
    Dim grp     As String ' ...
    Dim j       As Integer ' ...
    Dim pos     As Integer ' ...

    ' clear treeview
    Call trvUsers.Nodes.Clear
    
    ' create root node
    Call trvUsers.Nodes.Add(, , "Database", "Database")

    ' which tab index are we on?
    Select Case (tbsTabs.SelectedItem.index)
        Case 1: ' Users and Groups
            For i = LBound(m_DB) To UBound(m_DB)
                ' we're handling groups first; is this entry a group?
                If (StrComp(m_DB(i).Type, "GROUP", vbBinaryCompare) = 0) Then
                    ' is this group a member of other groups?
                    If (Len(m_DB(i).Groups) And (m_DB(i).Groups <> "%")) Then
                        ' is entry member of multiple groups?
                        If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            ' split up multiple groupings
                            grp = Split(m_DB(i).Groups, ",", 2)(0)
                        Else
                            ' no need for special handling...
                            grp = m_DB(i).Groups
                        End If
                    
                        ' has the group already been added or is database in an
                        ' incorrect order?
                        pos = Exists(grp, "Group")
                        
                        ' ... well, does it exist?
                        If (pos) Then
                            ' make node a child of existing group
                            Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(pos).Key, _
                                tvwChild, "Group: " & m_DB(i).Username, m_DB(i).Username, 1)
                        Else
                            ' lets make this guy a parent node for now until we can find
                            ' his real parent.
                            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                                "Group: " & m_DB(i).Username, m_DB(i).Username, 1)
                        End If
                    Else
                        Dim k   As Integer ' ...
                        Dim bln As Boolean ' ...
                    
                        ' Okay, is the group a lone ranger?  Or does he have children
                        ' that are already in the list?
                        For j = LBound(m_DB) To (i - 1)
                            ' we're only concerned with groups, atm.
                            If (StrComp(m_DB(j).Type, "GROUP", vbBinaryCompare) = 0) Then
                                ' we only need to check for groups that are members of
                                ' other groups
                                If (Len(m_DB(j).Groups) And (m_DB(j).Groups <> "%")) Then
                                    ' is entry member of multiple groups?
                                    If (InStr(1, m_DB(j).Groups, ",", vbBinaryCompare) <> 0) Then
                                        ' split up multiple groupings
                                        grp = Split(m_DB(j).Groups, ",", 2)(0)
                                    Else
                                        ' no need for special handling...
                                        grp = m_DB(j).Groups
                                    End If
                                
                                    ' is the current group a member of our group?
                                    If (StrComp(grp, m_DB(i).Username, vbTextCompare) = 0) Then
                                        ' indicate that we've found a match
                                        bln = True
                                        
                                        ' break from loop
                                        Exit For
                                    End If
                                End If
                            End If
                        Next j
                        
                        ' create node
                        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, "Group: " & _
                            m_DB(i).Username, m_DB(i).Username, 1)
                    
                        ' is this node a baby's daddy?
                        If (bln) Then
                            ' move node
                            Set trvUsers.Nodes(Exists(m_DB(j).Username, "Group")).Parent = _
                                newNode
                        End If

                        ' reset boolean
                        bln = False
                    End If
                    
                    ' change misc. settings
                    With newNode
                        .tag = "Group"
                    End With
                End If
            Next i

            ' loop through database... this time looking for users.
            For i = LBound(m_DB) To UBound(m_DB)
                ' is the entry a user?
                If (StrComp(m_DB(i).Type, "USER", vbBinaryCompare) = 0) Then
                    ' is the user a member of any groups?
                    If (Len(m_DB(i).Groups) And (m_DB(i).Groups <> "%")) Then
                        ' is entry member of multiple groups?
                        If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            ' split up multiple groupings
                            grp = Split(m_DB(i).Groups, ",", 2)(0)
                        Else
                            ' no need for special handling...
                            grp = m_DB(i).Groups
                        End If
                    
                        ' search for our group
                        pos = Exists(grp, "Group")
                            
                        ' does our group exist?
                        If (pos) Then
                            ' create user node and move into group
                            Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(pos).Key, _
                                tvwChild, "User: " & m_DB(i).Username, m_DB(i).Username, 3)
                        End If
                    Else
                        ' create new user node under root
                        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                            "User: " & m_DB(i).Username, m_DB(i).Username, 3)
                    End If
                    
                    ' change misc. settings
                    With newNode
                        .tag = "User"
                    End With
                End If
            Next i
            
            ' enable create user button
            btnCreateUser.Enabled = True
            
        Case 2: ' Clans
            ' loop through database searching for clans
            For i = LBound(m_DB) To UBound(m_DB)
                ' is entry a clan?
                If (StrComp(m_DB(i).Type, "CLAN", vbBinaryCompare) = 0) Then
                    ' create new node
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "Clan: " & m_DB(i).Username, m_DB(i).Username, 2)
                        
                    ' change misc. settings
                    With newNode
                        .tag = "Clan"
                    End With
                End If
            Next i

            ' disable create user button
            btnCreateUser.Enabled = False
            
        Case 3: ' Games
            ' loop through database searching for games
            For i = LBound(m_DB) To UBound(m_DB)
                ' is entry a game?
                If (StrComp(m_DB(i).Type, "GAME", vbBinaryCompare) = 0) Then
                    ' create new node
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "Group: " & m_DB(i).Username, m_DB(i).Username, 2)
                        
                    ' change misc. settings
                    With newNode
                        .tag = "Game"
                    End With
                End If
            Next i
            
            ' disable create user button
            btnCreateUser.Enabled = False
    End Select
    
    ' does our treeview contain any nodes?  hope so...
    If (trvUsers.Nodes.Count) Then
        ' change misc. settings for root node
        With trvUsers.Nodes(1)
            .Expanded = True
            .Image = 1
        End With
        
        ' refresh treeview
        Call trvUsers.Refresh
    End If
    
    ' ...
    Call UpdateGroupListBox
    
    ' ...
    Call LockGUI
End Sub

Private Sub LockGUI()
    Dim i As Integer ' ...
    
    ' set our default frame caption
    With frmDatabase
        .Caption = "Database"
    End With

    ' disable & clear rank
    txtRank.Enabled = False
    txtRank.text = vbNullString
    
    ' disable & clear flags
    txtFlags.Enabled = False
    txtFlags.text = vbNullString
    
    ' loop through listbox and clear selected items
    For i = 0 To (lstGroups.ListCount - 1)
        lstGroups.Selected(i) = False
    Next i
    
    ' disable group list
    lstGroups.Enabled = False
    
    ' reset created on & modified on labels
    lblCreatedOn.Caption = "(not applicable)"
    lblModifiedOn.Caption = "(not applicable)"
    
    ' reset created by & modified by labels
    lblCreatedBy.Caption = vbNullString
    lblModifiedBy.Caption = vbNullString
    
    ' disable entry buttons
    btnSave(1).Enabled = False
    btnDelete.Enabled = False
End Sub

Private Sub UnlockGUI()
    Dim i As Integer ' ...

    ' set our default frame caption
    With frmDatabase
        .Caption = "Database"
    End With

    ' enable & clear rank
    txtRank.Enabled = True
    txtRank.text = vbNullString
    
    ' enable & clear flags
    txtFlags.Enabled = True
    txtFlags.text = vbNullString

    ' loop through listbox and clear selected items
    For i = 0 To (lstGroups.ListCount - 1)
        lstGroups.Selected(i) = False
    Next i
    
    ' enable group list
    lstGroups.Enabled = True
    
    ' reset created on & modified on labels
    lblCreatedOn.Caption = "(not applicable)"
    lblModifiedOn.Caption = "(not applicable)"
    
    ' reset created by & modified by labels
    lblCreatedBy.Caption = vbNullString
    lblModifiedBy.Caption = vbNullString
    
    ' disable entry buttons
    btnSave(1).Enabled = False
    btnDelete.Enabled = False
End Sub

' handle node collapse
Private Sub trvUsers_Collapse(ByVal node As node)
    ' refresh tree view
    Call trvUsers.Refresh
End Sub

' handle node expand
Private Sub trvUsers_Expand(ByVal node As node)
    ' refresh tree view
    Call trvUsers.Refresh
End Sub

' ...
Private Sub trvUsers_NodeClick(ByVal node As MSComctlLib.node)
    Dim tmp As udtGetAccessResponse ' ...
    Dim i   As Integer ' ...
    
    ' ...
    If (node Is Nothing) Then
        Exit Sub
    End If
    
    ' ...
    m_group_index = -1

    ' ...
    If (node.index > 1) Then
        ' ...
        Call UnlockGUI
    
        ' ...
        frmDatabase.Caption = node.text
        
        ' ...
        tmp = GetAccess(node.text, node.tag)
        
        ' does entry have a rank?
        If (tmp.Access > 0) Then
            ' write rank to text box
            txtRank.text = tmp.Access
        Else
            ' clear rank from text box
            txtRank.text = vbNullString
        End If
        
        ' clear flags from text box
        txtFlags.text = tmp.Flags
        
        ' ...
        If ((tmp.AddedBy = vbNullString) Or (tmp.AddedBy = "%")) Then
            lblCreatedOn = "unknown"
            lblCreatedBy = "by unknown"
        Else
            lblCreatedOn = tmp.AddedOn & " Local Time"
            lblCreatedBy = "by " & tmp.AddedBy
        End If
        
        ' ...
        If ((tmp.ModifiedBy = vbNullString) Or (tmp.ModifiedBy = "%")) Then
            lblModifiedOn = "unknown"
            lblModifiedBy = "by unknown"
        Else
            lblModifiedOn = tmp.ModifiedOn & " Local Time"
            lblModifiedBy = "by " & tmp.ModifiedBy
        End If
        
        ' is entry a member of a group?
        If (Len(tmp.Groups) And (tmp.Groups <> "%")) Then
            Dim splt() As String  ' ...
            Dim j      As Integer ' ...
        
            ' is entry a member of multiple groups?
            If (InStr(1, tmp.Groups, ",", vbBinaryCompare) <> 0) Then
                ' store working copy of group memberships, splitting up
                ' multiple groupings by the ',' delimiter.
                splt() = Split(tmp.Groups, ",")
            Else
                ' redefine array size to store group name
                ReDim Preserve splt(0)
                
                ' store working copy of group membership
                splt(0) = tmp.Groups
            End If
            
            ' loop through entry's group memberships
            For i = LBound(splt) To UBound(splt)
                ' loop through our group listing, checking to see if we have any
                ' matches (since the entry is a member of a group, we better!)
                For j = 0 To (lstGroups.ListCount - 1)
                    ' is entry a member of group?
                    If (StrComp(splt(i), lstGroups.List(j), vbTextCompare) = 0) Then
                        ' ...
                        If (m_group_index = -1) Then
                            m_group_index = j
                        End If
                    
                        ' select group if entry is a member
                        lstGroups.Selected(j) = True
                    End If
                Next j
            Next i
        End If
    Else
        ' ...
        Call LockGUI
    End If
    
    ' ...
    Set trvUsers.SelectedItem = node

    ' refresh tree view
    Call trvUsers.Refresh
End Sub

' ...
Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' ...
    If (Button = vbLeftButton) Then
        ' ...
        Set trvUsers.SelectedItem = trvUsers.HitTest(X, Y)
        
        ' ...
        Call trvUsers_NodeClick(trvUsers.SelectedItem)
    End If
End Sub

' ...
Private Sub trvUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' ...
    If (Button = vbRightButton) Then
        Dim gAcc As udtGetAccessResponse ' ...

        ' ...
        If (Not (trvUsers.SelectedItem Is Nothing)) Then
            ' ...
            If (trvUsers.SelectedItem.index > 1) Then
                ' ...
                If (StrComp(trvUsers.SelectedItem.tag, "Group", vbTextCompare) = 0) Then
                    ' ...
                    mnuRename.Enabled = True
                Else
                    ' ...
                    mnuRename.Enabled = False
                End If
                
                ' ...
                mnuDelete.Enabled = True
            Else
                ' ...
                mnuRename.Enabled = False
                
                ' ...
                mnuDelete.Enabled = False
            End If
            
            ' ...
            Call Me.PopupMenu(mnuContext)
        End If
    End If
End Sub

' ...
Private Sub trvUsers_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, _
    Shift As Integer, X As Single, Y As Single, State As Integer)
    
    ' ...
    Set trvUsers.DropHighlight = trvUsers.HitTest(X, Y)
End Sub

' ...
Private Sub trvUsers_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    Dim nodePrev As node ' ...
    Dim nodeNow  As node ' ...

    Dim strKey   As String  ' ...
    Dim res      As Integer ' ...
    Dim i        As Integer ' ...
    Dim found    As Integer ' ...
    
    ' ...
    Set nodeNow = trvUsers.DropHighlight
    
    ' ...
    Set nodePrev = trvUsers.SelectedItem
        
    ' ...
    If (nodeNow.index = 1) Then
        ' ...
        For i = LBound(m_DB) To UBound(m_DB)
            ' ...
            If (StrComp(m_DB(i).Username, nodePrev.text, vbTextCompare) = 0) Then
                ' ...
                If (StrComp(m_DB(i).Type, nodePrev.tag, vbTextCompare) = 0) Then
                    ' ...
                    If ((Len(m_DB(i).Groups) > 0) And (m_DB(i).Groups <> "%")) Then
                        ' ...
                        Set nodePrev.Parent = nodeNow
                    End If
                    
                    ' ...
                    m_DB(i).Groups = vbNullString
                    
                    ' ...
                    Exit For
                End If
            End If
        Next i
    Else
        ' ...
        If (nodePrev.index <> 1) Then
            ' ...
            If (StrComp(nodeNow.tag, "Group", vbTextCompare) <> 0) Then
                ' ...
                Set nodeNow = nodeNow.Parent
                
                ' ...
                If (nodeNow.index = 1) Then
                    ' ...
                    Set trvUsers.DropHighlight = nodeNow
                
                    ' ...
                    Call trvUsers_OLEDragDrop(Data, Effect, Button, Shift, _
                        X, Y)
                
                    ' ...
                    Exit Sub
                End If
            End If
        
            ' ...
            If (IsInGroup(nodePrev.text, nodeNow.text) = False) Then
                ' ...
                For i = LBound(m_DB) To UBound(m_DB)
                    ' ...
                    If (StrComp(m_DB(i).Username, nodePrev.text, vbTextCompare) = 0) Then
                        ' ...
                        If (StrComp(m_DB(i).Type, nodePrev.tag, vbTextCompare) = 0) Then
                            ' ...
                            m_DB(i).Groups = nodeNow.text
                            
                            ' ...
                            Exit For
                        End If
                    End If
                Next i
                
                ' ...
                Set nodePrev.Parent = nodeNow
            End If
        End If
    End If
    
    ' ...
    Call trvUsers_NodeClick(nodePrev)
    
    ' ...
    Set trvUsers.DropHighlight = Nothing

ERROR_HANDLER:
    ' potential cycle introduction error
    If (Err.Number = 35614) Then
        MsgBox Err.description, vbCritical, "Error"
    End If
    
    ' ...
    Set trvUsers.DropHighlight = Nothing
    
    ' ...
    Exit Sub
End Sub

' ...
Private Sub trvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ...
    If (KeyCode = vbKeyDelete) Then
        ' ...
        If (Not (trvUsers.SelectedItem Is Nothing)) Then
            ' ...
            Call HandleDeleteEvent(trvUsers.SelectedItem)
        End If
    End If
End Sub

' ...
Private Sub trvUsers_BeforeLabelEdit(Cancel As Integer)
    ' is this a new entry that requires a name?
    If (m_new_entry) Then
        ' break from function
        Exit Sub
    End If

    ' is entry a group? if not, disallow edit.
    If (StrComp(trvUsers.SelectedItem.tag, "Group", vbTextCompare) <> 0) Then
        ' disallow editing of entry label
        Cancel = 1
    End If
End Sub

' ...
Private Sub trvUsers_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim i As Integer ' ...
    
    ' ...
    If (trvUsers.SelectedItem Is Nothing) Then
        ' ...
        Exit Sub
    End If
    
    ' ...
    If (m_new_entry) Then
        ' ...
        If (GetAccess(NewString, trvUsers.SelectedItem.tag).Username = vbNullString) Then
            ' ...
            With m_DB(UBound(m_DB))
                .Username = NewString
            End With
            
            ' ...
            m_new_entry = False
        Else
            ' ...
            MsgBox "There is already an entry of this type matching the specified name."
            
            ' ...
            Call trvUsers.StartLabelEdit
            
            ' ...
            Cancel = 1
        End If
    Else
        ' ...
        For i = LBound(m_DB) To UBound(m_DB)
            ' ...
            If (StrComp(trvUsers.SelectedItem.text, m_DB(i).Username, vbTextCompare) = 0) Then
                ' ...
                If (StrComp(trvUsers.SelectedItem.tag, m_DB(i).Type, vbTextCompare) = 0) Then
                    ' ...
                    m_DB(i).Username = NewString
            
                    ' ...
                    Exit For
                End If
            End If
        Next i
        
        ' ...
        If (StrComp(m_DB(i).Type, "Group", vbTextCompare) = 0) Then
            ' ...
            For i = LBound(m_DB) To UBound(m_DB)
                ' ...
                If ((Len(m_DB(i).Groups) > 0) And (m_DB(i).Groups <> "%")) Then
                    Dim splt() As String  ' ...
                    Dim j      As Integer ' ...
                
                    ' ...
                    If (InStr(1, m_DB(i).Groups, ",", vbTextCompare) <> 0) Then
                        ' ...
                        splt() = Split(m_DB(i).Groups, ",")
                    Else
                        ' ...
                        ReDim Preserve splt(0)
                        
                        ' ...
                        splt(0) = m_DB(i).Groups
                    End If
                    
                    ' ...
                    For j = LBound(splt) To UBound(splt)
                        ' ...
                        If (StrComp(splt(j), trvUsers.SelectedItem.text, vbTextCompare) = 0) Then
                            ' ...
                            splt(j) = NewString
                        End If
                    Next j
                    
                    ' ...
                    m_DB(i).Groups = Join(splt(), ",")
                End If
            Next i
        End If
    End If
End Sub

Private Function IsInGroup(ByVal Username As String, ByVal GroupName As String) As Boolean
    Dim i      As Integer ' ...
    Dim j      As Integer ' ...
    Dim splt() As String  ' ...
    
    ' ...
    For i = LBound(m_DB) To UBound(m_DB)
        ' ...
        If (StrComp(m_DB(i).Username, Username, vbTextCompare) = 0) Then
            ' ...
            If ((Len(m_DB(i).Groups) > 0) And (m_DB(i).Groups <> "%")) Then
                ' ...
                If (InStr(1, m_DB(i).Groups, "%", vbBinaryCompare) <> 0) Then
                    ' ...
                    splt() = Split(m_DB(i).Groups, "%")
                Else
                    ' ...
                    ReDim splt(0)
                    
                    ' ...
                    splt(0) = m_DB(i).Groups
                End If
                
                ' ...
                For j = LBound(splt) To UBound(splt)
                    If (StrComp(GroupName, splt(j), vbTextCompare) = 0) Then
                        ' ...
                        IsInGroup = True
                        
                        ' ...
                        Exit Function
                    End If
                Next j
            End If
        End If
    Next i
End Function

Private Sub HandleDeleteEvent(ByRef NodeToDelete As node)
    Dim Temp As node ' ...
    
    ' ...
    Set Temp = NodeToDelete

    ' ...
    If (Temp Is Nothing) Then
        ' ...
        Exit Sub
    End If

    ' ...
    If (Temp.index > 1) Then
        Dim response As Integer ' ...
        Dim isGroup  As Boolean ' ...
        
        ' ...
        isGroup = (StrComp(Temp.tag, "Group", vbTextCompare) = 0)
    
        ' ...
        If (isGroup) Then
            ' ...
            response = MsgBox("Are you sure you wish to delete this group and " & _
                "all of its members?", vbYesNo + vbInformation, "Information")
        End If
        
        ' ...
        If ((isGroup = False) Or ((isGroup) And (response = vbYes))) Then
            ' ...
            Call DB_remove(Temp.text, Temp.tag)
            
            ' ...
            If (Temp.Next Is Nothing) Then
                ' ...
                If (Temp.Previous Is Nothing) Then
                    ' ...
                    trvUsers.Nodes(Temp.Parent.index).Selected = True
                Else
                    ' ...
                    trvUsers.Nodes(Temp.Previous.index).Selected = True
                End If
            Else
                ' ...
                trvUsers.Nodes(Temp.Next.index).Selected = True
            End If
            
            ' ...
            Call trvUsers.Nodes.Remove(Temp.index)
            
            ' ...
            Call trvUsers_NodeClick(trvUsers.SelectedItem)
            
            ' ...
            Call UpdateGroupListBox
        End If
    End If
End Sub

Private Sub UpdateGroupListBox()
    Dim i As Integer ' ...

    ' clear group selection listing
    Call lstGroups.Clear

    ' go through group listing
    For i = LBound(m_DB) To UBound(m_DB)
        ' ...
        If (StrComp(m_DB(i).Type, "Group", vbTextCompare) = 0) Then
            ' add group to group selection listbox
            Call lstGroups.AddItem(m_DB(i).Username)
        End If
    Next i
End Sub

' ...
Private Function Exists(ByVal nodeName As String, Optional tag As String = vbNullString) As Integer
    Dim i As Integer ' ...
    
    ' ...
    For i = 1 To trvUsers.Nodes.Count
        ' ...
        If (StrComp(trvUsers.Nodes(i).text, nodeName, vbTextCompare) = 0) Then
            ' ...
            If (tag <> vbNullString) Then
                ' ...
                If (StrComp(trvUsers.Nodes(i).tag, tag, vbTextCompare) = 0) Then
                    ' ...
                    Exists = i
                    
                    ' ...
                    Exit Function
                End If
            Else
                ' ...
                Exists = i
            
                ' ...
                Exit Function
            End If
        End If
    Next i
    
    ' ...
    Exists = False
End Function

' ...
Private Sub txtFlags_Change()
    ' enable entry save button
    btnSave(1).Enabled = True
End Sub

' ...
Private Sub txtRank_Change()
    ' enable entry save button
    btnSave(1).Enabled = True
End Sub

' ...
Private Function GetAccess(ByVal Username As String, Optional dbType As String = _
    vbNullString) As udtGetAccessResponse
    
    Dim i   As Integer ' ...
    Dim bln As Boolean ' ...
    
    dbType = UCase$(dbType)

    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Username, Username, vbTextCompare) = 0) Then
            If (Len(dbType) > 0) Then
                If (StrComp(m_DB(i).Type, dbType, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
                
            If (bln = True) Then
                With GetAccess
                    .Username = m_DB(i).Username
                    .Access = m_DB(i).Access
                    .Flags = m_DB(i).Flags
                    .AddedBy = m_DB(i).AddedBy
                    .AddedOn = m_DB(i).AddedOn
                    .ModifiedBy = m_DB(i).ModifiedBy
                    .ModifiedOn = m_DB(i).ModifiedOn
                    .Type = m_DB(i).Type
                    .Groups = m_DB(i).Groups
                    .BanMessage = m_DB(i).BanMessage
                End With
                
                Exit Function
            End If
        End If
        
        bln = False
    Next i

    GetAccess.Access = -1
End Function

' ...
Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = _
    vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER

    Dim i     As Integer ' ...
    Dim found As Boolean ' ...
    
    dbType = UCase$(dbType)
    
    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean ' ...
        
            If (Len(dbType)) Then
                If (StrComp(m_DB(i).Type, dbType, vbBinaryCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
            
            If (bln) Then
                found = True
                
                Exit For
            End If
        End If
        
        bln = False
    Next i
    
    If (found) Then
        Dim bak As udtDatabase ' ...
        
        Dim j   As Integer ' ...
        
        ' ...
        bak = m_DB(i)

        ' we aren't removing the last array
        ' element, are we?
        If (i < UBound(m_DB)) Then
            For j = (i + 1) To UBound(m_DB)
                m_DB(j - 1) = m_DB(j)
            Next j
        End If
        
        If (UBound(m_DB)) Then
            ' redefine array size
            ReDim Preserve m_DB(UBound(m_DB) - 1)
        End If
        
        ' if we're removing a group, we need to also fix our
        ' group memberships, in case anything is broken now
        If (StrComp(bak.Type, "GROUP", vbBinaryCompare) = 0) Then
            Dim res As Boolean ' ...
       
            ' if we remove a user from the database during the
            ' execution of the inner loop, we have to reset our
            ' inner loop variables, otherwise we create errors
            ' due to incorrect database indexes.  Because of this,
            ' we have to dual-loop until our inner loop runs out
            ' of matching users.
            Do
                ' reset loop variable
                res = False
            
                ' loop through database checking for users that
                ' were members of the group that we just removed
                For i = LBound(m_DB) To UBound(m_DB)
                    If (Len(m_DB(i).Groups) And m_DB(i).Groups <> "%") Then
                        If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Dim splt()     As String ' ...
                            Dim innerfound As Boolean ' ...
                            
                            splt() = Split(m_DB(i).Groups, ",")
                            
                            For j = LBound(splt) To UBound(splt)
                                If (StrComp(bak.Username, splt(j), vbTextCompare) = 0) Then
                                    innerfound = True
                                
                                    Exit For
                                End If
                            Next j
                        
                            If (innerfound) Then
                                Dim k As Integer ' ...
                                
                                For k = (j + 1) To UBound(splt)
                                    splt(k - 1) = splt(k)
                                Next k
                                
                                ReDim Preserve splt(UBound(splt) - 1)
                                
                                m_DB(i).Groups = Join(splt(), vbNullString)
                            End If
                        Else
                            If (StrComp(bak.Username, m_DB(i).Groups, vbTextCompare) = 0) Then
                                res = DB_remove(m_DB(i).Username, m_DB(i).Type)
                                
                                Exit For
                            End If
                        End If
                    End If
                Next i
            Loop While (res)
        End If
        
        ' commit modifications
        Call WriteDatabase(GetFilePath("users.txt"))
        
        DB_remove = True
        
        Exit Function
    End If
    
    DB_remove = False
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Error: DB_remove() has encountered an error while " & _
        "removing a database entry.")
        
    DB_remove = False
    
    Exit Function
End Function
