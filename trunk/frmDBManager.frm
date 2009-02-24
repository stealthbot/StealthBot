VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDBManager 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   6255
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
   ScaleHeight     =   6255
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2880
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin VB.Frame frmDatabase 
      Caption         =   "Eric[nK]"
      Height          =   5295
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   3025
      Begin VB.ComboBox cbxGroups 
         Height          =   315
         ItemData        =   "frmDBManager.frx":0000
         Left            =   240
         List            =   "frmDBManager.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2543
         Width           =   2570
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   4845
         Width           =   855
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Top             =   4845
         Width           =   855
      End
      Begin VB.TextBox txtBanMessage 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   4440
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   1200
         Left            =   240
         TabIndex        =   19
         Top             =   2920
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   2117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
         EndProperty
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
      Begin VB.Label Label3 
         Caption         =   "Ban message:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   2535
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
         TabIndex        =   16
         Top             =   1845
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Group(s):"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   2535
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
         TabIndex        =   18
         Top             =   2005
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
         TabIndex        =   17
         Top             =   1375
         Width           =   2415
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
         Height          =   135
         Left            =   360
         TabIndex        =   13
         Top             =   1215
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
         TabIndex        =   15
         Top             =   1005
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
         TabIndex        =   14
         Top             =   1630
         Width           =   2535
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
            Picture         =   "frmDBManager.frx":0013
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0565
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0AB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnCreateGroup 
      Caption         =   "Create Group"
      Height          =   375
      Left            =   1800
      Picture         =   "frmDBManager.frx":1009
      TabIndex        =   2
      ToolTipText     =   "Create Group"
      Top             =   5378
      Width           =   1695
   End
   Begin VB.CommandButton btnCreateUser 
      Caption         =   "Create User"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "frmDBManager.frx":1471
      TabIndex        =   1
      ToolTipText     =   "Create User"
      Top             =   5378
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
            Object.ToolTipText     =   "Game entries allow access to be given based on game type"
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
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   5880
      Width           =   735
   End
   Begin MSComctlLib.TreeView trvUsers 
      Height          =   4725
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8334
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

Private Sub cbxGroups_Click()
    Dim I As Integer ' ...
    
    ' ...
    For I = 1 To lvGroups.ListItems.Count
        ' ...
        If (StrComp(cbxGroups.text, lvGroups.ListItems(I), vbTextCompare) = 0) Then
            ' ...
            lvGroups.ListItems(I).Checked = False
            
            ' ...
            Exit For
        End If
    Next I
    
    ' enable entry save command
    btnSave(1).Enabled = True
    
    ' ...
    m_group_change = True
End Sub

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
Public Sub ImportDatabase(strPath As String, dbType As Integer)
    
    Dim f    As Integer ' ...
    Dim buf  As String  ' ...
    Dim n    As node    ' ...
    Dim I    As Integer ' ...
    
    ' ...
    f = FreeFile

    ' ...
    If (dbType = 0) Then
    
        Open strPath For Input As #f
            ' ...
            Do While (EOF(f) = False)
                Line Input #f, buf
                
                ' ...
                If (buf <> vbNullString) Then
                    ' ...
                    If (GetAccess(buf, "USER").Username <> vbNullString) Then
                        ' ...
                        For I = 0 To UBound(m_DB)
                            ' ...
                            If ((StrComp(m_DB(I).Username, buf, vbTextCompare) = 0) And _
                                (StrComp(m_DB(I).Type, "USER", vbTextCompare) = 0)) Then
                            
                                ' ...
                                With m_DB(I)
                                    .Username = buf
                                    .Type = "USER"
                                    .ModifiedBy = "(console)"
                                    .ModifiedOn = Now
                                    
                                    ' ...
                                    If (InStr(1, .Flags, "S", vbBinaryCompare) = 0) Then
                                        .Flags = .Flags & "S"
                                    End If
                                    
                                    ' ...
                                    If (Not (trvUsers.DropHighlight Is Nothing)) Then
                                        ' ...
                                        If (StrComp(trvUsers.DropHighlight.Tag, "Group", vbTextCompare) = 0) Then
                                            .Groups = .Groups & "," & trvUsers.DropHighlight.text
                                        End If
                                    End If
                                    
                                    ' ...
                                    If (.Groups = vbNullString) Then
                                        .Groups = "%"
                                    End If
                                End With
                                
                                ' ...
                                Exit For
                            End If
                        Next I
                    Else
                        ' redefine array to support new entry
                        ReDim Preserve m_DB(UBound(m_DB) + 1)
                        
                        ' create new database entry
                        With m_DB(UBound(m_DB))
                            .Username = buf
                            .Type = "USER"
                            .AddedBy = "(console)"
                            .AddedOn = Now
                            .ModifiedBy = "(console)"
                            .ModifiedOn = Now
                            .Flags = "S"
                            
                            ' ...
                            If (Not (trvUsers.DropHighlight Is Nothing)) Then
                                ' ...
                                If (StrComp(trvUsers.DropHighlight.Tag, "Group", vbTextCompare) = 0) Then
                                    ' ...
                                    If (IsInGroup(buf, trvUsers.DropHighlight.text) = False) Then
                                        .Groups = .Groups & "," & trvUsers.DropHighlight.text
                                    End If
                                End If
                            End If
                            
                            ' ...
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
                    End If
                End If
            Loop ' end loop
        Close #f
        
    ElseIf ((dbType = 1) Or (dbType = 2)) Then
        
        Dim user As String ' ...
        Dim Msg  As String ' ...
    
        Open strPath For Input As #f
            ' ...
            Do While (EOF(f) = False)
                Line Input #f, buf
                
                ' ...
                If (buf <> vbNullString) Then
                    ' ...
                    If (InStr(1, buf, Space$(1), vbBinaryCompare) <> 0) Then
                        ' ...
                        user = Left$(buf, InStr(1, buf, Space$(1), vbBinaryCompare) - 1)
                        
                        ' ...
                        Msg = Mid$(buf, Len(user) + 1)
                    Else
                        user = buf
                    End If
                
                    ' ...
                    If (GetAccess(user, "USER").Username <> vbNullString) Then
                        ' ...
                        For I = 0 To UBound(m_DB)
                            ' ...
                            If ((StrComp(m_DB(I).Username, user, vbTextCompare) = 0) And _
                                (StrComp(m_DB(I).Type, "User", vbTextCompare) = 0)) Then
                                
                                ' ...
                                With m_DB(I)
                                    .Username = user
                                    .Type = "USER"
                                    .ModifiedBy = "(console)"
                                    .ModifiedOn = Now
                                    
                                    ' ...
                                    If (InStr(1, .Flags, "B", vbBinaryCompare) = 0) Then
                                        .Flags = .Flags & "B"
                                    End If
                                    
                                    ' ...
                                    If (Not (trvUsers.DropHighlight Is Nothing)) Then
                                        ' ...
                                        If (StrComp(trvUsers.DropHighlight.Tag, "Group", vbTextCompare) = 0) Then
                                            ' ...
                                            If (IsInGroup(user, trvUsers.DropHighlight.text) = False) Then
                                                .Groups = .Groups & "," & trvUsers.DropHighlight.text
                                            End If
                                        End If
                                    End If
                                    
                                    ' ...
                                    If (.Groups = vbNullString) Then
                                        .Groups = "%"
                                    End If
                                End With
                            End If
                        Next I
                    Else
                        ' redefine array to support new entry
                        ReDim Preserve m_DB(UBound(m_DB) + 1)
                        
                        ' create new database entry
                        With m_DB(UBound(m_DB))
                            .Username = user
                            .Type = "USER"
                            .AddedBy = "(console)"
                            .AddedOn = Now
                            .ModifiedBy = "(console)"
                            .ModifiedOn = Now
                            .Flags = "B"
                            .BanMessage = Msg
                            
                            ' ...
                            If (Not (trvUsers.DropHighlight Is Nothing)) Then
                                ' ...
                                If (StrComp(trvUsers.DropHighlight.Tag, "Group", vbTextCompare) = 0) Then
                                    .Groups = trvUsers.DropHighlight.text
                                End If
                            End If
                            
                            ' ...
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
                    End If
                End If
            Loop ' end loop
        Close #f
        
    End If
    
    ' ...
    tbsTabs_Click

End Sub

' ...
Private Sub btnCreateUser_Click()
    Static userCount As Integer ' ...
    
    Dim newNode      As node    ' ...
    Dim gAcc         As udtGetAccessResponse
    Dim Username     As String  ' ...
    
    ' ...
    Do
        ' ...
        Username = "New_User_#" & (userCount + 1)
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
        If (trvUsers.SelectedItem.Index = 1) Then
            Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                tvwChild, "User: " & Username, Username, 3)
                
        Else
            ' is the item a group?
            If (StrComp(trvUsers.SelectedItem.Tag, "Group", vbTextCompare) = 0) Then
                ' create new node under group node
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, tvwChild, _
                    "User: " & Username, Username, 3)
    
                ' ...
                With m_DB(UBound(m_DB))
                    .Groups = trvUsers.SelectedItem.text
                End With
            Else
                ' is our parent a group?
                If (StrComp(trvUsers.SelectedItem.Parent.Tag, "Group", vbTextCompare) = 0) Then
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
    
    ' ...
    If (Not (newNode Is Nothing)) Then
        ' change misc. settings
        With newNode
            .Tag = "User"
            .Selected = True
        End With
        
        ' ...
        m_new_entry = True
        
        ' open entry name for editing
        Call trvUsers.StartLabelEdit
        
        ' increment user count
        userCount = (userCount + 1)
    End If
End Sub

' ...
Private Sub btnCreateGroup_Click()
    Static groupCount As Integer ' ...
    Static clanCount  As Integer ' ...
    
    Dim newNode       As node    ' ...
    
    ' ...
    If (tbsTabs.SelectedItem.Index = 1) Then ' Users and Groups Tab
        Dim GroupName As String ' ...
    
        ' ...
        Do
            ' ...
            GroupName = "New_Group_#" & (groupCount + 1)
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
            If (trvUsers.SelectedItem.Index = 1) Then
                ' ...
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                    tvwChild, "Group: " & GroupName, GroupName, 1)
            Else
                ' ...
                If (StrComp(trvUsers.SelectedItem.Tag, "Group", vbTextCompare) = 0) Then
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
                    If (StrComp(trvUsers.SelectedItem.Parent.Tag, "Group", vbTextCompare) = 0) Then
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
        
        ' ...
        If (Not (newNode Is Nothing)) Then
            ' change misc. settings
            With newNode
                .Tag = "Group"
            End With
            
            ' increment group counter
            groupCount = (groupCount + 1)
        End If
        
    ElseIf (tbsTabs.SelectedItem.Index = 2) Then ' Clan Tab
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
        
        ' ...
        If (Not (newNode Is Nothing)) Then
            ' change misc. settings
            With newNode
                .Tag = "Clan"
            End With
            
            ' increment clan counter
            clanCount = (clanCount + 1)
        End If
        
    ElseIf (tbsTabs.SelectedItem.Index = 3) Then ' Game Tab
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
                
                ' ...
                If (Not (newNode Is Nothing)) Then
                    ' change misc. settings
                    With newNode
                        .Tag = "Game"
                    End With
                End If
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
        If ((tbsTabs.SelectedItem.Index = 1) Or (tbsTabs.SelectedItem.Index = 2)) Then
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
Private Sub btnSave_Click(Index As Integer)
    Dim I As Integer ' ...
    Dim j As Integer ' ...

    ' are we looking at a single entry or are we saving it all?
    If (Index = 1) Then
        ' if we have no selected user... escape quick!
        If (trvUsers.SelectedItem Is Nothing) Then
            ' break from function
            Exit Sub
        End If
    
        ' look for selected user in database
        For I = LBound(m_DB) To UBound(m_DB)
            ' is this the user we were looking for?
            If (StrComp(trvUsers.SelectedItem.text, m_DB(I).Username, vbTextCompare) = 0) Then
                ' ...
                If (StrComp(trvUsers.SelectedItem.Tag, m_DB(I).Type, vbTextCompare) = 0) Then
                    ' modifiy user data
                    With m_DB(I)
                        .Access = Val(txtRank.text)
                        .Flags = txtFlags.text
                        .ModifiedBy = "(console)"
                        .ModifiedOn = Now
                        .BanMessage = txtBanMessage.text
                        
                        ' ...
                        If (cbxGroups.ListIndex > 0) Then
                            .Groups = cbxGroups.text
                        Else
                            .Groups = vbNullString
                        End If
                        
                        ' ...
                        .Groups = .Groups & ","
                        
                        ' ...
                        For j = 1 To lvGroups.ListItems.Count
                            ' ...
                            If (lvGroups.ListItems(j).Checked) Then
                                .Groups = .Groups & lvGroups.ListItems(j).text & ","
                            End If
                        Next j
                        
                        ' ...
                        If (Len(.Groups) > 1) Then
                            .Groups = Left$(.Groups, Len(.Groups) - 1)
                        End If
                    End With
                    
                    ' ...
                    'If (m_group_change) Then
                    '    ' ...
                    '    If (m_group_index > -1) Then
                    '        m_DB(i).Groups = lstGroups.List(m_group_index)
                    '
                    '        ' ...
                    '        If (tbsTabs.SelectedItem.index = 1) Then
                    '            Set trvUsers.SelectedItem.Parent = _
                    '                    trvUsers.Nodes(Exists(m_DB(i).Groups, "Group"))
                    '        End If
                    '    Else
                    '        m_DB(i).Groups = vbNullString
                    '
                    '        ' ...
                    '        If (tbsTabs.SelectedItem.index = 1) Then
                    '            Set trvUsers.SelectedItem.Parent = trvUsers.Nodes(1)
                    '        End If
                    '    End If
                    '
                    '    ' ...
                    '    If (lstGroups.SelCount > 1) Then
                    '        ' ...
                    '        For j = 0 To (lstGroups.ListCount - 1)
                    '            ' ...
                    '            If (j <> m_group_index) Then
                    '                ' ...
                    '                If (lstGroups.Selected(j) = True) Then
                    '                    ' ...
                    '                    m_DB(i).Groups = m_DB(i).Groups & "," & _
                    '                        lstGroups.List(j)
                    '                End If
                    '            End If
                    '        Next j
                    '    End If
                    '
                    '    ' ...
                    '    m_group_change = False
                    'End If
                    
                    ' break loop
                    Exit For
                End If
            End If
        Next I
        
        ' disable entry save command
        btnSave(1).Enabled = False
        
        ' ...
        If (m_group_change) Then
            ' ...
            tbsTabs_Click
        
            ' ...
            m_group_change = False
        End If
    Else
        ' write temporary database to official
        DB() = m_DB()
        
        ' save database
        Call WriteDatabase(GetFilePath("users.txt"))
        
        ' check channel to find potential banned users
        Call g_Channel.CheckUsers
        
        ' close database form
        Call Unload(frmDBManager)
    End If
End Sub

Private Sub lvGroups_Click()
    Dim I As Integer ' ...
    
    ' ...
    'cbxGroups.ListIndex = 0
        
    ' ...
    ' ...
    For I = 1 To lvGroups.ListItems.Count
        ' ...
        If (lvGroups.ListItems(I).Selected) Then
            ' ...
            With lvGroups.ListItems(I)
                ' ...
                If (.Checked = True) Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End With
        End If
    
        ' ...
        If (StrComp(cbxGroups.text, lvGroups.ListItems(I), vbTextCompare) = 0) Then
            ' ...
            With lvGroups.ListItems(I)
                .Checked = False
                .Selected = False
            End With
        End If
    Next I
    
    ' enable entry save command
    btnSave(1).Enabled = True
    
    ' ...
    Set trvUsers.DropHighlight = Nothing
    Set lvGroups.SelectedItem = Nothing
End Sub

' ...
'Private Sub lstGroups_Click()
'    ' ...
'    m_group_change = True
'
'    ' ...
'    If (lstGroups.SelCount = 0) Then
'        m_group_index = -1
'    ElseIf (lstGroups.SelCount = 1) Then
'        m_group_index = lstGroups.ListIndex
'    Else
'        MsgBox lstGroups.ListIndex
'    End If
'
'    ' ...
'    btnSave(1).Enabled = True
'End Sub

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
    On Error GoTo ERROR_HANDLER

    Dim newNode As node ' ...

    Dim I       As Integer ' ...
    Dim grp     As String  ' ...
    Dim j       As Integer ' ...
    Dim pos     As Integer ' ...
    Dim blnDuplicateFound As Boolean

    ' clear treeview
    Call trvUsers.Nodes.Clear
    
    ' create root node
    Call trvUsers.Nodes.Add(, , "Database", "Database")

    ' which tab index are we on?
    Select Case (tbsTabs.SelectedItem.Index)
        Case 1: ' Users and Groups
            For I = LBound(m_DB) To UBound(m_DB)
                ' we're handling groups first; is this entry a group?
                If (StrComp(m_DB(I).Type, "GROUP", vbBinaryCompare) = 0) Then
                    ' is this group a member of other groups?
                    If (Len(m_DB(I).Groups) And (m_DB(I).Groups <> "%")) Then
                        ' is entry member of multiple groups?
                        If (InStr(1, m_DB(I).Groups, ",", vbBinaryCompare) <> 0) Then
                            ' split up multiple groupings
                            grp = Split(m_DB(I).Groups, ",", 2)(0)
                        Else
                            ' no need for special handling...
                            grp = m_DB(I).Groups
                        End If
                    
                        ' has the group already been added or is database in an
                        ' incorrect order?
                        pos = Exists(grp, "Group")
                        
                        ' ... well, does it exist?
                        If (pos) Then
                            ' make node a child of existing group
                            Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(pos).Key, _
                                tvwChild, "Group: " & m_DB(I).Username, m_DB(I).Username, 1)
                        Else
                            ' lets make this guy a parent node for now until we can find
                            ' his real parent.
                            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                                "Group: " & m_DB(I).Username, m_DB(I).Username, 1)
                        End If
                    Else
                        Dim k   As Integer ' ...
                        Dim bln As Boolean ' ...
                    
                        ' Okay, is the group a lone ranger?  Or does he have children
                        ' that are already in the list?
                        For j = LBound(m_DB) To (I - 1)
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
                                    If (StrComp(grp, m_DB(I).Username, vbTextCompare) = 0) Then
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
                            m_DB(I).Username, m_DB(I).Username, 1)
                    
                        ' is this node a baby's daddy?
                        If (bln) Then
                            ' move node
                            Set trvUsers.Nodes(Exists(m_DB(j).Username, "Group")).Parent = _
                                newNode
                        End If

                        ' reset boolean
                        bln = False
                    End If
                    
                    ' ...
                    If (Not (newNode Is Nothing)) Then
                        ' change misc. settings
                        With newNode
                            .Tag = "Group"
                        End With
                    End If
                End If
            Next I

            ' loop through database... this time looking for users.
            For I = LBound(m_DB) To UBound(m_DB)
                ' is the entry a user?
                If (StrComp(m_DB(I).Type, "USER", vbTextCompare) = 0) Then
                    ' is the user a member of any groups?
                    If (Len(m_DB(I).Groups) And (m_DB(I).Groups <> "%")) Then
                        ' is entry member of multiple groups?
                        If (InStr(1, m_DB(I).Groups, ",", vbBinaryCompare) <> 0) Then
                            ' split up multiple groupings
                            grp = Split(m_DB(I).Groups, ",", 2)(0)
                        Else
                            ' no need for special handling...
                            grp = m_DB(I).Groups
                        End If
                        
                        ' ...
                        If (grp = vbNullString) Then
                            pos = False
                        Else
                            ' search for our group
                            pos = Exists(grp, "Group")
                                
                            ' does our group exist?
                            If (pos) Then
                                ' create user node and move into group
                                Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(pos).Key, _
                                    tvwChild, "User: " & m_DB(I).Username, m_DB(I).Username, 3)
                            End If
                        End If
                    End If
                    
                    ' ...
                    If (pos = False) Then
                        ' create new user node under root
                        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                            "User: " & m_DB(I).Username, m_DB(I).Username, 3)
                    End If
                    
                    ' ...
                    If (Not (newNode Is Nothing)) Then
                        ' change misc. settings
                        With newNode
                            .Tag = "User"
                        End With
                    End If
                End If
                
                ' reset our variables
                pos = False
            Next I
            
            ' enable create user button
            btnCreateUser.Enabled = True
            
        Case 2: ' Clans
            ' loop through database searching for clans
            For I = LBound(m_DB) To UBound(m_DB)
                ' is entry a clan?
                If (StrComp(m_DB(I).Type, "CLAN", vbBinaryCompare) = 0) Then
                    ' create new node
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "Clan: " & m_DB(I).Username, m_DB(I).Username, 2)
                    
                    ' ...
                    If (Not (newNode Is Nothing)) Then
                         ' change misc. settings
                         With newNode
                             .Tag = "Clan"
                         End With
                     End If
                End If
            Next I

            ' disable create user button
            btnCreateUser.Enabled = False
            
        Case 3: ' Games
            ' loop through database searching for games
            For I = LBound(m_DB) To UBound(m_DB)
                ' is entry a game?
                If (StrComp(m_DB(I).Type, "GAME", vbBinaryCompare) = 0) Then
                    ' create new node
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "Group: " & m_DB(I).Username, m_DB(I).Username, 2)
                    
                    ' ...
                    If (Not (newNode Is Nothing)) Then
                        ' change misc. settings
                        With newNode
                            .Tag = "Game"
                        End With
                    End If
                End If
            Next I
            
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
    
    ' ...
    If (blnDuplicateFound = True) Then
        ' ...
        MsgBox "There were one or more duplicate database entries found which could not be loaded.", _
            vbExclamation, "Error"
    End If
    
    ' ...
    Exit Sub
    
ERROR_HANDLER:
    ' ...
    If (Err.Number = 35602) Then
        ' ...
        DB_remove m_DB(I).Username, m_DB(I).Type
    
        ' ...
        blnDuplicateFound = True
        
        ' ...
        Resume Next
    End If

    ' ...
    Exit Sub
End Sub

Private Sub LockGUI()
    Dim I As Integer ' ...
    
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
    Call ClearGroupList
    
    ' disable group lists
    lvGroups.Enabled = False
    cbxGroups.Enabled = False
    
    ' disable & clear ban message
    txtBanMessage.Enabled = False
    txtBanMessage.text = vbNullString
    
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

Private Sub ClearGroupList()
    Dim I As Integer ' ...
    
    ' ...
    cbxGroups.ListIndex = 0

    ' loop through listbox and clear selected items
    For I = 1 To lvGroups.ListItems.Count
        ' ...
        With lvGroups.ListItems(I)
            .Checked = False
            .Ghosted = False
        End With
    Next I
End Sub

Private Sub UnlockGUI()
    Dim I As Integer ' ...

    ' enable rank field
    txtRank.Enabled = True

    ' enable flags field
    txtFlags.Enabled = True
    
    ' enable ban message field
    txtBanMessage.Enabled = True
    
    ' disable entry save button
    btnSave(1).Enabled = False
    
    ' enable entry delete button
    btnDelete.Enabled = True
    
    ' enable group lists
    lvGroups.Enabled = True
    cbxGroups.Enabled = True
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
    Dim I   As Integer ' ...
    
    ' ...
    If (node Is Nothing) Then
        Exit Sub
    End If
    
    ' ...
    Call LockGUI
    
    ' ...
    m_group_index = -1

    ' ...
    If (node.Index > 1) Then
        ' ...
        Call ClearGroupList
    
        ' ...
        frmDatabase.Caption = node.text
        
        ' ...
        tmp = GetAccess(node.text, node.Tag)
        
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
            Dim Splt() As String  ' ...
            Dim j      As Integer ' ...
        
            ' is entry a member of multiple groups?
            If (InStr(1, tmp.Groups, ",", vbBinaryCompare) <> 0) Then
                ' store working copy of group memberships, splitting up
                ' multiple groupings by the ',' delimiter.
                Splt() = Split(tmp.Groups, ",")
            Else
                ' redefine array size to store group name
                ReDim Preserve Splt(0)
                
                ' store working copy of group membership
                Splt(0) = tmp.Groups
            End If
            
            ' loop through entry's group memberships
            For I = LBound(Splt) To UBound(Splt)
                ' ...
                If (I = 0) Then
                    ' loop through our group listing, checking to see if we have any
                    ' matches (since the entry is a member of a group, we better!)
                    For j = 1 To lvGroups.ListItems.Count
                        ' is entry a member of group?
                        If (StrComp(Splt(I), cbxGroups.List(j), vbTextCompare) = 0) Then
                            ' ...
                            cbxGroups.ListIndex = j
                            
                            ' ...
                            Exit For
                        End If
                    Next j
                Else
                    ' loop through our group listing, checking to see if we have any
                    ' matches (since the entry is a member of a group, we better!)
                    For j = 1 To lvGroups.ListItems.Count
                        ' ...
                        If (StrComp(Splt(I), lvGroups.ListItems(j), vbTextCompare) = 0) Then
                            ' select group if entry is a member
                            lvGroups.ListItems(j).Checked = True
                        End If
                    Next j
                End If
            Next I
        End If
        
        ' ...
        If ((tmp.BanMessage <> vbNullString) And (tmp.BanMessage <> "%")) Then
            txtBanMessage.text = tmp.BanMessage
        End If
        
        ' ...
        Call UnlockGUI
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
            If (trvUsers.SelectedItem.Index > 1) Then
                ' ...
                If (StrComp(trvUsers.SelectedItem.Tag, "Group", vbTextCompare) = 0) Then
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
    Dim I        As Integer ' ...
    Dim found    As Integer ' ...

    ' ...
    If (Data.GetFormat(15) = True) Then
        ' ...
        With frmDBType
            .setFilePath Data.Files(1)
            .Show
        End With
    Else
        ' ...
        Set nodeNow = trvUsers.DropHighlight
        
        ' ...
        Set nodePrev = trvUsers.SelectedItem
            
        ' ...
        If (nodeNow.Index = 1) Then
            ' ...
            For I = LBound(m_DB) To UBound(m_DB)
                ' ...
                If (StrComp(m_DB(I).Username, nodePrev.text, vbTextCompare) = 0) Then
                    ' ...
                    If (StrComp(m_DB(I).Type, nodePrev.Tag, vbTextCompare) = 0) Then
                        ' ...
                        If ((Len(m_DB(I).Groups) > 0) And (m_DB(I).Groups <> "%")) Then
                            ' ...
                            Set nodePrev.Parent = nodeNow
                        End If
                        
                        ' ...
                        m_DB(I).Groups = vbNullString
                        
                        ' ...
                        Exit For
                    End If
                End If
            Next I
        Else
            ' ...
            If (nodePrev.Index <> 1) Then
                ' ...
                If (StrComp(nodeNow.Tag, "Group", vbTextCompare) <> 0) Then
                    ' ...
                    Set nodeNow = nodeNow.Parent
                    
                    ' ...
                    If (nodeNow.Index = 1) Then
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
                    For I = LBound(m_DB) To UBound(m_DB)
                        ' ...
                        If (StrComp(m_DB(I).Username, nodePrev.text, vbTextCompare) = 0) Then
                            ' ...
                            If (StrComp(m_DB(I).Type, nodePrev.Tag, vbTextCompare) = 0) Then
                                ' ...
                                m_DB(I).Groups = nodeNow.text
                                
                                ' ...
                                Exit For
                            End If
                        End If
                    Next I
                    
                    ' ...
                    Set nodePrev.Parent = nodeNow
                End If
            End If
        End If
        
        ' ...
        Call trvUsers_NodeClick(nodePrev)
        
        ' ...
        Set trvUsers.DropHighlight = Nothing
    End If
    
    Exit Sub

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
    If (StrComp(trvUsers.SelectedItem.Tag, "Group", vbTextCompare) <> 0) Then
        ' disallow editing of entry label
        Cancel = 1
    End If
End Sub

' ...
Private Sub trvUsers_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim I As Integer ' ...
    
    ' ...
    If (trvUsers.SelectedItem Is Nothing) Then
        ' ...
        Exit Sub
    End If
    
    ' ...
    If ((NewString = vbNullString) Or (Len(NewString) > 30)) Then
        ' ...
        MsgBox "The specified name is of an invalid Length."

        ' ...
        Call trvUsers.StartLabelEdit
        
        ' ...
        Cancel = 1
    
        ' ...
        Exit Sub
    End If
    
    ' ...
    If (InStr(1, NewString, Space$(1), vbBinaryCompare <> 0)) Or _
       (InStr(1, NewString, "%", vbBinaryCompare <> 0)) Then
    
        ' ...
        MsgBox "The specified name contains one or more invalid characters."
        
        ' ...
        Call trvUsers.StartLabelEdit
        
        ' ...
        Cancel = 1
    
        ' ...
        Exit Sub
    End If
    
    ' ...
    If (m_new_entry) Then
        ' ...
        If (GetAccess(NewString, trvUsers.SelectedItem.Tag).Username = vbNullString) Then
            ' ...
            With m_DB(UBound(m_DB))
                .Username = NewString
            End With
            
            ' ...
            If (StrComp(m_DB(UBound(m_DB)).Type, "Group", vbTextCompare) = 0) Then
                Call UpdateGroupListBox
            End If
            
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
        For I = LBound(m_DB) To UBound(m_DB)
            ' ...
            If (StrComp(trvUsers.SelectedItem.text, m_DB(I).Username, vbTextCompare) = 0) Then
                ' ...
                If (StrComp(trvUsers.SelectedItem.Tag, m_DB(I).Type, vbTextCompare) = 0) Then
                    ' ...
                    m_DB(I).Username = NewString
            
                    ' ...
                    Exit For
                End If
            End If
        Next I
        
        ' ...
        If (StrComp(m_DB(I).Type, "Group", vbTextCompare) = 0) Then
            ' ...
            For I = LBound(m_DB) To UBound(m_DB)
                ' ...
                If ((Len(m_DB(I).Groups) > 0) And (m_DB(I).Groups <> "%")) Then
                    Dim Splt() As String  ' ...
                    Dim j      As Integer ' ...
                
                    ' ...
                    If (InStr(1, m_DB(I).Groups, ",", vbTextCompare) <> 0) Then
                        ' ...
                        Splt() = Split(m_DB(I).Groups, ",")
                    Else
                        ' ...
                        ReDim Preserve Splt(0)
                        
                        ' ...
                        Splt(0) = m_DB(I).Groups
                    End If
                    
                    ' ...
                    For j = LBound(Splt) To UBound(Splt)
                        ' ...
                        If (StrComp(Splt(j), trvUsers.SelectedItem.text, vbTextCompare) = 0) Then
                            ' ...
                            Splt(j) = NewString
                        End If
                    Next j
                    
                    ' ...
                    m_DB(I).Groups = Join(Splt(), ",")
                End If
            Next I
        End If
    End If
End Sub

Private Function IsInGroup(ByVal Username As String, ByVal GroupName As String) As Boolean
    Dim I      As Integer ' ...
    Dim j      As Integer ' ...
    Dim Splt() As String  ' ...
    
    ' ...
    For I = LBound(m_DB) To UBound(m_DB)
        ' ...
        If (StrComp(m_DB(I).Username, Username, vbTextCompare) = 0) Then
            ' ...
            If ((Len(m_DB(I).Groups) > 0) And (m_DB(I).Groups <> "%")) Then
                ' ...
                If (InStr(1, m_DB(I).Groups, "%", vbBinaryCompare) <> 0) Then
                    ' ...
                    Splt() = Split(m_DB(I).Groups, "%")
                Else
                    ' ...
                    ReDim Splt(0)
                    
                    ' ...
                    Splt(0) = m_DB(I).Groups
                End If
                
                ' ...
                For j = LBound(Splt) To UBound(Splt)
                    If (StrComp(GroupName, Splt(j), vbTextCompare) = 0) Then
                        ' ...
                        IsInGroup = True
                        
                        ' ...
                        Exit Function
                    End If
                Next j
            End If
        End If
    Next I
End Function

Private Sub HandleDeleteEvent(ByRef NodeToDelete As node)
    Dim temp As node ' ...
    
    ' ...
    Set temp = NodeToDelete

    ' ...
    If (temp Is Nothing) Then
        Exit Sub
    End If

    ' ...
    If (temp.Index > 1) Then
        Dim response As Integer ' ...
        Dim isGroup  As Boolean ' ...
        
        ' ...
        isGroup = (StrComp(temp.Tag, "Group", vbTextCompare) = 0)
    
        ' ...
        If (isGroup) Then
            ' ...
            response = MsgBox("Are you sure you wish to delete this group and " & _
                "all of its members?", vbYesNo + vbInformation, "Information")
        End If
        
        ' ...
        If ((isGroup = False) Or ((isGroup) And (response = vbYes))) Then
            ' ...
            Call DB_remove(temp.text, temp.Tag)
            
            ' ...
            If (temp.Next Is Nothing) Then
                ' ...
                If (temp.Previous Is Nothing) Then
                    ' ...
                    trvUsers.Nodes(temp.Parent.Index).Checked = True
                Else
                    ' ...
                    trvUsers.Nodes(temp.Previous.Index).Checked = True
                End If
            Else
                ' ...
                trvUsers.Nodes(temp.Next.Index).Checked = True
            End If
            
            ' ...
            Call trvUsers.Nodes.Remove(temp.Index)
            
            ' ...
            Call trvUsers_NodeClick(trvUsers.SelectedItem)
            
            ' ...
            Call UpdateGroupListBox
        End If
    End If
End Sub

Private Sub UpdateGroupListBox()
    Dim I As Integer ' ...

    ' clear group selection listing
    Call lvGroups.ListItems.Clear
    
    ' ...
    Call cbxGroups.Clear
    
    ' ...
    Call cbxGroups.AddItem("[none]", 0)

    ' go through group listing
    For I = LBound(m_DB) To UBound(m_DB)
        ' ...
        If (StrComp(m_DB(I).Type, "Group", vbTextCompare) = 0) Then
            ' add group to group selection listbox
            Call lvGroups.ListItems.Add(, , m_DB(I).Username)
            
            ' ...
            Call cbxGroups.AddItem(m_DB(I).Username)
        End If
    Next I
End Sub

' ...
Private Function Exists(ByVal nodeName As String, Optional Tag As String = vbNullString) As Integer
    Dim I As Integer ' ...
    
    ' ...
    For I = 1 To trvUsers.Nodes.Count
        ' ...
        If (StrComp(trvUsers.Nodes(I).text, nodeName, vbTextCompare) = 0) Then
            ' ...
            If (Tag <> vbNullString) Then
                ' ...
                If (StrComp(trvUsers.Nodes(I).Tag, Tag, vbTextCompare) = 0) Then
                    ' ...
                    Exists = I
                    
                    ' ...
                    Exit Function
                End If
            Else
                ' ...
                Exists = I
            
                ' ...
                Exit Function
            End If
        End If
    Next I
    
    ' ...
    Exists = False
End Function

Private Sub txtBanMessage_Change()
    ' enable entry save button
    btnSave(1).Enabled = True
End Sub

' ...
Private Sub txtFlags_Change()

    If (BotVars.CaseSensitiveFlags = False) Then
        txtFlags.text = UCase$(txtFlags.text)
    End If

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
    
    Dim I   As Integer ' ...
    Dim bln As Boolean ' ...
    
    dbType = UCase$(dbType)

    For I = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(I).Username, Username, vbTextCompare) = 0) Then
            If (Len(dbType) > 0) Then
                If (StrComp(m_DB(I).Type, dbType, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
                
            If (bln = True) Then
                With GetAccess
                    .Username = m_DB(I).Username
                    .Access = m_DB(I).Access
                    .Flags = m_DB(I).Flags
                    .AddedBy = m_DB(I).AddedBy
                    .AddedOn = m_DB(I).AddedOn
                    .ModifiedBy = m_DB(I).ModifiedBy
                    .ModifiedOn = m_DB(I).ModifiedOn
                    .Type = m_DB(I).Type
                    .Groups = m_DB(I).Groups
                    .BanMessage = m_DB(I).BanMessage
                End With
                
                Exit Function
            End If
        End If
        
        bln = False
    Next I

    GetAccess.Access = -1
End Function

' ...
Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = _
    vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER

    Dim I     As Integer ' ...
    Dim found As Boolean ' ...
    
    dbType = UCase$(dbType)
    
    For I = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(I).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean ' ...
        
            If (Len(dbType)) Then
                If (StrComp(m_DB(I).Type, dbType, vbTextCompare) = 0) Then
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
    Next I
    
    If (found) Then
        Dim bak As udtDatabase ' ...
        
        Dim j   As Integer ' ...
        
        ' ...
        bak = m_DB(I)

        ' we aren't removing the last array
        ' element, are we?
        If (UBound(m_DB) = 0) Then
            ' ...
            ReDim m_DB(0)
            
            ' ...
            With m_DB(0)
                .Username = vbNullString
                .Flags = vbNullString
                .Access = 0
                .Groups = vbNullString
                .AddedBy = vbNullString
                .ModifiedBy = vbNullString
                .AddedOn = Now
                .ModifiedOn = Now
            End With
        Else
            ' ...
            For j = I To UBound(m_DB) - 1
                m_DB(j) = m_DB(j + 1)
            Next j
            
            ' redefine array size
            ReDim Preserve m_DB(UBound(m_DB) - 1)
            
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
                    For I = LBound(m_DB) To UBound(m_DB)
                        If (Len(m_DB(I).Groups) And m_DB(I).Groups <> "%") Then
                            If (InStr(1, m_DB(I).Groups, ",", vbBinaryCompare) <> 0) Then
                                Dim Splt()     As String ' ...
                                Dim innerfound As Boolean ' ...
                                
                                Splt() = Split(m_DB(I).Groups, ",")
                                
                                For j = LBound(Splt) To UBound(Splt)
                                    If (StrComp(bak.Username, Splt(j), vbTextCompare) = 0) Then
                                        innerfound = True
                                    
                                        Exit For
                                    End If
                                Next j
                            
                                If (innerfound) Then
                                    Dim k As Integer ' ...
                                    
                                    For k = (j + 1) To UBound(Splt)
                                        Splt(k - 1) = Splt(k)
                                    Next k
                                    
                                    ReDim Preserve Splt(UBound(Splt) - 1)
                                    
                                    m_DB(I).Groups = Join(Splt(), vbNullString)
                                End If
                            Else
                                If (StrComp(bak.Username, m_DB(I).Groups, vbTextCompare) = 0) Then
                                    res = DB_remove(m_DB(I).Username, m_DB(I).Type)
                                    
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                Loop While (res)
            End If
        End If
        
        ' commit modifications
        'Call WriteDatabase(GetFilePath("users.txt"))
        
        DB_remove = True
        
        Exit Function
    End If
    
    DB_remove = False
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.description & " in DB_remove().")
        
    DB_remove = False
    
    Exit Function
End Function
