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
      Begin VB.TextBox txtFlags 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   580
         Width           =   1215
      End
      Begin VB.ListBox lstGroups 
         Enabled         =   0   'False
         Height          =   2010
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   2450
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1930
         TabIndex        =   10
         Top             =   4535
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1088
         TabIndex        =   11
         Top             =   4535
         Width           =   855
      End
      Begin VB.Label lblModifiedBy 
         Caption         =   "by Eric[nK]"
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
         TabIndex        =   20
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Label lblCreatedBy 
         Caption         =   "by Eric[nK]"
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
         Top             =   1350
         Width           =   2415
      End
      Begin VB.Label lblCreatedOn 
         Caption         =   "12/27/2007 at 11:32 P.M. Local Time"
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
         TabIndex        =   15
         Top             =   1180
         Width           =   2415
      End
      Begin VB.Label lblModifiedOn 
         Caption         =   "12/27/2007 at 11:32 P.M. Local Time"
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1590
         Width           =   2535
      End
      Begin VB.Label lblGroup 
         Caption         =   "Member of Group(s):"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2240
         Width           =   2535
      End
      Begin VB.Label lblFlags 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   340
         Width           =   1215
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank (1 - 200):"
         Height          =   255
         Left            =   240
         TabIndex        =   12
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "Apply and Cl&ose"
      Height          =   300
      Index           =   0
      Left            =   5280
      TabIndex        =   4
      Top             =   5540
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
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
      Indentation     =   575
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
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
Option Explicit

Public m_game      As String

Private m_DB()     As udtDatabase
Private m_Modified As Boolean
Private m_DBDate   As Long

Private Sub btnCreateUser_Click()
    Static userCount As Integer ' ...
    
    Dim newNode      As Node    ' ...
    Dim gAcc         As udtGetAccessResponse
    
    Dim Username     As String  ' ...
    
    Username = "New User #" & _
        (userCount + 1)
        
    ReDim Preserve m_DB(UBound(m_DB) + 1)
    
    With m_DB(UBound(m_DB))
        .Username = vbNullString
        .type = "USER"
        .AddedBy = "(console)"
        .AddedOn = Now
    End With

    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        If (trvUsers.SelectedItem.Index = 1) Then
            Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                tvwChild, "U:" & Username, Username, 3)
        ElseIf (trvUsers.SelectedItem.Image = 1) Then
            Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                tvwChild, "U:" & Username, Username, 3)

            With m_DB(UBound(m_DB))
                .Groups = trvUsers.SelectedItem.text
            End With
        Else
            Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Parent.Key, _
                tvwChild, "U:" & Username, Username, 3)
                
            If (trvUsers.SelectedItem.Parent.Image = 1) Then
                With m_DB(UBound(m_DB))
                    .Groups = trvUsers.SelectedItem.text
                End With
            End If
        End If
    Else
        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
            "U:" & Username, Username, 3)
    End If
        
    With trvUsers.Nodes(newNode.Index)
        .Selected = True
    End With
    
    Call trvUsers.StartLabelEdit
    
    userCount = (userCount + 1)
End Sub

Private Sub btnCreateGroup_Click()
    Static groupCount As Integer ' ...
    
    Dim newNode       As Node    ' ...
    
    Dim groupname     As String  ' ...
    
    If (tbsTabs.SelectedItem.Index = 1) Then
        groupname = "New Group #" & _
            (groupCount + 1)
            
        ReDim Preserve m_DB(UBound(m_DB) + 1)
        
        With m_DB(UBound(m_DB))
            .Username = vbNullString
            .type = "GROUP"
            .AddedBy = "(console)"
            .AddedOn = Now
        End With
    
        If (Not (trvUsers.SelectedItem Is Nothing)) Then
            If (trvUsers.SelectedItem.Index = 1) Then
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                    tvwChild, "G:" & groupname, groupname, 1)
            ElseIf (trvUsers.SelectedItem.Image = 1) Then
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
                    tvwChild, "G:" & groupname, groupname, 1)

                With m_DB(UBound(m_DB))
                    .Groups = trvUsers.SelectedItem.text
                End With
            Else
                Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Parent.Key, _
                    tvwChild, "G:" & groupname, groupname, 1)
                    
                If (trvUsers.SelectedItem.Parent.Image = 1) Then
                    With m_DB(UBound(m_DB))
                        .Groups = trvUsers.SelectedItem.Parent.text
                    End With
                End If
            End If
        Else
            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                "G:" & groupname, groupname, 1)
        End If
        
        ' ...
        Set trvUsers.SelectedItem = newNode
        
        ' ...
        Call trvUsers.StartLabelEdit
        
        ' ...
        groupCount = (groupCount + 1)
    ElseIf (tbsTabs.SelectedItem.Index = 2) Then
    
    ElseIf (tbsTabs.SelectedItem.Index = 3) Then
        ' ...
        Call frmGameSelection.Show(vbModal, frmDBManager)
        
        ' ...
        If (Len(m_game)) Then
            If (GetAccess(m_game, "GAME").Username = _
                vbNullString) Then
                
                ReDim Preserve m_DB(UBound(m_DB) + 1)
                
                With m_DB(UBound(m_DB))
                    .Username = m_game
                    .type = "GAME"
                    .AddedBy = "(console)"
                    .AddedOn = Now
                End With
            
                Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                    "G:" & m_game, m_game, 2)
                
                ' ...
                Set trvUsers.SelectedItem = newNode
                
                ' ...
                Call trvUsers_NodeClick(trvUsers.SelectedItem)
                
                ' ...
                Call trvUsers.SetFocus
            Else
                MsgBox "There is already an entry of this type matching the specified name."
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    If (Index = 0) Then
        Call Unload(frmDBManager)
    Else
        Dim response As Integer ' ...
        Dim isGroup  As Boolean ' ...
        
        isGroup = (trvUsers.Nodes(trvUsers.SelectedItem.Index).Image = 1)
    
        If (isGroup) Then
            response = MsgBox("Are you sure you wish to delete " & _
                "this group and " & "all of its members?", vbYesNo + _
                    vbInformation, "Information")
            
            If (response = vbYes) Then
                Call DB_remove(trvUsers.SelectedItem.text)

                Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                
                Call trvUsers_NodeClick(trvUsers.SelectedItem)
            End If
        Else
            Call DB_remove(trvUsers.SelectedItem.text)

            Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
            
            Call trvUsers_NodeClick(trvUsers.SelectedItem)
        End If
    End If
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Dim i As Integer ' ...

    If (Index = 1) Then
        For i = LBound(m_DB()) To UBound(m_DB())
            If (StrComp(trvUsers.SelectedItem.text, m_DB(i).Username, _
                vbTextCompare) = 0) Then
                
                With m_DB(i)
                    .Access = Val(txtRank.text)
                    .flags = txtFlags.text
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                End With
                
                Exit For
            End If
        Next i
        
        cmdSave(1).Enabled = False
    Else
        DB() = m_DB()
        
        Call checkUsers
    
        Call WriteDatabase(GetFilePath("users.txt"))
        
        Call Unload(frmDBManager)
    End If
End Sub

Private Sub Form_Load()
    If (DB(0).Username = vbNullString) Then
        Call LoadDatabase
    End If
    
    m_DB() = DB()
    
    Call tbsTabs_Click
End Sub

Private Sub lstGroups_Click()
    cmdSave(1).Enabled = True
End Sub

Private Sub mnuDelete_Click()
    If (trvUsers.SelectedItem.Index > 1) Then
        Dim response As Integer ' ...
        Dim isGroup  As Boolean ' ...
        
        isGroup = (trvUsers.Nodes(trvUsers.SelectedItem.Index).Image = 1)
    
        If (isGroup) Then
            response = MsgBox("Are you sure you wish to delete " & _
                "this group and " & "all of its members?", vbYesNo + _
                    vbInformation, "Information")
            
            If (response = vbYes) Then
                Call DB_remove(trvUsers.SelectedItem.text)

                Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                
                Call trvUsers_NodeClick(trvUsers.SelectedItem)
            End If
        Else
            Call DB_remove(trvUsers.SelectedItem.text)

            Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
            
            Call trvUsers_NodeClick(trvUsers.SelectedItem)
        End If
    End If
End Sub

Private Sub mnuOpenDatabase_Click()
    Call CommonDialog.ShowOpen
End Sub

Private Sub mnuRename_Click()
    Call trvUsers.StartLabelEdit
End Sub

Private Sub tbsTabs_Click()
    Dim newNode As Node ' ...

    Dim i       As Integer ' ...
    Dim Splt()  As String  ' ...
    Dim j       As Integer ' ...
    Dim Pos     As Integer ' ...

    Call trvUsers.Nodes.Clear
    
    Call trvUsers.Nodes.Add(, , "Database", "Database")
    
    trvUsers.Nodes(1).Sorted = True
    
    Select Case (tbsTabs.SelectedItem.Index)
        Case 1: ' Users and Groups
            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(m_DB(i).type, "GROUP", vbBinaryCompare) = 0) Then
                    If (Len(m_DB(i).Groups) And (m_DB(i).Groups <> "%")) Then
                        If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Splt() = Split(m_DB(i).Groups, ",")
                        Else
                            ReDim Preserve Splt(0)
                            
                            Splt(0) = m_DB(i).Groups
                        End If
                        
                        For j = LBound(Splt) To UBound(Splt)
                            Pos = Exists(Splt(j))
                            
                            If (Pos) Then
                                Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(Pos).Key, _
                                    tvwChild, "G:" & m_DB(i).Username, m_DB(i).Username, 1)
                            Else
                                Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                                    "G:" & m_DB(i).Username, m_DB(i).Username, 1)
                            End If
                        Next j
                    Else
                        Dim k   As Integer ' ...
                        Dim bln As Boolean ' ...
                    
                        For j = LBound(m_DB) To (i - 1)
                            If (StrComp(m_DB(j).type, "GROUP", vbBinaryCompare) = 0) Then
                                If (Len(m_DB(j).Groups) And (m_DB(j).Groups <> "%")) Then
                                    If (InStr(1, m_DB(j).Groups, ",", vbBinaryCompare) <> 0) Then
                                        Splt() = Split(m_DB(j).Groups, ",")
                                    Else
                                        ReDim Preserve Splt(0)
                                        
                                        Splt(0) = m_DB(j).Groups
                                    End If
                                    
                                    For k = LBound(Splt) To UBound(Splt)
                                        If (StrComp(Splt(k), m_DB(i).Username, _
                                            vbTextCompare) = 0) Then
                                        
                                            bln = True
                                            
                                            Exit For
                                        End If
                                    Next k
                                    
                                    If (bln) Then
                                        Exit For
                                    End If
                                End If
                            End If
                        Next j
                        
                        If (Not (Exists(m_DB(i).Username))) Then
                            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                                "G:" & m_DB(i).Username, m_DB(i).Username, 1)
                        
                            If (bln) Then
                                If (Exists(m_DB(j).Username)) Then
                                    Set trvUsers.Nodes(Exists(m_DB(j).Username)).Parent = _
                                        newNode
                                End If
                            End If
                        End If
                        
                        bln = False
                    End If
                End If
            Next i
            
            If (trvUsers.Nodes.Count > 1) Then
                Call lstGroups.Clear
            
                For i = 2 To trvUsers.Nodes.Count
                    Call lstGroups.AddItem(trvUsers.Nodes(i).text)
                Next i
            End If
            
            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(m_DB(i).type, "USER", vbBinaryCompare) = 0) Then
                    If (Len(m_DB(i).Groups) And (m_DB(i).Groups <> "%")) Then
                        If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Splt() = Split(m_DB(i).Groups, ",")
                        Else
                            ReDim Preserve Splt(0)
                            
                            Splt(0) = m_DB(i).Groups
                        End If
                        
                        For j = LBound(Splt) To UBound(Splt)
                            Pos = Exists(Splt(j))
                            
                            If (Pos) Then
                                Set newNode = trvUsers.Nodes.Add(trvUsers.Nodes(Pos).Key, _
                                    tvwChild, "U:" & m_DB(i).Username, m_DB(i).Username, 3)
                            End If
                        Next j
                    Else
                        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                            "U:" & m_DB(i).Username, m_DB(i).Username, 3)
                    End If
                End If
            Next i
            
            'btnCreateGroup.Caption = "Create Group"
            
            btnCreateUser.Enabled = True
            
        Case 2: ' Clans
            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(m_DB(i).type, "CLAN", vbBinaryCompare) = 0) Then
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "G:" & m_DB(i).Username, m_DB(i).Username, 2)
                End If
            Next i
            
            'btnCreateGroup.Caption = "Create Clan"
            
            btnCreateUser.Enabled = False
            
        Case 3: ' Games
            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(m_DB(i).type, "GAME", vbBinaryCompare) = 0) Then
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        "G:" & m_DB(i).Username, m_DB(i).Username, 2)
                End If
            Next i
            
            'btnCreateGroup.Caption = "Create Game"
            
            btnCreateUser.Enabled = False
    End Select
    
    If (trvUsers.Nodes.Count) Then
        With trvUsers.Nodes(1)
            .Expanded = True
            .Image = 1
        End With
        
        Call trvUsers.Refresh
    End If
    
    With frmDatabase
        .Caption = "Database"
    End With
    
    txtRank.Enabled = False
    txtRank.text = vbNullString
    
    txtFlags.Enabled = False
    txtFlags.text = vbNullString
    
    lstGroups.Enabled = False
    
    cmdSave(1).Enabled = False
    cmdCancel(1).Enabled = False
End Sub

Private Sub trvUsers_Collapse(ByVal Node As Node)
    Call trvUsers.Refresh
End Sub

Private Sub trvUsers_Expand(ByVal Node As Node)
    Call trvUsers.Refresh
End Sub

Private Sub trvUsers_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim tmp As udtGetAccessResponse ' ...
    
    Dim i   As Integer ' ...
    
    ' deselect groups
    For i = 0 To (lstGroups.ListCount - 1)
        lstGroups.Selected(i) = False
    Next i

    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' grab entry from database
        tmp = GetAccess(trvUsers.SelectedItem.text)

        If (Node.Index = 1) Then
            With frmDatabase
                .Caption = "Database"
            End With
            
            txtRank.Enabled = False
            
            txtFlags.Enabled = False
            
            lstGroups.Enabled = False
            
            cmdCancel(1).Enabled = False
            
            lblCreatedOn = "n/a"
                
            lblCreatedBy = vbNullString
            
            lblModifiedOn = "n/a"
                
            lblModifiedBy = vbNullString
        Else
            If (tmp.type = "USER") Then
                frmDatabase.Caption = tmp.Username
            ElseIf (tmp.type = "CLAN") Then
                frmDatabase.Caption = tmp.Username
            ElseIf (tmp.type = "GAME") Then
                frmDatabase.Caption = tmp.Username
            ElseIf (tmp.type = "GROUP") Then
                frmDatabase.Caption = tmp.Username
            Else
                frmDatabase.Caption = tmp.Username
            End If
            
            txtRank.Enabled = True
            
            txtFlags.Enabled = True
            
            lstGroups.Enabled = True
            
            cmdCancel(1).Enabled = True
            
            If (tmp.AddedBy = "%") Then
                lblCreatedOn = "unknown"
                
                lblCreatedBy = "by unknown"
            Else
                lblCreatedOn = tmp.AddedOn & _
                    " Local Time"
                    
                lblCreatedBy = "by " & _
                    tmp.AddedBy
            End If
            
            If (tmp.ModifiedBy = "%") Then
                lblModifiedOn = "unknown"
                
                lblModifiedBy = "by unknown"
            Else
                lblModifiedOn = tmp.ModifiedOn & _
                    " Local Time"
                    
                lblModifiedBy = "by " & _
                    tmp.ModifiedBy
            End If
            
            If (Len(tmp.Groups) And (tmp.Groups <> "%")) Then
                Dim Splt() As String  ' ...
                Dim j      As Integer ' ...
            
                If (InStr(1, tmp.Groups, ",", vbBinaryCompare) <> 0) Then
                    Splt() = Split(m_DB(i).Groups, ",")
                Else
                    ReDim Preserve Splt(0)
                    
                    Splt(0) = tmp.Groups
                End If
                
                For i = LBound(Splt) To UBound(Splt)
                    For j = 0 To (lstGroups.ListCount - 1)
                        If (StrComp(Splt(i), lstGroups.List(j), vbTextCompare) = 0) Then
                            lstGroups.Selected(j) = True
                        End If
                    Next j
                Next i
            End If
            
        End If
        
        If (tmp.Access > 0) Then
            txtRank.text = tmp.Access
        Else
            txtRank.text = vbNullString
        End If
        
        txtFlags.text = tmp.flags
    
        cmdSave(1).Enabled = False
        
        Call trvUsers.Refresh
    End If
End Sub

Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)

    If (Button = vbLeftButton) Then
        Set trvUsers.SelectedItem = _
            trvUsers.HitTest(X, Y)
            
        Call trvUsers_NodeClick(trvUsers.SelectedItem)
    End If
End Sub

Private Sub trvUsers_OLEStartDrag(Data As MSComctlLib.DataObject, _
    AllowedEffects As Long)
    
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Data.Clear
    
        Data.SetData trvUsers.SelectedItem.Key, _
            vbCFText
    End If
End Sub

Private Sub trvUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
        Dim gAcc As udtGetAccessResponse ' ...

        If (trvUsers.SelectedItem.Index > 1) Then
            ' ...
            gAcc = GetAccess(trvUsers.SelectedItem.text)
    
            ' ...
            If (gAcc.type = "GROUP") Then
                mnuRename.Enabled = True
            Else
                mnuRename.Enabled = False
            End If
            
            mnuDelete.Enabled = True
        Else
            mnuRename.Enabled = False
            mnuDelete.Enabled = False
        End If
        
        ' ...
        Call Me.PopupMenu(mnuContext)
    End If
End Sub

Private Sub trvUsers_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, _
    Button As Integer, Shift As Integer, X As Single, Y As Single, _
    State As Integer)
    
    Set trvUsers.DropHighlight = _
        trvUsers.HitTest(X, Y)
End Sub

Private Sub trvUsers_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, _
    Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' ...
    On Error GoTo ERROR_HANDLER
      
    If ((Not (trvUsers.DropHighlight Is Nothing)) And _
        (Not (trvUsers.SelectedItem Is Nothing))) Then
        
        Dim nodeprev As Node ' ...
        Dim nodenow  As Node ' ...
    
        Dim selnow   As udtGetAccessResponse ' ...
        Dim selprev  As udtGetAccessResponse ' ...
        
        Dim strKey   As String  ' ...
        Dim res      As Integer ' ...
        Dim i        As Integer ' ...
        Dim found    As Integer ' ...

        If (Data.GetFormat(vbCFText)) Then
            strKey = Data.GetData(vbCFText)
            
            If (Len(strKey)) Then
                Set nodeprev = trvUsers.Nodes(strKey)
            End If
        End If
        
        Set nodenow = trvUsers.DropHighlight
        
        If (trvUsers.DropHighlight.Index = 1) Then
            If (nodeprev.text <> vbNullString) Then
                selprev = GetAccess(nodeprev.text)
            Else
                selprev = GetAccess(strKey)
            End If

            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(selprev.Username, m_DB(i).Username, _
                    vbBinaryCompare) = 0) Then
                    
                    m_DB(i).Groups = vbNullString
                End If
            Next i
            
            ' ...
            Set nodeprev.Parent = nodenow
        Else
            If (trvUsers.SelectedItem.Index <> 1) Then
                selnow = GetAccess(trvUsers.DropHighlight.text)
                selprev = GetAccess(trvUsers.SelectedItem.text)
                
                If (selnow.Username <> selprev.Username) Then
                    If (StrComp(selnow.type, "GROUP", vbBinaryCompare) = 0) Then
                        For i = LBound(m_DB) To UBound(m_DB)
                            If (StrComp(selprev.Username, m_DB(i).Username, vbBinaryCompare) = 0) Then
                                m_DB(i).Groups = nodenow.text
                            End If
                        Next i
                        
                        ' ...
                        Set nodeprev.Parent = nodenow
                    End If
                End If
            End If
        End If
        
        ' ...
        Set trvUsers.DropHighlight = _
            Nothing
    ElseIf (Not (trvUsers.DropHighlight Is Nothing)) Then
        MsgBox "!!"
    
        If (Data.Files.Count) Then
            MsgBox Data.Files(1)
        End If
        
        ' ...
        Set trvUsers.DropHighlight = _
            Nothing
    Else
        Call Data.GetFormat(vbCFText)
    
            If (Data.Files.Count) Then
                MsgBox Data.Files(1)
            End If

        MsgBox Data.GetFormat(vbCFText)
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    If (Err.Number = 35614) Then
        MsgBox Err.description, vbCritical, _
            "Error"
    End If
    
    Set trvUsers.DropHighlight = _
        Nothing
    
    Exit Sub
End Sub

Private Sub trvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) Then
        If (trvUsers.SelectedItem.Index > 1) Then
            Dim response As Integer ' ...
            Dim isGroup  As Boolean ' ...
            
            isGroup = (trvUsers.Nodes(trvUsers.SelectedItem.Index).Image = 1)
        
            If (isGroup) Then
                response = MsgBox("Are you sure you wish to delete " & _
                    "this group and " & "all of its members?", vbYesNo + _
                        vbInformation, "Information")
                
                If (response = vbYes) Then
                    Call DB_remove(trvUsers.SelectedItem.text)
    
                    Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                    
                    Call trvUsers_NodeClick(trvUsers.SelectedItem)
                End If
            Else
                Call DB_remove(trvUsers.SelectedItem.text)
    
                Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                
                Call trvUsers_NodeClick(trvUsers.SelectedItem)
            End If
        End If
    End If
End Sub

Private Sub trvUsers_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim i As Integer ' ...
    
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        If (m_DB(UBound(m_DB)).Username = vbNullString) Then
            If (GetAccess(NewString, m_DB(UBound(m_DB)).type).Username <> _
                vbNullString) Then
                
                MsgBox "There is already an entry of this type matching the specified name."
                
                Call trvUsers.StartLabelEdit
                
                Cancel = 1
            End If
        Else
            For i = LBound(m_DB) To UBound(m_DB)
                If (StrComp(trvUsers.SelectedItem.text, m_DB(i).Username, _
                        vbTextCompare) = 0) Then
            
                    m_DB(i).Username = NewString
            
                    Exit For
                End If
            Next i
            
            If (StrComp(m_DB(i).type, "GROUP", vbBinaryCompare) = 0) Then
                For i = LBound(m_DB) To UBound(m_DB)
                    If ((Len(m_DB(i).Groups)) And (m_DB(i).Groups <> "%")) Then
                        Dim Splt() As String  ' ...
                        Dim j      As Integer ' ...
                    
                        If (InStr(1, m_DB(i).Groups, ",", vbTextCompare) <> 0) Then
                            Splt() = Split(m_DB(i).Groups, ",")
                        Else
                            ReDim Preserve Splt(0)
                            
                            Splt(0) = m_DB(i).Groups
                        End If
                        
                        For j = LBound(Splt) To UBound(Splt)
                            If (StrComp(Splt(j), trvUsers.SelectedItem.text, _
                                vbTextCompare) = 0) Then
                            
                                Splt(j) = NewString
                            End If
                        Next j
                        
                        m_DB(i).Groups = Join(Splt(), ",")
                    End If
                Next i
            End If
        End If
    End If
End Sub

Private Function Exists(ByVal nodeName As String) As Integer
    Dim i As Integer ' ...
    
    For i = 1 To trvUsers.Nodes.Count
        If (StrComp(trvUsers.Nodes(i).text, nodeName, _
            vbTextCompare) = 0) Then
            
            Exists = i
        
            Exit Function
        End If
    Next i
    
    Exists = False
End Function

Private Sub txtFlags_Change()
    cmdSave(1).Enabled = True
End Sub

Private Sub txtRank_Change()
    cmdSave(1).Enabled = True
End Sub

Private Function GetAccess(ByVal Username As String, Optional dbType As String = _
    vbNullString) As udtGetAccessResponse
    
    Dim i   As Integer ' ...
    Dim bln As Boolean ' ...

    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Username, Username, vbTextCompare) = 0) Then
            If (Len(dbType)) Then
                If (StrComp(m_DB(i).type, dbType, vbBinaryCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
                
            If (bln = True) Then
                With GetAccess
                    .Username = m_DB(i).Username
                    .Access = m_DB(i).Access
                    .flags = m_DB(i).flags
                    .AddedBy = m_DB(i).AddedBy
                    .AddedOn = m_DB(i).AddedOn
                    .ModifiedBy = m_DB(i).ModifiedBy
                    .ModifiedOn = m_DB(i).ModifiedOn
                    .type = m_DB(i).type
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

Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = _
    vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER

    Dim i     As Integer ' ...
    Dim found As Boolean ' ...
    
    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean ' ...
        
            If (Len(dbType)) Then
                If (StrComp(m_DB(i).type, dbType, vbBinaryCompare) = 0) Then
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
        If (StrComp(bak.type, "GROUP", vbBinaryCompare) = 0) Then
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
                            Dim Splt()     As String ' ...
                            Dim innerfound As Boolean ' ...
                            
                            Splt() = Split(m_DB(i).Groups, ",")
                            
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
                                
                                m_DB(i).Groups = Join(Splt(), vbNullString)
                            End If
                        Else
                            If (StrComp(bak.Username, m_DB(i).Groups, vbTextCompare) = 0) Then
                                res = DB_remove(m_DB(i).Username, m_DB(i).type)
                                
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
