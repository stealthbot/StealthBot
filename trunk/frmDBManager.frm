VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmDBManager 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
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
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton btnCreateGame 
      Caption         =   "Ga&me"
      Height          =   375
      Left            =   3240
      Picture         =   "frmDBManager.frx":0000
      TabIndex        =   23
      ToolTipText     =   "Create Group"
      Top             =   5378
      Width           =   735
   End
   Begin VB.CommandButton btnCreateClan 
      Caption         =   "C&lan"
      Height          =   375
      Left            =   2520
      Picture         =   "frmDBManager.frx":0468
      TabIndex        =   24
      ToolTipText     =   "Create Group"
      Top             =   5378
      Width           =   735
   End
   Begin MSComctlLib.ImageList Icons 
      Left            =   720
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0889
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0949
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":09B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":0D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":111C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":1526
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBManager.frx":17D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton btnCreateGroup 
      Caption         =   "&Group"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Create Group"
      Top             =   5378
      Width           =   735
   End
   Begin VB.CommandButton btnCreateUser 
      Caption         =   "Create &User..."
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      ToolTipText     =   "Create User"
      Top             =   5378
      Width           =   1695
   End
   Begin VB.CommandButton btnSaveForm 
      Caption         =   "Apply and Cl&ose"
      Default         =   -1  'True
      Height          =   300
      Left            =   5880
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5160
      TabIndex        =   3
      Top             =   5880
      Width           =   735
   End
   Begin vbalTreeViewLib6.vbalTreeView trvUsers 
      Height          =   4965
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8758
      BackColor       =   10040064
      ForeColor       =   16777215
      HotTracking     =   0   'False
      OLEDropMode     =   1
      OLEDragMode     =   1
      DragAutoExpand  =   -1
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
   Begin VB.Frame frmDatabase 
      BackColor       =   &H80000007&
      Caption         =   "Eric[nK]"
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   3025
      Begin VB.CommandButton btnSaveUser 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton btnRename 
         Caption         =   "&Rename"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox txtBanMessage 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   1320
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2328
         View            =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         SmallIcons      =   "Icons"
         ForeColor       =   16777215
         BackColor       =   10040064
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
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtRank 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblInherit 
         BackColor       =   &H00000000&
         Caption         =   "Inherits:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label lblBanMessage 
         BackColor       =   &H00000000&
         Caption         =   "Ban message:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label lblModifiedOn 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   360
         TabIndex        =   15
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblGroups 
         BackColor       =   &H00000000&
         Caption         =   "Groups:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblModifiedBy 
         BackColor       =   &H00000000&
         Caption         =   "(modified by)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   480
         TabIndex        =   17
         Top             =   2000
         Width           =   2415
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H00000000&
         Caption         =   "(created by)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   480
         TabIndex        =   16
         Top             =   1300
         Width           =   2415
      End
      Begin VB.Label lblFlags 
         BackColor       =   &H00000000&
         Caption         =   "Flags:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblRank 
         BackColor       =   &H00000000&
         Caption         =   "Rank (1 - 200):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCreatedOn 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   360
         TabIndex        =   12
         Top             =   1100
         Width           =   2415
      End
      Begin VB.Label lblCreated 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   240
         TabIndex        =   14
         Top             =   900
         Width           =   2535
      End
      Begin VB.Label lblLastMod 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   240
         TabIndex        =   13
         Top             =   1600
         Width           =   2535
      End
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00000000&
      Caption         =   "User Database"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSetPrimary 
         Caption         =   "Set This Primary"
         Enabled         =   0   'False
         Shortcut        =   ^P
         Visible         =   0   'False
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
Option Explicit

Public m_entrytype     As String
Public m_entryname     As String

' icons for groups list
Private Const IC_EMPTY     As Integer = 0
Private Const IC_UNCHECKED As Integer = 2
Private Const IC_CHECKED   As Integer = 3
Private Const IC_PRIMARY   As Integer = 4

' icons for database tree (for some reason it's 0-based!)
Private Const IC_UNKNOWN   As Integer = 0
Private Const IC_DATABASE  As Integer = 4
Private Const IC_USER      As Integer = 5
Private Const IC_GROUP     As Integer = 6
Private Const IC_CLAN      As Integer = 7
Private Const IC_GAME      As Integer = 8

' temporary DB working copy (TODO: USE Collection OF clsDBEntryObj!!)
Private m_DB()         As udtDatabase
' current entry index
Private m_currententry As Integer
' current entry node
Private m_currnode     As cTreeViewNode
' is this entry modified
Private m_modified     As Boolean
' is this a new entry (unused: if label editing worked, we'd use this)
Private m_new_entry    As Boolean
' root of DB ("Database")
Private m_root         As cTreeViewNode
' target for user node right-click menu
Private m_menutarget   As cTreeViewNode
' count of groups
Private m_glistcount   As Integer
' selected list item
Private m_glistsel     As ListItem
' target for group list right-click menu
Private m_gmenutarget  As ListItem

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ' this line is gay but for some reason I can't set the ImageList for vbalTV in the designer/VB properties -Ribose
    trvUsers.ImageList = Icons.hImageList
    
    ' has our database been loaded?
    If (DB(0).Username = vbNullString) Then
        ' load database, if for some reason that hasn't been done
        Call LoadDatabase
    End If
    
    ' store temporary copy of database
    m_DB() = DB()
    
    ' show database for default tab
    Call LoadView
End Sub ' end function Form_Load

Public Sub ImportDatabase(strPath As String, dbType As Integer)
    Dim f    As Integer
    Dim buf  As String
    Dim n    As cTreeViewNode
    Dim i    As Integer
    Dim tmp  As udtGetAccessResponse
    
    f = FreeFile

    If (dbType = 0) Then
    
        Open strPath For Input As #f
            Do While (EOF(f) = False)
                Line Input #f, buf
                
                If (buf <> vbNullString) Then
                    If (GetAccess(buf, tmp, "USER", i)) Then
                        ' do not save to tmp (it's a copy), but use m_DB(i)
                        With m_DB(i)
                            .Username = buf
                            .Type = "USER"
                            .ModifiedBy = "(console)"
                            .ModifiedOn = Now
                            
                            If (InStr(1, .Flags, "S", vbBinaryCompare) = 0) Then
                                .Flags = .Flags & "S"
                            End If
                            
                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
                                    .Groups = .Groups & "," & trvUsers.SelectedItem.Text
                                End If
                            End If
                            
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
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
                            
                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
                                    If (Not IsInGroup(m_DB(UBound(m_DB)).Groups, trvUsers.SelectedItem.Text)) Then
                                        .Groups = .Groups & "," & trvUsers.SelectedItem.Text
                                    End If
                                End If
                            End If
                            
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
                    End If
                End If
            Loop ' end loop
        Close #f
        
    ElseIf ((dbType = 1) Or (dbType = 2)) Then
        
        Dim User As String
        Dim Msg  As String
    
        Open strPath For Input As #f
            Do While (EOF(f) = False)
                Line Input #f, buf
                
                If (buf <> vbNullString) Then
                    If (Not InStr(1, buf, Space$(1), vbBinaryCompare) = 0) Then
                        User = Left$(buf, InStr(1, buf, Space$(1), vbBinaryCompare) - 1)
                        
                        Msg = Mid$(buf, Len(User) + 1)
                    Else
                        User = buf
                    End If
                
                    If (GetAccess(User, tmp, "USER", i)) Then
                        ' do not save to tmp (it's a copy), but use m_DB(i)
                        With m_DB(i)
                            .Username = User
                            .Type = "USER"
                            .ModifiedBy = "(console)"
                            .ModifiedOn = Now
                            
                            If (InStr(1, .Flags, "B", vbBinaryCompare) = 0) Then
                                .Flags = .Flags & "B"
                            End If
                            
                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
                                    If (Not IsInGroup(m_DB(i).Groups, trvUsers.SelectedItem.Text)) Then
                                        .Groups = .Groups & "," & trvUsers.SelectedItem.Text
                                    End If
                                End If
                            End If
                            
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
                    Else
                        ' redefine array to support new entry
                        ReDim Preserve m_DB(UBound(m_DB) + 1)
                        
                        ' create new database entry
                        With m_DB(UBound(m_DB))
                            .Username = User
                            .Type = "USER"
                            .AddedBy = "(console)"
                            .AddedOn = Now
                            .ModifiedBy = "(console)"
                            .ModifiedOn = Now
                            .Flags = "B"
                            .BanMessage = Msg
                            
                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
                                    .Groups = trvUsers.SelectedItem.Text
                                End If
                            End If
                            
                            If (.Groups = vbNullString) Then
                                .Groups = "%"
                            End If
                        End With
                    End If
                End If
            Loop ' end loop
        Close #f
        
    End If
    
    LoadView
End Sub

Private Sub btnCreateUser_Click()
    Dim newNode   As cTreeViewNode
    Dim Username  As String
    Dim tmp       As udtGetAccessResponse
    Dim pos       As Integer
    
    m_entrytype = "USER"
    m_entryname = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, frmDBManager)
    
    If (LenB(m_entryname) > 0) Then
    
        Username = m_entryname
    
        If (Not GetAccess(Username, tmp, "USER")) Then
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
            
            Set newNode = PlaceNewNode(Username, "USER", IC_USER)
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                With newNode
                    '.Image = 0
                    .Tag = "USER"
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already an entry of this type matching " & _
                "the specified name."
            pos = FindNodeIndex(Username, "USER")
            If pos > 0 Then
                trvUsers.Nodes(pos).Selected = True
            End If
        End If
    End If
End Sub

Private Sub btnCreateGroup_Click()
    Dim newNode   As cTreeViewNode
    Dim GroupName As String
    Dim tmp       As udtGetAccessResponse
    Dim pos       As Integer

    m_entrytype = "GROUP"
    m_entryname = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, frmDBManager)
    
    If (LenB(m_entryname) > 0) Then
    
        GroupName = m_entryname
    
        If (Not GetAccess(GroupName, tmp, "GROUP")) Then
            ReDim Preserve m_DB(UBound(m_DB) + 1)
            
            With m_DB(UBound(m_DB))
                .Username = GroupName
                .Type = "GROUP"
                .AddedBy = "(console)"
                .AddedOn = Now
                .ModifiedBy = "(console)"
                .ModifiedOn = Now
            End With
            
            Set newNode = PlaceNewNode(GroupName, "GROUP", IC_GROUP)
            
            Call UpdateGroupList
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                With newNode
                    .Tag = "GROUP"
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already an entry of this type matching " & _
                "the specified name."
            pos = FindNodeIndex(GroupName, "GROUP")
            If pos > 0 Then
                trvUsers.Nodes(pos).Selected = True
            End If
        End If
    End If
End Sub

Sub btnCreateClan_Click()
    Dim newNode   As cTreeViewNode
    Dim ClanName  As String
    Dim tmp       As udtGetAccessResponse
    Dim pos       As Integer
    
    m_entrytype = "CLAN"
    m_entryname = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, frmDBManager)
    
    If (LenB(m_entryname) > 0) Then
    
        ClanName = m_entryname
    
        If (Not GetAccess(ClanName, tmp, "CLAN")) Then
            ReDim Preserve m_DB(UBound(m_DB) + 1)
            
            With m_DB(UBound(m_DB))
                .Username = ClanName
                .Type = "CLAN"
                .AddedBy = "(console)"
                .AddedOn = Now
                .ModifiedBy = "(console)"
                .ModifiedOn = Now
            End With
            
            Set newNode = PlaceNewNode(ClanName, "CLAN", IC_CLAN)
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                With newNode
                    .Tag = "CLAN"
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already an entry of this type matching " & _
                "the specified name."
            pos = FindNodeIndex(ClanName, "CLAN")
            If pos > 0 Then
                trvUsers.Nodes(pos).Selected = True
            End If
        End If
    End If
End Sub
        
Sub btnCreateGame_Click()
    Dim newNode   As cTreeViewNode
    Dim GameEntry As String
    Dim tmp       As udtGetAccessResponse
    Dim pos       As Integer
    
    m_entryname = vbNullString
    
    Call frmDBGameSelection.Show(vbModal, frmDBManager)
    
    If (LenB(m_entryname) > 0) Then
    
        GameEntry = m_entryname
        
        If (Not GetAccess(GameEntry, tmp, "GAME")) Then
            ReDim Preserve m_DB(UBound(m_DB) + 1)
            
            With m_DB(UBound(m_DB))
                .Username = GameEntry
                .Type = "GAME"
                .AddedBy = "(console)"
                .AddedOn = Now
                .ModifiedBy = "(console)"
                .ModifiedOn = Now
            End With
        
            Set newNode = PlaceNewNode(GameEntry, "GAME", IC_GAME)
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                With newNode
                    .Tag = "GAME"
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already an entry of this type matching " & _
                "the specified name."
            pos = FindNodeIndex(GameEntry, "GAME")
            If pos > 0 Then
                trvUsers.Nodes(pos).Selected = True
            End If
        End If
    End If
End Sub

Private Function PlaceNewNode(EntryName As String, EntryType As String, EntryImage As Integer) As cTreeViewNode
    Dim NewParent As cTreeViewNode
    
    ' by default create the node under the root node
    Set NewParent = m_root
    
    ' do we have an item (hopefully a group) selected?
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' is the item a group?
        If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
            ' create new node under group node
            Set NewParent = trvUsers.SelectedItem
        Else
            ' is our parent a group?
            If Not trvUsers.SelectedItem.Parent Is Nothing Then
                If (StrComp(trvUsers.SelectedItem.Parent.Tag, "GROUP", vbTextCompare) = 0) Then
                    Set NewParent = trvUsers.SelectedItem.Parent
                End If
            End If
        End If
    End If
    
    Set PlaceNewNode = trvUsers.Nodes.Add(NewParent, etvwChild, EntryType & ": " & EntryName, EntryName, EntryImage, EntryImage)
                    
    ' set group settings on new database entry
    If (Not PlaceNewNode Is Nothing) Then
        If (StrComp(PlaceNewNode.Parent.Tag, "GROUP", vbTextCompare) = 0) Then
            With m_DB(UBound(m_DB))
                .Groups = PlaceNewNode.Parent.Text
            End With
        End If
    End If
End Function

Private Sub btnCancel_Click()
    m_modified = False
    
    Unload Me
End Sub

Private Sub btnSaveUser_Click()
    Dim i         As Integer
    Dim j         As Integer
    Dim OldGroups As String
    Dim NewGroups As String
    Dim OldPGroup As String
    Dim NewPGroup As String
    Dim pos       As Integer
    Dim NewParent As cTreeViewNode
    Dim Node      As cTreeViewNode

    ' if we have no selected user... escape quick!
    If (trvUsers.SelectedItem Is Nothing) Then
        ' break from function
        Exit Sub
    End If
    
    ' can't "save" the "Database"/root node
    If (StrComp(trvUsers.SelectedItem.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    ' disable entry save command
    Call HandleSaved

    ' look for selected user in database
    For i = LBound(m_DB) To UBound(m_DB)
        ' is this the user we were looking for?
        If (StrComp(trvUsers.SelectedItem.Text, m_DB(i).Username, vbTextCompare) = 0) Then
            If (StrComp(trvUsers.SelectedItem.Tag, m_DB(i).Type, vbTextCompare) = 0) Then
                ' modifiy user data
                With m_DB(i)
                    .Rank = Val(txtRank.Text)
                    .Flags = txtFlags.Text
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                    .BanMessage = txtBanMessage.Text
                    
                    ' save old groups...
                    OldGroups = .Groups
                    
                    ' generate new groups string
                    NewGroups = vbNullString
                    
                    If m_glistcount > 0 Then
                        For j = 1 To lvGroups.ListItems.Count
                            With lvGroups.ListItems(j)
                                If .Checked And Not .Ghosted Then
                                    If .ForeColor = vbYellow Then
                                        ' place first
                                        NewGroups = .Text & "," & NewGroups
                                    Else
                                        ' append
                                        NewGroups = NewGroups & .Text & ","
                                    End If
                                End If
                            End With
                        Next j
                    End If
                    
                    ' if ends with ",", trim it
                    If (Len(NewGroups) > 1) Then
                        NewGroups = Left$(NewGroups, Len(NewGroups) - 1)
                    End If
                    
                    ' store it
                    .Groups = NewGroups
                    If .Groups = vbNullString Then
                        .Groups = "%"
                    End If
                    
                    ' now to check if we need to move this node!
                    ' did the "primary" group change?
                    OldPGroup = GetPrimaryGroup(OldGroups)
                    NewPGroup = GetPrimaryGroup(NewGroups)
                    
                    If (StrComp(OldPGroup, NewPGroup, vbTextCompare) <> 0) Then
                        ' move under new primary
                        Set Node = trvUsers.SelectedItem
                        pos = FindNodeIndex(NewPGroup, "GROUP", Node)
                        ' well, does it exist?
                        If (pos > 0) Then
                            ' make node a child of existing group
                            Set NewParent = trvUsers.Nodes(pos)
                        Else
                            ' put it under DB root
                            Set NewParent = m_root
                        End If
                        
                        ' move node!!
                        Call Node.MoveNode(NewParent, etvwChild)
                        Node.Tag = .Type
                        Node.Selected = True
                    End If
                End With
                
                Exit For
            End If
        End If
    Next i
End Sub

Private Sub HandleSaved()
    m_modified = False
    btnSaveUser.Enabled = False
    If m_currententry < 0 Then
        frmDatabase.Caption = "Database"
        frmDBManager.Caption = "Database"
    Else
        frmDatabase.Caption = m_DB(m_currententry).Username & " (" & LCase$(m_DB(m_currententry).Type) & ")"
        frmDBManager.Caption = "Database - " & m_DB(m_currententry).Username & " (" & LCase$(m_DB(m_currententry).Type) & ")"
    End If
End Sub

Private Sub HandleUnsaved()
    If m_currententry < 0 Then
        m_modified = False
        frmDatabase.Caption = "Database"
        frmDBManager.Caption = "Database"
    Else
        m_modified = True
        btnSaveUser.Enabled = True
        frmDatabase.Caption = m_DB(m_currententry).Username & " (" & m_DB(m_currententry).Type & ")*"
        frmDBManager.Caption = "Database - " & m_DB(m_currententry).Username & " (" & m_DB(m_currententry).Type & ")*"
    End If
End Sub

Private Sub btnSaveForm_Click()
    ' save this user first
    If m_modified Then
        btnSaveUser_Click
    End If
    
    ' write temporary database to official
    DB() = m_DB()
    
    ' save database
    Call WriteDatabase(GetFilePath(FILE_USERDB))
    
    ' check channel to find potential banned users
    Call g_Channel.CheckUsers
    
    ' close database form
    Call Unload(frmDBManager)
End Sub

Private Sub lvGroups_Click()
    Set m_glistsel = lvGroups.SelectedItem
    
    If Not m_glistsel Is Nothing Then
        m_glistsel.Checked = Not m_glistsel.Checked
        Call lvGroups_ItemCheck(m_glistsel)
    End If
End Sub

Private Sub lvGroups_DblClick()
    Dim i As Integer
    Dim Item As ListItem
    
    Set Item = lvGroups.SelectedItem
    
    If Item.Ghosted Then
        Exit Sub
    End If
    
    ' set primary
    If Item.Checked Then
        Call SetLVPrimaryGroup(Item)
    End If
    
    ' enable entry save command
    Call HandleUnsaved
End Sub

Private Sub lvGroups_ItemCheck(ByVal Item As ListItem)
    Dim i As Integer
    Dim NewGroups As String

    Item.Selected = True

    If Item.Ghosted Then
        Item.Checked = False
        Item.SmallIcon = IIf(m_glistcount = 0, IC_EMPTY, IC_UNCHECKED)
        Exit Sub
    End If
    
    If Item.Checked Then
        ' if checked
        ' if no primary group
        If GetLVPrimaryGroup() Is Nothing Then
            ' set this
            Item.ForeColor = vbYellow
            Item.SmallIcon = IC_PRIMARY
        Else
            Item.SmallIcon = IC_CHECKED
        End If
    Else
        ' if not checked
        ' select the first "checked" item to be new primary
        If Item.ForeColor = vbYellow Then
            ' unset this
            Item.ForeColor = vbWhite
            For i = 1 To lvGroups.ListItems.Count
                If lvGroups.ListItems.Item(i).Checked Then
                    ' set if found
                    With lvGroups.ListItems.Item(i)
                        .ForeColor = vbYellow
                        .SmallIcon = IC_PRIMARY
                    End With
                    
                    Exit For
                End If
            Next i
        End If
        Item.SmallIcon = IC_UNCHECKED
    End If
    
    ' generate new groups string
    NewGroups = vbNullString
    
    For i = 1 To lvGroups.ListItems.Count
        With lvGroups.ListItems(i)
            If .Checked And Not .Ghosted Then
                If .ForeColor = vbYellow Then
                    ' place first
                    NewGroups = .Text & "," & NewGroups
                Else
                    ' append
                    NewGroups = NewGroups & .Text & ","
                End If
            End If
        End With
    Next i
    
    ' if ends with ",", trim it
    If (Len(NewGroups) > 1) Then
        NewGroups = Left$(NewGroups, Len(NewGroups) - 1)
    End If
    
    ' update inherits list
    Call UpdateInheritCaption(NewGroups)
    
    ' enable entry save command
    Call HandleUnsaved
End Sub

Private Sub lvGroups_KeyPress(KeyAscii As Integer)
    If m_glistsel <> lvGroups.SelectedItem Then
        Set m_glistsel = lvGroups.SelectedItem
    End If
    
    If KeyAscii = vbKeySpace Then
        If Not m_glistsel Is Nothing Then
            m_glistsel.Checked = Not m_glistsel.Checked
            Call lvGroups_ItemCheck(m_glistsel)
        End If
    End If
End Sub

Private Sub lvGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuSetPrimary.Visible = True
    mnuRename.Visible = False
    mnuDelete.Visible = False
    
    mnuSetPrimary.Enabled = False
    
    Set m_gmenutarget = Nothing
    
    If (Button = vbRightButton) Then
        Set m_gmenutarget = lvGroups.HitTest(x, y)
        
        If (Not m_gmenutarget Is Nothing) Then
            If (Not m_gmenutarget.Ghosted) Then
                mnuSetPrimary.Enabled = (m_gmenutarget.ForeColor <> vbYellow)
            
                Call Me.PopupMenu(mnuContext)
            End If
        End If
    End If
End Sub

Private Sub mnuSetPrimary_Click()
    Call SetLVPrimaryGroup(m_gmenutarget)
    
    Call HandleUnsaved
End Sub

Private Sub mnuDelete_Click()
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Call HandleDeleteEvent(m_menutarget)
    End If
End Sub

Private Sub btnDelete_Click()
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Call HandleDeleteEvent(trvUsers.SelectedItem)
    End If
End Sub

Private Sub mnuRename_Click()
    Call QueryRenameEvent(m_menutarget)
End Sub

Private Sub btnRename_Click()
    Call QueryRenameEvent(trvUsers.SelectedItem)
End Sub

'Private Sub mnuOpenDatabase_Click()
'    ' open file dialog
'    Call CommonDialog.ShowOpen
'End Sub

Private Sub HandleDeleteEvent(Target As cTreeViewNode)
    If (Target Is Nothing) Then
        Exit Sub
    End If
    
    If (StrComp(Target.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    Dim response As VbMsgBoxResult
    Dim isGroup  As Boolean
    
    If (m_modified And StrComp(Target.Text, m_currnode.Text, vbBinaryCompare) = 0) Then
        ' do not ask about "unsaved data" later!
        m_modified = False
    End If
    
    isGroup = (StrComp(Target.Tag, "GROUP", vbTextCompare) = 0)

    If (isGroup) Then
        response = MsgBox("Are you sure you wish to delete this group and " & _
            "all of its members?", vbYesNo Or vbInformation, "Database - Confirm Delete")
    End If
    
    If ((isGroup = False) Or ((isGroup) And (response = vbYes))) Then
        Call DB_remove(Target.Text, Target.Tag)
        
        Target.Delete
        
        If isGroup Then
            Call UpdateGroupList
        End If
        
        'Call trvUsers_NodeClick(trvUsers.SelectedItem)
    End If
End Sub

Private Sub QueryRenameEvent(Target As cTreeViewNode)
    If (Not (Target Is Nothing)) Then
        ' only "GROUP" entries can be renamed
        If (StrComp(Target.Tag, "GROUP", vbTextCompare) = 0) Then
            'trvUsers.SelectedItem.StartEdit
            
            m_entrytype = Target.Tag
            m_entryname = Target.Text
            
            Call frmDBNameEntry.Show(vbModal, frmDBManager)
            
            If HandleRenameEvent(Target, m_entryname) Then
            
                If LenB(m_entryname) > 0 Then
                    Target.Text = m_entryname
                    trvUsers.Refresh
                    
                    Call UpdateGroupList
                    
                    If m_modified Then
                        Call HandleUnsaved
                    Else
                        Call HandleSaved
                    End If
                End If
            
            Else
                ' alert user that entry already exists
                MsgBox "There is already an entry of this type matching " & _
                    "the specified name."
            End If
        End If
    End If
End Sub

Private Function HandleRenameEvent(Target As cTreeViewNode, NewString As String) As Boolean
    Dim i           As Integer
    Dim WasUpdated  As Boolean
    Dim tmp         As udtGetAccessResponse

    HandleRenameEvent = True
    WasUpdated = False
    
    If (Target Is Nothing) Then
        Exit Function
    End If
    
    If (StrComp(Target.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Function
    End If
    
    If NewString = Target.Text Or LenB(NewString) = 0 Then
        ' same name succeeds (no chnage); empty name success (cancelled)
        Exit Function
    End If
    
    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(Target.Text, m_DB(i).Username, vbTextCompare) = 0) Then
            If (StrComp(Target.Tag, m_DB(i).Type, vbTextCompare) = 0) Then
                If (GetAccess(NewString, tmp, m_DB(i).Type)) Then
                    ' already exists
                    HandleRenameEvent = False
                    Exit Function
                End If
                
                ' rename DB entry
                m_DB(i).Username = NewString
                WasUpdated = True
        
                Exit For
            End If
        End If
    Next i
    
    If Not WasUpdated Then
        ' this didn't already exist!! shouldn't happen...
        HandleRenameEvent = False
        Exit Function
    End If
    
    If (StrComp(m_DB(i).Type, "GROUP", vbTextCompare) = 0) Then
        For i = LBound(m_DB) To UBound(m_DB)
            If (Len(m_DB(i).Groups) > 0) Then
                Dim Splt() As String
                Dim j      As Integer
            
                If (Not InStr(1, m_DB(i).Groups, ",", vbTextCompare) = 0) Then
                    Splt() = Split(m_DB(i).Groups, ",")
                Else
                    ReDim Preserve Splt(0)
                    
                    Splt(0) = m_DB(i).Groups
                End If
                
                For j = LBound(Splt) To UBound(Splt)
                    If (StrComp(Splt(j), "%", vbBinaryCompare) <> 0) And (StrComp(Splt(j), Target.Text, vbTextCompare) = 0) Then
                        Splt(j) = NewString
                    End If
                Next j
                
                m_DB(i).Groups = Join(Splt(), ",")
            End If
        Next i
    End If
End Function

' handle tab clicks and initial loading
Private Sub LoadView()
    On Error GoTo ERROR_HANDLER

    Dim newNode As cTreeViewNode

    Dim i                 As Integer
    Dim j                 As Integer
    Dim K                 As Integer
    Dim grp               As String
    Dim pos               As Integer
    Dim bln               As Boolean
    Dim blnDuplicateFound As Boolean
    Dim TypeName          As String
    Dim TypeImage         As Integer
    Dim NewParent         As cTreeViewNode

    ' clear treeview
    Call trvUsers.Nodes.Clear
    
    ' create root node
    Set m_root = trvUsers.Nodes.Add(, , "Database", "Database", IC_DATABASE, IC_DATABASE)
    ' type DATABASE
    m_root.Tag = "DATABASE"

    ' which tab index are we on?
    For i = LBound(m_DB) To UBound(m_DB)
        ' we're handling groups first; is this entry a group?
        If (StrComp(m_DB(i).Type, "GROUP", vbBinaryCompare) = 0) And (Len(m_DB(i).Username) > 0) Then
            ' is this group a member of other groups?
            If (Len(m_DB(i).Groups) > 0) And (StrComp(m_DB(i).Groups, "%", vbBinaryCompare) <> 0) Then
                ' get the "primary" group (the first group) to put the node under
                grp = GetPrimaryGroup(m_DB(i).Groups)
            
                ' has the group already been added or is database in an
                ' incorrect order?
                pos = FindNodeIndex(grp, "GROUP")
                ' well, does it exist?
                If (pos > 0) Then
                    ' make node a child of existing group
                    Set NewParent = trvUsers.Nodes(pos)
                Else
                    ' lets make this guy a parent node for now until we can find
                    ' his real parent.
                    Set NewParent = m_root
                End If
                
                Set newNode = trvUsers.Nodes.Add(NewParent, etvwChild, "GROUP: " & m_DB(i).Username, m_DB(i).Username, IC_GROUP, IC_GROUP)
            Else
                ' create node
                Set NewParent = m_root
                
                Set newNode = trvUsers.Nodes.Add(NewParent, etvwChild, "GROUP: " & m_DB(i).Username, m_DB(i).Username, IC_GROUP, IC_GROUP)
                
                If (Not (newNode Is Nothing)) Then
                    ' Okay, is the group a lone ranger?  Or does he have children
                    ' that are already in the list?
                    j = LBound(m_DB)
                    Do
                        For j = j To (i - 1)
                            ' we're only concerned with groups, atm.
                            If (StrComp(m_DB(j).Type, "GROUP", vbBinaryCompare) = 0) And (Len(m_DB(j).Username) > 0) Then
                                ' we only need to check for groups that are members of
                                ' other groups
                                If (Len(m_DB(j).Groups) > 0) And (StrComp(m_DB(j).Groups, "%", vbBinaryCompare) <> 0) Then
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
                    
                        ' is this node a baby's daddy?
                        If (bln) Then
                            ' move node
                            pos = FindNodeIndex(m_DB(j).Username, "GROUP", newNode)
                            
                            If pos > 0 Then
                                Set trvUsers.Nodes(pos).Parent = newNode
                            End If
                        End If
        
                        ' reset boolean
                        bln = False
                    Loop Until j = i
                End If
            End If
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                newNode.Tag = "GROUP"
            End If
        End If
    Next i

    ' loop through database... this time looking for users, clans, and games (clans and games are "user like" in the tree)
    For i = LBound(m_DB) To UBound(m_DB)
        ' is the entry a user?
        If (StrComp(m_DB(i).Type, "GROUP", vbTextCompare) <> 0) And (Len(m_DB(i).Username) > 0) Then
            ' find the type name, used for the treeview
            If (StrComp(m_DB(i).Type, "USER", vbTextCompare) = 0) Then
                TypeName = "USER"
                TypeImage = IC_USER
            ElseIf (StrComp(m_DB(i).Type, "CLAN", vbTextCompare) = 0) Then
                TypeName = "CLAN"
                TypeImage = IC_CLAN
            ElseIf (StrComp(m_DB(i).Type, "GAME", vbTextCompare) = 0) Then
                TypeName = "GAME"
                TypeImage = IC_GAME
            Else
                TypeName = "USER"
                TypeImage = IC_UNKNOWN
            End If
            
            ' is the user a member of any groups?
            If (Len(m_DB(i).Groups) > 0) And (StrComp(m_DB(i).Groups, "%", vbBinaryCompare) <> 0) Then
                ' get the "primary" group (the first group) to put the node under
                grp = GetPrimaryGroup(m_DB(i).Groups)
            
                If (grp = vbNullString) Then
                    pos = 0
                Else
                    ' search for our group
                    pos = FindNodeIndex(grp, "GROUP")
                        
                    ' does our group exist?
                    If (pos > 0) Then
                        Set NewParent = trvUsers.Nodes(pos)
                    End If
                End If
            Else
                pos = 0
            End If
            
            If (pos <= 0) Then
                ' create new user node under root
                Set NewParent = m_root
            End If
            
            ' create user node and move into group
            Set newNode = trvUsers.Nodes.Add(NewParent, etvwChild, TypeName & ": " & m_DB(i).Username, m_DB(i).Username, TypeImage, TypeImage)
            
            If (Not (newNode Is Nothing)) Then
                ' change misc. settings
                newNode.Tag = TypeName
            End If
        End If
        
        ' reset our variables
        pos = 0
    Next i
    
    ' does our treeview contain any nodes?
    For i = 1 To trvUsers.NodeCount
        trvUsers.Nodes(i).Expanded = True
    Next i
    
    Call UpdateGroupList
    
    If trvUsers.NodeCount = 0 Then
        Call LockGUI
    'Else
    '    Call trvUsers_NodeClick(trvUsers.nodes(1))
    End If
    
    If (blnDuplicateFound = True) Then
        MsgBox "There were one or more duplicate database entries found which could not be loaded.", _
            vbExclamation, "Error"
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    ' duplicate node
    If (Err.Number = 35602) Then
        DB_remove m_DB(i).Username, m_DB(i).Type
        blnDuplicateFound = True
        Resume Next
    End If

    Exit Sub
End Sub

Private Sub LockGUI()
    Dim i As Integer
    
    ' set our default frame caption
    m_currententry = -1
    Call HandleSaved

    ' disable & clear rank
    txtRank.Enabled = False
    txtRank.Text = vbNullString
    
    ' disable & clear flags
    txtFlags.Enabled = False
    txtFlags.Text = vbNullString
    
    ' loop through listbox and clear selected items
    Call ClearGroupListChecks
    
    ' disable group lists
    'lvGroups.Enabled = False
    
    ' disable & clear ban message
    txtBanMessage.Enabled = False
    txtBanMessage.Text = vbNullString
    
    ' reset created on & modified on labels
    lblCreatedOn.Caption = "(not applicable)"
    lblModifiedOn.Caption = "(not applicable)"
    
    ' reset created by & modified by labels
    lblCreatedBy.Caption = vbNullString
    lblModifiedBy.Caption = vbNullString
    
    ' reset inherits caption
    lblInherit.Caption = vbNullString
    
    ' disable entry buttons
    btnRename.Enabled = False
    btnDelete.Enabled = False
End Sub

Private Sub UnlockGUI()
    Dim i As Integer

    ' enable rank field
    txtRank.Enabled = True

    ' enable flags field
    txtFlags.Enabled = True
    
    ' enable ban message field
    txtBanMessage.Enabled = True
    
    ' enable entry rename/delete buttons
    btnRename.Enabled = (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0)
    btnDelete.Enabled = True
    
    ' enable group lists
    'lvGroups.Enabled = True
    
    ' make sure save button and caption is up to date
    HandleSaved
End Sub

' handle node collapse
Private Sub trvUsers_Collapse(Node As cTreeViewNode)
    ' refresh tree view
    Call trvUsers.Refresh
End Sub

' handle node expand
Private Sub trvUsers_Expand(Node As cTreeViewNode)
    ' refresh tree view
    Call trvUsers.Refresh
End Sub

Private Sub trvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) Then
        If (Not (trvUsers.SelectedItem Is Nothing)) Then
            Call HandleDeleteEvent(trvUsers.SelectedItem)
        End If
    ElseIf (KeyCode = vbKeyF2) Then
        Call QueryRenameEvent(trvUsers.SelectedItem)
    End If
End Sub

'Private Sub trvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim node As cTreeViewNode
'
'    If (Button = vbLeftButton) Then
'        Set m_nodedragsrc = trvUsers.HitTest(x, y)
'        'frmChat.AddChat vbYellow, "[MOUSE] DOWN DRAG=" & m_nodedragsrc.Text
'    End If
'End Sub
'
'Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    'frmChat.AddChat vbYellow, "[MOUSE] MOVE" ' DRAG=" & m_nodedragsrc.Text
'    If (Button = vbLeftButton And m_nodedrag And Not m_nodedragsrc Is Nothing) Then
'        trvUsers_OLEDragDrop Nothing, vbDropEffectMove, Button, Shift, x, y 'm_dragtarget
'    End If
'End Sub

'Private Sub trvUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    'frmChat.AddChat vbYellow, "[MOUSE] UP" ' DRAG=" & m_nodedragsrc.Text
'End Sub

Private Sub trvUsers_NodeRightClick(Node As cTreeViewNode)
    mnuRename.Visible = True
    mnuDelete.Visible = True
    mnuSetPrimary.Visible = False
    
    mnuRename.Enabled = False
    mnuDelete.Enabled = False
    
    Set m_menutarget = Nothing
    
    If (Not (Node Is Nothing)) Then
        If (StrComp(Node.Tag, "DATABASE", vbTextCompare) <> 0) Then
            mnuRename.Enabled = (StrComp(Node.Tag, "GROUP", vbTextCompare) = 0)
            mnuDelete.Enabled = True
            
            Set m_menutarget = Node
        End If
    End If
            
    Call Me.PopupMenu(mnuContext)
End Sub

'Private Sub trvUsers_OLECompleteDrag(Effect As Long)
'    'frmChat.AddChat vbYellow, "[DRAG] COMPLETE: E=" & Effect
'End Sub
'
''// occurs when the user drops the object
''// this is where you move the node and its children.
''// this will not occur if Effect = vbDropEffectNone
'Private Sub trvUsers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim strSourceKey As String
'    Dim dragNode   As cTreeViewNode
'    Dim dropNode   As cTreeViewNode
'
'    '// get the carried data
'    'strSourceKey = Data.GetData(vbCFText)
'    Set dragNode = m_nodedragsrc
'    If dragNode Is Nothing Then
'        Exit Sub
'    End If
'    Set dropNode = trvUsers.HitTest(x, y)
'    If dropNode Is Nothing Then
'        Set dropNode = m_root
'    End If
'    '// get the target node
'    frmChat.AddChat vbYellow, "[DRAG] DROP: DROP=" & dropNode.Text & " E=" & Effect
'    '// if the target node is not a folder or the root item
'    '// then get it's parent (that is a folder or the root item)
'    If (StrComp(dropNode.Tag, "GROUP", vbTextCompare) <> 0) And (StrComp(dropNode.Tag, "DATABASE", vbTextCompare) <> 0) Then
'        '// the target must be a GROUP or the DATABASE
'        Effect = vbDropEffectNone
'        Exit Sub
'    End If
'
'    'Set dragNode = trvUsers.nodes(strSourceKey)
'    If Not dragNode Is Nothing Then
'        frmChat.AddChat vbYellow, "[DRAG] DROP: DRAG=" & dragNode.Text & " E=" & Effect
'
'        '// move the source node to the target node
'        Call dragNode.MoveNode(dropNode, etvwChild)
'        dragNode.Tag = "USER"
'        dragNode.Selected = True
'    End If
'    '// NOTE: You will also need to update the key to reflect the changes
'    '// if you are using it
'    '// we are not dragging from this control any more
'    m_nodedrag = False
'    '// cancel effect so that VB doesn't muck up your transfer
'    Effect = 0
'End Sub
'
''// occurs when the user starts dragging
''// this is where you assign the effect and the data.
'Private Sub trvUsers_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'    'Dim K As String
'    AllowedEffects = vbDropEffectNone
'
'    If (Not m_nodedragsrc Is Nothing) Then
'        If (StrComp(m_nodedragsrc.Tag, "DATABASE", vbTextCompare) <> 0) Then
'            '// Set the effect to move
'            AllowedEffects = vbDropEffectMove
'            '// Assign the selected item's key to the DataObject
'            'K = m_nodedragsrc.Key
'            'Call Data.SetData(m_nodedragsrc.Key)
'            '// we are dragging from this control
'            m_nodedrag = True
'        End If
'    End If
'
'    'frmChat.AddChat vbYellow, "[DRAG] START: AE=" & AllowedEffects & ", D=" & K
'End Sub
'
''// occurs when the object is dragged over the control.
''// this is where you check to see if the mouse is over
''// a valid drop object
'Private Sub trvUsers_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'    Dim dropNode As cTreeViewNode
'
'    '// set the effect
'    Effect = vbDropEffectMove
'    '// get the node that the object is being dragged over
'    Set dropNode = trvUsers.HitTest(x, y)
'
'    'Set m_nodedragdata = Data
'
'    If (Not dropNode Is Nothing And m_nodedrag) Then
'        If (StrComp(dropNode.Tag, "GROUP", vbTextCompare) = 0) Or (StrComp(dropNode.Tag, "DATABASE", vbTextCompare) = 0) Then
'            'dropNode.DropHighlighted = True
'        Else
'            '// the target must be a GROUP or the DATABASE
'            Effect = vbDropEffectNone
'            Exit Sub
'        End If
'    Else
'        'm_root.DropHighlighted = True
'        Exit Sub
'    End If
'
'    frmChat.AddChat vbYellow, "[DRAG] OVER: DROP=" & dropNode.Text & " E=" & Effect & " S=" & State
'End Sub

Private Sub trvUsers_SelectedNodeChanged()
Static skipupdate As Boolean
    Dim Node     As cTreeViewNode
    Dim tmp      As udtGetAccessResponse
    Dim Splt()   As String
    Dim j        As Integer
    Dim i        As Integer
    Dim pos      As Integer
    Dim response As VbMsgBoxResult
    Dim Disable  As Boolean
    
    Set Node = trvUsers.SelectedItem
    
    If (Node Is Nothing) Then
        Exit Sub
    End If
    
    If m_modified And Not skipupdate Then
        ' check if we should allow this node change? (is Unsaved?)
        response = MsgBox("Are you sure you wish to discard changes to the " & _
            m_currnode.Text & " (" & UCase$(m_currnode.Tag) & ") database entry?", _
            vbYesNo Or vbInformation, "Database - Confirm Discard Changes")
            
        If response = vbNo Then
            skipupdate = True
            m_currnode.Selected = True
            skipupdate = False
            Exit Sub
        End If
    End If
    
    If skipupdate Then Exit Sub
    
    Set m_currnode = Node
    
    Call LockGUI
    
    Node.Expanded = True
    
    If (StrComp(Node.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    Call GetAccess(Node.Text, tmp, Node.Tag, m_currententry)
    
    ' does entry have a rank?
    If (tmp.Rank > 0) Then
        ' write rank to text box
        txtRank.Text = tmp.Rank
    Else
        ' clear rank from text box
        txtRank.Text = vbNullString
    End If
    
    ' clear flags from text box
    txtFlags.Text = tmp.Flags
    
    If ((tmp.AddedBy = vbNullString) Or (tmp.AddedBy = "%")) Then
        lblCreatedOn = "unknown"
        lblCreatedBy = "by unknown"
    Else
        lblCreatedOn = tmp.AddedOn & " Local Time"
        lblCreatedBy = "by " & tmp.AddedBy
    End If
    
    If ((tmp.ModifiedBy = vbNullString) Or (tmp.ModifiedBy = "%")) Then
        lblModifiedOn = "unknown"
        lblModifiedBy = "by unknown"
    Else
        lblModifiedOn = tmp.ModifiedOn & " Local Time"
        lblModifiedBy = "by " & tmp.ModifiedBy
    End If
    
    ' is entry a member of a group?
    If (Len(tmp.Groups) > 0) And (StrComp(tmp.Groups, "%", vbBinaryCompare) <> 0) Then
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
    End If
    
    Call UpdateInheritCaption(tmp.Groups)
    
    ' loop through our listview, checking for matches
    If m_glistcount > 0 Then
        For j = 1 To lvGroups.ListItems.Count
            With lvGroups.ListItems(j)
                ' loop through entry's group memberships
                .Checked = False
                .SmallIcon = IC_UNCHECKED
                .Ghosted = False
                .ForeColor = vbWhite
                
                If (Len(tmp.Groups) > 0) And (StrComp(tmp.Groups, "%", vbBinaryCompare) <> 0) Then
                    For i = LBound(Splt) To UBound(Splt)
                        If (StrComp(Splt(i), "%", vbBinaryCompare) <> 0) And (StrComp(Splt(i), .Text, vbTextCompare) = 0) Then
                            ' select group if entry is a member
                            .Checked = True
                            .SmallIcon = IC_CHECKED
                            
                            ' highlight group if "primary" (first group)
                            If (i = LBound(Splt)) Then
                                .ForeColor = vbYellow
                                .SmallIcon = IC_PRIMARY
                            End If
                            
                            Exit For
                        End If
                    Next i
                End If
                
                If (StrComp(tmp.Type, "GROUP", vbTextCompare) = 0) Then
                    Disable = False
                    
                    ' don't allow a group to be in itself
                    If (StrComp(tmp.Username, .Text, vbTextCompare) = 0) Then
                        Disable = True
                    End If
                    
                    ' don't allow a group to be in its children
                    pos = FindNodeIndex(.Text)
                    If pos > 0 Then
                        If Node.IsParentOf(trvUsers.Nodes(pos)) Then
                            Disable = True
                        End If
                    End If
                    
                    If Disable Then
                        .Checked = False
                        .SmallIcon = IC_UNCHECKED
                        .Ghosted = True
                        .ForeColor = &H888888
                    End If
                End If
            End With
        Next j
    End If
    
    If ((tmp.BanMessage <> vbNullString) And (tmp.BanMessage <> "%")) Then
        txtBanMessage.Text = tmp.BanMessage
    End If
    
    Call UnlockGUI
    
    Node.Selected = True

    ' refresh tree view
    Call trvUsers.Refresh
End Sub

'Private Sub trvUsers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Dim node As cTreeViewNode
    
    'If (Button = vbLeftButton) Then
    '    Set node = trvUsers.HitTest(x, y)
    '    If Not node Is Nothing Then
    '        node.Selected = True
    '        Call trvUsers_NodeClick(trvUsers.SelectedItem)
    '    End If
    'End If
'End Sub

'Private Sub trvUsers_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, _
'    Shift As Integer, x As Single, y As Single, State As Integer)
'
'    Set trvUsers.SelectedItem = trvUsers.HitTest(x, y)
'End Sub

'Private Sub trvUsers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, _
'    Shift As Integer, x As Single, y As Single)
'
'    On Error GoTo ERROR_HANDLER
'
'    Dim nodePrev As cTreeViewNode
'    Dim nodeNow  As cTreeViewNode
'
'    Dim strKey   As String
'    Dim res      As Integer
'    Dim i        As Integer
'    Dim found    As Integer '
'
'    If (Data.GetFormat(15) = True) Then
'        With frmDBType
'            .setFilePath Data.Files(1)
'            .Show
'        End With
'    Else
'        Set nodeNow = trvUsers.NodeFromDragData(Data)
'
'        Set nodePrev = trvUsers.SelectedItem
'
'        If (nodeNow.Index = 1) Then
'            For i = LBound(m_DB) To UBound(m_DB)
'                If (StrComp(m_DB(i).Username, nodePrev.Text, vbTextCompare) = 0) Then
'                    If (StrComp(m_DB(i).Type, nodePrev.Tag, vbTextCompare) = 0) Then
'                        If ((Len(m_DB(i).Groups) > 0) And (m_DB(i).Groups <> "%")) Then
'                            Set nodePrev.Parent = nodeNow
'                        End If
'
'                        m_DB(i).Groups = vbNullString
'                        Exit For
'                    End If
'                End If
'            Next i
'        Else
'            If (Not nodePrev.Index = 1) Then
'                If (Not StrComp(nodeNow.Tag, "GROUP", vbTextCompare) = 0) Then
'                    Set nodeNow = nodeNow.Parent
'
'                    If (nodeNow.Index = 1) Then
'                        Set trvUsers.SelectedItem = nodeNow
'
'                        Call trvUsers_OLEDragDrop(Data, Effect, Button, Shift, x, y)
'                        Exit Sub
'                    End If
'                End If
'
'                'If (IsInGroup(nodePrev.Text, nodeNow.Text) = False) Then
'                '    For i = LBound(m_DB) To UBound(m_DB)
'                '        If (StrComp(m_DB(i).Username, nodePrev.Text, vbTextCompare) = 0) Then
'                '            If (StrComp(m_DB(i).Type, nodePrev.Tag, vbTextCompare) = 0) Then
'                '                m_DB(i).Groups = nodeNow.Text
'                '                Exit For
'                '            End If
'                '        End If
'                '    Next i
'                '
'                '    Set nodePrev.Parent = nodeNow
'                'End If
'            End If
'        End If
'
'        Call trvUsers_NodeClick(nodePrev)
'        'Set trvUsers.DropHighlight = Nothing
'    End If
'
'    Exit Sub
'
'ERROR_HANDLER:
'    ' potential cycle introduction error
'    If (Err.Number = 35614) Then
'        MsgBox Err.Description, vbCritical, "Error"
'    End If
'
'    'Set trvUsers.DropHighlight = Nothing
'
'    Exit Sub
'End Sub

Private Function IsInGroup(Groups As String, ByVal GroupName As String) As Boolean
    Dim j      As Integer
    Dim Splt() As String
    
    IsInGroup = False
    
    If (Len(Groups) > 0) Then
        If (Not InStr(1, Groups, ",", vbBinaryCompare) = 0) Then
            Splt() = Split(Groups, ",")
        Else
            ReDim Splt(0)
            Splt(0) = Groups
        End If
        
        For j = LBound(Splt) To UBound(Splt)
            If (StrComp(Splt(j), "%", vbBinaryCompare) <> 0) And (StrComp(GroupName, Splt(j), vbTextCompare) = 0) Then
                IsInGroup = True
                
                Exit Function
            End If
        Next j
    End If
End Function

Private Sub UpdateGroupList()
    Dim i As Integer
    Dim Count As Integer

    ' clear group selection listing
    Call lvGroups.ListItems.Clear
    
    m_glistcount = 0

    ' go through group listing
    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Type, "GROUP", vbTextCompare) = 0) Then
            m_glistcount = m_glistcount + 1
            ' add group to group selection listbox
            With lvGroups.ListItems.Add(, , m_DB(i).Username, , IC_UNCHECKED)
                .ForeColor = vbWhite
            End With
        End If
    Next i
    
    If m_glistcount = 0 Then
        With lvGroups.ListItems.Add(, , "[none]", , IC_EMPTY)
            .Ghosted = True
            .ForeColor = &H888888
        End With
    End If
End Sub

Private Sub UpdateInheritCaption(Groups As String)
    Dim grpwlk As udtGetAccessResponse
    Dim n As Integer
    Dim s As String
    
    s = vbNullString
    n = GetAccessGroupWalk(Groups, grpwlk)
    If n <> 1 Then s = "s"
    
    lblInherit.Caption = vbNullString
    If n > 0 Then
        If grpwlk.Rank > 0 And LenB(grpwlk.Flags) > 0 Then
            lblInherit.Caption = StringFormat("Inherits rank {2} and flags {3} from {0} group{1}.", n, s, grpwlk.Rank, grpwlk.Flags)
        ElseIf grpwlk.Rank > 0 Then
            lblInherit.Caption = StringFormat("Inherits rank {2} from {0} group{1}.", n, s, grpwlk.Rank)
        ElseIf LenB(grpwlk.Flags) > 0 Then
            lblInherit.Caption = StringFormat("Inherits flags {2} from {0} group{1}.", n, s, grpwlk.Flags)
        End If
    End If
End Sub

Private Sub ClearGroupListChecks()
    Dim i As Integer

    ' loop through listbox and clear selected items
    For i = 1 To lvGroups.ListItems.Count
        With lvGroups.ListItems(i)
            .Checked = False
            .SmallIcon = IIf(m_glistcount = 0, IC_EMPTY, IC_UNCHECKED)
            .Ghosted = True
            .ForeColor = &H888888
        End With
    Next i
End Sub

Private Function GetPrimaryGroup(ByVal Groups As String) As String
    Dim grp    As String
    Dim Splt() As String
    
    ' is it not in a group?
    If (LenB(Groups) = 0 Or StrComp(Groups, "%", vbBinaryCompare) = 0) Then
        grp = vbNullString
    ' is entry member of multiple groups?
    ElseIf (InStr(1, Groups, ",", vbBinaryCompare) <> 0) Then
        ' split up multiple groupings
        Splt() = Split(Groups, ",")
        grp = Splt(0)
    Else
        ' no need for special handling...
        grp = Groups
    End If
    
    GetPrimaryGroup = grp
End Function

Private Function GetLVPrimaryGroup() As ListItem
    Dim i As Integer
    
    For i = 1 To lvGroups.ListItems.Count
        With lvGroups.ListItems
            If (.Item(i).ForeColor = vbYellow And Not .Item(i).Ghosted And .Item(i).Checked) Then
                Set GetLVPrimaryGroup = .Item(i)
                Exit Function
            End If
        End With
    Next i
    
    Set GetLVPrimaryGroup = Nothing
End Function

Private Sub SetLVPrimaryGroup(ListItem As ListItem)
    Dim i As Integer
    
    If (Not ListItem Is Nothing) Then
        For i = 1 To lvGroups.ListItems.Count
            With lvGroups.ListItems.Item(i)
                If (StrComp(.Text, ListItem.Text, vbTextCompare) = 0) Then
                    .ForeColor = vbYellow
                    .Checked = True
                    .SmallIcon = IC_PRIMARY
                ElseIf (Not .Ghosted) Then
                    .ForeColor = vbWhite
                    .SmallIcon = IIf(.Checked, IC_CHECKED, IC_UNCHECKED)
                End If
            End With
        Next i
    End If
End Sub

Private Function FindNodeIndex(ByVal nodeName As String, Optional ByVal Tag As String = vbNullString, Optional ByVal NotChildOf As cTreeViewNode = Nothing) As Integer
    Dim i As Integer
    
    For i = 1 To trvUsers.NodeCount
        If (StrComp(trvUsers.Nodes(i).Text, nodeName, vbTextCompare) = 0) Then
            If (LenB(Tag) > 0) Then
                If (StrComp(trvUsers.Nodes(i).Tag, Tag, vbTextCompare) = 0) Then
                    If (Not trvUsers.Nodes(i).IsParentOf(NotChildOf)) Then
                        FindNodeIndex = i
                        Exit Function
                    End If
                End If
            Else
                If (Not trvUsers.Nodes(i).IsParentOf(NotChildOf)) Then
                    FindNodeIndex = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindNodeIndex = 0
End Function

Private Sub txtBanMessage_Change()
    ' enable entry save button
    Call HandleUnsaved
End Sub

Private Sub txtFlags_KeyPress(KeyAscii As Integer)
    
    Const AZ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    ' disallow entering space
    If (KeyAscii = 32) Then KeyAscii = 0
    
    ' if key is A-Z, then make uppercase
    If (InStr(1, AZ, ChrW$(KeyAscii), vbTextCompare) > 0) Then
        If (BotVars.CaseSensitiveFlags = False) Then
            If (KeyAscii > 90) Then ' lowercase if greater than "Z"
                KeyAscii = AscW(UCase$(ChrW$(KeyAscii)))
            End If
        End If
    ' else disallow entering that character (if not a control character)
    ElseIf (KeyAscii > 32) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtFlags_Change()
    ' enable entry save button
    Call HandleUnsaved
End Sub

Private Sub txtRank_KeyPress(KeyAscii As Integer)
    Const n09 As String = "0123456789"
    
    ' disallow entering space
    If (KeyAscii = 32) Then KeyAscii = 0
    
    ' if key is not 0-9, disallow entering that character (if not a control character)
    If (InStr(1, n09, ChrW$(KeyAscii), vbTextCompare) = 0 And KeyAscii > 32) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtRank_Change()
    Dim SelStart As Long
    
    If (Val(txtRank.Text) > 200) Then
        With txtRank
            SelStart = .SelStart
            .Text = "200"
            .SelStart = SelStart
        End With
    End If

    ' enable entry save button
    Call HandleUnsaved
End Sub

Private Function GetAccess(ByVal Username As String, ByRef Result As udtGetAccessResponse, _
    Optional ByVal dbType As String = vbNullString, Optional ByRef Index As Integer) As Boolean
    
    Dim i   As Integer
    Dim bln As Boolean
    
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
                Index = i
                
                GetAccess = True
                
                With Result
                    .Username = m_DB(i).Username
                    .Rank = m_DB(i).Rank
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

    GetAccess = False
End Function

' gets combined access of all Groups containing this item
Private Function GetAccessGroupWalk(ByVal Groups As String, ByRef Result As udtGetAccessResponse) As Integer
    Dim Splt()    As String
    Dim Group     As String
    Dim AllGroups As New Collection
    Dim tmp       As udtGetAccessResponse
    Dim MaxRank   As Integer
    Dim CombFlags As String
    Dim i         As Integer
    Dim j         As Integer
    
    If LenB(Groups) > 0 Then
        Splt() = Split(Groups, ",")
        For j = LBound(Splt) To UBound(Splt)
            If (StrComp(Splt(j), "%", vbBinaryCompare) <> 0) Then
                On Error GoTo ERROR_HANDLER
                Call AllGroups.Add(Splt(j), Splt(j))
                On Error GoTo 0
            End If
        Next j
    End If
    
    i = 1
    Do While i <= AllGroups.Count
        If GetAccess(AllGroups.Item(i), tmp, "GROUP") Then
            If tmp.Rank > MaxRank Then MaxRank = tmp.Rank
            CombFlags = CombFlags & tmp.Flags
            
            If LenB(tmp.Groups) > 0 And StrComp(tmp.Groups, "%", vbTextCompare) <> 0 Then
                Splt() = Split(tmp.Groups, ",")
                For j = LBound(Splt) To UBound(Splt)
                    If (StrComp(Splt(j), "%", vbBinaryCompare) <> 0) Then
                        On Error GoTo ERROR_HANDLER
                        Call AllGroups.Add(Splt(j), Splt(j))
                        On Error GoTo 0
                    End If
                Next j
            End If
        End If
        i = i + 1
    Loop
    
    GetAccessGroupWalk = AllGroups.Count
    
    With Result
        .Username = "(all groups)"
        .Rank = MaxRank
        .Flags = CombFlags
        .AddedBy = "(console)"
        .AddedOn = Now
        .ModifiedBy = vbNullString
        .ModifiedOn = 0
        .Type = "GROUP"
        .Groups = vbNullString
        .BanMessage = vbNullString
    End With
    
    Exit Function
ERROR_HANDLER:
    If (Err.Number = 457) Then
        Resume Next
    End If
End Function

Public Function DB_remove(ByVal entry As String, Optional ByVal dbType As String = _
    vbNullString) As Boolean
    
    On Error GoTo ERROR_HANDLER

    Dim i     As Integer
    Dim found As Boolean
    
    dbType = UCase$(dbType)
    
    For i = LBound(m_DB) To UBound(m_DB)
        If (StrComp(m_DB(i).Username, entry, vbTextCompare) = 0) Then
            Dim bln As Boolean
        
            If (Len(dbType)) Then
                If (StrComp(m_DB(i).Type, dbType, vbTextCompare) = 0) Then
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
        Dim bak As udtDatabase
        
        Dim j   As Integer
        
        bak = m_DB(i)

        ' we aren't removing the last array
        ' element, are we?
        If (UBound(m_DB) = 0) Then
            ReDim m_DB(0)
            
            With m_DB(0)
                .Username = vbNullString
                .Flags = vbNullString
                .Rank = 0
                .Groups = vbNullString
                .AddedBy = vbNullString
                .ModifiedBy = vbNullString
                .AddedOn = Now
                .ModifiedOn = Now
            End With
        Else
            For j = i To UBound(m_DB) - 1
                m_DB(j) = m_DB(j + 1)
            Next j
            
            ' redefine array size
            ReDim Preserve m_DB(UBound(m_DB) - 1)
            
            ' if we're removing a group, we need to also fix our
            ' group memberships, in case anything is broken now
            If (StrComp(bak.Type, "GROUP", vbBinaryCompare) = 0) Then
                Dim res As Boolean
            
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
                        If (Len(m_DB(i).Groups) > 0) Then
                            If (InStr(1, m_DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                                Dim Splt()     As String
                                Dim innerfound As Boolean
                                
                                Splt() = Split(m_DB(i).Groups, ",")
                                
                                For j = LBound(Splt) To UBound(Splt)
                                    If (StrComp(Splt(j), "%", vbBinaryCompare) <> 0) And (StrComp(bak.Username, Splt(j), vbTextCompare) = 0) Then
                                        innerfound = True
                                    
                                        Exit For
                                    End If
                                Next j
                            
                                If (innerfound) Then
                                    Dim K As Integer
                                    
                                    For K = (j + 1) To UBound(Splt)
                                        Splt(K - 1) = Splt(K)
                                    Next K
                                    
                                    ReDim Preserve Splt(UBound(Splt) - 1)
                                    
                                    m_DB(i).Groups = Join(Splt(), vbNullString)
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
        End If
        
        ' commit modifications
        'Call WriteDatabase(GetFilePath(FILE_USERDB))
        
        DB_remove = True
        
        Exit Function
    End If
    
    DB_remove = False
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in DB_remove().")
        
    DB_remove = False
    
    Exit Function
End Function
