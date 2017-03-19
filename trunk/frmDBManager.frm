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
   ClientWidth     =   7605
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
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCreateGame 
      Caption         =   "Ga&me"
      Height          =   375
      Left            =   3240
      Picture         =   "frmDBManager.frx":0000
      TabIndex        =   5
      ToolTipText     =   "Create game entry. Dynamically matches a user's current product."
      Top             =   5378
      Width           =   735
   End
   Begin VB.CommandButton cmdCreateClan 
      Caption         =   "C&lan"
      Height          =   375
      Left            =   2520
      Picture         =   "frmDBManager.frx":0468
      TabIndex        =   4
      ToolTipText     =   "Create WarCraft III Clan tag entry. Dynamically matches a user's clan tag."
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
   Begin VB.CommandButton cmdCreateGroup 
      Caption         =   "&Group"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Create group of entries."
      Top             =   5378
      Width           =   735
   End
   Begin VB.CommandButton cmdCreateUser 
      Caption         =   "Create &User..."
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   2
      ToolTipText     =   "Create user entry. Matches a user by exact name, or include ""*""s to match multiple users."
      Top             =   5378
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveForm 
      Caption         =   "Apply and Cl&ose"
      Default         =   -1  'True
      Height          =   300
      Left            =   6120
      TabIndex        =   25
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5400
      TabIndex        =   26
      Top             =   5880
      Width           =   735
   End
   Begin vbalTreeViewLib6.vbalTreeView trvUsers 
      Height          =   4965
      Left            =   120
      TabIndex        =   1
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
   Begin VB.Frame fraEntry 
      BackColor       =   &H80000007&
      Caption         =   "Eric[nK]"
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdSaveUser 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   22
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton cmdDiscardUser 
         Caption         =   "D&iscard"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   27
         Top             =   5160
         Width           =   735
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton cmdRenameUser 
         Caption         =   "Re&name"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox txtBanMessage 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   1320
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblInherit 
         BackColor       =   &H00000000&
         Caption         =   "Inherits:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label lblBanMessage 
         BackColor       =   &H00000000&
         Caption         =   "&Ban message"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3960
         Width           =   2775
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
         Width           =   2655
      End
      Begin VB.Label lblGroups 
         BackColor       =   &H00000000&
         Caption         =   "Grou&ps"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   2775
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
         TabIndex        =   16
         Top             =   1995
         Width           =   2535
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
         TabIndex        =   13
         Top             =   1300
         Width           =   2415
      End
      Begin VB.Label lblFlags 
         BackColor       =   &H00000000&
         Caption         =   "&Flags"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblRank 
         BackColor       =   &H00000000&
         Caption         =   "&Rank (1-200)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
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
         Top             =   1095
         Width           =   2655
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
         TabIndex        =   11
         Top             =   900
         Width           =   2775
      End
      Begin VB.Label lblModified 
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
         TabIndex        =   14
         Top             =   1605
         Width           =   2775
      End
   End
   Begin VB.Label lblDB 
      BackColor       =   &H00000000&
      Caption         =   "User &Database"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
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

Public m_EntryType     As String
Public m_EntryName     As String

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

' temporary DB working copy
Private m_DB           As clsDatabase
' current entry index
Private m_CurrentEntry As clsDBEntryObj
' current entry node
Private m_CurrNode     As cTreeViewNode
' is this entry modified
Private m_Modified     As Boolean
' is this a new entry (unused: if label editing worked, we'd use this)
Private m_New_Entry    As Boolean
' root of DB ("Database")
Private m_Root         As cTreeViewNode
' target for user node right-click menu
Private m_MenuTarget   As cTreeViewNode
' count of groups
Private m_GListCount   As Integer
' selected list item
Private m_GListSel     As ListItem
' target for group list right-click menu
Private m_GMenuTarget  As ListItem
' GUI is clearing and form items shouldn't change the header
Private m_ClearingUI   As Boolean

Private Sub Form_Load()

    Me.Icon = frmChat.Icon
    
    ' this line is gay but for some reason I can't set the ImageList for vbalTV in the designer/VB properties -Ribose
    trvUsers.ImageList = Icons.hImageList
    
    ' has our database been loaded?
    If Not Database.IsLoaded Then
        ' load database, if for some reason that hasn't been done
        Call Database.Load(GetFilePath(FILE_USERDB))
    End If
    
    ' store temporary copy of database
    Set m_DB = Database.CreateCopy()
    
    ' Set current entry
    Set m_CurrentEntry = Nothing
    
    ' show database for default tab
    Call LoadView
End Sub ' end function Form_Load

Public Sub ImportDatabase(strPath As String, dbType As Integer)
'    Dim f    As Integer
'    Dim buf  As String
'    Dim n    As cTreeViewNode
'    Dim i    As Integer
'    Dim tmp  As udtUserAccess
'
'    f = FreeFile
'
'    If (dbType = 0) Then
'
'        Open strPath For Input As #f
'            Do While (EOF(f) = False)
'                Line Input #f, buf
'
'                If (buf <> vbNullString) Then
'                    If (GetAccess(buf, tmp, DB_TYPE_USER, i)) Then
'                        ' do not save to tmp (it's a copy), but use m_DB(i)
'                        With m_DB(i)
'                            .Username = buf
'                            .Type = DB_TYPE_USER
'                            .ModifiedBy = "(console)"
'                            .ModifiedOn = Now
'
'                            If (InStr(1, .Flags, "S", vbBinaryCompare) = 0) Then
'                                .Flags = .Flags & "S"
'                            End If
'
'                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
'                                If (StrComp(trvUsers.SelectedItem.Tag, DB_TYPE_GROUP, vbTextCompare) = 0) Then
'                                    .Groups = .Groups & "," & trvUsers.SelectedItem.Text
'                                End If
'                            End If
'
'                            If (.Groups = vbNullString) Then
'                                .Groups = "%"
'                            End If
'                        End With
'                    Else
'                        ' redefine array to support new entry
'                        ReDim Preserve m_DB(UBound(m_DB) + 1)
'
'                        ' create new database entry
'                        With m_DB(UBound(m_DB))
'                            .Username = buf
'                            .Type = "USER"
'                            .AddedBy = "(console)"
'                            .AddedOn = Now
'                            .ModifiedBy = "(console)"
'                            .ModifiedOn = Now
'                            .Flags = "S"
'
'                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
'                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
'                                    If (Not IsInGroup(m_DB(UBound(m_DB)).Groups, trvUsers.SelectedItem.Text)) Then
'                                        .Groups = .Groups & "," & trvUsers.SelectedItem.Text
'                                    End If
'                                End If
'                            End If
'
'                            If (.Groups = vbNullString) Then
'                                .Groups = "%"
'                            End If
'                        End With
'                    End If
'                End If
'            Loop ' end loop
'        Close #f
'
'    ElseIf ((dbType = 1) Or (dbType = 2)) Then
'
'        Dim User As String
'        Dim Msg  As String
'
'        Open strPath For Input As #f
'            Do While (EOF(f) = False)
'                Line Input #f, buf
'
'                If (buf <> vbNullString) Then
'                    If (Not InStr(1, buf, Space$(1), vbBinaryCompare) = 0) Then
'                        User = Left$(buf, InStr(1, buf, Space$(1), vbBinaryCompare) - 1)
'
'                        Msg = Mid$(buf, Len(User) + 1)
'                    Else
'                        User = buf
'                    End If
'
'                    If (GetAccess(User, tmp, DB_TYPE_USER, i)) Then
'                        ' do not save to tmp (it's a copy), but use m_DB(i)
'                        With m_DB(i)
'                            .Username = User
'                            .Type = DB_TYPE_USER
'                            .ModifiedBy = "(console)"
'                            .ModifiedOn = Now
'
'                            If (InStr(1, .Flags, "B", vbBinaryCompare) = 0) Then
'                                .Flags = .Flags & "B"
'                            End If
'
'                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
'                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
'                                    If (Not IsInGroup(m_DB(i).Groups, trvUsers.SelectedItem.Text)) Then
'                                        .Groups = .Groups & "," & trvUsers.SelectedItem.Text
'                                    End If
'                                End If
'                            End If
'
'                            If (.Groups = vbNullString) Then
'                                .Groups = "%"
'                            End If
'                        End With
'                    Else
'                        ' redefine array to support new entry
'                        ReDim Preserve m_DB(UBound(m_DB) + 1)
'
'                       ' create new database entry
'                        With m_DB(UBound(m_DB))
'                            .Username = User
'                            .Type = "USER"
'                            .AddedBy = "(console)"
'                            .AddedOn = Now
'                            .ModifiedBy = "(console)"
'                            .ModifiedOn = Now
'                            .Flags = "B"
'                            .BanMessage = Msg
'
'                            If (Not (trvUsers.SelectedItem Is Nothing)) Then
'                                If (StrComp(trvUsers.SelectedItem.Tag, "GROUP", vbTextCompare) = 0) Then
'                                    .Groups = trvUsers.SelectedItem.Text
'                                End If
'                            End If
'
'                            If (.Groups = vbNullString) Then
'                                .Groups = "%"
'                            End If
'                        End With
'                    End If
'                End If
'            Loop ' end loop
'        Close #f
'
'    End If
'
'    LoadView
End Sub

Private Sub cmdCreateUser_Click()
    Dim oNewNode     As cTreeViewNode
    Dim sName       As String
    Dim iPos        As Integer
    Dim oEntry      As clsDBEntryObj
    
    m_EntryType = DB_TYPE_USER
    m_EntryName = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, Me)
    
    If (LenB(m_EntryName) > 0) Then
    
        sName = m_EntryName
    
        If (Not m_DB.ContainsEntry(sName, DB_TYPE_USER)) Then
            Set oEntry = m_DB.CreateNewEntry(sName, , DB_TYPE_USER)
            Call m_DB.AddEntry(oEntry)
            
            Set oNewNode = PlaceNewNode(oEntry, IC_USER)
            
            If (Not (oNewNode Is Nothing)) Then
                ' change misc. settings
                With oNewNode
                    '.Image = 0
                    .Tag = DB_TYPE_USER
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already a user with that name in the database."
            iPos = FindNodeIndex(sName, DB_TYPE_USER)
            If iPos > 0 Then
                trvUsers.Nodes(iPos).Selected = True
            End If
        End If
    End If
End Sub

Private Sub cmdCreateGroup_Click()
    Dim oNewNode    As cTreeViewNode
    Dim sGroup      As String
    Dim iPos         As Integer
    Dim oEntry      As clsDBEntryObj

    m_EntryType = DB_TYPE_GROUP
    m_EntryName = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, Me)
    
    If (LenB(m_EntryName) > 0) Then
    
        sGroup = m_EntryName
    
        If (Not m_DB.ContainsEntry(sGroup, DB_TYPE_GROUP)) Then
            Set oEntry = m_DB.CreateNewEntry(sGroup, , DB_TYPE_GROUP)
            Call m_DB.AddEntry(oEntry)
            
            Set oNewNode = PlaceNewNode(oEntry, IC_GROUP)
            
            Call UpdateGroupList
            
            If (Not (oNewNode Is Nothing)) Then
                ' change misc. settings
                With oNewNode
                    .Tag = DB_TYPE_GROUP
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "There is already a group with this name in the database."
            iPos = FindNodeIndex(sGroup, DB_TYPE_GROUP)
            If iPos > 0 Then
                trvUsers.Nodes(iPos).Selected = True
            End If
        End If
    End If
End Sub

Private Sub cmdCreateClan_Click()
    Dim oNewNode    As cTreeViewNode
    Dim sClan       As String
    Dim iPos        As Integer
    Dim oEntry      As clsDBEntryObj
    
    m_EntryType = DB_TYPE_CLAN
    m_EntryName = vbNullString
    
    Call frmDBNameEntry.Show(vbModal, Me)
    
    If (LenB(m_EntryName) > 0) Then
    
        sClan = m_EntryName
    
        If (Not m_DB.ContainsEntry(sClan, DB_TYPE_CLAN)) Then
            Set oEntry = m_DB.CreateNewEntry(sClan, , DB_TYPE_CLAN)
            Call m_DB.AddEntry(oEntry)
            
            Set oNewNode = PlaceNewNode(oEntry, IC_CLAN)
            
            If (Not (oNewNode Is Nothing)) Then
                ' change misc. settings
                With oNewNode
                    .Tag = DB_TYPE_CLAN
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "The specified clan is already in the database."
            iPos = FindNodeIndex(sClan, DB_TYPE_CLAN)
            If iPos > 0 Then
                trvUsers.Nodes(iPos).Selected = True
            End If
        End If
    End If
End Sub
        
Private Sub cmdCreateGame_Click()
    Dim oNewNode    As cTreeViewNode
    Dim sGame       As String
    Dim iPos        As Integer
    Dim oEntry      As clsDBEntryObj
    
    m_EntryName = vbNullString
    
    Call frmDBGameSelection.Show(vbModal, Me)
    
    If (LenB(m_EntryName) > 0) Then
    
        sGame = m_EntryName
        
        If (Not m_DB.ContainsEntry(sGame, DB_TYPE_GAME)) Then
            Set oEntry = m_DB.CreateNewEntry(sGame, , DB_TYPE_GAME)
            Call m_DB.AddEntry(oEntry)
        
            Set oNewNode = PlaceNewNode(oEntry, IC_GAME)
            
            If (Not (oNewNode Is Nothing)) Then
                ' change misc. settings
                With oNewNode
                    .Tag = DB_TYPE_GAME
                    .Selected = True
                End With
                
                'Call trvUsers_NodeClick(newNode)
            End If
        Else
            ' alert user that entry already exists
            MsgBox "The specified game is already in the database."
            iPos = FindNodeIndex(sGame, DB_TYPE_GAME)
            If iPos > 0 Then
                trvUsers.Nodes(iPos).Selected = True
            End If
        End If
    End If
End Sub

Private Function PlaceNewNode(oEntry As clsDBEntryObj, iEntryImage As Integer) As cTreeViewNode
    Dim oNewParent As cTreeViewNode
    
    ' by default create the node under the root node
    Set oNewParent = m_Root
    
    ' do we have an item (hopefully a group) selected?
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        ' is the item a group?
        If (StrComp(trvUsers.SelectedItem.Tag, DB_TYPE_GROUP, vbTextCompare) = 0) Then
            ' create new node under group node
            Set oNewParent = trvUsers.SelectedItem
        Else
            ' is our parent a group?
            If Not trvUsers.SelectedItem.Parent Is Nothing Then
                If (StrComp(trvUsers.SelectedItem.Parent.Tag, DB_TYPE_GROUP, vbTextCompare) = 0) Then
                    Set oNewParent = trvUsers.SelectedItem.Parent
                End If
            End If
        End If
    End If
    
    Set PlaceNewNode = trvUsers.Nodes.Add(oNewParent, etvwChild, _
        oEntry.EntryType & ": " & oEntry.Name, oEntry.Name, iEntryImage, iEntryImage)
                    
    ' set group settings on new database entry
    If (Not PlaceNewNode Is Nothing) Then
        If (StrComp(PlaceNewNode.Parent.Tag, DB_TYPE_GROUP, vbTextCompare) = 0) Then
            oEntry.AddGroup PlaceNewNode.Parent.Text
        End If
    End If
End Function

Private Sub cmdCancel_Click()
    m_Modified = False
    
    Unload Me
End Sub

Private Sub cmdDiscardUser_Click()
    m_Modified = False
    Call trvUsers_SelectedNodeChanged
End Sub

Private Sub cmdSaveUser_Click()
    Dim i               As Integer
    Dim sOldPrimary     As String
    Dim iPos            As Integer
    Dim oNewParent      As cTreeViewNode
    Dim oNode           As cTreeViewNode
    Dim oEntry          As clsDBEntryObj

    ' if we have no selected user... escape quick!
    If (trvUsers.SelectedItem Is Nothing) Then
        Exit Sub
    End If
    
    ' can't "save" the "Database"/root node
    If (StrComp(trvUsers.SelectedItem.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    ' disable entry save command
    Call HandleSaved

    ' get the selected user from database
    Set oEntry = m_DB.GetEntry(trvUsers.SelectedItem.Text, trvUsers.SelectedItem.Tag)

    ' modifiy user data
    With oEntry
        If Len(txtRank.Text) > 0 And StrictIsNumeric(txtRank.Text) Then
            .Rank = Int(txtRank.Text)
        End If
        .Flags = txtFlags.Text
        .ModifiedBy = m_DB.GetConsoleAccess().Username
        .ModifiedOn = Now()
        .BanMessage = txtBanMessage.Text
                    
        ' get old primary group
        If .Groups.Count = 0 Then
            sOldPrimary = vbNullString
        Else
            sOldPrimary = .Groups.Item(1)
        End If
                    
        ' build collection of new groups
        Call .ClearGroups
        If m_GListCount > 0 Then
            For i = 1 To lvGroups.ListItems.Count
                With lvGroups.ListItems(i)
                    If .Checked And Not .Ghosted Then
                        If ((.ForeColor = vbYellow) And (oEntry.Groups.Count > 0)) Then
                            oEntry.Groups.Add .Text, .Text, 1
                        Else
                            ' append
                            oEntry.Groups.Add .Text, .Text
                        End If
                    End If
                End With
            Next i
        End If
                    
        Set oNode = trvUsers.SelectedItem
        Set oNewParent = Nothing
        
        ' now to check if we need to move this node!
        ' did the "primary" group change?
        If .Groups.Count = 0 Then
            ' no longer assigned to any groups, go to root
            Set oNewParent = m_Root
        Else
            If (StrComp(sOldPrimary, .Groups.Item(1), vbTextCompare) <> 0) Then
                ' move under new primary
                iPos = FindNodeIndex(.Groups.Item(1), DB_TYPE_GROUP, oNode)
            
                ' well, does it exist?
                If (iPos > 0) Then
                    ' make node a child of existing group
                    Set oNewParent = trvUsers.Nodes(iPos)
                End If
            End If
        End If
        
        If Not (oNewParent Is Nothing) Then
            ' move node!!
            Call oNode.MoveNode(oNewParent, etvwChild)
            oNode.Tag = .EntryType
            oNode.Selected = True
        End If
    End With
                
End Sub

Private Sub HandleSaved()
    If Not m_ClearingUI Then
        m_Modified = False
        cmdSaveUser.Enabled = False
        cmdDiscardUser.Enabled = False
        If m_CurrentEntry Is Nothing Then
            fraEntry.Caption = vbNullString
            Me.Caption = "User Database Manager"
        Else
            fraEntry.Caption = m_CurrentEntry.ToString()
            Me.Caption = "User Database Manager - " & m_CurrentEntry.ToString()
        End If
    End If
End Sub

Private Sub HandleUnsaved()
    If Not m_ClearingUI Then
        If m_CurrentEntry Is Nothing Then
            m_Modified = False
            fraEntry.Caption = vbNullString
            Me.Caption = "User Database Manager *"
        Else
            m_Modified = True
            cmdSaveUser.Enabled = True
            cmdDiscardUser.Enabled = True
            fraEntry.Caption = m_CurrentEntry.ToString() & " *"
            Me.Caption = "User Database Manager - " & m_CurrentEntry.ToString() & " *"
        End If
    End If
End Sub

Private Sub cmdSaveForm_Click()
    ' save this user first
    If m_Modified Then
        cmdSaveUser_Click
    End If
    
    ' write temporary database to disk
    Call m_DB.Save(Database.FilePath)
    
    ' reload primary database
    Call Database.Load(Database.FilePath)
    
    ' check channel to find potential banned users
    Call g_Channel.CheckUsers
    
    ' close database form
    Call Unload(Me)
End Sub

Private Sub lvGroups_Click()
    Set m_GListSel = lvGroups.SelectedItem
    
    If Not m_GListSel Is Nothing Then
        m_GListSel.Checked = Not m_GListSel.Checked
        Call lvGroups_ItemCheck(m_GListSel)
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
    Dim i           As Integer
    Dim oGroups     As Collection

    Item.Selected = True

    If Item.Ghosted Then
        Item.Checked = False
        Item.SmallIcon = IIf(m_GListCount = 0, IC_EMPTY, IC_UNCHECKED)
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
    
    ' generate new group list
    Set oGroups = New Collection
    
    For i = 1 To lvGroups.ListItems.Count
        With lvGroups.ListItems(i)
            If .Checked And Not .Ghosted Then
                If ((.ForeColor = vbYellow) And (oGroups.Count > 0)) Then
                    ' place first
                    oGroups.Add .Text, .Text, 1
                Else
                    ' append
                    oGroups.Add .Text, .Text
                End If
            End If
        End With
    Next i
    
    ' update inherits list
    Call UpdateInheritCaption(oGroups)
    
    ' enable entry save command
    Call HandleUnsaved
End Sub

Private Sub lvGroups_KeyPress(KeyAscii As Integer)
    If m_GListSel <> lvGroups.SelectedItem Then
        Set m_GListSel = lvGroups.SelectedItem
    End If
    
    If KeyAscii = vbKeySpace Then
        If Not m_GListSel Is Nothing Then
            m_GListSel.Checked = Not m_GListSel.Checked
            Call lvGroups_ItemCheck(m_GListSel)
        End If
    End If
End Sub

Private Sub lvGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuSetPrimary.Visible = True
    mnuRename.Visible = False
    mnuDelete.Visible = False
    
    mnuSetPrimary.Enabled = False
    
    Set m_GMenuTarget = Nothing
    
    If (Button = vbRightButton) Then
        Set m_GMenuTarget = lvGroups.HitTest(x, y)
        
        If (Not m_GMenuTarget Is Nothing) Then
            If (Not m_GMenuTarget.Ghosted) Then
                mnuSetPrimary.Enabled = (m_GMenuTarget.ForeColor <> vbYellow)
            
                Call Me.PopupMenu(mnuContext)
            End If
        End If
    End If
End Sub

Private Sub mnuSetPrimary_Click()
    Call SetLVPrimaryGroup(m_GMenuTarget)
    
    Call HandleUnsaved
End Sub

Private Sub mnuDelete_Click()
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Call HandleDeleteEvent(m_MenuTarget)
    End If
End Sub

Private Sub cmdDeleteUser_Click()
    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Call HandleDeleteEvent(trvUsers.SelectedItem)
    End If
End Sub

Private Sub mnuRename_Click()
    Call QueryRenameEvent(m_MenuTarget)
End Sub

Private Sub cmdRenameUser_Click()
    Call QueryRenameEvent(trvUsers.SelectedItem)
End Sub

'Private Sub mnuOpenDatabase_Click()
'    ' open file dialog
'    Call CommonDialog.ShowOpen
'End Sub

Private Sub HandleDeleteEvent(oTarget As cTreeViewNode)
    If (oTarget Is Nothing) Then
        Exit Sub
    End If
    
    If (StrComp(oTarget.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    Dim rResponse   As VbMsgBoxResult   ' confirm delete response
    Dim bIsGroup    As Boolean          ' is the node a group
    Dim oEntry      As clsDBEntryObj    ' entry associated with deleted node
    
    If (m_Modified And StrComp(oTarget.Text, m_CurrNode.Text, vbBinaryCompare) = 0) Then
        ' do not ask about "unsaved data" later!
        m_Modified = False
    End If
    
    ' Is the target a group?
    bIsGroup = (StrComp(oTarget.Tag, DB_TYPE_GROUP, vbTextCompare) = 0)
    If (bIsGroup) Then
        rResponse = MsgBox("Are you sure you wish to delete this group? This " & _
            "may have unforeseen consequences.", vbYesNo Or vbInformation, "Database - Confirm Delete")
    End If
    
    ' If it's not a group or they said yes, go ahead and delete it.
    If ((bIsGroup = False) Or ((bIsGroup) And (rResponse = vbYes))) Then
        ' Remove from database
        Set oEntry = m_DB.GetEntry(oTarget.Text, oTarget.Tag)
        If Not (oEntry Is Nothing) Then
            Call m_DB.RemoveEntry(oEntry)
        End If
        
        ' Remove tree node
        oTarget.Delete
        
        If bIsGroup Then
            Call UpdateGroupList
            Call LoadView
        End If

        'Call trvUsers_NodeClick(trvUsers.SelectedItem)
    End If
End Sub

Private Sub QueryRenameEvent(Target As cTreeViewNode)
    If (Not (Target Is Nothing)) Then
        ' only "GROUP" entries can be renamed
        If (StrComp(Target.Tag, "GROUP", vbTextCompare) = 0) Then
            'trvUsers.SelectedItem.StartEdit
            
            m_EntryType = Target.Tag
            m_EntryName = Target.Text
            
            Call frmDBNameEntry.Show(vbModal, Me)
            
            If HandleRenameEvent(Target, m_EntryName) Then
            
                If LenB(m_EntryName) > 0 Then
                    Target.Text = m_EntryName
                    trvUsers.Refresh
                    
                    Call UpdateGroupList
                    
                    If m_Modified Then
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

Private Function HandleRenameEvent(oTarget As cTreeViewNode, sNewString As String) As Boolean
    Dim i, j        As Integer
    Dim oEntry      As clsDBEntryObj

    HandleRenameEvent = True
    
    If (oTarget Is Nothing) Then
        Exit Function
    End If
    
    If (StrComp(oTarget.Tag, "DATABASE", vbTextCompare) = 0) Then
        Exit Function
    End If
    
    If sNewString = oTarget.Text Or LenB(sNewString) = 0 Then
        ' same name succeeds (no chnage); empty name success (cancelled)
        Exit Function
    End If
    
    ' Check if the new name is in use.
    If m_DB.ContainsEntry(oTarget.Text, oTarget.Tag) Then
        HandleRenameEvent = False
        Exit Function
    End If
    
    ' Get the current entry
    Set oEntry = m_DB.GetEntry(oTarget.Text, oTarget.Tag)
    If oEntry Is Nothing Then
        HandleRenameEvent = False
        Exit Function
    End If
    
    ' If we are renaming a group, update all the members
    If (StrComp(oEntry.EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0) Then
        For i = 1 To m_DB.Entries.Count
            With m_DB.Entries.Item(i)
                If .IsInGroup(oEntry.Name) Then
                    ' Is it the primary?
                    If StrComp(.Groups.Item(1), oEntry.Name, vbTextCompare) = 0 Then
                        .Groups.Remove oEntry.Name
                        .Groups.Add oEntry.Name, oEntry.Name, 1
                    Else
                        .Groups.Remove oEntry.Name
                        .Groups.Add oEntry.Name, oEntry.Name
                    End If
                End If
            End With
        Next i
    End If
    
    ' Remove and re-add the entry with its new name.
    Call m_DB.RemoveEntry(oEntry)
    oEntry.Name = sNewString
    Call m_DB.AddEntry(oEntry)
End Function

' handle tab clicks and initial loading
Private Sub LoadView()
    Dim oNode           As cTreeViewNode    ' Recently added node
    Dim iImage          As Integer          ' Image value
    
    Dim i               As Integer          ' counter
    Dim iPos            As Integer          ' node position
    
    Dim cGroups         As New Collection   ' group list
    Dim cRemove         As New Collection   ' groups that have been added

    ' clear treeview
    Call trvUsers.Nodes.Clear
    
    ' create root node
    Set m_Root = trvUsers.Nodes.Add(, , "Database", "Database", IC_DATABASE, IC_DATABASE)
    ' type DATABASE
    m_Root.Tag = "DATABASE"
    
    ' If the database is empty, exit early
    If m_DB.Entries.Count = 0 Then
        Call UpdateGroupList
        Exit Sub
    End If

    ' Get all groups from the database
    For i = 1 To m_DB.Entries.Count
        With m_DB.Entries.Item(i)
            If StrComp(.EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0 Then
                cGroups.Add m_DB.Entries.Item(i), .Name
            End If
        End With
    Next i
    
    ' Add groups to their parents
    While cGroups.Count > 0
        For i = 1 To cGroups.Count
            With cGroups.Item(i)
                Set oNode = Nothing
                If .Groups.Count = 0 Then
                    ' Top level node
                    Set oNode = m_Root.AddChildNode(.ToString(), .Name, IC_GROUP, IC_GROUP)
                    cRemove.Add .Name
                Else
                    ' See if we've added this group's parent yet.
                    iPos = FindNodeIndex(.Groups.Item(1), DB_TYPE_GROUP)
                    If iPos > 0 Then
                        ' We have, so we can add this
                        Set oNode = trvUsers.Nodes(iPos).AddChildNode(.ToString(), .Name, IC_GROUP, IC_GROUP)
                        cRemove.Add .Name
                    End If
                End If
                
                ' Tag entry type for searching
                If Not (oNode Is Nothing) Then
                    oNode.Tag = DB_TYPE_GROUP
                End If
            End With
        Next i
        
        ' Remove placed groups
        If cRemove.Count > 0 Then
            For i = 1 To cRemove.Count
                cGroups.Remove cRemove.Item(i)
            Next i
            Set cRemove = New Collection
        End If
    Wend
    
    ' Add non-group entries
    For i = 1 To m_DB.Entries.Count
        With m_DB.Entries.Item(i)
            If StrComp(.EntryType, DB_TYPE_GROUP, vbBinaryCompare) <> 0 Then
                ' Determine which icon to use
                Select Case .EntryType
                    Case DB_TYPE_CLAN
                        iImage = IC_CLAN
                    Case DB_TYPE_GAME
                        iImage = IC_GAME
                    Case Else
                        iImage = IC_USER
                End Select
                
                Set oNode = Nothing
                If .Groups.Count = 0 Then
                    ' Top level node
                    Set oNode = m_Root.AddChildNode(.ToString(), .Name, iImage, iImage)
                Else
                    ' Try to find the parent node.
                    iPos = FindNodeIndex(.Groups.Item(1), DB_TYPE_GROUP)
                    If iPos > 0 Then
                        Set oNode = trvUsers.Nodes(iPos).AddChildNode(.ToString(), .Name, iImage, iImage)
                    Else
                        ' Uh oh. We couldn't find the parent even though they should all be there.
                        '   ... just add it to root
                        Set oNode = m_Root.AddChildNode(.ToString(), .Name, iImage, iImage)
                    End If
                End If
                
                ' Tag entry type for searching
                If Not (oNode Is Nothing) Then
                    oNode.Tag = .EntryType
                End If
            End If
        End With
    Next i
    
    ' Expand all of the nodes
    For i = 1 To trvUsers.NodeCount
        trvUsers.Nodes(i).Expanded = True
    Next i
    
    Call UpdateGroupList
End Sub

Private Sub LockGUI()
    Dim i As Integer

    m_ClearingUI = True

    ' set our default frame caption
    Set m_CurrentEntry = Nothing
    Call HandleSaved

    ' disable & clear rank
    lblRank.Enabled = False
    txtRank.Enabled = False
    txtRank.Text = vbNullString

    ' disable & clear flags
    lblFlags.Enabled = False
    txtFlags.Enabled = False
    txtFlags.Text = vbNullString

    ' loop through listbox and clear selected items
    Call ClearGroupListChecks

    ' disable group lists
    lblGroups.Enabled = False
    'lvGroups.Enabled = False

    ' disable & clear ban message
    lblBanMessage.Enabled = False
    txtBanMessage.Enabled = False
    txtBanMessage.Text = vbNullString

    ' reset created & modified labels
    lblCreated.Enabled = False
    lblCreatedOn.Enabled = False
    lblCreatedOn.Caption = "(not applicable)"
    lblCreatedBy.Enabled = False
    lblCreatedBy.Caption = vbNullString
    lblModified.Enabled = False
    lblModifiedOn.Enabled = False
    lblModifiedOn.Caption = "(not applicable)"
    lblModifiedBy.Enabled = False
    lblModifiedBy.Caption = vbNullString

    ' reset inherits caption
    lblInherit.Caption = vbNullString

    ' disable entry buttons
    cmdRenameUser.Enabled = False
    cmdDeleteUser.Enabled = False

    m_ClearingUI = False

    HandleSaved
End Sub

Private Sub UnlockGUI()
    Dim i As Integer

    m_ClearingUI = True

    ' enable rank field
    lblRank.Enabled = True
    txtRank.Enabled = True

    ' enable flags field
    lblFlags.Enabled = True
    txtFlags.Enabled = True
    
    ' enable ban message field
    lblBanMessage.Enabled = True
    txtBanMessage.Enabled = True
    
    ' enable labels
    lblCreated.Enabled = True
    lblCreatedOn.Enabled = True
    lblCreatedBy.Enabled = True
    lblModified.Enabled = True
    lblModifiedOn.Enabled = True
    lblModifiedBy.Enabled = True
    
    ' enable entry rename/delete buttons
    cmdRenameUser.Enabled = (StrComp(trvUsers.SelectedItem.Tag, DB_TYPE_GROUP, vbTextCompare) = 0)
    cmdDeleteUser.Enabled = True
    
    ' enable group lists
    lblGroups.Enabled = True
    'lvGroups.Enabled = True

    m_ClearingUI = False
    
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
    
    Set m_MenuTarget = Nothing
    
    If (Not (Node Is Nothing)) Then
        If (StrComp(Node.Tag, "DATABASE", vbTextCompare) <> 0) Then
            mnuRename.Enabled = (StrComp(Node.Tag, DB_TYPE_GROUP, vbTextCompare) = 0)
            mnuDelete.Enabled = True
            
            Set m_MenuTarget = Node
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
    Static bSkipUpdate  As Boolean
    
    Dim oNode       As cTreeViewNode
    Dim i, j        As Integer
    Dim iPos        As Integer
    Dim rResponse   As VbMsgBoxResult
    Dim bDisable    As Boolean
    Dim oEntry      As clsDBEntryObj
    
    Set oNode = trvUsers.SelectedItem
    
    If (oNode Is Nothing) Then
        Exit Sub
    End If
    
    If m_Modified And Not bSkipUpdate And (Not (m_CurrNode Is Nothing)) Then
        ' check if we should allow this node change? (is Unsaved?)
        rResponse = MsgBox("Are you sure you wish to discard changes to the " & _
            m_CurrNode.Text & " (" & UCase$(m_CurrNode.Tag) & ") database entry?", _
            vbYesNo Or vbInformation, "Database - Confirm Discard Changes")
        
        If rResponse = vbNo Then
            bSkipUpdate = True
            m_CurrNode.Selected = True
            bSkipUpdate = False
            Exit Sub
        End If
    End If
    
    If bSkipUpdate Then Exit Sub
    
    Set m_CurrNode = oNode
    
    Call LockGUI
    
    oNode.Expanded = True
    
    If (StrComp(oNode.Tag, "DATABASE", vbTextCompare) = 0) Then
        Set m_CurrentEntry = Nothing
        Exit Sub
    End If
    
    ' Get the selected entry.
    Set oEntry = m_DB.GetEntry(oNode.Text, oNode.Tag)
    Set m_CurrentEntry = oEntry

    ' does entry have a rank?
    If (oEntry.Rank > 0) Then
        ' write rank to text box
        txtRank.Text = oEntry.Rank
    Else
        ' clear rank from text box
        txtRank.Text = vbNullString
    End If
    
    ' clear flags from text box
    txtFlags.Text = oEntry.Flags
    
    If (oEntry.CreatedBy = vbNullString) Then
        lblCreatedOn = "unknown"
        lblCreatedBy = "by unknown"
    Else
        lblCreatedOn = oEntry.CreatedOn & " Local Time"
        lblCreatedBy = "by " & oEntry.CreatedBy
    End If
    
    If (oEntry.ModifiedBy = vbNullString) Then
        lblModifiedOn = "unknown"
        lblModifiedBy = "by unknown"
    Else
        lblModifiedOn = oEntry.ModifiedOn & " Local Time"
        lblModifiedBy = "by " & oEntry.ModifiedBy
    End If
    
    Call UpdateInheritCaption(oEntry.Groups)
    
    ' loop through our listview, checking for matches
    If m_GListCount > 0 Then
        For j = 1 To lvGroups.ListItems.Count
            With lvGroups.ListItems(j)
                ' loop through entry's group memberships
                .Checked = False
                .SmallIcon = IC_UNCHECKED
                .Ghosted = False
                .ForeColor = vbWhite
                
                If oEntry.Groups.Count > 0 Then
                    For i = 1 To oEntry.Groups.Count
                        If (StrComp(oEntry.Groups.Item(i), .Text, vbTextCompare) = 0) Then
                            ' select group if entry is a member
                            .Checked = True
                            .SmallIcon = IC_CHECKED
                            
                            ' highlight group if "primary" (first group)
                            If (i = 1) Then
                                .ForeColor = vbYellow
                                .SmallIcon = IC_PRIMARY
                            End If
                            
                            Exit For
                        End If
                    Next i
                End If
                
                If (StrComp(oEntry.EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0) Then
                    bDisable = False
                    
                    ' don't allow a group to be in itself
                    If (StrComp(oEntry.Name, .Text, vbTextCompare) = 0) Then
                        bDisable = True
                    End If
                    
                    ' don't allow a group to be in its children
                    iPos = FindNodeIndex(.Text)
                    If iPos > 0 Then
                        If oNode.IsParentOf(trvUsers.Nodes(iPos)) Then
                            bDisable = True
                        End If
                    End If
                    
                    If bDisable Then
                        .Checked = False
                        .SmallIcon = IC_UNCHECKED
                        .Ghosted = True
                        .ForeColor = &H888888
                    End If
                End If
            End With
        Next j
    End If
    
    If Len(oEntry.BanMessage) > 0 Then
        txtBanMessage.Text = oEntry.BanMessage
    End If
    
    Call UnlockGUI
    
    oNode.Selected = True

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

Private Sub UpdateGroupList()
    Dim i As Integer
    Dim Count As Integer

    ' clear group selection listing
    Call lvGroups.ListItems.Clear
    
    m_GListCount = 0

    ' go through group listing
    For i = 1 To m_DB.Entries.Count
        If (StrComp(m_DB.Entries.Item(i).EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0) Then
            m_GListCount = m_GListCount + 1
            ' add group to group selection listbox
            With lvGroups.ListItems.Add(, , m_DB.Entries.Item(i).Name, , IC_UNCHECKED)
                .ForeColor = vbWhite
            End With
        End If
    Next i
    
    If m_GListCount = 0 Then
        With lvGroups.ListItems.Add(, , "[none]", , IC_EMPTY)
            .Ghosted = True
            .ForeColor = &H888888
        End With
    End If
End Sub

Private Sub UpdateInheritCaption(ByRef oGroups As Collection)
    Dim oAccess     As udtUserAccess
    Dim oFakeEntry  As clsDBEntryObj
    Dim s           As String           ' group(s)
    Dim n           As Integer          ' number of groups
    Dim i           As Integer          ' counter
    
    ' Create a dummy entry to assign to these groups
    Set oFakeEntry = m_DB.CreateNewEntry("<dummy>")
    For i = 1 To oGroups.Count
        oFakeEntry.AddGroup oGroups.Item(i)
    Next i
    
    ' Get the access available to an entry with all these groups
    oAccess = m_DB.GetEntryAccess(oFakeEntry)
    
    ' Set formatting variables
    n = oAccess.Groups.Count
    s = IIf(n <> 1, "s", vbNullString)
    
    ' Set caption
    lblInherit.Caption = vbNullString
    If n > 0 Then
        If oAccess.Rank > 0 And LenB(oAccess.Flags) > 0 Then
            lblInherit.Caption = StringFormat("Inherits rank {2} and flags {3} from {0} group{1}.", n, s, oAccess.Rank, oAccess.Flags)
        ElseIf oAccess.Rank > 0 Then
            lblInherit.Caption = StringFormat("Inherits rank {2} from {0} group{1}.", n, s, oAccess.Rank)
        ElseIf LenB(oAccess.Flags) > 0 Then
            lblInherit.Caption = StringFormat("Inherits flags {2} from {0} group{1}.", n, s, oAccess.Flags)
        End If
    End If
End Sub

Private Sub ClearGroupListChecks()
    Dim i As Integer

    ' loop through listbox and clear selected items
    For i = 1 To lvGroups.ListItems.Count
        With lvGroups.ListItems(i)
            .Checked = False
            .SmallIcon = IIf(m_GListCount = 0, IC_EMPTY, IC_UNCHECKED)
            .Ghosted = True
            .ForeColor = &H888888
        End With
    Next i
End Sub

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
    Call HandleUnsaved
End Sub

Private Sub txtRank_KeyPress(KeyAscii As Integer)
    ' disallow entering space
    If (KeyAscii = vbKeySpace) Then KeyAscii = 0
    
    ' if key is not 0-9, disallow entering that character (if not a control character)
    If (InStr(1, Num09, ChrW$(KeyAscii), vbTextCompare) = 0 And KeyAscii > 32) Then
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

