VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDBManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   5895
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
   ScaleHeight     =   5895
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.Frame frmDatabase 
      Caption         =   "Database"
      Enabled         =   0   'False
      Height          =   4920
      Left            =   3600
      TabIndex        =   5
      Top             =   487
      Width           =   3025
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Index           =   1
         Left            =   1930
         TabIndex        =   10
         Top             =   4530
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   4530
         Width           =   855
      End
      Begin VB.ListBox lstGroups 
         Height          =   3180
         Left            =   240
         TabIndex        =   8
         Top             =   1280
         Width           =   2535
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFlags 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rank (1 - 200):"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Group:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1050
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnCreateGroup 
      Caption         =   "Create Group"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   5016
      Width           =   1695
   End
   Begin VB.CommandButton btnCreateUser 
      Caption         =   "Create User"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5016
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip tbsTabs 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   135
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
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clans"
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
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   5520
      Width           =   735
   End
   Begin MSComctlLib.TreeView trvUsers 
      Height          =   4350
      Left            =   120
      TabIndex        =   14
      Top             =   578
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7673
      _Version        =   393217
      Indentation     =   575
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "icons"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
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
      End
   End
End
Attribute VB_Name = "frmDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DB() As udtDatabase

Private Sub btnCreateUser_Click()
    Static userCount As Integer ' ...
    
    Dim newNode      As Node    ' ...
    
    Dim Username     As String  ' ...
    
    Username = "New User #" & _
        (userCount + 1)

    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
            tvwChild, Username, Username, 3)
    Else
        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, Username, _
            Username, 3)
    End If
        
    trvUsers.Nodes(newNode.Index).Selected = True
    
    Call trvUsers.StartLabelEdit
    
    userCount = (userCount + 1)
End Sub

Private Sub btnCreateGroup_Click()
    Static groupCount As Integer ' ...
    
    Dim newNode       As Node    ' ...
    
    Dim groupname     As String  ' ...
    
    groupname = "New Group #" & _
        (groupCount + 1)

    If (Not (trvUsers.SelectedItem Is Nothing)) Then
        Set newNode = trvUsers.Nodes.Add(trvUsers.SelectedItem.Key, _
            tvwChild, groupname, groupname, 1)
    Else
        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, groupname, _
            groupname, 1)
    End If
        
    trvUsers.Nodes(newNode.Index).Selected = True
    
    Call trvUsers.StartLabelEdit
    
    groupCount = (groupCount + 1)
End Sub

Private Sub Form_Load()
    If (m_DB(0).Username = vbNullString) Then
        Call LoadDatabase
    End If
    
    m_DB() = DB()
    
    Call tbsTabs_Click
End Sub

Private Sub mnuDelete_Click()
    Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
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
    
    Select Case (tbsTabs.SelectedItem.Index)
        Case 1: ' Users and Groups
            For i = LBound(DB) To UBound(DB)
                If (StrComp(m_DB(i).Type, "GROUP", vbBinaryCompare) = 0) Then
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
                                    tvwChild, m_DB(i).Username, m_DB(i).Username, 1)
                            End If
                        Next j
                    Else
                        If (Not (Exists(m_DB(i).Username))) Then
                            Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                                m_DB(i).Username, m_DB(i).Username, 1)
                        End If
                    End If
                End If
            Next i
            
            If (trvUsers.Nodes.Count > 1) Then
                Call lstGroups.Clear
            
                For i = 2 To trvUsers.Nodes.Count
                    Call lstGroups.AddItem(trvUsers.Nodes(i).text)
                Next i
            End If
            
            For i = LBound(DB) To UBound(DB)
                If (StrComp(m_DB(i).Type, "USER", vbBinaryCompare) = 0) Then
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
                                    tvwChild, m_DB(i).Username, m_DB(i).Username, 3)
                            End If
                        Next j
                    Else
                        Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                            m_DB(i).Username, m_DB(i).Username, 3)
                    End If
                End If
            Next i
            
        Case 2: ' Clans
            For i = LBound(DB) To UBound(DB)
                If (StrComp(m_DB(i).Type, "CLAN", vbBinaryCompare) = 0) Then
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, m_DB(i).Username, _
                            m_DB(i).Username, 2)
                End If
            Next i
            
        Case 3: ' Games
            For i = LBound(DB) To UBound(DB)
                If (StrComp(m_DB(i).Type, "GAME", vbBinaryCompare) = 0) Then
                    Set newNode = trvUsers.Nodes.Add("Database", tvwChild, _
                        m_DB(i).Username, m_DB(i).Username, 2)
                End If
            Next i
    End Select
    
    If (trvUsers.Nodes.Count) Then
        With trvUsers.Nodes(1)
            .Expanded = True
            .Image = 1
        End With
        
        Call trvUsers.Refresh
    End If
End Sub

Private Sub trvUsers_Collapse(ByVal Node As Node)
    If (Node.Index = 1) Then
        'trvUsers.Nodes(1).Image = 1
    End If
    
    Call trvUsers.Refresh
End Sub

Private Sub trvUsers_Expand(ByVal Node As Node)
    If (Node.Index = 1) Then
        'trvUsers.Nodes(1).Image = 2
    End If
    
    Call trvUsers.Refresh
End Sub

Private Sub trvUsers_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim tmp As udtGetAccessResponse ' ...
    
    Dim i   As Integer ' ...
    
    ' deselect groups
    For i = 0 To (lstGroups.ListCount - 1)
        lstGroups.Selected(i) = False
    Next i

    ' grab entry from database
    tmp = GetAccess(trvUsers.SelectedItem.text)
    
    If (Node.Index = 1) Then
        frmDatabase.Caption = "Database"
    Else
        If (tmp.Type = "USER") Then
            frmDatabase.Caption = "User: " & _
                tmp.Username
        ElseIf (tmp.Type = "CLAN") Then
            frmDatabase.Caption = "Clan: " & _
                tmp.Username
        ElseIf (tmp.Type = "GAME") Then
            frmDatabase.Caption = "Game: " & _
                tmp.Username
        ElseIf (tmp.Type = "GROUP") Then
            frmDatabase.Caption = "Group: " & _
                tmp.Username
        Else
            frmDatabase.Caption = _
                tmp.Username
        End If
    End If
    
    If (tmp.access > 0) Then
        txtRank.text = tmp.access
    Else
        txtRank.text = vbNullString
    End If
    
    txtFlags.text = tmp.Flags
    
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

        ' ...
        gAcc = GetAccess(trvUsers.SelectedItem.text)

        ' ...
        If (gAcc.Type = "GROUP") Then
            mnuRename.Enabled = True
        Else
            mnuRename.Enabled = False
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
      
    If (Not (trvUsers.DropHighlight Is Nothing)) Then
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
            selprev = GetAccess(nodeprev.text)

            For i = LBound(DB) To UBound(DB)
                If (StrComp(selprev.Username, m_DB(i).Username, vbBinaryCompare) = 0) Then
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
                    If (StrComp(selnow.Type, "GROUP", vbBinaryCompare) = 0) Then
                        For i = LBound(DB) To UBound(DB)
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
        
        Set trvUsers.DropHighlight = Nothing
    End If
End Sub

Private Function Exists(ByVal nodeName As String) As Integer
    Dim i As Integer ' ...
    
    For i = 1 To trvUsers.Nodes.Count
        If (StrComp(trvUsers.Nodes(i).text, nodeName, vbTextCompare) = 0) Then
            Exists = i
        
            Exit Function
        End If
    Next i
    
    Exists = False
End Function

Private Function GetIconIndex(ByVal Name As String) As Integer
    ' ...
End Function
