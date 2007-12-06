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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Create User"
      Height          =   375
      Index           =   3
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
      HideSelection   =   0   'False
      Indentation     =   575
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "icons"
      Appearance      =   1
      OLEDragMode     =   1
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

Private m_dragging As Boolean ' ...
Private m_selnode  As Node    ' ...

Private Sub Form_Load()
    Call tbsTabs_Click
End Sub

Private Sub mnuDelete_Click()
    Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
End Sub

Private Sub mnuRename_Click()
    ' ...
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub tbsTabs_Click()
    Dim i      As Integer ' ...
    Dim Splt() As String  ' ...
    Dim j      As Integer ' ...
    Dim Pos    As Integer ' ...
    
    If (DB(0).Username = vbNullString) Then
        Call LoadDatabase
    End If

    Call trvUsers.Nodes.Clear
    
    Call trvUsers.Nodes.Add(, , "Database", "Database")
    
    Select Case (tbsTabs.SelectedItem.Index)
        Case 1: ' Users and Groups
            For i = LBound(DB) To UBound(DB)
                If (StrComp(DB(i).Type, "GROUP", vbBinaryCompare) = 0) Then
                    If (Len(DB(i).Groups) And (DB(i).Groups <> "%")) Then
                        If (InStr(1, DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Splt() = Split(DB(i).Groups, ",")
                        Else
                            ReDim Preserve Splt(0)
                            
                            Splt(0) = DB(i).Groups
                        End If
                        
                        For j = LBound(Splt) To UBound(Splt)
                            Pos = Exists(Splt(j))
                            
                            If (Pos) Then
                                Call trvUsers.Nodes.Add(trvUsers.Nodes(Pos).Key, _
                                    tvwChild, DB(i).Username, DB(i).Username, 1)
                            End If
                        Next j
                    Else
                        If (Not (Exists(DB(i).Username))) Then
                            Call trvUsers.Nodes.Add("Database", tvwChild, DB(i).Username, _
                                DB(i).Username, 1)
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
                If (StrComp(DB(i).Type, "USER", vbBinaryCompare) = 0) Then
                    If (Len(DB(i).Groups) And (DB(i).Groups <> "%")) Then
                        If (InStr(1, DB(i).Groups, ",", vbBinaryCompare) <> 0) Then
                            Splt() = Split(DB(i).Groups, ",")
                        Else
                            ReDim Preserve Splt(0)
                            
                            Splt(0) = DB(i).Groups
                        End If
                        
                        For j = LBound(Splt) To UBound(Splt)
                            Pos = Exists(Splt(j))
                            
                            If (Pos) Then
                                Call trvUsers.Nodes.Add(trvUsers.Nodes(Pos).Key, _
                                    tvwChild, DB(i).Username, DB(i).Username, 3)
                            End If
                        Next j
                    Else
                        Call trvUsers.Nodes.Add("Database", tvwChild, DB(i).Username, _
                                DB(i).Username, 3)
                    End If
                End If
            Next i
            
        Case 2: ' Clans
            For i = LBound(DB) To UBound(DB)
                If (StrComp(DB(i).Type, "CLAN", vbBinaryCompare) = 0) Then
                    Call trvUsers.Nodes.Add("Database", tvwChild, DB(i).Username, _
                            DB(i).Username, 2)
                End If
            Next i
            
        Case 3: ' Games
            For i = LBound(DB) To UBound(DB)
                If (StrComp(DB(i).Type, "GAME", vbBinaryCompare) = 0) Then
                    Dim gameName As String ' ...
                
                    ' check for invalid game entry
                    'Select Case (DB(i).Username)
                    '    Case "CHAT": gameName = "Chat"
                    '    Case "DRTL": gameName = "Diablo I: Retail"
                    '    Case "DSHR": gameName = "Diablo I: Shareware"
                    '    Case "W2BN": gameName = "WarCraft II: Battle.net Edition"
                    '    Case "STAR": gameName = "StarCraft"
                    '    Case "SSHR": gameName = "StarCraft: Shareware"
                    '    Case "JSTR": gameName = "StarCraft: Japanese"
                    '    Case "SEXP": gameName = "StarCraft: Brood War"
                    '    Case "D2DV": gameName = "Diablo II"
                    '    Case "D2XP": gameName = "Diablo II: Lord of Destruction"
                    '    Case "WAR3": gameName = "WarCraft III: Reign of Chaos"
                    '    Case "W3XP": gameName = "WarCraft III: The Frozen Throne"
                    '    Case Else:   gameName = DB(i).Username
                    'End Select
                    
                    gameName = DB(i).Username
                
                    Call trvUsers.Nodes.Add("Database", tvwChild, gameName, _
                            gameName, 2)
                End If
            Next i
    End Select
    
    If (trvUsers.Nodes.Count) Then
        With trvUsers.Nodes(1)
            .Expanded = True
            .Image = 1
        End With
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

    Set m_selnode = Node

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
            Splt() = Split(DB(i).Groups, ",")
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
        m_dragging = True
        
        Call trvUsers.Drag(vbBeginDrag)
    End If
End Sub

Private Sub trvUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
        Call Me.PopupMenu(mnuContext)
    End If
End Sub

Private Sub trvUsers_DragOver(ByRef Source As Control, ByRef X As Single, _
    ByRef Y As Single, ByRef State As Integer)
    
    If (m_dragging) Then
        If (Source.Name = "trvUsers") Then
            Set trvUsers.DropHighlight = trvUsers.HitTest(X, Y)
        End If
    End If
End Sub

Private Sub trvUsers_DragDrop(ByRef Source As Control, ByRef X As Single, _
    ByRef Y As Single)
    
    If (m_dragging) Then
        If (Source.Name = "trvUsers") Then
            If (Not (trvUsers.DropHighlight Is Nothing)) Then
                Dim current As Node ' ...
                Dim child   As Node ' ...
            
                Dim selnow  As udtGetAccessResponse ' ...
                Dim selprev As udtGetAccessResponse ' ...
                Dim gAcc    As udtGetAccessResponse ' ...
                
                Dim res     As Integer ' ...
                Dim i       As Integer ' ...
                Dim found   As Integer ' ...
                
                If (trvUsers.DropHighlight.Index = 1) Then
                    selprev = GetAccess(trvUsers.SelectedItem.text)
                    
                    If (Len(selprev.Groups) And (selprev.Groups <> "%")) Then
                        For i = LBound(DB) To UBound(DB)
                            If (StrComp(selprev.Username, DB(i).Username, vbBinaryCompare) = 0) Then
                                DB(i).Groups = vbNullString
                            End If
                        Next i
                        
                        ' ...
                        gAcc = GetCumulativeAccess(selprev.Username)
                        
                        Call trvUsers.Nodes.Add(trvUsers.DropHighlight.text, tvwChild, , _
                            selprev.Username, 1)
                            
                        If (trvUsers.Nodes(trvUsers.SelectedItem.Index).children) Then
                            'Set child = trvUsers.SelectedItem.Next

                            'Do While (Not (child Is Nothing))
                            '    Call trvUsers.Nodes.Add(trvUsers.DropHighlight.text, _
                            '        tvwChild, , selprev.Username)
                            '
                            '    Set child = trvUsers.SelectedItem.Next
                            'Loop
                        End If
                            
                        Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                        
                        If ((gAcc.access = 0) And _
                            (gAcc.Flags = vbNullString) And _
                            ((gAcc.Groups = vbNullString) Or _
                             (gAcc.Groups = "%"))) Then
                           
                            found = Exists(selprev.Username)
                            
                            If (found > 0) Then
                                With trvUsers.Nodes(found)
                                    .ForeColor = vbRed
                                End With
                            End If
                        End If
                    End If
                Else
                    If (trvUsers.SelectedItem.Index <> 1) Then
                        selnow = GetAccess(trvUsers.DropHighlight.text)
                        selprev = GetAccess(trvUsers.SelectedItem.text)
                        
                        If (selnow.Username <> selprev.Username) Then
                            If (StrComp(selnow.Type, "GROUP", vbBinaryCompare) = 0) Then
                                For i = LBound(DB) To UBound(DB)
                                    If (StrComp(selprev.Username, DB(i).Username, vbBinaryCompare) = 0) Then
                                        DB(i).Groups = trvUsers.DropHighlight.text
                                    End If
                                Next i
                                
                                ' ...
                                gAcc = GetCumulativeAccess(selprev.Username)
                                
                                Call trvUsers.Nodes.Add(trvUsers.DropHighlight.text, tvwChild, , _
                                    selprev.Username, 2)
    
                                Call trvUsers.Nodes.Remove(trvUsers.SelectedItem.Index)
                                
                                If ((gAcc.access = 0) And _
                                    (gAcc.Flags = vbNullString) And _
                                    ((gAcc.Groups = vbNullString) Or _
                                     (gAcc.Groups = "%"))) Then
                                   
                                    found = Exists(selprev.Username)
                                    
                                    If (found > 0) Then
                                        With trvUsers.Nodes(found)
                                            .ForeColor = vbRed
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            Set m_selnode = Nothing
            
            Set trvUsers.DropHighlight = Nothing
            
            m_dragging = False
        End If
    End If
End Sub

Private Sub btnCreateGroup_Click()
    Call frmGroupSelection.Show(vbModal, Me)
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
