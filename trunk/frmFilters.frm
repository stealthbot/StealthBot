VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFilters 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Username and Text Filters"
   ClientHeight    =   6750
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   7875
   Icon            =   "frmFilters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOutAdd 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   22
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txtOutAdd 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   3735
   End
   Begin MSComctlLib.ListView lvReplace 
      Height          =   3615
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6376
      View            =   3
      Arrange         =   2
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   10040064
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Words To Replace"
         Object.Width           =   6528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Replace With"
         Object.Width           =   6528
      EndProperty
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00000000&
      Caption         =   "Outgoing Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
      Width           =   2175
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00000000&
      Caption         =   "Incoming Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "&Remove Selected Item(s)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add It!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtAdd 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   7575
   End
   Begin VB.ListBox lbBlock 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   7575
   End
   Begin VB.ListBox lbText 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      ItemData        =   "frmFilters.frx":0CCA
      Left            =   120
      List            =   "frmFilters.frx":0CCC
      TabIndex        =   0
      Top             =   720
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Selected Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   1935
   End
   Begin VB.OptionButton optText 
      BackColor       =   &H00000000&
      Caption         =   "Message Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.OptionButton optBlock 
      BackColor       =   &H00000000&
      Caption         =   "Block List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdOutRem 
      Caption         =   "&Remove Selected Row"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdOutAdd 
      Caption         =   "&Add It!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label IncomingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Add to which list?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label IncomingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Username / Phrase to add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label OutgoingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Phrase to find:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000012&
      Caption         =   "More Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label IncomingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Username-Based Block List"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "-- StealthBot Custom Filters --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label IncomingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Text Message Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblMI 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   7575
   End
   Begin VB.Label OutgoingLbl 
      BackColor       =   &H80000012&
      Caption         =   "Phrase to replace with:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldMaxBlockIndex    As Integer
Private OldMaxFilterIndex   As Integer
' Updated later to erase old entries

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub cmdDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub cmdOutAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Allows you to edit the selected item."
End Sub

Private Sub cmdOutAdd_Click()
    If txtOutAdd(0).text <> vbNullString And txtOutAdd(1).text <> vbNullString Then
        lvReplace.ListItems.Add lvReplace.ListItems.Count + 1, , txtOutAdd(0).text
        lvReplace.ListItems.Item(lvReplace.ListItems.Count).ListSubItems.Add , , txtOutAdd(1).text
        txtOutAdd(0).text = vbNullString
        txtOutAdd(1).text = vbNullString
        txtOutAdd(0).SetFocus
    End If
End Sub

Private Sub cmdEdit_Click()
    If lbText.ListIndex <> -1 Then
        txtAdd.text = lbText.text
        Call cmdRem_CLick
    ElseIf lbBlock.ListIndex <> -1 Then
        txtAdd.text = lbBlock.text
        Call cmdRem_CLick
    ElseIf Not (lvReplace.SelectedItem Is Nothing) Then
        txtOutAdd(0).text = lvReplace.ListItems.Item(lvReplace.SelectedItem.Index).text
        txtOutAdd(1).text = lvReplace.ListItems.Item(lvReplace.SelectedItem.Index).ListSubItems.Item(1).text
        Call cmdOutRem_Click
    End If
End Sub

Private Sub cmdOutRem_Click()
    If Not (lvReplace.SelectedItem Is Nothing) Then
        If lvReplace.SelectedItem.Index > 0 Then lvReplace.ListItems.Remove lvReplace.SelectedItem.Index
    End If
End Sub

Private Sub cmdOutRem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Removes the selected item from the list."
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    Dim i As Integer, s As String
    
    optText.Value = True
    
    s = ReadINI("TextFilters", "Total", FILE_FILTERS)
    If StrictIsNumeric(s) Then
        i = Val(s)
        OldMaxFilterIndex = i
        
        If i > 0 Then
            For i = 1 To i
                s = ReadINI("TextFilters", "Filter" & i, FILE_FILTERS)
                
                If LenB(s) > 0 Then
                    lbText.AddItem s
                End If
            Next i
        End If
    End If
    
    s = ReadINI("BlockList", "Total", FILE_FILTERS)
    If StrictIsNumeric(s) Then
        i = Val(s)
        OldMaxBlockIndex = i
        
        If i > 0 Then
            For i = 1 To i
                s = ReadINI("BlockList", "Filter" & i, FILE_FILTERS)
                
                If LenB(s) > 0 Then
                    lbBlock.AddItem s
                End If
            Next i
        End If
    End If
    
    s = ReadINI("Outgoing", "Total", FILE_FILTERS)
    If StrictIsNumeric(s) Then
        i = Val(s)
        
        For i = 1 To i
            s = Replace(ReadINI("Outgoing", "Find" & i, FILE_FILTERS), "¦", " ")
            
            If LenB(s) > 0 Then
                lvReplace.ListItems.Add lvReplace.ListItems.Count + 1, , s
                lvReplace.ListItems.Item(lvReplace.ListItems.Count).ListSubItems.Add , , Replace(ReadINI("Outgoing", "Replace" & i, FILE_FILTERS), "¦", " ")
            End If
        Next i
    End If
    
    Call optType_Click(0)
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdRem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Clicking this button will remove any usernames or filters that you have selected on either list."
End Sub

Private Sub cmdRem_CLick()
    If lbText.ListIndex <> -1 Then lbText.RemoveItem lbText.ListIndex
    If lbBlock.ListIndex <> -1 Then lbBlock.RemoveItem lbBlock.ListIndex
End Sub

Private Sub lbBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Adding usernames to this list will cause the bot to completely block any and all messages from their username. The filter system supports wildcards. Example: Adding 'floodbot*' to the list will cause the bot to stop any messages coming from anybody whose name starts with the word 'FloodBot'. The filter is not applied to whispers."
End Sub

Private Sub lbText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Adding filters to this list will cause the bot to completely block any and all messages containing the text you added. Useful for blocking annoying floodbot or other spams."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMI.Caption = "Incoming chat filters are toggled by pressing CTRL + F inside the bot." & _
        vbNewLine & "Outgoing filters are permanently active."
End Sub

Private Sub cmdAdd_Click()
    If optBlock.Value = True Then
        lbBlock.AddItem txtAdd.text
    Else
        lbText.AddItem txtAdd.text
    End If
    txtAdd.text = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    i = -1
    
    ' Write text filters
    If lbText.ListCount > 0 Then
        For i = 1 To lbText.ListCount
            WriteINI "TextFilters", "Filter" & i, lbText.List(i - 1), FILE_FILTERS
        Next i
        
        WriteINI "TextFilters", "Total", lbText.ListCount, FILE_FILTERS
    Else
        WriteINI "TextFilters", "Total", 0, FILE_FILTERS
    End If
    
    ' Erase old text filters
    If i >= 0 And i < OldMaxFilterIndex Then
        For i = i To OldMaxFilterIndex
            WriteINI "TextFilters", "Filter" & i, "-", FILE_FILTERS
        Next i
    End If
    
    ' Write block list
    i = -1
    If lbBlock.ListCount <> 0 Then
        For i = 1 To lbBlock.ListCount
            WriteINI "BlockList", "Filter" & i, lbBlock.List(i - 1), FILE_FILTERS
        Next i
        
        WriteINI "BlockList", "Total", lbBlock.ListCount, FILE_FILTERS
    Else
        WriteINI "BlockList", "Total", 0, FILE_FILTERS
    End If
    
    ' Erase old blocked items
    If i >= 0 And i < OldMaxBlockIndex Then
        For i = i To OldMaxBlockIndex
            WriteINI "BlockList", "Filter" & i, "-", FILE_FILTERS
        Next i
    End If
    
    ' Write new outgoing filters
    If lvReplace.ListItems.Count <> 0 Then
        For i = 1 To lvReplace.ListItems.Count
            WriteINI "Outgoing", "Find" & i, Replace(lvReplace.ListItems.Item(i).text, " ", "¦"), FILE_FILTERS
            WriteINI "Outgoing", "Replace" & i, Replace(lvReplace.ListItems.Item(i).ListSubItems.Item(1).text, " ", "¦"), FILE_FILTERS
        Next i
        
        WriteINI "Outgoing", "Total", lvReplace.ListItems.Count, FILE_FILTERS
    Else
        WriteINI "Outgoing", "Total", 0, FILE_FILTERS
    End If
    
    Call frmChat.LoadOutFilters
    Call frmChat.LoadArray(LOAD_FILTERS, gFilters())
End Sub

Private Sub optBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lbBlock_MouseMove(0, 0, 0, 0)
End Sub

Private Sub optText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lbText_MouseMove(0, 0, 0, 0)
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Byte
    Dim A As Boolean
    
    If Index = 0 Then   'Incoming Filters Clicked
        optType(1).Value = False
        optType(0).Value = True
        A = True
    Else                'Outgoing Filters Clicked
        optType(0).Value = False
        optType(1).Value = True
        A = False
    End If
    
    lbText.Visible = A
    lbBlock.Visible = A
    txtAdd.Visible = A
    cmdAdd.Visible = A
    cmdRem.Visible = A
    
    For i = 0 To IncomingLbl.UBound
        IncomingLbl(i).Visible = A
    Next i
    
    For i = 0 To OutgoingLbl.UBound
        OutgoingLbl(i).Visible = Not A
    Next i
    
    optText.Visible = A
    optBlock.Visible = A
    lvReplace.Visible = Not A
    txtOutAdd(0).Visible = Not A
    txtOutAdd(1).Visible = Not A
    cmdOutRem.Visible = Not A
    cmdOutAdd.Visible = Not A
End Sub

Private Sub optType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        lblMI.Caption = "Change your incoming chat filters (message filters) here."
    Else
        lblMI.Caption = "Change your outgoing chat filters here."
    End If
End Sub
