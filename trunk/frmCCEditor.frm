VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCCEditor 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Command Editor"
   ClientHeight    =   5775
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Save All and E&xit"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvCCList 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   8493
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Selected"
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
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtCCAction 
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
      Height          =   4485
      Left            =   2160
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.TextBox txtCCAccess 
      Alignment       =   2  'Center
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
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename Selected"
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
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Current Commands:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Click here for more information on Custom Commands."
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Actions to take:"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Access required:"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'- frmCCEditor.frm
Option Explicit

Private CurrentCmd As Integer
Private Commands() As udtCustomCommandData
Private Modified As Boolean


Private Sub Form_Load()
    'On Error Resume Next
    Me.Icon = frmChat.Icon
    Dim i As Integer, f As Integer, r As Integer
    Dim ccIn As udtCustomCommandData
    
    ReDim Commands(0)
    
    f = FreeFile
    
    If LenB(Dir$(GetFilePath("commands.dat"))) Then
        Open (GetFilePath("commands.dat")) For Random As #f Len = LenB(ccIn)
        
        r = LOF(f) \ LenB(ccIn)
        If LOF(f) Mod LenB(ccIn) <> 0 Then r = r + 1
        
        For i = 1 To r
            Get #f, i, ccIn
            If ccIn.reqAccess < 1001 And Len(RTrim(ccIn.Action)) > 0 Then
                Commands(UBound(Commands)) = ccIn
                lvCCList.ListItems.Add , , RTrim(Commands(UBound(Commands)).Query)
                If i <> r Then ReDim Preserve Commands(UBound(Commands) + 1)
            End If
        Next i
        
        Close #f
                
        If lvCCList.ListItems.Count > 0 Then
            lvCCList.ListItems(1).Selected = True
            CurrentCmd = 0
            ShowCurrentCommand
        Else
            CurrentCmd = -1
        End If
            
    End If
End Sub

Private Sub cmdApply_Click()
    Call SaveCurrent
    Call SaveAll
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer, Iterations As Integer
    
    If Not (lvCCList.SelectedItem Is Nothing) Then
        
        'Debug.Print "CurrentCmd is " & CurrentCmd
        If UBound(Commands) > 0 Then
            While i <= (UBound(Commands))
                If i >= CurrentCmd And i < UBound(Commands) Then
                    'Debug.Print "Moving command from position " & i + 1 & " to position " & i
                    Commands(i).Action = Commands(i + 1).Action
                    Commands(i).Query = Commands(i + 1).Query
                    Commands(i).reqAccess = Commands(i + 1).reqAccess
                End If
                
                i = i + 1
                Iterations = Iterations + 1
                
                If Iterations > 7000 Then
                    frmChat.AddChat RTBColors.ErrorMessageText, "Warning: Loop size limit exceeded in cmdDelete_Click()!"
                    Exit Sub
                End If
            Wend
            
            ReDim Preserve Commands(UBound(Commands) - 1)
        Else
            ReDim Commands(0)
        End If
        
        If lvCCList.ListItems.Count > 1 Then
            lvCCList.ListItems.Remove lvCCList.SelectedItem.Index
            
            If CurrentCmd > 1 Then
                lvCCList.ListItems(CurrentCmd - 1).Selected = True
            Else
                lvCCList.ListItems(1).Selected = True
            End If
            
            Modified = False
            Call lvCCList_Click
            
            ShowCurrentCommand
        Else
            CurrentCmd = -1
            txtCCAction.text = ""
            txtCCAccess.text = ""
            If lvCCList.ListItems.Count = 1 Then lvCCList.ListItems.Remove 1
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    'Dim f As Integer, i As Integer, c As Integer
    
    Call SaveCurrent
    Call SaveAll
    
    Modified = False
    Call Form_Unload(0)
End Sub

Private Sub cmdNew_Click()
    Call SaveCurrent
    
    If UBound(Commands) = 0 Then
        If lvCCList.ListItems.Count = 0 Then
            CurrentCmd = 0
        Else
            ReDim Preserve Commands(UBound(Commands) + 1)
            CurrentCmd = 1
        End If
    Else
        ReDim Preserve Commands(UBound(Commands) + 1)
        CurrentCmd = UBound(Commands)
    End If
    
    txtCCAccess.text = 0
    txtCCAction.text = vbNullString
    
    Modified = False
    
    With lvCCList
        .ListItems.Add , , "NewCommand"
        .ListItems.Item(.ListItems.Count).Selected = True
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modified Then
        Call cmdExit_Click
    End If
    
    Unload Me
End Sub

Private Sub Label3_Click()
    OpenReadme
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbBlue
End Sub

Private Sub lvCCList_Click()
    Call SaveCurrent
    
    If Not (lvCCList.SelectedItem Is Nothing) Then
        CurrentCmd = lvCCList.SelectedItem.Index - 1
    
        ShowCurrentCommand
    End If
End Sub

Private Sub txtCCAccess_Change()
    Modified = True
End Sub

Private Sub txtCCAccess_LostFocus()
    ' Warning added 8/4/06
    If Val(txtCCAccess.text) > 1000 Then
        MsgBox "Adding commands greater than access 1,000 will render them totally useless. Please enter a valid access level, from 0 to 1,000." _
                , vbExclamation + vbOKOnly, "StealthBot Custom Command Editor"
                
        With txtCCAccess
            .SelStart = 0
            .SelLength = Len(.text)
            .SetFocus
        End With
    End If
End Sub

Private Sub txtCCAction_Change()
    Modified = True
End Sub

Private Sub txtCCAction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbWhite
End Sub

Function SaveCurrent() As Boolean
    If Modified And CurrentCmd >= 0 Then
        Commands(CurrentCmd).Action = Replace(txtCCAction.text, vbCrLf, "& ")
        Commands(CurrentCmd).reqAccess = Val(txtCCAccess.text)
        Commands(CurrentCmd).Query = lvCCList.ListItems(CurrentCmd + 1).text
        Modified = False
        SaveCurrent = True
    End If
End Function

Sub SaveAll()
    Dim f As Integer, i As Integer, c As Integer

    If lvCCList.ListItems.Count > 0 Then
        f = FreeFile
        
        If LenB(Dir(GetFilePath("commands.dat"))) Then
            Open GetFilePath("commands.dat") For Output As #f: Close #f
        End If
        
        Open GetFilePath("commands.dat") For Random As #f Len = LenB(Commands(0))
        
        For i = 0 To UBound(Commands)
            With Commands(i)
                If .reqAccess < 1001 Then
                    c = c + 1
                    Put #f, c, Commands(i)
                End If
            End With
        Next i
        
        Close #f
    End If
End Sub

Private Sub cmdRename_Click()
    Dim toName As String
    
    If Not (lvCCList.SelectedItem Is Nothing) Then
        toName = InputBox("What would you like the command to be named?", "Rename your Custom Command", lvCCList.SelectedItem.text)
        
        lvCCList.SelectedItem.text = toName
        Commands(lvCCList.SelectedItem.Index - 1).Query = toName
        SaveCurrent
    End If
End Sub

Sub ShowCurrentCommand()
    With Commands(CurrentCmd)
        txtCCAccess.text = .reqAccess
        txtCCAction.text = Replace(RTrim(.Action), "& ", vbCrLf)
    End With
End Sub
