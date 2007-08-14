VERSION 5.00
Begin VB.Form frmTTT 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Enabled         =   0   'False
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
      Left            =   4440
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdInitConn 
      BackColor       =   &H00000000&
      Caption         =   "&Go"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtInitiate 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
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
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton cmd23 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmd22 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmd21 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmd33 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmd11 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmd32 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmd31 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmd13 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmd12 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label labelX 
      Caption         =   "Label4"
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00000000&
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
      Left            =   1800
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Currently connected to:"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Initiate a connection with:"
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
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Status"
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
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmTTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Status As Byte
Private ReceivedInitial As Byte
Private MyPlayer As Byte

Option Explicit
'STATUS
'0 = No Activity
'1 = Awaiting Response
'PLAYERS
'0 = O
'1 = X

Private Sub cmd11_Click()
If cmd11.Caption = vbNullString And Status = 0 Then
    Call SendTTT(1, 1, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd11.Caption = "O"
        Case 1
            cmd11.Caption = "X"
    End Select
End If
Call CheckWin
End Sub

Private Sub cmd12_Click()
If cmd12.Caption = vbNullString And Status = 0 Then
    Call SendTTT(1, 2, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd12.Caption = "O"
        Case 1
            cmd12.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd13_Click()
If cmd13.Caption = vbNullString And Status = 0 Then
    Call SendTTT(1, 3, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd13.Caption = "O"
        Case 1
            cmd13.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd21_Click()
If cmd21.Caption = vbNullString And Status = 0 Then
    Call SendTTT(2, 1, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd21.Caption = "O"
        Case 1
            cmd21.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd22_Click()
If cmd22.Caption = vbNullString And Status = 0 Then
    Call SendTTT(2, 2, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd22.Caption = "O"
        Case 1
            cmd22.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd23_Click()
If cmd23.Caption = vbNullString And Status = 0 Then
    Call SendTTT(2, 3, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd23.Caption = "O"
        Case 1
            cmd23.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd31_Click()
If cmd31.Caption = vbNullString And Status = 0 Then
    Call SendTTT(3, 1, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd31.Caption = "O"
        Case 1
            cmd31.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd32_Click()
If cmd32.Caption = vbNullString And Status = 0 Then
    Call SendTTT(3, 2, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd32.Caption = "O"
        Case 1
            cmd32.Caption = "X"
    End Select
    End If
    Call CheckWin
End Sub

Private Sub cmd33_Click()
If cmd33.Caption = vbNullString And Status = 0 Then
    Call SendTTT(3, 3, TTTEngaged)
    Select Case MyPlayer
        Case 0
            cmd33.Caption = "O"
        Case 1
            cmd33.Caption = "X"
    End Select
End If
Call CheckWin
End Sub

Sub cmdInitConn_Click()
    If txtInitiate.text = vbNullString Then
        MsgBox "You need to enter a person to play with."
        Exit Sub
    End If
    frmChat.AddQ "/w " & txtInitiate.text & "  ß~ßi" & MyPlayer & " Initialization request for Tic Tac Toe"
    uStatus "Requesting Tic Tac Toe."
    TTTEngaged = txtInitiate.text
End Sub

Sub cmdReset_Click()
    txtStatus.text = vbNullString
    frmChat.AddQ "/w " & txtInitiate.text & "  ß~ßi" & MyPlayer & " Initialization request for Tic Tac Toe"
    uStatus "Requesting another game."
    cmd11.Caption = vbNullString
    cmd12.Caption = vbNullString
    cmd13.Caption = vbNullString
    cmd21.Caption = vbNullString
    cmd22.Caption = vbNullString
    cmd23.Caption = vbNullString
    cmd31.Caption = vbNullString
    cmd32.Caption = vbNullString
    cmd33.Caption = vbNullString
    cmdReset.Enabled = False
End Sub

Sub cmdContinue_Click()
    cmd11.Caption = vbNullString
    cmd12.Caption = vbNullString
    cmd13.Caption = vbNullString
    cmd21.Caption = vbNullString
    cmd22.Caption = vbNullString
    cmd23.Caption = vbNullString
    cmd31.Caption = vbNullString
    cmd32.Caption = vbNullString
    cmd33.Caption = vbNullString
    txtStatus.text = vbNullString
    uStatus "Reset."
    Status = 0
    TTTEngaged = vbNullString
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    uStatus "Welcome to the StealthBot InterBot Tic Tac Toe application."
    cmdReset.Enabled = False
End Sub

Sub uStatus(ByVal Message As String)
    txtStatus.text = txtStatus.text & vbCrLf & "[" & Format(Time, "HH:MM:SS") & "] " & Message
    txtStatus.SelStart = Len(txtStatus.text)
End Sub

Sub SendTTT(ByVal Row As Byte, Column As Byte, Username As String)
    If Status = 0 Then
        If Dii Then
            frmChat.AddQ "/msg *" & Username & "  ß~ß" & Row & MyPlayer & Column
        Else
            frmChat.AddQ "/msg " & Username & "  ß~ß" & Row & MyPlayer & Column
        End If
        Status = 1
        
        uStatus "Message sent, awaiting reply."
    End If
End Sub

Sub uBoard(ByVal Row As Integer, Column As Integer, Player As Byte)
    Dim sPlayer As String
    
    Select Case Player
        Case 0
            sPlayer = "O"
        Case 1
            sPlayer = "X"
    End Select
    
    If Row = 1 Then
        If Column = 1 Then
            cmd11.Caption = sPlayer
        ElseIf Column = 2 Then
            cmd12.Caption = sPlayer
        Else
            cmd13.Caption = sPlayer
        End If
    ElseIf Row = 2 Then
        If Column = 1 Then
            cmd21.Caption = sPlayer
        ElseIf Column = 2 Then
            cmd22.Caption = sPlayer
        Else
            cmd23.Caption = sPlayer
        End If
    Else
        If Column = 1 Then
            cmd31.Caption = sPlayer
        ElseIf Column = 2 Then
            cmd32.Caption = sPlayer
        Else
            cmd33.Caption = sPlayer
        End If
    End If
    
    Call CheckWin
End Sub

Sub TTTArrival(ByVal Message As String)
    If Status = 0 And ReceivedInitial = 1 Then
        uStatus "Opponent has sent an extra message."
    Else
        uStatus "Received message."
        Status = 0
        Message = Right(Message, 3)
        If StrictIsNumeric(Message) Then
            uBoard Left$(Message, 1), Right(Message, 1), Left$(Right(Message, 2), 1)
            lblUser.Caption = TTTEngaged
            ReceivedInitial = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TTTEngaged = vbNullString
    Status = 0
    txtStatus.text = vbNullString
End Sub

Sub SetMyByte(ByVal PBYTE As Byte)
    MyPlayer = PBYTE
End Sub

Sub CheckWin()
    Dim b As Boolean, s As String, z As String
    'horizontal
    If StrComp(cmd11.Caption, cmd12.Caption, vbTextCompare) = 0 And StrComp(cmd11.Caption, cmd13.Caption, vbTextCompare) = 0 Then
         b = True
         s = cmd11.Caption
        End If
    If StrComp(cmd21.Caption, cmd22.Caption, vbTextCompare) = 0 And StrComp(cmd21.Caption, cmd23.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd21.Caption
        End If
    If StrComp(cmd31.Caption, cmd32.Caption, vbTextCompare) = 0 And StrComp(cmd31.Caption, cmd33.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd31.Caption
        End If
    'vertical
    If StrComp(cmd11.Caption, cmd21.Caption, vbTextCompare) = 0 And StrComp(cmd11.Caption, cmd31.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd11.Caption
        End If
    If StrComp(cmd12.Caption, cmd22.Caption, vbTextCompare) = 0 And StrComp(cmd12.Caption, cmd32.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd12.Caption
        End If
    If StrComp(cmd13.Caption, cmd23.Caption, vbTextCompare) = 0 And StrComp(cmd13.Caption, cmd33.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd13.Caption
        End If
    'diagonals
    If StrComp(cmd11.Caption, cmd22.Caption, vbTextCompare) = 0 And StrComp(cmd11.Caption, cmd33.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd11.Caption
        End If
    If StrComp(cmd13.Caption, cmd22.Caption, vbTextCompare) = 0 And StrComp(cmd13.Caption, cmd31.Caption, vbTextCompare) = 0 Then
        b = True
         s = cmd13.Caption
        End If
    
    If b = True Then
        If s = vbNullString Then
            Exit Sub
        End If
        Select Case MyPlayer
            Case 0
                z = "O"
            Case 1
                z = "X"
        End Select
        labelX.Caption = s & z
        If StrComp(s, z, vbTextCompare) = 0 Then
            uStatus "You have won! Click Continue to ask for another game."
        Else
            uStatus "You have lost. Click Continue to ask for another game."
        End If
        cmdReset.Enabled = True
        Status = 0
        ReceivedInitial = 1
        Exit Sub
    End If
    If cmd11.Caption <> vbNullString And cmd12.Caption <> vbNullString And cmd13.Caption <> vbNullString _
        And cmd21.Caption <> vbNullString And cmd22.Caption <> vbNullString And cmd23.Caption <> vbNullString _
            And cmd31.Caption <> vbNullString And cmd32.Caption <> vbNullString And cmd33.Caption <> vbNullString Then
        uStatus "Tie! Press Continue to ask for another game."
        cmdReset.Enabled = True
    End If
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
