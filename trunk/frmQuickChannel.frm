VERSION 5.00
Begin VB.Form frmQuickChannel 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuickChannel Manager"
   ClientHeight    =   3090
   ClientLeft      =   2610
   ClientTop       =   1545
   ClientWidth     =   4035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Channel 
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
      Index           =   9
      Left            =   600
      MaxLength       =   31
      TabIndex        =   18
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   8
      Left            =   600
      MaxLength       =   31
      TabIndex        =   16
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   7
      Left            =   600
      MaxLength       =   31
      TabIndex        =   14
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   6
      Left            =   600
      MaxLength       =   31
      TabIndex        =   12
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   5
      Left            =   600
      MaxLength       =   31
      TabIndex        =   10
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   4
      Left            =   600
      MaxLength       =   31
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   3
      Left            =   600
      MaxLength       =   31
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Index           =   2
      Left            =   600
      MaxLength       =   31
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Channel 
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
      Left            =   600
      MaxLength       =   31
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Save"
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
      Left            =   1560
      TabIndex        =   19
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&9:"
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
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&8:"
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
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&7:"
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
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&6:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&5:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&4:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&3:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&2:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "F&1:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblChannel 
      BackColor       =   &H00000000&
      Caption         =   "Set your QuickChannels:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmQuickChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    frmQuickChannel.KeyPreview = True
    
    Dim i As Integer
    
    ' bounds of Channel controls
    For i = Channel.LBound To Channel.UBound
        Channel(i) = QC(i)
    Next i
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdDone_Click
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub cmdDone_Click()
    Dim i As Integer
    
    'write the qc list
    ' bounds of Channel controls
    For i = Channel.LBound To Channel.UBound
        QC(i) = Channel(i).Text
    Next i
    
    SaveQuickChannels
    
    PrepareQuickChannelMenu
    
    Unload Me
End Sub
