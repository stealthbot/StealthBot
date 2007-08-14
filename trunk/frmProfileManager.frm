VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProfileManager 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Manager"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPath 
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
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Path"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   4080
         X2              =   120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Server"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Product"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Username"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboProfile 
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
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Choose a profile to view, or click ""New"" below."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmProfileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Saved As Boolean
Private CurrentIndex As Integer

Private Sub cmdBrowse_Click()
    With cdl
        .InitDir = App.Path
        .ShowSave
    
        If Len(.FileName) > 0 Then
            txtPath.text = .FileName
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    If Not Saved Then
        If MsgBox("You haven't saved your changes to this profile. Do you still want to exit?", vbExclamation + vbYesNo, "Profile Manager") = vbYes Then
            SaveAllProfiles
            Unload Me
        End If
    End If
End Sub

Private Sub cmdNew_Click()
    SaveCurrentProfile
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    Set colProfiles = New Collection
End Sub

Sub SaveCurrentProfile()
    If CurrentIndex > colProfiles.Count Then
        
    Else
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colProfiles = Nothing
End Sub

Sub SaveAllProfiles()
    Dim f As Integer, i As Integer
    
    If colProfiles.Count > 0 Then
        Open App.Path & "\profilelist.txt" For Output As #f
            For i = 1 To colProfiles.Count
                Print #f, CStr(colProfiles.Item(i))
            Next i
        Close #f
    End If
End Sub
