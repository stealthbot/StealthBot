VERSION 5.00
Begin VB.Form frmCatch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catch Phrases"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFlashOnCaughtPhrase 
      BackColor       =   &H00000000&
      Caption         =   "Flash window on caught phrases"
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
      Left            =   600
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
   End
   Begin VB.ListBox lbCatch 
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
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
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
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
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
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtModify 
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
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton cmdOutRem 
      Caption         =   "&Remove Selected Item"
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
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
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
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCatch.frx":0000
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
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "-- StealthBot Catch Phrases --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmCatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFlashOnCaughtPhrase_Click()
    If Config.FlashOnCatchPhrases <> CBool(chkFlashOnCaughtPhrase.Value) Then
        Config.FlashOnCatchPhrases = CBool(chkFlashOnCaughtPhrase.Value)
        Call Config.Save
    End If
End Sub

Private Sub cmdDone_Click()
    Dim i As Integer, f As Integer
    ReDim Preserve Catch(0)
    If lbCatch.ListCount < 0 Then
        Unload Me
        Exit Sub
    End If
    
    f = FreeFile
    Open GetFilePath(FILE_CATCH_PHRASES) For Output As #f
    
    For i = 0 To lbCatch.ListCount
        Catch(i) = lbCatch.List(i)
        Print #f, lbCatch.List(i)
        If i <> lbCatch.ListCount Then ReDim Preserve Catch(0 To UBound(Catch) + 1)
    Next i
    
    Close #f
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    If lbCatch.ListIndex >= 0 Then
        txtModify.Text = lbCatch.Text
        lbCatch.RemoveItem lbCatch.ListIndex
    End If
End Sub

Private Sub cmdOutAdd_Click()
    If txtModify.Text <> vbNullString Then
        lbCatch.AddItem txtModify.Text
        txtModify.Text = vbNullString
    End If
End Sub

Private Sub cmdOutRem_Click()
    If lbCatch.ListIndex >= 0 Then
        lbCatch.RemoveItem lbCatch.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    If Config.FlashOnCatchPhrases Then
        chkFlashOnCaughtPhrase.Value = vbChecked
    End If
    
    Dim i As Integer
    For i = LBound(Catch) To UBound(Catch)
        If Catch(i) <> vbNullString Then
            lbCatch.AddItem Catch(i)
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call cmdDone_Click
End Sub
