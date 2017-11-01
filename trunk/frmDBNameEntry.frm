VERSION 5.00
Begin VB.Form frmDBNameEntry 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Name Entry"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      MaxLength       =   30
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblEntry 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Choose a name for the %s entry."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   440
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmDBNameEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmDBManager.m_EntryName = vbNullString

    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmDBManager.m_EntryName = txtEntry.Text
    
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = StringFormat("New Entry - {0} Name", frmDBManager.m_EntryType)
    
    With txtEntry
        If LenB(frmDBManager.m_EntryName) = 0 Then
            lblEntry.Caption = StringFormat("Choose the name for this new {0} entry.", frmDBManager.m_EntryType)
            .Text = vbNullString
        Else
            lblEntry.Caption = StringFormat("Rename this {0} entry.", frmDBManager.m_EntryType)
            .Text = frmDBManager.m_EntryName
            .SelStart = 0
            .SelLength = Len(frmDBManager.m_EntryName)
        End If
        
        If StrComp(frmDBManager.m_EntryType, "Clan", vbTextCompare) = 0 Then
            .MaxLength = 4
        Else
            .MaxLength = 30
        End If
         
        cmdOK.Enabled = CanSave()
    End With
End Sub

Private Sub txtEntry_Change()
    With txtEntry
        cmdOK.Enabled = CanSave()
    End With
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdOK_Click
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Function CanSave() As Boolean
    With txtEntry
        CanSave = True
        CanSave = CanSave And (LenB(.Text) > 0)
        CanSave = CanSave And (Len(.Text) <= .MaxLength)
        CanSave = CanSave And (StrComp(.Text, "%", vbBinaryCompare) <> 0)
        CanSave = CanSave And (InStr(1, .Text, " ", vbBinaryCompare) = 0)
        CanSave = CanSave And (InStr(1, .Text, ",", vbBinaryCompare) = 0)
        CanSave = CanSave And (InStr(1, .Text, Chr$(34), vbBinaryCompare) = 0)
        
        ' Groups have slightly stricter names
        If StrComp(frmDBManager.m_EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0 Then
            CanSave = CanSave And (StrComp(Left(.Text, 1), "+", vbBinaryCompare) <> 0)
            CanSave = CanSave And (StrComp(Left(.Text, 1), "-", vbBinaryCompare) <> 0)
        End If
    End With
End Function

