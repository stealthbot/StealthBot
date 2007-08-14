VERSION 5.00
Begin VB.Form frmManageKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage CDKeys"
   ClientHeight    =   2685
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete Selected"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Selected"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtActiveKey 
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
      TabIndex        =   1
      Top             =   2280
      Width           =   3375
   End
   Begin VB.ListBox lstKeys 
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
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmManageKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_FINDSTRINGEXACT = &H1A2
Private NumCDKeysOnOpen As Long

Private Sub cmdAdd_Click()
    Dim s As String
    
    s = txtActiveKey.text
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Trim$(s)
    
    AddUnique lstKeys, s
    
    txtActiveKey.text = vbNullString
    txtActiveKey.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If lstKeys.ListIndex > -1 Then
        lstKeys.RemoveItem lstKeys.ListIndex
    End If
End Sub

Private Sub cmdDone_Click()
    Call Local_WriteCDKeys
    
    If Not (frmChat.SettingsForm Is Nothing) Then
        frmChat.SettingsForm.cboCDKey.Clear
        Call LoadCDKeys(frmChat.SettingsForm.cboCDKey)
    End If
    
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    If lstKeys.ListIndex > -1 Then
        txtActiveKey.text = lstKeys.List(lstKeys.ListIndex)
        lstKeys.RemoveItem lstKeys.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    Call Local_LoadCDKeys
    
    NumCDKeysOnOpen = lstKeys.ListCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmChat.SettingsForm.Show
    frmChat.SettingsForm.SetFocus
End Sub

Private Sub lstKeys_Click()
    cmdEdit.Enabled = (lstKeys.ListIndex > -1)
End Sub

Private Sub txtActiveKey_Change()
    cmdAdd.Enabled = (Len(txtActiveKey.text) > 0)
End Sub

Private Sub Local_WriteCDKeys()
    Dim i As Integer

    WriteINI "StoredKeys", "Count", lstKeys.ListCount + 1
    
    For i = 0 To NumCDKeysOnOpen
        If i <= lstKeys.ListCount Then
            If Len(lstKeys.List(i)) > 0 Then
                WriteINI "StoredKeys", "Key" & (i + 1), lstKeys.List(i)
            Else
                WriteINI "StoredKeys", "Key" & (i + 1), ""
            End If
        Else
            WriteINI "StoredKeys", "Key" & (i + 1), ""
        End If
    Next i
    
    If i < lstKeys.ListCount Then
        If Len(lstKeys.List(i)) > 0 Then
            WriteINI "StoredKeys", "Key" & (i + 1), lstKeys.List(i)
        End If
    End If
End Sub

'// thanks Grok[vL]
'//     http://forum.valhallalegends.com/phpbbs/index.php?board=31;action=display;threadid=5872;start=msg50449#msg50449
'// what would we do without Grok?
Private Sub AddUnique(ByVal LB As ListBox, ByVal strNewValue As String)
    Dim errCode As Integer
    If Len(strNewValue) = 0 Then Exit Sub      'LB_FINDSTRINGEXACT does not detect blank entries
    errCode = SendMessageByString(LB.hWnd, LB_FINDSTRINGEXACT, -1, strNewValue)
    
    If errCode = -1 Then
        LB.AddItem strNewValue
    End If
End Sub

Private Sub Local_LoadCDKeys()
    Dim Count As Integer
    Count = Val(ReadCFG("StoredKeys", "Count"))
    
    Dim sKey As String
    
    If Count > 0 Then
        For Count = 1 To Count
            sKey = ReadCFG("StoredKeys", "Key" & Count)
            If Len(sKey) > 0 Then lstKeys.AddItem sKey
        Next Count
    End If
End Sub
