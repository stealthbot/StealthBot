VERSION 5.00
Begin VB.Form frmUserManager 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Userlist Manager"
   ClientHeight    =   4920
   ClientLeft      =   2685
   ClientTop       =   1620
   ClientWidth     =   5280
   ForeColor       =   &H00000000&
   Icon            =   "frmUserManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFlags 
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
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Leave these blank if you don't know what they are."
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ListBox List 
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
      Height          =   3570
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5055
   End
   Begin VB.TextBox txtAccess 
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
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtUsername 
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
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Flags (optional)"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblRem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[&Remove Selected User]"
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
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[&Add User]"
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
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Access"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Username"
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
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblDone 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[&Done]"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmUserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmUserManager
' DateTime  : 4/5/2007 16:45
' Author    : Andy Trevino (andytrevino@gmail.com)
' Purpose   : Updated to revamp VERY old (circa early 2004) code and provide support
'               for new database features introduced in 2.7
'---------------------------------------------------------------------------------------
Option Explicit

Private colOtherInfo As Collection

'  0        1       2       3       4       5       6
' username access flags addedby addedon modifiedby modifiedon
Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    Dim i As Integer
    Dim lbString As String, oiString As String
    
    Set colOtherInfo = New Collection
    
    For i = 0 To UBound(DB)
        lbString = DB(i).Username & " " & DB(i).Access & " " & DB(i).Flags
        oiString = DB(i).AddedBy & " " & IIf(DB(i).AddedOn > 0, DateCleanup(DB(i).AddedOn), "%") & " " & DB(i).ModifiedBy & " " & IIf(DB(i).ModifiedOn > 0, DateCleanup(DB(i).ModifiedOn), "%")
        
        List.AddItem lbString
        colOtherInfo.Add oiString
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colOtherInfo = Nothing
End Sub

Private Sub lblAdd_Click()
    Dim s As String
    
    If Len(txtAccess.text) > 0 Or Len(txtFlags.text) > 0 Then
        s = txtUsername.text
        
        If Len(txtAccess.text) > 0 Then s = s & Space(1) & txtAccess.text
        If Len(txtFlags.text) > 0 Then s = s & Space(1) & txtFlags.text
        
        List.AddItem (s)
        colOtherInfo.Add "(console) " & DateCleanup(Now) & " (console) " & DateCleanup(Now)
        
        txtUsername.text = vbNullString
        txtAccess.text = vbNullString
        txtFlags.text = vbNullString
    Else
        MsgBox "Please enter an access level (0-999) or flags (A-Z) for this person!", vbOKOnly + vbInformation, "StealthBot Userlist Manager"
    End If
End Sub

Private Sub lblDone_Click()
    Dim i As Integer
    Dim f As Integer
    Dim sPath As String
    Dim s() As String

    On Error GoTo lblDone_Click_Error

    f = FreeFile
    sPath = GetFilePath("users.txt")
    
    Open sPath For Output As #f
    
    For i = 0 To (List.ListCount - 1)
        s() = Split(List.List(i), " ")
        Print #f, s(0) & " "; ' Username
        
        If (StrictIsNumeric(s(1))) Then
            Print #f, s(1) & " "; 's(1) is access
            
            If UBound(s) > 1 Then
                'we have flags as well
                Print #f, s(2) & " ";
            Else
                Print #f, "% ";
            End If
        Else
            's(1) is flags. process s(2)
            If UBound(s) > 1 Then
                Print #f, s(2) & " ";
            Else
                Print #f, "% ";
            End If
            
            'then s(1)
            Print #f, s(1) & " ";
        End If
        
        Print #f, colOtherInfo.Item(i + 1)
    Next i
    
    Close #f
    
    Call LoadDatabase
    Unload Me

lblDone_Click_Exit:
    On Error GoTo 0
    Exit Sub

lblDone_Click_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDone_Click of Form frmUserManager"
    Resume lblDone_Click_Exit
End Sub

Private Sub lblRem_Click()
    On Error Resume Next
    colOtherInfo.Remove List.ListIndex + 1
    List.RemoveItem (List.ListIndex)
End Sub

Private Sub List_DblClick()
    Dim s() As String
    
    If List.ListIndex > -1 Then
        s() = Split(List.List(List.ListIndex), " ")
        
        txtUsername.text = s(0)
        If UBound(s) > 0 Then
            If StrictIsNumeric(s(1)) Then
                txtAccess.text = s(1)
            Else
                txtFlags.text = s(1)
            End If
        End If
        
        If UBound(s) > 1 Then
            If StrictIsNumeric(s(2)) Then
                txtAccess.text = s(2)
            Else
                txtFlags.text = s(2)
            End If
        End If
        
        colOtherInfo.Remove List.ListIndex + 1
        List.RemoveItem List.ListIndex
    End If
End Sub

Private Sub txtAccess_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(txtAccess.text) > 0 Then
        Call lblAdd_Click
    End If
End Sub

Private Sub txtFlags_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(txtFlags.text) > 0 Then
        Call lblAdd_Click
    End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtFlags.text) > 0 Or Len(txtAccess.text) > 0 Then Call lblAdd_Click
    End If
End Sub

