VERSION 5.00
Begin VB.Form frmRenameProfile 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename Profile"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Enter the name you want to rename the profile to."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRenameProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const OBJECT_NAME = "frmRenameProfile"

Private previousProfileName As String
Private previousProfileIndex As String

Private Sub Form_Load()
On Error GoTo ERROR_HANDLER:

    Me.Icon = frmLauncher.Icon
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Load"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ERROR_HANDLER:

    Unload Me
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdCancel_Click"
End Sub

Private Sub cmdOk_Click()
On Error GoTo ERROR_HANDLER:
    
    Dim i As Integer
    Dim Text As String
    Dim Char As String * 1
    Dim originalPath As String
    Dim destinationPath As String
    
    Text = txtName.Text
    
    If (LenB(Text) = 0) Then
        MsgBox "You must enter a profile name!", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To Len(INVALID_CHARS)
        Char = Mid$(INVALID_CHARS, i, 1)
        If (InStr(1, Text, Char, vbBinaryCompare) > 0) Then
            MsgBox "Invalid character in profile name: " & Char, vbExclamation
            Exit Sub
        End If
    Next i
    
    If (ProfileExists(Text)) Then
        MsgBox "That profile already exists!"
        Exit Sub
    End If
    
    originalPath = StringFormat("{0}\StealthBot\{1}", ReplaceEnvironmentVars("%APPDATA%"), previousProfileName)
    destinationPath = StringFormat("{0}\StealthBot\{1}", ReplaceEnvironmentVars("%APPDATA%"), Text)
    
    If (CopyFolder(originalPath, destinationPath)) Then
        KillFolder originalPath
        frmLauncher.renameProfileInList Text, previousProfileIndex
    End If
    
    Unload Me

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdOk_Click"
End Sub

Public Sub setOriginalProfile(ByVal profileName As String, ByVal profileIndex As Integer)
On Error GoTo ERROR_HANDLER:

    previousProfileName = profileName
    previousProfileIndex = profileIndex
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "setOriginalProfile"
End Sub

