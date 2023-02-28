VERSION 5.00
Begin VB.Form frmNameDialog 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Title"
   ClientHeight    =   1095
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtName 
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
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "[message]"
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
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.Line line 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6360
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmNameDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' this dialog allows you to input a name for your profile

Private Const OBJECT_NAME As String = "frmNameDialog"
Private Const INVALID_CHARS As String = "\/*?"":<>|"

Private m_profileOption As ProfileOption
Private previousProfileName As String
Private previousProfileIndex As Integer

Private Sub Form_Load()
On Error GoTo ERROR_HANDLER
    txtName.Text = vbNullString
    Me.Icon = frmLauncher.Icon
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Load"
End Sub

Private Sub cmdOk_Click()
On Error GoTo ERROR_HANDLER
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
    
    Select Case m_profileOption
        Case CREATE_PROFILE
            If (CreateProfile(Text)) Then
                frmLauncher.ListProfile Text
            Else
                'MsgBox "Failed to create profile!"
            End If
        Case RENAME_PROFILE
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
    End Select
    
    Unload Me
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdOK_Click"
    
    Resume Next
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ERROR_HANDLER
    Unload Me
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdCancel_Click"
End Sub

Public Sub setWindowData(ByVal title As String, ByVal message As String, ByVal po As ProfileOption)
On Error GoTo ERROR_HANDLER
    Me.Caption = title
    lblText.Caption = message
    m_profileOption = po

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "setWindowData"
End Sub

Public Sub setOldProfileInfo(ByVal profileName As String, ByVal profileIndex As Integer)
On Error GoTo ERROR_HANDLER
    previousProfileName = profileName
    previousProfileIndex = profileIndex

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "setOldProfileInfo"
End Sub
