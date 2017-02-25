VERSION 5.00
Begin VB.Form frmAccountManager 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Manager"
   ClientHeight    =   3990
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
   ScaleHeight     =   3990
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtField3 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "&Save and Connect"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtField2 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtField1 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox cboMode 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmAccountManager.frx":0000
      Left            =   240
      List            =   "frmAccountManager.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblField3 
      BackColor       =   &H00000000&
      Caption         =   "Field 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblField2 
      BackColor       =   &H00000000&
      Caption         =   "Field 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblField1 
      BackColor       =   &H00000000&
      Caption         =   "Field 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H00000000&
      Caption         =   "Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Choose what to do after connecting."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblModeDetail 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Log on to this account:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmAccountManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const M_LOGON As Byte = 0
Private Const M_CREAT As Byte = 1
Private Const M_CHPWD As Byte = 2
Private Const M_RSPWD As Byte = 3
Private Const M_CHREG As Byte = 4
Private Const M_EMREG As Byte = 5

Private Username As String
Private Password As String
Private Email As String
Private NPassword1 As String
Private NPassword2 As String
Private NEmail1 As String
Private NEmail2 As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnConnect_Click()
    SaveFieldsAndConnect
End Sub

Private Sub SaveFieldsAndConnect()
    Dim ErrorMsg As String
    
    ' save values
    txtUsername_Change
    txtField1_Change
    txtField2_Change
    txtField3_Change

    ' check that the "verify" fields match, if applicable
    Select Case cboMode.ListIndex
        Case M_CREAT, M_CHPWD
            If (LenB(NPassword1) > 0) And (LenB(NPassword2) > 0) And (StrComp(NPassword1, NPassword2, vbBinaryCompare) <> 0) Then
                ErrorMsg = "Passwords do not match! Please make sure you did not make a mistake."
            End If
        Case M_CHREG, M_EMREG
            If (LenB(NEmail1) > 0) And (LenB(NEmail2) > 0) And (StrComp(NEmail1, NEmail2, vbTextCompare) <> 0) Then
                ErrorMsg = "Email addresses do not match! Please make sure you did not make a mistake."
            End If
    End Select

    ' if no no-match errors, check that there's data in the fields
    If LenB(ErrorMsg) = 0 Then
        Dim MissingFields As String
        Dim MissingFieldCount As Integer
        Dim Comma As String
        Dim ActionName As String

        If (LenB(Username) = 0) Then
            MissingFieldCount = MissingFieldCount + 1
            MissingFields = MissingFields & Comma & "Username"
            Comma = ", "
        End If

        Select Case cboMode.ListIndex
            Case M_LOGON
                ActionName = "log on"
                If (LenB(Password) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Password"
                    Comma = ", "
                End If
            Case M_CREAT
                ActionName = "create an account"
                If (LenB(NPassword1) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Password"
                    Comma = ", "
                End If
                If (LenB(NPassword2) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Retype Password"
                    Comma = ", "
                End If
            Case M_CHPWD
                ActionName = "change the password"
                If (LenB(Password) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Old Password"
                    Comma = ", "
                End If
                If (LenB(NPassword1) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "New Password"
                    Comma = ", "
                End If
                If (LenB(NPassword2) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Retype New Password"
                    Comma = ", "
                End If
            Case M_RSPWD
                ActionName = "reset the password"
                If (LenB(Password) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Registered Email"
                    Comma = ", "
                End If
            Case M_CHREG
                ActionName = "change the registered email"
                If (LenB(Email) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Old Registered Email"
                    Comma = ", "
                End If
                If (LenB(NEmail1) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "New Email"
                    Comma = ", "
                End If
                If (LenB(NEmail2) = 0) Then
                    MissingFieldCount = MissingFieldCount + 1
                    MissingFields = MissingFields & Comma & "Retype New Email"
                    Comma = ", "
                End If
        End Select

        If MissingFieldCount > 0 Then
            ErrorMsg = StringFormat("The {0} field{1} {2} required to {3}.", MissingFields, _
                    IIf(MissingFieldCount = 1, vbNullString, "s"), _
                    IIf(MissingFieldCount = 1, "is", "are"), _
                    ActionName)
        End If
    End If

    ' display error and abort
    If LenB(ErrorMsg) > 0 Then
        MsgBox "Wait! You can't do that yet! " & ErrorMsg, vbOKOnly Or vbExclamation, "StealthBot"
        Exit Sub
    End If
    
    ' save fields
    Config.Username = Username
    Select Case cboMode.ListIndex
        Case M_LOGON
            Config.Password = Password
            Config.AccountMode = ACCOUNT_MODE_LOGON
        Case M_CREAT
            Config.Password = NPassword1
            Config.AccountMode = ACCOUNT_MODE_CREAT
        Case M_CHPWD
            Config.Password = Password
            Config.NewPassword = NPassword1
            Config.AccountMode = ACCOUNT_MODE_CHPWD
        Case M_RSPWD
            Config.RegisterEmailDefault = Email
            Config.AccountMode = ACCOUNT_MODE_RSPWD
        Case M_CHREG
            Config.RegisterEmailDefault = Email
            Config.RegisterEmailChange = NEmail1
            Config.AccountMode = ACCOUNT_MODE_CHREG
    End Select
    Config.AutoAccountAction = False
    Config.Save
    
    If g_Connected And ds.AccountEntry Then
        Call modBNCS.DoAccountAction
    Else
        Call frmChat.DoConnect
    End If
End Sub

Private Sub SetFieldMode(Label As Label, TextBox As TextBox, ByVal Visible As Boolean, _
        Optional ByVal Caption As String = vbNullString, Optional ByVal PasswordChar As String = vbNullString)
    Label.Visible = Visible
    Label.Caption = Caption
    TextBox.Visible = Visible
    TextBox.PasswordChar = PasswordChar
End Sub

Private Sub cboMode_Change()
    ' set up fields based on mode
    Select Case cboMode.ListIndex
        Case M_LOGON
            lblModeDetail.Caption = "Log on to this account:"
            Call SetFieldMode(lblField1, txtField1, True, "Password", "*")
            Call SetFieldMode(lblField2, txtField2, False)
            Call SetFieldMode(lblField3, txtField3, False)
        Case M_CREAT
            lblModeDetail.Caption = "Create this account:"
            Call SetFieldMode(lblField1, txtField1, True, "Password", "*")
            Call SetFieldMode(lblField2, txtField2, True, "Retype Password", "*")
            Call SetFieldMode(lblField3, txtField3, False)
        Case M_CHPWD
            lblModeDetail.Caption = "Change the password for this account:"
            Call SetFieldMode(lblField1, txtField1, True, "Old Password", "*")
            Call SetFieldMode(lblField2, txtField2, True, "New Password", "*")
            Call SetFieldMode(lblField3, txtField3, True, "Retype New Password", "*")
        Case M_RSPWD
            lblModeDetail.Caption = "Request a password reset for this account:"
            Call SetFieldMode(lblField1, txtField1, True, "Registered Email", vbNullString)
            Call SetFieldMode(lblField2, txtField2, False)
            Call SetFieldMode(lblField3, txtField3, False)
        Case M_CHREG
            lblModeDetail.Caption = "Change the registered email address:"
            Call SetFieldMode(lblField1, txtField1, True, "Old Registered Email", vbNullString)
            Call SetFieldMode(lblField2, txtField2, True, "New Email", vbNullString)
            Call SetFieldMode(lblField3, txtField3, True, "Retype New Email", vbNullString)
    End Select
    
    ' fill in fields
    txtUsername.Text = Username
    Select Case cboMode.ListIndex
        Case M_LOGON, M_CHPWD: txtField1.Text = Password
        Case M_RSPWD, M_CHREG: txtField1.Text = Email
        Case M_CREAT:          txtField1.Text = vbNullString
    End Select
    txtField2.Text = vbNullString
    txtField3.Text = vbNullString
End Sub

Private Sub cboMode_Click()
    cboMode_Change
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon

    ' populate modes
    cboMode.AddItem "Account Logon (default)", M_LOGON
    cboMode.AddItem "Create Account", M_CREAT
    cboMode.AddItem "Change Password", M_CHPWD
    cboMode.AddItem "Reset Password", M_RSPWD
    cboMode.AddItem "Change Registered Email", M_CHREG
End Sub

Public Sub LeftAccountEntryMode()
    Call ShowMode(vbNullString)
End Sub

Public Sub ShowMode(ByVal Mode As String, Optional ByVal ControlIndex As Integer = 1)
    On Error Resume Next
    Show
    On Error GoTo 0

    Username = Config.Username
    Password = Config.Password
    Email = Config.RegisterEmailDefault

    Select Case UCase$(Mode)
        Case vbNullString:          ' no change
        Case ACCOUNT_MODE_CREAT:    cboMode.ListIndex = M_CREAT
        Case ACCOUNT_MODE_CHPWD:    cboMode.ListIndex = M_CHPWD
        Case ACCOUNT_MODE_RSPWD:    cboMode.ListIndex = M_RSPWD
        Case ACCOUNT_MODE_CHREG:    cboMode.ListIndex = M_CHREG
        Case Else:                  cboMode.ListIndex = M_LOGON
    End Select

    Call cboMode_Change

    If frmChat.sckBNet.State <> sckConnected Then
        lblMode.Caption = "Choose what to do after connecting:"
        btnClose.Caption = "&Close"
    ElseIf ds.AccountEntry Then
        lblMode.Caption = "Choose what to do:"
        btnClose.Caption = "Dis&connect"
    Else
        lblMode.Caption = "Choose what to do after reconnecting:"
        btnClose.Caption = "&Close"
    End If

    On Error Resume Next
    Select Case ControlIndex
        Case 0: cboMode.SetFocus
        Case 1: txtUsername.SetFocus
        Case 2: txtField1.SetFocus
        Case 3: txtField2.SetFocus
        Case 4: txtField3.SetFocus
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (frmChat.sckBNet.State = sckConnected) And (ds.AccountEntry) Then
        frmChat.DoDisconnect
    End If
End Sub

Private Sub txtField1_Change()
    Select Case cboMode.ListIndex
        Case M_LOGON, M_CHPWD:  Password = txtField1.Text
        Case M_RSPWD, M_CHREG:  Email = txtField1.Text
        Case M_CREAT:           NPassword1 = txtField1.Text
    End Select
    If frmChat.sckBNet.State = sckConnected And ds.AccountEntry Then
        btnClose.Caption = "Dis&connect"
    Else
        btnClose.Caption = "&Cancel"
    End If
End Sub

Private Sub txtField2_Change()
    Select Case cboMode.ListIndex
        Case M_CHPWD:   NPassword1 = txtField2.Text
        Case M_CREAT:   NPassword2 = txtField2.Text
        Case M_CHREG:   NEmail1 = txtField2.Text
    End Select
    If frmChat.sckBNet.State = sckConnected And ds.AccountEntry Then
        btnClose.Caption = "Dis&connect"
    Else
        btnClose.Caption = "&Cancel"
    End If
End Sub

Private Sub txtField3_Change()
    Select Case cboMode.ListIndex
        Case M_CHPWD:   NPassword2 = txtField3.Text
        Case M_CHREG:   NEmail2 = txtField3.Text
    End Select
    If frmChat.sckBNet.State = sckConnected And ds.AccountEntry Then
        btnClose.Caption = "Dis&connect"
    Else
        btnClose.Caption = "&Cancel"
    End If
End Sub

Private Sub txtUsername_Change()
    Username = txtUsername.Text
    If frmChat.sckBNet.State = sckConnected And ds.AccountEntry Then
        btnClose.Caption = "Dis&connect"
    Else
        btnClose.Caption = "&Cancel"
    End If
End Sub
