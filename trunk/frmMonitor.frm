VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMonitor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Monitor"
   ClientHeight    =   5160
   ClientLeft      =   1920
   ClientTop       =   1035
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
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
      IMEMode         =   3  'DISABLE
      Left            =   5400
      TabIndex        =   13
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin StealthBot.ctlMonitor monConn 
      Left            =   3000
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "&Shutdown"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdDisc 
      Caption         =   "Manually &Disconnect"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Manually Connect"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "R&efresh List From Textfile"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Close"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdRem 
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
      Left            =   5400
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtAdd 
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
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvMonitor 
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8070
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   10079232
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "name"
         Text            =   "Username"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "status"
         Text            =   "Status"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "last"
         Text            =   "Last Check"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5400
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "¯`- StealthBot User         Monitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000012&
      Caption         =   "Offline"
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
      Left            =   720
      TabIndex        =   5
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Status:"
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
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StatusWatch() As Byte
Private Sent() As Byte
Attribute Sent.VB_VarHelpID = -1

Public Sub Shutdown()
  monConn.Disconnect
  Call frmChat.DeconstructMonitor
End Sub

Private Sub cmdConnect_Click()
    Call Connect(True)
End Sub
Public Function Connect(blDisplayError As Boolean) As Boolean
    Dim blSuccess As Boolean
    blSuccess = monConn.Connect(blDisplayError)
    Connect = blSuccess
    If (blSuccess) Then
        cmdConnect.Enabled = False
        cmdDisc.Enabled = True
    End If
End Function


Private Sub cmdRefresh_Click()
    Call Form_Load
End Sub

Private Sub cmdDisc_Click()
    On Error Resume Next
    monConn.Disconnect
    cmdConnect.Enabled = True
    cmdDisc.Enabled = False
End Sub

Private Sub cmdShutdown_Click()
    Call ShutdownMonitor
End Sub

Public Sub ShutdownMonitor()
    monConn.Disconnect
    Call frmChat.DeconstructMonitor
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Me.Icon = frmChat.Icon
    monConn.LoadMonitorConfig
    txtUsername.text = monConn.Username
    txtPassword.text = monConn.Password
    'If Not DisableMonitor Then
    '    monConn.Connect
    'Else
        Call cmdDisc_Click
    'End If
    
    Dim Users As Collection, X As Integer
    With lvMonitor
        .SmallIcons = frmChat.imlIcons
        .icons = frmChat.imlIcons
        .View = lvwReport
        .ListItems.Clear
        Set Users = monConn.getList
        For X = 1 To Users.Count
            .ListItems.Add , Users.Item(X).Username, Users.Item(X).Username, , ICSQUELCH
            .ListItems(.ListItems.Count).ListSubItems.Add 1, "status", "Offline", MONITOR_OFFLINE
            .ListItems(.ListItems.Count).ListSubItems.Add 2, "last", "None"
            .ListItems(.ListItems.Count).Tag = "0"
        Next X
    End With
End Sub

Private Sub cmdDone_Click()
    Me.Hide
End Sub

Private Sub cmdRem_CLick()
    If lvMonitor.ListItems.Count = 1 Then
        MsgBox "You can't remove the last person in the monitor. " & vbNewLine & "If you'd like to stop using the monitor, go to the Settings menu and choose Bot Settings. Select the panel labeled 'Miscellaneous Settings' and check the 'Disable User Monitor' checkbox.", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    If Not (lvMonitor.SelectedItem Is Nothing) Then
        Call monConn.RemoveUser(lvMonitor.SelectedItem.text)
        lvMonitor.ListItems.Remove (lvMonitor.SelectedItem.index)
    End If
End Sub


Public Function RemoveUser(strUser As String) As Boolean
  RemoveUser = False
  Dim usrItem As ListItem
  Set usrItem = lvMonitor.FindItem(strUser)
  If (usrItem Is Nothing) Then Exit Function
  Call monConn.RemoveUser(usrItem.text)
  lvMonitor.ListItems.Remove usrItem.index
  RemoveUser = True
End Function


Sub cmdAdd_Click()
    'On Error Resume Next
    If txtAdd.text <> vbNullString Then
        AddUser txtAdd.text
        txtAdd.text = vbNullString
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then Exit Sub Else Call frmChat.DeconstructMonitor
End Sub

Private Sub monConn_BNETClose()
  Debug.Print "BNET Close"
  lblStatus.Caption = "Offline"
  lblStatus.ToolTipText = "Discnnected"
End Sub

Private Sub monConn_BNETConnect()
  Debug.Print "BNET Connect"
  lblStatus.Caption = "Connecting..."
End Sub

Private Sub monConn_BNETError(ByVal Number As Integer, ByVal description As String)
  Debug.Print "BNET " & Number & " " & description
  lblStatus.Caption = "[BNET] " & Number & ": " & description
End Sub

Private Sub monConn_BNLSClose()
  Debug.Print "BNLS Close"
End Sub

Private Sub monConn_BNLSConnect()
  Debug.Print "BNLS Connet"
End Sub

Private Sub monConn_BNLSError(ByVal Number As Integer, ByVal description As String)
  Debug.Print "BNLS " & Number & " " & description
  lblStatus.Caption = "[BNLS] " & Number & ": " & description
  lblStatus.ToolTipText = "BNLS Hash encountered an error. The server may be down."
End Sub

Private Sub monConn_OnChatJoin(ByVal UniqueName As String)
  Debug.Print "Logged in as " & UniqueName
  lblStatus.Caption = "Connected as " & UniqueName
  lblStatus.ToolTipText = "Logged on an Monitoring Users"
End Sub

Private Sub monConn_OnCreateAccount(ByVal blSucces As Boolean)
  If blSucces Then
    lblStatus.Caption = "Account creation successful, Logging in"
    lblStatus.ToolTipText = vbNullString
  Else
    lblStatus.Caption = "Account creation failed"
    lblStatus.ToolTipText = "This could mean you have invalid charecters is your account (){}&*#@! or space. Or, It could mean that the account is already taken by someone else."
  End If
End Sub

Private Sub monConn_OnLogin(ByVal Success As Boolean)
  Debug.Print "Login " & IIf(Success, "Success", "Failed")
  If (Success) Then
    lblStatus.Caption = "Login successful"
    lblStatus.ToolTipText = vbNullString
  Else
    lblStatus.Caption = "Failed to login to account: Invalid password."
    lblStatus.ToolTipText = "Enter a valid account/password in the text boxes to the right."
  End If
End Sub

Private Sub monConn_OnVersionCheck(ByVal result As Long, PatchFile As String)
  Debug.Print "Version: 0x" & Hex(result)
  If (result = 2) Then
    lblStatus.Caption = "Version Check Passed"
    lblStatus.ToolTipText = "Version check for Diablo Retail passed"
  Else
    lblStatus.Caption = "Version Check Failed: " & PatchFile
    lblStatus.ToolTipText = "Faild to validate Diablo Retail version. BNLS Server is outdated."
  End If
End Sub

Private Sub monConn_UserInfo(user As clsFriendObj)
  Debug.Print "User info: " & user.Username & ": " & user.Status
  Call UpdateList(user)
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
        KeyAscii = 0
    End If
End Sub
Private Sub UpdateList(user As clsFriendObj)
    Dim X As ListItem, Holder As Integer, b As Byte
    
    If InStr(1, user.Product, "starcraft", vbTextCompare) <> 0 Then
        If InStr(1, user.Product, "broodwar", vbTextCompare) <> 0 Then
            Holder = ICSEXP
        ElseIf InStr(1, user.Product, "japanese", vbTextCompare) <> 0 Then
            Holder = ICJSTR
        ElseIf InStr(1, user.Product, "shareware", vbTextCompare) <> 0 Then
            Holder = ICSCSW
        Else
            Holder = ICSTAR
        End If
    ElseIf InStr(1, user.Product, "diablo", vbTextCompare) <> 0 Then
        If InStr(1, user.Product, "ii", vbTextCompare) <> 0 Then
            If InStr(1, user.Product, "lord of destruction", vbTextCompare) <> 0 Then
                Holder = ICD2XP
            Else
                Holder = ICD2DV
            End If
        Else
            If InStr(1, user.Product, "shareware", vbTextCompare) <> 0 Then
                Holder = ICDIABLOSW
            Else
                Holder = ICDIABLO
            End If
        End If
    ElseIf InStr(1, user.Product, "chat", vbTextCompare) <> 0 Then
        Holder = ICCHAT
    ElseIf InStr(1, user.Product, "warcraft", vbTextCompare) <> 0 Then
        If InStr(1, user.Product, "iii", vbTextCompare) <> 0 Then
            Holder = ICWAR3
            If InStr(1, user.Product, "frozen throne", vbTextCompare) <> 0 Then
                Holder = ICWAR3X
            End If
        Else
            Holder = ICW2BN
        End If
    End If
    If Holder = 0 Then Holder = ICUNKNOWN
    Set X = lvMonitor.FindItem(user.Username)
    If Not X Is Nothing Then
        If user.Location = 1 And Not (X.Icon = 1) Then
            StatusOnline user.Username
            X.Icon = 1
        End If
        With lvMonitor.ListItems(X.index)
            On Error Resume Next
            .SmallIcon = Holder
            .ListSubItems.Clear
            If (user.Status = 1) Then
                .ListSubItems.Add , "status", "Online", MONITOR_ONLINE
            Else
                .ListSubItems.Add , "status", "Offline", MONITOR_OFFLINE
            End If
            .ListSubItems.Add , "last", Time
        End With
    End If
End Sub

Function AddUser(ByVal Username As String)
    On Error Resume Next
    AddUser = False
    If (Not monConn.AddUser(Username)) Then Exit Function
    With lvMonitor
        .ListItems.Add , Username, Username, , ICSQUELCH
        'Debug.Print .ListItems(.ListItems.Count).ListSubItems.Count
        .ListItems(.ListItems.Count).ListSubItems.Add 0, "status", "Offline", MONITOR_OFFLINE
        .ListItems(.ListItems.Count).ListSubItems.Add 1, "last", "None"
        .ListItems(.ListItems.Count).Tag = "0"
    End With
    AddUser = True
End Function

Function SetStatusWatch(ByVal Val As Byte, ByVal Username As String) As Byte
    Dim X As ListItem
    
    Set X = lvMonitor.FindItem(Username)
    
    If Not (X Is Nothing) Then
        X.Icon = Val
        SetStatusWatch = 1
    End If
    Exit Function

SetStatusWatch_Error:
    SetStatusWatch = 2
    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure SetStatusWatch of Form frmMonitor"

End Function

Sub StatusOnline(ByVal Username As String)
    frmChat.AddQ Username & " has signed onto Battle.net."
End Sub

Function GetStatusWatch(ByVal Username As String) As Byte
    Dim X As Collection, i As Integer
    
    Set X = monConn.getList
    For i = 1 To X.Count
        If (StrComp(Username, X.Item(i).Username, vbTextCompare) = 0) Then
          GetStatusWatch = X.Item(i).Location
          Exit Function
        End If
    Next i
End Function

Function GetUserStatus(ByVal Username As String) As Integer
    Dim X As ListItem
    
    Set X = lvMonitor.FindItem(Username)
    
    If Not (X Is Nothing) Then
        If X.ListSubItems(1).text = "Online" Then
            GetUserStatus = 1
        Else
            GetUserStatus = 0
        End If
        
        Set X = Nothing
    Else
        GetUserStatus = -1
    End If
End Function

Public Function OnlineUsers() As String
    Dim X As Integer
    Dim tmpBuf As String
    Dim Count As Integer
    For X = 1 To lvMonitor.ListItems.Count
        If (lvMonitor.ListItems(X).ListSubItems(1).text = "Online") Then
            Count = Count + 1
            If (Count Mod 6 = 0) Then tmpBuf = tmpBuf & " [more]" & vbNewLine
            tmpBuf = tmpBuf & lvMonitor.ListItems(X).text & ", "
        End If
    Next X
    tmpBuf = "Online users " & Count & ": " & tmpBuf
    tmpBuf = Left(tmpBuf, Len(tmpBuf) - 2)
    OnlineUsers = tmpBuf
End Function

Function GetFullUserStatus(ByVal Username As String, ByRef Online As Boolean, ByRef LastChecked As String, ByRef LastWhois As String) As Integer
    Dim X As ListItem
    
    Set X = lvMonitor.FindItem(Username)
    
    If Not (X Is Nothing) Then
        LastWhois = X.Tag
        Online = (X.ListSubItems(1).text = "Online")
        LastChecked = X.ListSubItems(2).text
        Set X = Nothing
        GetFullUserStatus = 0
    Else
        Online = False
        LastChecked = vbNullString
        LastWhois = vbNullString
        GetFullUserStatus = 1
    End If
End Function

Private Sub txtPassword_Change()
  Call WriteINI("Monitor", "Password", txtPassword.text)
  monConn.Password = txtPassword.text
End Sub

Private Sub txtUsername_Change()
  Call WriteINI("Monitor", "Username", txtUsername.text)
  monConn.Username = txtUsername.text
End Sub
