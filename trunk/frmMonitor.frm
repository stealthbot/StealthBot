VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonitor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Monitor"
   ClientHeight    =   4800
   ClientLeft      =   390
   ClientTop       =   525
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin StealthBot.ctlMonitor monConn 
      Left            =   5640
      Top             =   3000
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
      Top             =   4080
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
      Top             =   2520
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
      Top             =   2160
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
   Begin VB.Timer tmrDelay 
      Left            =   5640
      Top             =   2520
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
      Top             =   3720
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
   Begin MSComctlLib.ListView lvMonitor 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7435
      View            =   3
      Arrange         =   1
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
      Top             =   4440
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
      Top             =   4440
      Width           =   495
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strUsers() As String
Private LastCheck As Integer
Private StatusWatch() As Byte
Private Sent() As Byte
Attribute Sent.VB_VarHelpID = -1

Private Sub cmdConnect_Click()
    monConn.Connect
    cmdConnect.Enabled = False
    cmdDisc.Enabled = True
End Sub

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
    monConn.Disconnect
    Call frmChat.DeconstructMonitor
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Me.Icon = frmChat.Icon
    If Not DisableMonitor Then
      monConn.LoadMonitorConfig
      monConn.Connect
    Else
        Call cmdDisc_Click
    End If
    
    With lvMonitor
        .SmallIcons = frmChat.imlIcons
        .Icons = frmChat.imlIcons
        .View = lvwReport
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
        Call RemUser(lvMonitor.SelectedItem.Index)
    End If
End Sub
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
End Sub

Private Sub monConn_BNETConnect()
  Debug.Print "BNET Connect"
End Sub

Private Sub monConn_BNETError(ByVal Number As Integer, ByVal Description As String)
  Debug.Print "BNET " & Number & " " & Description
End Sub

Private Sub monConn_BNLSClose()
  Debug.Print "BNLS Close"
End Sub

Private Sub monConn_BNLSConnect()
  Debug.Print "BNLS Connet"
End Sub

Private Sub monConn_BNLSError(ByVal Number As Integer, ByVal Description As String)
  Debug.Print "BNLS " & Number & " " & Description
End Sub

Private Sub monConn_OnChatJoin(ByVal UniqueName As String)
  Debug.Print "Logged in as " & UniqueName
End Sub

Private Sub monConn_OnLogin(ByVal Success As Boolean)
  Debug.Print "Login " & IIf(Success, "Success", "Failed")
End Sub

Private Sub monConn_OnVersionCheck(ByVal result As Long, PatchFile As String)
  Debug.Print "Version: 0x" & Hex(result)
End Sub

Private Sub monConn_UserInfo(user As clsFriend)
  Debug.Print user.Username & ": " & user.Status
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
        KeyAscii = 0
    End If
End Sub
Private Sub UpdateList(ByVal Msg As String, Optional Disable As Byte)
    Dim X As ListItem, Holder As Integer, b As Byte
    
    If Disable = 1 Then
        If LastCheck = 0 Then
            If Len(strUsers(0)) > 0 Then
                Set X = lvMonitor.FindItem(strUsers(0))
                Sent(0) = 0
                b = 1
            End If
        Else
            If Len(strUsers(LastCheck - 1)) > 0 Then
                Set X = lvMonitor.FindItem(strUsers(LastCheck - 1))
                Sent(LastCheck - 1) = 0
                b = 1
            End If
        End If
        
        If b = 1 And (Not (X Is Nothing)) Then
            With lvMonitor
                .ListItems(X.Index).SmallIcon = ICSQUELCH
                .ListItems(X.Index).ListSubItems.Clear
                .ListItems(X.Index).ListSubItems.Add , "status", "Offline", MONITOR_OFFLINE
                .ListItems(X.Index).ListSubItems.Add , "last", Time
            End With
        End If
    Else
        
        If LastCheck = 0 Then
            If GetStatusWatch(strUsers(0)) = 1 And Not Sent(0) = 1 Then
                StatusOnline strUsers(0)
                Sent(0) = 1
            End If
        Else
            If GetStatusWatch(strUsers(LastCheck - 1)) = 1 And Not (Sent(LastCheck - 1) = 1) Then
                StatusOnline strUsers(LastCheck - 1)
                Sent(LastCheck - 1) = 1
            End If
        End If
        
        Msg = LCase(Right(Msg, Len(Msg) - InStr(1, Msg, "using", vbTextCompare)))
        
        If InStr(1, Msg, "starcraft", vbTextCompare) <> 0 Then
            If InStr(1, Msg, "broodwar", vbTextCompare) <> 0 Then
                Holder = ICSEXP
            ElseIf InStr(1, Msg, "japanese", vbTextCompare) <> 0 Then
                Holder = ICJSTR
            ElseIf InStr(1, Msg, "shareware", vbTextCompare) <> 0 Then
                Holder = ICSCSW
            Else
                Holder = ICSTAR
            End If
        ElseIf InStr(1, Msg, "diablo", vbTextCompare) <> 0 Then
            If InStr(1, Msg, "ii", vbTextCompare) <> 0 Then
                If InStr(1, Msg, "lord of destruction", vbTextCompare) <> 0 Then
                    Holder = ICD2XP
                Else
                    Holder = ICD2DV
                End If
            Else
                If InStr(1, Msg, "shareware", vbTextCompare) <> 0 Then
                    Holder = ICDIABLOSW
                Else
                    Holder = ICDIABLO
                End If
            End If
        ElseIf InStr(1, Msg, "chat", vbTextCompare) <> 0 Then
            Holder = ICCHAT
        ElseIf InStr(1, Msg, "warcraft", vbTextCompare) <> 0 Then
            If InStr(1, Msg, "iii", vbTextCompare) <> 0 Then
                Holder = ICWAR3
                If InStr(1, Msg, "frozen throne", vbTextCompare) <> 0 Then
                    Holder = ICWAR3X
                End If
            Else
                Holder = ICW2BN
            End If
        End If

        If Holder = 0 Then Holder = ICUNKNOWN
        
        If LastCheck <> 0 Then
            Set X = lvMonitor.FindItem(strUsers(LastCheck - 1))
        Else
            Set X = lvMonitor.FindItem(strUsers(0))
        End If
        
        If Not X Is Nothing Then
            With lvMonitor.ListItems(X.Index)
                On Error Resume Next
                .Tag = Msg
                .SmallIcon = Holder
                .ListSubItems.Clear
                .ListSubItems.Add , "status", "Online", MONITOR_ONLINE
                .ListSubItems.Add , "last", Time
            End With
        End If
    End If
End Sub

Sub AddUser(ByVal Username As String)
    On Error Resume Next
    If (Not monConn.AddAccount(Username)) Then Exit Sub
    With lvMonitor
        .ListItems.Add , Username, Username, , ICSQUELCH
        'Debug.Print .ListItems(.ListItems.Count).ListSubItems.Count
        .ListItems(.ListItems.Count).ListSubItems.Add 0, "status", "Offline", MONITOR_OFFLINE
        .ListItems(.ListItems.Count).ListSubItems.Add 1, "last", "None"
        .ListItems(.ListItems.Count).Tag = "0"
    End With
End Sub

Function SetStatusWatch(ByVal Val As Byte, ByVal Username As String) As Byte
    Dim i As Integer
        
    On Error GoTo SetStatusWatch_Error

    For i = 0 To UBound(strUsers)
        Debug.Print "Comparing " & strUsers(i) & " to " & Username
        
        If StrComp(strUsers(i), Username, vbTextCompare) = 0 Then
            StatusWatch(i) = Val
            SetStatusWatch = 1
            
            Exit Function
        End If
    Next i

    On Error GoTo 0
    Exit Function

SetStatusWatch_Error:
    SetStatusWatch = 2
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure SetStatusWatch of Form frmMonitor"

End Function

Sub StatusOnline(ByVal Username As String)
    frmChat.AddQ Username & " has signed onto Battle.net."
End Sub

Function GetStatusWatch(ByVal Username As String) As Byte

    Dim i As Integer
    
    Username = LCase(Username)
    
    For i = 0 To UBound(strUsers)
        If StrComp(Username, strUsers(i), vbTextCompare) = 0 Then
            
            GetStatusWatch = StatusWatch(i)
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

