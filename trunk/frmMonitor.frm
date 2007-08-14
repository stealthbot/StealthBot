VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Monitor"
   ClientHeight    =   4800
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
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
   Begin MSWinsockLib.Winsock wskBNet 
      Left            =   6120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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

Private Sub cmdConnect_Click()
    If wskBNet.State = 0 Then
        wskBNet.RemoteHost = ReadINI("Main", "Server", "config.ini")
        wskBNet.RemotePort = 6112
        wskBNet.Connect
    End If
    cmdConnect.Enabled = False
    cmdDisc.Enabled = True
End Sub

Private Sub cmdRefresh_Click()
    Call Form_Load
End Sub

Private Sub cmdDisc_Click()
    On Error Resume Next
    
    If wskBNet.State <> 0 Then
        wskBNet.Close
        lblStatus.Caption = "User monitor manually disconnected."
        cmdConnect.Enabled = True
        tmrDelay.Interval = 0
    End If
    
    cmdConnect.Enabled = True
    cmdDisc.Enabled = False
End Sub

Private Sub cmdShutdown_Click()
    Call frmChat.DeconstructMonitor
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Me.Icon = frmChat.Icon
    
    ReDim strUsers(0)
    ReDim StatusWatch(0)
    ReDim Sent(0)
        
    Call LoadList

    If Not DisableMonitor Then
        If wskBNet.State <> 0 Then
            wskBNet.Close
        End If
    
        wskBNet.RemoteHost = ReadINI("Main", "Server", "config.ini")
        wskBNet.RemotePort = 6112
        
        If lvMonitor.ListItems.Count > 0 Then
            Call cmdConnect_Click
        End If
        
    Else
        Call cmdDisc_Click
    End If
    
    With lvMonitor
        .SmallIcons = frmChat.imlIcons
        .Icons = frmChat.imlIcons
        .View = lvwReport
    End With
End Sub

Sub LoadList()
    Dim s As String, f As Integer
    

    f = FreeFile
    
    If Dir$(GetProfilePath() & "\monitor.txt") = vbNullString Then
        Open (GetProfilePath() & "\monitor.txt") For Output As #f
            Print #f, "Stealth"
        Close #f
        
        AddUser "Stealth"
        
    Else
    
        Open (GetProfilePath() & "\monitor.txt") For Input As #f
        
        If LOF(f) > 1 Then
        
            lvMonitor.ListItems.Clear
            Do
                Input #f, s
                
                If s <> vbNullString And s <> " " Then
                
                    AddUser s, 1
                
                End If
                
            Loop Until EOF(f)
            
        Else
            
            AddUser "Stealth"
            
        End If
        
    End If
    
    Close #f
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

Private Sub tmrDelay_Timer()
    On Error Resume Next
    If wskBNet.State = 7 Then
        Static SentFirst As Byte
        
        If SentFirst = 0 Then
            SentFirst = 1
            Exit Sub
        ElseIf SentFirst = 1 Then
            wskBNet.SendData "/dnd your StealthBot User Monitor at work" & vbCrLf
            SentFirst = 2
            Exit Sub
        End If
        
        If LastCheck > UBound(strUsers) Then
            LastCheck = 0
        End If
        
        If Len(strUsers(LastCheck)) > 0 Then
            wskBNet.SendData "/whois " & strUsers(LastCheck) & vbCrLf
        End If
        LastCheck = LastCheck + 1
    End If
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
        KeyAscii = 0
    End If
End Sub

Private Sub wskBNet_Connect()
    On Error Resume Next
    wskBNet.SendData Chr(3) & "anonymous" & vbCrLf  'BotVars.Username & vbCrLf & BotVars.Password & vbCrLf
    lblStatus.Caption = "Chat dummy connected, login sent."
End Sub

Private Sub wskBNet_DataArrival(ByVal bytesTotal As Long)
    Dim s As String, i As Integer
    Dim ary() As String
    
    wskBNet.GetData s, vbString
    
    'Debug.Print s
    
    ary = Split(s, Chr(13))
    For i = LBound(ary) To UBound(ary)
        If Left$(ary(i), 1) = Chr(10) Then ary(i) = Right(ary(i), Len(ary(i)) - 1)

        Select Case Left$(ary(i), 4)
            Case "1018"
                If InStr(1, s, "using", vbTextCompare) <> 0 Then
                    s = Right(s, Len(s) - 10)
                    s = Replace(s, Chr(34), vbNullString)
                    UpdateList s
                End If
            Case "2010"
                lblStatus.Caption = "Connected and monitoring users."
                'If InStr(1, s, "Public Chat 1", vbTextCompare) <> 0 Or InStr(1, s, "Public Chat 2", vbTextCompare) Then wskBNet.SendData "/join Public Chat SBUserMonitor-" & Int(Rnd * 50) & vbCrLf
                tmrDelay.Interval = 4900
            Case "1019"
                s = Right(s, Len(s) - 10)
                s = Replace(s, Chr(34), vbNullString)
                UpdateList s, 1
        End Select
    Next i
End Sub

Private Sub wskBNet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Number & Space(1) & Description
    lblStatus.Caption = "A winsock error, " & Number & ", has been encountered. Please manually reconnect the user monitor."
    tmrDelay.Interval = 0
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

Sub RemUser(ByVal i As Integer)
    Dim tmp() As String
    Dim swtmp() As Byte
    Dim sntmp() As Byte
    Dim c As Integer
    Dim f As Integer
    Dim Username As String
    
    f = FreeFile
    
    If i < 1 Then Exit Sub
    
    With lvMonitor.ListItems
        Username = .Item(i).text
        .Item(i).ListSubItems.Clear
        .Remove i
    End With
    
    ReDim tmp(UBound(strUsers) - 1)
    ReDim swtmp(UBound(strUsers) - 1)
    ReDim sntmp(UBound(strUsers) - 1)
    
    Open GetProfilePath() & "\monitor.txt" For Output As #f
    
    For i = 0 To UBound(strUsers)
        If StrComp(Username, strUsers(i), vbTextCompare) <> 0 Then
            tmp(c) = strUsers(i)
            swtmp(c) = StatusWatch(i)
            sntmp(c) = Sent(i)
            Print #f, strUsers(i)
            c = c + 1
        End If
    Next i
    
    Close #f
    
    ReDim strUsers(UBound(tmp))
    ReDim StatusWatch(UBound(tmp))
    ReDim Sent(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        strUsers(i) = tmp(i)
        StatusWatch(i) = swtmp(i)
        Sent(i) = sntmp(i)
    Next i
End Sub

Sub AddUser(ByVal Username As String, Optional ByVal DoNotWriteFile As Byte)
    Dim f As Integer
    f = FreeFile
    
    On Error Resume Next
    With lvMonitor
        .ListItems.Add , Username, Username, , ICSQUELCH
        'Debug.Print .ListItems(.ListItems.Count).ListSubItems.Count
        .ListItems(.ListItems.Count).ListSubItems.Add 0, "status", "Offline", MONITOR_OFFLINE
        .ListItems(.ListItems.Count).ListSubItems.Add 1, "last", "None"
        .ListItems(.ListItems.Count).Tag = "0"
    End With
    
    ReDim Preserve strUsers(UBound(strUsers) + 1)
    ReDim Preserve StatusWatch(UBound(StatusWatch) + 1)
    ReDim Preserve Sent(UBound(Sent) + 1)
    strUsers(UBound(strUsers)) = Username
    
    If Dir$(GetProfilePath() & "\monitor.txt") = vbNullString Then
        Open GetProfilePath() & "\monitor.txt" For Output As #f
        Close #f
    End If
    
    If DoNotWriteFile = 0 Then
    
        Open GetProfilePath() & "\monitor.txt" For Append As #f
        
        Print #f, Username
        
        Close #f
        
    End If
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
        'Debug.Print "Is something! (" & X.Tag & ")"
        If X.ListSubItems(1).text = "Online" Then
            GetUserStatus = 1
        Else
            GetUserStatus = 0
        End If
        
        Set X = Nothing
    Else
        'Debug.Print "Is nothing!"
        GetUserStatus = -1
    End If
End Function

Function GetLastWhoisResponse(ByVal Username As String) As String
    Dim X As ListItem
    
    Set X = lvMonitor.FindItem(Username)
    
    If Not (X Is Nothing) Then
        GetLastWhoisResponse = X.Tag
        
        Set X = Nothing
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
        LastChecked = ""
        LastWhois = ""
        
        GetFullUserStatus = 1
    End If
End Function

