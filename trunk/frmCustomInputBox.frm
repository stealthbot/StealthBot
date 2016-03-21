VERSION 5.00
Begin VB.Form frmCustomInputBox 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(caption)"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboGame 
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
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "Choose One"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ComboBox cboServer 
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
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "Choose One"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "X &Cancel"
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
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< &Back"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >>"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtInput 
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
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "[ message ]"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmCustomInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentPos As Integer
Private InputMessages() As String
Private InputValues() As String
Private Captions() As String
Private NeedsExpansionKey As Boolean

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ReDim InputMessages(9)
    ReDim InputValues(7)
    ReDim Captions(9)
    
    InputMessages(0) = "Welcome to StealthBot's Step-By-Step Setup. Click Next to begin. You may click Cancel at any point and no changes will be made."
    InputMessages(1) = "Please enter the username you'd like your bot to use. If it's not already existent, the bot will create it."
    InputMessages(2) = "Enter the password corresponding with the account you just entered."
    InputMessages(3) = "Which game would you like the bot to connect with?"
    InputMessages(4) = "Please enter a valid cd-key for the game you specified earlier."
    InputMessages(5) = "Please enter a valid cd-key for the %n expansion. (Both a Regular and an Expansion key are required.)"
    InputMessages(6) = "What channel would you like the bot to use as its home?"
    InputMessages(7) = "To which Battle.net gateway will you be connecting? (USEast, USWest, Asia, Europe)"
    InputMessages(8) = "Enter your main Battle.net account name. This will act as the bot's ""owner"" account -- you can leave it blank." & vbCrLf & vbCrLf & "IMPORTANT: Correct the Owner Name later if the bot sees your account name differently once it has connected -- it must be EXACT and include any @Realm stuff that the bot sees on your name."
    InputMessages(9) = "Congratulations, you're all set up! Enjoy your StealthBot, and remember to visit http://www.stealthbot.net if you have problems."
    
    Captions(0) = "Welcome to StealthBot!"
    Captions(1) = "Username"
    Captions(2) = "Password"
    Captions(3) = "Game Selection"
    Captions(4) = "CD-Key"
    Captions(5) = "Expansion CD-Key"
    Captions(6) = "Home Channel"
    Captions(7) = "Gateway Selection"
    Captions(8) = "Bot Owner"
    Captions(9) = "Finished!"
    
    Call cboGame_Click
    
    With cboServer
        .AddItem "(Choose One)"
        .AddItem "USEast / Azeroth"
        .AddItem "USWest / Lordaeron"
        .AddItem "Asia / Kalimdor"
        .AddItem "Europe / Northrend"
        .ListIndex = 0
        .Visible = False
    End With
    
    With cboGame
        .AddItem "(Choose One)"
        .AddItem "Warcraft II Battle.Net Edition"
        .AddItem "Starcraft"
        .AddItem "Starcraft: Brood War"
        .AddItem "Diablo II"
        .AddItem "Diablo II: Lord of Destruction"
        .AddItem "Warcraft III"
        .AddItem "Warcraft III: The Frozen Throne"
        .ListIndex = 0
        .Visible = False
    End With
    
    CurrentPos = 0
    ShowCurrentPos
End Sub

Private Sub cmdBack_Click()
    If CurrentPos < 9 Then InputValues(CurrentPos - 1) = txtInput.text
    CurrentPos = CurrentPos - 1
    ShowCurrentPos (True)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If CurrentPos < 9 Then
        If CurrentPos > 0 Then InputValues(CurrentPos - 1) = txtInput.text
        CurrentPos = CurrentPos + 1
        ShowCurrentPos
        
        On Error Resume Next
        If (CurrentPos <> 3 And CurrentPos <> 7) Then txtInput.SetFocus
    Else
        If frmChat.SettingsForm Is Nothing Then
            Set frmChat.SettingsForm = New frmSettings
            frmChat.SettingsForm.Show
        End If
        
        With frmChat.SettingsForm
            .ShowPanel spConnectionConfig
            .txtUsername.text = InputValues(0)
            .txtPassword.text = InputValues(1)
            
'            .AddItem "(Choose One)"
'            .AddItem "Warcraft II Battle.Net Edition"
'            .AddItem "Starcraft"
'            .AddItem "Starcraft: Brood War"
'            .AddItem "Diablo II"
'            .AddItem "Diablo II: Lord of Destruction"
'            .AddItem "Warcraft III"
'            .AddItem "Warcraft III: The Frozen Throne"
        
            Select Case cboGame.ListIndex
                Case 0
                    .optSTAR.Value = True
                    Call .optSTAR_Click
                Case 1
                    .optW2BN.Value = True
                    Call .optW2BN_Click
                Case 2
                    .optSTAR.Value = True
                    Call .optSTAR_Click
                Case 3
                    .optSEXP.Value = True
                    Call .optSEXP_Click
                Case 4
                    .optD2DV.Value = True
                    Call .optD2DV_Click
                Case 5
                    .optD2XP.Value = True
                    Call .optD2XP_Click
                Case 6
                    .optWAR3.Value = True
                    Call .optWAR3_Click
                Case 7
                    .optW3XP.Value = True
                    Call .optW3XP_Click
            End Select
            
            .cboCDKey.text = InputValues(3)
            .lblAddCurrentKey_Click
            .txtExpKey.text = InputValues(4)
            .txtHomeChan.text = InputValues(5)
            
'            .AddItem "(Choose One)"
'            .AddItem "USEast / Azeroth"
'            .AddItem "USWest / Lordaeron"
'            .AddItem "Asia / Kalimdor"
'            .AddItem "Europe / Northrend"

            Select Case cboServer.ListIndex
                Case 0: .cboServer.text = "useast.battle.net"
                Case 1: .cboServer.text = "useast.battle.net"
                Case 2: .cboServer.text = "uswest.battle.net"
                Case 3: .cboServer.text = "asia.battle.net"
                Case 4: .cboServer.text = "europe.battle.net"
            End Select
            
            .txtOwner.text = InputValues(7)
            
            Unload Me
        End With
    End If
End Sub

Private Sub ShowCurrentPos(Optional ByVal GoingBackwards As Boolean = False)
    Dim s As String
    
    'debug.print "CurrentPos: " & CurrentPos
    'debug.print "GoingBackwards: " & GoingBackwards
    

    If CurrentPos = 5 Then
        If GoingBackwards Then
            CurrentPos = 4
            lblText.Caption = InputMessages(CurrentPos)
        Else
            If NeedsExpansionKey Then
                s = FormatOutput(InputMessages(CurrentPos))
                CurrentPos = 5
            Else
                s = InputMessages(CurrentPos + 1)
                CurrentPos = 6
            End If
            
            lblText.Caption = s
        End If
    Else
        lblText.Caption = InputMessages(CurrentPos)
    End If
    
    Me.Caption = Captions(CurrentPos)
    
    txtInput.Visible = ((CurrentPos > 0 And CurrentPos < 9) And (CurrentPos <> 3 And CurrentPos <> 7))
    cboServer.Visible = (CurrentPos = 7)
    cboGame.Visible = (CurrentPos = 3)
    cmdBack.Enabled = (CurrentPos > 0)
    
    If CurrentPos > 0 And CurrentPos < 9 Then
        txtInput.text = InputValues(CurrentPos - 1)
    End If
    
    If CurrentPos = 9 Then
        cmdNext.Caption = "&Finish!"
    Else
        cmdNext.Caption = ">> &Next"
    End If
End Sub

Function FormatOutput(ByVal sIn As String) As String
    Select Case cboGame.ListIndex
        Case 5: FormatOutput = Replace(sIn, "%n", "Diablo II: Lord of Destruction")
        Case 7: FormatOutput = Replace(sIn, "%n", "Warcraft III: The Frozen Throne")
        Case Else: CurrentPos = CurrentPos + 1
    End Select
End Function

Sub cboGame_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub cboGame_Click()
    NeedsExpansionKey = (cboGame.ListIndex = 5 Or cboGame.ListIndex = 7)
    'debug.print "NeedsExpansionKey is now " & CBool(cboGame.ListIndex = 5 Or cboGame.ListIndex = 7)
End Sub

Sub cboServer_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdNext_Click
        KeyAscii = 0
    End If
End Sub
