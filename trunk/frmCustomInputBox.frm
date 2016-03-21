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
Private MaxLengthVal() As Integer
Private IsExpansion As Boolean
Private NeedsExpansionKey As Boolean

Private Const STEP_WELC = 0
Private Const STEP_NAME = 1
Private Const STEP_PASS = 2
Private Const STEP_PROD = 3
Private Const STEP_KEY1 = 4
Private Const STEP_KEY2 = 5
Private Const STEP_CHAN = 6
Private Const STEP_SERV = 7
Private Const STEP_OWNR = 8
Private Const STEP_DONE = 9

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ReDim InputMessages(9)
    ReDim Captions(9)
    ReDim InputValues(7)
    ReDim MaxLengthVal(7)
    
    InputMessages(STEP_WELC) = "Welcome to StealthBot's Step-By-Step Setup. Click Next to begin. You may click Cancel at any point and no changes will be made."
    InputMessages(STEP_NAME) = "Please enter the username you'd like your bot to use. If it's not already existent, the bot will create it."
    InputMessages(STEP_PASS) = "Enter the password corresponding with the account you just entered."
    InputMessages(STEP_PROD) = "Which game would you like the bot to connect with?"
    InputMessages(STEP_KEY1) = "Please enter a valid CDKey for the game %s."
    InputMessages(STEP_KEY2) = "Please enter a valid CDKey for the %s expansion. (Both a Regular and an Expansion key are required.)"
    InputMessages(STEP_CHAN) = "What channel would you like the bot to use as its home?"
    InputMessages(STEP_SERV) = "To which Battle.net gateway will you be connecting? (USEast, USWest, Asia, Europe)"
    InputMessages(STEP_OWNR) = "Enter your main Battle.net account name. This will act as the bot's ""owner"" account -- you can leave it blank." & vbCrLf & vbCrLf & "IMPORTANT: Correct the Owner Name later if the bot sees your account name differently once it has connected -- it must be EXACT and include any @Realm stuff that the bot sees on your name."
    InputMessages(STEP_DONE) = "Congratulations, you're all set up! Enjoy your StealthBot, and remember to visit http://www.stealthbot.net if you have problems."
    
    Captions(STEP_WELC) = "Welcome to StealthBot!"
    Captions(STEP_NAME) = "Username"
    Captions(STEP_PASS) = "Password"
    Captions(STEP_PROD) = "Game Selection"
    Captions(STEP_KEY1) = "CD-Key"
    Captions(STEP_KEY2) = "Expansion CD-Key"
    Captions(STEP_CHAN) = "Home Channel"
    Captions(STEP_SERV) = "Gateway Selection"
    Captions(STEP_OWNR) = "Bot Owner"
    Captions(STEP_DONE) = "Finished!"
    
    MaxLengthVal(STEP_NAME - 1) = 15
    MaxLengthVal(STEP_PASS - 1) = 0
    MaxLengthVal(STEP_KEY1 - 1) = 60
    MaxLengthVal(STEP_KEY2 - 1) = 60
    MaxLengthVal(STEP_CHAN - 1) = 30
    MaxLengthVal(STEP_OWNR - 1) = 30
    
    NeedsExpansionKey = False
    IsExpansion = False
    
    With cboGame
        .AddItem "WarCraft II: Battle.net Edition"
        .AddItem "StarCraft"
        .AddItem "StarCraft: Brood War"
        .AddItem "Diablo II"
        .AddItem "Diablo II: Lord of Destruction"
        .AddItem "WarCraft III"
        .AddItem "WarCraft III: The Frozen Throne"
        .Text = "- choose one -"
        .Visible = False
    End With
    
    With cboServer
        .AddItem "USEast / Azeroth"
        .AddItem "USWest / Lordaeron"
        .AddItem "Asia / Kalimdor"
        .AddItem "Europe / Northrend"
        .Text = "- choose one -"
        .Visible = False
    End With
    
    CurrentPos = 0
    ShowCurrentPos
End Sub

Private Sub cmdBack_Click()
    If CurrentPos < STEP_DONE Then InputValues(CurrentPos - 1) = txtInput.Text
    CurrentPos = CurrentPos - 1
    Call ShowCurrentPos(True)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If CurrentPos < STEP_DONE Then
        If CurrentPos > STEP_WELC Then InputValues(CurrentPos - 1) = txtInput.Text
        CurrentPos = CurrentPos + 1
        ShowCurrentPos
        
        On Error Resume Next
        If (CurrentPos <> STEP_PROD And CurrentPos <> STEP_SERV) Then txtInput.SetFocus
    Else
        If frmChat.SettingsForm Is Nothing Then
            Set frmChat.SettingsForm = New frmSettings
            frmChat.SettingsForm.Show
        End If
        
        With frmChat.SettingsForm
            .ShowPanel spConnectionConfig
            .txtUsername.Text = InputValues(STEP_NAME - 1)
            .txtPassword.Text = InputValues(STEP_PASS - 1)
            
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
                    .optW2BN.Value = True
                    Call .optW2BN_Click
                Case 1
                    .optSTAR.Value = True
                    Call .optSTAR_Click
                Case 2
                    .optSEXP.Value = True
                    Call .optSEXP_Click
                Case 3
                    .optD2DV.Value = True
                    Call .optD2DV_Click
                Case 4
                    .optD2XP.Value = True
                    Call .optD2XP_Click
                Case 5
                    .optWAR3.Value = True
                    Call .optWAR3_Click
                Case 6
                    .optW3XP.Value = True
                    Call .optW3XP_Click
            End Select
            
            .cboCDKey.Text = InputValues(STEP_KEY1 - 1)
            .lblAddCurrentKey_Click
            .txtExpKey.Text = InputValues(STEP_KEY2 - 1)
            .txtHomeChan.Text = InputValues(STEP_CHAN - 1)
            
'            .AddItem "(Choose One)"
'            .AddItem "USEast / Azeroth"
'            .AddItem "USWest / Lordaeron"
'            .AddItem "Asia / Kalimdor"
'            .AddItem "Europe / Northrend"

            Select Case cboServer.ListIndex
                Case 0: .cboServer.Text = "useast.battle.net"
                Case 1: .cboServer.Text = "uswest.battle.net"
                Case 2: .cboServer.Text = "asia.battle.net"
                Case 3: .cboServer.Text = "europe.battle.net"
            End Select
            
            .txtOwner.Text = InputValues(STEP_OWNR - 1)
            
            Unload Me
        End With
    End If
End Sub

Private Sub ShowCurrentPos(Optional ByVal GoingBackwards As Boolean = False)
    Dim s As String
    Dim InputPresent As Boolean
    
    'debug.print "CurrentPos: " & CurrentPos
    'debug.print "GoingBackwards: " & GoingBackwards

    If CurrentPos = STEP_KEY2 Then
        If Not NeedsExpansionKey Then
            If GoingBackwards Then
                CurrentPos = STEP_KEY1
            Else
                CurrentPos = STEP_CHAN
            End If
        End If
        
        lblText.Caption = FormatOutput(InputMessages(CurrentPos))
    Else
        lblText.Caption = FormatOutput(InputMessages(CurrentPos))
    End If
    
    Me.Caption = Captions(CurrentPos)
    
    txtInput.Visible = ((CurrentPos > 0 And CurrentPos < 9) And (CurrentPos <> 3 And CurrentPos <> 7))
    cboServer.Visible = (CurrentPos = STEP_SERV)
    cboGame.Visible = (CurrentPos = STEP_PROD)
    cmdBack.Enabled = (CurrentPos > STEP_WELC)
    txtInput.PasswordChar = IIf(CurrentPos = STEP_PASS, "*", vbNullString)
    
    InputPresent = True
    If CurrentPos > STEP_WELC And CurrentPos < STEP_DONE Then
        txtInput.Text = InputValues(CurrentPos - 1)
        txtInput.MaxLength = MaxLengthVal(CurrentPos - 1)
        InputPresent = False
        If CurrentPos = STEP_PROD And StrComp(cboGame.Text, "- choose one -", vbBinaryCompare) <> 0 Then
            InputPresent = True
        ElseIf CurrentPos = STEP_SERV And StrComp(cboServer.Text, "- choose one -", vbBinaryCompare) <> 0 Then
            InputPresent = True
        ElseIf LenB(txtInput.Text) > 0 Then
            InputPresent = True
        ElseIf CurrentPos = STEP_OWNR Then
            InputPresent = True
        End If
    End If
    cmdNext.Enabled = InputPresent
    
    If CurrentPos = STEP_DONE Then
        cmdNext.Caption = "&Finish!"
    Else
        cmdNext.Caption = ">> &Next"
    End If
End Sub

Function FormatOutput(ByVal sIn As String) As String
    FormatOutput = sIn
    If CurrentPos = STEP_KEY1 Then
        If IsExpansion Then
            FormatOutput = Replace(sIn, "%s", cboGame.List(cboGame.ListIndex - 1))
        Else
            FormatOutput = Replace(sIn, "%s", cboGame.List(cboGame.ListIndex))
        End If
    ElseIf CurrentPos = STEP_KEY2 Then
        If NeedsExpansionKey Then
            FormatOutput = Replace(sIn, "%s", cboGame.List(cboGame.ListIndex))
        End If
    End If
End Function

Sub cboGame_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub cboGame_Click()
    NeedsExpansionKey = (cboGame.ListIndex = 4 Or cboGame.ListIndex = 6)
    IsExpansion = (cboGame.ListIndex = 2 Or cboGame.ListIndex = 4 Or cboGame.ListIndex = 6)
    cmdNext.Enabled = (StrComp(cboGame.Text, "- choose one -", vbBinaryCompare) <> 0)
    'debug.print "NeedsExpansionKey is now " & CBool(cboGame.ListIndex = 4 Or cboGame.ListIndex = 6)
End Sub

Sub cboServer_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub cboServer_Click()
    cmdNext.Enabled = (StrComp(cboServer.Text, "- choose one -", vbBinaryCompare) <> 0)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdNext_Click
        KeyAscii = 0
    End If
End Sub

Sub txtInput_Change()
    cmdNext.Enabled = (LenB(txtInput.Text) > 0)
    If CurrentPos = STEP_OWNR Then cmdNext.Enabled = True
End Sub
