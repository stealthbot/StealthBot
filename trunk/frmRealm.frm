VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRealm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diablo II Realm Login"
   ClientHeight    =   4350
   ClientLeft      =   525
   ClientTop       =   840
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10920
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrLoginTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   2760
   End
   Begin VB.OptionButton optCreateNew 
      BackColor       =   &H00000000&
      Caption         =   "Create New Character"
      Enabled         =   0   'False
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
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optViewExisting 
      BackColor       =   &H00000000&
      Caption         =   "View Existing Characters"
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
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imlChars 
      Left            =   9600
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   103
      ImageHeight     =   201
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":0B2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":1C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":2E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":3EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":4FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":5FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealm.frx":7159
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwChars 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7223
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imlChars"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
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
      NumItems        =   0
   End
   Begin VB.Frame fraCreateNew 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtCharName 
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
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox chkLadder 
         BackColor       =   &H00000000&
         Caption         =   "Ladder"
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
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkHardcore 
         BackColor       =   &H00000000&
         Caption         =   "Hardcore"
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
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkExpansion 
         BackColor       =   &H00000000&
         Caption         =   "Expansion"
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
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "&Create"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Assassin"
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
         Height          =   255
         Index           =   7
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Druid"
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
         Height          =   255
         Index           =   6
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Barbarian"
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
         Height          =   255
         Index           =   5
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Paladin"
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
         Height          =   255
         Index           =   4
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Necromancer"
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
         Height          =   255
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Sorceress"
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
         Height          =   255
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optNewCharType 
         BackColor       =   &H00000000&
         Caption         =   "Amazon"
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
         Height          =   255
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Images (c) Blizzard Entertainment"
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
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Desired character name:"
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
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Image imgCharPortrait 
         Height          =   3015
         Left            =   2760
         Picture         =   "frmRealm.frx":82CA
         Top             =   480
         Width           =   1545
      End
   End
   Begin VB.Label lblExpiration 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   615
      Left            =   9480
      TabIndex        =   20
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblSecondsCap 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "seconds."
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
      Height          =   255
      Left            =   9480
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblSeconds 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   855
      Left            =   9480
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete This Character"
      End
   End
End
Attribute VB_Name = "frmRealm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MCPHandler As clsMCPHandler
Attribute MCPHandler.VB_VarHelpID = -1

Private Unload_SuccessfulLogin As Boolean
Private CharListReceived As Boolean
Private mTicks As Long
Private CharIsExpansion As Collection
Private CharExpiration As Collection
Private IndexToDelete As Integer
Private CreatedExpRealmChar As Integer
Private CharacterCount As Integer

'    Unknown& = &H0
'    Amazon& = &H1
'    Sorceress& = &H2
'    Necromancer& = &H3
'    Paladin& = &H4
'    Barbarian& = &H5
'    Druid& = &H6
'    Assassin& = &H7

Private Sub Form_Load()
    Dim B As Boolean
    
    Me.Icon = frmChat.Icon
    
    Set MCPHandler = New clsMCPHandler
    Set CharIsExpansion = New Collection
    Set CharExpiration = New Collection
    
    Call optNewCharType_Click(1)
    Call optViewExisting_Click
    
    lblExpiration.Visible = True
    Unload_SuccessfulLogin = False
    
    B = (BotVars.Product = "PX2D")
    
    chkExpansion.Enabled = B
    optNewCharType(6).Enabled = B
    optNewCharType(7).Enabled = B
    
    'lvwChars.ListItems.Add , "temp", "Please wait..."
    lblExpiration.Caption = "Please wait..."
    
    IndexToDelete = -1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lvwChars.ListItems.Clear
    
    If ((Not (Unload_SuccessfulLogin)) Or (RealmError)) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Login cancelled, proceeding with non-realm login."
        
        If frmChat.sckMCP.State <> 0 Then
            frmChat.sckMCP.Close
        End If
        
        SendEnterChatSequence
    End If
    
    RealmError = False
    mTicks = 0
    
    Set MCPHandler = Nothing
    Set CharIsExpansion = Nothing
    Set CharExpiration = Nothing
End Sub

Private Sub lvwChars_Click()
    StopLoginTimer
End Sub

Private Sub lvwChars_DblClick()
    'On Error Resume Next

    With lvwChars
        If Not (.SelectedItem Is Nothing) And CharListReceived Then
            If CharIsExpansion.Item(.SelectedItem.Key) And _
                Not (StrComp(BotVars.Product, "PX2D") = 0) And _
                CreatedExpRealmChar <> .SelectedItem.Index Then
                
                frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] That is an expansion character. Please log on using Diablo II: Lord of Destruction."
                frmChat.SetFocus
            Else
                MCPHandler.LogonToCharacter .SelectedItem.Key
                Unload_SuccessfulLogin = True
                RealmError = False
            End If
            'Debug.Print "-- " & CharIsExpansion.Item(.SelectedItem.Key)
        End If
    End With
End Sub

Private Sub lvwChars_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not (lvwChars.SelectedItem Is Nothing) Then
        lblExpiration.Caption = CharExpiration.Item(lvwChars.SelectedItem.Key)
    End If
End Sub

Private Sub lvwChars_KeyDown(KeyCode As Integer, Shift As Integer)
    StopLoginTimer
End Sub

Private Sub lvwChars_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Call lvwChars_DblClick
    ElseIf KeyAscii = vbKeyEscape Then
        Unload Me
    Else
        StopLoginTimer
    End If
End Sub

Private Sub cmdCreate_Click()
    Dim i As Integer
    Dim Flags As Long
    
    CreatedExpRealmChar = 0
    
    If lvwChars.ListItems.Count > 7 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Your account is full! Delete a character before trying to create another."
    Else
        If Len(txtCharName.Text) > 2 Then
            If chkLadder.Value = 1 Then Flags = Flags Or &H40
            
            If chkExpansion.Value = 1 Then
                Flags = Flags Or &H20
                CreatedExpRealmChar = lvwChars.ListItems.Count
            End If
            
            If chkHardcore.Value = 1 Then Flags = Flags Or &H4
            
            For i = 1 To 7
                If optNewCharType(i).Value = True Then
                    MCPHandler.CreateMCPCharacter i - 1, Flags, txtCharName.Text
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub lvwChars_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If Not (lvwChars.SelectedItem Is Nothing) Then
            PopupMenu mnuPop
        End If
    End If
End Sub

Private Sub MCPHandler_CharDeleteResponse(ByVal Success As Boolean)
    If Success Then
        frmChat.AddChat RTBColors.SuccessText, "[REALM] Character successfully deleted."
        
        If IndexToDelete > 0 Then
            lvwChars.ListItems.Remove IndexToDelete
            CharIsExpansion.Remove IndexToDelete
            IndexToDelete = -1
        End If
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] There was a problem deleting your character."
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Please log in using the actual Diablo II game to delete it."
    End If
End Sub

Private Sub MCPHandler_CharListResponse(ByVal NumCharacters As Integer)
    lvwChars.ListItems.Clear

    ClearExpansionCollection
    ClearExpirationCollection
    ClearExpirationLabel
    
    CharListReceived = True
    
    CharacterCount = NumCharacters
    
    If CharacterCount = 0 Then
        optCreateNew.Enabled = True
    End If
    
    'tmrLoginTimeout.Enabled = True
End Sub

Private Sub MCPHandler_CharCreateResponse(ByVal Status As Byte, ByVal Message As String)
    If Status = 0 Then
        frmChat.AddChat RTBColors.SuccessText, "[REALM] " & Message
        lvwChars.ListItems.Clear
        
        CharListReceived = False
        
        MCPHandler.RequestCharacterList
        Call optViewExisting_Click
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] " & Message
        txtCharName.Text = vbNullString
    End If
    
    ClearExpirationLabel
End Sub

Private Sub MCPHandler_CharListEntry(ByVal CharName As String, ByVal Statstring As String, ByVal ExpirationDate As Date)
    Dim sOut As String
    Dim Hardcore As Boolean, Ladder As Boolean, Dead As Boolean, Expired As Boolean ', Expansion As Boolean
    Dim IsExpansion As Boolean
    Dim Level As Byte, ClassByte As Byte
    Dim Class As String
    
    Expired = (Sgn(DateDiff("s", Now, ExpirationDate)) = -1)
    
    MCPHandler.GetD2CharStats Statstring, Class, ClassByte, Level, Hardcore, Dead, Ladder, IsExpansion

    sOut = CharName & vbCr & "(a" & IIf(Expired, "n expired ", " ") & IIf(Dead, IIf(Hardcore, "dead ", ""), "") & _
                                    IIf(Hardcore, "hardcore ", "") _
                                    & "level " & Level & Space(1) & _
                                    IIf(IsExpansion, "expansion ", "") & Class & ")"
    
    With lvwChars
        'If .ListItems.Count > 0 Then
        '    .ListItems.Clear
        'End If
            
        CharIsExpansion.Add IsExpansion, CharName
    
        If LenB(CharName) > 0 Then
            
            If Not FindKey(CharName) Then
            
                .ListItems.Add , CharName, sOut, ClassByte + 1
                
                'frmChat.AddChat vbRed, CharName & " (" & .ListItems(1).Key & ")"
                
                If Expired Then
                    .ListItems.Item(.ListItems.Count).ForeColor = vbRed
                End If
                
                CharExpiration.Add IIf(Expired, "Expired ", "Expires ") & vbCrLf & ExpirationDate, CharName
                
                .SetFocus
                
                If (.ListItems.Count = CharacterCount) Then
                    optCreateNew.Enabled = True
                    tmrLoginTimeout.Enabled = True
                End If
            End If
        End If
    End With
End Sub

Private Sub MCPHandler_CharLogonResponse(ByVal Status As Byte, ByVal Message As String)
    If Status = 0 Then
        frmChat.AddChat RTBColors.SuccessText, "[REALM] " & Message
        
        SendEnterChatSequence
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] " & Message
        RealmError = True
    End If
    
    Unload Me
End Sub

Private Sub MCPHandler_RealmStartup(ByVal Status As Byte, ByVal Message As String)
    If Status = 0 Then
        frmChat.AddChat RTBColors.SuccessText, "[REALM] " & Message
        MCPHandler.RequestCharacterList
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] " & Message
        RealmError = True
        Unload frmRealm
    End If
End Sub

Private Sub mnuPopDelete_Click()
    Dim s As String
    
    If Not (lvwChars.SelectedItem Is Nothing) Then
        s = Split(lvwChars.SelectedItem.Text, vbCr)(0)
        IndexToDelete = lvwChars.SelectedItem.Index
        
        MCPHandler.DeleteCharacter s
    End If
End Sub

Private Sub optNewCharType_Click(Index As Integer)
    Dim i As Integer
    
    imgCharPortrait.Picture = imlChars.ListImages.Item(Index + 1).Picture
    
    For i = 1 To 7
        If i <> Index Then optNewCharType(i).Value = False
    Next i
End Sub

Private Sub optViewExisting_Click()
    optCreateNew.Value = False
    
    StopLoginTimer
    
    fraCreateNew.Visible = False
    lvwChars.Visible = True
    lblExpiration.Visible = True
End Sub

Private Sub optCreateNew_Click()
    optViewExisting.Value = False
    
    StopLoginTimer
    
    fraCreateNew.Visible = True
    lvwChars.Visible = False
    lblExpiration.Visible = False
End Sub

Sub StopLoginTimer()
    tmrLoginTimeout.Enabled = False
    lblWarning.Caption = ""
    lblSeconds.Caption = ""
    lblSecondsCap.Caption = ""
End Sub

Private Sub tmrLoginTimeout_Timer()
    Static indexValid As Integer
    
    If (indexValid = 0) Then
        Dim i As Integer
        Dim j As Integer
        
        For i = 1 To lvwChars.ListItems.Count
            If (Len(CharExpiration(i)) >= Len("Expired ")) Then
                If (Left$(CharExpiration(i), Len("Expired ")) <> "Expired ") Then
                    indexValid = i
                    
                    Exit For
                End If
            End If
        Next i
    End If

    mTicks = mTicks + 1

    If (indexValid > 0) Then
        lblWarning.Caption = lvwChars.ListItems(indexValid).Key & vbCrLf & " will be chosen automatically in"
        
        If mTicks >= 30 Then
            MCPHandler.LogonToCharacter lvwChars.ListItems(indexValid).Key
            Unload_SuccessfulLogin = True
            tmrLoginTimeout.Enabled = False
        End If
    Else
        lblWarning.Caption = "You have no characters! Realm login will be cancelled in"
        
        If mTicks >= 30 Then
            frmChat.sckMCP.Close
            SendEnterChatSequence
            Unload Me
        End If
    End If
    
    lblSeconds.Caption = (30 - mTicks)
End Sub

Private Function FindKey(ByVal sKey As String) As Boolean
    Dim i As Long
    
    With lvwChars.ListItems
        If .Count > 0 Then
            For i = 1 To .Count
                If .Item(i).Key = sKey Then
                    FindKey = True
                    Exit Function
                End If
            Next i
        End If
    End With
    
End Function

Sub ClearExpansionCollection()
    'While CharIsExpansion.Count > 0
    '    CharIsExpansion.Remove 1
    'Wend
    
    Set CharIsExpansion = New Collection
End Sub

Sub ClearExpirationCollection()
    'While CharExpiration.Count > 0
    '    CharExpiration.Remove 1
    'Wend
    
    Set CharExpiration = New Collection
End Sub

Sub ClearExpirationLabel()
    lblExpiration.Caption = ""
End Sub

