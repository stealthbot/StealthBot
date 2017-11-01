VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRealm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diablo II Realm Login"
   ClientHeight    =   4755
   ClientLeft      =   525
   ClientTop       =   840
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   10920
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboOtherRealms 
      BackColor       =   &H00993300&
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
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btnDisconnect 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      TabIndex        =   29
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton btnChoose 
      Caption         =   "&Logon"
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
      Height          =   300
      Left            =   9480
      TabIndex        =   28
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Timer tmrLoginTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   2760
   End
   Begin VB.OptionButton optCreateNew 
      BackColor       =   &H00000000&
      Caption         =   "Create ne&w character"
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
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton optViewExisting 
      BackColor       =   &H00000000&
      Caption         =   "&View existing characters"
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
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
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
   Begin VB.CommandButton btnUpgrade 
      Caption         =   "&Upgrade"
      Height          =   300
      Left            =   9480
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   9480
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraCreateNew 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      TabIndex        =   11
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
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   21
         Top             =   720
         Width           =   1815
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2280
         Width           =   1815
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton btnCreate 
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
         Left            =   4800
         TabIndex        =   26
         Top             =   2760
         Width           =   1815
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
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   7
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
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
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   6
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblRealm 
         BackColor       =   &H00000000&
         Caption         =   "C&haracter class"
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
         Index           =   9
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblRealm 
         BackColor       =   &H00000000&
         Caption         =   "Character &options"
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
         Index           =   8
         Left            =   4800
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblRealm 
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
         Index           =   7
         Left            =   2400
         TabIndex        =   27
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblRealm 
         BackColor       =   &H00000000&
         Caption         =   "Character &name"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image imgCharPortrait 
         Height          =   3015
         Left            =   2880
         Picture         =   "frmRealm.frx":82CA
         Top             =   480
         Width           =   1545
      End
   End
   Begin MSComctlLib.ListView lvwChars 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7223
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin VB.Label lblRealm 
      BackColor       =   &H00000000&
      Caption         =   "Other &Realms"
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
      Left            =   9480
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblRealm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Expires:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Index           =   1
      Left            =   9480
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblRealm 
      BackColor       =   &H00000000&
      Caption         =   "{0}{1} is a {2} {3} on Realm {4}"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4380
      UseMnemonic     =   0   'False
      Width           =   7695
   End
   Begin VB.Label lblRealm 
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
      Index           =   4
      Left            =   9480
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblRealm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "#"
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
      Left            =   9480
      TabIndex        =   30
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblRealm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Auto-choose X in"
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
      Height          =   735
      Index           =   2
      Left            =   9480
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPopUpgrade 
         Caption         =   "&Upgrade to Expansion"
         Shortcut        =   ^U
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRealm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Unload_SuccessfulLogin As Boolean
' ticks until auto choose
Private m_Ticks As Long
' current auto choose target
Private m_Choice As Integer
' are we on expansion?
Private m_IsExpansion As Boolean
' save selected item for refreshes and menu clicking
Private m_Selection As Integer
' save new character name so m_Selection gets set to the new character we just created
Private m_NewCharacterName As String

'Private Const WM_NCDESTROY = &H82

'    Unknown& = &H0
'    Amazon& = &H1
'    Sorceress& = &H2
'    Necromancer& = &H3
'    Paladin& = &H4
'    Barbarian& = &H5
'    Druid& = &H6
'    Assassin& = &H7

Private Sub Form_Load()
    Dim i As Integer

    Me.Icon = frmChat.Icon

    ' must have a MCPHandler
    If ds.MCPHandler Is Nothing Then
        Set ds.MCPHandler = New clsMCPHandler
        ds.MCPHandler.IsRealmError = True
        Unload Me
    End If

    ' must be on D2
    If (BotVars.Product <> "PX2D" And BotVars.Product <> "VD2D") Then
        ds.MCPHandler.IsRealmError = True
        Unload Me
    End If

    ' store if expansion
    m_IsExpansion = (BotVars.Product = "PX2D")

    ' this is for deciding whether to enter chat after a form close
    m_Unload_SuccessfulLogin = False

    With lblRealm(0) ' detail
        .Caption = "Please wait..."
        .ForeColor = &H888888
    End With

    ' subclass for listview...
    #If COMPILE_DEBUG = 0 Then
        HookWindowProc hWnd
        'm_OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SkipDragLVItem)
    #End If

    ' read auto choose settings from handler
    m_Ticks = ds.MCPHandler.AutoChooseWait
    m_Choice = ds.MCPHandler.AutoChooseTarget
    m_Selection = m_Choice

    ' UI setup
    Call RealmStartupResponse

    ' set up char creation defaults
    chkExpansion.Enabled = m_IsExpansion
    chkExpansion.Value = IIf(m_IsExpansion, 1, 0)
    chkLadder.Value = 1
    optNewCharType(6).Enabled = m_IsExpansion
    optNewCharType(7).Enabled = m_IsExpansion
    optNewCharType(1).Value = True
    Call optNewCharType_Click(1)

    ' view existing
    optViewExisting.Value = True
    Call CharListResponse

    ' MCP handler state
    ds.MCPHandler.FormActive = True

    ' setup timer
    If m_Ticks > 0 Then
        ' hide delete button here as it's been made visible (over seconds remaining in timer)
        ' will be made visible by stopping timer (if character is selected)
        btnUpgrade.Visible = False
        btnDelete.Visible = False

        ' enable the timer and call in last to avoid infinite loops
        tmrLoginTimeout.Enabled = True
        tmrLoginTimeout_Timer
    End If
End Sub

Private Sub Form_Click()
    Call StopLoginTimer
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call StopLoginTimer
End Sub

' Display message names.
'Public Function SkipDragLVItem(ByVal hWnd As Long, ByVal Msg As Long, _
'        ByVal wParam As Long, ByVal lParam As Long) As Long
'
'    Dim nm_hdr As NMHDR
'
'    ' If we're being destroyed,
'    ' restore the original WindowProc.
'    If Msg = WM_NCDESTROY Then
'        SetWindowLong hWnd, GWL_WNDPROC, OldWindowProc
'    ElseIf Msg = WM_NOTIFY Then
'        ' Copy info into the NMHDR structure.
'        CopyMemory nm_hdr, ByVal lParam, Len(nm_hdr)
'
'    End If
'
'    NewWindowProc = CallWindowProc( _
'        OldWindowProc, hWnd, Msg, wParam, _
'        lParam)
'End Function

Private Sub AddCharacterItem(CharacterName As String, CharacterStats As clsUserStats, CharacterExpires As Date, ByVal Index As Integer)

    Dim Expired As Boolean
    Dim NewItem As ListItem
    
    With lvwChars
        If (LenB(CharacterName) > 0) Then
            If (FindCharacter(CharacterName) < 0) Then
                Set NewItem = .ListItems.Add(, _
                        CharacterName, CharacterStats.CharacterTitleAndName, _
                        CharacterStats.CharacterClassID + 1)
                
                NewItem.Tag = Index
                
                If CharacterStats.IsExpansionCharacter Then
                    NewItem.ForeColor = vbGreen
                End If
                
                If IsDateExpired(CharacterExpires) Then
                    NewItem.ForeColor = vbRed
                End If
                
            End If
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lvwChars.ListItems.Clear
    
    tmrLoginTimeout.Enabled = False
    m_Ticks = -1
    
    If ((Not (m_Unload_SuccessfulLogin)) Or (ds.MCPHandler.IsRealmError)) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Logon cancelled."
        
        If frmChat.sckMCP.State <> sckClosed Then
            frmChat.sckMCP.Close
        End If
        
        Call SendEnterChatSequence
        frmChat.mnuRealmSwitch.Enabled = True
        m_Unload_SuccessfulLogin = True
        ds.MCPHandler.IsRealmError = False
    End If
    
    ds.MCPHandler.FormActive = False
    ds.MCPHandler.IsRealmError = False
End Sub

Private Sub cboOtherRealms_Click()
    Dim CurrRealmIndex As Integer
    Dim CurrRealmTitle As String
    Dim NewRealmTitle  As String
    Dim RealmPassword  As String

    Call StopLoginTimer
    
    CurrRealmIndex = ds.MCPHandler.RealmServerSelectedIndex
    CurrRealmTitle = ds.MCPHandler.RealmServerTitle(CurrRealmIndex)
    
    If (StrComp(cboOtherRealms.Text, CurrRealmTitle, vbTextCompare) <> 0) Then
        NewRealmTitle = cboOtherRealms.Text
        
        DisableGUI
        
        ' close connection and switch realms
        If frmChat.sckMCP.State <> sckClosed Then
            frmChat.sckMCP.Close
        End If
        
        Call ds.MCPHandler.RealmServerLogon(NewRealmTitle)
        
        CharListResponse
    End If
End Sub

Private Sub cboOtherRealms_KeyPress(KeyAscii As Integer)
    Call StopLoginTimer
End Sub

Private Sub DisableGUI()
    btnChoose.Enabled = False
    btnDisconnect.Enabled = False
    btnDelete.Visible = False
    btnUpgrade.Visible = False
    cboOtherRealms.Enabled = False
    optCreateNew.Enabled = False
    btnCreate.Enabled = False
    lvwChars.ListItems.Clear
End Sub

' labels are all in a control array so clicking on any one of them calls this:
Private Sub lblRealm_Click(Index As Integer)
    Call StopLoginTimer
End Sub

Private Sub lvwChars_DblClick()
    btnChoose_Click
End Sub

Private Sub lvwChars_ItemClick(ByVal Item As ListItem)
    Dim ExpireText As String
    Dim DetailText As String
    
    If Not (lvwChars.SelectedItem Is Nothing) Then
        m_Selection = lvwChars.SelectedItem.Tag
        
        ExpireText = GetCharacterExpireText(lvwChars.SelectedItem.Tag)
        DetailText = GetCharacterDetailText(lvwChars.SelectedItem.Tag)
        
        With lblRealm(0) ' detail
            .Caption = DetailText
            .ForeColor = vbWhite
        End With
        
        With lblRealm(1) ' expires
            .Caption = ExpireText
            If IsDateExpired(ds.MCPHandler.CharacterExpires(lvwChars.SelectedItem.Tag)) Then
                .ForeColor = vbRed
            Else
                .ForeColor = vbYellow
            End If
        End With
        
        btnChoose.Enabled = CanChooseCharacter(lvwChars.SelectedItem.Tag)
        btnDelete.Visible = True
        btnUpgrade.Visible = CanUpgradeCharacter(lvwChars.SelectedItem.Tag)
    Else
        ' clear detail
        lblRealm(0).Caption = vbNullString
        ' clear expires
        lblRealm(1).Caption = vbNullString
        btnChoose.Enabled = False
        btnDelete.Visible = False
        btnUpgrade.Visible = False
    End If
End Sub

Private Sub lvwChars_KeyDown(KeyCode As Integer, Shift As Integer)
    Call StopLoginTimer
End Sub

Private Sub lvwChars_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        Call btnChoose_Click
    ElseIf KeyCode = vbKeyEscape Then
        Call UnloadNormal
    ElseIf KeyCode = vbKeyDelete Then
        Call HandleCharacterDelete(lvwChars.SelectedItem)
    ElseIf KeyCode = vbKeyU And Shift = vbCtrlMask Then
        Call HandleCharacterUpgrade(lvwChars.SelectedItem)
    End If
End Sub

Private Sub btnCreate_Click()
    Dim i As Integer
    Dim Flags As Long
    
    If lvwChars.ListItems.Count > 7 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] Your account is full! Delete a character before trying to create another."
    Else
        If Len(txtCharName.Text) > 2 Then
            Flags = 0
            If chkLadder.Value = 1 Then Flags = Flags Or &H40
            If chkExpansion.Value = 1 Then Flags = Flags Or &H20
            If chkHardcore.Value = 1 Then Flags = Flags Or &H4
            
            For i = optNewCharType.LBound To optNewCharType.UBound
                If optNewCharType(i).Value = True Then
                    m_NewCharacterName = txtCharName.Text
                    
                    ds.MCPHandler.SEND_MCP_CHARCREATE (i - 1), Flags, txtCharName.Text
                    
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub lvwChars_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call StopLoginTimer
End Sub

Private Sub lvwChars_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Item As ListItem
    
    m_Selection = -1
    
    If Button = vbRightButton Then
        Set Item = lvwChars.HitTest(x, y)
        
        If Not (Item Is Nothing) Then
            m_Selection = Item.Tag
            
            mnuPopUpgrade.Visible = CanUpgradeCharacter(m_Selection)
            
            PopupMenu mnuPop
        End If
    End If
End Sub

Public Sub RealmStartupResponse()
    Dim i          As Integer
    Dim RealmIndex As Integer
    Dim RealmTitle As String
    Dim RealmDescr As String
    Dim RealmIP    As String
    Dim RealmPort  As Integer
    
    RealmIndex = ds.MCPHandler.RealmServerSelectedIndex
    RealmTitle = ds.MCPHandler.RealmServerTitle(RealmIndex)
    RealmDescr = ds.MCPHandler.RealmServerDescription(RealmIndex)
    RealmIP = ds.MCPHandler.RealmSelectedServerIP
    RealmPort = ds.MCPHandler.RealmSelectedServerPort
    
    Me.Caption = StringFormat("Realm {0} - {1} ({2}:{3})", RealmTitle, RealmDescr, RealmIP, CStr(RealmPort))
    
    ' build "other realm" list
    If ds.MCPHandler.RealmServerCount > 1 Then
        cboOtherRealms.Clear
        
        For i = 0 To ds.MCPHandler.RealmServerCount - 1
            cboOtherRealms.AddItem ds.MCPHandler.RealmServerTitle(i)
            
            If (StrComp(RealmTitle, ds.MCPHandler.RealmServerTitle(i), vbTextCompare) = 0) Then
                cboOtherRealms.ListIndex = i
            End If
        Next i
        
        cboOtherRealms.Text = RealmTitle
        
        ' other realms label
        lblRealm(5).Caption = "Other &Realms"
        cboOtherRealms.Visible = True
    Else
        ' other realms label
        lblRealm(5).Caption = "Realm: " & RealmTitle
        cboOtherRealms.Visible = False
    End If
End Sub

' form callback for character list
' also if this form needs to check existing character list
' updates UI after checking character list
Public Sub CharListResponse()
    Dim i As Integer
    Dim NewSelection As Integer
    
    lvwChars.ListItems.Clear
    
    optCreateNew.Value = False
    
    Call StopLoginTimer
    
    fraCreateNew.Visible = False
    lvwChars.Visible = True
    ' clear expires
    lblRealm(1).Caption = vbNullString
    btnDelete.Visible = False
    btnUpgrade.Visible = False
    
    cboOtherRealms.Enabled = True
    btnDisconnect.Enabled = True
    btnCreate.Enabled = True
    btnChoose.Enabled = False
    
    If ds.MCPHandler.RetrievingCharacterList Then
        With lblRealm(0) ' detail
            .Caption = "Retrieving characters..."
            .ForeColor = &H888888
        End With
    ElseIf Not ds.MCPHandler.RealmServerConnected Then
        With lblRealm(0) ' detail
            .Caption = "Switching realm..."
            .ForeColor = &H888888
        End With
    Else
        For i = 0 To ds.MCPHandler.CharacterCount - 1
        
            Call AddCharacterItem(ds.MCPHandler.CharacterName(i), ds.MCPHandler.CharacterStats(i), ds.MCPHandler.CharacterExpires(i), i)
        
        Next i
        
        optCreateNew.Enabled = (ds.MCPHandler.CharacterCount < 8)
        
        If LenB(m_NewCharacterName) > 0 Then
            NewSelection = FindCharacter(m_NewCharacterName)
            
            If NewSelection >= 0 Then m_Selection = NewSelection
            
            m_NewCharacterName = vbNullString
        End If
        
        If m_Selection < 0 Then m_Selection = 0
        
        With lvwChars.ListItems
            If .Count > m_Selection Then
                .Item(m_Selection + 1).Selected = True
                Call lvwChars_ItemClick(.Item(m_Selection + 1))
                .Item(m_Selection + 1).EnsureVisible
            End If
        End With
        
        If ds.MCPHandler.CharacterCount = 0 Then
            With lblRealm(0)
                .Caption = "No characters found."
                .ForeColor = vbRed
            End With
        End If
    End If
End Sub

' form callback for character create
Public Sub CharCreateResponse(ByVal Success As Boolean, ByVal Message As String)
    If Success Then
        ' go back to character list and refresh
        HandleRefresh
        
        ' clear field for returning to this panel
        txtCharName.Text = vbNullString
    Else
        ' clear saved name so we don't try to look for it later
        m_NewCharacterName = vbNullString
        
        ' focus on the textbox since the character create failed
        With txtCharName
            .SelStart = 0
            .SelLength = Len(.Text)
            On Error Resume Next
            .SetFocus
        End With
    End If
End Sub

' form callback for character delete
Public Sub CharDeleteResponse(ByVal Success As Boolean, ByVal Message As String)
    If Success Then
        'If m_Selection >= 0 Then
        '    lvwChars.ListItems.Remove m_Selection + 1
        'Else
        ' always re-request (or MCPHandler.CharacterList will be different)
        HandleRefresh
        'End If
    End If
End Sub

' form callback for character upgrade
Public Sub CharUpgradeResponse(ByVal Success As Boolean, ByVal Message As String)
    If Success Then
        'If m_Selection >= 0 Then
        '    lvwChars.ListItems.Item(m_Selection + 1).ForeColor = vbGreen
        '    Call lvwChars_ItemClick(lvwChars.ListItems.Item(m_Selection + 1))
        'Else
        ' always re-request (or MCPHandler.CharacterList will be different)
        HandleRefresh
        'End If
    End If
End Sub

' form callback for character logon
Public Sub CharLogonResponse(ByVal Success As Boolean, ByVal Message As String)
    If Success Then
        UnloadNormal
    End If
End Sub

' form callback for BNCS close/DoDisconnect
Public Sub UnloadAfterBNCSClose()
    m_Unload_SuccessfulLogin = True
    ds.MCPHandler.IsRealmError = False
    
    #If COMPILE_DEBUG = 0 Then
        UnhookWindowProc hWnd
    #End If
    
    Unload Me
End Sub

Public Sub UnloadRealmError()
    If ds.MCPHandler Is Nothing Then
        Set ds.MCPHandler = New clsMCPHandler
    End If
    
    ds.MCPHandler.IsRealmError = True
    
    #If COMPILE_DEBUG = 0 Then
        UnhookWindowProc hWnd
    #End If
    
    Unload Me
End Sub

Private Sub UnloadNormal()
    #If COMPILE_DEBUG = 0 Then
        UnhookWindowProc hWnd
    #End If
    
    Unload Me
End Sub

Private Sub mnuPopDelete_Click()
    If (m_Selection >= 0 And optViewExisting.Value) Then
        If lvwChars.ListItems.Count > m_Selection Then
            Call HandleCharacterDelete(lvwChars.ListItems(m_Selection + 1))
        End If
    End If
End Sub

Private Sub mnuPopUpgrade_Click()
    If (m_Selection >= 0) And CanUpgradeCharacter(m_Selection And optViewExisting.Value) Then
        If lvwChars.ListItems.Count > m_Selection Then
            Call HandleCharacterUpgrade(lvwChars.ListItems(m_Selection + 1))
        End If
    End If
End Sub

Private Sub btnDelete_Click()
    Call StopLoginTimer
    
    If (Not lvwChars.SelectedItem Is Nothing And optViewExisting.Value) Then
        Call HandleCharacterDelete(lvwChars.SelectedItem)
    End If
End Sub

Private Sub btnUpgrade_Click()
    Call StopLoginTimer
    
    If (Not lvwChars.SelectedItem Is Nothing And optViewExisting.Value) Then
        If CanUpgradeCharacter(lvwChars.SelectedItem.Tag) Then
            Call HandleCharacterUpgrade(lvwChars.SelectedItem)
        End If
    End If
End Sub

Private Sub HandleRefresh()
    lvwChars.ListItems.Clear
    
    ds.MCPHandler.ClearInternalCharacters
    
    optViewExisting.Value = True
    Call CharListResponse
    
    ds.MCPHandler.DoRequestCharacters
End Sub

Private Sub HandleCharacterDelete(ByVal CharacterItem As ListItem)
    Dim Result As VbMsgBoxResult
    
    If Not CharacterItem Is Nothing And optViewExisting.Value Then
        Result = MsgBox(CharacterItem.Text & " will be deleted. This action is irreversable! " & vbNewLine & _
                "Are you sure you want to do that?", vbYesNo Or vbExclamation, "Realm Confirm Delete")
        
        If Result = vbYes Then
            ds.MCPHandler.SEND_MCP_CHARDELETE CharacterItem.Key
        End If
    End If
End Sub

Private Sub HandleCharacterUpgrade(ByVal CharacterItem As ListItem)
    Dim Result As VbMsgBoxResult
    
    If Not CharacterItem Is Nothing And optViewExisting.Value Then
        If CanUpgradeCharacter(CharacterItem.Tag) Then
            Result = MsgBox(CharacterItem.Text & " will be upgraded to a Lord of Destruction character. This action is irreversable! " & vbNewLine & _
                    "Are you sure you want to do that?", vbYesNo Or vbQuestion, "Realm Confirm Upgrade")
            
            If Result = vbYes Then
                ds.MCPHandler.SEND_MCP_CHARUPGRADE CharacterItem.Key
            End If
        End If
    End If
End Sub

Private Sub btnDisconnect_Click()
    Call StopLoginTimer
    
    UnloadNormal
End Sub

Private Sub btnChoose_Click()
    Call StopLoginTimer
    
    With lvwChars
        If Not (.SelectedItem Is Nothing) Then
            If Not CanChooseCharacter(.SelectedItem.Tag) Then
                frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] You must use Diablo II: Lord of Destruction to choose that character."
            Else
                Call ds.MCPHandler.SEND_MCP_CHARLOGON(.SelectedItem.Key)
                m_Unload_SuccessfulLogin = True
                ds.MCPHandler.IsRealmError = False
            End If
        End If
    End With
End Sub

Private Sub optNewCharType_Click(Index As Integer)
    Dim i As Integer
    
    imgCharPortrait.Picture = imlChars.ListImages.Item(Index + 1).Picture
    
    For i = optNewCharType.LBound To optNewCharType.UBound
        If i <> Index Then optNewCharType(i).Value = False
    Next i
End Sub

Private Sub chkExpansion_Click()
    Dim Enable As Boolean
    
    If Not m_IsExpansion Then
        chkExpansion.Value = 0
        Exit Sub
    End If
    
    Enable = (chkExpansion.Value <> 0)
    optNewCharType(6).Enabled = Enable
    optNewCharType(7).Enabled = Enable
    If Not Enable Then
        If optNewCharType(6).Value Or optNewCharType(7).Value Then
            optNewCharType(1).Value = True
            Call optNewCharType_Click(1)
        End If
    End If
End Sub

Private Sub optViewExisting_Click()
    Call CharListResponse
End Sub

Private Sub optCreateNew_Click()
    optViewExisting.Value = False
    
    Call StopLoginTimer
    
    fraCreateNew.Visible = True
    lvwChars.Visible = False
    With lblRealm(0) ' detail
        .Caption = "You may create a character."
        .ForeColor = vbYellow
    End With
    ' expires
    lblRealm(1).Caption = vbNullString
    btnChoose.Enabled = False
    btnDelete.Visible = False
    btnUpgrade.Visible = False
End Sub

Sub StopLoginTimer()
    tmrLoginTimeout.Enabled = False
    ' clear 3 parts of timer labels
    lblRealm(2).Caption = vbNullString
    lblRealm(3).Caption = vbNullString
    lblRealm(4).Caption = vbNullString
End Sub

Private Sub tmrLoginTimeout_Timer()
    Static indexValid As Integer

    indexValid = m_Choice

    ' if selecting nothing, find first unexpired account (no choose setting)
    If (indexValid = -1) Then
        Dim i As Integer
        
        For i = 0 To ds.MCPHandler.CharacterCount - 1
            If CanChooseCharacter(indexValid) Then
                If Not IsDateExpired(ds.MCPHandler.CharacterExpires(i)) Then
                    indexValid = i
                    Exit For
                End If
            End If
        Next i
    ' if choose setting, then select only if not expired
    Else
        If Not CanChooseCharacter(indexValid) Then
            indexValid = -1
        ElseIf IsDateExpired(ds.MCPHandler.CharacterExpires(indexValid)) Then
            indexValid = -1
        End If
    End If

    ' seconds label (part 2 of timer labels)
    lblRealm(3).Caption = CStr(m_Ticks)
    ' seconds cap label (part 2 of timer labels)
    lblRealm(4).Caption = "second" & IIf(m_Ticks <> 1, "s", vbNullString) & "."

    If (indexValid >= 0) Then
        ' warning label (part 1 of timer labels)
        lblRealm(2).Caption = lvwChars.ListItems(indexValid + 1).Text & vbCrLf & " will be chosen automatically in"

        'If m_Selection < 0 Then
        '    lvwChars.ListItems.Item(indexValid + 1).Selected = True
        '    Call lvwChars_ItemClick(lvwChars.ListItems.Item(indexValid + 1))
        'End If

        If m_Ticks <= 0 Then
            tmrLoginTimeout.Enabled = False
            If Not CanChooseCharacter(indexValid) Then
                frmChat.AddChat RTBColors.ErrorMessageText, "[REALM] You must use Diablo II: Lord of Destruction to choose that character."
            Else
                Call ds.MCPHandler.SEND_MCP_CHARLOGON(ds.MCPHandler.CharacterName(indexValid))
                m_Unload_SuccessfulLogin = True
                ds.MCPHandler.IsRealmError = False
            End If
        End If
    Else
        ' warning label
        lblRealm(2).Caption = "No unexpired characters! Realm logon will be cancelled in"

        If m_Ticks <= 0 Then
            tmrLoginTimeout.Enabled = False
            UnloadNormal
        End If
    End If

    m_Ticks = m_Ticks - 1
End Sub

Private Function FindCharacter(ByVal sKey As String) As Integer

    Dim i As Integer
    
    With lvwChars.ListItems
        If .Count > 0 Then
            For i = 1 To .Count
                If .Item(i).Key = sKey Then
                    FindCharacter = i - 1
                    Exit Function
                End If
            Next i
        End If
    End With
    
    FindCharacter = -1
    
End Function

Private Function IsDateExpired(ByVal Expires As Date) As Boolean
    IsDateExpired = (Sgn(DateDiff("s", UtcNow, Expires)) = -1)
End Function

Private Function CanUpgradeCharacter(ByVal CharIndex As Integer) As Boolean
    Dim Stats As clsUserStats

    Set Stats = ds.MCPHandler.CharacterStats(CharIndex)

    CanUpgradeCharacter = False
    If Stats Is Nothing Then Exit Function

    CanUpgradeCharacter = (Not Stats.IsExpansionCharacter And m_IsExpansion)
End Function

Private Function CanChooseCharacter(ByVal CharIndex As Integer) As Boolean
    Dim Stats As clsUserStats

    Set Stats = ds.MCPHandler.CharacterStats(CharIndex)

    CanChooseCharacter = False
    If Stats Is Nothing Then Exit Function

    ' must be PX2D if isExpansion, otherwise doesn't matter
    CanChooseCharacter = (Stats.IsExpansionCharacter Imp m_IsExpansion)
End Function

Private Function GetCharacterExpireText(ByVal CharIndex As Integer) As String
    Dim Expires    As Date
    Dim ExpireType As String
    Dim ExpireDiff As Long
    Dim ExpireVal  As String
    Dim ExpireUnit As String
    
    Expires = ds.MCPHandler.CharacterExpires(CharIndex)
    
    ExpireDiff = Abs(DateDiff("h", UtcNow, Expires))
    ExpireUnit = "hour"
    If ExpireDiff <> 1 Then ExpireUnit = ExpireUnit & "s"
    
    If ExpireDiff > 24 Then
        ExpireDiff = Round(ExpireDiff / 24)
        ExpireUnit = "day"
        If ExpireDiff <> 1 Then ExpireUnit = ExpireUnit & "s"
    End If
    
    If IsDateExpired(Expires) Then
        ExpireType = "Expired"
        ExpireVal = CStr(ExpireDiff) & " " & ExpireUnit & " ago"
    Else
        ExpireType = "Expires"
        ExpireVal = "in " & CStr(ExpireDiff) & " " & ExpireUnit
    End If
    
    If ExpireDiff = 0 Then
        If IsDateExpired(Expires) Then
            ExpireVal = "minutes ago"
        Else
            ExpireType = "Expires"
            ExpireVal = "in minutes"
        End If
    End If
    
    GetCharacterExpireText = StringFormat("{0} on{1}{2}{1}({3})", ExpireType, vbCrLf, _
            CStr(UtcToLocal(Expires)), ExpireVal)
End Function

Private Function GetCharacterDetailText(ByVal CharIndex As Integer) As String
    Dim Stats        As clsUserStats
    Dim NonExpansion As String
    Dim NonLadder    As String
    Dim IsDead       As String
    Dim NonHardcore  As String
    
    Set Stats = ds.MCPHandler.CharacterStats(CharIndex)
    
    If Not Stats.IsLadderCharacter Then
        NonLadder = "non-"
    End If
    
    If Not Stats.IsHardcoreCharacter Then
        NonHardcore = "non-"
    ElseIf Stats.IsCharacterDead Then
        IsDead = "dead "
    End If
    
    If Not Stats.IsExpansionCharacter Then
        NonExpansion = "non-"
    End If
    
    GetCharacterDetailText = StringFormat("{0} is a {1}ladder, {2}hardcore, {3}expansion {4} in {5}.", _
            Stats.CharacterTitleAndName, NonLadder, NonHardcore, NonExpansion, Stats.CharacterClass, Stats.CurrentActAndDifficulty)
End Function



