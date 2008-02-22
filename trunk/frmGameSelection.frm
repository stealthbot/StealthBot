VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGameSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Selection"
   ClientHeight    =   3600
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
   ScaleHeight     =   3600
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":3046
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":61AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":93DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":B9E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":EEAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":11691
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":1483C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":17885
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":1A921
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":1DE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGameSelection.frx":207F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvGames 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   670
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Game"
         Object.Width           =   5071
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   3220
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   2602
      TabIndex        =   3
      Top             =   3220
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Which game do you wish to create an entry for?"
      Height          =   495
      Left            =   440
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmGameSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmDBManager.m_game = _
        vbNullString

    Call Unload(frmGameSelection)
End Sub

Private Sub cmdOK_Click()
    Dim game As String ' ...
    
    Select Case (lvGames.SelectedItem.index)
        Case 1:  game = "CHAT"
        Case 2:  game = "STAR"
        Case 3:  game = "SSHR"
        Case 4:  game = "JSTR"
        Case 5:  game = "SEXP"
        Case 6:  game = "DRTL"
        Case 7:  game = "DSHR"
        Case 8:  game = "D2DV"
        Case 9:  game = "D2XP"
        Case 10: game = "W2BN"
        Case 11: game = "WAR3"
        Case 12: game = "W3XP"
    End Select

    frmDBManager.m_game = _
        game
    
    Call Unload(frmGameSelection)
End Sub

Private Sub Form_Load()
    Call lvGames.ListItems.Add(, , "Chat", , 7)
    Call lvGames.ListItems.Add(, , "StarCraft", , 1)
    Call lvGames.ListItems.Add(, , "StarCraft: Shareware", , 12)
    Call lvGames.ListItems.Add(, , "StarCraft: Japanese", , 10)
    Call lvGames.ListItems.Add(, , "StarCraft: Brood War", , 2)
    Call lvGames.ListItems.Add(, , "Diablo I: Retail", , 8)
    Call lvGames.ListItems.Add(, , "Diablo I: Shareware", , 9)
    Call lvGames.ListItems.Add(, , "Diablo II", , 3)
    Call lvGames.ListItems.Add(, , "Diablo II: Lord of Destruction", , 4)
    Call lvGames.ListItems.Add(, , "WarCraft II: Battle.net Edition", , 5)
    Call lvGames.ListItems.Add(, , "WarCraft III: Reign of Chaos", , 6)
    Call lvGames.ListItems.Add(, , "WarCraft III: The Frozen Throne", , 11)
End Sub
