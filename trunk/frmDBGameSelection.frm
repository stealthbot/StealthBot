VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBGameSelection 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Entry - Select Game"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
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
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   28
      ImageHeight     =   14
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":023A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":03A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":05C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":0904
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":0C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":1130
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":128E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":13E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":1632
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":1C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":20F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBGameSelection.frx":2250
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvGames 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
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
      ForeColor       =   16777215
      BackColor       =   10040064
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Game"
         Object.Width           =   5071
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   3220
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2602
      TabIndex        =   2
      Top             =   3220
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Which game do you wish to create an entry for?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   440
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDBGameSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmDBManager.m_EntryName = vbNullString

    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Game As String
    
    Select Case (lvGames.SelectedItem.Index)
        Case 1:  Game = PRODUCT_W2BN
        Case 2:  Game = PRODUCT_STAR
        Case 3:  Game = PRODUCT_SEXP
        Case 4:  Game = PRODUCT_D2DV
        Case 5:  Game = PRODUCT_D2XP
        Case 6:  Game = PRODUCT_WAR3
        Case 7:  Game = PRODUCT_W3XP
        Case 8:  Game = PRODUCT_JSTR
        Case 9:  Game = PRODUCT_SSHR
        Case 10: Game = PRODUCT_DRTL
        Case 11: Game = PRODUCT_DSHR
        Case 12: Game = PRODUCT_CHAT
    End Select

    frmDBManager.m_EntryName = Game
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_W2BN).FullName, , 5)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_STAR).FullName, , 1)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_SEXP).FullName, , 2)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_D2DV).FullName, , 3)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_D2XP).FullName, , 4)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_WAR3).FullName, , 6)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_W3XP).FullName, , 11)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_JSTR).FullName, , 10)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_SSHR).FullName, , 12)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_DRTL).FullName, , 8)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_DSHR).FullName, , 9)
    Call lvGames.ListItems.Add(, , GetProductInfo(PRODUCT_CHAT).FullName, , 7)
End Sub

Private Sub lvGames_DblClick()
    Call cmdOK_Click
End Sub

Private Sub lvGames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdOK_Click
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

