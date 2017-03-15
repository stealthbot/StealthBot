VERSION 5.00
Begin VB.Form frmClanInvite 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warcraft III Clan Invitation"
   ClientHeight    =   1905
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDecline 
      Cancel          =   -1  'True
      Caption         =   "&Decline"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
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
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00000000&
      Caption         =   "A. Random Person"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblClan 
      BackColor       =   &H00000000&
      Caption         =   "Clan %clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblInvite 
      BackColor       =   &H00000000&
      Caption         =   "has invited you to join"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmClanInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
    Call modWar3Clan.InvitationResponse(ClanResponseAccept, Clan.Token, Clan.DWName, Clan.Creator, Clan.IsNew)
    Clan.IsNew = False

    frmChat.AddChat RTBColors.SuccessText, "[CLAN] Invitation accepted."
    AwaitingClanMembership = 1

    Unload Me
End Sub

Sub cmdDecline_Click()
    Call modWar3Clan.InvitationResponse(ClanResponseDecline, Clan.Token, Clan.DWName, Clan.Creator, Clan.IsNew)
    Clan.IsNew = False

    frmChat.AddChat RTBColors.ErrorMessageText, "[CLAN] Invitation declined."

    Unload Me
End Sub

Private Sub Form_Load()
    lblUser.Caption = Clan.Creator
    lblClan.Caption = "Clan " & StrReverse(Clan.DWName)
    Me.Icon = frmChat.Icon
    cmdAccept.Enabled = False
    cmdDecline.Enabled = False
    
    ClanAcceptTimerID = SetTimer(frmClanInvite.hWnd, 0, 2000, AddressOf ClanInviteTimerProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ClanAcceptTimerID > 0 Then
        KillTimer 0, ClanAcceptTimerID
    End If
End Sub
