VERSION 5.00
Begin VB.Form frmClanInvite 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warcraft III Clan Invitation"
   ClientHeight    =   1905
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDecline 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
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
      TabIndex        =   4
      Top             =   120
      Width           =   3015
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
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
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
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmClanInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
    frmChat.AddChat RTBColors.SuccessText, "[CLAN] Invitation accepted."
    
    With PBuffer
        .InsertNonNTString Clan.Token
        .InsertNonNTString Clan.DWName
        .InsertNTString Clan.Creator
        .InsertByte &H6
    
        If Clan.isNew = 1 Then
            .SendPacket SID_CLANCREATIONINVITATION
            Clan.isNew = 0
        Else
            .SendPacket SID_CLANINVITATIONRESPONSE
        End If
    End With
    AwaitingClanMembership = 1
    
    Unload Me
End Sub

Sub cmdDecline_Click()
    frmChat.AddChat RTBColors.ErrorMessageText, "[CLAN] Invitation declined."
    
    With PBuffer
        .InsertNonNTString Clan.Token
        .InsertNonNTString Clan.DWName
        .InsertNTString Clan.Creator
        .InsertByte &H4
    
        If Clan.isNew = 1 Then
            .SendPacket SID_CLANCREATIONINVITATION
            Clan.isNew = 0
        Else
            .SendPacket SID_CLANINVITATIONRESPONSE
        End If
    End With
    
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
