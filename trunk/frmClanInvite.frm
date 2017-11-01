VERSION 5.00
Begin VB.Form frmClanInvite 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warcraft III Clan Invitation"
   ClientHeight    =   2145
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
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
      TabIndex        =   6
      Top             =   1680
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
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Timer"
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
      TabIndex        =   4
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblInv 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Inviter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Clan Full Name"
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
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblClan 
      Alignment       =   2  'Center
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
      TabIndex        =   1
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblInvite 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You have been invited to join"
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
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmClanInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REQUEST_COOKIE As Integer = 0
Private Const REQUEST_TAG    As Integer = 1
Private Const REQUEST_NAME   As Integer = 2
Private Const REQUEST_INV    As Integer = 3
Private Const REQUEST_ISNEW  As Integer = 4

Private Const CLAN_TIMER_TIMEOUT As Integer = 30
Private Const CLAN_BUTTON_TIMEOUT As Integer = 28

Private Ticks As Integer
Private vArray() As Variant

Private Sub cmdAccept_Click()
    ' send invitation response
    Call frmChat.ClanHandler.InvitationResponse( _
            CBool(vArray(REQUEST_ISNEW)), CLng(vArray(REQUEST_COOKIE)), _
            CStr(vArray(REQUEST_TAG)), CStr(vArray(REQUEST_INV)), _
            clresAccept)

    frmChat.AddChat RTBColors.SuccessText, "[CLAN] Invitation accepted."

    Unload Me
End Sub

Private Sub cmdDecline_Click()
    ' send invitation response
    Call frmChat.ClanHandler.InvitationResponse( _
            CBool(vArray(REQUEST_ISNEW)), CLng(vArray(REQUEST_COOKIE)), _
            CStr(vArray(REQUEST_TAG)), CStr(vArray(REQUEST_INV)), _
            clresDecline)

    frmChat.AddChat RTBColors.ErrorMessageText, "[CLAN] Invitation declined."

    Unload Me
End Sub

Private Sub Form_Load()
    Dim oRequest As udtServerRequest

    If (g_Clan.PendingInvitation) Then
        g_Clan.PendingInvitation = False
        If FindServerRequest(oRequest, g_Clan.PendingInvitationCookie, SID_CLANINVITATIONRESPONSE) Then
            ' set variables for instance
            vArray = oRequest.Tag
            Ticks = CLAN_TIMER_TIMEOUT

            ' set up appearance
            Me.Icon = frmChat.Icon

            lblClan.Caption = StringFormat("Clan {0}", CStr(vArray(REQUEST_TAG)))
            lblClan.ForeColor = RTBColors.JoinUsername
            lblName.Caption = CStr(vArray(REQUEST_NAME))
            lblInv.Caption = StringFormat("Invited by:  {0}", CStr(vArray(REQUEST_INV)))

            ' disable buttons
            cmdAccept.Enabled = False
            cmdDecline.Enabled = False

            ' tick once and enable 1-second timer
            Call tmrTimeout_Timer
            tmrTimeout.Enabled = True

            Exit Sub
        End If
    End If

    ' no pending invite...
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrTimeout.Enabled = False
End Sub

Private Sub tmrTimeout_Timer()
    ' re-enable buttons?
    If Ticks = CLAN_BUTTON_TIMEOUT Then
        cmdAccept.Enabled = True
        cmdDecline.Enabled = True
    End If

    ' time out?
    If Ticks <= 0 Then
        cmdDecline_Click
        Exit Sub
    End If

    ' set timer caption
    lblTimer.Caption = StringFormat("Invitation expires in {0} second{1}...", _
            Ticks, IIf(Ticks <> 1, "s", vbNullString))

    ' count down
    Ticks = Ticks - 1
End Sub

