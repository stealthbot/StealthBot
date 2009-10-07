VERSION 5.00
Begin VB.Form frmEMailReg 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail Registration"
   ClientHeight    =   2640
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAskLater 
      Caption         =   "&Ask Me Later"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&OK"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtAddress 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "click here for more information"
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
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmEMailReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StealthBot EMail Registration Form
'stealth@stealthbot.net
Option Explicit
Private ClosedProperly As Boolean

Private Sub CloseEmailReg()
    If Not g_Online Then
        If Dii And BotVars.UseRealm Then
            frmRealm.Show
        Else
            Call SendEnterChatSequence
        End If
    End If
    
    ClosedProperly = True
    ds.WaitingForEmail = False
End Sub

Private Sub cmdGo_Click()

    modBNCS.SEND_SID_SETEMAIL txtAddress.Text
    
    CloseEmailReg
    
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    
    modBNCS.SEND_SID_SETEMAIL vbNullString
    
    CloseEmailReg
    
    Unload Me
End Sub

Private Sub cmdAskLater_Click()
    frmChat.AddChat RTBColors.SuccessText, ">> E-mail registration ignored."
    
    CloseEmailReg
    
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ClosedProperly = False
    
    Label1.Caption = "Battle.net would like to know if you want to register an e-mail address " & _
                    "with your account. If you want to do so, type a valid e-mail address in the box " & _
                    "below. If you don't want to register an e-mail address, click Ignore." & _
                    "To be asked again on your next login, click Ask Me Later. OK and Ignore are permanent."
                    
    Label1.Caption = Label1.Caption & vbCrLf & vbCrLf & "Choose an option below to proceed."
    
    Label2.Caption = "Click here for more information."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ClosedProperly Then
        frmChat.AddChat RTBColors.SuccessText, ">> E-mail registration ignored."
        CloseEmailReg
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdGo_Click
    End If
End Sub
