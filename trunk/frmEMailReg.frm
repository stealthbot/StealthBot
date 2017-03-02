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
      Cancel          =   -1  'True
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
      Caption         =   "&Never Ask Again"
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
      MaxLength       =   254
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

' this is the main functionality of email registration
' depending on the "Action", (result of clicking a button in the prompt OR the config values)
' will do the specified task, then continue logon sequence
Public Sub DoRegisterEmail(ByVal EMailAction As String, Optional ByVal EMailValue As String = vbNullString)
    Select Case UCase$(EMailAction)
        Case EMAIL_ACT_ASKLATER
            ' "ASKLATER"/ask later: do nothing here
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] E-mail address registration ignored. You may be prompted later."
            
            ContinueLogonSequence
            
        Case EMAIL_ACT_NEVERASK
            ' "NEVERASK"/never ask: register an empty email address
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] E-mail address registration declined."
        
            modBNCS.SEND_SID_SETEMAIL vbNullString
            
            ContinueLogonSequence
        
        Case Else
            ' "VALUE" or "PROMPT" [default behavior]: use the provided value, or prompt with the form if empty
            ' note that "VALUE" and "PROMPT" behave the same
            If LenB(EMailValue) = 0 Then
                ' prompt: show email registration form
                ' this is the default behavior if no config value is specified
                ' (then depending on the user's selection, another one of this functions' actions will happen)
                Show
                
                On Error Resume Next
                txtAddress.SetFocus
            Else
                ' value: send the provided email
                frmChat.AddChat RTBColors.SuccessText, "[BNCS] E-mail address registered."
        
                SEND_SID_SETEMAIL EMailValue
                
                ContinueLogonSequence
            End If
            
    End Select
End Sub

' this function continues logon sequence from where it left off
Private Sub ContinueLogonSequence()
    If g_Connected And Not g_Online Then
        If Dii And BotVars.UseRealm Then
            Call DoQueryRealms
        Else
            Call SendEnterChatSequence
        End If
    End If
    
    ClosedProperly = True
    ds.WaitingForEmail = False
End Sub

Private Sub cmdGo_Click()
    If LenB(txtAddress.Text) > 0 Then
        Call DoRegisterEmail(EMAIL_ACT_VALUE, txtAddress.Text)
        
        Unload Me
    End If
End Sub

Private Sub cmdIgnore_Click()
    Call DoRegisterEmail(EMAIL_ACT_NEVERASK)
    
    Unload Me
End Sub

Private Sub cmdAskLater_Click()
    Call DoRegisterEmail(EMAIL_ACT_ASKLATER)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ClosedProperly = False
    
    Label1.Caption = "Battle.net would like to know if you want to register an e-mail address " & _
                    "with your account. If you want to do so, type a valid e-mail address in the box " & _
                    "below. If you don't want to register an e-mail address, click ""Never Ask Again""." & _
                    "To be asked again on your next logon, click ""Ask Me Later"". ""OK"" and ""Never Ask Again"" are permanent."
                    
    Label1.Caption = Label1.Caption & vbCrLf & vbCrLf & "Choose an option below to proceed."
    
    Label2.Caption = "Click here for more information."
    
    txtAddress.Text = vbNullString
    cmdGo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ClosedProperly And g_Online Then
        Call DoRegisterEmail(EMAIL_ACT_ASKLATER)
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdGo_Click
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdAskLater_Click
    End If
End Sub

Private Sub txtAddress_Change()
    cmdGo.Enabled = (LenB(txtAddress.Text) > 0)
End Sub
