VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About StealthBot"
   ClientHeight    =   5805
   ClientLeft      =   2745
   ClientTop       =   2340
   ClientWidth     =   8775
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4006.714
   ScaleMode       =   0  'User
   ScaleWidth      =   8240.177
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtDescr 
      Height          =   3525
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6218
      _Version        =   393217
      BackColor       =   10040064
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAbout.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   436.659
      X2              =   7662.66
      Y1              =   3105.98
      Y2              =   3105.98
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "StealthBot Contributors"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   4680
      UseMnemonic     =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSpecialThanks 
      BackColor       =   &H00000000&
      Caption         =   "Special thanks to..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   8055
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "StealthBot Wiki"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   4680
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Send Me Email"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   4680
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The StealthBot Website and Support Forum"
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
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      UseMnemonic     =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ OK ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   7200
      TabIndex        =   8
      Top             =   5160
      UseMnemonic     =   0   'False
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   422.573
      X2              =   7662.66
      Y1              =   3095.627
      Y2              =   3095.627
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   7995
   End
   Begin VB.Label lblBottom 
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
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   6855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lNumClicks As Long

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    lblTitle.ForeColor = vbWhite
    lblTitle.Caption = ".: " & CVERSION

    EnableURLDetect txtDescr.hWnd
    #If COMPILE_DEBUG <> 1 Then
        HookWindowProc Me.hWnd
    #End If

    With txtDescr
        .SelTabCount = 1
        .SelTabs(0) = 15 * Screen.TwipsPerPixelX
        .SelHangingIndent = .SelTabs(0)
    End With

    AddLine "The list of current StealthBot project contributors can be found at"
    AddLine "-> http://contributors.stealthbot.net <-"
    AddCrLf
    AddLine "StealthBot is updated and maintained by a team of developers who mostly volunteer their time."
    AddLine "-> The current list of developers can be found at http://devteam.stealthbot.net."
    AddLine "-> Donations to the project are split between developers after expenses."
    AddCrLf
    AddLine "All Blizzard copyrights are (c) 1996-present Blizzard Entertainment."
    AddLine "For a detailed listing of Blizzard copyrights, please visit"
    AddLine "  http://www.blizzard.com/copyright.shtml"
    AddLine "For any further legal information regarding StealthBot and the StealthBot website, visit"
    AddLine "  http://www.stealthbot.net/legal/web/"
    AddCrLf
    AddLine "- Hdx for his continued help and maintenance of the JBLS server at jbls.org"
    AddLine "- The staff and users at StealthBot.net for their continued support"
    AddCrLf
    AddLine "- Thanks to all of the beta testers and people who suggested features to me -- StealthBot wouldn't be possible without you"
    AddLine "- Thanks to Retain, Jack, Hdx, and Berzerker, for their long lists of suggestions and bug reports"
    AddLine "- Thanks to PhiX for his excellent continued support help on StealthBot.net"
    AddCrLf
    AddLine "- My website administrators & moderators, for their continued help in managing the stealthbot.net forums"
    AddCrLf
    AddLine "- And, an extra-special thanks to:"
    AddLine "-> The tech support people on StealthBot.net and in Clan SBS, for giving their time to help others"
    AddLine "-> The people who have donated to the StealthBot project - you can find a current list at http://contributors.stealthbot.net - THANK YOU!"
    
    SetTextSelection txtDescr.hWnd, 0, 0
    ScrollToCaret txtDescr.hWnd
    
    lblBottom.Caption = "(c)2002-2009 Andy T - all rights reserved." & vbCrLf & "Use of this program is subject to the License Agreement found at http://eula.stealthbot.net."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Byte
    For i = lblURL.LBound To lblURL.UBound
        lblURL(i).ForeColor = vbWhite
    Next i
    lblOK.ForeColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect txtDescr.hWnd
    #If COMPILE_DEBUG <> 1 Then
        UnhookWindowProc Me.hWnd
    #End If
End Sub

Private Sub lblOK_Click()
    lNumClicks = 0
    Unload Me
End Sub

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Byte
    For i = lblURL.LBound To lblURL.UBound
        lblURL(i).ForeColor = vbWhite
    Next i
    lblOK.ForeColor = &HFFCC99

    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

Private Sub lblURL_Click(Index As Integer)
    Select Case Index
        Case 0: ShellOpenURL "http://www.stealthbot.net", "the StealthBot Forum"
        Case 1: ShellOpenURL "mailto:stealth@stealthbot.net"
        Case 2: ShellOpenURL "http://contributors.stealthbot.net", "the StealthBot Contributors List"
        Case 3: ShellOpenURL "http://www.stealthbot.net/wiki/Main_Page", "the StealthBot Wiki"
    End Select
End Sub

Private Sub lblURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Byte
    For i = lblURL.LBound To lblURL.UBound
        If i = Index Then
            lblURL(i).ForeColor = &HFFCC99
        Else
            lblURL(i).ForeColor = vbWhite
        End If
    Next i
    lblOK.ForeColor = vbWhite

    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

Private Sub txtDescr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub lblTitle_Click()
    TitleClicked
End Sub

Private Sub lblTitle_DblClick()
    TitleClicked
End Sub

Private Sub AddLine(ByVal sIn As String)
    Dim saElements(0 To 2) As Variant
    saElements(0) = Config.ChatFont
    saElements(1) = vbWhite
    saElements(2) = sIn & vbCrLf
    DisplayRichTextElement txtDescr, saElements(), 0
End Sub

Private Sub AddCrLf()
    Dim saElements(0 To 2) As Variant
    saElements(0) = Config.ChatFont
    saElements(1) = vbWhite
    saElements(2) = vbCrLf
    DisplayRichTextElement txtDescr, saElements(), 0
End Sub

Private Sub TitleClicked()
    lNumClicks = lNumClicks + 1
    
    If lNumClicks Mod 14 = 0 Then
        lblTitle.Caption = " hamtaro rocks :P "
    ElseIf lNumClicks Mod 7 = 0 Then
        lblTitle.Caption = "< Think Outside the Bun >"
    Else
        lblTitle.Caption = ".: " & CVERSION
    End If
End Sub
