VERSION 5.00
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
   Begin VB.TextBox txtDescr 
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
      Height          =   3525
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   840
      Width           =   7935
   End
   Begin VB.Label lblURL 
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
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
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
      Index           =   5
      Left            =   4560
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label lblURL 
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
      Index           =   4
      Left            =   4800
      TabIndex        =   8
      Top             =   4920
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
      Left            =   120
      TabIndex        =   7
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   8055
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Online ReadMe"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Send Me E-mail"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblURL 
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
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
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
      TabIndex        =   1
      Top             =   5160
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7995
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
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   6975
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
    txtDescr.ForeColor = vbWhite
    lblTitle.Caption = ".: " & CVERSION
    
    txtDescr.Locked = True
    
    AddLine "The list of current StealthBot project contributors can be found at"
    AddLine "-> http://contributors.stealthbot.net <-"
    AddCrLf
    AddLine "All Blizzard copyrights are (c) 1996-2007 Blizzard Entertainment."
    AddLine "For a detailed listing of Blizzard copyrights, please visit"
    AddLine "  http://www.blizzard.com/copyright.shtml"
    AddLine "For any further legal information and copyrights regarding StealthBot and the StealthBot website, visit"
    AddLine "  http://www.stealthbot.net/legal/web/"
    AddCrLf
    AddLine "- Grok[vL] - AddChat and DebugOutput are invaluable"
    AddLine "- Skywing[vL] and Yoni[vL] for their development and maintenance of the BNLS system"
    AddLine "- Hdx for his maintenance of the JBLS server at jbls.org"
    AddLine "- API information from NSDN for Winamp features"
    AddLine "- Barumonk and Pr0pHeT]ZeR0[ for their help during my struggle with 0x51 :)"
    AddCrLf
    AddLine "- 6th period Intro to Programming, sophomore year in high school, and my teacher Mr. P"
    AddCrLf
    AddLine "- Thanks to all of my beta testers and people who suggested features to me -- StealthBot wouldn't be possible without you, you know who you are"
    AddLine "- Thanks to Retain, Jack, Hdx, and Berzerker, for their long lists of suggestions and bug reports"
    AddLine "- Thanks to UserLoser and Eric for their help with Diablo II Realm-related things"
    AddCrLf
    AddLine "- My website administrators, liQuid, MetaLMilitiA and Eric for their continued help in managing the stealthbot.net forums"
    AddLine "- Eric, DaRk-FeAnOr and Ethereal for their help regarding Warcraft III"
    AddLine "- Arta[vL]'s BNetDocs and BNU-Camel's OpenBNetDocs wiki, as references for all things BNCS-related"
    AddCrLf
    AddLine "- InternationaL, for his help with the CHM readme"
    AddCrLf
    AddLine "- LW-Killbound for the splash screen image"
    AddLine "- ZeKe for getting me the account Stealth@USEast"
    AddLine "- Thanks to Jake and Marc for getting me Stealth@Azeroth"
    AddCrLf
    AddLine "- And, an extra-special thanks to:"
    AddLine "-> The tech support people on StealthBot.net and in Clan SBS, for giving their time"
    AddLine "-> The people who have donated to the StealthBot project - you can find a current list at http://contributors.stealthbot.net - THANK YOU!"
                             
    lblBottom.Caption = "(c)2002-2007 Andy T (" & Chr(34) & "Stealth@USEast" & Chr(34) & ") - All Rights Reserved." & vbCrLf & "Use of this program is subject to the License Agreement found at http://eula.stealthbot.net."
End Sub

Private Sub AddLine(ByVal sIn As String)
    txtDescr.text = txtDescr.text & sIn & vbCrLf
End Sub

Private Sub AddCrLf()
    txtDescr.text = txtDescr.text & vbCrLf
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    For i = 0 To 6
        lblURL(i).ForeColor = vbWhite
    Next i
End Sub

Private Sub lblOK_Click()
    lNumClicks = 0
    Unload Me
End Sub

Private Sub lblTitle_Click()
    TitleClicked
End Sub

Private Sub lblTitle_DblClick()
    TitleClicked
End Sub

Sub TitleClicked()
    lNumClicks = lNumClicks + 1
    
    If lNumClicks Mod 14 = 0 Then
        lblTitle.Caption = " hamtaro rocks :P "
    ElseIf lNumClicks Mod 7 = 0 Then
        lblTitle.Caption = "< Think Outside the Bun >"
    Else
        lblTitle.Caption = ".: " & CVERSION
    End If
End Sub

Private Sub lblURL_Click(Index As Integer)
    If Index = 3 Then
        ShellExecute Me.hWnd, "Open", "mailto:stealth@stealthbot.net", 0&, 0&, 0&
    ElseIf Index = 0 Then
        ShellExecute Me.hWnd, "Open", "http://www.stealthbot.net", 0&, 0&, 0&
    ElseIf Index = 6 Then
        ShellExecute Me.hWnd, "Open", "http://www.stealthbot.net/readme/", 0&, 0&, 0&
    ElseIf Index = 5 Then
        ShellExecute Me.hWnd, "Open", "http://contributors.stealthbot.net", 0&, 0&, 0&
    End If
End Sub

Private Sub lblURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    For i = 0 To 6
        lblURL(i).ForeColor = vbWhite
    Next i
    lblURL(Index).ForeColor = vbBlue
End Sub
