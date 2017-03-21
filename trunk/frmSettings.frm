VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StealthBot Settings"
   ClientHeight    =   5310
   ClientLeft      =   1575
   ClientTop       =   1935
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cboProfile 
      Appearance      =   0  'Flat
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
      Left            =   120
      TabIndex        =   0
      Text            =   "Profile Selector"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "&Forum"
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
      Left            =   3840
      TabIndex        =   211
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdReadme 
      Caption         =   "&Wiki"
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
      Left            =   3000
      TabIndex        =   210
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdStepByStep 
      Caption         =   "&Step-By-Step Configuration"
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
      Left            =   4800
      TabIndex        =   212
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Apply and Cl&ose"
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
      Left            =   8280
      TabIndex        =   213
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   7440
      TabIndex        =   214
      Top             =   4920
      Width           =   855
   End
   Begin vbalTreeViewLib6.vbalTreeView tvw 
      Height          =   4620
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8149
      BackColor       =   10040064
      ForeColor       =   16777215
      Style           =   2
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
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   8
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      Begin VB.Label lblSplash 
         BackColor       =   &H00000000&
         Caption         =   "Splash message container label."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   6015
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   31
         X1              =   360
         X2              =   6240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Welcome to StealthBot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCDKey 
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
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox chkSHR 
         BackColor       =   &H00000000&
         Caption         =   "S&hareware"
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
         Left            =   4800
         TabIndex        =   32
         Top             =   3000
         Width           =   1212
      End
      Begin VB.CheckBox chkJPN 
         BackColor       =   &H00000000&
         Caption         =   "Japanese"
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
         Left            =   4800
         TabIndex        =   33
         Top             =   3240
         Width           =   1212
      End
      Begin VB.OptionButton optDRTL 
         BackColor       =   &H00000000&
         Caption         =   "Diablo"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3240
         Width           =   1455
      End
      Begin VB.OptionButton optW3XP 
         BackColor       =   &H00000000&
         Caption         =   "The Frozen Throne"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton optSTAR 
         BackColor       =   &H00000000&
         Caption         =   "StarCraft"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1800
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optSEXP 
         BackColor       =   &H00000000&
         Caption         =   "Brood War"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton optD2DV 
         BackColor       =   &H00000000&
         Caption         =   "Diablo II"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optD2XP 
         BackColor       =   &H00000000&
         Caption         =   "Lord of Destruction"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton optW2BN 
         BackColor       =   &H00000000&
         Caption         =   "WarCraft II BNE"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton optWAR3 
         BackColor       =   &H00000000&
         Caption         =   "WarCraft III"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   240
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtHomeChan 
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
         Left            =   240
         MaxLength       =   31
         TabIndex        =   19
         ToolTipText     =   "This is the channel that the bot will attempt to join when it logs on."
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox cboServer 
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
         Height          =   315
         Left            =   3240
         TabIndex        =   21
         Text            =   "Choose One"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtExpKey 
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
         Height          =   285
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Only required for Lord of Destruction and The Frozen Throne."
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CheckBox chkUseRealm 
         BackColor       =   &H00000000&
         Caption         =   "Use Diablo II &Realm"
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
         Left            =   3240
         TabIndex        =   35
         ToolTipText     =   "Allows you to log on as your Diablo II character as well as enter realm-restricted channels."
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CheckBox chkSpawn 
         BackColor       =   &H00000000&
         Caption         =   "Use Key as Sp&awned Client"
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
         Left            =   3240
         TabIndex        =   34
         Top             =   3600
         Width           =   3252
      End
      Begin VB.Label lblAccountManage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "( manage account )"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblAddCurrentKey 
         BackColor       =   &H00000000&
         Caption         =   "( add current )"
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblManageKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "( manage )"
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
         Left            =   1920
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblBasic 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   4800
         TabIndex        =   23
         Top             =   1680
         Width           =   735
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   4680
         X2              =   4800
         Y1              =   2640
         Y2              =   2760
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   4680
         X2              =   4800
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   4680
         X2              =   4800
         Y1              =   1920
         Y2              =   2040
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   3000
         X2              =   3000
         Y1              =   960
         Y2              =   4080
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "Pro&duct"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   1560
         Width           =   615
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "Ser&ver"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "&Username"
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "&Password"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "CD&Key"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "&Home Channel"
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
         Left            =   240
         TabIndex        =   18
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblBasic 
         BackColor       =   &H00000000&
         Caption         =   "E&xpansion CDKey"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Basic configuration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   2
      Left            =   3000
      TabIndex        =   36
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkShowNews 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show StealthBot news on startup"
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
         Left            =   3480
         TabIndex        =   39
         ToolTipText     =   "Show the StealthBot news in the chat window after startup."
         Top             =   840
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkMinimizeOnStartup 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Minimize on startup"
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
         Left            =   360
         TabIndex        =   40
         ToolTipText     =   "Automatically minimize on startup."
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox chkURLDetect 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Enable URL detection"
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
         Left            =   3480
         TabIndex        =   45
         ToolTipText     =   "Enables automatic URL detection and highlighting in the main chat window."
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CheckBox chkShowPrefixBox 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show prefix box"
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
         Left            =   360
         TabIndex        =   51
         ToolTipText     =   "Enables the small prefix box to the left of the send box used to put text before all messages."
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkShowSuffixBox 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show suffix box"
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
         Left            =   360
         TabIndex        =   52
         ToolTipText     =   "Enables the small suffix box to the right of the send box used to put text after all messages."
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkShowUserGameStatsIcons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show user game stats icons"
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
         Left            =   3480
         TabIndex        =   49
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkShowUserFlagsIcons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show user flag-based icons"
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
         Left            =   3480
         TabIndex        =   50
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkListColoring 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Enable channel list name coloring"
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
         Left            =   360
         TabIndex        =   48
         ToolTipText     =   "Name coloring changes the color of people's names in the channel list based on their status or activity."
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkNameAutocomplete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Enable name autocompletion"
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
         Left            =   3480
         TabIndex        =   47
         ToolTipText     =   "Enables the highlighted display of suggested usernames as you type in the send box."
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkMinimizeToTray 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Always minimize to System Tray"
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
         Left            =   3480
         TabIndex        =   41
         ToolTipText     =   "Always minimize to the System Tray instead of the Taskbar."
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkFlash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Flash window on events"
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
         Left            =   360
         TabIndex        =   46
         ToolTipText     =   "Flash the bot window when events occur."
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox chkUTF8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Use UTF-8 encoding/decoding"
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
         Left            =   360
         TabIndex        =   42
         ToolTipText     =   "Blizzard games encode their messages to UTF-8 format. Enable this setting to properly see special characters sent by these games."
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.ComboBox cboTimestamp 
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
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4080
         Width           =   3495
      End
      Begin VB.CheckBox chkSplash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show splash screen on startup"
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
         Left            =   360
         TabIndex        =   38
         ToolTipText     =   "Show the splash screen on startup."
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox chkFilter 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Use chat filtering"
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
         Left            =   3480
         TabIndex        =   43
         ToolTipText     =   "Enable chat filtering (lowers CPU usage)."
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox chkJoinLeaves 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show join/leave notifications"
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
         Left            =   360
         TabIndex        =   44
         ToolTipText     =   "Enable Join and Leave messages."
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   2520
         X2              =   2520
         Y1              =   3720
         Y2              =   4680
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   6240
         X2              =   360
         Y1              =   3560
         Y2              =   3560
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   6240
         X2              =   360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   360
         X2              =   6240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblInterface 
         BackColor       =   &H00000000&
         Caption         =   "&Timestamp"
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
         Left            =   2760
         TabIndex        =   53
         Top             =   3840
         Width           =   975
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "General interface settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   37
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   3
      Left            =   3000
      TabIndex        =   55
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdSaveColor 
         Caption         =   "Sa&ve changes to this color"
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
         Left            =   4200
         TabIndex        =   89
         Top             =   2160
         Width           =   2055
      End
      Begin VB.ComboBox cboColorList 
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
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox txtValue 
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
         Left            =   360
         TabIndex        =   71
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdColorPicker 
         Caption         =   "Color Picker"
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
         Left            =   360
         TabIndex        =   74
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txtR 
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
         Left            =   600
         TabIndex        =   76
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtG 
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
         Left            =   1560
         TabIndex        =   78
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtB 
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
         Left            =   2520
         TabIndex        =   80
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdGetRGB 
         Caption         =   "Generate New Value from RGB"
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
         Left            =   360
         TabIndex        =   85
         Top             =   3720
         Width           =   2775
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import ColorList"
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
         Left            =   3240
         TabIndex        =   86
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdDefaults 
         Caption         =   "Restore &Default Colors"
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
         Left            =   4200
         TabIndex        =   67
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export ColorList"
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
         Left            =   4800
         TabIndex        =   87
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtHTML 
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   83
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdHTMLGen 
         Caption         =   "Generate"
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
         Left            =   5160
         TabIndex        =   84
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtChanFont 
         Alignment       =   2  'Center
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
         Left            =   1920
         TabIndex        =   64
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtChatFont 
         Alignment       =   2  'Center
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
         Left            =   1920
         TabIndex        =   59
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtChanSize 
         Alignment       =   1  'Right Justify
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   66
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtChatSize 
         Alignment       =   1  'Right Justify
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   61
         Top             =   840
         Width           =   615
      End
      Begin MSComDlg.CommonDialog cDLG 
         Left            =   5400
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblCurrentValue 
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
         Left            =   2040
         TabIndex        =   72
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblEg 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3360
         TabIndex        =   73
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "Color to modify:"
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
         Left            =   360
         TabIndex        =   68
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "New Value:                   Current Value:       Example:"
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
         Left            =   360
         TabIndex        =   70
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "R:"
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
         Left            =   360
         TabIndex        =   75
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "G:"
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
         Left            =   1320
         TabIndex        =   77
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "B:"
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
         Left            =   2280
         TabIndex        =   79
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Caption         =   "Use HTML hexadecimal colors:"
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
         Left            =   3360
         TabIndex        =   81
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblColor 
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
         Index           =   6
         Left            =   3360
         TabIndex        =   82
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "Size:"
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
         Left            =   4200
         TabIndex        =   60
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "Channel &List"
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
         Left            =   360
         TabIndex        =   62
         ToolTipText     =   "Changes the font settings for the channel list."
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "C&hat Window"
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
         Left            =   360
         TabIndex        =   57
         ToolTipText     =   "Changes the font setting for the main chat window."
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "Font:"
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
         Left            =   1440
         TabIndex        =   58
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "Size:"
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
         Left            =   4200
         TabIndex        =   65
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblFont 
         BackColor       =   &H00000000&
         Caption         =   "Font:"
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
         Left            =   1440
         TabIndex        =   63
         Top             =   1200
         Width           =   495
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Interface font and color settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   56
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblColorStatus 
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
         Left            =   360
         TabIndex        =   88
         Top             =   4200
         Width           =   5895
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   4
      Left            =   3000
      TabIndex        =   90
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkPhraseKick 
         BackColor       =   &H00000000&
         Caption         =   "Kick"
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
         Left            =   2160
         TabIndex        =   99
         ToolTipText     =   "Instead of banning users for saying banned phrases, the bot will kick them."
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox chkQuietKick 
         BackColor       =   &H00000000&
         Caption         =   "Kick"
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
         Left            =   2160
         TabIndex        =   97
         ToolTipText     =   "Instead of banning quiet users, the bot will kick them."
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtPingLevel 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   102
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox chkBanEvasion 
         BackColor       =   &H00000000&
         Caption         =   "Use Ban Evasion"
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
         Left            =   360
         TabIndex        =   95
         ToolTipText     =   "Ban Evasion attempts to keep people who are banned out of your channel."
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkIdleKick 
         BackColor       =   &H00000000&
         Caption         =   "Kick"
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
         Left            =   4680
         TabIndex        =   108
         ToolTipText     =   "Instead of banning idle users, the bot will kick them."
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkPeonbans 
         BackColor       =   &H00000000&
         Caption         =   "Ban Warcraft III Peons"
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
         Left            =   3240
         TabIndex        =   119
         ToolTipText     =   "Ban Warcraft III users who have the Peon icon."
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtLevelBanMsg 
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
         Left            =   3240
         MaxLength       =   180
         TabIndex        =   126
         Text            =   "You are below the required level for entry."
         Top             =   4305
         Width           =   3135
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "The Frozen Throne"
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
         TabIndex        =   117
         Top             =   2400
         Width           =   1812
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Lord of Destruction"
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
         Left            =   4560
         TabIndex        =   115
         Top             =   2160
         Width           =   1812
      End
      Begin VB.CheckBox chkPhrasebans 
         BackColor       =   &H00000000&
         Caption         =   "Enable Phrasebans"
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
         Left            =   360
         TabIndex        =   98
         ToolTipText     =   "Ban unsafelisted users who say banned phrases."
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkIPBans 
         BackColor       =   &H00000000&
         Caption         =   "Enable IPBanning"
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
         Left            =   360
         TabIndex        =   92
         ToolTipText     =   "Ban squelched users."
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Starcraft"
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
         Left            =   3240
         TabIndex        =   112
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Brood War"
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
         Left            =   4560
         TabIndex        =   113
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Diablo II"
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
         Left            =   3240
         TabIndex        =   114
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Warcraft II"
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
         Left            =   3240
         TabIndex        =   118
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Warcraft III"
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
         Left            =   3240
         TabIndex        =   116
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkQuiet 
         BackColor       =   &H00000000&
         Caption         =   "Enable Quiet-Time"
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
         Left            =   360
         TabIndex        =   96
         ToolTipText     =   "Ban unsafelisted users that talk."
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtProtectMsg 
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
         Left            =   360
         MaxLength       =   180
         TabIndex        =   106
         Text            =   "Lockdown Enabled"
         Top             =   4305
         Width           =   2415
      End
      Begin VB.CheckBox chkProtect 
         BackColor       =   &H00000000&
         Caption         =   "Enable Channel Protection"
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
         Left            =   360
         TabIndex        =   104
         ToolTipText     =   "Ban unsafelisted users who join the channel."
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CheckBox chkKOY 
         BackColor       =   &H00000000&
         Caption         =   "Enable Kick-On-Yell"
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
         Left            =   360
         TabIndex        =   94
         ToolTipText     =   "Kick users who yell (uppercase message longer than 5 letters)."
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtBanW3 
         Alignment       =   1  'Right Justify
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
         Left            =   5760
         MaxLength       =   25
         TabIndex        =   124
         Top             =   3720
         Width           =   495
      End
      Begin VB.CheckBox chkPlugban 
         BackColor       =   &H00000000&
         Caption         =   "Enable PlugBans"
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
         Left            =   360
         TabIndex        =   93
         ToolTipText     =   "Ban users with a UDP plug."
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtBanD2 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         MaxLength       =   25
         TabIndex        =   122
         Top             =   3720
         Width           =   495
      End
      Begin VB.CheckBox chkIdlebans 
         BackColor       =   &H00000000&
         Caption         =   "Ban idle users"
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
         Left            =   3240
         TabIndex        =   107
         ToolTipText     =   "Ban users who have been idle for the specified number of seconds."
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtIdleBanDelay 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         MaxLength       =   25
         TabIndex        =   110
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox chkPingBan 
         BackColor       =   &H00000000&
         Caption         =   "PingBan"
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
         Left            =   360
         TabIndex        =   100
         ToolTipText     =   "Ban unsafelisted users who have a ping, in milliseconds, higher than specified."
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Above:"
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
         Left            =   1320
         TabIndex        =   101
         Top             =   3120
         Width           =   615
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   3000
         X2              =   3000
         Y1              =   4560
         Y2              =   720
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "ms"
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
         Left            =   2760
         TabIndex        =   103
         Top             =   3120
         Width           =   255
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   2880
         X2              =   240
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   3120
         X2              =   6360
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   3120
         X2              =   6360
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Levelban message"
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
         Left            =   3240
         TabIndex        =   125
         Top             =   4065
         Width           =   1335
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "Seconds before ban:"
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
         Left            =   3600
         TabIndex        =   109
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "Clientbans"
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
         Left            =   3240
         TabIndex        =   111
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "Message"
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
         Left            =   360
         TabIndex        =   105
         ToolTipText     =   "Shorter is better"
         Top             =   4065
         Width           =   1935
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "LevelBans: Set to 0 to disable."
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
         Left            =   3240
         TabIndex        =   120
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Diablo II:"
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
         Left            =   3240
         TabIndex        =   121
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Warcraft III:"
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
         Left            =   4680
         TabIndex        =   123
         Top             =   3720
         Width           =   975
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "General moderation settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   91
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   6
      Left            =   3000
      TabIndex        =   127
      Top             =   0
      Width           =   6615
      Begin VB.OptionButton optMsg 
         BackColor       =   &H00000000&
         Caption         =   "Message"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optQuote 
         BackColor       =   &H00000000&
         Caption         =   "Quote"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton optUptime 
         BackColor       =   &H00000000&
         Caption         =   "Uptime"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optMP3 
         BackColor       =   &H00000000&
         Caption         =   "MP3"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkIdles 
         BackColor       =   &H00000000&
         Caption         =   "Show anti-&idle messages"
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
         Left            =   360
         TabIndex        =   129
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtIdleMsg 
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
         Left            =   360
         TabIndex        =   139
         Text            =   "/me is a %ver"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtIdleWait 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   136
         Text            =   "6"
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "minutes"
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
         Left            =   3000
         TabIndex        =   137
         Top             =   1200
         Width           =   735
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   3840
         X2              =   5280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   21
         X1              =   5040
         X2              =   5280
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   20
         X1              =   5040
         X2              =   5280
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   19
         X1              =   5040
         X2              =   5280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   5040
         X2              =   5040
         Y1              =   1080
         Y2              =   2160
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Idle message type:"
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
         Left            =   3840
         TabIndex        =   130
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "&Delay between messages:"
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
         Left            =   360
         TabIndex        =   135
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Idle &message"
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
         Left            =   360
         TabIndex        =   138
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblIdleVars 
         BackColor       =   &H00000000&
         Caption         =   "idle variable container label"
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
         Height          =   1935
         Left            =   360
         TabIndex        =   140
         Top             =   2280
         Width           =   5895
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Idle message settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   128
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   7
      Left            =   3000
      TabIndex        =   141
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkClanInvitations 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Allow Clan Invitations"
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
         Left            =   360
         TabIndex        =   146
         ToolTipText     =   "If not in a WarCraft III clan, allow the bot to process invitations sent to you."
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtOwner 
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
         Left            =   360
         MaxLength       =   30
         TabIndex        =   148
         ToolTipText     =   "This account has full control over the bot. Use carefully!"
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtTrigger 
         Alignment       =   2  'Center
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
         Left            =   1920
         TabIndex        =   150
         Text            =   "."
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox chkD2Naming 
         BackColor       =   &H00000000&
         Caption         =   "Use &Diablo II naming conventions"
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
         Left            =   3480
         TabIndex        =   176
         ToolTipText     =   "Show usernames with Diablo II naming conventions."
         Top             =   4320
         Width           =   2895
      End
      Begin VB.OptionButton optNaming 
         BackColor       =   &H00000000&
         Caption         =   "Show all"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "Show usernames with all gateways."
         Top             =   3960
         Width           =   1215
      End
      Begin VB.OptionButton optNaming 
         BackColor       =   &H00000000&
         Caption         =   "WarCraft III"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Show usernames as they would appear to a WarCraft III user."
         Top             =   3960
         Width           =   1215
      End
      Begin VB.OptionButton optNaming 
         BackColor       =   &H00000000&
         Caption         =   "Legacy"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   158
         ToolTipText     =   "Show usernames as they would appear to a StarCraft or WarCraft II user."
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optNaming 
         BackColor       =   &H00000000&
         Caption         =   "Default"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Show usernames as they would appear to the selected client."
         Top             =   3600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox chkShowOffline 
         BackColor       =   &H00000000&
         Caption         =   "Show offline friends"
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
         Left            =   3480
         TabIndex        =   152
         ToolTipText     =   "Determines whether or not offline friends are hidden from the /f list command."
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtBackupChan 
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
         Left            =   3480
         MaxLength       =   31
         TabIndex        =   155
         ToolTipText     =   "The channel to go to when kicked. Leave blank to stay in The Void."
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CheckBox chkAllowMP3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Allow MP3 commands"
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
         Left            =   360
         TabIndex        =   143
         ToolTipText     =   "Allow commands such as .next and .back that change your current Winamp/iTunes song."
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkWhisperCmds 
         BackColor       =   &H00000000&
         Caption         =   "Whisper command responses"
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
         Left            =   3480
         TabIndex        =   151
         ToolTipText     =   "Whisper the return messages of all bot commands."
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkPAmp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Use ProfileAmp"
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
         Left            =   360
         TabIndex        =   144
         ToolTipText     =   "ProfileAmp writes Winamp's currently played song to your profile every 30 seconds."
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkMail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Enable Mail System"
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
         Left            =   360
         TabIndex        =   145
         ToolTipText     =   "Check the mail.ini file when users join."
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkBackup 
         BackColor       =   &H00000000&
         Caption         =   "Join a backup channel when kicked"
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
         Left            =   3480
         TabIndex        =   153
         ToolTipText     =   "The bot will join a specified channel when kicked, instead of rejoining."
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lblMisc 
         BackColor       =   &H00000000&
         Caption         =   "Bot &Owner"
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
         Left            =   360
         TabIndex        =   147
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblMisc 
         BackColor       =   &H00000000&
         Caption         =   "Command &trigger:"
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
         Left            =   480
         TabIndex        =   149
         ToolTipText     =   "The command trigger is used to identify a chat message as a command."
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Gateway naming convention"
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
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   156
         Top             =   3360
         Width           =   2040
      End
      Begin VB.Label lblMisc 
         BackColor       =   &H00000000&
         Caption         =   "Backup &channel"
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
         Left            =   3480
         TabIndex        =   154
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   23
         X1              =   3120
         X2              =   3120
         Y1              =   840
         Y2              =   4560
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   22
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Miscellaneous general settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   142
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   1
      Left            =   3000
      TabIndex        =   161
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkProxyBNLS 
         BackColor       =   &H00000000&
         Caption         =   "Proxy BNLS"
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
         Left            =   4560
         TabIndex        =   187
         ToolTipText     =   "Routes the BNLS connection through the same proxy."
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox chkProxyMCP 
         BackColor       =   &H00000000&
         Caption         =   "Proxy Diablo II Realm"
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
         Left            =   4560
         TabIndex        =   188
         ToolTipText     =   "Routes a Diablo II Realm connection through the same proxy."
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CheckBox chkConnectOnStartup 
         BackColor       =   &H00000000&
         Caption         =   "Co&nnect on startup"
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
         Left            =   240
         TabIndex        =   169
         ToolTipText     =   "Automatically connect when the bot starts up."
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   3120
         TabIndex        =   179
         ToolTipText     =   "Created accounts will automatically be registered with this address. Leave blank to prompt each time."
         Top             =   2880
         Width           =   3015
      End
      Begin VB.ComboBox cboBNLSServer 
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
         Height          =   315
         ItemData        =   "frmSettings.frx":0000
         Left            =   2400
         List            =   "frmSettings.frx":0002
         TabIndex        =   166
         Text            =   "cboBNLSServer"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtReconDelay 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   171
         Text            =   "1000"
         Top             =   3000
         Width           =   615
      End
      Begin VB.OptionButton optSocks5 
         BackColor       =   &H00000000&
         Caption         =   "SOCKS&5"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   3480
         Width           =   735
      End
      Begin VB.OptionButton optSocks4 
         BackColor       =   &H00000000&
         Caption         =   "SOCKS&4"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   3480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtProxyPort 
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
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   186
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtProxyIP 
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   184
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CheckBox chkUseProxies 
         BackColor       =   &H00000000&
         Caption         =   "Use a pro&xy"
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
         Left            =   3000
         TabIndex        =   180
         ToolTipText     =   "Routes your Battle.net connection through a SOCKS4 or SOCKS5 proxy."
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkUDP 
         BackColor       =   &H00000000&
         Caption         =   "Use &Lag Plug"
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
         Left            =   240
         TabIndex        =   175
         ToolTipText     =   "Sets whether or not you have a lag plug when you sign on. If you don't know what this is, leave it off."
         Top             =   4320
         Width           =   1215
      End
      Begin VB.ComboBox cboSpoof 
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
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   174
         Top             =   3840
         Width           =   2295
      End
      Begin VB.ComboBox cboConnMethod 
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
         Height          =   315
         ItemData        =   "frmSettings.frx":0004
         Left            =   1920
         List            =   "frmSettings.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   840
         Width           =   4335
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   26
         X1              =   2640
         X2              =   240
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "(for account registration)"
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
         Left            =   4320
         TabIndex        =   178
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblConn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Email &Address"
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
         Left            =   3000
         TabIndex        =   177
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   27
         X1              =   2880
         X2              =   6360
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "&BNLS server, if applicable:"
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
         Left            =   360
         TabIndex        =   165
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "ms"
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
         Left            =   2280
         TabIndex        =   172
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "&Reconnect delay:"
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
         Left            =   240
         TabIndex        =   170
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblConn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Port:"
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
         Left            =   3000
         TabIndex        =   185
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label lblConn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "&IP/Host:"
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
         Left            =   3000
         TabIndex        =   183
         Top             =   3840
         Width           =   615
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   25
         X1              =   2760
         X2              =   2760
         Y1              =   2520
         Y2              =   4680
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "&Ping spoofing"
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
         Left            =   240
         TabIndex        =   173
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblHashPath 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
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
         Left            =   360
         TabIndex        =   168
         Top             =   1920
         Width           =   6015
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "Connection &method:"
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
         Left            =   360
         TabIndex        =   163
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblConn 
         BackColor       =   &H00000000&
         Caption         =   "Local hashing is supported for all game clients. Your current hash file path is:"
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
         Left            =   360
         TabIndex        =   167
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   24
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Advanced connection settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   162
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   5
      Left            =   3000
      TabIndex        =   189
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkWhisperGreet 
         BackColor       =   &H00000000&
         Caption         =   "&Whisper the greet message"
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
         Left            =   360
         TabIndex        =   192
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtGreetMsg 
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
         Left            =   360
         MaxLength       =   200
         TabIndex        =   194
         Text            =   "Welcome to %c, %0! I am a %v."
         Top             =   1800
         Width           =   5895
      End
      Begin VB.CheckBox chkGreetMsg 
         BackColor       =   &H00000000&
         Caption         =   "&Greet users who join the channel"
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
         Left            =   360
         TabIndex        =   191
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblGreetVars 
         BackColor       =   &H00000000&
         Caption         =   "greet variable container label"
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
         Height          =   2172
         Left            =   360
         TabIndex        =   195
         Top             =   2280
         Width           =   5892
      End
      Begin VB.Label lblGreet 
         BackColor       =   &H00000000&
         Caption         =   "Greet &message"
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
         Left            =   360
         TabIndex        =   193
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   28
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         Caption         =   "Greet message settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   190
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   9
      Left            =   3000
      TabIndex        =   196
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkLogDBActions 
         BackColor       =   &H00000000&
         Caption         =   "Log &database changes"
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
         Left            =   480
         TabIndex        =   201
         ToolTipText     =   "Any actions that change the database will be logged in the log folder in a file called 'database.txt'."
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox chkLogAllCommands 
         BackColor       =   &H00000000&
         Caption         =   "&Log all commands"
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
         Left            =   480
         TabIndex        =   202
         ToolTipText     =   "Any commands issued to the bot will be logged in a file in the bot's Logs folder called 'commandlog.txt'."
         Top             =   2610
         Width           =   2295
      End
      Begin VB.TextBox txtMaxLogSize 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   208
         Text            =   "0"
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cboLogging 
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
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   200
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox txtMaxBackLogSize 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   205
         Text            =   "10000"
         ToolTipText     =   $"frmSettings.frx":0008
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "Size limits (set 0 for unlimited)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   203
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   30
         X1              =   360
         X2              =   3720
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   $"frmSettings.frx":00A6
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
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   198
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "&Text logging"
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
         Left            =   360
         TabIndex        =   199
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "&Maximum logfile size:"
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
         Left            =   360
         TabIndex        =   207
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "  megabytes"
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
         Left            =   2880
         TabIndex        =   209
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "  bytes"
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
         Left            =   2880
         TabIndex        =   206
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblLog 
         BackColor       =   &H00000000&
         Caption         =   "Maximum &backlog size:"
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
         Left            =   360
         TabIndex        =   204
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Logging settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   197
         Top             =   240
         Width           =   3255
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   29
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// frmSettings.frm - project StealthBot - january 2004 - authored by Stealth (stealth@stealthbot.net)
'// To switch between panels on the UI display, click the current top panel
'//     and use CTRL+K to send it to the back. Rinse and repeat until you are looking
'//     at the panel you want to edit.
' u 7/13/09 to fix topic 42703 -andy
Option Explicit

Private mColors()           As Long
Private FirstRun            As Byte
Private ModifiedColors      As Boolean
Private InitChanFont        As String
Private InitChatFont        As String
Private InitChanSize        As Integer
Private InitChatSize        As Integer
Private PanelsInitialized   As Boolean
Private OldBotOwner         As String

Const SC    As Byte = 0
Const BW    As Byte = 1
Const D2    As Byte = 2
Const D2X   As Byte = 3
Const W3    As Byte = 4
Const W3X   As Byte = 5
Const W2    As Byte = 6

Private Const BNLS_NOT_SET As String = "No server set"

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    Dim nRoot As cTreeViewNode
    Dim nCurrent As cTreeViewNode
    Dim nTopLevel As cTreeViewNodes
    Dim nOptLevel As cTreeViewNodes
    Dim lMouseOver As Long
    Dim s As String
    Dim serverList() As String
    Dim i As Long, j As Long, K As Long
    Dim colBNLS As Collection
    
    '##########################################
    ' TREEVIEW INITIALIZATION CODE
    '##########################################
    
    lMouseOver = vbWhite
    
    With tvw
        
        .Nodes.Clear
        
        Set nRoot = .Nodes.Add(, etvwFirst, "root", "StealthBot Settings")
            nRoot.MouseOverForeColor = lMouseOver
        
            Set nTopLevel = nRoot.Children
            
                Set nCurrent = nTopLevel.Add(, etvwChild, "connection", "Connection Settings")
                    nCurrent.MouseOverForeColor = lMouseOver
                    
                    Set nOptLevel = nCurrent.Children
                        nOptLevel.Add , etvwChild, "conn_config", "Basic Setup"  'general setup
                        nOptLevel.Add , etvwChild, "conn_advanced", "Advanced"     'proxies/spoofing/bnls
                        
                    nCurrent.Expanded = True
                
                Set nCurrent = nTopLevel.Add(, etvwChild, "interface", "Interface Settings")
                    nCurrent.MouseOverForeColor = lMouseOver
                    
                    Set nOptLevel = nCurrent.Children
                        nOptLevel.Add , etvwChild, "int_general", "General Settings"
                        nOptLevel.Add , etvwChild, "int_fonts", "Fonts and Colors"
                        
                    nCurrent.Expanded = True
                        
                Set nCurrent = nTopLevel.Add(, etvwChild, "general", "General Settings")
                    nCurrent.MouseOverForeColor = lMouseOver
                    
                    Set nOptLevel = nCurrent.Children
                        nOptLevel.Add , etvwChild, "op_moderation", "Moderation"
                        nOptLevel.Add , etvwChild, "op_logging", "Logging"
                        nOptLevel.Add , etvwChild, "op_greets", "Greet Message"
                        nOptLevel.Add , etvwChild, "op_idles", "Idle Message"
                        nOptLevel.Add , etvwChild, "op_misc", "Miscellaneous"
                        
                    nCurrent.Expanded = True
                    
                    Set nOptLevel = Nothing
                Set nCurrent = Nothing
                
            Set nTopLevel = Nothing
                    
            nRoot.Expanded = True
            
        Set nRoot = Nothing
        
    End With
    
    '##########################################
    ' PROFILE SELECTOR COMBO BOX STUFF
    '##########################################
    
    With cboProfile
        .Text = "[default profile]"
    End With
    
    Set colProfiles = New Collection
    
    'Call LoadProfileList(cboProfile)
    
    
    '##########################################
    ' INTERFACE DISPLAY
    '##########################################
    
    lblSplash.Caption = vbCrLf & "If you're new to bots, click " & Chr(39) & "Step-By-Step Configuration" & Chr(39) & " below for a walkthrough to get the bot set up." & vbCrLf & vbCrLf & "Otherwise, click a section on the left to change settings."
    
    With cboSpoof
        .AddItem "Disabled"
        .AddItem "0ms (send first)"
        .AddItem "-1ms (ignore)"
        .ListIndex = 0
    End With
    
    With cboConnMethod
        .AddItem "BNLS - Battle.net Logon Server"
        .AddItem "ADVANCED - Local hashing"
        .ListIndex = 0
    End With
    
    With cboBNLSServer
        Dim BNLSServers As Collection
        
        .AddItem "Automatic (Server Finder)"
        
        ' If the user has a server set, add it.
        If Len(Config.BNLSServer) > 0 Then
            AddBNLSServer Config.BNLSServer
        End If
        
        ' Figure out the default selection
        If Config.BNLSFinder Or Len(Config.BNLSServer) = 0 Then
            .ListIndex = 0
        ElseIf Len(Config.BNLSServer) > 0 Then
            .ListIndex = GetBnlsIndex(Config.BNLSServer)
        Else
            .ListIndex = 0
        End If
        
        ' Add servers from the user's local list.
        Set colBNLS = ListFileLoad(GetFilePath(FILE_BNLS_LIST))
        
        If colBNLS.Count > 0 Then
            For i = 1 To colBNLS.Count
                AddBNLSServer colBNLS.Item(i)
            Next i
        End If
        
        Set BNLSServers = Nothing
    End With
    
    With cboLogging
        .AddItem "Full - text is logged and a dated logfile is saved during operation."
        .AddItem "Temporary - text is logged. The logfile is deleted on shutdown."
        .AddItem "Disabled"
        .ListIndex = 0
    End With
    
    With cboTimestamp
        .AddItem "[HH:MM:SS PM] - Seconds with time of day"
        .AddItem "[HH:MM:SS] - Seconds"
        .AddItem "[HH:MM:SS:mmm] - Milliseconds"
        .AddItem "No timestamp"
        .ListIndex = 0
    End With
    
    lblGreetVars.Caption = "Greet message variables:" & vbNewLine & _
        "%c or %chan = Current channel" & vbNewLine & _
        "%0 or %user = Username of the person who joins" & vbNewLine & _
        "%1 or %me = Bot's current username" & vbNewLine & _
        "%p or %ping = Ping of the person who joins" & vbNewLine & _
        "%v or %ver = The bot's current version" & vbNewLine & _
        "%a or %r = Database access / %f = Database flags of the person who joins" & vbNewLine & _
        "%m or %mail = Number of unread messages for the person who joins" & vbNewLine & _
        "%bc = Mumber of people the bot has banned" & vbNewLine & _
        "%t = Current time / %d = Current date"

    lblIdleVars.Caption = "Idle message variables:" & vbNewLine & _
        "%c or %chan = Current channel" & vbNewLine & _
        "%1 or %me = Current username" & vbNewLine & _
        "%v or %ver = Bot version" & vbNewLine & _
        "%botup = Bot uptime / %cpuup = System uptime" & vbNewLine & _
        "%mp3 = Current MP3" & vbNewLine & _
        "%quote = Random quote" & vbNewLine & _
        "%rnd = Random person in the channel" & vbNewLine & _
        "%bc = Mumber of people the bot has banned" & vbNewLine & _
        "%t = Current time / %d = Current date"
    
    '##########################################
    'COLOR STUFF
    '##########################################
    Call LoadColors
    
    cDLG.Filter = "StealthBot ColorList Files|*.sclf"
    cboColorList.ListIndex = 0
    '##########################################
    '##########################################
    ShowCurrentColor
    
    Call InitAllPanels
    
    PanelsInitialized = True
    
    '##########################################
    '   LAST SETTINGS PANEL
    '##########################################
    
    If Config.LastSettingsPanel = -1 Then
        ShowPanel spSplash, False
    Else
        ShowPanel CInt(Config.LastSettingsPanel)
    End If
End Sub

Private Function GetBnlsIndex(ByVal sServerName As String) As Integer
    Dim i As Integer
    GetBnlsIndex = -1
    
    For i = 1 To (cboBNLSServer.ListCount - 1)
        If StrComp(sServerName, cboBNLSServer.List(i), vbTextCompare) = 0 Then
            GetBnlsIndex = i
            Exit Function
        End If
    Next
End Function

' Adds the server to the list (if it does not already exist) and returns its position.
Private Function AddBNLSServer(ByVal sServerHost As String) As Integer
    Dim i As Integer    ' counter
    sServerHost = Trim$(sServerHost)
    
    If Len(sServerHost) = 0 Then
        AddBNLSServer = -1
        Exit Function
    End If
    
    ' Check if the server is already in the list.
    i = GetBnlsIndex(sServerHost)
    If i = -1 Then
        cboBNLSServer.AddItem sServerHost
        i = GetBnlsIndex(sServerHost)
    End If
    
    AddBNLSServer = i
End Function

Private Sub SetHashPath(ByVal Path As String)
    lblHashPath.Caption = Path
    lblHashPath.ToolTipText = Path
    
    ' Try to adjust font size for long paths
    If ((Len(Path) > 80) And (Not InStr(Path, Space(1)) > 0)) Then
        If Len(Path) > 100 Then
            lblHashPath.FontSize = 6
        Else
            lblHashPath.FontSize = 7
        End If
    End If
    
End Sub

Function KeyToIndex(ByVal sKey As String) As Byte
    
    Select Case sKey
        Case "splash":          KeyToIndex = 8
    
        Case "conn_config":     KeyToIndex = 0
        Case "conn_advanced":   KeyToIndex = 1
        
        Case "int_general":     KeyToIndex = 2
        Case "int_fonts":       KeyToIndex = 3
    
        Case "op_moderation":   KeyToIndex = 4
        Case "op_logging":      KeyToIndex = 9
        Case "op_greets":       KeyToIndex = 5
        Case "op_idles":        KeyToIndex = 6
        Case "op_misc":         KeyToIndex = 7
        
        Case Else:              KeyToIndex = 8
    End Select
    
End Function

Sub ShowPanel(ByVal Index As enuSettingsPanels, Optional ByVal SavePanel As Boolean = True)

    Dim Control As Control
    Dim nod     As cTreeViewNode
    Dim i       As Integer
    Dim j       As Integer
    Dim TabStop As Boolean

    If PanelsInitialized Then
        ' set the selected node
        If Index <> spSplash Then
            For i = 1 To tvw.NodeCount
                Set nod = tvw.Nodes.Item(i)
                If Not nod.Selected And KeyToIndex(nod.Key) = Index Then
                    nod.Selected = True
                    Exit For
                End If
            Next i
            Set nod = Nothing
        End If

        ' set panel visibility
        For i = fraPanel.LBound To fraPanel.UBound
            If Index = i Then
                fraPanel(i).Visible = True
            Else
                fraPanel(i).Visible = False
            End If
        Next i

        ' save the latest panel choice to config
        If SavePanel Then
            Config.LastSettingsPanel = Index
        End If
    End If

End Sub

Private Sub cboConnMethod_Click()
    'Disables the server selection box when hashing is selected
    cboBNLSServer.Enabled = CBool(cboConnMethod.ListIndex = 0)
    chkProxyBNLS.Enabled = CBool(cboConnMethod.ListIndex = 0) And chkUseProxies.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReadme_Click()
    OpenReadme
End Sub

Private Sub cmdSave_Click()
    If ModifiedColors Then Call cmdSaveColor_Click
    
    If Not InvalidConfigValues Then
        If SaveSettings Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdSaveColor_Click()
    ModifiedColors = True
    RecordCurrentColor
    lblColorStatus.Caption = "Color settings saved."
End Sub

Private Sub cmdStepByStep_Click()
    Call frmCustomInputBox.Show(vbModal, frmSettings)
End Sub

Private Sub cmdWebsite_Click()
    Call frmChat.mnuHelpWebsite_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmChat.DeconstructSettings
    Set colProfiles = Nothing
End Sub

Private Sub lblAccountManage_Click()
    ' save the values that Account Manager reads
    Config.Username = txtUsername.Text
    Config.Password = txtPassword.Text
    Config.RegisterEmailDefault = txtEmail.Text
    Config.Save

    ' load Account Manager modal
    Call frmAccountManager.ShowMode(Config.AccountMode, vbModal, frmSettings)
End Sub

Sub lblAddCurrentKey_Click()
    Dim keys As Collection
    Dim Item As Variant
    Dim Key1 As String
    Dim Key2 As String
    
    ' Load the list
    Set keys = ListFileLoad(GetFilePath(FILE_KEY_LIST))
    
    Key1 = UCase$(CDKeyReplacements(txtCDKey.Text))
    Key2 = UCase$(CDKeyReplacements(txtExpKey.Text))
    
    ' if it's already there, do nothing.
    For Each Item In keys
        If StrComp(CStr(Item), Key1, vbTextCompare) = 0 Then Key1 = vbNullString
        If StrComp(CStr(Item), Key2, vbTextCompare) = 0 Then Key2 = vbNullString
    Next Item
    
    ' Add the keys
    If LenB(Key1) > 0 Then keys.Add Key1
    If LenB(Key2) > 0 Then keys.Add Key2
    
    ' Save the list
    If LenB(Key1) > 0 Or LenB(Key2) > 0 Then
        ListFileSave GetFilePath(FILE_KEY_LIST), keys
    End If
    
    Set keys = Nothing
End Sub

Private Sub lblManageKeys_Click()
    If LenB(txtCDKey.Text) > 0 Then
        Call lblAddCurrentKey_Click
    End If

    Call frmKeyManager.Show(vbModal, frmSettings)
End Sub

Sub lblAccountManage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAccountManage.ForeColor = &HFFCC99
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = vbWhite

    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

Sub lblAddCurrentKey_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAccountManage.ForeColor = vbWhite
    lblAddCurrentKey.ForeColor = &HFFCC99
    lblManageKeys.ForeColor = vbWhite

    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

Sub lblManageKeys_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAccountManage.ForeColor = vbWhite
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = &HFFCC99

    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAccountManage.ForeColor = vbWhite
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = vbWhite
End Sub

Sub fraPanel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAccountManage.ForeColor = vbWhite
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = vbWhite
End Sub

'Private Sub tvw_Click()
'    Call tvw_SelectedNodeChanged
'End Sub
'
'Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call tvw_SelectedNodeChanged
'End Sub
'
'Private Sub tvw_NodeClick(Node As vbalTreeViewLib6.cTreeViewNode)
'    Call tvw_SelectedNodeChanged
'End Sub

Private Sub tvw_SelectedNodeChanged()
    If FirstRun = 0 Then
        ShowPanel KeyToIndex(tvw.SelectedItem.Key)
    Else
        FirstRun = 0
    End If
End Sub

Private Function DoCDKeyLengthCheck(ByVal sKey As String, ByVal sProd As String) As Boolean
    sKey = CDKeyReplacements(sKey)
    
    DoCDKeyLengthCheck = True
    If Config.IgnoreCDKeyLength Then Exit Function
    
    Select Case sProd
        Case PRODUCT_STAR, PRODUCT_SEXP
            If ((Len(sKey) <> 13) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case PRODUCT_D2DV, PRODUCT_D2XP
            If ((Len(sKey) <> 16) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case PRODUCT_W2BN
            If (Len(sKey) <> 16) Then DoCDKeyLengthCheck = False
            
        Case PRODUCT_WAR3, PRODUCT_W3XP
            If (Len(sKey) <> 26) Then DoCDKeyLengthCheck = False
        
        Case PRODUCT_SSHR, PRODUCT_DSHR, PRODUCT_DRTL
            DoCDKeyLengthCheck = True
        
        Case PRODUCT_JSTR
            If (Len(sKey) <> 13) Then DoCDKeyLengthCheck = False
        
        Case Else
            DoCDKeyLengthCheck = False
    End Select
End Function

Private Function SaveSettings() As Boolean
    Dim s As String
    Dim Clients(6) As String
    Dim i As Long, j As Long
    Dim colBNLS As New Collection
    
    ' First, CDKey Length check and corresponding stuff that needs to run first:
    Select Case True
        Case optSTAR.Value:
            If (chkSHR.Value) Then
                s = PRODUCT_SSHR
            ElseIf (chkJPN.Value) Then
                s = PRODUCT_JSTR
            Else
                s = PRODUCT_STAR
            End If
        Case optSEXP.Value: s = PRODUCT_SEXP
        Case optD2DV.Value: s = PRODUCT_D2DV
        Case optD2XP.Value: s = PRODUCT_D2XP
        Case optWAR3.Value: s = PRODUCT_WAR3
        Case optW3XP.Value: s = PRODUCT_W3XP
        Case optW2BN.Value: s = PRODUCT_W2BN
        Case optDRTL.Value:
            If (chkSHR.Value) Then
                s = PRODUCT_DSHR
            Else
                s = PRODUCT_DRTL
            End If
        'Case optCHAT.Value: s = PRODUCT_CHAT
    End Select
    
    If Not DoCDKeyLengthCheck(txtCDKey.Text, s) Then
        If MsgBox("Your CD key is of an invalid Length for the product you have chosen. Do you want to save anyway?", vbExclamation + vbYesNo, "StealthBot Settings") = vbNo Then
            ShowPanel spConnectionConfig
            txtCDKey.SetFocus
            SaveSettings = False
            Exit Function
        End If
    End If
    
    If txtExpKey.Enabled And Not DoCDKeyLengthCheck(txtExpKey.Text, s) Then
        If MsgBox("Your expansion CD key is of an invalid Length for the product you have chosen. Do you want to anyway?", vbInformation + vbYesNo, "StealthBot Settings") = vbNo Then
            ShowPanel spConnectionConfig
            txtExpKey.SetFocus
            SaveSettings = False
            Exit Function
        End If
    End If
    
    Config.Game = s
    
    ' The rest of the basic config now
    Config.Username = txtUsername.Text
    Config.Password = txtPassword.Text
    Config.CDKey = CDKeyReplacements(txtCDKey.Text)
    Config.ExpKey = CDKeyReplacements(txtExpKey.Text)
    Config.HomeChannel = txtHomeChan.Text
    Config.Server = cboServer.Text

    Config.UseSpawn = CBool(CanSpawn(Config.Game, Len(Config.CDKey)) And CBool(chkSpawn.Value))
    Config.UseD2Realms = CBool(chkUseRealm.Value)

    ' reset these when save & close normal config
    Config.NewPassword = vbNullString
    Config.RegisterEmailChange = vbNullString
    
    ' Advanced connection settings
    Config.UseBNLS = CBool(cboConnMethod.ListIndex = 0)
    Config.BNLSFinder = CBool(cboBNLSServer.ListIndex = 0)
    
    If cboBNLSServer.ListIndex > 0 Then
        Config.BNLSServer = cboBNLSServer.Text
    ElseIf cboBNLSServer.ListIndex = -1 Then
        If StrComp(cboBNLSServer.Text, BNLS_NOT_SET, vbBinaryCompare) <> 0 Then
            Config.BNLSServer = cboBNLSServer.Text
            AddBNLSServer Config.BNLSServer
        End If
    End If
    
    ' Save the BNLS server list
    With cboBNLSServer
        j = -1
        
        ' Check if the set server is in the list
        For j = 1 To .ListCount
            If StrComp(.Text, .List(j)) = 0 Then
                j = -1
                Exit For
            End If
        Next j
        
        If j >= 0 Or .ListCount > 0 Then
            For j = 1 To .ListCount
                If Not (Len(.List(j)) = 0) Then colBNLS.Add .List(j)
            Next
            
            ListFileSave GetFilePath(FILE_BNLS_LIST), colBNLS
        End If
    End With
    
    Config.AutoConnect = CBool(chkConnectOnStartup.Value)
    Config.RegisterEmailDefault = Trim$(txtEmail.Text)
    Config.PingSpoofing = cboSpoof.ListIndex
    Config.UseUDP = CBool(chkUDP.Value)
    Config.UseProxy = CBool(chkUseProxies.Value)
    Config.ProxyPort = StringToNumber(txtProxyPort.Text, Config.ProxyPort)
    Config.ProxyIP = Trim$(txtProxyIP.Text)
    Config.ProxyType = IIf(CBool(optSocks5.Value), PROXY_SETTING_SOCKS5, PROXY_SETTING_SOCKS4)
    Config.ProxyBNLS = CBool(chkProxyBNLS.Value)
    Config.ProxyMCP = CBool(chkProxyMCP.Value)

    Config.ReconnectDelay = StringToNumber(txtReconDelay.Text, Config.ReconnectDelay)
    
    ' General Interface settings
    Config.ShowSplashScreen = CBool(chkSplash.Value)
    Config.DisableNews = Not CBool(chkShowNews.Value)
    Config.MinimizeOnStartup = CBool(chkMinimizeOnStartup.Value)
    Config.MinimizeToTray = CBool(chkMinimizeToTray.Value)
    
    Config.UseUTF8 = CBool(chkUTF8.Value)
    Config.ShowJoinLeaves = CBool(chkJoinLeaves.Value)
    Config.FlashOnEvents = CBool(chkFlash.Value)
    Config.ChatFilters = CBool(chkFilter.Value)
    Config.UrlDetection = CBool(chkURLDetect.Value)
    Config.NameAutoComplete = CBool(chkNameAutocomplete.Value)
    
    Config.NameColoring = CBool(chkListColoring.Value)
    Config.ShowStatsIcons = CBool(chkShowUserGameStatsIcons.Value)
    Config.ShowFlagIcons = CBool(chkShowUserFlagsIcons.Value)
    
    Config.DisablePrefixBox = Not CBool(chkShowPrefixBox.Value)
    Config.DisableSuffixBox = Not CBool(chkShowSuffixBox.Value)
    Config.TimestampMode = cboTimestamp.ListIndex
    
    ' Font and color
    SaveFontSettings
    
    ' Moderation settings
    Config.IPBans = CBool(chkIPBans.Value)
    Config.UDPBan = CBool(chkPlugban.Value)
    Config.KickOnYell = CBool(chkKOY.Value)
    Config.BanEvasion = CBool(chkBanEvasion.Value)
    
    Config.QuietTime = CBool(chkQuiet.Value)
    Config.QuietTimeKick = CBool(chkQuietKick.Value)
    
    Config.PhraseBans = CBool(chkPhrasebans.Value)
    Config.PhraseKick = CBool(chkPhraseKick.Value)
    
    Config.PingBan = CBool(chkPingBan.Value)
    Config.PingBanLevel = StringToNumber(txtPingLevel.Text, Config.PingBanLevel)
    
    Config.ChannelProtection = CBool(chkProtect.Value)
    Config.ChannelProtectionMessage = txtProtectMsg.Text
    
    Config.IdleBan = CBool(chkIdlebans.Value)
    Config.IdleBanKick = CBool(chkIdleKick.Value)
    Config.IdleBanDelay = StringToNumber(txtIdleBanDelay.Text, Config.IdleBanDelay)
    
    Call SaveClientBans
    
    Config.PeonBan = CBool(chkPeonbans.Value)
    
    Config.LevelBanW3 = StringToNumber(txtBanW3.Text, Config.LevelBanW3)
    Config.LevelBanD2 = StringToNumber(txtBanD2.Text, Config.LevelBanD2)
    Config.LevelBanMessage = txtLevelBanMsg.Text
    
    '// Logging options
    If (cboLogging.ListIndex = 0) Then
        Config.LoggingMode = 2
    ElseIf (cboLogging.ListIndex = 1) Then
        Config.LoggingMode = 1
    ElseIf (cboLogging.ListIndex = 2) Then
        Config.LoggingMode = 0
    End If
    
    Config.LogDBActions = CBool(chkLogDBActions.Value)
    Config.LogCommands = CBool(chkLogAllCommands.Value)
    
    Config.MaxBacklogSize = StringToNumber(txtMaxBackLogSize.Text, Config.MaxBacklogSize)
    Config.MaxLogFileSize = StringToNumber(txtMaxLogSize.Text, Config.MaxLogFileSize)
    
    ' Greet Message Settings
    Config.GreetMessageText = txtGreetMsg.Text
    Config.WhisperGreet = CBool(chkWhisperGreet.Value)
    Config.GreetMessage = CBool(chkGreetMsg.Value)
    
    ' Idle message settings
    Config.IdleMessage = CBool(chkIdles.Value)
    If IsNumeric(txtIdleWait.Text) Then
        Config.IdleMessageDelay = CLng(txtIdleWait.Text * 2)
    End If
    
    Select Case True
        Case optMsg.Value:       Config.IdleMessageType = "msg"
        Case optUptime.Value:    Config.IdleMessageType = "uptime"
        Case optMP3.Value:       Config.IdleMessageType = "mp3"
        Case optQuote.Value:     Config.IdleMessageType = "quote"
        Case Else: Config.IdleMessageType = "msg"
    End Select
    Config.IdleMessageText = txtIdleMsg.Text
    
    
    '// Misc General Settings
    Config.Mp3Commands = CBool(chkAllowMP3.Value)
    Config.ProfileAmp = CBool(chkPAmp.Value)
    Config.BotMail = CBool(chkMail.Value)
    Config.IgnoreClanInvites = Not CBool(chkClanInvitations.Value)
    
    Config.BotOwner = txtOwner.Text
    Config.Trigger = txtTrigger.Text

    Config.WhisperCommands = CBool(chkWhisperCmds.Value)
    Config.ShowOfflineFriends = CBool(chkShowOffline.Value)
    Config.UseBackupChannel = CBool(chkBackup.Value)
    Config.BackupChannel = txtBackupChan.Text
    
    For i = optNaming.LBound To optNaming.UBound
        If optNaming(i).Value Then Exit For ' i = index of opt checked
    Next i
    If i = 4 Then i = 0 ' if none were checked, then set to default
    Config.NamespaceConvention = i
    
    Config.UseD2Naming = CBool(chkD2Naming.Value)
    
    '// Save the config instance to disk
    Call Config.Save
    
    '// Load the config into the form
    Call frmChat.ReloadConfig(1)
    
    '// RESIZE FORM TO FIX ANY UI CHANGES
    Call frmChat.Form_Resize
    
    '// Take care of the colors.
    If ModifiedColors Then
        Call SaveColors
        Call GetColorLists
    End If
    
    SaveSettings = True
End Function

' check for potential invalid config stuffs
Public Function InvalidConfigValues() As Boolean

    Dim s As String
    
    If optW3XP.Value Or optD2XP.Value Then
        If LenB(txtExpKey.Text) = 0 Then
            If optW3XP.Value Then
                s = "Warcraft III and a Frozen Throne"
            Else
                s = "Diablo II and a Lord of Destruction"
            End If
            
            MsgBox "You must enter both a " & s & " CD-key to connect with an Expansion game.", vbOKOnly + vbInformation
            ShowPanel spConnectionConfig
            txtExpKey.SetFocus
            InvalidConfigValues = True
        End If
    End If
    
End Function


'##########################################
' COLOR-RELATED CODE
'##########################################

Sub CAdd(ByVal colorName As String, ColorValue As Long, Optional Append As Byte)
    mColors(UBound(mColors)) = ColorValue
    
    If Append = 1 Then colorName = "Event | " & colorName
    
    cboColorList.AddItem colorName
    
    ReDim Preserve mColors(UBound(mColors) + 1)
End Sub

Private Sub cmdExport_Click()
    With cDLG
        .FileName = vbNullString
        .ShowSave
        If .FileName <> vbNullString Then
            SaveColors .FileName
            MsgBox "ColorList exported.", vbOKOnly
        End If
    End With
End Sub

Private Sub cmdImport_Click()
    With cDLG
        .FileName = vbNullString
        .ShowOpen
        If .FileName <> vbNullString Then
            GetColorLists (.FileName)
            cboColorList.Clear
            Call Form_Load
        End If
    End With
End Sub

Private Sub cmdGetRGB_Click()
    On Error Resume Next
    txtValue.Text = RGB(txtR.Text, txtG.Text, txtB.Text)
End Sub

Private Sub cmdHTMLGen_Click()
    If Left$(txtHTML.Text, 1) = "#" Then txtHTML.Text = Mid$(txtHTML.Text, 2)
    
    txtValue.Text = HTMLToRGBColor(txtHTML.Text)
End Sub

Private Sub cmdDefaults_Click()
    If MsgBox("Are you sure you want to restore the default Values()?" & vbCrLf & _
            "(All current color data will be lost unless exported)", vbYesNo + vbExclamation) = vbYes Then
            
        If LenB(Dir$(GetFilePath(FILE_COLORS))) > 0 Then
            Kill GetFilePath(FILE_COLORS)
            Call GetColorLists
            Call LoadColors
        End If
    End If
End Sub

Private Sub SaveColors(Optional sPath As String)
    Dim f As Integer
    Dim i As Integer
    
    f = FreeFile
    
    If LenB(sPath) = 0 Then sPath = GetFilePath(FILE_COLORS)
    
    Open sPath For Random As #f Len = 4
    
    For i = LBound(mColors) To UBound(mColors)
        Put #f, i + 1, CLng(mColors(i))
        'Debug.Print "Putting color; " & i & ":" & mColors(i)
    Next i
    
    Close #f
End Sub

Private Sub txtValue_Change()
    On Error Resume Next
    lblEg.BackColor = Val(txtValue.Text)
End Sub

Private Sub cmdColorPicker_Click()
    cDLG.ShowColor
    txtValue.Text = cDLG.Color
End Sub


Sub ShowCurrentColor()
    On Error GoTo ShowCurrentColor_Error

    lblEg.BackColor = mColors(cboColorList.ListIndex)
    txtValue.Text = mColors(cboColorList.ListIndex)
    lblCurrentValue.Caption = mColors(cboColorList.ListIndex)

ShowCurrentColor_Exit:
    Exit Sub

ShowCurrentColor_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowCurrentColor of Form frmSettings"
    Resume ShowCurrentColor_Exit
End Sub

Private Sub RecordCurrentColor()
    If cboColorList.ListIndex > -1 Then
        mColors(cboColorList.ListIndex) = Val(txtValue.Text)
        'Debug.Print "Recording current color."
    End If
End Sub

Private Sub cboColorList_GotFocus()
    ShowCurrentColor
End Sub

Private Sub cboColorList_Scroll()
    ShowCurrentColor
End Sub

Private Sub cboColorList_Click()
    ModifiedColors = True
    lblColorStatus.Caption = "Be sure to click 'Save Changes to This Color' before proceeding."
    ShowCurrentColor
End Sub



'##########################################
' ENABLE/DISABLE CODE
'##########################################

Private Sub chkUseProxies_Click()
    txtProxyIP.Enabled = chkUseProxies.Value
    txtProxyPort.Enabled = chkUseProxies.Value
    optSocks4.Enabled = chkUseProxies.Value
    optSocks5.Enabled = chkUseProxies.Value
    chkProxyBNLS.Enabled = CBool(cboConnMethod.ListIndex = 0) And chkUseProxies.Value
    chkProxyMCP.Enabled = chkUseRealm.Value And chkUseProxies.Value
End Sub

Private Sub chkUseRealm_Click()
    chkProxyMCP.Enabled = chkUseRealm.Value And chkUseProxies.Value
End Sub

Private Sub chkBackup_Click()
    txtBackupChan.Enabled = chkBackup.Value
End Sub

Private Sub chkIdlebans_click()
    chkIdleKick.Enabled = chkIdlebans.Value
    txtIdleBanDelay.Enabled = chkIdlebans.Value
End Sub

Private Sub chkIdles_Click()
    optMsg.Enabled = chkIdles.Value
    optUptime.Enabled = chkIdles.Value
    optMP3.Enabled = chkIdles.Value
    optQuote.Enabled = chkIdles.Value
    txtIdleWait.Enabled = chkIdles.Value
    txtIdleMsg.Enabled = (optMsg.Enabled And optMsg.Value)
End Sub

Private Sub optMsg_Click()
    txtIdleMsg.Enabled = True
End Sub

Private Sub optUptime_Click()
    txtIdleMsg.Enabled = False
End Sub

Private Sub optMP3_Click()
    txtIdleMsg.Enabled = False
End Sub

Private Sub optQuote_Click()
    txtIdleMsg.Enabled = False
End Sub

Sub optSTAR_Click()
    chkSHR.Visible = True
    chkSpawn.Enabled = True
    chkJPN.Visible = True
    txtCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    If (chkSHR.Value) Then
        SetHashPath GetGamePath("RHSS")
        chkSpawn.Enabled = False
        txtCDKey.Enabled = False
    ElseIf (chkJPN.Value) Then
        SetHashPath GetGamePath("RTSJ")
    Else
        SetHashPath GetGamePath("RATS")
    End If
    chkUDP.Enabled = True
End Sub

Sub optWAR3_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = False
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("3RAW")
    chkUDP.Enabled = False
End Sub

Sub optD2DV_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = False
    chkSpawn.Value = vbUnchecked
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = True
    SetHashPath GetGamePath("VD2D")
    chkUDP.Enabled = False
End Sub

Sub optW2BN_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = True
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("NB2W")
    chkUDP.Enabled = True
End Sub

Sub optSEXP_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = False
    chkSpawn.Value = vbUnchecked
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("RATS")
    chkUDP.Enabled = True
End Sub

Sub optD2XP_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = False
    chkSpawn.Value = vbUnchecked
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = True
    SetHashPath GetGamePath("PX2D")
    chkUDP.Enabled = False
End Sub

Sub optW3XP_Click()
    chkSHR.Visible = False
    chkSpawn.Enabled = False
    chkSpawn.Value = vbUnchecked
    chkJPN.Visible = False
    txtCDKey.Enabled = True
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("PX3W")
    chkUDP.Enabled = False
End Sub

Sub optDRTL_Click()
    chkSHR.Visible = True
    chkSpawn.Enabled = False
    chkSpawn.Value = vbUnchecked
    chkJPN.Visible = False
    txtCDKey.Enabled = False
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    If (chkSHR.Value) Then
        SetHashPath GetGamePath("RHSD")
    Else
        SetHashPath GetGamePath("LTRD")
    End If
    chkUDP.Enabled = True
End Sub

Private Sub chkJPN_Click()
    Dim Checked As Boolean
    Checked = CBool(chkJPN.Value)
    If (Checked) Then chkSHR.Value = vbUnchecked
    If (optSTAR.Value) Then
        If (Checked) Then
            SetHashPath GetGamePath("RTSJ")
        Else
            SetHashPath GetGamePath("RATS")
        End If
    End If
End Sub

Private Sub chkSHR_Click()
    Dim Checked As Boolean
    Checked = CBool(chkSHR.Value)
    If (Checked) Then chkJPN.Value = vbUnchecked
    If (optSTAR.Value) Then
        chkSpawn.Enabled = Not Checked
        txtCDKey.Enabled = Not Checked
        If (Checked) Then
            SetHashPath GetGamePath("RHSS")
        Else
            SetHashPath GetGamePath("RATS")
        End If
    ElseIf (optDRTL.Value) Then
        If (Checked) Then
            SetHashPath GetGamePath("RHSD")
        Else
            SetHashPath GetGamePath("LTRD")
        End If
    End If
End Sub

Private Sub chkGreetMsg_Click()
    chkWhisperGreet.Enabled = chkGreetMsg.Value
    txtGreetMsg.Enabled = chkGreetMsg.Value
End Sub

Private Sub chkProtect_Click()
    txtProtectMsg.Enabled = chkProtect.Value
End Sub

Private Sub chkQuiet_Click()
    chkQuietKick.Enabled = chkQuiet.Value
End Sub

Private Sub chkPhrasebans_Click()
    chkPhraseKick.Enabled = chkPhrasebans.Value
End Sub

Private Sub chkPingBan_Click()
    txtPingLevel.Enabled = chkPingBan.Value
End Sub

'##########################################
' INIT SUBS
'##########################################

Private Sub InitAllPanels()
    InitGenMisc
    InitConnAdvanced
    InitFontsColors
    InitGenInterface
    InitGenMod
    InitBasicConfig
    InitGenGreets
    InitGenIdles
    InitLogging
    
    If Config.FileExists Then
        ShowPanel spConnectionConfig
        FirstRun = 1
    End If
End Sub

Private Sub InitBasicConfig()
    Dim i As Integer
    Dim AddCurrent As Boolean
    Dim Item As String
    Dim AdditionalServerList As Collection
    
    txtUsername.Text = Config.Username
    txtPassword.Text = Config.Password
    txtCDKey.Text = Config.CDKey
    txtExpKey.Text = Config.ExpKey
    
    txtHomeChan.Text = Config.HomeChannel

    With cboServer
        ' add the 4 default servers
        .AddItem "useast.battle.net"
        .AddItem "uswest.battle.net"
        .AddItem "europe.battle.net"
        .AddItem "asia.battle.net"
        
        ' get additional servers
        Set AdditionalServerList = ListFileLoad(GetFilePath(FILE_SERVER_LIST))
        
        ' if additional servers, add a blank line, then add them
        If AdditionalServerList.Count > 0 Then
            .AddItem ""
            
            For i = 1 To AdditionalServerList.Count
                .AddItem AdditionalServerList.Item(i)
            Next i
        End If
        
        ' check if "currently selected" is in list
        AddCurrent = True
        
        For i = 0 To .ListCount - 1
            Item = .List(i)
            
            If StrComp(Item, Config.Server, vbBinaryCompare) = 0 Then
                AddCurrent = False
                .ListIndex = i
            End If
        Next i
        
        ' if not, add it (first)
        If AddCurrent Then
            .AddItem Config.Server, 0
            .ListIndex = 0
        End If
    End With
    
    Select Case GetProductInfo(Config.Game).Code
        Case PRODUCT_STAR:    Call optSTAR_Click: optSTAR.Value = True: chkSHR.Value = vbUnchecked: chkJPN.Value = vbUnchecked
        Case PRODUCT_SEXP:    Call optSEXP_Click: optSEXP.Value = True
        Case PRODUCT_D2DV:    Call optD2DV_Click: optD2DV.Value = True
        Case PRODUCT_D2XP:    Call optD2XP_Click: optD2XP.Value = True
        Case PRODUCT_W2BN:    Call optW2BN_Click: optW2BN.Value = True
        Case PRODUCT_WAR3:    Call optWAR3_Click: optWAR3.Value = True
        Case PRODUCT_W3XP:    Call optW3XP_Click: optW3XP.Value = True
        Case PRODUCT_DRTL:    Call optDRTL_Click: optDRTL.Value = True: chkSHR.Value = vbUnchecked
        Case PRODUCT_DSHR:    Call optDRTL_Click: optDRTL.Value = True: chkSHR.Value = vbChecked
        Case PRODUCT_SSHR:    Call optSTAR_Click: optSTAR.Value = True: chkSHR.Value = vbChecked ' unchecks jpn
        Case PRODUCT_JSTR:    Call optSTAR_Click: optSTAR.Value = True: chkJPN.Value = vbChecked ' unchecks shr
        Case Else:      Call optSTAR_Click: optSTAR.Value = True: chkSHR.Value = vbUnchecked: chkJPN.Value = vbUnchecked
    End Select
    
    chkSpawn.Value = Abs(Config.UseSpawn)
    chkUseRealm.Value = Abs(Config.UseD2Realms)
    
End Sub

Private Sub InitConnAdvanced()
    ' Connection method
    cboConnMethod.ListIndex = CInt(Abs(Not Config.UseBNLS))
    cboBNLSServer.Enabled = Abs(Config.UseBNLS)
    
    ' Set selected BNLS server
    If Config.BNLSFinder Then
        cboBNLSServer.ListIndex = 0
    ElseIf Len(Config.BNLSServer) = 0 Then
        cboBNLSServer.Text = BNLS_NOT_SET
    ElseIf Len(Config.BNLSServer) > 0 Then
        cboBNLSServer.ListIndex = GetBnlsIndex(Config.BNLSServer)
    End If
    
    txtEmail.Text = Config.RegisterEmailDefault
    
    cboSpoof.ListIndex = Config.PingSpoofing
    chkUDP.Value = Abs(Config.UseUDP)
    
    chkConnectOnStartup.Value = Abs(Config.AutoConnect)
    txtReconDelay.Text = Config.ReconnectDelay
    If LenB(txtReconDelay.Text) = 0 Then txtReconDelay.Text = 1000
    
    chkUseProxies.Value = Abs(Config.UseProxy)
    Call chkUseProxies_Click
    
    txtProxyPort.Text = Config.ProxyPort
    txtProxyIP.Text = Config.ProxyIP
    
    Select Case UCase$(Config.ProxyType)
        Case PROXY_SETTING_SOCKS5:
            optSocks5.Value = True
            optSocks4.Value = False
        Case PROXY_SETTING_SOCKS4:
            optSocks5.Value = False
            optSocks4.Value = True
    End Select
    
    chkProxyBNLS.Value = Abs(Config.ProxyBNLS)
    chkProxyMCP.Value = Abs(Config.ProxyMCP)
    
    ' Adjust "BNLS server" label 2 pixels down
    lblConn(1).Top = lblConn(1).Top + (2 * Screen.TwipsPerPixelY)
End Sub

Private Sub InitGenInterface()
    chkSplash.Value = Abs(Config.ShowSplashScreen)
    chkShowNews.Value = Abs(Not Config.DisableNews)
    chkMinimizeOnStartup.Value = Abs(Config.MinimizeOnStartup)
    chkMinimizeToTray.Value = Abs(Config.MinimizeToTray)
    
    chkUTF8.Value = Abs(Config.UseUTF8)
    chkJoinLeaves.Value = Abs(Config.ShowJoinLeaves)
    chkFlash.Value = Abs(Config.FlashOnEvents)
    chkFilter.Value = Abs(Config.ChatFilters)
    chkURLDetect.Value = Abs(Config.UrlDetection)
    chkNameAutocomplete.Value = Abs(Config.NameAutoComplete)
    
    chkListColoring.Value = Abs(Config.NameColoring)
    chkShowUserGameStatsIcons.Value = Abs(Config.ShowStatsIcons)
    chkShowUserFlagsIcons.Value = Abs(Config.ShowFlagIcons)
    
    chkShowPrefixBox.Value = Abs(Not Config.DisablePrefixBox)
    chkShowSuffixBox.Value = Abs(Not Config.DisableSuffixBox)
    cboTimestamp.ListIndex = Config.TimestampMode
End Sub

Private Sub InitFontsColors()
    txtChatFont.Text = frmChat.rtbChat.Font.Name
    InitChatFont = txtChatFont.Text
    
    txtChanFont.Text = frmChat.lvChannel.Font.Name
    InitChanFont = txtChanFont.Text
    
    txtChatSize.Text = CInt(frmChat.rtbChat.Font.Size)
    InitChatSize = CInt(frmChat.rtbChat.Font.Size)
    
    txtChanSize.Text = CInt(frmChat.lvChannel.Font.Size)
    InitChanSize = CInt(frmChat.lvChannel.Font.Size)
    
    cboColorList.ListIndex = 0
End Sub

Private Sub InitGenMod()
    chkIPBans.Value = Abs(Config.IPBans)
    chkPlugban.Value = Abs(Config.UDPBan)
    chkKOY.Value = Abs(Config.KickOnYell)
    chkBanEvasion.Value = Abs(Config.BanEvasion)
    
    chkQuiet.Value = Abs(Config.QuietTime)
    chkQuietKick.Value = Abs(Config.QuietTimeKick)
    Call chkQuiet_Click
    
    chkPhrasebans.Value = Abs(Config.PhraseBans)
    chkPhraseKick.Value = Abs(Config.PhraseKick)
    Call chkPhrasebans_Click
    
    chkPingBan.Value = Abs(Config.PingBan)
    txtPingLevel.Text = Config.PingBanLevel
    Call chkPingBan_Click
    
    chkProtect.Value = Abs(Config.ChannelProtection)
    txtProtectMsg.Text = Config.ChannelProtectionMessage
    Call chkProtect_Click
    
    chkIdlebans.Value = Abs(Config.IdleBan)
    chkIdleKick.Value = Abs(Config.IdleBanKick)
    txtIdleBanDelay.Text = Config.IdleBanDelay
    Call chkIdlebans_click
    
    ' grab client ban settings from database
    chkCBan(SC).Value = Abs(IsClientBanned(PRODUCT_STAR))
    chkCBan(BW).Value = Abs(IsClientBanned(PRODUCT_SEXP))
    chkCBan(D2).Value = Abs(IsClientBanned(PRODUCT_D2DV))
    chkCBan(D2X).Value = Abs(IsClientBanned(PRODUCT_D2XP))
    chkCBan(W2).Value = Abs(IsClientBanned(PRODUCT_W2BN))
    chkCBan(W3).Value = Abs(IsClientBanned(PRODUCT_WAR3))
    chkCBan(W3X).Value = Abs(IsClientBanned(PRODUCT_W3XP))

    chkPeonbans.Value = Abs(Config.PeonBan)
    
    txtBanD2.Text = Config.LevelBanD2
    txtBanW3.Text = Config.LevelBanW3
    txtLevelBanMsg.Text = Config.LevelBanMessage
End Sub

Private Sub InitLogging()
    Select Case Config.LoggingMode
        Case 0:
            cboLogging.ListIndex = 2
        Case 1:
            cboLogging.ListIndex = 1
        Case Else:
            cboLogging.ListIndex = 0
    End Select
    
    chkLogDBActions.Value = Abs(Config.LogDBActions)
    chkLogAllCommands.Value = Abs(Config.LogCommands)

    txtMaxBackLogSize.Text = Config.MaxBacklogSize
    txtMaxLogSize.Text = Config.MaxLogFileSize
End Sub

Private Sub InitGenGreets()
    txtGreetMsg.Text = Config.GreetMessageText
    chkGreetMsg.Value = Abs(Config.GreetMessage)
    Call chkGreetMsg_Click
    
    chkWhisperGreet.Value = Abs(Config.WhisperGreet)
End Sub

Private Sub InitGenIdles()
    txtIdleWait.Text = (Config.IdleMessageDelay / 2)
    
    Select Case Config.IdleMessageType
        Case "msg", vbNullString
            optMsg.Value = True
            Call optMsg_Click
        Case "quote"
            optQuote.Value = True
            Call optQuote_Click
        Case "uptime"
            optUptime.Value = True
            Call optUptime_Click
        Case "mp3"
            optMP3.Value = True
            Call optMP3_Click
        Case Else
            optMsg.Value = True
            Call optMsg_Click
    End Select
    
    txtIdleMsg.Text = Config.IdleMessageText
    If LenB(txtIdleMsg.Text) = 0 Then txtIdleMsg.Text = "/me is a %v by Stealth - http://www.stealthbot.net"
    
    chkIdles.Value = Abs(Config.IdleMessage)
    Call chkIdles_Click
    
End Sub

Private Sub InitGenMisc()
    chkAllowMP3.Value = Abs(Config.Mp3Commands)
    chkPAmp.Value = Abs(Config.ProfileAmp)
    chkMail.Value = Abs(Config.BotMail)
    chkClanInvitations.Value = Abs(Not Config.IgnoreClanInvites)

    chkWhisperCmds.Value = Abs(Config.WhisperCommands)
    chkShowOffline.Value = Abs(Config.ShowOfflineFriends)
    
    txtOwner.Text = Config.BotOwner
    txtTrigger.Text = Config.Trigger
    
    chkBackup.Value = Abs(Config.UseBackupChannel)
    Call chkBackup_Click
    txtBackupChan.Text = Config.BackupChannel
    
    optNaming(Config.NamespaceConvention).Value = True
    chkD2Naming.Value = Abs(Config.UseD2Naming)
End Sub

' END INIT SUBS


Function CDKeyReplacements(ByVal inString As String) As String
    inString = Replace(inString, "-", "")
    inString = Replace(inString, " ", "")
    CDKeyReplacements = Trim$(inString)
End Function

' Returns true if the client is banned
Private Function IsClientBanned(ByVal sProductCode As String) As Boolean
    IsClientBanned = (InStr(1, Database.GetAccess(sProductCode, DB_TYPE_GAME).Flags, "B", _
        vbBinaryCompare) > 0)
End Function

Private Sub SaveClientBans()
    Dim Clients(6)  As String
    Dim i           As Integer          ' counters
    Dim bChanged    As Boolean          ' TRUE if changes have been made for a single entry
    Dim bNeedSave   As Boolean          ' TRUE if anyt entry was changed
    
    Dim oEntry      As clsDBEntryObj
    Dim eAction     As enuDBActions
    
    Clients(SC) = PRODUCT_STAR
    Clients(BW) = PRODUCT_SEXP
    Clients(D2) = PRODUCT_D2DV
    Clients(D2X) = PRODUCT_D2XP
    Clients(W3) = PRODUCT_WAR3
    Clients(W3X) = PRODUCT_W3XP
    Clients(W2) = PRODUCT_W2BN

    For i = chkCBan.LBound To chkCBan.UBound
        ' Get the entry
        Set oEntry = Database.GetEntry(Clients(i), DB_TYPE_GAME)
        
        ' Is the ban enabled?
        If (chkCBan(i).Value = 1) Then
            ' Is there an existing entry for this game?
            If oEntry Is Nothing Then
                ' Nope, create a new one.
                Set oEntry = Database.CreateNewEntry(Clients(i), , DB_TYPE_GAME)
                oEntry.Flags = "B"
                Call Database.AddEntry(oEntry)
                
                eAction = AddEntry
                bChanged = True
            Else
                ' Does the entry need the B flag?
                If Not oEntry.HasFlag("B") Then
                    oEntry.Flags = oEntry.Flags & "B"
                    
                    eAction = ModEntry
                    bChanged = True
                End If
            End If
        Else
            If Not (oEntry Is Nothing) Then
                With oEntry
                
                    ' Will something be left after we remove this flag?
                    If Len(.Flags) > 1 Or .Rank > 0 Or .Groups.Count > 0 Then
                        ' If the entry has the B flag, remove it.
                        If .HasFlag("B") Then
                            .Flags = Replace(.Flags, "B", vbNullString)
                        
                            eAction = ModEntry
                            bChanged = True
                        End If
                    Else
                        ' Remove the entry
                        Call Database.RemoveEntry(oEntry)
                        eAction = RemEntry
                        bChanged = True
                    End If
                End With
            End If
        End If
        
        ' Log this change if it was made.
        If bChanged Then
            bNeedSave = True
            
            ' Update metadata
            With oEntry
                .ModifiedBy = Database.GetConsoleAccess().Username
                .ModifiedOn = Now()
            End With
            
            ' Log actions
            If Config.LogDBActions Then
                Call LogDbAction(eAction, oEntry.ModifiedBy, oEntry.Name, oEntry.EntryType, oEntry.Rank, oEntry.Flags)
            End If
        End If
        
        ' Reset change tracker
        bChanged = False
          
    Next i
    
    ' Save the database
    If bNeedSave Then
        Call Database.Save
    End If
End Sub

Sub SaveFontSettings()
    Dim ResizeChatElements As Boolean
    Dim ResizeChannelElements As Boolean
    
    If (StrComp(InitChatFont, txtChatFont.Text, vbTextCompare) <> 0) Then
        Config.ChatFont = txtChatFont.Text
        
        'frmChat.rtbChat.Font.Name = Config.ChatFont
        frmChat.cboSend.Font.Name = Config.ChatFont
        frmChat.txtPre.Font.Name = Config.ChatFont
        frmChat.txtPost.Font.Name = Config.ChatFont
        'frmChat.rtbWhispers.Font.Name = Config.ChatFont
        ResizeChatElements = True
    End If
    
    If Not InitChatSize = CInt(txtChatSize.Text) Then
        Config.ChatFontSize = Val(txtChatSize.Text)
        
        'frmChat.rtbChat.Font.Size = Config.ChatFontSize
        frmChat.cboSend.Font.Size = Config.ChatFontSize
        frmChat.txtPre.Font.Size = Config.ChatFontSize
        frmChat.txtPost.Font.Size = Config.ChatFontSize
        'frmChat.rtbWhispers.Font.Size = Config.ChatFontSize
        ResizeChatElements = True
    End If
    
    If (StrComp(InitChanFont, txtChanFont.Text, vbTextCompare) <> 0) Then
        Config.ChannelListFont = txtChanFont.Text
        
        frmChat.lvChannel.Font.Name = Config.ChannelListFont
        frmChat.lvClanList.Font.Name = Config.ChannelListFont
        frmChat.lvFriendList.Font.Name = Config.ChannelListFont
        frmChat.ListviewTabs.Font.Name = Config.ChannelListFont
        frmChat.lblCurrentChannel.Font.Name = Config.ChannelListFont
        ResizeChannelElements = True
    End If
    
    If Not InitChanSize = CInt(txtChanSize.Text) Then
        Config.ChannelListFontSize = Val(txtChanSize.Text)
        
        frmChat.lvChannel.Font.Size = Config.ChannelListFontSize
        frmChat.lvClanList.Font.Size = Config.ChannelListFontSize
        frmChat.lvFriendList.Font.Size = Config.ChannelListFontSize
        frmChat.ListviewTabs.Font.Size = Config.ChannelListFontSize
        frmChat.lblCurrentChannel.Font.Size = Config.ChannelListFontSize
        ResizeChannelElements = True
    End If
    
    If ResizeChannelElements Then
        Dim lblHeight As Single
        
        frmChat.lblCurrentChannel.AutoSize = True
        lblHeight = frmChat.lblCurrentChannel.Height + 40
        frmChat.lblCurrentChannel.AutoSize = False
        frmChat.lblCurrentChannel.Height = lblHeight
        ResizeChatElements = True
    End If
    
    If ResizeChatElements Then
        Call ChangeRTBFont(frmChat.rtbChat, Config.ChatFont, Config.ChannelListFontSize)
        Call ChangeRTBFont(frmChat.rtbWhispers, Config.ChatFont, Config.ChannelListFontSize)
        
        frmChat.Form_Resize
    End If
End Sub

Private Sub ChangeRTBFont(rtb As RichTextBox, ByVal NewFont As String, ByVal NewSize As Integer)
    Dim tmpBuffer As String
    
    With rtb
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = NewSize
        .SelFontName = NewFont
        tmpBuffer = .TextRTF
        .Text = vbNullString
        .Font.Name = NewFont
        .Font.Size = NewSize
        .TextRTF = tmpBuffer
        .SelStart = Len(.Text)
    End With
End Sub

Sub LoadColors()
    ReDim mColors(0)
    cboColorList.Clear
    
    With FormColors
        CAdd "Current Channel Label | Background", .ChannelLabelBack
        CAdd "Current Channel Label | Text", .ChannelLabelText
        CAdd "Channel List | Background", .ChannelListBack
        CAdd "Channel List | Normal Users", .ChannelListText
        CAdd "Channel List | Self", .ChannelListSelf
        CAdd "Channel List | Idle Users", .ChannelListIdle
        CAdd "Channel List | Squelched Users", .ChannelListSquelched
        CAdd "Channel List | Operators", .ChannelListOps
        CAdd "Chat Window | Background", .RTBBack
        CAdd "Send Boxes | Background", .SendBoxesBack
        CAdd "Send Boxes | Text", .SendBoxesText
    End With
    
    With RTBColors
        CAdd "Talk - Bot Username", .TalkBotUsername, 1
        CAdd "Talk - Normal Usernames", .TalkUsernameNormal, 1
        CAdd "Talk - Op Usernames", .TalkUsernameOp, 1
        CAdd "Talk - Normal Text", .TalkNormalText, 1
        CAdd "Talk - Carat Color", .Carats, 1
        CAdd "Emote - Text", .EmoteText, 1
        CAdd "Emote - Username", .EmoteUsernames, 1
        CAdd "Information - Neutral", .InformationText, 1
        CAdd "Information - Success", .SuccessText, 1
        CAdd "Information - Errors", .ErrorMessageText, 1
        CAdd "Information - Timestamps", .TimeStamps, 1
        CAdd "Information - Server Information", .ServerInfoText, 1
        CAdd "Information - Console Messages", .ConsoleText, 1
        CAdd "Join/Leave - Text", .JoinText, 1
        CAdd "Join/Leave - Username", .JoinUsername, 1
        CAdd "Channel Join - Name", .JoinedChannelName, 1
        CAdd "Channel Join - Text", .JoinedChannelText, 1
        CAdd "Whisper - Carat Color", .WhisperCarats, 1
        CAdd "Whisper - Text", .WhisperText, 1
        CAdd "Whisper - Usernames", .WhisperUsernames, 1
    End With
End Sub

Private Function StringToNumber(ByVal sNumber As String, Optional ByVal lDefault As Long = 0) As Long
    sNumber = Trim$(sNumber)
    
    If IsNumeric(sNumber) Then
        StringToNumber = CLng(sNumber)
    Else
        StringToNumber = lDefault
    End If
End Function




